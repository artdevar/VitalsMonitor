
import os
import sys
import json
import requests
import threading
import time
import logging
from winreg import OpenKey, EnumValue, HKEY_CURRENT_USER


logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    datefmt='%Y-%m-%d %H:%M:%S',
    handlers=[
        logging.FileHandler('log.txt', mode='w'),
        logging.StreamHandler()
    ]
)

class SharedSensorsValues:
    def __init__(self, lines: list):
        self.values = {}
        for line in lines:
            for key in line.sensors:
                self.values[key] = 0.0

    def set_value(self, key: str, value: float):
        self.values[key] = value

    def get_value(self, key: str, default=0.0):
        return self.values.get(key, default)

    def get(self):
        return self.values

class LineConfig:
    def __init__(self, display_format, sensors: list):
        self.display_format = display_format
        self.sensors = sensors
        self.line_key = '_'.join(self.sensors)


class Config:
    CONFIG_FILE = "sensor_config.json"

    def __init__(self):
        self.version = None
        self.update_interval_ms = -1
        self.coreprops_path = ''
        self.device_type = ''
        self.lines = []

    def load(self):
        try:
            with open(self.CONFIG_FILE, 'r', encoding='utf-8') as f:
                data = json.load(f)

            self.version = data.get('version')
            self.update_interval_ms = data.get('update_interval_ms', 1000)
            self.coreprops_path = data.get('coreprops_path', r"%ProgramData%\SteelSeries\SteelSeries Engine 3\coreProps.json")
            self.device_type = data.get('device_type', "screened-128x48")

            self.lines = []
            lines_data = data.get('lines', data.get('sensors', []))
            for line_data in lines_data:
                line = LineConfig(
                    display_format=line_data.get('display_format', ''),
                    sensors=line_data.get('registry_keys', line_data.get('registry_key', []))
                )
                self.lines.append(line)

            logging.info(f"Configuration loaded: {len(self.lines)} line(s), device: {self.device_type}")
            return True
        except FileNotFoundError:
            logging.warning(f"Configuration file '{self.CONFIG_FILE}' not found. Using defaults.")
            return False
        except Exception as e:
            logging.error(f"Error loading configuration: {e}")
            return False


class SensorMonitorThread():
    HWINFO_KEY_PATH = r"SOFTWARE\\HWiNFO64\\VSB"

    def __init__(self, config, shared_values, update_event):
        self.config = config
        self.shared_values = shared_values
        self.thread = None
        self.started = False
        self.update_event = update_event

    def __del__(self):
        self.stop()

    def start(self):
        self.started = True
        self.thread = threading.Thread(target=self.__monitor_sensors, daemon=True)
        self.thread.start()

    def stop(self):
        self.started = False
        if self.thread and self.thread.is_alive():
            self.thread.join(timeout=5)

    def __monitor_sensors(self):
        with OpenKey(HKEY_CURRENT_USER, self.HWINFO_KEY_PATH) as key:
            while self.started:
                is_updated = self.__pull_sensors_values(key)
                if is_updated:
                    self.update_event.set()
                time.sleep(self.config.update_interval_ms / 1000.0)

    def __pull_sensors_values(self, key):
        index = 0
        changed = False

        while True:
            try:
                key_name, value_data, _ = EnumValue(key, index)
                index += 1
            except OSError:
                break

            if key_name in self.shared_values.get():
                new_value = float(value_data)
                self.shared_values.set_value(key_name, new_value)
                changed = True

        return changed


class DisplayUpdaterThread():
    CMD_REGISTER_GAME = "game_metadata"
    CMD_REGISTER_EVENT = "bind_game_event"
    CMD_SEND_EVENT = "game_event"
    DISPLAY_NAME = "Vitals Monitor"
    GAME_NAME = "VITALS_MONITOR_APP"
    EVENT_NAME = "DISPLAY_TEXT"
    DEVELOPER = "artdevar"

    def __init__(self, config, shared_values, gamesense: str, update_event):
        self.config = config
        self.shared_values = shared_values
        self.gamesense_address = gamesense
        self.thread = None
        self.started = False
        self.update_event = update_event

    def __del__(self):
        self.stop()

    def start(self):
        self.__init_display()
        self.started = True
        self.thread = threading.Thread(target=self.__update_loop, daemon=True)
        self.thread.start()

    def stop(self):
        self.started = False
        if self.thread and self.thread.is_alive():
            self.thread.join(timeout=5)

    def __update_loop(self):
        while self.started:
            self.update_event.wait(timeout=3)
            if self.update_event.is_set():
                self.__send_event()
                self.update_event.clear()

    def __init_display(self):
        self.__register()
        self.__bind_event()

    def __register(self):
        url = f"{self.gamesense_address}/{self.CMD_REGISTER_GAME}"
        payload = {
            "game": self.GAME_NAME,
            "game_display_name": self.DISPLAY_NAME,
            "developer": self.DEVELOPER
        }
        requests.post(url, json=payload)

    def __bind_event(self):
        url = f"{self.gamesense_address}/{self.CMD_REGISTER_EVENT}"
        lines = []
        for line in self.config.lines:
            lines.append({
                "has-text": True,
                "context-frame-key": line.line_key
            })

        payload = {
            "game": self.GAME_NAME,
            "event": self.EVENT_NAME,
            "value_optional": True,
            "handlers": [
                {
                    "device-type": self.config.device_type,
                    "mode": "screen",
                    "zone": "one",
                    "datas": [
                        {
                            "lines": lines
                        }
                    ]
                }
            ]
        }
        requests.post(url, json=payload)

    def __send_event(self):
        url = f"{self.gamesense_address}/{self.CMD_SEND_EVENT}"
        frame = {}
        for line in self.config.lines:
            values = [self.shared_values.get_value(key) for key in line.sensors]
            frame[line.line_key] = line.display_format.format(*values)

        payload = {
            "game": self.GAME_NAME,
            "event": self.EVENT_NAME,
            "data": {
                "frame": frame
            }
        }
        requests.post(url, json=payload)


def find_core_props_path(path):
    core_props_path = os.path.expandvars(path)
    if os.path.exists(core_props_path):
        return core_props_path
    logging.error(f"coreProps.json not found at {core_props_path}")
    return None


def get_gamesense_address(config):
    core_props = find_core_props_path(config.coreprops_path)
    with open(core_props, 'r') as f:
        data = json.load(f)
        address = data.get("address")
        return f"http://{address}"


if __name__ == "__main__":
    config = Config()
    is_loaded = config.load()
    if not is_loaded:
        logging.error("The sensor_config.json is not loaded. Run config_gui.pyw first and save a config.")
        sys.exit()

    try:
        gamesense_address = get_gamesense_address(config)
        logging.info(f"SteelSeries GameSense server: {gamesense_address}")
    except Exception as e:
        logging.error(f"Unable to retrieve the GameSense address: {e}")
        sys.exit()

    sensor_monitor = None
    display_updater = None
    shared_values = SharedSensorsValues(config.lines)
    update_event = threading.Event()

    try:
        sensor_monitor = SensorMonitorThread(config, shared_values, update_event)
        sensor_monitor.start()

        display_updater = DisplayUpdaterThread(config, shared_values, gamesense_address, update_event)
        display_updater.start()
    except Exception as e:
        logging.error(f"Unable to start threads: {e}")
        sys.exit()

    try:
        while True:
            time.sleep(3600)
    except Exception as e:
        logging.error(f"Something went wrong: {e}")
    except KeyboardInterrupt:
        logging.info("Closing")
