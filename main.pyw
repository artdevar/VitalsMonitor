import os
import json
import requests
import threading
import time
from winreg import OpenKey, EnumValue, HKEY_CURRENT_USER

# Global params
# chatgpt says non-atomic variables are ok in this case
cpuTemp = 0.0
gpuTemp = 0.0
mem1Temp = 0.0
mem2Temp = 0.0

# Constants
# timeout 15 seconds
CMD_REGISTER_GAME = "game_metadata"
CMD_REGISTER_EVENT = "bind_game_event"
CMD_SEND_EVENT = "game_event"

PULL_SENSORS_DATA_INTERVAL = 1.7 # seconds
UPDATE_DISPLAY_INVERVAL = 2 # seconds

CORE_PROPS_PATH = r"%ProgramData%\SteelSeries\SteelSeries Engine 3\coreProps.json"
GAME_NAME = "MONITOR_APP"
EVENT_NAME = "DISPLAY_TEXT"
CELSIUS = "°C"

HWINFO_KEY_PATH = r"SOFTWARE\HWiNFO64\VSB"
CPU_TEMP_KEY_PATH = "ValueRaw0"
MEM1_TEMP_KEY_PATH = "ValueRaw1"
MEM2_TEMP_KEY_PATH = "ValueRaw2"
GPU_TEMP_KEY_PATH = "ValueRaw3"

CPU_FRAME_NAME = "cpu"
GPU_FRAME_NAME = "gpu"
RAM_FRAME_NAME = "ram"

def FindCorePropsPath():
    core_props_path = os.path.expandvars(CORE_PROPS_PATH)
    if os.path.exists(core_props_path):
        return core_props_path
    else:
        print(f"coreProps.json not found at {core_props_path}")
        return None

def GetGamesenseAddress():
    try:
        with open(FindCorePropsPath(), 'r') as f:
            data = json.load(f)
            address = data.get("address")
            return f"http://{address}"
    except Exception as e:
        print(f"Error reading coreProps.json: {e}")

# Step 2: Register your application
def Register():
    url = f"{address}/{CMD_REGISTER_GAME}"
    payload = {
        "game": GAME_NAME,
        "game_display_name": "Monitor",
        "developer": "AB"
    }
    requests.post(url, json=payload)

  # Step 3: Bind an event to use the OLED screen
def BindEvent():
    url = f"{address}/{CMD_REGISTER_EVENT}"
    payload = {
        "game": GAME_NAME,
        "event": EVENT_NAME,
        "value_optional": True,
        "handlers": [
            {
                # OLED on Arctis Pro Wireless
                "device-type": "screened-128x48",  # https://github.com/SteelSeries/gamesense-sdk/blob/master/doc/api/standard-zones.md#screened-screened-widthxheight
                "mode": "screen",
                "zone": "one",
                "datas": [
                    {
                        "lines": [
                            {
                                "has-text": True,
                                "context-frame-key": CPU_FRAME_NAME
                            },
                            {
                                "has-text": True,
                                "context-frame-key": GPU_FRAME_NAME
                            },
                            {
                                "has-text": True,
                                "context-frame-key": RAM_FRAME_NAME
                            }
                        ]
                    }
                ]
            }
        ]
    }
    requests.post(url, json=payload)


# Step 4: Send the actual message to the screen
def SendEvent():
    url = f"{address}/{CMD_SEND_EVENT}"
    payload = {
        "game": GAME_NAME,
        "event": EVENT_NAME,
                "data": {
                    "frame": {
                        CPU_FRAME_NAME: f"CPU: {cpuTemp:.1f}{CELSIUS}",
                        GPU_FRAME_NAME: f"GPU: {gpuTemp:.1f}{CELSIUS}",
                        RAM_FRAME_NAME: f"RAM: {int(mem1Temp)}|{int(mem2Temp)}{CELSIUS}"
                    }
                }
    }
    requests.post(url, json=payload)

def PullSensorsValues(key):
    index = 0
    while True:
        try:
            value_name, value_data, _ = EnumValue(key, index)
            index += 1
        except OSError:
            break

        if value_name == CPU_TEMP_KEY_PATH:
            global cpuTemp
            value_data = float(value_data)
            cpuTemp = value_data
        elif value_name == MEM1_TEMP_KEY_PATH:
            global mem1Temp
            value_data = float(value_data)
            mem1Temp = value_data
        elif value_name == MEM2_TEMP_KEY_PATH:
            global mem2Temp
            value_data = float(value_data)
            mem2Temp = value_data
        elif value_name == GPU_TEMP_KEY_PATH:
            global gpuTemp
            value_data = float(value_data)
            gpuTemp = value_data

def MonitorSensors():
    with OpenKey(HKEY_CURRENT_USER, HWINFO_KEY_PATH) as key:
        while True:
            PullSensorsValues(key)
            time.sleep(PULL_SENSORS_DATA_INTERVAL)

if __name__ == "__main__":
    address = GetGamesenseAddress()
    if not address:
        print("Could not find SteelSeries GameSense address")
        exit()
    print(f"SteelSeries GameSense server: {address}")

    monitoringThread = threading.Thread(
            target=MonitorSensors,
            daemon=True  # Thread exits when main program does
        )
    monitoringThread.start()
    time.sleep(0.5)  # Give the thread some time to start

    # Run all steps
    Register()
    BindEvent()

    while True:
        SendEvent()
        time.sleep(UPDATE_DISPLAY_INVERVAL)
