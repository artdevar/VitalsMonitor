import tkinter as tk
from tkinter import ttk, messagebox
import json
import os
from winreg import OpenKey, EnumValue, HKEY_CURRENT_USER, QueryInfoKey


class ToolTip:
    """Create a tooltip for a given widget"""
    def __init__(self, widget, text):
        self.widget = widget
        self.text = text
        self.tooltip_window = None
        self.widget.bind("<Enter>", self.show_tooltip)
        self.widget.bind("<Leave>", self.hide_tooltip)

    def show_tooltip(self, event=None):
        if self.tooltip_window or not self.text:
            return
        x, y, _, _ = self.widget.bbox("insert")
        x += self.widget.winfo_rootx() + 25
        y += self.widget.winfo_rooty() + 25

        self.tooltip_window = tw = tk.Toplevel(self.widget)
        tw.wm_overrideredirect(True)
        tw.wm_geometry(f"+{x}+{y}")

        label = tk.Label(tw, text=self.text, justify=tk.LEFT,
                        background="#ffffe0", relief=tk.SOLID, borderwidth=1,
                        font=("tahoma", "8", "normal"))
        label.pack(ipadx=1)

    def hide_tooltip(self, event=None):
        if self.tooltip_window:
            self.tooltip_window.destroy()
            self.tooltip_window = None


class SensorData:
    def __init__(self, sensorName, regName, regNum):
        self.sensorName = sensorName
        self.regName = regName
        self.regNum = regNum

class SensorConfigGUI:
    """GUI for configuring sensor mappings"""

    HWINFO_KEY_PATH = r"SOFTWARE\HWiNFO64\VSB"
    CONFIG_FILE = "sensor_config.json"
    CONFIG_VERSION = "0.1"
    DEFAULT_COREPROPS_PATH = r"%ProgramData%\SteelSeries\SteelSeries Engine 3\coreProps.json"

    # Device types mapping: display_name -> device_type
    DEVICE_TYPES = {
        "Rival 700, Rival 710": "screened-128x36",
        "Apex 7, Apex 7 TKL, Apex Pro, Apex Pro TKL": "screened-128x40",
        "Arctis Pro Wireless": "screened-128x48",
        "GameDAC / Arctis Pro + GameDAC": "screened-128x52"
    }

    def __init__(self, root):
        self.root = root
        self.root.title("Sensor Configuration")
        self.root.geometry("600x400")

        self.sensor_entries = []  # List to store sensor combobox widgets
        self.available_sensors = self.get_available_sensors()

        # Main container
        main_frame = ttk.Frame(root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))

        # Configure grid weights for resizing
        root.columnconfigure(0, weight=1)
        root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(0, weight=1)
        main_frame.rowconfigure(4, weight=1)

        # Title
        title_label = ttk.Label(main_frame, text="Configure Sensor Mappings",
                               font=('Arial', 12, 'bold'))
        title_label.grid(row=0, column=0, pady=(0, 10), sticky=tk.W)

        # Update interval settings
        interval_frame = ttk.Frame(main_frame)
        interval_frame.grid(row=1, column=0, pady=(0, 10), sticky=(tk.W, tk.E))

        interval_label = ttk.Label(interval_frame, text="Update Interval (ms):")
        interval_label.pack(side=tk.LEFT, padx=(0, 5))
        ToolTip(interval_label, "How often to update sensor readings in milliseconds (e.g., 1000 = 1 second)")

        self.update_interval_entry = ttk.Entry(interval_frame, width=10)
        self.update_interval_entry.insert(0, "1000")
        self.update_interval_entry.pack(side=tk.LEFT)
        ToolTip(self.update_interval_entry, "SteelSeries recommends to not update the display to often. \n1000 ms should be plenty")

        # CoreProps.json path settings
        coreprops_frame = ttk.Frame(main_frame)
        coreprops_frame.grid(row=2, column=0, pady=(0, 10), sticky=(tk.W, tk.E))

        coreprops_label = ttk.Label(coreprops_frame, text="SteelSeries Engine coreProps.json path:")
        coreprops_label.pack(side=tk.LEFT, padx=(0, 5))
        ToolTip(coreprops_label, "Path to SteelSeries Engine coreProps.json configuration file")

        self.coreprops_entry = ttk.Entry(coreprops_frame, width=50)
        self.coreprops_entry.insert(0, self.DEFAULT_COREPROPS_PATH)
        self.coreprops_entry.pack(side=tk.LEFT, fill=tk.X, expand=True)
        self.coreprops_entry.bind('<KeyRelease>', self.validate_coreprops_path)
        ToolTip(self.coreprops_entry, "Make sure the SteelSeries Engine is installed")

        # Initial validation
        self.validate_coreprops_path()

        # Device selection
        device_frame = ttk.Frame(main_frame)
        device_frame.grid(row=3, column=0, pady=(0, 10), sticky=(tk.W, tk.E))

        device_label = ttk.Label(device_frame, text="Device:")
        device_label.pack(side=tk.LEFT, padx=(0, 5))
        ToolTip(device_label, "Select your SteelSeries device with OLED screen")

        device_display_values = list(self.DEVICE_TYPES.keys())
        self.device_combo = ttk.Combobox(device_frame, values=device_display_values, state="readonly", width=40)
        if device_display_values:
            self.device_combo.current(0)  # Default to first device
        self.device_combo.pack(side=tk.LEFT)
        ToolTip(self.device_combo, "This determines the screen resolution for your device")

        # Scrollable frame for sensor entries
        canvas_frame = ttk.Frame(main_frame)
        canvas_frame.grid(row=4, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        canvas_frame.columnconfigure(0, weight=1)
        canvas_frame.rowconfigure(0, weight=1)

        canvas = tk.Canvas(canvas_frame, highlightthickness=0)
        scrollbar = ttk.Scrollbar(canvas_frame, orient="vertical", command=canvas.yview)
        self.sensor_frame = ttk.Frame(canvas)

        canvas.configure(yscrollcommand=scrollbar.set)

        scrollbar.grid(row=0, column=1, sticky=(tk.N, tk.S))
        canvas.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))

        canvas_window = canvas.create_window((0, 0), window=self.sensor_frame, anchor="nw")

        def configure_scroll_region(event):
            canvas.configure(scrollregion=canvas.bbox("all"))
            canvas.itemconfig(canvas_window, width=event.width)

        self.sensor_frame.bind("<Configure>", configure_scroll_region)
        canvas.bind("<Configure>", lambda e: canvas.itemconfig(canvas_window, width=e.width))

        # Button frame
        button_frame = ttk.Frame(main_frame)
        button_frame.grid(row=5, column=0, pady=(10, 0), sticky=(tk.W, tk.E))

        ttk.Button(button_frame, text="Add Line", command=self.add_sensor_entry).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="Save Config", command=self.save_config).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="Load Config", command=self.load_config).pack(side=tk.LEFT, padx=5)

        # Add initial sensor entry
        self.add_sensor_entry()

    def get_available_sensors(self):
        """Read available sensors from Windows Registry (HWiNFO64) - only Sensor%N keys"""
        sensors = {}
        try:
            with OpenKey(HKEY_CURRENT_USER, self.HWINFO_KEY_PATH) as key:
              index = 0
              while True:
                  try:
                      value_name, value_data, _ = EnumValue(key, index)
                      index += 1
                  except OSError:
                      break

                  if value_name.startswith('Sensor'):
                    sensors[value_name] = SensorData(str(value_data), value_name, int(value_name[6:]))

        except FileNotFoundError:
            messagebox.showwarning("Registry Key Not Found",
                                  f"HWiNFO64 registry key not found at:\n{self.HWINFO_KEY_PATH}\n\n"
                                  "Make sure HWiNFO64 is running with shared memory enabled.")
        except Exception as e:
            messagebox.showerror("Error", f"Error reading registry: {str(e)}")

        return sensors

    def validate_coreprops_path(self, event=None):
        """Validate if coreProps.json path exists and change color accordingly"""
        path = self.coreprops_entry.get()
        expanded_path = os.path.expandvars(path)

        if os.path.isfile(expanded_path):
            self.coreprops_entry.configure(foreground='green')
        else:
            self.coreprops_entry.configure(foreground='red')

    def add_sensor_entry(self):
        """Add a new sensor selection row"""
        row = len(self.sensor_entries)

        entry_frame = ttk.Frame(self.sensor_frame)
        entry_frame.grid(row=row, column=0, sticky=(tk.W, tk.E), pady=5, padx=5)
        entry_frame.columnconfigure(2, weight=1)

        label = ttk.Label(entry_frame, text=f"Line {row + 1}:")
        label.grid(row=0, column=0, sticky=tk.W, padx=(0, 10))

        name_label = ttk.Label(entry_frame, text="Display Format:")
        name_label.grid(row=0, column=1, sticky=tk.W, padx=(0, 5))
        ToolTip(name_label, "Format string for display. Use {0}, {1}, {2}... for sensor values.\nExamples: 'CPU: {0:.1f}°C' or 'CPU: {0:.0f} | GPU: {1:.0f}°C'\nRead Python format() documentation for more info")

        name_entry = ttk.Entry(entry_frame, width=30)
        name_entry.insert(0, f"Line {row + 1}: {{0:.1f}}°C")
        name_entry.grid(row=0, column=2, sticky=(tk.W, tk.E), padx=(0, 10))
        ToolTip(name_entry, "Use {0} for first sensor, {1} for second, etc. Format: {0:.1f} = 1 decimal, {0:.0f} = no decimals")

        remove_btn = ttk.Button(entry_frame, text="Remove Line", width=12,
                               command=lambda: self.remove_sensor_entry(entry_frame, row))
        remove_btn.grid(row=0, column=3, sticky=tk.E)

        sensors_subframe = ttk.Frame(entry_frame)
        sensors_subframe.grid(row=1, column=0, columnspan=4, sticky=(tk.W, tk.E), pady=(5, 0))
        sensors_subframe.columnconfigure(0, weight=1)

        add_sensor_btn = ttk.Button(sensors_subframe, text="+ Add Sensor", width=15,
                                    command=lambda: self.add_sensor_to_line(sensors_subframe, entry_frame))
        add_sensor_btn.grid(row=0, column=0, sticky=tk.W, padx=(30, 0))

        self.sensor_entries.append({
            'frame': entry_frame,
            'name': name_entry,
            'sensors_frame': sensors_subframe,
            'sensor_combos': [],
            'row': row,
            'next_sensor_row': 1
        })

        self.add_sensor_to_line(sensors_subframe, entry_frame)

    def add_sensor_to_line(self, sensors_frame, parent_frame):
        """Add a sensor combobox to a line"""
        entry = next((e for e in self.sensor_entries if e['frame'] == parent_frame), None)
        if not entry:
            return

        sensor_row = entry['next_sensor_row']
        entry['next_sensor_row'] += 1

        sensor_combo_frame = ttk.Frame(sensors_frame)
        sensor_combo_frame.grid(row=sensor_row, column=0, sticky=(tk.W, tk.E), pady=2)

        sensor_index = len(entry['sensor_combos'])
        sensor_label = ttk.Label(sensor_combo_frame, text=f"Sensor {{{sensor_index}}}:")
        sensor_label.grid(row=0, column=0, sticky=tk.W, padx=(30, 5))

        display_values = [sensor.sensorName for sensor in self.available_sensors.values()]
        combo = ttk.Combobox(sensor_combo_frame, values=display_values, state="readonly", width=40)
        if display_values:
            combo.current(min(sensor_index, len(display_values) - 1))
        combo.grid(row=0, column=1, sticky=(tk.W, tk.E), padx=(0, 5))

        remove_sensor_btn = ttk.Button(sensor_combo_frame, text="Remove", width=8,
                                       command=lambda: self.remove_sensor_from_line(sensor_combo_frame, entry))
        remove_sensor_btn.grid(row=0, column=2, sticky=tk.E)

        entry['sensor_combos'].append({
            'frame': sensor_combo_frame,
            'combo': combo,
            'label': sensor_label
        })

    def remove_sensor_from_line(self, sensor_frame, entry):
        """Remove a sensor from a line"""
        if len(entry['sensor_combos']) <= 1:
            messagebox.showinfo("Info", "At least one sensor must remain per line.")
            return

        sensor_frame.destroy()
        entry['sensor_combos'] = [s for s in entry['sensor_combos'] if s['frame'] != sensor_frame]

        for i, sensor_info in enumerate(entry['sensor_combos']):
            sensor_info['label'].config(text=f"Sensor {{{i}}}:")

    def remove_sensor_entry(self, frame, row):
        """Remove a sensor entry"""
        if len(self.sensor_entries) <= 1:
            messagebox.showinfo("Info", "At least one sensor entry must remain.")
            return

        # Remove from UI
        frame.destroy()

        # Remove from list
        self.sensor_entries = [entry for entry in self.sensor_entries if entry['frame'] != frame]

        # Re-index remaining entries
        for i, entry in enumerate(self.sensor_entries):
            entry['row'] = i
            for widget in entry['frame'].winfo_children():
                if isinstance(widget, ttk.Label) and widget.cget('text').startswith('Line'):
                    widget.config(text=f"Line {i + 1}:")
                    break

    def save_config(self):
        try:
            update_interval = int(self.update_interval_entry.get())
            if update_interval < 0:
                messagebox.showerror("Invalid Input", "Update interval cannot be negative!")
                return
        except ValueError:
            messagebox.showerror("Invalid Input", "Update interval must be a valid number!")
            return

        coreprops_path = self.coreprops_entry.get()
        expanded_path = os.path.expandvars(coreprops_path)
        if not os.path.exists(expanded_path):
            messagebox.showerror("Invalid Path", f"coreProps.json not found at:\n{expanded_path}\n\nPlease verify the path and ensure SteelSeries Engine is installed.")
            return

        device_name = self.device_combo.get()
        device_type = self.DEVICE_TYPES.get(device_name)

        config = {
            'version': self.CONFIG_VERSION,
            'update_interval_ms': update_interval,
            'coreprops_path': self.coreprops_entry.get(),
            'device_type': device_type,
            'lines': []
        }

        for entry in self.sensor_entries:
            display_format = entry['name'].get().strip()
            registry_keys = []

            for sensor_info in entry['sensor_combos']:
                sensor_display = sensor_info['combo'].get()
                if sensor_display:
                    sensor_key = None
                    for key, sensor_data in self.available_sensors.items():
                        if sensor_data.sensorName == sensor_display:
                            sensor_key = key
                            break

                    if sensor_key:
                        value_raw_key = sensor_key.replace('Sensor', 'ValueRaw', 1)
                        registry_keys.append(value_raw_key)

            if display_format and registry_keys:
                config['lines'].append({
                    'display_format': display_format,
                    'registry_keys': registry_keys
                })

        if not config['lines']:
            messagebox.showwarning("Warning", "No sensors configured!")
            return

        try:
            with open(self.CONFIG_FILE, 'w', encoding='utf-8') as f:
                json.dump(config, f, indent=4, ensure_ascii=False)
            messagebox.showinfo("Success", f"Configuration saved to {self.CONFIG_FILE}")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to save configuration: {str(e)}")

    def load_config(self):
        """Load configuration from JSON file"""
        try:
            with open(self.CONFIG_FILE, 'r') as f:
                config = json.load(f)

            # Check version
            config_version = config.get('version', 'unknown')
            if config_version != self.CONFIG_VERSION:
                messagebox.showwarning("Version Mismatch",
                                      f"Config file version ({config_version}) differs from current version ({self.CONFIG_VERSION}).\n"
                                      "Loading will continue, but some settings may not be compatible.")

            # Load update interval
            if 'update_interval_ms' in config:
                self.update_interval_entry.delete(0, tk.END)
                self.update_interval_entry.insert(0, str(config['update_interval_ms']))

            # Load coreProps path
            if 'coreprops_path' in config:
                self.coreprops_entry.delete(0, tk.END)
                self.coreprops_entry.insert(0, config['coreprops_path'])
                self.validate_coreprops_path()

            # Load device type
            if 'device_type' in config:
                device_type = config['device_type']
                # Find device name by device type value
                for device_name, dev_type in self.DEVICE_TYPES.items():
                    if dev_type == device_type:
                        self.device_combo.set(device_name)
                        break

            # Clear existing entries
            for entry in self.sensor_entries[:]:
                entry['frame'].destroy()
            self.sensor_entries.clear()

            lines_data = config.get('lines', config.get('sensors', []))
            if lines_data:
                for line in lines_data:
                    self.add_sensor_entry()
                    last_entry = self.sensor_entries[-1]
                    last_entry['name'].delete(0, tk.END)
                    display_format = line.get('display_format', line.get('name', ''))
                    last_entry['name'].insert(0, display_format)

                    registry_keys = line.get('registry_keys', line.get('registry_key', []))
                    if isinstance(registry_keys, str):
                        registry_keys = [registry_keys]

                    for sensor_combo_info in last_entry['sensor_combos'][:]:
                        sensor_combo_info['frame'].destroy()
                    last_entry['sensor_combos'].clear()
                    last_entry['next_sensor_row'] = 1

                    for i, registry_key in enumerate(registry_keys):
                        self.add_sensor_to_line(last_entry['sensors_frame'], last_entry['frame'])
                        sensor_key = registry_key.replace('ValueRaw', 'Sensor', 1)
                        if sensor_key in self.available_sensors:
                            display_text = self.available_sensors[sensor_key].sensorName
                            last_entry['sensor_combos'][-1]['combo'].set(display_text)

            messagebox.showinfo("Success", "Configuration loaded successfully!")
        except FileNotFoundError:
            messagebox.showwarning("Warning", f"Configuration file '{self.CONFIG_FILE}' not found.")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to load configuration: {str(e)}")


def main():
    root = tk.Tk()
    app = SensorConfigGUI(root)
    try:
        root.mainloop()
    except KeyboardInterrupt:
        print("Interrupted by user")
        root.destroy()


if __name__ == "__main__":
    main()
