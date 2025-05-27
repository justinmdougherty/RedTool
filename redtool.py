import serial
import time
import threading
import tkinter as tk
from tkinter import ttk, scrolledtext, messagebox, filedialog
import queue
from datetime import datetime
import re
import traceback
import xml.etree.ElementTree as ET
import os
import json

class BoltTerminalGUI:
    CONFIG_FILENAME = "redtool_config.json"

    def __init__(self, root):
        self.root = root
        self.root.title("BOLT Terminal & Configurator")
        self.root.geometry("900x850")
        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)

        # --- Serial Communication Variables ---
        self.serial_port = tk.StringVar(value="COM2")
        self.baudrate = tk.IntVar(value=19200)
        self.is_connected = False
        self.serial_thread = None
        self.stop_thread = threading.Event()
        self.serial = None
        self.is_bolt_connected = False
        self.incoming_line = ""
        self.message_queue = queue.Queue()
        self.start_time = 0
        self.auto_connect = tk.BooleanVar(value=True)
        self.auto_scroll = tk.BooleanVar(value=True)
        self.char_delay = tk.DoubleVar(value=0.01)
        # --- Events for specific prompt detection ---
        self.main_prompt_event = threading.Event()
        self.needs_key_event = threading.Event()
        # --- Variables to hold expected prompts ---
        self.currently_expected_prompt = None
        self.currently_needs_key_prompt = None

        # --- Device Information Variables ---
        self.fset_value = tk.StringVar(value="--")
        self.unit_id = tk.StringVar(value="--")
        self.brick_number = tk.StringVar(value="--")
        self.start_time_value = tk.StringVar(value="--")
        self.loaded_item_value = tk.StringVar(value="--")
        self.slot_value = tk.StringVar(value="--")

        # --- Configuration Variables ---
        self.slot_xml_paths = [tk.StringVar(value="") for _ in range(4)]
        self.ame_tek_path = tk.StringVar(value="")
        self.wfc_tek_path = tk.StringVar(value="")

        # --- Regex Patterns ---
        self.patterns = {
            'fset': re.compile(r'^\s*\bFSET\b\s*[:=]?\s*(\d+)', re.IGNORECASE),
            'unit_id': re.compile(r'^\s*\bUnitID\b\s*[:=]?\s*(\w+)', re.IGNORECASE),
            'brick_number': re.compile(r'^\s*\bBrickNumber\b\s*[:=]?\s*(\w+)', re.IGNORECASE),
            'start_time': re.compile(r'^\s*\bStart\b\s+(\d{2}:\d{2}:\d{2})', re.IGNORECASE),
            'loading_info': re.compile(r'^\s*Loading\s+(\S+)\s+in\s+slot\s+(\d+)', re.IGNORECASE),
            'wfmid_slot': re.compile(r'^\s*WFMID\s+(\d+)', re.IGNORECASE),
            'waveform_id': re.compile(r'^\s*Waveform\s+(\d+)', re.IGNORECASE)
        }

        # GUI element variables
        self.rx_count = tk.IntVar(value=0); self.tx_count = tk.IntVar(value=0)
        self.status_var = tk.StringVar(value="Disconnected")

        # Create the GUI elements
        self.create_menu()
        self.create_widgets()
        self.load_app_settings()
        self.process_messages()

    # --- Method Definitions ---

    def set_command_entry(self, command_text):
        """Clears and sets the text in the command entry box."""
        self.cmd_entry.delete(0, tk.END)
        self.cmd_entry.insert(0, command_text)

    def toggle_connection(self):
        """Connects if disconnected, disconnects if connected."""
        if not self.is_connected:
            self.connect()
        else:
            self.disconnect()

    def request_device_info(self):
        """Sends commands to the device to request status information."""
        if not self.is_connected or not self.is_bolt_connected:
             messagebox.showwarning("Not Connected", "Connect to BOLT first.")
             return
        # Define the specific commands needed to refresh info
        commands = ["info", "bricknumber", "unitid"] # Example commands
        self.add_message(f"Manual request for device info: {', '.join(commands)}", "info")
        def send_next_command(index=0):
             if index < len(commands):
                 cmd_term = commands[index] + '\r\n'
                 self.send_command(command_to_send=cmd_term, from_gui=False)
                 # Schedule next command after a delay
                 self.root.after(500, lambda: send_next_command(index + 1))
        if commands: send_next_command(0)

    def copy_device_info(self):
        """Copies the currently displayed device info to the clipboard."""
        info_text = f"""BOLT Info:
Connection: {self.status_var.get()}
Start Time: {self.start_time_value.get()}
FSET: {self.fset_value.get()}
Loaded Item: {self.loaded_item_value.get()}
Slot: {self.slot_value.get()}
UnitID: {self.unit_id.get()}
Brick Num: {self.brick_number.get()}"""
        try:
            self.root.clipboard_clear()
            self.root.clipboard_append(info_text)
            messagebox.showinfo("Info Copied", "Device info copied to clipboard.")
        except Exception as e:
            messagebox.showerror("Clipboard Error", f"Could not copy info:\n{e}")

    def send_command_event(self, event=None):
        """Handles the Enter key press or Send button click for command input."""
        command = self.cmd_entry.get().strip()
        if command:
             cmd_term = command + '\r\n' # Assume commands need CR+LF
             self.send_command(command_to_send=cmd_term, from_gui=True)
             self.cmd_entry.delete(0, tk.END) # Clear entry after sending

    def clear_output(self):
        """Clears the terminal output text area."""
        try:
             self.output_text.config(state=tk.NORMAL)
             self.output_text.delete(1.0, tk.END)
             self.output_text.config(state=tk.DISABLED)
        except Exception as e:
            # Log error to GUI instead of just console if possible
            self.add_message(f"Error clearing output: {e}", "error")
            print(f"Error clearing output: {e}") # Keep console print for critical errors

    def save_log(self):
        """Opens a dialog to save the terminal output to a file."""
        try:
            default_filename = f"bolt_log_{datetime.now().strftime('%Y%m%d_%H%M%S')}.log"
            filename = filedialog.asksaveasfilename(
                defaultextension=".log",
                filetypes=[("Log files", "*.log"), ("Text files", "*.txt"), ("All files", "*.*")],
                title="Save Log As",
                initialfile=default_filename
            )
            if not filename: return # User cancelled
            # Get all text from the ScrolledText widget
            log_content = self.output_text.get(1.0, tk.END)
            with open(filename, 'w', encoding='utf-8') as f:
                f.write(log_content)
            messagebox.showinfo("Log Saved", f"Log successfully saved to:\n{filename}")
        except Exception as e:
            messagebox.showerror("Error Saving Log", f"Failed to save log:\n{str(e)}")

    def get_timestamp(self):
        """Returns a formatted timestamp string."""
        return datetime.now().strftime("%H:%M:%S.%f")[:-3]

    def load_app_settings(self):
        """Loads settings like TEK file paths from the config file."""
        try:
            if os.path.exists(self.CONFIG_FILENAME):
                with open(self.CONFIG_FILENAME, 'r') as f:
                    settings = json.load(f)
                    ame_path = settings.get("ame_tek_path", "")
                    wfc_path = settings.get("wfc_tek_path", "")
                    # Use set() method for Tkinter StringVars
                    self.ame_tek_path.set(ame_path)
                    self.wfc_tek_path.set(wfc_path)
                    self.add_message("Loaded saved TEK file paths.", "info")
                    # Only log if path is actually set
                    if ame_path: self.add_message(f"  AME: {os.path.basename(ame_path)}", "info")
                    if wfc_path: self.add_message(f"  WFC: {os.path.basename(wfc_path)}", "info")
            else:
                self.add_message(f"Settings file '{self.CONFIG_FILENAME}' not found. Select TEK files via File menu.", "info")
        except json.JSONDecodeError as e:
             self.add_message(f"Error decoding settings file {self.CONFIG_FILENAME}: {e}", "error")
        except Exception as e:
            self.add_message(f"Error loading settings: {e}", "error")

    def save_app_settings(self):
        """Saves current settings (like TEK paths) to the config file."""
        settings = {
            "ame_tek_path": self.ame_tek_path.get(),
            "wfc_tek_path": self.wfc_tek_path.get()
            # Add other settings to save here if needed
        }
        try:
            with open(self.CONFIG_FILENAME, 'w') as f:
                json.dump(settings, f, indent=4)
            # Optional: Log success
            # self.add_message(f"Settings saved to {self.CONFIG_FILENAME}", "info")
        except Exception as e:
            # Show error in GUI log instead of just printing
            self.add_message(f"Error saving settings to {self.CONFIG_FILENAME}: {e}", "error")
            print(f"Error saving settings to {self.CONFIG_FILENAME}: {e}") # Keep console print

    def on_closing(self):
        """Handles the window closing event, saves settings, disconnects."""
        print("Closing application...") # Console message
        self.add_message("Closing application, saving settings...", "info") # GUI message
        self.save_app_settings()
        self.stop_thread.set() # Signal threads to stop *before* disconnecting
        self.disconnect()
        print("Disconnect complete. Destroying window.")
        # Wait briefly for threads to potentially finish after disconnect signal
        # Note: Joining non-daemon threads here would be better practice if they weren't daemons
        # time.sleep(0.2)
        self.root.destroy()

    def select_slot_xml(self, slot_index):
        """Opens dialog to select XML config file for a specific slot."""
        # Determine initial directory (e.g., current dir or last used)
        initial_dir = "."
        current_slot_path = self.slot_xml_paths[slot_index].get()
        if current_slot_path and os.path.exists(os.path.dirname(current_slot_path)):
            initial_dir = os.path.dirname(current_slot_path)

        filepath = filedialog.askopenfilename(
            title=f"Select Configuration XML for Slot {slot_index + 1}",
            initialdir=initial_dir,
            filetypes=[("XML files", "*.xml"), ("All files", "*.*")]
        )
        if filepath:
             self.slot_xml_paths[slot_index].set(filepath)
             self.add_message(f"Selected Slot {slot_index + 1} XML: {os.path.basename(filepath)}", "info")

    def select_tek_file(self, tek_type):
        """Opens dialog to select AME or WFC TEK file."""
        title = f"Select {tek_type.upper()} TEK File"
        initial_dir = "." # Default directory
        current_path = ""
        if tek_type == 'ame': current_path = self.ame_tek_path.get()
        elif tek_type == 'wfc': current_path = self.wfc_tek_path.get()

        # Set initial directory to the directory of the currently selected file, if valid
        if current_path and os.path.exists(os.path.dirname(current_path)):
             initial_dir = os.path.dirname(current_path)

        filepath = filedialog.askopenfilename(
            title=title,
            initialdir=initial_dir,
            filetypes=[("TEK files", "*.tek"), ("XML files", "*.xml"), ("All files", "*.*")]
        )
        if filepath:
            if tek_type == 'ame':
                 self.ame_tek_path.set(filepath)
                 self.add_message(f"Selected AME TEK File: {os.path.basename(filepath)}", "info")
            elif tek_type == 'wfc':
                 self.wfc_tek_path.set(filepath)
                 self.add_message(f"Selected WFC TEK File: {os.path.basename(filepath)}", "info")
            # Save settings immediately after selection for persistence
            self.save_app_settings()    

    # --- GUI Creation Methods --- (create_menu, create_widgets - unchanged, assuming they are correct)
    def create_menu(self):
        self.menu_bar = tk.Menu(self.root); self.root.config(menu=self.menu_bar)
        file_menu = tk.Menu(self.menu_bar, tearoff=0); self.menu_bar.add_cascade(label="File", menu=file_menu)
        file_menu.add_command(label="Choose AME TEK File...", command=lambda: self.select_tek_file('ame'))
        file_menu.add_command(label="Choose WFC TEK File...", command=lambda: self.select_tek_file('wfc'))
        file_menu.add_separator(); file_menu.add_command(label="Exit", command=self.on_closing)

    def create_widgets(self):
        # Connection frame
        conn_frame = ttk.LabelFrame(self.root, text="Connection Settings"); conn_frame.pack(side=tk.TOP, fill=tk.X, padx=10, pady=5)
        ttk.Label(conn_frame, text="Port:").grid(row=0, column=0, padx=5, pady=5, sticky=tk.W); ttk.Entry(conn_frame, textvariable=self.serial_port, width=10).grid(row=0, column=1, padx=5, pady=5, sticky=tk.W)
        ttk.Label(conn_frame, text="Baudrate:").grid(row=0, column=2, padx=5, pady=5, sticky=tk.W)
        self.baudrate_combo = ttk.Combobox(conn_frame, textvariable=self.baudrate, width=10, state="readonly"); self.baudrate_combo['values'] = ('9600', '19200', '38400', '57600', '115200'); self.baudrate_combo.grid(row=0, column=3, padx=5, pady=5, sticky=tk.W); self.baudrate_combo.set(str(self.baudrate.get()))
        self.connect_button = ttk.Button(conn_frame, text="Connect", command=self.toggle_connection); self.connect_button.grid(row=0, column=4, padx=5, pady=5, sticky=tk.W)
        ttk.Checkbutton(conn_frame, text="Auto-connect to BOLT", variable=self.auto_connect).grid(row=0, column=5, padx=5, pady=5, sticky=tk.W)
        ttk.Checkbutton(conn_frame, text="Auto-scroll", variable=self.auto_scroll).grid(row=0, column=6, padx=5, pady=5, sticky=tk.W)
        sending_frame = ttk.Frame(conn_frame); sending_frame.grid(row=1, column=0, columnspan=7, padx=5, pady=5, sticky=tk.W)
        ttk.Label(sending_frame, text="Char Delay (ms):").pack(side=tk.LEFT, padx=5)
        delay_values = ('0', '1', '5', '10', '15', '20', '50'); initial_delay_ms_str = str(int(self.char_delay.get() * 1000))
        self.delay_combo = ttk.Combobox(sending_frame, width=5, values=delay_values, state="readonly"); self.delay_combo.pack(side=tk.LEFT, padx=5); self.delay_combo.set(initial_delay_ms_str)
        self.delay_combo.bind("<<ComboboxSelected>>", lambda e: self.char_delay.set(float(self.delay_combo.get()) / 1000.0))

        # Status indicator frame
        status_frame = ttk.LabelFrame(self.root, text="Device Status"); status_frame.pack(side=tk.TOP, fill=tk.X, padx=10, pady=5)
        status_left = ttk.Frame(status_frame); status_left.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5, pady=5)
        ttk.Label(status_left, text="Connection:").grid(row=0, column=0, padx=5, pady=2, sticky=tk.W); self.status_label = ttk.Label(status_left, textvariable=self.status_var, foreground="red"); self.status_label.grid(row=0, column=1, padx=5, pady=2, sticky=tk.W)
        ttk.Label(status_left, text="Bytes Rcvd:").grid(row=1, column=0, padx=5, pady=2, sticky=tk.W); ttk.Label(status_left, textvariable=self.rx_count).grid(row=1, column=1, padx=5, pady=2, sticky=tk.W)
        ttk.Label(status_left, text="Bytes Sent:").grid(row=2, column=0, padx=5, pady=2, sticky=tk.W); ttk.Label(status_left, textvariable=self.tx_count).grid(row=2, column=1, padx=5, pady=2, sticky=tk.W)
        status_middle = ttk.Frame(status_frame); status_middle.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5, pady=5)
        ttk.Label(status_middle, text="FSET:").grid(row=0, column=0, padx=5, pady=2, sticky=tk.W); ttk.Label(status_middle, textvariable=self.fset_value).grid(row=0, column=1, padx=5, pady=2, sticky=tk.W)
        ttk.Label(status_middle, text="Loaded Item:").grid(row=1, column=0, padx=5, pady=2, sticky=tk.W); ttk.Label(status_middle, textvariable=self.loaded_item_value).grid(row=1, column=1, padx=5, pady=2, sticky=tk.W)
        ttk.Label(status_middle, text="Start Time:").grid(row=2, column=0, padx=5, pady=2, sticky=tk.W); ttk.Label(status_middle, textvariable=self.start_time_value).grid(row=2, column=1, padx=5, pady=2, sticky=tk.W)
        status_right = ttk.Frame(status_frame); status_right.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5, pady=5)
        ttk.Label(status_right, text="Unit ID:").grid(row=0, column=0, padx=5, pady=2, sticky=tk.W); ttk.Label(status_right, textvariable=self.unit_id).grid(row=0, column=1, padx=5, pady=2, sticky=tk.W)
        ttk.Label(status_right, text="Brick Num:").grid(row=1, column=0, padx=5, pady=2, sticky=tk.W); ttk.Label(status_right, textvariable=self.brick_number).grid(row=1, column=1, padx=5, pady=2, sticky=tk.W)
        ttk.Label(status_right, text="Slot:").grid(row=2, column=0, padx=5, pady=2, sticky=tk.W); ttk.Label(status_right, textvariable=self.slot_value).grid(row=2, column=1, padx=5, pady=2, sticky=tk.W)
        button_frame_info = ttk.Frame(status_frame); button_frame_info.pack(side=tk.RIGHT, padx=10, pady=5)
        ttk.Button(button_frame_info, text="Refresh Info", command=self.request_device_info).grid(row=0, column=0, padx=5, pady=2)
        ttk.Button(button_frame_info, text="Copy All Info", command=self.copy_device_info).grid(row=1, column=0, padx=5, pady=2)

        # Configuration Frame (Bolt Setup)
        config_frame = ttk.LabelFrame(self.root, text="Bolt Setup"); config_frame.pack(side=tk.TOP, fill=tk.X, padx=10, pady=5)
        self.slot_entry_widgets = []
        for i in range(4):
            slot_num = i + 1; ttk.Label(config_frame, text=f"Slot {slot_num} XML:").grid(row=i, column=0, padx=5, pady=(5,2), sticky=tk.W)
            entry = ttk.Entry(config_frame, textvariable=self.slot_xml_paths[i], width=45, state="readonly"); entry.grid(row=i, column=1, padx=5, pady=(5,2), sticky=tk.W); self.slot_entry_widgets.append(entry)
            ttk.Button(config_frame, text="Load XML", command=lambda s=i: self.select_slot_xml(s)).grid(row=i, column=2, padx=5, pady=(5,2))
        self.full_config_button = ttk.Button(config_frame, text="Load TEK Keys and Configure BOLT", command=self.start_full_configuration)
        self.full_config_button.grid(row=4, column=1, padx=5, pady=10, sticky=tk.W)

        # Terminal output
        output_frame = ttk.LabelFrame(self.root, text="Terminal Output"); output_frame.pack(side=tk.TOP, fill=tk.BOTH, expand=True, padx=10, pady=5)
        self.output_text = scrolledtext.ScrolledText(output_frame, wrap=tk.WORD, height=15, state=tk.DISABLED); self.output_text.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        self.output_text.tag_configure("timestamp", foreground="gray"); self.output_text.tag_configure("command", foreground="blue", font=('TkDefaultFont', 9, 'bold')); self.output_text.tag_configure("error", foreground="red"); self.output_text.tag_configure("info", foreground="green"); self.output_text.tag_configure("special_char", foreground="purple", font=('TkDefaultFont', 9, 'bold')); self.output_text.tag_configure("sent", foreground="darkorange", font=('TkDefaultFont', 9, 'bold')); self.output_text.tag_configure("standard", foreground="black")
        button_frame_output = ttk.Frame(output_frame); button_frame_output.pack(fill=tk.X, padx=5, pady=5)
        ttk.Button(button_frame_output, text="Clear", command=self.clear_output).pack(side=tk.LEFT, padx=5); ttk.Button(button_frame_output, text="Save Log", command=self.save_log).pack(side=tk.LEFT, padx=5)

        # Command input
        cmd_frame = ttk.LabelFrame(self.root, text="Command Input"); cmd_frame.pack(side=tk.TOP, fill=tk.X, padx=10, pady=5)
        self.cmd_entry = ttk.Entry(cmd_frame, width=50); self.cmd_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5, pady=5); self.cmd_entry.bind("<Return>", self.send_command_event)
        self.send_button = ttk.Button(cmd_frame, text="Send", command=self.send_command_event); self.send_button.pack(side=tk.LEFT, padx=5, pady=5)

        # Preset command buttons
        quick_cmd_frame = ttk.LabelFrame(self.root, text="Quick Commands"); quick_cmd_frame.pack(side=tk.TOP, fill=tk.X, padx=10, pady=5)
        cmds_row1 = ["help", "info", "tempc", "rfoff", "echo 1", "echo 0"]; cmds_row2 = ["powerlevel low", "powerlevel medium", "powerlevel high", "x 4 13", "q", "reboot"]
        for i, cmd in enumerate(cmds_row1): ttk.Button(quick_cmd_frame, text=cmd, command=lambda c=cmd: self.send_command(c + '\r\n', from_gui=True)).grid(row=0, column=i, padx=5, pady=5)
        for i, cmd in enumerate(cmds_row2): ttk.Button(quick_cmd_frame, text=cmd, command=lambda c=cmd: self.send_command(c + '\r\n', from_gui=True)).grid(row=1, column=i, padx=5, pady=5)


    # --- Serial Communication Methods ---
    def connect(self):
        if self.is_connected:
            self.add_message("Already connected.", "info")
            return
        try:
            port = self.serial_port.get()
            baudrate = self.baudrate.get()
            self.serial = serial.Serial(port, baudrate, timeout=0.1)
            time.sleep(0.5) # Allow port to open

            self.add_message(f"Opened {port} at {baudrate} baud.", "info")
            self.is_connected = True
            self.start_time = time.time()

            self.connect_button.config(text="Disconnect")
            self.status_var.set("Port Open")
            self.status_label.config(foreground="darkorange")

            # Reset connection-specific info
            self.rx_count.set(0); self.tx_count.set(0)
            self.fset_value.set("--"); self.unit_id.set("--"); self.brick_number.set("--")
            self.start_time_value.set("--"); self.loaded_item_value.set("--"); self.slot_value.set("--")
            self.is_bolt_connected = False # Reset BOLT connection status

            # Reset slot XML paths for the new session
            for path_var in self.slot_xml_paths:
                 path_var.set("")
            # DO NOT reset self.ame_tek_path or self.wfc_tek_path here

            self.stop_thread.clear()
            self.serial_thread = threading.Thread(target=self.read_serial, daemon=True)
            self.serial_thread.start()

            if self.auto_connect.get():
                self.root.after(1000, self.connect_to_bolt)

        except Exception as e:
            messagebox.showerror("Connection Error", f"Failed to connect: {str(e)}")
            self.status_var.set("Connection Failed")
            self.status_label.config(foreground="red")
            self.is_connected = False

    def disconnect(self):
        self.stop_thread.set()
        if self.serial_thread and self.serial_thread.is_alive():
             self.serial_thread.join(timeout=1.0)
        if self.serial and self.serial.is_open:
            try:
                self.serial.close()
                self.add_message("Serial port closed.", "info")
            except Exception as e:
                self.add_message(f"Error closing port: {e}", "error")
        self.serial = None
        self.is_connected = False
        self.is_bolt_connected = False
        self.connect_button.config(text="Connect")
        self.status_var.set("Disconnected")
        self.status_label.config(foreground="red")
        # Clear slot paths on disconnect as well? Optional, doing it here too.
        for path_var in self.slot_xml_paths:
            path_var.set("")

    def read_serial(self):
        while not self.stop_thread.is_set():
            if self.serial and self.serial.is_open:
                try:
                    # Auto-connect attempt (rate-limited in connect_to_bolt check)
                    if self.auto_connect.get() and not self.is_bolt_connected and time.time() - self.start_time >= 3:
                         if self.is_connected: # Ensure port still logically connected
                            # Use root.after to schedule the check in the main thread
                            self.root.after(0, self.connect_to_bolt)
                            # Prevent rapid re-attempts by setting start_time far in the future
                            # connect_to_bolt will reset it if connection fails or isn't needed
                            self.start_time = time.time() + 3600

                    if self.serial.in_waiting > 0:
                        data = self.serial.read(self.serial.in_waiting)
                        if data:
                            current_rx = self.rx_count.get()
                            self.rx_count.set(current_rx + len(data))
                            self.process_data_chunk(data)
                except serial.SerialException as e:
                    self.message_queue.put(("error", f"Serial error: {str(e)} - Disconnecting."))
                    # Schedule disconnect in main thread
                    self.root.after(0, self.disconnect)
                    break # Exit read thread
                except Exception as e:
                    self.message_queue.put(("error", f"Read error: {str(e)}"))
            else:
                # Port not open/object doesn't exist, wait before checking again
                if self.stop_thread.is_set(): break
                time.sleep(0.1) # Reduce CPU usage when disconnected
            time.sleep(0.005) # Small sleep to prevent busy-waiting

    def process_data_chunk(self, data_chunk):
        """Processes received bytes, checks for expected prompts/keys."""
        for byte_val in data_chunk:
            byte = bytes([byte_val])
            # Process line endings first
            if byte == b'\r': continue
            elif byte == b'\n' or byte == b'\xB6': # Treat CR or ¶ as line end
                line_content = self.incoming_line # Process before adding newline/marker
                self.incoming_line = "" # Reset buffer *before* processing line
                if line_content: # Process non-empty lines
                    self.check_for_device_info(line_content) # Check for general info patterns first
                    # Now check for specific expected prompts for config flow
                    clean_line = line_content.strip()
                    # Check if it matches the main prompt we might be waiting for
                    if self.currently_expected_prompt and clean_line == self.currently_expected_prompt:
                        # print(f"DEBUG: Matched main prompt: '{clean_line}'") # Debug print
                        self.main_prompt_event.set()
                    # Check if it matches the "Needs Key" prompt we might be waiting for
                    elif self.currently_needs_key_prompt and clean_line == self.currently_needs_key_prompt:
                        # print(f"DEBUG: Matched NeedsKey prompt: '{clean_line}'") # Debug print
                        self.needs_key_event.set()

                # Add the line ending character to the output queue
                if byte == b'\n': self.message_queue.put(("char", "\n"))
                else: self.message_queue.put(("special_char", "¶\n"))

            elif byte == b'\xA1': self.message_queue.put(("special_char", "¡")) # Display special char
            else: # Regular character
                try:
                    char = byte.decode('ascii', errors='replace')
                    self.incoming_line += char # Append to buffer
                    self.message_queue.put(("char", char)) # Display character
                    # Check for initializing message inline (still useful)
                    if self.incoming_line.endswith("Initializing:."):
                        self.message_queue.put(("info", "\nDevice Initializing detected."))
                        self.root.after(0, self.send_gps_timeout_command)
                except UnicodeDecodeError: self.message_queue.put(("char", "?"))


    def check_for_device_info(self, line):
        """Uses regex patterns to extract device info from a line."""
        clean_line = line.strip()
        if not clean_line: return
        for key, pattern in self.patterns.items():
            match = pattern.search(clean_line)
            if match:
                if key == 'loading_info' or key == 'wfmid_slot':
                    slot_value = match.group(1).strip() if key == 'wfmid_slot' else match.group(2).strip()
                    self.slot_value.set(slot_value); self.add_message(f"Extracted Slot: {slot_value}", "info")
                    if key == 'loading_info':
                         item_value = match.group(1).strip(); self.loaded_item_value.set(item_value); self.add_message(f"Extracted Loaded Item: {item_value}", "info")
                elif key == 'waveform_id':
                    waveform_num_str = match.group(1).strip(); item_name = "--"
                    if waveform_num_str == '6': item_name = "AME-6"
                    elif waveform_num_str == '8': item_name = "AME-8"
                    # Add other mappings if needed
                    if item_name != "--": self.loaded_item_value.set(item_name); self.add_message(f"Extracted Item (WF {waveform_num_str}): {item_name}", "info")
                else: # Other single-value patterns
                    value = match.group(1).strip(); variable_updated = False
                    if key == 'fset' and 'UTCOFFSET' not in clean_line.upper(): self.fset_value.set(value); variable_updated = True
                    elif key == 'unit_id': self.unit_id.set(value); variable_updated = True
                    elif key == 'brick_number': self.brick_number.set(value); variable_updated = True
                    elif key == 'start_time': self.start_time_value.set(value); variable_updated = True
                    if variable_updated: self.add_message(f"Extracted {key.replace('_', ' ').title()}: {value}", "info")
                # Don't break, allow multiple matches per line if necessary, though unlikely with ^ anchor


    def send_gps_timeout_command(self):
        if self.serial and self.serial.is_open:
             command = "*GPSTimeOut\r\n"
             self.send_command(command_to_send=command, from_gui=False)
             self.add_message(f"Auto-sent: {command.strip()}", "info")
        else: self.add_message("Cannot send GPS Timeout: Not connected.", "error")

    def connect_to_bolt(self):
        if not self.serial or not self.serial.is_open:
            # self.add_message("Cannot connect to BOLT: Port not open.", "error") # Too noisy for auto-connect
            return
        if self.is_bolt_connected: return # Already connected

        self.add_message("Attempting BOLT protocol connect...", "info")
        connection_command = b'\xA1O.ECP 2\xB6'
        final_baudrate = 115200
        try:
            self.serial.write(connection_command)
            current_tx = self.tx_count.get()
            self.tx_count.set(current_tx + len(connection_command))
            self.add_message(f"Sent BOLT Connect Sequence", "sent")

            # Change baud rate and assume connection
            time.sleep(0.5) # Wait for command to be processed
            self.serial.baudrate = final_baudrate
            time.sleep(0.1) # Short pause after baud change

            self.is_bolt_connected = True
            self.baudrate.set(final_baudrate) # Update internal variable

            # Update combobox in GUI thread
            try:
                 if hasattr(self, 'baudrate_combo') and self.root.winfo_exists():
                    self.root.after(0, lambda: self.baudrate_combo.set(str(final_baudrate)))
                 elif not hasattr(self, 'baudrate_combo'):
                      self.add_message("Baudrate combo not found.", "error")
            except Exception as e:
                 self.add_message(f"Combo update error: {e}", "error")

            self.add_message(f"Changed script baud to {final_baudrate}. Assuming BOLT connected.", "info")
            # Update status in GUI thread
            if self.root.winfo_exists():
                 self.root.after(0, lambda: self.status_var.set("Connected to BOLT"))
                 self.root.after(0, lambda: self.status_label.config(foreground="blue"))

            # Reset start_time to prevent immediate re-attempts by read_serial
            self.start_time = time.time()

        except Exception as e:
            self.add_message(f"BOLT connect error: {e}", "error")
            self.is_bolt_connected = False
             # Update status in GUI thread
            if self.root.winfo_exists():
                 self.root.after(0, lambda: self.status_var.set("BOLT Connect Failed"))
                 self.root.after(0, lambda: self.status_label.config(foreground="red"))
            # Reset start_time so auto-connect might try again later
            self.start_time = time.time()

    def request_device_info(self):
        if not self.is_connected or not self.is_bolt_connected:
             messagebox.showwarning("Not Connected", "Connect to BOLT first.")
             return
        commands = ["info", "bricknumber"] # Example commands to get info
        self.add_message(f"Manual request for device info: {', '.join(commands)}", "info")
        def send_next_command(index=0):
             if index < len(commands):
                 cmd_term = commands[index] + '\r\n'
                 self.send_command(command_to_send=cmd_term, from_gui=False)
                 # Schedule next command after a delay
                 self.root.after(500, lambda: send_next_command(index + 1))
        if commands: send_next_command(0)

    def copy_device_info(self):
        info_text = f"""BOLT Info:
Connection: {self.status_var.get()}
Start Time: {self.start_time_value.get()}
FSET: {self.fset_value.get()}
Loaded Item: {self.loaded_item_value.get()}
Slot: {self.slot_value.get()}
UnitID: {self.unit_id.get()}
Brick Num: {self.brick_number.get()}"""
        try:
            self.root.clipboard_clear()
            self.root.clipboard_append(info_text)
            messagebox.showinfo("Info Copied", "Device info copied to clipboard.")
        except Exception as e:
            messagebox.showerror("Clipboard Error", f"Could not copy info:\n{e}")

    def send_command_event(self, event=None):
        command = self.cmd_entry.get().strip()
        if command:
             cmd_term = command + '\r\n' # Assume commands need CR+LF
             self.send_command(command_to_send=cmd_term, from_gui=True)
             self.cmd_entry.delete(0, tk.END)

    def send_command(self, command_to_send, from_gui=True):
        """Sends a command string character by character with optional delay."""
        if not self.is_connected or not self.serial or not self.serial.is_open:
            if from_gui: messagebox.showwarning("Not Connected", "Please connect first.")
            else: self.add_message("Cannot send command: Not connected.", "error")
            return
        try:
            char_delay_sec = self.char_delay.get()
            bytes_sent = 0
            encoded_command = command_to_send.encode('ascii', errors='replace') # Encode whole command once
            for byte_val in encoded_command:
                self.serial.write(bytes([byte_val]))
                bytes_sent += 1
                if char_delay_sec > 0:
                    time.sleep(char_delay_sec)
            current_tx = self.tx_count.get()
            self.tx_count.set(current_tx + bytes_sent)
            # Log sent command (optional, can be noisy)
            # self.add_message(f"Sent Cmd: {command_to_send.strip()}", "sent")
        except Exception as e:
            self.add_message(f"Send command error: {str(e)}", "error")

    def send_bytes(self, byte_sequence):
        """Sends a raw byte sequence to the serial port."""
        if not self.is_connected or not self.serial or not self.serial.is_open:
            self.add_message("Cannot send bytes: Not connected.", "error")
            return
        try:
            self.serial.write(byte_sequence)
            current_tx = self.tx_count.get()
            self.tx_count.set(current_tx + len(byte_sequence))
            # self.add_message(f"Sent Bytes: {repr(byte_sequence)}", "sent")
        except Exception as e:
            self.add_message(f"Send bytes error: {str(e)}", "error")

    # --- GUI Update/Utility Methods ---
    def clear_output(self):
        try:
             self.output_text.config(state=tk.NORMAL)
             self.output_text.delete(1.0, tk.END)
             self.output_text.config(state=tk.DISABLED)
        except Exception as e: print(f"Error clearing output: {e}")

    def save_log(self):
        try:
            default_filename = f"bolt_log_{datetime.now().strftime('%Y%m%d_%H%M%S')}.log"
            filename = filedialog.asksaveasfilename(
                defaultextension=".log",
                filetypes=[("Log files", "*.log"), ("Text files", "*.txt"), ("All files", "*.*")],
                title="Save Log As",
                initialfile=default_filename
            )
            if not filename: return # User cancelled
            with open(filename, 'w', encoding='utf-8') as f:
                f.write(self.output_text.get(1.0, tk.END))
            messagebox.showinfo("Log Saved", f"Log successfully saved to:\n{filename}")
        except Exception as e:
            messagebox.showerror("Error Saving Log", f"Failed to save log:\n{str(e)}")

    def add_message(self, message, message_type="info"):
        """Adds a message to the queue for display in the text area."""
        # Add timestamp here if desired, or handle in process_messages
        self.message_queue.put((message_type, str(message)))

    def process_messages(self):
        """Processes messages from the queue and updates the text area."""
        try:
            processed_count = 0
            # Process multiple messages per call for efficiency
            for _ in range(100): # Limit messages per update cycle
                if self.message_queue.empty(): break

                item = self.message_queue.get_nowait()
                if isinstance(item, tuple) and len(item) == 2:
                    message_type, message_content = item
                else:
                    print(f"Skipping badly formatted queue item: {item}")
                    continue

                processed_count += 1
                try: # Update GUI safely
                    self.output_text.config(state=tk.NORMAL)
                    # Add timestamp universally for log-style messages
                    if message_type != "char": # Don't add timestamp for every character
                        ts = self.get_timestamp()
                        self.output_text.insert(tk.END, f"\n[{ts}] ", "timestamp")

                    # Apply specific tags based on type
                    if message_type == "char":
                        self.output_text.insert(tk.END, message_content)
                    elif message_type == "special_char":
                         self.output_text.insert(tk.END, message_content, "special_char")
                    elif message_type == "sent":
                        self.output_text.insert(tk.END, message_content, "sent")
                    elif message_type == "error":
                         self.output_text.insert(tk.END, message_content, "error")
                    elif message_type == "info":
                         self.output_text.insert(tk.END, message_content, "info")
                    elif message_type == "command":
                         self.output_text.insert(tk.END, message_content, "command")
                    else: # Default/standard tag
                         self.output_text.insert(tk.END, message_content, "standard")

                    self.output_text.config(state=tk.DISABLED)
                except tk.TclError as e: # Handle potential errors if GUI is destroyed
                    print(f"GUI update error (TclError): {e}")
                    break
                except Exception as e:
                    print(f"General GUI update error: {e}")

            # Auto-scroll if messages were processed
            if processed_count > 0 and self.auto_scroll.get():
                try: self.output_text.see(tk.END)
                except tk.TclError: pass # Ignore error if widget destroyed

        except queue.Empty: pass # No messages to process
        except Exception as e: print(f"Error in process_messages loop: {e}")

        # Reschedule this method to run again
        if self.root.winfo_exists(): # Check if root window still exists
            self.root.after(20, self.process_messages)

    def get_timestamp(self):
        return datetime.now().strftime("%H:%M:%S.%f")[:-3]

    # --- Settings Persistence Methods ---
    def load_app_settings(self):
        """Loads settings like TEK file paths from the config file."""
        try:
            if os.path.exists(self.CONFIG_FILENAME):
                with open(self.CONFIG_FILENAME, 'r') as f:
                    settings = json.load(f)
                    ame_path = settings.get("ame_tek_path", "")
                    wfc_path = settings.get("wfc_tek_path", "")
                    # Use set() method for Tkinter StringVars
                    self.ame_tek_path.set(ame_path)
                    self.wfc_tek_path.set(wfc_path)
                    self.add_message("Loaded saved TEK file paths.", "info")
                    if ame_path: self.add_message(f"  AME: {os.path.basename(ame_path)}", "info")
                    if wfc_path: self.add_message(f"  WFC: {os.path.basename(wfc_path)}", "info")
            else:
                self.add_message("No config file found. Select TEK files via File menu.", "info")
        except Exception as e:
            self.add_message(f"Error loading settings: {e}", "error")

    def save_app_settings(self):
        """Saves current settings (like TEK paths) to the config file."""
        settings = {
            "ame_tek_path": self.ame_tek_path.get(),
            "wfc_tek_path": self.wfc_tek_path.get()
            # Add other settings to save here if needed
        }
        try:
            with open(self.CONFIG_FILENAME, 'w') as f:
                json.dump(settings, f, indent=4)
        except Exception as e:
            # Show error in GUI log instead of just printing
            self.add_message(f"Error saving settings to {self.CONFIG_FILENAME}: {e}", "error")

    def on_closing(self):
        """Handles the window closing event, saves settings, disconnects."""
        print("Closing application...")
        self.save_app_settings()
        self.disconnect()
        print("Disconnect complete. Destroying window.")
        self.root.destroy()

    # --- Configuration File Handling Methods ---
    def select_slot_xml(self, slot_index):
        filepath = filedialog.askopenfilename(
            title=f"Select Configuration XML for Slot {slot_index + 1}",
            filetypes=[("XML files", "*.xml"), ("All files", "*.*")]
        )
        if filepath:
             self.slot_xml_paths[slot_index].set(filepath)
             self.add_message(f"Selected Slot {slot_index + 1} XML: {os.path.basename(filepath)}", "info")

    def select_tek_file(self, tek_type):
        title = f"Select {tek_type.upper()} TEK File"
        initial_dir = "." # Default directory
        current_path = ""
        if tek_type == 'ame': current_path = self.ame_tek_path.get()
        elif tek_type == 'wfc': current_path = self.wfc_tek_path.get()

        if current_path and os.path.exists(os.path.dirname(current_path)):
             initial_dir = os.path.dirname(current_path)

        filepath = filedialog.askopenfilename(
            title=title,
            initialdir=initial_dir,
            filetypes=[("TEK files", "*.tek"), ("XML files", "*.xml"), ("All files", "*.*")]
        )
        if filepath:
            if tek_type == 'ame':
                 self.ame_tek_path.set(filepath)
                 self.add_message(f"Selected AME TEK File: {os.path.basename(filepath)}", "info")
            elif tek_type == 'wfc':
                 self.wfc_tek_path.set(filepath)
                 self.add_message(f"Selected WFC TEK File: {os.path.basename(filepath)}", "info")
            self.save_app_settings() # Save settings immediately after selection

    def parse_config_xml(self, filepath):
        """Parses the XML configuration file for a specific slot."""
        config_data = {}
        try:
            tree = ET.parse(filepath)
            root = tree.getroot()

            # Elements to extract text content
            for elem_name in ['name', 'Waveform', 'NeedsKey', 'passwordPrompt',
                              'password', 'prompt', 'TEKLoad', 'TEKFile',
                              'TEKOffset', 'KeyingID', 'WFType']: # Added WFType
                element = root.find(elem_name)
                if element is not None and element.text:
                    config_data[elem_name] = element.text.strip()

            # Elements with attributes
            brick_elem = root.find('BrickNumber')
            if brick_elem is not None:
                 config_data['BrickNumberPrompt'] = brick_elem.get('Prompt')
                 config_data['BrickNumberTrigger'] = brick_elem.get('Trigger')

            unitid_elem = root.find('UnitID')
            if unitid_elem is not None:
                 config_data['UnitIDPrompt'] = unitid_elem.get('Prompt')
                 config_data['UnitIDTrigger'] = unitid_elem.get('Trigger')

            # KeyOrder parsing
            keyorder_elem = root.find('KeyOrder')
            if keyorder_elem is not None and keyorder_elem.text:
                try:
                    config_data['KeyOrder'] = [int(x.strip()) for x in keyorder_elem.text.split(',') if x.strip()]
                except ValueError:
                     self.add_message(f"Error parsing KeyOrder in {os.path.basename(filepath)}.", "error")
                     # Decide if this is fatal for the config? Return None or continue?
                     # return None # Example: Treat as fatal

            # LightningWFChange parsing
            lwf_elem = root.find('LightningWFChange')
            if lwf_elem is not None and lwf_elem.text:
                 config_data['LightningWFChange'] = lwf_elem.text.strip().lower() in ['true', '1', 'yes']

            # Commands parsing
            commands_list = []
            commands_root = root.find('Commands')
            if commands_root is not None:
                for command_elem in commands_root.findall('Command'):
                    # Store command attributes as a dictionary
                    commands_list.append(dict(command_elem.attrib)) # More concise way
            config_data['Commands'] = commands_list

            return config_data

        except ET.ParseError as e:
             self.add_message(f"XML Parse Error in {os.path.basename(filepath)}: {e}", "error")
             return None
        except Exception as e:
            self.add_message(f"Error parsing {os.path.basename(filepath)}: {e}", "error")
            return None

    def find_tek_keys(self, tek_filepath, device_id):
        """Finds the TEK_1 key for a given device ID in a TEK XML file."""
        # This function currently only finds TEK_1. Modify if multiple keys needed.
        tek_key_1 = None
        if not device_id or device_id == "--":
            self.add_message("Cannot find TEK: Brick Number missing.", "error")
            return None # Return None on failure

        # Ensure TEK file path exists before trying to parse
        if not tek_filepath:
             self.add_message(f"TEK file path is missing.", "error")
             return None
        if not os.path.exists(tek_filepath):
             # Log full path for debugging if it doesn't exist
             self.add_message(f"TEK file not found at path: {tek_filepath}", "error")
             return None

        try:
            tree = ET.parse(tek_filepath)
            root = tree.getroot()
            # Use XPath to find the Device element with the matching ID
            # Note: XPath syntax might vary slightly depending on XML structure
            device_element = root.find(f".//Device[ID='{device_id}']")

            if device_element is None:
                self.add_message(f"Device ID '{device_id}' not found in TEK file: {os.path.basename(tek_filepath)}", "error")
                return None # Return None if device ID not found

            tek1_element = device_element.find('TEK_1') # Assuming TEK_1 is direct child
            if tek1_element is not None and tek1_element.text:
                tek_key_1 = tek1_element.text.strip()
                # Add check for empty key?
                if not tek_key_1:
                     self.add_message(f"TEK_1 found but is empty for ID '{device_id}'.", "warning")
                     return None # Treat empty key as failure? Or allow? Returning None for now.
            else:
                self.add_message(f"TEK_1 element not found or empty for ID '{device_id}'.", "warning")
                return None # Return None if TEK_1 not found

            return tek_key_1

        except ET.ParseError as e:
             self.add_message(f"XML Parse Error reading TEK file {os.path.basename(tek_filepath)}: {e}", "error")
             return None
        except Exception as e:
            self.add_message(f"Error reading TEK file {os.path.basename(tek_filepath)}: {e}", "error")
            return None

    # --- Main Configuration Logic ---
    def start_full_configuration(self):
        """Initiates the multi-slot configuration process."""
        self.add_message("--- Starting Full Configuration Process ---", "info")

        # 1. Check Prerequisites
        if not self.is_bolt_connected:
            messagebox.showerror("Error", "BOLT device not connected. Please connect first.")
            return
        brick_num = self.brick_number.get()
        if not brick_num or brick_num == "--":
             messagebox.showerror("Error", "Brick Number not available. Cannot proceed with TEK loading.")
             return

        # 2. Prepare List of Slots to Configure
        first_slot_index = -1
        slots_to_configure = []
        for i in range(4):
            slot_path = self.slot_xml_paths[i].get()
            if slot_path:
                if not os.path.exists(slot_path):
                     self.add_message(f"Warning: Slot {i+1} XML file not found at '{slot_path}'. Skipping.", "warning")
                     continue # Skip if file doesn't exist

                if first_slot_index == -1: first_slot_index = i
                self.add_message(f"Parsing configuration for Slot {i+1} from {os.path.basename(slot_path)}...", "info")
                slot_config_data = self.parse_config_xml(slot_path)
                if slot_config_data:
                    slots_to_configure.append({'index': i, 'path': slot_path, 'data': slot_config_data})
                else:
                    # parse_config_xml already logs the error
                    self.add_message(f"Failed to parse Slot {i+1} XML. Configuration for this slot will be skipped.", "error")
                    # Optionally show a warning messagebox here too?

        # 3. Validate Slot List and Order
        if not slots_to_configure or first_slot_index == -1:
            messagebox.showerror("Error", "No valid Slot XML files loaded or parsed correctly.")
            return

        # Configure slots starting from the first one loaded
        ordered_slots_to_configure = [s for s in slots_to_configure if s['index'] >= first_slot_index]
        if not ordered_slots_to_configure: # Should not happen if slots_to_configure is not empty
            messagebox.showerror("Error", "Internal logic error: No slots found to configure.")
            return

        self.add_message(f"Starting configuration sequence from Slot {first_slot_index + 1}.", "info")

        # 4. Check Required TEK Files are Selected
        required_tek_files = set()
        for slot in ordered_slots_to_configure:
             # Ensure 'data' exists (should always if parse was successful)
             if 'data' in slot and slot['data']:
                 tek_file = slot['data'].get('TEKFile', '').lower()
                 if tek_file:
                     required_tek_files.add(tek_file)

        missing_tek = False
        # Check AME TEK
        if "ame.tek" in required_tek_files:
             ame_path = self.ame_tek_path.get()
             if not ame_path:
                 messagebox.showerror("Error", "AME TEK file needed (required by XML) but not selected.")
                 missing_tek = True
             elif not os.path.exists(ame_path):
                  messagebox.showerror("Error", f"Selected AME TEK file not found at:\n{ame_path}")
                  missing_tek = True

        # Check WFC TEK (if needed)
        if "wfc.tek" in required_tek_files:
             wfc_path = self.wfc_tek_path.get()
             if not wfc_path:
                 messagebox.showerror("Error", "WFC TEK file needed (required by XML) but not selected.")
                 missing_tek = True
             elif not os.path.exists(wfc_path):
                  messagebox.showerror("Error", f"Selected WFC TEK file not found at:\n{wfc_path}")
                  missing_tek = True

        if missing_tek:
            self.add_message("Configuration stopped due to missing or invalid TEK file paths.", "error")
            return # Stop if required TEK files are missing or not found

        # 5. Start Configuration Thread
        self.add_message("Starting configuration thread...", "info") # Log before starting
        config_thread = threading.Thread(
             target=self._run_config_thread,
             args=(ordered_slots_to_configure, brick_num),
             daemon=True
         )
        config_thread.start()
        self.add_message("Configuration thread started.", "info") # Log after starting


    def _run_config_thread(self, slots_to_configure, brick_num):
        """Worker thread executing config sequence, waiting for specific prompts."""
        print(f"DEBUG: _run_config_thread entered. Received {len(slots_to_configure)} slots to configure.")
        try:
            if not self.root.winfo_exists():
                 print("DEBUG: Root window closed, exiting config thread early.")
                 return

            # Disable button via main thread
            self.root.after(0, lambda: self.full_config_button.config(state=tk.DISABLED))
            print("DEBUG: Config button disable scheduled.")

            for slot_info in slots_to_configure:
                if self.stop_thread.is_set():
                    self.add_message("Config stopped by user.", "warning")
                    break # Exit the loop if stopped

                slot_index = slot_info['index']
                slot_path = slot_info['path']
                config_data = slot_info['data']
                slot_num = slot_index + 1
                self.add_message(f"--- Configuring Slot {slot_num} ---", "info")
                print(f"DEBUG: Starting loop for slot index {slot_index}")

                # Get expected prompts from config for this slot
                expected_prompt = config_data.get('prompt')
                needs_key_prompt = config_data.get('NeedsKey')
                if not expected_prompt:
                    self.add_message(f"Config Error for Slot {slot_num}: Missing <prompt> tag in XML. Skipping slot.", "error")
                    continue # Skip this slot


                # --- Step 1: Handle TEK Key (Conditional) ---
                proceed_after_tek = False # Default to failure/skip
                tek_filename_from_config = config_data.get('TEKFile')
                tek_load_cmd_base = config_data.get('TEKLoad')

                if tek_filename_from_config and tek_load_cmd_base:
                    # TEK is required by this XML

                    # Determine actual TEK file path (AME or WFC)
                    tek_filepath = ""
                    tek_file_lower = tek_filename_from_config.lower()
                    if tek_file_lower == "ame.tek":
                         tek_filepath = self.ame_tek_path.get()
                    elif tek_file_lower == "wfc.tek":
                         tek_filepath = self.wfc_tek_path.get()
                    else:
                         # Handle potential relative paths if needed, assuming absolute for now
                         config_dir = os.path.dirname(slot_path)
                         tek_filepath = os.path.join(config_dir, tek_filename_from_config)
                         self.add_message(f"Assuming TEK relative path (ensure file exists): {tek_filepath}", "info")

                    # find_tek_keys already checks existence, but double check path validity
                    if not tek_filepath: # Should have been caught by start_full_config, but check again
                         self.add_message(f"Internal Error: TEK path for {tek_filename_from_config} is empty. Skipping TEK.", "error")
                         continue # Skip slot

                    # Now, determine if we need to wait for 'NeedsKey' prompt or just the main prompt
                    if needs_key_prompt:
                         # We might get 'NeedsKey' OR the main prompt if key is already loaded
                         self.add_message(f"Waiting for prompt ('{expected_prompt}') or NeedsKey ('{needs_key_prompt}')...", "info")
                         self.currently_expected_prompt = expected_prompt
                         self.currently_needs_key_prompt = needs_key_prompt
                         self.main_prompt_event.clear()
                         self.needs_key_event.clear()

                         wait_start = time.time()
                         tek_timeout = 15 # Timeout for this initial wait
                         key_needed = False
                         prompt_received_instead = False

                         while time.time() - wait_start < tek_timeout:
                            if self.stop_thread.is_set(): self.add_message("Stop requested.", "warning"); return
                            if self.needs_key_event.wait(timeout=0.1): key_needed = True; break
                            if self.main_prompt_event.wait(timeout=0.1): prompt_received_instead = True; break

                         # Clear expectations now that wait is over
                         self.currently_expected_prompt = None
                         self.currently_needs_key_prompt = None

                         if key_needed:
                            self.add_message(f"'{needs_key_prompt}' received. Proceeding with TEK load.", "info")
                            tek_key_1 = self.find_tek_keys(tek_filepath, brick_num)
                            if not tek_key_1:
                                 self.add_message(f"Stopping Slot {slot_num}: TEK_1 key not found or is invalid.", "error")
                                 continue # Skip slot

                            self.add_message(f"Sending TEK_1 for Slot {slot_num}...", "info")
                            tek_command = f"{tek_load_cmd_base} {tek_key_1}"
                            # Log only part of key for security if needed
                            self.add_message(f"Sending Key Cmd: {tek_load_cmd_base} ...", "info")
                            self.send_command(tek_command + '\r\n', from_gui=False)

                            # Wait specifically for the main prompt after sending TEK
                            if self.wait_for_specific_prompt(expected_prompt, timeout_sec=10):
                                 proceed_after_tek = True # Success!
                            else:
                                 self.add_message("Timeout or error waiting for prompt AFTER sending TEK. Stopping slot.", "error")
                                 continue # Skip slot
                         elif prompt_received_instead:
                              self.add_message("Main prompt received before NeedsKey. Assuming key already loaded or not needed. Proceeding.", "info")
                              proceed_after_tek = True # Assume okay to proceed
                         else: # Timeout occurred waiting for either prompt
                              self.add_message(f"Timeout waiting for NeedsKey or Main Prompt for Slot {slot_num}. Stopping slot.", "error")
                              continue # Skip slot
                    else:
                         # No 'NeedsKey' prompt defined, just wait for the main prompt initially
                         self.add_message("No NeedsKey prompt defined. Waiting for initial main prompt...", "info")
                         if self.wait_for_specific_prompt(expected_prompt, timeout_sec=15):
                            # Now send the TEK key unconditionally since no NeedsKey check
                            tek_key_1 = self.find_tek_keys(tek_filepath, brick_num)
                            if not tek_key_1:
                                 self.add_message(f"Stopping Slot {slot_num}: TEK_1 key not found or is invalid.", "error")
                                 continue # Skip slot

                            self.add_message(f"Sending TEK_1 for Slot {slot_num} (unconditional)...", "info")
                            tek_command = f"{tek_load_cmd_base} {tek_key_1}"
                            self.add_message(f"Sending Key Cmd: {tek_load_cmd_base} ...", "info")
                            self.send_command(tek_command + '\r\n', from_gui=False)

                            # Wait for prompt again after sending key
                            if self.wait_for_specific_prompt(expected_prompt, timeout_sec=10):
                                 proceed_after_tek = True # Success!
                            else:
                                 self.add_message("Timeout or error waiting for prompt AFTER sending TEK. Stopping slot.", "error")
                                 continue # Skip slot
                         else:
                             self.add_message("Initial main prompt not received. Stopping slot.", "error")
                             continue # Skip slot
                else:
                    # No TEKFile or TEKLoad defined in XML, skip TEK process
                    self.add_message(f"No TEKFile/TEKLoad in XML for Slot {slot_num}. Skipping TEK step.", "info")
                    # Need to wait for initial prompt if not doing TEK
                    if self.wait_for_specific_prompt(expected_prompt, timeout_sec=15):
                         proceed_after_tek = True # Got initial prompt, okay to proceed
                    else:
                         self.add_message("Initial main prompt not received. Stopping slot.", "error")
                         continue # Skip slot

                # --- Check if TEK step succeeded ---
                if not proceed_after_tek:
                     # Error messages already logged inside the logic above
                     self.add_message(f"Configuration for Slot {slot_num} aborted due to TEK/Initial Prompt failure.", "error")
                     continue # Move to the next slot

                # --- Step 2: Send 'amereset' (Optional) ---
                # UNCOMMENT AND MODIFY IF NEEDED for your specific device workflow
                # self.add_message(f"Sending amereset for Slot {slot_num}...", "info")
                # self.send_command("amereset\r\n", from_gui=False)
                # if not self.wait_for_specific_prompt(expected_prompt, timeout_sec=5): # Check prompt quickly?
                #      self.add_message(f"Warning: No prompt received shortly after amereset.", "warning")


                # --- Step 3: Wait for reconnect/re-login (Optional, only if Step 2 causes it) ---
                # UNCOMMENT AND MODIFY IF NEEDED
                # needs_reconnect_wait = False # Set to True if amereset causes disconnect
                # if needs_reconnect_wait:
                #     self.add_message("Waiting for device reset & reconnect...", "info")
                #     self.is_bolt_connected = False # Assume disconnect
                #     # Update status visually (optional)
                #     if self.root.winfo_exists():
                #          self.root.after(0, lambda: self.status_var.set("Resetting..."))
                #          self.root.after(0, lambda: self.status_label.config(foreground="orange"))
                #
                #     time.sleep(10) # Initial sleep after reset command
                #     reconnect_wait_start = time.time()
                #     reconnect_timeout = 45 # seconds
                #     reconnected = False
                #     while time.time() - reconnect_wait_start < reconnect_timeout:
                #         if self.stop_thread.is_set(): self.add_message("Stop requested.", "warning"); return
                #         # Attempt BOLT connection via connect_to_bolt (which checks is_bolt_connected)
                #         if self.root.winfo_exists(): self.root.after(0, self.connect_to_bolt)
                #         time.sleep(1) # Check connection status periodically
                #         if self.is_bolt_connected: # connect_to_bolt sets this
                #             reconnected = True
                #             break
                #         time.sleep(1.5) # Wait a bit longer between checks
                #
                #     if not reconnected:
                #          self.add_message("Timeout waiting for device to reconnect after reset. Stopping slot.", "error")
                #          continue # Skip slot
                #
                #     self.add_message("Device reconnected to BOLT.", "info")
                #     # Wait for the main prompt after reconnecting
                #     if not self.wait_for_specific_prompt(expected_prompt, timeout_sec=15):
                #         self.add_message("Prompt not received after reconnect. Stopping config for slot.", "error")
                #         continue # Skip slot
                # else:
                #      self.add_message("Skipping reconnect wait (amereset not configured).", "info")


                # --- Send Dynamic WFI Command (AFTER successful Step 1/3) ---
                self.add_message(f"Attempting to send WFI command for Slot {slot_num}...", "info")
                wf_type_string = config_data.get('WFType')
                current_brick_num = self.brick_number.get()
                current_unit_id = self.unit_id.get()

                if wf_type_string and current_brick_num != "--" and current_unit_id != "--":
                    wf_value = -1
                    # Map WFType (case-insensitive)
                    wf_type_upper = wf_type_string.upper()
                    if wf_type_upper == 'LYNX-6': wf_value = 2
                    elif wf_type_upper == 'LYNX-8': wf_value = 3
                    elif wf_type_upper == 'COPPER': wf_value = 4
                    # Add more mappings if needed

                    if wf_value != -1:
                        try:
                            current_slot_num_str = str(slot_num)
                            wf_value_str = str(wf_value)
                            cmd_bytes = (
                                b'\xA1C.WFI ' + current_slot_num_str.encode('ascii') + b' ' +
                                wf_value_str.encode('ascii') + b' ' + current_brick_num.encode('ascii') + b' ' +
                                current_unit_id.encode('ascii') + b' ' + wf_type_string.encode('ascii') + b'\xB6'
                            )
                            self.send_bytes(cmd_bytes)
                            self.add_message(f"Sent WFI Cmd: ¡C.WFI {current_slot_num_str} {wf_value_str} {current_brick_num} {current_unit_id} {wf_type_string}¶", "sent")

                            # Wait for prompt after WFI command
                            if not self.wait_for_specific_prompt(expected_prompt, timeout_sec=10):
                                 self.add_message(f"No prompt after WFI command. Continuing...", "warning")
                        except UnicodeEncodeError:
                             self.add_message(f"Error encoding components for WFI cmd (non-ASCII?).", "error")
                        except Exception as e:
                             self.add_message(f"Error sending WFI command: {e}", "error")
                    else:
                        self.add_message(f"Skipping WFI command: Unknown WFType '{wf_type_string}' found in XML.", "warning")
                elif not wf_type_string:
                     self.add_message(f"Skipping WFI command: <WFType> tag missing or empty in XML.", "info")
                else:
                     self.add_message(f"Skipping WFI command: Brick Number or Unit ID not available.", "warning")


                # --- Step 4: Send Commands from <Commands> block ---
                commands_to_send = config_data.get('Commands', [])
                if commands_to_send:
                    self.add_message(f"Sending {len(commands_to_send)} config commands from XML...", "info")
                    all_commands_sent_ok = True
                    for i, cmd_info in enumerate(commands_to_send):
                        if self.stop_thread.is_set(): self.add_message("Stop requested.", "warning"); return

                        cmd_type = cmd_info.get('Type') # 'Type' holds the command string
                        # Check NumArguments; only send if 0 for now (simplest case)
                        if cmd_type and cmd_info.get('NumArguments', '0') == '0':
                            self.add_message(f"Sending Cmd {i+1}: {cmd_type}", "info")
                            self.send_command(cmd_type + '\r\n', from_gui=False)

                            # Wait for the expected prompt after each command
                            if not self.wait_for_specific_prompt(expected_prompt, timeout_sec=5): # Shorter timeout for sequential commands
                                self.add_message(f"No prompt or error after '{cmd_type}'. Stopping commands for this slot.", "error")
                                all_commands_sent_ok = False
                                break # Stop sending commands for this slot
                        elif cmd_type:
                             self.add_message(f"Skipping XML cmd '{cmd_type}' (NumArguments != 0 or missing).", "warning")
                        else:
                             self.add_message(f"Skipping invalid XML cmd entry {i+1} (missing Type).", "warning")

                    if all_commands_sent_ok:
                         self.add_message(f"Finished sending config commands from XML.", "info")
                    else:
                         self.add_message(f"Aborted sending XML commands due to error for Slot {slot_num}.", "error")
                         continue # Skip to next slot if commands failed

                else:
                    self.add_message(f"No commands found in <Commands> block for Slot {slot_num}.", "info")


                # --- Step 5 & 6: Determine and Send wfchange ---
                is_last_slot_in_sequence = (slot_info == slots_to_configure[-1])
                wfchange_cmd_str = None # Reset for this slot

                if not is_last_slot_in_sequence:
                    # Intermediate Slot: Change to next sequential slot
                    self.add_message("Determining current slot post-config...", "info")
                    time.sleep(1) # Brief pause to allow potential GUI updates
                    current_slot = -1
                    try:
                        current_slot_str = self.slot_value.get()
                        if current_slot_str.isdigit(): current_slot = int(current_slot_str)
                        else: self.add_message(f"Warning: Invalid current slot value '{current_slot_str}'.", "warning")
                    except Exception as e: self.add_message(f"Error reading current slot: {e}", "error")

                    if current_slot != -1:
                        wfchange_target_slot = (current_slot % 4) + 1
                        wfchange_cmd_str = f"wfchange {wfchange_target_slot}"
                        self.add_message(f"Preparing to change from slot {current_slot} to next slot: {wfchange_target_slot}", "info")
                    else:
                        self.add_message("Skipping wfchange to next slot (unknown current slot).", "warning")

                else:
                    # Last Slot: Change back to Slot 1 (if not already there)
                    self.add_message(f"Finished last configured slot ({slot_num}). Checking slot...", "info")
                    current_slot = -1
                    try:
                         current_slot_str = self.slot_value.get()
                         if current_slot_str.isdigit(): current_slot = int(current_slot_str)
                    except Exception: pass

                    if current_slot != 1:
                        wfchange_cmd_str = "wfchange 1"
                        self.add_message("Preparing to change back to slot 1.", "info")
                    else:
                        self.add_message("Device already in slot 1. No wfchange needed.", "info")

                # Send the determined wfchange command (if any)
                if wfchange_cmd_str:
                    self.add_message(f"Sending: {wfchange_cmd_str}", "info")
                    self.send_command(wfchange_cmd_str + '\r\n', from_gui=False)
                    # Optional: Wait for prompt after wfchange
                    # Note: The prompt might change AFTER wfchange. Using current 'expected_prompt' might be wrong.
                    # Consider waiting for a generic prompt or skipping the wait here.
                    # if not self.wait_for_specific_prompt(expected_prompt, timeout_sec=15):
                    #     self.add_message(f"Warning: No prompt or error after {wfchange_cmd_str}.", "warning")


                # --- Finished Slot ---
                self.add_message(f"--- Finished Slot {slot_num} Processing ---", "info")
                if not is_last_slot_in_sequence:
                    time.sleep(1) # Small delay before starting next slot

            # --- End of the main for loop for slots ---
            self.add_message("--- Full Configuration Process Completed ---", "info")

        except Exception as e:
            self.add_message(f"Config thread encountered an unhandled error: {e}", "error")
            self.add_message(traceback.format_exc(), "error") # Log full traceback

        finally:
            # Ensure button is re-enabled in the main thread
            if self.root.winfo_exists():
                self.root.after(0, lambda: self.full_config_button.config(state=tk.NORMAL))
            print("DEBUG: Config thread finished.")


    def wait_for_specific_prompt(self, prompt_string, timeout_sec):
        """Waits for main_prompt_event (triggered by matching prompt_string)."""
        if not prompt_string:
            self.add_message("Wait error: No prompt string specified.", "error")
            return False

        self.add_message(f"Waiting for prompt '{prompt_string}' (timeout: {timeout_sec}s)...", "info")
        self.currently_expected_prompt = prompt_string # Tell reader thread
        self.main_prompt_event.clear()
        wait_start_time = time.time()
        prompt_ok = False

        while time.time() - wait_start_time < timeout_sec:
            if self.stop_thread.is_set():
                self.add_message("Stop requested while waiting for prompt.", "warning")
                self.currently_expected_prompt = None
                return False
            # Wait for event with a short timeout for responsiveness
            if self.main_prompt_event.wait(timeout=0.2): # Check event more frequently
                # self.add_message("Prompt received.", "info") # Less verbose
                prompt_ok = True
                break # Exit wait loop

        self.currently_expected_prompt = None # Clear expectation
        if not prompt_ok:
            self.add_message(f"Timeout waiting for prompt: '{prompt_string}' ({timeout_sec}s).", "error")
        return prompt_ok

    # --- Other Methods (unchanged assuming correct) ---
    # set_command_entry
    # toggle_connection (uses connect/disconnect)
    # request_device_info
    # copy_device_info
    # send_command_event (uses send_command)
    # clear_output
    # save_log
    # get_timestamp
    # load_app_settings
    # save_app_settings
    # on_closing (uses save_app_settings, disconnect)
    # select_slot_xml
    # select_tek_file (uses save_app_settings)

# --- Main execution block ---
def main():
    root = tk.Tk()
    app = BoltTerminalGUI(root)
    root.mainloop()

if __name__ == "__main__":
    main()