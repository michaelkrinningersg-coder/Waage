import serial
import serial.tools.list_ports
import pyautogui
import tkinter as tk
from tkinter import ttk
import threading
import re

class BalanceWedgeApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Waagen PC-Direct Wedge Pro")
        self.root.geometry("500x750") # Fenster vergrößert für den Monitor
        
        self.serial_port = None
        self.is_running = False
        self.setup_ui()
        self.update_ports()
        
    def setup_ui(self):
        # --- Schnittstellen-Einstellungen ---
        tk.Label(self.root, text="--- Schnittstelle (RS232/USB) ---", font=("Arial", 10, "bold")).pack(pady=(10,5))
        
        port_frame = tk.Frame(self.root)
        port_frame.pack(padx=20, fill="x")

        tk.Label(port_frame, text="COM Port:").pack(side="left", padx=5)
        self.port_var = tk.StringVar()
        self.port_combo = ttk.Combobox(port_frame, textvariable=self.port_var, width=20)
        self.port_combo.pack(side="left", padx=5)
        
        self.refresh_btn = tk.Button(port_frame, text="?", command=self.update_ports, width=3)
        self.refresh_btn.pack(side="left", padx=5)
        
        settings_frame = tk.Frame(self.root)
        settings_frame.pack(padx=20, pady=10)

        # Baudrate
        tk.Label(settings_frame, text="Baudrate:").grid(row=0, column=0, sticky="w", pady=2)
        self.baud_var = tk.StringVar(value="9600")
        ttk.Combobox(settings_frame, textvariable=self.baud_var, values=["1200", "2400", "4800", "9600", "19200", "38400", "115200"], width=25).grid(row=0, column=1, pady=2)

        # Datenbits / Parität
        tk.Label(settings_frame, text="Datenbits:").grid(row=1, column=0, sticky="w", pady=2)
        self.databits_var = tk.StringVar(value="8")
        ttk.Combobox(settings_frame, textvariable=self.databits_var, values=["7", "8"], width=25).grid(row=1, column=1, pady=2)

        tk.Label(settings_frame, text="Parität:").grid(row=2, column=0, sticky="w", pady=2)
        self.parity_var = tk.StringVar(value="None")
        ttk.Combobox(settings_frame, textvariable=self.parity_var, values=["None", "Odd", "Even"], width=25).grid(row=2, column=1, pady=2)

        # --- Datenverarbeitung ---
        tk.Label(self.root, text="--- Datenverarbeitung ---", font=("Arial", 10, "bold")).pack(pady=(15,5))
        
        process_frame = tk.Frame(self.root)
        process_frame.pack(padx=20, anchor="w")

        self.replace_comma_var = tk.BooleanVar(value=True)
        tk.Checkbutton(process_frame, text="Punkt '.' in Komma ',' umwandeln", variable=self.replace_comma_var).pack(anchor="w")

        self.strip_unit_var = tk.BooleanVar(value=True)
        tk.Checkbutton(process_frame, text="Einheit entfernen (nur Zahlenwert)", variable=self.strip_unit_var).pack(anchor="w")

        # --- Excel-Steuerung ---
        tk.Label(self.root, text="--- Excel-Steuerung ---", font=("Arial", 10, "bold")).pack(pady=(15,5))
        
        excel_frame = tk.Frame(self.root)
        excel_frame.pack(padx=20, anchor="w")

        self.direction_var = tk.StringVar(value="enter")
        tk.Radiobutton(excel_frame, text="Nächste Zeile (Enter)", variable=self.direction_var, value="enter").pack(anchor="w")
        tk.Radiobutton(excel_frame, text="Nächste Spalte (Tab)", variable=self.direction_var, value="tab").pack(anchor="w")
        
        # --- Start/Stop ---
        self.toggle_btn = tk.Button(self.root, text="START (Verbinden)", command=self.toggle_connection, bg="lightgreen", font=("Arial", 12, "bold"), height=2)
        self.toggle_btn.pack(pady=15, fill="x", padx=40)
        
        self.status_label = tk.Label(self.root, text="Status: Nicht verbunden", fg="grey")
        self.status_label.pack()

        # --- Rohdaten Monitor ---
        tk.Label(self.root, text="--- Rohdaten-Monitor ---", font=("Arial", 10, "bold")).pack(pady=(10,5))
        
        monitor_frame = tk.Frame(self.root)
        monitor_frame.pack(padx=20, pady=(0, 10), fill="both", expand=True)
        
        # Textfeld mit Scrollbar
        self.monitor_text = tk.Text(monitor_frame, height=5, width=50, bg="#f4f4f4", state="disabled", font=("Courier", 9))
        self.monitor_text.pack(side="left", fill="both", expand=True)
        
        scrollbar = tk.Scrollbar(monitor_frame, command=self.monitor_text.yview)
        scrollbar.pack(side="right", fill="y")
        self.monitor_text.config(yscrollcommand=scrollbar.set)

    def log_to_monitor(self, text):
        """Schreibt Text sicher in das Monitor-Fenster und scrollt nach unten."""
        self.monitor_text.config(state="normal")
        self.monitor_text.insert("end", text + "\n")
        self.monitor_text.see("end")
        self.monitor_text.config(state="disabled")

    def update_ports(self):
        ports = [port.device for port in serial.tools.list_ports.comports()]
        self.port_combo['values'] = ports
        if ports:
            if self.port_var.get() not in ports:
                self.port_var.set(ports[0])
            self.status_label.config(text=f"{len(ports)} Port(s) gefunden.", fg="blue")
        else:
            self.port_var.set("")
            self.status_label.config(text="Kein COM Port gefunden!", fg="red")

    def toggle_connection(self):
        if not self.is_running:
            self.start_reading()
        else:
            self.stop_reading()

    def start_reading(self):
        port = self.port_var.get()
        if not port:
            self.status_label.config(text="Fehler: Bitte COM Port wählen!", fg="red")
            return

        parity_map = {"None": serial.PARITY_NONE, "Odd": serial.PARITY_ODD, "Even": serial.PARITY_EVEN}
        bytesize_map = {"7": serial.SEVENBITS, "8": serial.EIGHTBITS}
        
        try:
            self.serial_port = serial.Serial(
                port=port,
                baudrate=int(self.baud_var.get()),
                bytesize=bytesize_map[self.databits_var.get()],
                parity=parity_map[self.parity_var.get()],
                stopbits=serial.STOPBITS_ONE,
                timeout=1
            )
            self.is_running = True
            self.toggle_btn.config(text="STOP (Trennen)", bg="salmon")
            self.status_label.config(text=f"Verbunden mit {port}.", fg="green")
            
            # Hinweis im Monitor ausgeben
            self.log_to_monitor(f"--- Verbunden auf {port} ---")
            
            self.read_thread = threading.Thread(target=self.read_from_port, daemon=True)
            self.read_thread.start()
        except Exception as e:
            self.status_label.config(text=f"Verbindungsfehler: {e}", fg="red")

    def stop_reading(self):
        self.is_running = False
        if self.serial_port and self.serial_port.is_open:
            self.serial_port.close()
        self.toggle_btn.config(text="START (Verbinden)", bg="lightgreen")
        self.status_label.config(text="Verbindung getrennt.", fg="grey")
        self.log_to_monitor("--- Verbindung getrennt ---")

    def read_from_port(self):
        while self.is_running:
            try:
                if self.serial_port.in_waiting > 0:
                    # 1. Die rohen Bytes auslesen
                    raw_bytes = self.serial_port.readline()
                    
                    if raw_bytes:
                        # 2. Im Monitor exakt anzeigen, was empfangen wurde (als Debug-String)
                        raw_debug_str = repr(raw_bytes)
                        self.root.after(0, self.log_to_monitor, f"Empfangen: {raw_debug_str}")
                        
                        # 3. Das Signal für Excel aufbereiten
                        raw_data = raw_bytes.decode('ascii', errors='ignore').strip()
                        processed_data = raw_data
                        
                        if self.strip_unit_var.get():
                            processed_data = re.sub(r'[^0-9.,-]', '', processed_data)
                        if self.replace_comma_var.get():
                            processed_data = processed_data.replace('.', ',')
                        
                        # 4. In Excel eintragen
                        pyautogui.write(processed_data)
                        
                        if self.direction_var.get() == "enter":
                            pyautogui.press('enter')
                        else:
                            pyautogui.press('tab')
            except Exception as e:
                pass

if __name__ == "__main__":
    root = tk.Tk()
    app = BalanceWedgeApp(root)
    root.mainloop()