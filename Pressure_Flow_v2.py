# Version 2.0
# Dual Alicat Pressure Flow Test
# Test catheter for pressure decay and flow rate using two Alicat devices
# Setup:
# - Alicat A: Sets input pressure and measures pressure decay
# - Alicat B: Measures flow rate, acts as valve control
# - Both Alicats connected via serial port at 38400 baud via BB9
# - Catheter connected to test circuit
# 
# Test Process:
# Flow Test Phase:
#   1. Release valve on Alicat A
#   2. Set Alicat A to flow test pressure
#   3. Set Alicat B to flow test pressure
#   4. Wait for stabilization
#   5. Record mass flow from both devices
# Pressure Decay Test Phase:
#   1. Set Alicat B pressure to decay test pressure (closes valve)
#   2. Wait for Alicat A pressure to stabilize
#   3. Close valve on Alicat A
#   4. Record pressure decay over time

import tkinter as tk
from tkinter import messagebox
import serial
import time
import datetime
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
import os
from openpyxl import Workbook

# Change path name for your box folder
path = r"C:\Users\patri\RnD\SW Test Data"

class DualAlicatTestApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Catheter Pressure Flow Test v1.0")
        
        # Initialize serial connection
        try:
            self.ser = serial.Serial('COM23', 38400, timeout=1)  #  Connect throug BB9.
        except serial.SerialException as e:
            messagebox.showerror("Serial Error", f"Could not open serial port: {e}")
            self.root.quit()
            return
        
        # Initialize test parameters - Set pressures using absolute pressure (PSI)
        self.ambient_pressure = 14.7  # PSI
        self.a_flow_test_pressure = self.ambient_pressure + 10.0  # PSI
        self.b_flow_test_pressure = 0.0
        self.a_decay_test_pressure = 0.0
        self.b_decay_test_pressure = 0.0
        self.flow_sample_time = 5.0 # seconds
        self.pressure_sample_time = 20.0
        self.read_rate = .25
        self.pressure_read_rate = 1.0
        self.pressurize_time = 10.0
        
        # Read configuration from ini file
        self.read_ini()
        
        # Data storage
        self.time_data = []
        self.pressure_a_data = []
        self.pressure_b_data = []
        self.flow_a_data = []
        self.flow_b_data = []
        
        # Excel workbook
        self.workbook = None
        self.settings_sheet = None
        self.data_sheet = None
        self.excel_path = None
        
        # Build GUI
        self.build_gui()
        
    def read_ini(self):
        """Read settings from Test.ini file"""
        try:
            with open("Test.ini", "r") as file:
                for line in file:
                    line = line.strip()
                    if not line or line.startswith('#'):
                        continue
                    if "=" in line:
                        key, value = line.split('=', 1)
                        key = key.strip()
                        value = value.strip()
                        
                        if key == "A_FLOW_TEST_PRESSURE":
                            self.a_flow_test_pressure = float(value)
                        elif key == "B_FLOW_TEST_PRESSURE":
                            self.b_flow_test_pressure = float(value)
                        elif key == "A_DECAY_TEST_PRESSURE":
                            self.a_decay_test_pressure = float(value)
                        elif key == "B_DECAY_TEST_PRESSURE":
                            self.b_decay_test_pressure = float(value)
                        elif key == "FLOW_SAMPLE_TIME":
                            self.flow_sample_time = float(value)
                        elif key == "PRESSURE_SAMPLE_TIME":
                            self.pressure_sample_time = float(value)
                        elif key == "READ_RATE":
                            self.read_rate = float(value)
                        elif key == "PRESSURE_READ_RATE":
                            self.pressure_read_rate = float(value)
                        elif key == "PRESSURIZE_TIME":
                            self.pressurize_time = float(value)
        except FileNotFoundError:
            messagebox.showwarning("Warning", "Test.ini file not found! Using default values.")
            self.create_default_ini()
    
    def create_default_ini(self):
        """Create a default Test.ini file"""
        with open("Test.ini", "w") as file:
            file.write("# Dual Alicat Test Configuration\n")
            file.write("# Pressures in PSI\n")
            file.write("# Times in seconds\n\n")
            file.write("A_FLOW_TEST_PRESSURE=10.0\n")
            file.write("B_FLOW_TEST_PRESSURE=0.0\n")
            file.write("A_DECAY_TEST_PRESSURE=0.0\n")
            file.write("B_DECAY_TEST_PRESSURE=0.0\n")
            file.write("FLOW_SAMPLE_TIME=5.0\n")
            file.write("PRESSURE_SAMPLE_TIME=20.0\n")
            file.write("READ_RATE=1.0\n")
            file.write("PRESSURE_READ_RATE=1.0\n")
            file.write("PRESSURIZE_TIME=10.0\n")
    
    def build_gui(self):
        """Build the GUI interface"""
        # Part Number Entry
        tk.Label(self.root, text="Part Number:").grid(row=0, column=0, padx=5, pady=5, sticky='e')
        self.part_number = tk.StringVar()
        tk.Entry(self.root, textvariable=self.part_number, width=30).grid(row=0, column=1, padx=5, pady=5)
        
        # Test Parameters Display
        params_frame = tk.LabelFrame(self.root, text="Test Parameters", padx=10, pady=10)
        params_frame.grid(row=1, column=0, columnspan=2, padx=10, pady=10, sticky='ew')
        
        tk.Label(params_frame, text=f"A Flow Test Pressure: {self.a_flow_test_pressure} PSI").grid(row=0, column=0, sticky='w')
        tk.Label(params_frame, text=f"B Flow Test Pressure: {self.b_flow_test_pressure} PSI").grid(row=1, column=0, sticky='w')
        tk.Label(params_frame, text=f"B Decay Test Pressure: {self.b_decay_test_pressure} PSI").grid(row=2, column=0, sticky='w')
        tk.Label(params_frame, text=f"Flow Sample Time: {self.flow_sample_time} s").grid(row=0, column=1, padx=20, sticky='w')
        tk.Label(params_frame, text=f"Pressure Sample Time: {self.pressure_sample_time} s").grid(row=1, column=1, padx=20, sticky='w')
        tk.Label(params_frame, text=f"Read Rate: {self.read_rate} s").grid(row=2, column=1, padx=20, sticky='w')
        
        # Real-time Data Display
        data_frame = tk.LabelFrame(self.root, text="Real-time Data", padx=10, pady=10)
        data_frame.grid(row=2, column=0, columnspan=2, padx=10, pady=10, sticky='ew')
        
        tk.Label(data_frame, text="Alicat A Pressure:").grid(row=0, column=0, sticky='e')
        self.pressure_a_display = tk.StringVar(value="0.00")
        tk.Label(data_frame, textvariable=self.pressure_a_display, font=('Arial', 12, 'bold')).grid(row=0, column=1, sticky='w')
        tk.Label(data_frame, text="PSI").grid(row=0, column=2, sticky='w')
        
        tk.Label(data_frame, text="Alicat B Pressure:").grid(row=1, column=0, sticky='e')
        self.pressure_b_display = tk.StringVar(value="0.00")
        tk.Label(data_frame, textvariable=self.pressure_b_display, font=('Arial', 12, 'bold')).grid(row=1, column=1, sticky='w')
        tk.Label(data_frame, text="PSI").grid(row=1, column=2, sticky='w')
        
        tk.Label(data_frame, text="Alicat A Flow:").grid(row=0, column=3, sticky='e', padx=(20, 0))
        self.flow_a_display = tk.StringVar(value="0.000")
        tk.Label(data_frame, textvariable=self.flow_a_display, font=('Arial', 12, 'bold')).grid(row=0, column=4, sticky='w')
        tk.Label(data_frame, text="SLPM").grid(row=0, column=5, sticky='w')
        
        tk.Label(data_frame, text="Alicat B Flow:").grid(row=1, column=3, sticky='e', padx=(20, 0))
        self.flow_b_display = tk.StringVar(value="0.000")
        tk.Label(data_frame, textvariable=self.flow_b_display, font=('Arial', 12, 'bold')).grid(row=1, column=4, sticky='w')
        tk.Label(data_frame, text="SLPM").grid(row=1, column=5, sticky='w')
        
        # Test Status
        status_frame = tk.LabelFrame(self.root, text="Test Status", padx=10, pady=10)
        status_frame.grid(row=3, column=0, columnspan=2, padx=10, pady=10, sticky='ew')
        
        tk.Label(status_frame, text="Phase:").grid(row=0, column=0, sticky='e')
        self.test_phase = tk.StringVar(value="Idle")
        tk.Label(status_frame, textvariable=self.test_phase, font=('Arial', 12, 'bold'), fg='blue').grid(row=0, column=1, sticky='w')
        
        tk.Label(status_frame, text="Time Remaining:").grid(row=0, column=2, sticky='e', padx=(20, 0))
        self.time_remaining = tk.StringVar(value="0")
        self.time_remaining_label = tk.Label(status_frame, textvariable=self.time_remaining, 
                                             font=('Arial', 12, 'bold'), fg='black')
        self.time_remaining_label.grid(row=0, column=3, sticky='w')
        tk.Label(status_frame, text="s").grid(row=0, column=4, sticky='w')
        
        # Control Buttons
        button_frame = tk.Frame(self.root, padx=10, pady=10)
        button_frame.grid(row=4, column=0, columnspan=2)
        
        self.start_button = tk.Button(button_frame, text="Start Test", command=self.start_test, 
                                      bg='green', fg='white', width=15, height=2)
        self.start_button.grid(row=0, column=0, padx=5)
        
        self.stop_button = tk.Button(button_frame, text="Stop Test", command=self.stop_test, 
                                     bg='red', fg='white', width=15, height=2, state='disabled')
        self.stop_button.grid(row=0, column=1, padx=5)
        
        # Plot Area
        plot_frame = tk.LabelFrame(self.root, text="Pressure Decay Plot", padx=10, pady=10)
        plot_frame.grid(row=5, column=0, columnspan=2, padx=10, pady=10, sticky='nsew')
        
        self.fig, self.ax = plt.subplots(figsize=(8, 4))
        self.ax.set_xlabel('Time (s)')
        self.ax.set_ylabel('Pressure (PSI)')
        self.ax.set_title('Alicat A Pressure Decay')
        self.ax.grid(True)
        
        self.canvas = FigureCanvasTkAgg(self.fig, master=plot_frame)
        self.canvas.draw()
        self.canvas.get_tk_widget().pack()
        
        # Configure grid weights for resizing
        self.root.grid_rowconfigure(5, weight=1)
        self.root.grid_columnconfigure(0, weight=1)
        
    def send_command(self, command):
        """Send command over serial port"""
        self.ser.write(command.encode())
        time.sleep(0.1)
    
    def read_alicat(self, device):
        """Read data from specified Alicat device (A or B)"""
        command = f"{device}\r"
        self.ser.reset_input_buffer()
        self.send_command(command)
        response = self.ser.read_until(b"\r").decode('utf-8', errors='ignore')
        data = response.split()
        
        if len(data) >= 5:
            # Alicat response format: [ID, Pressure, Temperature, VolumetricFlow, MassFlow, SetPoint, Gas]
            return {
                'pressure': float(data[1]),
                'temperature': float(data[2]),
                'volumetric_flow': float(data[3]),
                'mass_flow': float(data[4])
            }
        return None
    
    def start_test(self):
        """Start the dual Alicat test sequence"""
        if not self.part_number.get():
            messagebox.showerror("Error", "Please enter a part number.")
            return
        
        # Disable start button, enable stop button
        self.start_button.config(state='disabled')
        self.stop_button.config(state='normal')
        
        # Create Excel file
        part_number = self.part_number.get()
        timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
        os.makedirs(path, exist_ok=True)
        self.excel_path = os.path.join(path, f"{part_number}_{timestamp}.xlsx")
        
        # Create workbook with two sheets
        self.workbook = Workbook()
        self.settings_sheet = self.workbook.active
        self.settings_sheet.title = "Settings"
        self.data_sheet = self.workbook.create_sheet(title="Data")
        
        # Write settings
        self.settings_sheet.append(["Setting", "Value"])
        self.settings_sheet.append(["Part Number", part_number])
        self.settings_sheet.append(["Timestamp", timestamp])
        self.settings_sheet.append(["A Flow Test Pressure (PSI)", self.a_flow_test_pressure])
        self.settings_sheet.append(["B Flow Test Pressure (PSI)", self.b_flow_test_pressure])
        self.settings_sheet.append(["B Decay Test Pressure (PSI)", self.b_decay_test_pressure])
        self.settings_sheet.append(["Flow Sample Time (s)", self.flow_sample_time])
        self.settings_sheet.append(["Pressure Sample Time (s)", self.pressure_sample_time])
        self.settings_sheet.append(["Read Rate (s)", self.read_rate])
        self.settings_sheet.append(["Pressurize Time (s)", self.pressurize_time])
        
        # Write data headers
        self.data_sheet.append(["Phase", "Time (s)", "A Pressure (PSI)", "B Pressure (PSI)", 
                               "A Flow (SLPM)", "B Flow (SLPM)"])
        
        self.workbook.save(self.excel_path)
        
        # Reset data arrays
        self.time_data = []
        self.pressure_a_data = []
        self.pressure_b_data = []
        self.flow_a_data = []
        self.flow_b_data = []
        
        # Run test sequence
        try:
            self.run_flow_test()
            self.run_pressure_decay_test()
            messagebox.showinfo("Test Complete", f"Test completed successfully!\nData saved to:\n{self.excel_path}")
        except Exception as e:
            messagebox.showerror("Test Error", f"An error occurred during the test:\n{e}")
        finally:
            self.test_phase.set("Complete")
            self.start_button.config(state='normal')
            self.stop_button.config(state='disabled')
            if self.workbook:
                self.workbook.save(self.excel_path)
    
    def run_flow_test(self):
        """Execute the flow test phase"""
        self.test_phase.set("Flow Test - Setup")
        self.root.update_idletasks()
        
        # Step 1: Release valve on A
        self.send_command("AC\r")
        time.sleep(0.5)
        
        # Step 2: Set A to flow test pressure
        self.send_command(f"AS{self.a_flow_test_pressure}\r")
        
        # Step 3: Set B to flow test pressure
        self.send_command(f"BS{self.b_flow_test_pressure}\r")
        
        # Step 4: Wait for pressure stabilization
        self.test_phase.set("Flow Test - Stabilizing")
        remaining = int(self.pressurize_time)
        while remaining > 0:
            self.time_remaining.set(str(remaining))
            
            # Read and display pressures during stabilization
            data_a = self.read_alicat('A')
            data_b = self.read_alicat('B')
            if data_a:
                self.pressure_a_display.set(f"{data_a['pressure']:.2f}")
                self.flow_a_display.set(f"{data_a['mass_flow']:.3f}")
            if data_b:
                self.pressure_b_display.set(f"{data_b['pressure']:.2f}")
                self.flow_b_display.set(f"{data_b['mass_flow']:.3f}")
            
            self.root.update_idletasks()
            time.sleep(1)
            remaining -= 1
        
        # Step 5: Record mass flow for both devices
        self.test_phase.set("Flow Test - Recording")
        flow_start_time = time.time()
        flow_end_time = flow_start_time + self.flow_sample_time
        
        while time.time() < flow_end_time:
            elapsed = time.time() - flow_start_time
            remaining_time = int(flow_end_time - time.time())
            self.time_remaining.set(str(remaining_time))
            
            # Read data from both Alicats
            data_a = self.read_alicat('A')
            data_b = self.read_alicat('B')
            
            if data_a and data_b:
                # Update displays
                self.pressure_a_display.set(f"{data_a['pressure']:.2f}")
                self.pressure_b_display.set(f"{data_b['pressure']:.2f}")
                self.flow_a_display.set(f"{data_a['mass_flow']:.3f}")
                self.flow_b_display.set(f"{data_b['mass_flow']:.3f}")
                
                # Store data
                self.data_sheet.append(["Flow Test", f"{elapsed:.2f}", 
                                       f"{data_a['pressure']:.2f}", f"{data_b['pressure']:.2f}",
                                       f"{data_a['mass_flow']:.3f}", f"{data_b['mass_flow']:.3f}"])
                self.workbook.save(self.excel_path)
            
            self.root.update_idletasks()
        time.sleep(self.read_rate)
    
    def run_pressure_decay_test(self):
        """Execute the pressure decay test phase"""
        self.test_phase.set("Decay Test - Setup")
        self.root.update_idletasks()
        
        # Step 1: Set B to decay test pressure (close valve)
        self.send_command(f"BS{self.b_decay_test_pressure}\r")
        
        # Step 2: Wait for A pressure to stabilize
        self.send_command(f"AS{self.a_decay_test_pressure}\r")  # Ensure A is at flow test pressure for decay test
        self.test_phase.set("Decay Test - Stabilizing")
        remaining = int(self.pressurize_time)
        while remaining > 0:
            self.time_remaining.set(str(remaining))
            
            # Read and display pressures during stabilization
            data_a = self.read_alicat('A')
            data_b = self.read_alicat('B')
            if data_a:
                self.pressure_a_display.set(f"{data_a['pressure']:.2f}")
                self.flow_a_display.set(f"{data_a['mass_flow']:.3f}")
            if data_b:
                self.pressure_b_display.set(f"{data_b['pressure']:.2f}")
                self.flow_b_display.set(f"{data_b['mass_flow']:.3f}")
            
            self.root.update_idletasks()
            time.sleep(self.read_rate)
            remaining -= 1
        
        # Step 3: Close valve on A
        self.send_command("AHC\r")
        time.sleep(0.25)
        
        # Step 4: Record pressure decay
        self.test_phase.set("Decay Test - Recording")
        self.time_remaining.set("0")
        
        decay_start_time = time.time()
        decay_end_time = decay_start_time + self.pressure_sample_time
        
        # Clear plot data for decay test
        self.time_data = []
        self.pressure_a_data = []
        
        while time.time() < decay_end_time:
            elapsed = time.time() - decay_start_time
            remaining_time = int(decay_end_time - time.time())
            self.time_remaining.set(str(remaining_time))
            
            # Read data from Alicat A
            data_a = self.read_alicat('A')
            data_b = self.read_alicat('B')
            
            if data_a:
                # Update displays
                self.pressure_a_display.set(f"{data_a['pressure']:.2f}")
                if data_b:
                    self.pressure_b_display.set(f"{data_b['pressure']:.2f}")
                
                # Store data for plotting
                self.time_data.append(elapsed)
                self.pressure_a_data.append(data_a['pressure'])
                
                # Store data in Excel
                self.data_sheet.append(["Pressure Decay", f"{elapsed:.2f}", 
                                       f"{data_a['pressure']:.2f}", 
                                       f"{data_b['pressure']:.2f}" if data_b else "N/A",
                                       f"{data_a['mass_flow']:.3f}",
                                       f"{data_b['mass_flow']:.3f}" if data_b else "N/A"])
                self.workbook.save(self.excel_path)
                
                # Update plot
                self.update_plot()
            
            self.root.update_idletasks()
            time.sleep(self.pressure_read_rate)
        
        self.time_remaining.set("0")
    
    def update_plot(self):
        """Update the pressure decay plot"""
        self.ax.clear()
        if self.time_data and self.pressure_a_data:
            self.ax.plot(self.time_data, self.pressure_a_data, 'b-', linewidth=2, label='Alicat A Pressure')
            self.ax.set_xlabel('Time (s)')
            self.ax.set_ylabel('Pressure (PSI)')
            self.ax.set_title('Alicat A Pressure Decay')
            self.ax.grid(True, alpha=0.3)
            self.ax.legend()
            
            # Set reasonable y-axis limits
            if len(self.pressure_a_data) > 0:
                min_p = min(self.pressure_a_data)
                max_p = max(self.pressure_a_data)
                padding = (max_p - min_p) * 0.1 or 1
                self.ax.set_ylim(min_p - padding, max_p + padding)
        
        self.canvas.draw()
    
    def stop_test(self):
        """Stop the test and close valves"""
        self.test_phase.set("Stopping")
        self.send_command("AHC\r")  # Close valve A
        self.send_command(f"BS{self.b_decay_test_pressure}\r")  # Set B to decay pressure
        self.test_phase.set("Stopped")
        self.start_button.config(state='normal')
        self.stop_button.config(state='disabled')
        if self.workbook:
            self.workbook.save(self.excel_path)
    
    def __del__(self):
        """Cleanup on exit"""
        if hasattr(self, 'ser') and self.ser.is_open:
            self.ser.close()

# Create and run the application
if __name__ == "__main__":
    root = tk.Tk()
    app = DualAlicatTestApp(root)
    root.mainloop()
