# Version 1.0
# Test catheter for pressure decay over time using Alicat and measure flow rate
# Setup
# 2 alicats operating over serial port at 38400 baud connect via BB9
# Alicat A - Sets input Pressure
# Alicat B - Measures Flow rate
# Catheter is connected to Alicat B
# Process
# 1. Send command to Alicat A to set pressure to desired level
# 2. Set Alicat B to 0 pressure to measure flow through catheter once stablized
# 3. Set Alicate B to 2x of Alicat A pressure to close valve Wait to stabilize
# 3. Close Alicat Value to hold pressure and measure pressure decay over time
# 
#Change directory for Box and remove Storing gas Flow data
import tkinter as tk
from tkinter import messagebox
import serial
import csv
import time
import datetime
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
import os
import openpyxl
from openpyxl import Workbook

# change path name for your box folder
path = r"C:\Users\patri\RnD\SW Test Data"  # Define the path as a string variable


    def read_ini(self):
        # Read settings from Test.ini file
        try:
            with open("Test.ini", "r") as file:
                for line in file:
                    if "SetPressure" in line:
                        self.set_pressure = float(line.split('=')[1].strip())
                        self.starting_pressure.set(f"{self.set_pressure:.2f}")  # Update starting pressure display
                    elif "ThresholdPressure" in line:
                        self.threshold_pressure = float(line.split('=')[1].strip())
                    elif "SampleRate" in line:
                        self.sample_rate = float(line.split('=')[1].strip())
                    elif "TestDuration" in line:
                        self.test_duration = float(line.split('=')[1].strip()) 
                    elif "PressureWaitTime" in line:
                        self.pressurize_time = float(line.split('=')[1].strip())
                    elif "TestRepeats" in line:
                        self.test_repeats = int(line.split('=')[1].strip())   
        except FileNotFoundError:
            messagebox.showerror("Error", "Test.ini file not found!")

    def update_timer(self):
        """Update the countdown timer display"""
        if hasattr(self, 'end_test_time') and hasattr(self, 'start_time'):
            time_left = self.end_test_time - time.time()
            if time_left > 0:
                minutes = int(time_left // 60)
                seconds = int(time_left % 60)
                if minutes > 0:
                    self.time_remaining.set(f"{minutes}:{seconds:02d}")
                else:
                    self.time_remaining.set(f"{seconds}")
                
                # Change color based on time remaining
                if time_left <= 30:  # Last 30 seconds
                    self.time_remaining_label.config(fg="red")
                elif time_left <= 60:  # Last minute
                    self.time_remaining_label.config(fg="orange")
                else:
                    self.time_remaining_label.config(fg="green")
            else:
                self.time_remaining.set("0")
                self.time_remaining_label.config(fg="red")
        else:
            self.time_remaining.set("0")
            self.time_remaining_label.config(fg="black")
            
    def send_command(self, command):
        """ Send command over serial port """
        self.ser.write(command.encode())
        time.sleep(0.1)  # Wait for command to be processed

    def start_test(self):

        """ Start the test by sending the initial command to set pressure """
        if not self.part_number.get():
            messagebox.showerror("Error", "Please enter a part number.")
            return

        # Disable the "Start Test" button
        self.start_button.config(state='disabled')
        # Enable the other buttons
        self.release_valve_button.config(state='normal')
        self.repeat_test_button.config(state='normal')
        self.end_test_button.config(state='normal')

        part_number = self.part_number.get()
        timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
        os.makedirs(path, exist_ok=True)
        # Create a new Excel workbook and add two sheets: Settings and Data
        self.excel_path = os.path.join(path, f"{part_number}_{timestamp}.xlsx")
        self.workbook = Workbook()
        self.settings_sheet = self.workbook.active
        self.settings_sheet.title = "Settings"
        self.data_sheet = self.workbook.create_sheet(title="Data")
        self.workbook.save(self.excel_path)

        # Write settings to the Settings sheet
        self.settings_sheet.append(["Setting", "Value"])
        self.settings_sheet.append(["Set Pressure", self.set_pressure])
        self.settings_sheet.append(["Threshold Pressure", self.threshold_pressure])
        self.settings_sheet.append(["Sample Rate", self.sample_rate])
        self.settings_sheet.append(["Test Duration", self.test_duration])
        self.settings_sheet.append(["Part Number", part_number])
        self.settings_sheet.append(["Timestamp", timestamp])
        self.workbook.save(self.excel_path)

        # Write header to the Data sheet
        self.data_sheet.append(["Time (s)", "Pressure", "Gas Flow"])

        # Save the workbook path for later use
        self.excel_path = os.path.join(path, f"{part_number}_{timestamp}.xlsx")
        self.collect_data()
        if self.number_of_repeats > 0:
            self.number_of_repeats -= 1
            self.repeats_remaining.set(self.number_of_repeats)
            self.repeat_test()

    def collect_data(self):
        """ Collect data by sending the command to start sampling """
        
        self.send_release_valve()  # Release the valve to start the test
        time.sleep(.5)  # for valve to open
        # Send initial set pressure command
        # Show a "pressurizing" icon before sleep


        self.root.update_idletasks()

        self.send_command(f"AS{self.set_pressure}\r")
        # Wait for the pressure to stabilize and update pressurizing_remaining every second
        remaining = int(self.pressurize_time)
        while remaining > 0:
            self.pressurizing_remaining.set(str(remaining))
            self.root.update_idletasks()
            time.sleep(1)
            remaining -= 1
        self.pressurizing_remaining.set("0")

        self.root.update_idletasks()
        self.send_hold_valve_closed()  # hold pressure
        
        # Start sampling loop
        self.start_time = time.time()
        self.end_test_time = self.start_time + self.test_duration
        #clear the serial buffer-
        self.ser.reset_input_buffer()  # Clear the serial buffer
        #self.send_command("A@ @\r")  # start streaming data
        # self.send_command("A\r")
        # response_buffer = self.ser.read_until("\r").decode('utf-8') # waste first response to sync up with the data stream
        # ali_string = response_buffer.split()
        # new_pressure = round(float(ali_string[1]),2)
        # self.current_pressure.set(f"{new_pressure:.2f}")
        
        # Wait until pressure from Alicat is equal to or less than set_pressure
        while True:
            self.send_command("A\r")
            response_buffer = self.ser.read_until("\r").decode('utf-8')
            ali_string = response_buffer.split()
            if len(ali_string) > 1:
                new_pressure = round(float(ali_string[1]), 2)
                self.current_pressure.set(f"{new_pressure:.2f}")
                self.root.update_idletasks()  # Update the GUI to show current pressure
            if new_pressure <= self.set_pressure:
                    break
            time.sleep(0.1)  # Small delay to avoid busy waiting

        # Reset timer for correct test duration
        self.start_time = time.time()
        self.end_test_time = self.start_time + self.test_duration

        # Initialize timer display
        self.update_timer()
        elapsed_time = 0.0
        #self.current_flow.set(f"{new_flow:.3f}")        
        #time.sleep(.1)  # Wait for the command to be processed 
        self.pressure_data = []
        #save first pressure reading =< setpressure
        self.pressure_data.append(new_pressure)
        self.time_data.append(elapsed_time)
        gas_flow = round(float(ali_string[3]), 3)

        ali_string = ""
        counter = 0
        # Save the first pressure reading (already read above) at time 0

        self.time_data = [elapsed_time]
        self.pressure_data = [new_pressure]
        self.gas_flow.set(f"{gas_flow:.3f}")

        # Store the first datapoint in the Excel workbook's Data sheet
        self.data_sheet.append([elapsed_time, new_pressure, gas_flow])
        self.workbook.save(self.excel_path)

        # Update GUI to show initial values
        self.root.update_idletasks()

        # Continue collecting data for the rest of the test duration
        while time.time() < self.end_test_time:
            elapsed_time = time.time() - self.start_time
            self.send_command("A\r")
            response_buffer = self.ser.read_until("\r").decode('utf-8')
            ali_string = response_buffer.split()
            if len(ali_string) > 3:
                pressure = round(float(ali_string[1]), 2)
                gas_flow = round(float(ali_string[3]), 3)
            else:
                pressure = 0.0
                gas_flow = 0.0
                elapsed_time = round(elapsed_time, 3)

            # Update timer display
            self.update_timer()

            # Log data for plotting and Excel
            self.time_data.append(elapsed_time)
            self.pressure_data.append(pressure)
            self.gas_flow.set(f"{gas_flow:.3f}")

            # Store data in the Excel workbook's Data sheet
            self.data_sheet.append([elapsed_time, pressure, gas_flow])
            self.workbook.save(self.excel_path)

            # Update GUI to show real-time changes
            self.root.update_idletasks()

        self.send_command("@@ A\r")  # stop streaming data

        # Reset timer display
        self.time_remaining.set("0")
        self.time_remaining_label.config(fg="red")

        # Update plot
        self.ax.clear()
        self.ax.plot(self.time_data, self.pressure_data, label="Pressure")
        self.ax.set_title('Pressure vs Time')
        self.ax.set_xlabel('Time (s)')
        self.ax.set_ylabel('Pressure')
        self.ax.legend()  # Add legend to the plot

        # Ensure proper y-axis labeling by setting limits dynamically
        if self.pressure_data:
            self.ax.set_ylim(min(self.pressure_data) - 5, max(self.pressure_data) + 5)

        self.canvas.draw()
        self.root.update_idletasks()  # Ensure the plot updates in real-time


        # if self.pressure < self.threshold_pressure:
        #     self.final_decay_time = time.time() - self.start_time
        #     self.csv_writer.writerow([self.final_decay_time, pressure]) 
        #     messagebox.showinfo("Threshold Reached", f"Pressure dropped below threshold in {self.final_decay_time:.2f} seconds.")
            

        # Schedule the next sampling
        # time.sleep(self.sample_rate)
            


    def finish_test(self):
        """ Finish the test and close the valve """
        self.send_hold_valve_closed()   # Close the valve


    def send_hold_valve_closed(self):
        """ Send 'HC' command to close valve """
        self.send_command("AHC\r")


    def send_release_valve(self):
        """ Send 'AC' command to release valve """
        self.send_command("AC\r")
        self.root.update_idletasks()  # Ensure GUI updates


    def repeat_test(self):
        """ Repeat the test by opening the valve and starting the test again """
        # Reset time and pressure data
        self.time_data = []
        self.pressure_data = []

        # Add a header row to the Excel Data sheet for the repeat test
        if hasattr(self, 'data_sheet'):
            self.data_sheet.append(["Time (s) - Repeat", "Pressure - Repeat", "Gas Flow - Repeat"])
            self.workbook.save(self.excel_path)

        # Log repeat event with timestamp and current settings in the Settings sheet
        if hasattr(self, 'settings_sheet'):
            # Add a blank row for separation
            self.settings_sheet.append([])
            now = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            self.settings_sheet.append(["Repeat Test", "Started", "Time", now])
            # Log current settings
            self.settings_sheet.append(["Set Pressure", self.set_pressure])
            self.settings_sheet.append(["Threshold Pressure", self.threshold_pressure])
            self.settings_sheet.append(["Sample Rate", self.sample_rate])
            self.settings_sheet.append(["Test Duration", self.test_duration])
            self.workbook.save(self.excel_path)

        self.send_release_valve()
        self.collect_data()
        self.root.update_idletasks()  # Ensure GUI updates

        if self.number_of_repeats > 0:
            self.number_of_repeats -= 1
            self.repeats_remaining.set(self.number_of_repeats)
            self.repeat_test()

    def end_test(self):
        """ End the test and close the valve """
        self.send_hold_valve_closed()  # Close the valve
        if self.csv_file:
            self.csv_file.close()
        if hasattr(self, 'workbook'):
            self.workbook.save(self.excel_path)
        messagebox.showinfo("Test Ended", "Test has been ended and data saved.")
        self.root.update_idletasks()  # Ensure GUI updates
        self.root.quit()
        self.root.destroy()
        
    def set_pressure_command(self):
        """ Send 'AS new pressure' command over serial port and log to Settings sheet """
        new_pressure = self.new_pressure_entry.get()
        try:
            new_pressure = float(new_pressure)
            self.starting_pressure.set(f"{new_pressure:.2f}")  # Update starting pressure display
            self.send_command(f"AS{new_pressure}\r")
            messagebox.showinfo("Pressure Set", f"Pressure set to {new_pressure} successfully.")

            # Store the time and new pressure in the Settings sheet if workbook exists
            if hasattr(self, 'workbook') and hasattr(self, 'settings_sheet'):
                now = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                self.settings_sheet.append(["Pressure Change", f"{new_pressure}", "Time", now])
                self.workbook.save(self.excel_path)
        except ValueError:
            messagebox.showerror("Invalid Input", "Please enter a valid number for pressure.")

# Create Tkinter root window
root = tk.Tk()
app = PressureTestApp(root)
root.mainloop()
