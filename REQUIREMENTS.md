# Pressure Flow Test Application - Requirements & Features

## Project Overview
A Tkinter-based GUI application for testing catheter pressure decay over time using Alicat mass flow controllers connected via serial ports. The application automates pressure testing with configurable parameters and logs data to Excel workbooks.

---

## System Architecture

### Hardware Components
- **Alicat A**: Sets input pressure via serial port (38400 baud)
- **Alicat B**: Measures flow rate via serial port (38400 baud)
- **BB9 Connector**: Connects both Alicats
- **Catheter**: Connected to Alicat B for testing

### Software Components
- **GUI Framework**: Tkinter
- **Data Storage**: OpenPyXL (Excel workbooks)
- **Serial Communication**: PySerial
- **Data Visualization**: Matplotlib
- **Configuration**: Test.ini file

---

## Current Features

### 1. Configuration Management
- **File**: Test.ini
- **Parameters**:
  - `SetPressure`: Initial pressure setting (PSI)
  - `ThresholdPressure`: Minimum pressure threshold
  - `SampleRate`: Data sampling rate
  - `TestDuration`: Total test duration (seconds)
  - `PressureWaitTime`: Time to wait for pressure stabilization
  - `TestRepeats`: Number of test repeats

### 2. Serial Communication
- Send commands to Alicat devices
- Read pressure and flow rate data from Alicat A
- Format: Command string ending with `\r`
- Baud rate: 38400

### 3. Valve Control
- **Release Valve** (`AC`): Opens valve to start test
- **Hold Valve Closed** (`HC`): Closes valve to hold pressure

### 4. Data Collection
- **Time**: Elapsed time from test start (seconds)
- **Pressure**: Real-time pressure readings (PSI)
- **Gas Flow**: Flow rate measurements

### 5. Excel Data Logging
- **Sheets**:
  - **Settings**: Test configuration, part number, timestamp, parameter changes
  - **Data**: Time-series pressure and flow data
- **Auto-save**: Data saved after each reading
- **Repeat Tracking**: Separate data columns for repeated tests
- **File naming**: `{PartNumber}_{Timestamp}.xlsx`

### 6. Timer Display
- **Countdown**: Shows remaining test time (mm:ss or seconds)
- **Color Coding**:
  - Green: > 60 seconds remaining
  - Orange: 30-60 seconds remaining
  - Red: < 30 seconds remaining

### 7. GUI Elements
- Part number entry field
- Starting pressure display
- Pressure set/change controls
- Current pressure display
- Gas flow display
- Repeats remaining counter
- Pressurizing countdown
- Real-time pressure vs. time plot
- Control buttons:
  - Start Test
  - Release Valve
  - Repeat Test
  - End Test

### 8. Test Workflow
1. Enter part number
2. Click "Start Test"
3. Release valve opens to test
4. Pressure stabilizes (pressurize_time countdown)
5. Valve closes, test begins
6. Data collected for test_duration
7. Option to repeat test
8. Save complete results to Excel

### 9. Data Visualization
- Real-time matplotlib plot
- X-axis: Time (seconds)
- Y-axis: Pressure (PSI)
- Dynamic Y-axis limits based on data range

---

## Data Storage Locations

- **Test Data Path**: `C:\Users\patri\RnD\SW Test Data\`
- **Excel Files**: Stored by part number and timestamp
- **Format**: `.xlsx` with multiple sheets

---

## Known Limitations / TODO

- [ ] Error handling for serial port disconnection
- [ ] Validation of Alicat responses
- [ ] Recovery from malformed serial data
- [ ] GUI responsiveness during data collection
- [ ] Optional data export to CSV format
- [ ] Real-time data validation and alerts
- [ ] Pressure threshold monitoring during test
- [ ] Multiple simultaneous tests
- [ ] Test results summary/statistics sheet
- [ ] Unit configuration (PSI vs. other units)

---

## Code Structure

### Key Methods

| Method | Purpose |
|--------|---------|
| `read_ini()` | Load configuration from Test.ini |
| `start_test()` | Initialize Excel file and begin test sequence |
| `collect_data()` | Main data collection loop |
| `update_timer()` | Update countdown display with color coding |
| `send_command(command)` | Send command over serial port |
| `send_hold_valve_closed()` | Send HC command to close valve |
| `send_release_valve()` | Send AC command to open valve |
| `repeat_test()` | Run test again without resetting data file |
| `end_test()` | Close valve, save data, and exit |
| `set_pressure_command()` | Change pressure mid-test |

---

## Version History

- **v1.0**: Initial release with basic pressure testing functionality
- **v2.0**: Refactored for dual Alicat operation with separate flow and decay test phases

---

## Requirements for New Features / Changes

When adding new features or making changes, update this file with:
1. **Feature Description**: What does it do?
2. **Affected Components**: Which methods/files are modified?
3. **Data Impact**: Does it affect Excel structure or data collection?
4. **GUI Impact**: Does it add/modify UI elements?
5. **Configuration Impact**: Does it require new Test.ini parameters?
6. **Testing Notes**: How to verify the feature works correctly?

---

## Development Notes

- Serial communication uses non-blocking reads
- Excel workbook is saved after each data point for data safety
- Timer is updated via GUI event loop (not threads)
- All time measurements use `time.time()` for consistency
- Part number is required before starting tests

---

## Version 2.0 Changes (NEW)

### Architecture Changes
- **Dual Alicat Communication**: Commands prefixed with device ID (A or B)
  - Alicat A: Pressure control and decay measurement
  - Alicat B: Flow measurement and valve control
- **Two-Phase Test Sequence**: Separated flow test and pressure decay test

### New Test Workflow

#### Phase 1: Flow Test
1. Release valve on Alicat A (`AC\r`)
2. Set Alicat A to `A_FLOW_TEST_PRESSURE` (`AS#.#\r`)
3. Set Alicat B to `B_FLOW_TEST_PRESSURE` (`BS#.#\r`)
4. Wait for stabilization (`PRESSURIZE_TIME`)
5. Record mass flow from both devices for `FLOW_SAMPLE_TIME`

#### Phase 2: Pressure Decay Test
1. Set Alicat B to `B_DECAY_TEST_PRESSURE` (`BS#.#\r`) to close valve
2. Wait for Alicat A pressure to stabilize
3. Close valve on Alicat A (`AHC\r`)
4. Record pressure decay for `PRESSURE_SAMPLE_TIME` at `READ_RATE` intervals

### New Configuration Parameters (Test.ini)
- `A_FLOW_TEST_PRESSURE`: Pressure setting for Alicat A during flow test (PSI)
- `B_FLOW_TEST_PRESSURE`: Pressure setting for Alicat B during flow test (PSI)
- `B_DECAY_TEST_PRESSURE`: Pressure for Alicat B during decay test (default 0 PSI)
- `FLOW_SAMPLE_TIME`: Duration to record flow data (default 5s)
- `PRESSURE_SAMPLE_TIME`: Duration to record pressure decay (default 20s)
- `READ_RATE`: Data sampling interval (default 1s)
- `PRESSURIZE_TIME`: Stabilization wait time (default 10s)

### Data Collection Changes
- **Excel Structure**:
  - Sheet 1 (Settings): All test parameters and timestamps
  - Sheet 2 (Data): Phase-tagged measurements
    - Columns: Phase, Time, A Pressure, B Pressure, A Flow, B Flow
    - Phases: "Flow Test" and "Pressure Decay"

### GUI Enhancements
- Real-time display for both Alicat A and B (pressure and flow)
- Test phase indicator (Idle, Flow Test, Decay Test, etc.)
- Simplified control: Start Test and Stop Test buttons
- Live pressure decay plot during decay phase

### Command Format Changes
| Action | V1.0 Command | V2.0 Command A | V2.0 Command B |
|--------|-------------|----------------|----------------|
| Set Pressure | `AS10.0\r` | `AS10.0\r` | `BS10.0\r` |
| Release Valve | `AC\r` | `AC\r` | `BC\r` |
| Hold/Close Valve | `AHC\r` | `AHC\r` | `BHC\r` |
| Poll Data | `A\r` | `A\r` | `B\r` |

### Files
- **Main Program**: `Pressure_Flow_v2.py`
- **Configuration**: `Test.ini` (new format)
- **Data Output**: Excel files with dual-device measurements
