# Pressure Flow Test Data Plotter

## Overview
The **Data Plotting Utility** (`plot_test_data.py`) is a companion application to the main Pressure Flow Test application. It allows you to load and visualize test data from Excel files, compare multiple test runs, and analyze pressure decay and flow rate patterns.

## Features

### 📊 Data Visualization
- **Pressure Decay Plots**: Visualizes Alicat A pressure decay over time from the Pressure Decay test phase
- **Multiple File Comparison**: Load multiple Excel test files and plot them together on the same graph with different colors for easy comparison
- **Flow Rate Analysis**: Calculates and displays average flow rates from both Alicat A and Alicat B

### 🎛️ Controls
- **Load Excel File**: Browse and select test data Excel files (default directory: `C:\Users\patri\RnD\SW Test Data`)
- **Clear Plot**: Remove all loaded files and reset the plot to start fresh comparisons
- **Remove Last File**: Remove only the most recently loaded file from the plot while keeping others

### 📈 Display Information
- **Loaded Files Table**: Shows all currently loaded files with their calculated average flow rates
- **Color-Coded Lines**: Each file is displayed with a unique color for easy identification
- **Legend**: Plot legend shows filenames for quick reference

## Usage

### Starting the Application
```bash
python plot_test_data.py
```

### Loading Test Data
1. Click the **"Load Excel File"** button
2. Navigate to your test data directory (defaults to `C:\Users\patri\RnD\SW Test Data`)
3. Select an Excel file from a previous test run
4. The file will be loaded and plotted automatically

### Comparing Multiple Tests
1. Load the first test file as described above
2. Click **"Load Excel File"** again and select another test
3. Repeat for additional files - all will be plotted together with different colors
4. Use the legend to identify each file

### Clearing and Starting Over
- Click **"Clear Plot"** to remove all loaded files and reset
- Click **"Remove Last File"** to undo the most recent file load

## Data Processing

### Extracted Information
The utility reads data from the **"Data"** sheet in the Excel file and extracts:
- **Pressure Decay Phase**: Records time and pressure from Alicat A during the decay test
- **Flow Test Phase**: Records flow rates from both Alicats during the flow test phase

### Calculations
- **Average Flow A (SLPM)**: Mean mass flow rate from Alicat A across both phases
- **Average Flow B (SLPM)**: Mean mass flow rate from Alicat B across both phases

## File Structure

### Excel File Requirements
Your test Excel files must contain:
- **Data Sheet**: With columns:
  1. Phase (Flow Test or Pressure Decay)
  2. Time (s)
  3. A Pressure (PSI)
  4. B Pressure (PSI)
  5. A Flow (SLPM)
  6. B Flow (SLPM)

## Default Directory
The plotter looks for Excel files in:
```
C:\Users\patri\RnD\SW Test Data
```

If this directory doesn't exist, the file browser will default to your user home directory.

## Example Workflow

1. **Run Pressure Flow Test**: Execute `Pressure_Flow_v2.py` to generate test data
2. **Data is Saved**: Excel file is created in the default test data directory
3. **Open Plotter**: Run `plot_test_data.py`
4. **Load Test Files**: Click "Load Excel File" and select your test Excel files
5. **Compare Results**: Multiple files are displayed for comparison
6. **Analyze**: Check flow rates in the info panel and pressure decay curves in the plot

## Tips

- **Color Coding**: The plot uses distinct colors to differentiate between test runs
- **Multiple Comparisons**: You can load up to 8 files before colors repeat (easily extensible)
- **Non-Destructive Removal**: Use "Remove Last File" to undo a load without clearing everything
- **Fresh Start**: Use "Clear Plot" before starting a new comparison session

## Troubleshooting

### "Excel file must contain a 'Data' sheet"
- Ensure your Excel file has a sheet named "Data" (case-sensitive)
- Only Excel files created by `Pressure_Flow_v2.py` are compatible

### File Already Loaded Warning
- Each file can only be loaded once at a time
- Use "Remove Last File" or "Clear Plot" if you want to reload the same file

### No Data Appears in Plot
- Verify the Excel file contains "Pressure Decay" phase data
- Check that the Pressure Decay phase has valid numeric values in the data columns

## Integration with Main Application

This utility works alongside the main `Pressure_Flow_v2.py` application:
- Run `Pressure_Flow_v2.py` to collect test data
- Run `plot_test_data.py` to analyze and compare results
- Both applications are independent and can be run separately

---

**Version**: 1.0  
**Created**: February 2026  
**Compatible With**: Pressure_Flow_v2.py test data format
