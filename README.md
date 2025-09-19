# PressTech - Siphony GUI

An integrated graphical user interface for foam processing analysis tools used in the PressTech research project.

## Overview

This GUI integrates six different analysis modules that were previously independent scripts with hardcoded folder paths. Now users can select input files through a user-friendly interface for each analysis type.

## Modules

### 1. Analysis Results
- **Purpose**: Analyze combined results from All_Results.xlsx files
- **Input**: Excel file with combined experimental results
- **Output**: Statistical analysis, correlation matrices, and visualizations
- **Features**: Summary statistics, distribution plots, scatter plots by polymer type

### 2. Combine Results
- **Purpose**: Combine multiple data sources into a single results file
- **Input**: DoE file, Density file, PDR file, Histogram file
- **Output**: Combined All_Results.xlsx file
- **Features**: Automatic data merging, column mapping, polymer filtering

### 3. DSC Analysis
- **Purpose**: Process DSC (Differential Scanning Calorimetry) data files
- **Input**: Text files containing DSC measurement data
- **Output**: Excel files with extracted thermal properties (Tm, Xc, Tg)
- **Features**: Support for both semicrystalline and amorphous polymers

### 4. SEM Image Editor
- **Purpose**: Edit and process SEM (Scanning Electron Microscopy) images
- **Input**: Image files (PNG, JPG, TIFF, BMP)
- **Output**: Processed images with scale bars and borders
- **Features**: Scale calibration, image cropping, scale bar addition, colored borders

### 5. Open-Cell Content
- **Purpose**: Calculate open-cell content from picnometry data
- **Input**: Density Excel file, Picnometry Excel files by polymer
- **Output**: OC calculation results by polymer
- **Features**: Automatic sample matching, statistical summaries

### 6. Pressure Drop Rate
- **Purpose**: Calculate pressure drop rate from CSV data files
- **Input**: CSV files with time-series pressure data
- **Output**: Excel files with PDR calculations and graphs
- **Features**: Automatic pressure curve analysis, graph generation

## Installation

1. **Clone or download** this repository to your local machine

2. **Install Python dependencies**:
   ```bash
   pip install -r requirements.txt
   ```

3. **Required Python packages**:
   - pandas (≥1.5.0)
   - matplotlib (≥3.5.0)
   - seaborn (≥0.11.0)
   - openpyxl (≥3.0.9)
   - Pillow (≥9.0.0)
   - numpy (≥1.21.0)

## Usage

### Starting the Application

Run the main GUI application:
```bash
python main_gui.py
```

### Using Individual Modules

Each module can be accessed through the main GUI by clicking the corresponding button:

- 📊 **ANALYSIS RESULTS** - Statistical analysis and visualization
- 🔄 **COMBINE RESULTS** - Data combination and merging
- 🌡️ **DSC ANALYSIS** - Thermal analysis data processing
- 🔬 **EDITING SEM IMAGES** - Image editing and processing
- 🔓 **OPEN-CELL CONTENT** - Picnometry analysis
- 📉 **PRESSURE DROP RATE** - Pressure analysis

### File Format Requirements

#### DoE File (Excel)
- Columns: Label, m(g), Water (g), T(ºC), P CO2(bar), t(min)
- Used for experimental design parameters

#### Density File (Excel)
- Columns: Label, density measurements, porosity calculations
- Used for foam density analysis

#### PDR Files (CSV)
- Columns: Time, T1(ºC), T2(ºC), P(bar)
- Time-series pressure data for PDR calculation

#### DSC Files (Text)
- Plain text files with DSC measurement results
- Must contain 'Sample:' and 'Results:' sections

#### Histogram Files (Excel)
- Cell size distribution data
- Used for morphological analysis

#### Picnometry Files (Excel)
- Volume measurements for open-cell content calculation
- One file per polymer type

## Features

### User-Friendly Interface
- Intuitive button-based navigation
- File selection dialogs instead of hardcoded paths
- Progress indicators and status updates
- Tooltips and help text

### Data Validation
- Input file format checking
- Missing data detection
- Error handling and user feedback

### Flexible Output
- User-selectable output folders
- Customizable output filenames
- Multiple output formats (Excel, images, reports)

### Visualization
- Real-time data preview
- Interactive plots and charts
- Export-ready graphs and figures

## Project Structure

```
Siphony GIU/
├── main_gui.py              # Main application entry point
├── requirements.txt         # Python dependencies
├── README.md               # This file
├── modules/                # Analysis modules
│   ├── __init__.py
│   ├── analysis_module.py  # Statistical analysis
│   ├── combine_module.py   # Data combination
│   ├── dsc_module.py       # DSC processing
│   ├── sem_module.py       # SEM image editing
│   ├── oc_module.py        # Open-cell content
│   └── pdr_module.py       # Pressure drop rate
└── [original folders]      # Original independent scripts
```

## Migration from Original Scripts

This GUI replaces the original independent scripts while maintaining the same core functionality:

- **Analysis All_Results/Analysis.py** → Analysis Results module
- **Combine/Combine.py** → Combine Results module  
- **DSC/DSCfoams_*.py** → DSC Analysis module
- **Editing SEM images/SEMimages.py** → SEM Image Editor module
- **OC (Open-cell content)/OC.py** → Open-Cell Content module
- **PDR (Pressure Drop Rate)/PDR.py** → Pressure Drop Rate module

The main advantages of the GUI version:
- No hardcoded file paths
- Interactive file selection
- Integrated workflow
- Better error handling
- Consistent user interface

## Troubleshooting

### Common Issues

1. **Module import errors**: Ensure all dependencies are installed
2. **File not found errors**: Check file paths and permissions
3. **Excel formatting issues**: Verify input file column names and structure
4. **Memory issues with large files**: Process files in smaller batches

### Getting Help

- Check the status bar for current operation status
- Use the tooltips on buttons for quick help
- Validate input files before processing
- Check the console output for detailed error messages

## Contributing

When adding new analysis modules:
1. Create a new module file in the `modules/` directory
2. Follow the existing module structure pattern
3. Add the import and button in `main_gui.py`
4. Update this README with the new module documentation

## Version History

- **v1.0** - Initial release with all six integrated modules
  - Migrated from independent scripts
  - Added file selection dialogs
  - Integrated user interface
  - Added progress indicators and status updates
