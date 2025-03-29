# FilterPointsInsidePologon
Filter Points Inside Pologon Used to get Points Name inside Pologon or LineString

A Python application that processes KML files to identify points contained within polygons or converted LineStrings, with results exported to Excel.

![image](https://github.com/user-attachments/assets/4a0bc894-b7f3-41c9-ad71-ba67599c3980)

## 🚀 Features

- **Drag & Drop Interface**: Simply drop your KML file to start processing
- **Smart Conversion**: Automatically converts LineStrings to Polygons for analysis
- **Multi-Feature Support**: Process multiple polygons/linestrings simultaneously
- **Interactive Selection**: Visually select which features to analyze
- **Excel Export**: Clean, organized output in Excel format with multiple sheets
- **Progress Tracking**: Real-time progress updates during processing

## 📦 Installation

1. Ensure you have Python 3.8+ installed
2. Clone this repository:
   ```bash
   git clone https://github.com/shadywardy/FilterPointsInsidePologon.git
   cd FilterPointsInsidePologon

## 📦Install dependencies:
pip install -r requirements.txt

🛠️ Usage

## 📦Run the application:
python FilterPointsInsidePologon.py
or use EXE File inside Folder dis

Either:
*Drag & drop a KML file onto the window, or
*Click "Select File" to browse for your KML
*Choose whether to convert LineStrings to Polygons
*Select which features to analyze in the interactive dialog
*Specify an output Excel file location
*Get your results!

## 📦Technical Details
Backend: Shapely for geometric operations, lxml for KML parsing
Frontend: Tkinter with modern UI elements and drag-and-drop support
Performance: Multi-threaded processing for large files
Output: Excel files with three sheets:
Polygon: Points inside original polygons
LineString: Points inside converted linestrings
Unassigned: Points not contained in any selected feature

## 🌟 Why This Tool?
Precision: Accurate point-in-polygon calculations using robust geometric libraries

User-Friendly: No GIS expertise required - perfect for field researchers

Flexible: Works with both KML and (through conversion) KMZ files

Time-Saving: Processes complex spatial relationships in seconds

## 📜 License
MIT License - Feel free to use and modify for your projects!

## 🤝 Contributing
Pull requests welcome! Please open an issue first to discuss major changes.

