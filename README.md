# FilterPointsInsidePologon
Filter Points Inside Pologon Used to get Points Name inside Pologon or LineString

A Python application that processes KML files to identify points contained within polygons or converted LineStrings, with results exported to Excel.

![Application Screenshot](screenshot.png)

## Features

- Processes KML files containing Points, Polygons, and LineStrings
- Optionally converts LineStrings to Polygons (closed shapes)
- Interactive feature selection interface
- Identifies which points are contained within which polygons
- Generates Excel reports with:
  - Points contained in each Polygon
  - Points contained in converted LineStrings (as Polygons)
  - Unassigned points not in any selected feature
- Drag-and-drop file support
- Progress tracking and status updates

## Requirements

- Python 3.7+
- Required packages (install via `pip install -r requirements.txt`):
-pandas
-shapely
-pillow
-lxml
-pykml
-openpyxl
-tkinterdnd2


