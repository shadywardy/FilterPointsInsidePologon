import xml.etree.ElementTree as ET
import pandas as pd
from shapely.geometry import Point, Polygon
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import threading
import os
from PIL import Image, ImageTk
import tempfile
from lxml import etree
from pykml import parser
import time
from collections import defaultdict

try:
    from tkinterdnd2 import DND_FILES, TkinterDnD
except ImportError:
    messagebox.showerror("Error", "tkinterdnd2 is not installed or has issues!")
    exit()


class KMLToExcelConverter:
    def __init__(self, root):
        self.root = root
        self.root.protocol("WM_DELETE_WINDOW", self.on_close)
        self.setup_ui()
        self.load_logo()
        self.running = True
        self.current_process = None

    def setup_ui(self):
        """Initialize the main application UI."""
        self.root.geometry("600x450")
        self.frame = tk.Frame(self.root, padx=20, pady=20)
        self.frame.pack(fill=tk.BOTH, expand=True)

        # Header
        header_frame = tk.Frame(self.frame)
        header_frame.pack(fill=tk.X, pady=(0, 10))
        
        self.label = tk.Label(
            header_frame, 
            text="Points inside Polygon Or LineString", 
            font=('Helvetica', 12, 'bold')
        )
        self.label.pack(side=tk.LEFT, expand=True)

        # Drag and drop area
        drop_frame = tk.LabelFrame(self.frame, text="File Input", padx=10, pady=10)
        drop_frame.pack(fill=tk.BOTH, expand=True, pady=5)
        
        drop_label = tk.Label(
            drop_frame, 
            text="Drag and drop a KML file here\nor click 'Select File' below", 
            wraplength=350
        )
        drop_label.pack(pady=10)

        # Buttons
        button_frame = tk.Frame(self.frame)
        button_frame.pack(fill=tk.X, pady=5)
        
        self.select_button = tk.Button(
            button_frame, 
            text="Select File", 
            command=self.select_file,
            width=15
        )
        self.select_button.pack(side=tk.LEFT, padx=5)

        self.convert_lines_var = tk.BooleanVar(value=True)
        self.convert_lines_check = tk.Checkbutton(
            button_frame, 
            text="Convert LineStrings", 
            variable=self.convert_lines_var,
            command=self.toggle_conversion
        )
        self.convert_lines_check.pack(side=tk.LEFT, padx=5)

        # Progress area
        progress_frame = tk.LabelFrame(self.frame, text="Progress", padx=10, pady=10)
        progress_frame.pack(fill=tk.BOTH, expand=True, pady=5)
        
        self.progress_var = tk.DoubleVar()
        self.progress_bar = ttk.Progressbar(
            progress_frame, 
            variable=self.progress_var, 
            maximum=100,
            length=400
        )
        self.progress_bar.pack(fill=tk.X, expand=True, pady=5)

        self.progress_label = tk.Label(progress_frame, text="Ready", height=1)
        self.progress_label.pack()

        self.try_again_button = tk.Button(
            progress_frame, 
            text="Process Another File", 
            command=self.reset_application,
            state=tk.DISABLED
        )
        self.try_again_button.pack(pady=5)

        # Status bar
        self.status_var = tk.StringVar()
        self.status_var.set("Ready")
        status_bar = tk.Label(
            self.frame, 
            textvariable=self.status_var, 
            bd=1, 
            relief=tk.SUNKEN,
            anchor=tk.W
        )
        status_bar.pack(fill=tk.X, pady=(5, 0))

        # Drag and drop setup
        self.root.drop_target_register(DND_FILES)
        self.root.dnd_bind("<<Drop>>", self.drop_file)

    def toggle_conversion(self):
        """Toggle the LineString to Polygon conversion setting."""
        status = "enabled" if self.convert_lines_var.get() else "disabled"
        self.status_var.set(f"LineString to Polygon conversion {status}")

    def select_file(self):
        """Open file dialog to select a KML file."""
        file_path = filedialog.askopenfilename(filetypes=[("KML files", "*.kml")])
        if file_path:
            self.process_kml(file_path)

    def process_kml(self, kml_file):
        """Process the selected KML file."""
        if not self.running:
            return
            
        try:
            self.status_var.set(f"Processing: {os.path.basename(kml_file)}")
            self.update_progress(0, "Starting processing...")
            
            converted_kml_file = None
            if self.convert_lines_var.get():
                self.update_progress(0, "Converting LineStrings to Polygons...")
                converted_kml_file = self.convert_linestrings_to_polygons(kml_file)
                if not converted_kml_file:
                    self.status_var.set("No LineStrings found to convert")

            # Start processing in a separate thread
            self.current_process = threading.Thread(
                target=self.process_file,
                args=(kml_file, converted_kml_file)
            )
            self.current_process.start()

        except Exception as e:
            self.show_error(f"Error processing file: {str(e)}")

    def process_file(self, original_kml_file, converted_kml_file):
        """Main processing function that runs in a separate thread."""
        try:
            start_time = time.time()
            
            # Parse original KML to get points
            self.update_progress(5, "Reading original KML...")
            tree_original = ET.parse(original_kml_file)
            root_original = tree_original.getroot()
            namespace = {'kml': 'http://www.opengis.net/kml/2.2'}
            
            # Get all points
            placemarks = []
            all_placemarks = root_original.findall(".//kml:Placemark", namespace)
            total_steps = len(all_placemarks) + 10  # Estimate for progress tracking
            
            for index, placemark in enumerate(all_placemarks):
                if not self.running:
                    return
                    
                name = placemark.find("kml:name", namespace)
                point = placemark.find(".//kml:Point/kml:coordinates", namespace)
                
                if name is not None and point is not None:
                    try:
                        coords = list(map(float, point.text.strip().split(",")[:2]))
                        placemarks.append((name.text.strip(), Point(coords[0], coords[1])))
                    except Exception as e:
                        print(f"Error parsing point {name.text if name else 'unnamed'}: {e}")
                
                self.update_progress(
                    5 + (index + 1) / len(all_placemarks) * 25, 
                    f"Reading points ({index+1}/{len(all_placemarks)})"
                )

            # Get original polygons
            self.update_progress(30, "Finding original polygons...")
            original_polygons = self.extract_polygons(root_original, namespace)

            # Get converted polygons (from LineStrings)
            converted_polygons = []
            if converted_kml_file:
                self.update_progress(40, "Processing converted LineStrings...")
                tree_converted = ET.parse(converted_kml_file)
                root_converted = tree_converted.getroot()
                converted_polygons = self.extract_polygons(root_converted, namespace, is_converted=True)

            # Show feature selection dialog in main thread
            polygon_names = [name for name, _ in original_polygons]
            linestring_names = [name for name, _ in converted_polygons]
            
            if not polygon_names and not linestring_names:
                self.show_warning("No Features", "No polygons or linestrings found in the KML file.")
                return
            
            self.root.after(0, lambda: self.show_selection_dialog(
                placemarks, original_polygons, converted_polygons, polygon_names, linestring_names
            ))

        except Exception as e:
            self.show_error(f"Error during processing: {str(e)}")

    def extract_polygons(self, root_element, namespace, is_converted=False):
        """Extract polygons from KML document."""
        polygons = []
        for placemark in root_element.findall(".//kml:Placemark", namespace):
            name = placemark.find("kml:name", namespace)
            polygon = placemark.find(".//kml:Polygon/kml:outerBoundaryIs/kml:LinearRing/kml:coordinates", namespace)
            
            if name is not None and polygon is not None:
                try:
                    coords_list = [
                        list(map(float, coord.strip().split(",")[:2])) 
                        for coord in polygon.text.strip().split()
                    ]
                    polygon_geom = Polygon([(c[0], c[1]) for c in coords_list])
                    display_name = name.text.strip()
                    if is_converted:
                        display_name = display_name.replace("_Polygon", "")
                    polygons.append((display_name, polygon_geom))
                except Exception as e:
                    print(f"Error parsing polygon {name.text}: {e}")
        return polygons

    def show_selection_dialog(self, placemarks, original_polygons, converted_polygons, polygon_names, linestring_names):
        """Show the feature selection dialog."""
        try:
            selected_polygons, selected_linestrings = self.show_feature_selection(polygon_names, linestring_names)
            
            if not selected_polygons and not selected_linestrings:
                self.status_var.set("Processing canceled - no features selected")
                return
            
            output_path = filedialog.asksaveasfilename(
                defaultextension=".xlsx",
                filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")],
                initialfile="Features_Points.xlsx"
            )
            
            if not output_path:
                self.status_var.set("Processing canceled - no output file selected")
                return
            
            # Start the export process
            threading.Thread(
                target=self.export_to_excel,
                args=(placemarks, original_polygons, converted_polygons, 
                     selected_polygons, selected_linestrings, output_path)
            ).start()

        except Exception as e:
            self.show_error(f"Error in selection dialog: {str(e)}")

    def export_to_excel(self, placemarks, original_polygons, converted_polygons, 
                      selected_polygons, selected_linestrings, output_path):
        """Export the results to an Excel file."""
        try:
            start_time = time.time()
            self.update_progress(60, "Preparing data for export...")
            
            # Use defaultdict for better performance with large datasets
            polygon_data = defaultdict(list)
            linestring_data = defaultdict(list)
            assigned_points = set()
            
            # Process original polygons
            for poly_name, polygon in original_polygons:
                if not self.running:
                    return
                    
                if poly_name in selected_polygons:
                    for pt_name, point in placemarks:
                        if polygon.contains(point):
                            polygon_data[poly_name].append(pt_name)
                            assigned_points.add(pt_name)
            
            # Process converted polygons (from LineStrings)
            for line_name, polygon in converted_polygons:
                if not self.running:
                    return
                    
                if line_name in selected_linestrings:
                    for pt_name, point in placemarks:
                        if polygon.contains(point):
                            linestring_data[line_name].append(pt_name)
                            assigned_points.add(pt_name)
            
            # Prepare final data structures
            self.update_progress(80, "Formatting data...")
            
            # Polygon sheet data
            polygon_output = []
            for poly_name, points in polygon_data.items():
                for pt_name in points:
                    polygon_output.append({
                        "Polygon Name": poly_name,
                        "included Point Name": pt_name
                    })
            
            # LineString sheet data
            linestring_output = []
            for line_name, points in linestring_data.items():
                for pt_name in points:
                    linestring_output.append({
                        "Converted LineString to Polygon Name": line_name,
                        "included Point Name": pt_name
                    })
            
            # Unassigned points
            unassigned_output = []
            for pt_name, _ in placemarks:
                if pt_name not in assigned_points:
                    unassigned_output.append({
                        "Point Name": pt_name,
                        "Status": "Not in any selected feature"
                    })
            
            # Write to Excel
            self.update_progress(90, "Writing to Excel...")
            with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
                if polygon_output:
                    pd.DataFrame(polygon_output).to_excel(
                        writer, 
                        sheet_name='Polygon', 
                        index=False
                    )
                
                if linestring_output:
                    pd.DataFrame(linestring_output).to_excel(
                        writer, 
                        sheet_name='LineString', 
                        index=False
                    )
                
                if unassigned_output:
                    pd.DataFrame(unassigned_output).to_excel(
                        writer, 
                        sheet_name='Unassigned', 
                        index=False
                    )
            
            elapsed_time = time.time() - start_time
            self.update_progress(100, f"Processing complete in {elapsed_time:.1f} seconds!")
            self.status_var.set(f"Completed: {os.path.basename(output_path)}")
            self.try_again_button.config(state=tk.NORMAL)
            messagebox.showinfo("Completed", f"Excel file saved at:\n{output_path}")

        except Exception as e:
            self.show_error(f"Error during export: {str(e)}")

    def convert_linestrings_to_polygons(self, input_file):
        """Convert LineStrings in KML to Polygons."""
        try:
            self.status_var.set(f"Converting LineStrings in {os.path.basename(input_file)}")
            
            with open(input_file, 'r', encoding='utf-8') as f:
                kml_root = parser.parse(f).getroot()

            ns = {'kml': 'http://www.opengis.net/kml/2.2'}
            polygons = []
            
            for placemark in kml_root.findall(".//kml:Placemark", ns):
                line_string = placemark.find(".//kml:LineString", ns)
                if line_string is not None:
                    coordinates = line_string.find(".//kml:coordinates", ns)
                    if coordinates is not None and coordinates.text:
                        coords = [coord.strip() for coord in coordinates.text.strip().split()]
                        if len(coords) >= 3:
                            coords.append(coords[0])  # Close the polygon
                            name = placemark.find(".//kml:name", ns)
                            polygons.append((
                                name.text if name is not None else "Unnamed", 
                                coords
                            ))

            if not polygons:
                return None

            temp_file = tempfile.NamedTemporaryFile(suffix='.kml', delete=False)
            temp_file.close()

            kml_doc = self.create_kml_document(polygons)
            with open(temp_file.name, "w", encoding="utf-8") as f:
                f.write(etree.tostring(kml_doc, pretty_print=True).decode("utf-8"))

            return temp_file.name

        except Exception as e:
            self.show_warning("Conversion Warning", f"Could not convert LineStrings: {str(e)}")
            return None

    def create_kml_document(self, polygons):
        """Create a KML document from polygon data."""
        kml_ns = "{http://www.opengis.net/kml/2.2}"
        kml_element = etree.Element(kml_ns + "kml")
        document = etree.SubElement(kml_element, kml_ns + "Document")
        
        for name, coords in polygons:
            placemark = etree.SubElement(document, kml_ns + "Placemark")
            name_element = etree.SubElement(placemark, kml_ns + "name")
            name_element.text = name + "_Polygon"
            
            polygon = etree.SubElement(placemark, kml_ns + "Polygon")
            outer_boundary = etree.SubElement(polygon, kml_ns + "outerBoundaryIs")
            linear_ring = etree.SubElement(outer_boundary, kml_ns + "LinearRing")
            coordinates = etree.SubElement(linear_ring, kml_ns + "coordinates")
            coordinates.text = "\n".join(coords)
        
        return kml_element

    def show_feature_selection(self, polygons, linestrings):
        """Show the feature selection window."""
        selection_window = tk.Toplevel(self.root)
        selection_window.title("Select Features to Process")
        selection_window.geometry("700x600")
        selection_window.transient(self.root)
        selection_window.grab_set()
        
        notebook = ttk.Notebook(selection_window)
        notebook.pack(fill='both', expand=True, padx=10, pady=10)
        
        # Polygon tab
        polygon_frame = self.create_feature_selection_tab(notebook, "Polygons", polygons)
        
        # LineString tab
        linestring_frame = self.create_feature_selection_tab(notebook, "LineStrings", linestrings)
        
        selected_polygons = []
        selected_linestrings = []
        
        def process_selected():
            nonlocal selected_polygons, selected_linestrings
            selected_polygons = [name for name, var in polygon_frame.vars if var.get()]
            selected_linestrings = [name for name, var in linestring_frame.vars if var.get()]
            selection_window.destroy()
        
        button_frame = tk.Frame(selection_window)
        button_frame.pack(fill=tk.X, padx=10, pady=10)
        
        tk.Button(
            button_frame, 
            text="Process Selected", 
            command=process_selected,
            width=15
        ).pack(side=tk.RIGHT)
        
        tk.Button(
            button_frame, 
            text="Cancel", 
            command=selection_window.destroy,
            width=15
        ).pack(side=tk.RIGHT, padx=5)
        
        selection_window.wait_window()
        return selected_polygons, selected_linestrings

    def create_feature_selection_tab(self, notebook, tab_name, features):
        """Create a tab in the feature selection notebook."""
        frame = tk.Frame(notebook)
        notebook.add(frame, text=f"{tab_name} ({len(features)})")
        
        # Header with select all/none
        header_frame = tk.Frame(frame)
        header_frame.pack(fill=tk.X, padx=5, pady=5)
        
        select_all_var = tk.BooleanVar(value=True)
        tk.Checkbutton(
            header_frame, 
            text=f"Select All {tab_name}", 
            variable=select_all_var,
            command=lambda: self.toggle_all(frame.vars, select_all_var.get())
        ).pack(side=tk.LEFT)
        
        search_frame = tk.Frame(header_frame)
        search_frame.pack(side=tk.RIGHT)
        
        search_label = tk.Label(search_frame, text="Search:")
        search_label.pack(side=tk.LEFT)
        
        search_var = tk.StringVar()
        search_entry = tk.Entry(search_frame, textvariable=search_var, width=20)
        search_entry.pack(side=tk.LEFT, padx=5)
        search_entry.bind("<KeyRelease>", lambda e: self.filter_features(frame.scrollable_frame, search_var.get().lower()))
        
        # Scrolled list of features
        canvas = tk.Canvas(frame)
        scrollbar = ttk.Scrollbar(frame, orient="vertical", command=canvas.yview)
        scrollable_frame = tk.Frame(canvas)
        
        scrollable_frame.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        
        vars = []
        unique_features = sorted(set(features))
        
        for feature in unique_features:
            var = tk.BooleanVar(value=True)
            cb = tk.Checkbutton(
                scrollable_frame, 
                text=feature, 
                variable=var,
                anchor="w",
                width=50
            )
            cb.pack(anchor="w", padx=20, pady=2)
            vars.append((feature, var))
        
        frame.vars = vars
        frame.scrollable_frame = scrollable_frame
        return frame

    def toggle_all(self, vars, state):
        """Toggle all checkboxes in the feature selection list."""
        for _, var in vars:
            var.set(state)

    def filter_features(self, frame, search_text):
        """Filter features in the selection list based on search text."""
        for widget in frame.winfo_children():
            if isinstance(widget, tk.Checkbutton):
                text = widget.cget("text").lower()
                if search_text in text:
                    widget.pack(anchor="w", padx=20, pady=2)
                else:
                    widget.pack_forget()

    def update_progress(self, value, message=None):
        """Update the progress bar and status message."""
        if not self.running:
            return
            
        self.progress_var.set(value)
        if message:
            self.progress_label.config(text=message)
        self.root.update_idletasks()

    def drop_file(self, event):
        """Handle file drop event."""
        file_path = event.data.strip().strip('{}')
        if file_path.lower().endswith('.kml'):
            self.process_kml(file_path)
        else:
            self.show_error("Please drop a KML file")

    def reset_application(self):
        """Reset the application state for new processing."""
        self.progress_var.set(0)
        self.progress_bar.update()
        self.progress_label.config(text="Ready")
        self.try_again_button.config(state=tk.DISABLED)
        self.status_var.set("Ready for new file")

    def load_logo(self):
        """Attempt to load application logo from various paths."""
        logo_paths = ["ZTE LOGO.png", "logo.png", "icon.png"]
        for path in logo_paths:
            if os.path.exists(path):
                try:
                    img = Image.open(path)
                    img = img.resize((32, 32))
                    logo_icon = ImageTk.PhotoImage(img)
                    self.root.iconphoto(False, logo_icon)
                    break
                except Exception as e:
                    print(f"Could not load logo {path}: {e}")

    def show_error(self, message):
        """Show an error message."""
        self.root.after(0, lambda: messagebox.showerror("Error", message))
        self.status_var.set(f"Error: {message}")
        self.reset_application()

    def show_warning(self, title, message):
        """Show a warning message."""
        self.root.after(0, lambda: messagebox.showwarning(title, message))
        self.status_var.set(f"Warning: {message}")

    def on_close(self):
        """Handle application close event."""
        self.running = False
        if self.current_process and self.current_process.is_alive():
            self.status_var.set("Waiting for process to finish...")
            self.current_process.join(timeout=2)
        self.root.destroy()


if __name__ == "__main__":
    root = TkinterDnD.Tk()
    root.title("Points inside Polygon Or LineString")
    app = KMLToExcelConverter(root)
    root.mainloop()