from pykml.factory import KML_ElementMaker as KML
from pykml import parser
import zipfile
from shapely.geometry import Polygon
from lxml import etree  # ✅ هذا هو السطر الناقص

def read_kml_kmz(file_path):
    if file_path.lower().endswith(".kmz"):
        with zipfile.ZipFile(file_path, 'r') as kmz:
            kml_file = [f for f in kmz.namelist() if f.endswith('.kml')][0]
            with kmz.open(kml_file) as kml_content:
                return parser.parse(kml_content).getroot()
    else:
        with open(file_path, 'r', encoding='utf-8') as f:
            return parser.parse(f).getroot()

def convert_path_to_polygon(kml_root):
    """ يحول المسارات إلى مضلعات عن طريق إغلاق النقاط """
    ns = {'kml': 'http://www.opengis.net/kml/2.2'}
    
    polygons = []
    for placemark in kml_root.findall(".//kml:Placemark", ns):
        line_string = placemark.find(".//kml:LineString", ns)
        if line_string is not None:
            coordinates = line_string.find(".//kml:coordinates", ns).text.strip()
            coords = [coord.strip() for coord in coordinates.split()]
            if len(coords) >= 3:  # لازم يكون عندي 3 نقاط على الأقل عشان المضلع
                coords.append(coords[0])  # إغلاق المضلع بتكرار أول نقطة
                name = placemark.find(".//kml:name", ns)
                polygons.append((name.text if name is not None else "Unnamed", coords))
    return polygons

def write_kml(output_file, polygons):
    """ يكتب المضلع إلى ملف KML بصيغة XML صحيحة """
    kml_doc = KML.kml(
        KML.Document(
            *[
                KML.Placemark(
                    KML.name(name + "_Polygon"),
                    KML.Polygon(
                        KML.outerBoundaryIs(
                            KML.LinearRing(
                                KML.coordinates("\n".join(coords))
                            )
                        )
                    )
                )
                for name, coords in polygons
            ]
        )
    )

    with open(output_file, "w", encoding="utf-8") as f:
        f.write(etree.tostring(kml_doc, pretty_print=True).decode("utf-8"))


def convert_coords_to_polygon(coordinates_str):
    """Converts raw coordinates string to a closed polygon"""
    coords = [coord.strip() for coord in coordinates_str.split()]
    if len(coords) >= 3:
        coords.append(coords[0])  # Close the polygon
        return coords
    return None


if __name__ == "__main__":
    input_file = "input.kml"  # استبدلها بملف KML/KMZ الخاص بك
    output_file = "output.kml"
    
    kml_root = read_kml_kmz(input_file)
    polygons = convert_path_to_polygon(kml_root)
    write_kml(output_file, polygons)
    
    print(f"✅ التحويل اكتمل! تم حفظ الملف في: {output_file}")
