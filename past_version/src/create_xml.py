import mammoth
import xml.etree.ElementTree as ET

def create_xml(filepath, output_folder, filename):
    with open(filepath, "rb") as docx_file:
        result = mammoth.convert_to_html(docx_file)
    html = result.value

    html = f"<xml>{html}</xml>"  # Wrap the HTML in a root element
    root = ET.fromstring(html)
    xml = ET.ElementTree(root)

    with open(f"{output_folder}/{filename}.xml", "w", encoding="utf-8") as xml_file:
        xml_file.write(ET.tostring(xml.getroot(), encoding="utf-8").decode("utf-8"))