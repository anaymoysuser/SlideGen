"""
recolor_theme  — replace the color scheme of a PowerPoint template.

Steps:
1. Load the template .pptx.
2. Locate the theme XML (usually ppt/theme/theme1.xml).
3. Overwrite <a:clrScheme> with a new RGB palette.
4. Save the presentation under a new file name.
 
"""

from pathlib import Path
from pptx import Presentation
from lxml import etree  # python-pptx already depends on lxml
 
TEMPLATE_PATH = Path("my_template.pptx")         
OUTPUT_PATH   = Path("my_template_recolored.pptx") 

 
# new RGB values (hex, without the leading '#')
# keys must follow the ECMA-376 standard: dk1, lt1, dk2, lt2,
# accent1–accent6, hlink, folHlink
 
NEW_COLORS = {
    "dk1":      "000000",
    "lt1":      "FFFFFF",
    "dk2":      "003049",
    "lt2":      "F8F9FA",
    "accent1":  "1F78B4",
    "accent2":  "FF9800",
    "accent3":  "FBC02D",
    "accent4":  "4CAF50",
    "accent5":  "26C6DA",
    "accent6":  "F06365",
    "hlink":    "004CFF",
    "folHlink": "7B1FA2",
}

 
def replace_color(node, rgb_hex, ns_uri):
    """Remove existing children and insert a new <a:srgbClr val="..."/>."""
    for child in list(node):
        node.remove(child)
    srgb = etree.SubElement(node, f"{{{ns_uri}}}srgbClr")
    srgb.set("val", rgb_hex.upper())

def main():
    presentation = Presentation(TEMPLATE_PATH) 
    # locate the first theme part (theme1.xml)
    theme_part = next(
        p for p in presentation.part.related_parts.values()
        if p.partname.endswith("theme/theme1.xml")
    )

    # parse XML
    theme_tree = etree.fromstring(theme_part.blob)
    NSMAP = {"a": "http://schemas.openxmlformats.org/drawingml/2006/main"}

    color_scheme = theme_tree.find(".//a:clrScheme", namespaces=NSMAP)
    namespace_uri = NSMAP["a"]

    # iterate over dk1 / lt1 / accent1 … nodes
    for child in color_scheme:
        tag = etree.QName(child).localname
        if tag in NEW_COLORS:
            replace_color(child, NEW_COLORS[tag], namespace_uri)

    # write back
    theme_part._blob = etree.tostring(
        theme_tree, encoding="UTF-8", xml_declaration=True
    )

    presentation.save(OUTPUT_PATH)
    print(f"Color scheme replaced → {OUTPUT_PATH.resolve()}")




if __name__ == "__main__":
    main()