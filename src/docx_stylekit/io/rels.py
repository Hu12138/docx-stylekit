from lxml import etree as ET
from ..constants import NS

def parse_document_rels(xml_bytes):
    root = ET.fromstring(xml_bytes)
    rels = {}
    for el in root.findall("rels:Relationship", namespaces=NS):
        rId = el.get("Id")
        target = el.get("Target")
        rels[rId] = {"type": el.get("Type"), "target": target}
    return rels
