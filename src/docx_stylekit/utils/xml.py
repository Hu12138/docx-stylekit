from lxml import etree as ET
from ..constants import NS

def parse_bytes(xml_bytes):
    return ET.fromstring(xml_bytes)

def find(node, xpath):
    return node.find(xpath, namespaces=NS)

def findall(node, xpath):
    return node.findall(xpath, namespaces=NS)

def attr(node, name, default=None):
    if node is None:
        return default
    return node.get(name, default)

def text(node, default=None):
    if node is None or node.text is None:
        return default
    return node.text
