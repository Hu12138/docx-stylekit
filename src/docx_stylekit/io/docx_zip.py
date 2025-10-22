import zipfile
from pathlib import Path

class DocxZip:
    def __init__(self, path):
        self.path = Path(path)
        self.zf = zipfile.ZipFile(self.path, "r")

    def read_xml(self, member):
        return self.zf.read(member)

    def has(self, member):
        try:
            self.zf.getinfo(member)
            return True
        except KeyError:
            return False

    def close(self):
        self.zf.close()

    # 常用部件路径
    @staticmethod
    def parts():
        return {
            "document": "word/document.xml",
            "styles": "word/styles.xml",
            "numbering": "word/numbering.xml",
            "settings": "word/settings.xml",
            "theme": "word/theme/theme1.xml",
            "doc_rels": "word/_rels/document.xml.rels",
        }

    def list_headers(self):
        return [n for n in self.zf.namelist() if n.startswith("word/header") and n.endswith(".xml")]

    def list_footers(self):
        return [n for n in self.zf.namelist() if n.startswith("word/footer") and n.endswith(".xml")]
