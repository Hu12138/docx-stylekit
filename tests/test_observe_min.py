from pathlib import Path
from docx import Document
import subprocess, sys

def test_observe_runs(tmp_path):
    sample = tmp_path / "small.docx"
    doc = Document()
    doc.add_paragraph("Test content")
    doc.save(sample)
    out = tmp_path / "observed.yaml"
    cmd = [sys.executable, "-m", "docx_stylekit.cli", "observe", str(sample), "-o", str(out)]
    subprocess.run(cmd, check=True)
    assert out.exists() and out.stat().st_size > 0
