from pathlib import Path
import subprocess, sys

def test_observe_runs():
    sample = Path(__file__).parent / "assets" / "small.docx"
    out = Path(__file__).parent / "observed.yaml"
    cmd = [sys.executable, "-m", "docx_stylekit.cli", "observe", str(sample), "-o", str(out)]
    subprocess.run(cmd, check=True)
    assert out.exists() and out.stat().st_size > 0
