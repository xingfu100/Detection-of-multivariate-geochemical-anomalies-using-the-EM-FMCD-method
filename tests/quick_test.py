"""Check that the distributed source files parse without launching their GUIs."""
import ast
from pathlib import Path

source_dir = Path(__file__).resolve().parents[1] / "src"
files = ["01_preprocessing.py", "02_aic_element_selection.py", "03_em_clustering.py", "04_robust_pca.py", "05_fmcd.py"]
for name in files:
    path = source_dir / name
    ast.parse(path.read_text(encoding="utf-8-sig"), filename=str(path))
    print(f"PASS: {name}")
print("Syntax check passed; numerical and GUI behavior were not tested.")
