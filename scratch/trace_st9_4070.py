import sys, os
sys.path.insert(0, r'c:\Users\Tasarımcı\Desktop\trendyol')
from pathlib import Path
import pandas as pd
from input_files import load_upload_set, load_counter_configs, load_plus_extra_configs
from komisyon_hesaplayici import calculate_all
from app import UPLOAD_DIR, INPUT_MANIFEST, BASE_DIR

input_files = load_upload_set(UPLOAD_DIR, INPUT_MANIFEST)
target_bc = "ST9-4070"

print(f"=== TRACING BARCODE: {target_bc} ACROSS ALL FILES ===")

for key, p in input_files.items():
    if not p or not os.path.exists(p): continue
    try:
        df = pd.read_excel(p)
        bcol = [c for c in df.columns if "barkod" in str(c).lower()]
        if bcol:
            found = df[df[bcol[0]].astype(str).str.strip() == target_bc]
            if not found.empty:
                print(f"\n[+] Found in file '{key}' ({Path(p).name}):")
                for col in found.columns:
                    val = found[col].values[0]
                    if pd.notna(val):
                        print(f"     {col}: {val}")
    except Exception as e:
        print(f"Error checking {key}: {e}")

# Check all excel files in Girdiler/Yuklenen
print("\n--- Checking all excel files in UPLOAD_DIR recursively ---")
for root_dir, dirs, files in os.walk(UPLOAD_DIR):
    for fn in files:
        if fn.endswith(".xlsx"):
            fpath = os.path.join(root_dir, fn)
            try:
                df = pd.read_excel(fpath)
                bcol = [c for c in df.columns if "barkod" in str(c).lower()]
                if bcol:
                    found = df[df[bcol[0]].astype(str).str.strip() == target_bc]
                    if not found.empty:
                        print(f"\n[+] Found in File '{fn}' (Path: {fpath}):")
                        for col in found.columns:
                            val = found[col].values[0]
                            if pd.notna(val):
                                print(f"     {col}: {val}")
            except Exception as e:
                pass
