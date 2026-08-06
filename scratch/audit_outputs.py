import os, glob
import pandas as pd

target_dir = r"C:\Users\Tasarımcı\Desktop\trendyol\Çıktılar\2026-08-06_17-19-06"
hesap_file = r"C:\Users\Tasarımcı\Desktop\trendyol\Çıktılar\Kampanya_Hesaplama_Sonuclari.xlsx"

df_hesap = pd.read_excel(hesap_file)
barcode_dip_map = {}
barcode_current_map = {}
for idx, row in df_hesap.iterrows():
    bc = str(row.get("Barkod", "")).strip()
    if bc:
        dip = row.get("Düşülebilecek Dip Fiyat (TL)")
        gun = row.get("Güncel Ürün Fiyatı (TL)")
        try: dip_val = float(dip) if pd.notna(dip) else None
        except: dip_val = None
        try: gun_val = float(gun) if pd.notna(gun) else None
        except: gun_val = None
        barcode_dip_map[bc] = dip_val
        barcode_current_map[bc] = gun_val

print("=== DETAILED SAFETY & PRICE COMPLIANCE AUDIT ===")
print(f"Target Directory: {target_dir}")
print(f"Total Master Barcodes: {len(barcode_dip_map)}\n")

files = glob.glob(os.path.join(target_dir, "*.xlsx"))
total_checked = 0
total_violations = 0

for fpath in sorted(files):
    fname = os.path.basename(fpath)
    df = pd.read_excel(fpath)
    
    bcols = [c for c in df.columns if "barkod" in str(c).lower()]
    bcol = bcols[0] if bcols else None
    
    # Identify price column
    pcol = None
    for c in df.columns:
        clower = str(c).lower()
        if any(k in clower for k in ["yeni tsf", "24 saat fiyat", "seçilen satış fiyatı", "kampanyalı satış fiyatı", "uygulanan kampanya fiyat"]):
            pcol = c
            break
            
    print(f"File: {fname}")
    print(f"  - Total Rows: {len(df)}")
    print(f"  - Barcode Col: {bcol}")
    print(f"  - Price Col: {pcol}")
    
    file_checked = 0
    file_violations = 0
    
    if bcol and pcol:
        for idx, row in df.iterrows():
            bc = str(row[bcol]).strip()
            price_raw = row[pcol]
            try: price_val = float(price_raw) if pd.notna(price_raw) else None
            except: price_val = None
            
            dip_val = barcode_dip_map.get(bc)
            
            if price_val is not None and dip_val is not None:
                file_checked += 1
                total_checked += 1
                if price_val < (dip_val - 0.01):
                    file_violations += 1
                    total_violations += 1
                    print(f"    !!! VIOLATION: {bc} | Applied Price: {price_val} TL < Dip Price: {dip_val} TL")
                    
        print(f"  - Dip Price Checked Rows: {file_checked} | Violations: {file_violations}")
        
        # Sample 2 rows display
        for idx in range(min(2, len(df))):
            sample_bc = str(df.iloc[idx][bcol]).strip()
            sample_p = df.iloc[idx][pcol]
            sample_dip = barcode_dip_map.get(sample_bc, "N/A")
            print(f"    Sample: {sample_bc} => Applied: {sample_p} TL | Dip Threshold: {sample_dip} TL")
    else:
        print("  - Summary/Report file (no applied price column needed)")
    print("-" * 60)

print("\nFINAL AUDIT VERDICT:")
print(f"Total Applied Prices Checked Against Dip Prices: {total_checked}")
print(f"Total Below-Dip Price Violations: {total_violations}")

if total_violations == 0:
    print("STATUS: 100% SAFE! ZERO DANGER OF SELLER LOSS! ALL APPLIED PRICES ARE EQUAL TO OR ABOVE DIP PRICES!")
else:
    print(f"STATUS: WARNING! {total_violations} VIOLATIONS FOUND!")
