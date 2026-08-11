import os
from pathlib import Path
import pandas as pd
from komisyon_hesaplayici import calculate_komisyon

uploads_dir = Path("uploads")
files = {
    'discount': uploads_dir / "discount.xlsx",
    'commission': uploads_dir / "commission.xlsx",
    'current': uploads_dir / "current.xlsx",
    'advantage': uploads_dir / "advantage.xlsx" if (uploads_dir / "advantage.xlsx").exists() else None,
    'flash': uploads_dir / "flash.xlsx" if (uploads_dir / "flash.xlsx").exists() else None,
    'plus': uploads_dir / "plus.xlsx" if (uploads_dir / "plus.xlsx").exists() else None,
    'plus_extra': uploads_dir / "plus_extra.xlsx" if (uploads_dir / "plus_extra.xlsx").exists() else None,
}

target_barcodes = [
    'KOKO3360-kpmk429', 'KOKO5090-kpmk429',
    'LAZERTOZMODEL1-SIYAH-R-40x70', 'LAZERTOZMODEL1-SIYAH-R-50x90',
    'NA100OZELK-150200', 'NA100OZELK-150300',
    'TA200-8-3360', 'TA200-8-4060', 'TA200-8-4070', 'TA200-8-5090',
    'kpbn4294060'
]

results, error = calculate_komisyon(files)
if error:
    print("Error:", error)
else:
    df_res = pd.DataFrame(results)
    for bc in target_barcodes:
        row = df_res[df_res['Barkod'] == bc]
        if row.empty:
            print(f"Barcode {bc}: NOT FOUND in calculated results!")
        else:
            r = row.iloc[0]
            print(f"--- Barcode: {bc} ---")
            print("  Katılabilir Kampanyalar:", r.get('eligible_campaigns'))
            print("  Düşülebilecek Dip Fiyat:", r.get('Düşülebilecek Dip Fiyat (TL)'))
            print("  İndirim Uygulanabilir:", r.get('İndirim Uygulanabilir'))
            print("  Güncel Fiyat:", r.get('Güncel Ürün Fiyatı (TL)'))
            print("  Plus Ek İndirim 5 Fiyatı:", r.get('Plus Ek Fiyatı %5 (TL)'))
            print("  Plus Ek İndirim 5 Net:", r.get('Plus Ek Net %5 (TL)'))
            # Check if present in input files
            for fname in ['plus_extra', 'plus', 'advantage', 'flash']:
                if files.get(fname):
                    try:
                        df_f = pd.read_excel(files[fname])
                        b_col = 'Barkod' if 'Barkod' in df_f.columns else ('BARKOD' if 'BARKOD' in df_f.columns else None)
                        if b_col:
                            found = df_f[df_f[b_col].astype(str).str.strip() == bc]
                            if not found.empty:
                                print(f"  Found in {fname}.xlsx: YES")
                            else:
                                print(f"  Found in {fname}.xlsx: NO")
                    except Exception as e:
                        print(f"  Error checking {fname}: {e}")
