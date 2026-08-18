import sys
import pandas as pd
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
from komisyon_hesaplayici import calculate_all

base_dir = Path(r"c:\Users\Tasarımcı\Desktop\trendyol\Girdiler\Yuklenen")
inputs = {
    "discount": base_dir / "discount.xlsx",
    "commission": base_dir / "commission.xlsx",
    "current": base_dir / "current.xlsx",
    "advantage": base_dir / "advantage.xlsx",
    "flash": base_dir / "flash.xlsx",
    "plus": base_dir / "plus.xlsx",
}

out_dir = Path(r"c:\Users\Tasarımcı\Desktop\trendyol\scratch\analysis_output")
res = calculate_all(inputs, output_dir=out_dir)
results = res["results"]

# Analysis of Komisyon Tarifesi eligible
eligible_kom = [r for r in results if "Komisyon Tarifesi" in r.get("eligible_campaigns", [])]
print(f"Toplam Hesaplanan: {len(results)}")
print(f"Uygun Komisyon Tarifesi Olan Ürün Sayısı (Net Bazlı): {len(eligible_kom)}")

# Kademe dağılımı
df_kom = pd.read_excel(inputs["commission"])
b_col = "Barkod" if "Barkod" in df_kom.columns else "BARKOD"
kom_dict = {str(r[b_col]).strip(): r for _, r in df_kom.iterrows() if pd.notna(r[b_col])}

k_counts = {1: 0, 2: 0, 3: 0, 4: 0}
sub_dip_count = 0
for r in eligible_kom:
    b = str(r["Barkod"]).strip()
    p = float(r["Komisyon Tarifesi Fiyatı (TL)"])
    dip = float(r.get("Düşülebilecek Dip Fiyat (TL)", 0) or 0)
    if dip > 0 and p < dip:
        sub_dip_count += 1
    
    kom_row = kom_dict.get(b)
    if kom_row is not None:
        for k, col in [(4, "4.Fiyat Üst Limiti"), (3, "3.Fiyat Üst Limiti"), (2, "2.Fiyat Üst Limiti"), (1, "1.Fiyat Alt Limit")]:
            val = kom_row.get(col)
            if val is not None and not pd.isna(val) and abs(float(val) - p) < 0.01:
                k_counts[k] += 1
                break

print(f"Dip Fiyatın Altında Olup Yeni Optimize Dip Fiyat Olarak Tanımlanan Ürün Sayısı: {sub_dip_count}")
print("Kademe Dağılımı:")
for k, count in sorted(k_counts.items(), reverse=True):
    print(f"  - {k}. Kademe: {count} ürün")
