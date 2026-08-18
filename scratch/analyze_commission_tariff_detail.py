import os
import sys
import json
import pandas as pd
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
from komisyon_hesaplayici import calculate_all, evaluate_commission_tariff_bracket

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

# Read original commission.xlsx directly to get raw row info
df_kom_raw = pd.read_excel(inputs["commission"])
b_col = "Barkod" if "Barkod" in df_kom_raw.columns else "BARKOD"
total_raw_rows = len(df_kom_raw)
raw_barcodes = set(df_kom_raw[b_col].dropna().astype(str).str.strip())

# Analyze each product
kom_in_calc = 0
kom_eligible = 0
only_kom_eligible = 0
kom_and_others_eligible = 0
kademe_distribution = {1: 0, 2: 0, 3: 0, 4: 0}
bracket_stats = []

for row in results:
    b = str(row.get("Barkod", "")).strip()
    if b not in raw_barcodes:
        continue
    
    kom_in_calc += 1
    kom_price = row.get("Komisyon Tarifesi Fiyatı (TL)")
    kom_net = row.get("Komisyon Tarifesi Net (TL)")
    kom_comm = row.get("Komisyon Tarifesi Komisyon (%)")
    eligible_campaigns = row.get("eligible_campaigns", [])
    
    if "Komisyon Tarifesi" in eligible_campaigns:
        kom_eligible += 1
        other_mains = [c for c in eligible_campaigns if c in ("Avantajlı", "Flaş", "Plus")]
        if other_mains:
            kom_and_others_eligible += 1
        else:
            only_kom_eligible += 1
        
        # Check which bracket was chosen
        # Compare kom_price with raw row kademe prices
        raw_rows = df_kom_raw[df_kom_raw[b_col].astype(str).str.strip() == b]
        if not raw_rows.empty:
            r = raw_rows.iloc[0].to_dict()
            chosen_k = None
            for k, col in [(4, "4.Fiyat Üst Limiti"), (3, "3.Fiyat Üst Limiti"), (2, "2.Fiyat Üst Limiti"), (1, "1.Fiyat Alt Limit")]:
                val = r.get(col)
                if val is not None and not pd.isna(val) and abs(float(val) - float(kom_price)) < 0.01:
                    chosen_k = k
                    break
            if chosen_k:
                kademe_distribution[chosen_k] += 1

print("="*60)
print("ÜRÜN KOMİSYON TARİFELERİ ANALİZ RAPORU")
print("="*60)
print(f"1. Ürün Komisyon Tarifeleri Dosyasındaki Toplam Satır Sayısı : {total_raw_rows}")
print(f"2. Hesaplamaya Dahil Olan Ürün Sayısı                     : {kom_in_calc}")
print(f"3. Dip Fiyat/Net Kriterlerine Uyan (Uygun Kademeli) Ürün : {kom_eligible} (oran: %{kom_eligible/kom_in_calc*100:.1f})")
print(f"   - SADECE Komisyon Tarifesi Uygun Olan Ürün Sayısı     : {only_kom_eligible}")
print(f"   - Diğer Kampanyalarla Birlikte Uygun Olan Ürün Sayısı  : {kom_and_others_eligible}")
print(f"4. Dip Fiyat/Net Kriterine Uymayan (Dip Altı Kalan) Ürün : {kom_in_calc - kom_eligible}")
print("\n5. Seçilen Kademe Dağılımı (Müşteriye En Düşük Fiyatı Sunan):")
for k, count in sorted(kademe_distribution.items(), reverse=True):
    print(f"   - {k}. Kademe: {count} ürün")
print("="*60)
