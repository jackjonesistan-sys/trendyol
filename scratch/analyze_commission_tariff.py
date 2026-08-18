import os
import sys
import json
import pandas as pd
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
from komisyon_hesaplayici import calculate_all

base_dir = Path(r"c:\Users\Tasarımcı\Desktop\trendyol\Girdiler\Yuklenen")
discount_file = base_dir / "discount.xlsx"
commission_file = base_dir / "commission.xlsx"
current_file = base_dir / "current.xlsx"
advantage_file = base_dir / "advantage.xlsx"
flash_file = base_dir / "flash.xlsx"
plus_file = base_dir / "plus.xlsx"

inputs = {
    "discount": discount_file,
    "commission": commission_file,
    "current": current_file,
    "advantage": advantage_file,
    "flash": flash_file,
    "plus": plus_file,
}

out_dir = Path(r"c:\Users\Tasarımcı\Desktop\trendyol\scratch\analysis_output")
out_dir.mkdir(parents=True, exist_ok=True)

res = calculate_all(inputs, output_dir=out_dir)

results = res["results"]
print(f"Toplam Hesaplanan Urun Sayisi: {len(results)}")

total_kom_in_file = 0
kom_eligible_count = 0
kom_selected_as_main = 0
eligible_samples = []

for row in results:
    kom_price = row.get("Komisyon Tarifesi Fiyatı (TL)")
    kom_net = row.get("Komisyon Tarifesi Net (TL)")
    kom_comm = row.get("Komisyon Tarifesi Komisyon (%)")
    kom_sel = row.get("Komisyon Tarifesi Seçimi")
    initial_sel = row.get("İlk Kampanya Seçimi")
    eligible_campaigns = row.get("eligible_campaigns", [])
    
    if kom_price is not None and not pd.isna(kom_price):
        total_kom_in_file += 1
        is_eligible = "Komisyon Tarifesi" in eligible_campaigns
        if is_eligible:
            kom_eligible_count += 1
            if initial_sel == "Komisyon Tarifesi":
                kom_selected_as_main += 1
            
            if len(eligible_samples) < 10:
                eligible_samples.append({
                    "Barkod": row.get("Barkod"),
                    "Guncel Fiyat": row.get("Güncel Ürün Fiyatı (TL)"),
                    "Guncel Net": row.get("Güncel Ürün Kalan Net (TL)"),
                    "Dip Fiyat": row.get("Düşülebilecek Dip Fiyat (TL)"),
                    "Dip Net": row.get("Düşülebilecek Dip Net (TL)"),
                    "Kom Fiyat": kom_price,
                    "Kom Oran": kom_comm,
                    "Kom Net": kom_net,
                    "Ilk Secim": initial_sel,
                    "Eligible": eligible_campaigns,
                })

print(f"Komisyon Tarifesi Tablosundaki Toplam Urun: {total_kom_in_file}")
print(f"Dip Fiyat/Net Kriterine Gore Uygun Komisyon Tarifesi Olan Urun Sayisi: {kom_eligible_count}")
print(f"Onerilen / Ilk Secim Olarak Komisyon Tarifesi Secilen: {kom_selected_as_main}")

print("\nOrnek 5 Urun:")
for s in eligible_samples[:5]:
    print(json.dumps(s, ensure_ascii=False, indent=2))
