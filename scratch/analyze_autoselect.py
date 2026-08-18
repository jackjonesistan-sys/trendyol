import sys
import json
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

# Simulate autoSelect in JS
# Rule 1: Priority rule active (["Avantajlı", "Flaş", "Plus"])
# When priority campaign exists in eligible -> pick it
# If NONE of priority campaigns exist, but Komisyon Tarifesi is eligible -> pick Komisyon Tarifesi!
priority = ["Avantajlı", "Flaş", "Plus"]
sel_with_rule = {"Avantajlı": 0, "Flaş": 0, "Plus": 0, "Komisyon Tarifesi": 0, "Hiçbiri": 0}
for r in results:
    eligible = r.get("eligible_main_campaigns", [])
    rec = None
    for p in priority:
        if p in eligible:
            rec = p
            break
    if not rec:
        if "Komisyon Tarifesi" in eligible:
            rec = "Komisyon Tarifesi"
        else:
            rec = "Hiçbiri"
    sel_with_rule[rec] += 1

print("Önerilenleri Seç (Öncelik Kuralı Aktifken: Avantajlı > Flaş > Plus > Komisyon Tarifesi):")
for k, v in sel_with_rule.items():
    print(f"  - {k}: {v} ürün")

# Rule 2: Priority rule disabled (Best net profit)
sel_no_rule = {"Avantajlı": 0, "Flaş": 0, "Plus": 0, "Komisyon Tarifesi": 0, "Hiçbiri": 0}
for r in results:
    eligible = r.get("eligible_main_campaigns", [])
    metrics = {
        'Avantajlı': ('Avantajlı Ürün Kalan Net (TL)', 'Avantajlı Ürün Fiyatı (YENİ TSF) (TL)'),
        'Flaş': ('Flaş Ürün Kalan Net (TL)', 'Flaş Ürün 24 Saat Fiyatı (TL)'),
        'Plus': ('Plus Net (TL)', 'Plus Fiyatı (TL)'),
        'Komisyon Tarifesi': ('Komisyon Tarifesi Net (TL)', 'Komisyon Tarifesi Fiyatı (TL)')
    }
    best = None
    for c in ['Avantajlı', 'Flaş', 'Plus', 'Komisyon Tarifesi']:
        if c in eligible:
            net = r.get(metrics[c][0])
            price = r.get(metrics[c][1])
            if net is not None and not pd.isna(net):
                candidate = {"campaign": c, "net": float(net), "price": float(price or 0)}
                if not best or candidate["net"] > best["net"]:
                    best = candidate
                elif candidate["net"] == best["net"] and candidate["price"] > best["price"]:
                    best = candidate
    rec = best["campaign"] if best else "Hiçbiri"
    sel_no_rule[rec] += 1

print("\nÖnerilenleri Seç (Öncelik Kuralı Pasifken - En Yüksek Net Getiri):")
for k, v in sel_no_rule.items():
    print(f"  - {k}: {v} ürün")
