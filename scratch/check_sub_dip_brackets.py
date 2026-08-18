import sys
import json
import pandas as pd
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
from komisyon_hesaplayici import calculate_all, to_float

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

# Load raw commission.xlsx
df_kom = pd.read_excel(inputs["commission"])
b_col = "Barkod" if "Barkod" in df_kom.columns else "BARKOD"
kom_dict = {str(r[b_col]).strip(): r for _, r in df_kom.iterrows() if pd.notna(r[b_col])}

# Let's check products where bracket_price <= dip_price and bracket_net >= dip_net
matching_cases = []
total_analyzed = 0

for row in results:
    b = str(row.get("Barkod", "")).strip()
    if b not in kom_dict:
        continue
    
    total_analyzed += 1
    dip_price = row.get("Düşülebilecek Dip Fiyat (TL)")
    dip_net = row.get("Düşülebilecek Dip Net (TL)")
    guncel_fiyat = row.get("Güncel Ürün Fiyatı (TL)")
    guncel_net = row.get("Güncel Ürün Kalan Net (TL)")
    guncel_kom = row.get("Güncel Ürün Komisyon (%)")
    
    if dip_price is None or pd.isna(dip_price) or dip_price <= 0:
        continue
    
    # Calculate dip_net if not present
    if dip_net is None or pd.isna(dip_net):
        if guncel_kom is not None:
            dip_net = round(dip_price * (1 - guncel_kom / 100.0), 2)
        else:
            continue
    
    kom_row = kom_dict[b]
    brackets = [
        ("4. Kademe", to_float(kom_row.get("4.Fiyat Üst Limiti")), to_float(kom_row.get("4.KOMİSYON"))),
        ("3. Kademe", to_float(kom_row.get("3.Fiyat Üst Limiti")), to_float(kom_row.get("3.KOMİSYON"))),
        ("2. Kademe", to_float(kom_row.get("2.Fiyat Üst Limiti")), to_float(kom_row.get("2.KOMİSYON"))),
        ("1. Kademe", to_float(kom_row.get("1.Fiyat Alt Limit")), to_float(kom_row.get("1.KOMİSYON"))),
    ]
    
    product_matches = []
    for k_name, p, r in brackets:
        if p is not None and p > 0 and r is not None and r > 0:
            net = round(p - (p * (r / 100.0)), 2)
            # Check condition: Price <= Dip Price AND Net > Dip Net (or Net >= Dip Net)
            if p <= dip_price + 0.01:
                net_diff = round(net - dip_net, 2)
                price_diff = round(p - dip_price, 2)
                if net > dip_net:
                    product_matches.append({
                        "Kademe": k_name,
                        "Kademe Fiyatı": p,
                        "Kademe Komisyon": r,
                        "Kademe Net": net,
                        "Dip Fiyat": dip_price,
                        "Dip Net": dip_net,
                        "Fiyat Farkı (TL)": price_diff,
                        "Net Kazanç Farkı (TL)": net_diff,
                    })
    
    if product_matches:
        matching_cases.append({
            "Barkod": b,
            "Ürün Adı": row.get("Ürün Adı", ""),
            "Güncel Fiyat": guncel_fiyat,
            "Güncel Komisyon": guncel_kom,
            "Güncel Net": guncel_net,
            "Dip Fiyat": dip_price,
            "Dip Net": dip_net,
            "Kademeler": product_matches
        })

print(f"Toplam Analiz Edilen Komisyon Tarifesi Ürün Sayısı: {total_analyzed}")
print(f"Fiyatı Dip Fiyata Eşit/Altında Olup Dip Netten DAHA İYİ Net Bırakan Ürün Sayısı: {len(matching_cases)}")

if matching_cases:
    print("\n--- ÖRNEK 10 EŞLEŞME ---")
    for item in matching_cases[:10]:
        print(f"\nBarkod: {item['Barkod']} | Güncel: {item['Güncel Fiyat']} TL (%{item['Güncel Komisyon']}) -> Net: {item['Güncel Net']} TL")
        print(f"Dip Fiyat: {item['Dip Fiyat']} TL -> Dip Net: {item['Dip Net']} TL")
        for k in item["Kademeler"]:
            print(f"  * {k['Kademe']}: Fiyat = {k['Kademe Fiyatı']} TL (%{k['Kademe Komisyon']}) -> Net = {k['Kademe Net']} TL [Net Farkı: +{k['Net Kazanç Farkı (TL)']} TL, Fiyat: {k['Fiyat Farkı (TL)']} TL]")
else:
    print("\nDip fiyatın altında olup dip netten daha yüksek getiri bırakan ürün bulunamadı.")
