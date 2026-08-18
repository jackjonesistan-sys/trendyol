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

# Filter Komisyon Tarifesi eligible
eligible_list = [r for r in results if "Komisyon Tarifesi" in r.get("eligible_campaigns", [])]

# Calculate averages
avg_curr_price = sum(r["Güncel Ürün Fiyatı (TL)"] for r in eligible_list) / len(eligible_list)
avg_kom_price = sum(r["Komisyon Tarifesi Fiyatı (TL)"] for r in eligible_list) / len(eligible_list)
avg_curr_net = sum(r["Güncel Ürün Kalan Net (TL)"] for r in eligible_list) / len(eligible_list)
avg_kom_net = sum(r["Komisyon Tarifesi Net (TL)"] for r in eligible_list) / len(eligible_list)
avg_curr_comm = sum(r["Güncel Ürün Komisyon (%)"] for r in eligible_list) / len(eligible_list)
avg_kom_comm = sum(r["Komisyon Tarifesi Komisyon (%)"] for r in eligible_list) / len(eligible_list)

print(f"Ortalama Güncel Fiyat: {avg_curr_price:.2f} TL -> Ortalama Komisyon Tarifesi Fiyatı: {avg_kom_price:.2f} TL (İndirim: %{(1 - avg_kom_price/avg_curr_price)*100:.1f})")
print(f"Ortalama Güncel Net: {avg_curr_net:.2f} TL -> Ortalama Komisyon Tarifesi Net: {avg_kom_net:.2f} TL")
print(f"Ortalama Güncel Komisyon: %{avg_curr_comm:.2f} -> Ortalama Yeni Komisyon: %{avg_kom_comm:.2f}")

# Top 5 sample rows for display
samples = []
for r in eligible_list[:6]:
    samples.append({
        "Barkod": r["Barkod"],
        "Güncel Fiyat": f"{r['Güncel Ürün Fiyatı (TL)']:.2f} TL",
        "Güncel Net": f"{r['Güncel Ürün Kalan Net (TL)']:.2f} TL",
        "Dip Fiyat": f"{r.get('Düşülebilecek Dip Fiyat (TL)', 0):.2f} TL",
        "Tarife Fiyatı": f"{r['Komisyon Tarifesi Fiyatı (TL)']:.2f} TL",
        "Tarife Komisyon": f"%{r['Komisyon Tarifesi Komisyon (%)']:.1f}",
        "Tarife Net": f"{r['Komisyon Tarifesi Net (TL)']:.2f} TL",
        "Tarife Seçimi": r.get("Komisyon Tarifesi Seçimi", "7 Günlük Fiyat"),
    })

print("\nÖrnek Tablo:")
df_samples = pd.DataFrame(samples)
print(df_samples.to_markdown(index=False))
