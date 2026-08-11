import os
import pandas as pd

target_dir = r"C:\Users\Tasarımcı\Desktop\trendyol\Çıktılar\2026-08-06_17-19-06"
hesap_file = r"C:\Users\Tasarımcı\Desktop\trendyol\Çıktılar\Kampanya_Hesaplama_Sonuclari.xlsx"
avan_file = os.path.join(target_dir, "Avantajlı Ürün.xlsx")

df_hesap = pd.read_excel(hesap_file)
df_avan = pd.read_excel(avan_file)

print(f"Total rows in Avantajlı Ürün.xlsx: {len(df_avan)}")

# Merge with calculation results
df_merged = pd.merge(df_avan, df_hesap, left_on="BARKOD", right_on="Barkod", how="left")

print("\nSample violating rows in Avantajlı Ürün.xlsx:")
viol_rows = []
for idx, r in df_merged.iterrows():
    bc = r["BARKOD"]
    p_avan = float(r["YENİ TSF (FİYAT GÜNCELLE)"]) if pd.notna(r.get("YENİ TSF (FİYAT GÜNCELLE)")) else None
    dip = float(r["Düşülebilecek Dip Fiyat (TL)"]) if pd.notna(r.get("Düşülebilecek Dip Fiyat (TL)")) else None
    has_dip = r.get("İndirim Uygulanabilir")
    el_camp = r.get("eligible_campaigns")
    
    if p_avan is not None and dip is not None and p_avan < (dip - 0.01):
        viol_rows.append((bc, p_avan, dip, has_dip, el_camp))

print(f"Total Violating Rows in Avantajlı Ürün.xlsx: {len(viol_rows)}")
for v in viol_rows[:10]:
    print(f"Barcode: {v[0]} | TSF: {v[1]} TL | Dip: {v[2]} TL | İndirimUygul: {v[3]} | Eligible: {v[4]}")
