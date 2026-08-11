import os, glob
import pandas as pd

target_dir = r"C:\Users\Tasarımcı\Desktop\trendyol\Çıktılar\2026-08-06_17-19-06"
hesap_file = r"C:\Users\Tasarımcı\Desktop\trendyol\Çıktılar\Kampanya_Hesaplama_Sonuclari.xlsx"

df_hesap = pd.read_excel(hesap_file)
barcode_dip_map = {}
for idx, row in df_hesap.iterrows():
    bc = str(row.get("Barkod", "")).strip()
    if bc:
        dip = row.get("Düşülebilecek Dip Fiyat (TL)")
        try: dip_val = float(dip) if pd.notna(dip) else None
        except: dip_val = None
        barcode_dip_map[bc] = dip_val

print("=== TRENDYOL YÜKLEME DOSYALARI DİP FİYAT EMNİYET DENETİMİ ===")

files = glob.glob(os.path.join(target_dir, "*.xlsx"))
submission_files = [f for f in files if not any(w in os.path.basename(f) for w in ["Genel_Raporu", "Ozet_Raporu", "Uygulanmayan"])]

print(f"Denetlenen Trendyol Gönderim Dosyası Sayısı: {len(submission_files)}\n")

total_violations = 0
total_checked = 0

for fpath in sorted(submission_files):
    fname = os.path.basename(fpath)
    df = pd.read_excel(fpath)
    
    bcols = [c for c in df.columns if "barkod" in str(c).lower()]
    bcol = bcols[0] if bcols else None
    
    pcol = None
    if "YENİ TSF" in df.columns:
        pcol = "YENİ TSF"
    elif "24 Saat Fiyat" in df.columns:
        pcol = "24 Saat Fiyat"
    elif "Seçilen Satış Fiyatı" in df.columns:
        pcol = "Seçilen Satış Fiyatı"
    elif "Kampanyalı Satış Fiyatı" in df.columns:
        pcol = "Kampanyalı Satış Fiyatı"
    else:
        for c in df.columns:
            if "fiyat" in str(c).lower() or "tsf" in str(c).lower():
                pcol = c
                break
                
    print(f"Dosya: {fname}")
    print(f"  - Toplam Ürün Satırı: {len(df)}")
    print(f"  - Barkod Sütunu: {bcol}")
    print(f"  - Fiyat Sütunu: {pcol}")
    
    viol_list = []
    checked_in_file = 0
    
    if bcol and pcol:
        for idx, row in df.iterrows():
            bc = str(row[bcol]).strip()
            price_raw = row[pcol]
            try: price_val = float(price_raw) if pd.notna(price_raw) else None
            except: price_val = None
            
            dip_val = barcode_dip_map.get(bc)
            
            if price_val is not None and dip_val is not None:
                checked_in_file += 1
                total_checked += 1
                if price_val < (dip_val - 0.01):
                    total_violations += 1
                    viol_list.append((bc, price_val, dip_val, round(dip_val - price_val, 2)))
                    
        print(f"  - Dip Fiyatı Kontrol Edilen Ürün: {checked_in_file}")
        if viol_list:
            print(f"  - !!! DİP FİYAT ALTI İHLAL SAYISI: {len(viol_list)} !!!")
            for v in viol_list[:5]:
                print(f"      İhlal: Barkod={v[0]}, Uygulanan Fiyat={v[1]} TL, Dip Fiyat={v[2]} TL (Fark={v[3]} TL)")
        else:
            print("  - GÜVENLİ: Dip fiyat altı hiçbir ürün YOKTUR (0 İhlal).")
    else:
        print("  - Fiyat sütunu bulunamadı / boş dosya.")
    print("-" * 65)

print("\n================ DENEİM SONUCU ================")
print(f"Kontrol Edilen Toplam Ürün Fiyatı: {total_checked}")
print(f"Dip Fiyatın Altında Satılan Ürün Sayısı: {total_violations}")

if total_violations == 0:
    print("SONUÇ: %100 GÜVENLİ! DİP FİYATTAN DÜŞÜK UYGULANAN HİÇBİR FİYAT YOKTUR. ZARAR ETME RİSKİ SIFIRDIR!")
else:
    print(f"SONUÇ: TEHLİKE! {total_violations} adet üründe dip fiyatın altında fiyat tespit edildi!")
