import os
import re
import pandas as pd
import numpy as np

# Updated File paths (10 Ağustos catalog + Trendyol Damping_güncellendi + Dip Fiyatlar + İndirim Listesi)
f_curr = r"Girdiler/10agustos/Ürün Listesi.xlsx"
f_ind = r"Girdiler/İndirim Uygulanabilecek Ürünler.xlsx"
f_damp = r"Girdiler/muhasebe/7 ağustos/Trendyol Damping_güncellendi.xlsx"
f_dip = r"Girdiler/muhasebe/7 ağustos/DİP FİYATLAR.xlsx"

out_dir = r"Çıktılar"
os.makedirs(out_dir, exist_ok=True)
out_excel_path = os.path.join(out_dir, "Tek_Seferlik_Birlestirilmis_Indirim_Listesi.xlsx")

# 1. Read files
df_curr = pd.read_excel(f_curr)
df_ind = pd.read_excel(f_ind)
df_damp = pd.read_excel(f_damp)
df_dip = pd.read_excel(f_dip)

# 2. Build current price lookup
curr_col = [c for c in df_curr.columns if 'Satılacak' in c or 'Satilacak' in c][0]
curr_map = {}
curr_name_map = {}
for _, r in df_curr.iterrows():
    bc = str(r.get('Barkod', '')).strip()
    price = r.get(curr_col)
    name = r.get('Ürün Adı', '')
    if bc and pd.notna(price):
        try:
            curr_map[bc] = float(price)
            curr_name_map[bc] = str(name)
        except Exception:
            pass

# Collect price proposals per barcode
# bc -> list of dicts: {'source': str, 'yeni': float, 'eski': float}
bc_entries = {}

# Process Trendyol Damping Güncellendi
for _, r in df_damp.iterrows():
    bc = str(r.get('BARKOD', '')).strip()
    yeni = r.get('YENİ Fiyat')
    if bc and pd.notna(yeni):
        try:
            y_val = round(float(yeni), 2)
            eski_val = curr_map.get(bc, np.nan)
            if bc not in bc_entries: bc_entries[bc] = []
            bc_entries[bc].append({'source': 'Trendyol Damping', 'yeni': y_val, 'eski': eski_val})
        except Exception:
            pass

# Process DİP FİYATLAR
for _, r in df_dip.iterrows():
    bc = str(r.get('BARKOD', '')).strip()
    dip = r.get('DİP Fiyatlar')
    if bc and pd.notna(dip):
        try:
            y_val = round(float(dip), 2)
            eski_val = curr_map.get(bc, np.nan)
            if bc not in bc_entries: bc_entries[bc] = []
            bc_entries[bc].append({'source': 'DİP FİYATLAR', 'yeni': y_val, 'eski': eski_val})
        except Exception:
            pass

# Process İndirim Uygulanabilecek Ürünler
for _, r in df_ind.iterrows():
    bc = str(r.get('BARKOD', '')).strip()
    eski = r.get('Eski Fiyat')
    yeni = r.get('YENİ Fiyat')
    if bc and pd.notna(yeni):
        try:
            y_val = round(float(yeni), 2)
            e_val = float(eski) if (pd.notna(eski) and float(eski) > 0) else curr_map.get(bc, np.nan)
            if bc not in bc_entries: bc_entries[bc] = []
            bc_entries[bc].append({'source': 'İndirim Listesi', 'yeni': y_val, 'eski': e_val})
        except Exception:
            pass

# Process conflict resolutions and build final rows
# Rule: In case of conflict, choose LOWER YENİ Fiyat (min)
final_rows = []
conflict_reports = []

same_price_conflicts = 0
different_price_conflicts = 0

indirim_won_conflicts = 0
dip_won_conflicts = 0
damping_won_conflicts = 0

for bc, entries in bc_entries.items():
    eski_fiyat = curr_map.get(bc)
    if eski_fiyat is None:
        eski_candidates = [e['eski'] for e in entries if e['eski'] is not None and not pd.isna(e['eski'])]
        if eski_candidates:
            eski_fiyat = eski_candidates[0]

    # Conflict resolution: Take LOWEST YENİ Fiyat
    lowest_entry = min(entries, key=lambda x: x['yeni'])
    min_yeni_fiyat = lowest_entry['yeni']

    sources_str = " + ".join(dict.fromkeys([e['source'] for e in entries]))
    has_conflict = len(entries) > 1

    if has_conflict:
        prices_list = [e['yeni'] for e in entries]
        unique_prices = set(prices_list)
        winning_sources = [e['source'] for e in entries if e['yeni'] == min_yeni_fiyat]
        winner_source_str = " & ".join(dict.fromkeys(winning_sources))

        if len(unique_prices) == 1:
            same_price_conflicts += 1
            price_status = 'Birebir Aynı Fiyat'
        else:
            different_price_conflicts += 1
            price_status = 'Farklı Fiyatlar Var'
            
            if 'Trendyol Damping' in winning_sources:
                damping_won_conflicts += 1
            elif 'DİP FİYATLAR' in winning_sources:
                dip_won_conflicts += 1
            elif 'İndirim Listesi' in winning_sources:
                indirim_won_conflicts += 1

        prices_detail = " | ".join([f"{e['source']}: {e['yeni']:.2f} TL" for e in entries])
        conflict_reports.append({
            'BARKOD': bc,
            'Eski Fiyat (TL)': eski_fiyat,
            'Farklı Fiyatlar': prices_detail,
            'Seçilen Fiyat (Düşük Olan TL)': min_yeni_fiyat,
            'Kazanan Kaynak': winner_source_str,
            'Fiyat Eşitlik Durumu': price_status
        })

    # Calculations
    indirim_tutar = round(eski_fiyat - min_yeni_fiyat, 2) if (eski_fiyat is not None and not pd.isna(eski_fiyat)) else np.nan
    indirim_yuzde = round(((eski_fiyat - min_yeni_fiyat) / eski_fiyat) * 100, 2) if (eski_fiyat is not None and not pd.isna(eski_fiyat) and eski_fiyat > 0) else np.nan

    final_rows.append({
        'BARKOD': bc,
        'Ürün Adı': curr_name_map.get(bc, ''),
        'Eski Fiyat (TL)': eski_fiyat,
        'YENİ Fiyat (TL)': min_yeni_fiyat,
        'İndirim Tutarı (TL)': indirim_tutar,
        'İndirim Oranı (%)': indirim_yuzde,
        'Çakışma Durumu': 'Çakışma Var' if has_conflict else 'Tek Kaynak',
        'Kaynak Dosya(lar)': sources_str
    })

df_final = pd.DataFrame(final_rows)

# Export Excel with formatting
with pd.ExcelWriter(out_excel_path, engine='openpyxl') as writer:
    df_final.to_excel(writer, index=False, sheet_name='Birleştirilmiş İndirim Listesi')
    if conflict_reports:
        df_conflicts = pd.DataFrame(conflict_reports)
        df_conflicts.to_excel(writer, index=False, sheet_name='Çakışma Raporu')

print("SUCCESSFULLY GENERATED MERGED EXCEL AT:", out_excel_path)
print("TOTAL FINAL UNIQUE PRODUCTS:", len(df_final))
print("TOTAL CONFLICTS PROCESSED:", len(conflict_reports))
print(f"  - Birebir Aynı Fiyata Sahip Çakışmalar: {same_price_conflicts}")
print(f"  - Farklı Fiyatlara Sahip Çakışmalar: {different_price_conflicts}")
print(f"     * Trendyol Damping Düşük Olan (Kazanan): {damping_won_conflicts}")
print(f"     * DİP FİYATLAR Düşük Olan (Kazanan): {dip_won_conflicts}")
print(f"     * İndirim Listesi Düşük Olan (Kazanan): {indirim_won_conflicts}")
