import pandas as pd
import numpy as np
import os
import glob

def clean_price(val):
    if pd.isna(val) or str(val).strip() == '' or str(val).strip() == 'None':
        return None
    try:
        s = str(val).strip()
        if ',' in s and '.' in s:
            s = s.replace('.', '').replace(',', '.')
        elif ',' in s:
            s = s.replace(',', '.')
        return float(s)
    except:
        return None

def generate_fiyat_farki_raporu(output_dir='Çıktılar'):
    print("Veriler Yükleniyor...")
    input_dir = 'Girdiler'
    
    indirim_files = [f for f in os.listdir(input_dir) if 'ndirim' in f and f.endswith('.xlsx') and not f.startswith('~$')]
    if not indirim_files:
        print("İndirim Uygulanabilecek Ürünler dosyası Girdiler klasöründe bulunamadı.")
        return
    df_indirim = pd.read_excel(os.path.join(input_dir, indirim_files[0]), dtype=str)
    
    uygulanmayan_path = os.path.join(output_dir, 'Uygulanmayan_Urunler_Raporu.xlsx')
    if not os.path.exists(uygulanmayan_path):
        print("Uygulanmayan_Urunler_Raporu.xlsx Çıktılar klasöründe bulunamadı.")
        return
    df_uygulanmayan = pd.read_excel(uygulanmayan_path, dtype=str)
    
    ind_barkod_col = [c for c in df_indirim.columns if 'barkod' in c.lower()][0]
    ind_eski_col = [c for c in df_indirim.columns if 'eski fiyat' in c.lower()][0]
    ind_yeni_col = [c for c in df_indirim.columns if 'yen' in c.lower() and 'fiyat' in c.lower()][0]
    
    uyg_barkod_col = [c for c in df_uygulanmayan.columns if 'barkod' in c.lower()][0]
    uyg_guncel_col = [c for c in df_uygulanmayan.columns if 'ncel' in c.lower() and 'fiyat' in c.lower()][0]
    uyg_avan_col = [c for c in df_uygulanmayan.columns if 'avantajl' in c.lower() and 'fiyat' in c.lower()][0]
    
    df_indirim['BARKOD_CLN'] = df_indirim[ind_barkod_col].astype(str).str.strip()
    df_uygulanmayan['BARKOD_CLN'] = df_uygulanmayan[uyg_barkod_col].astype(str).str.strip()
    
    # Build dictionary from İndirim file to always use it as the source of truth
    indirim_dict = {}
    for row in df_indirim.to_dict('records'):
        b = row['BARKOD_CLN']
        eski = clean_price(row.get(ind_eski_col))
        yeni = clean_price(row.get(ind_yeni_col))
        indirim_dict[b] = (eski, yeni)
        
    uygulanmayan_barkodlar = set(df_uygulanmayan['BARKOD_CLN'].unique())
    
    results = []
    
    for row in df_uygulanmayan.to_dict('records'):
        b = row['BARKOD_CLN']
        
        # Eğer İndirim dosyasında varsa, GÜVENİLİR BİLGİYİ oradan al.
        if b in indirim_dict:
            eski, yeni = indirim_dict[b]
        else:
            eski = clean_price(row.get(uyg_guncel_col))
            yeni = clean_price(row.get(uyg_avan_col))
            if yeni is None or pd.isna(yeni):
                yeni = eski
                
        results.append({
            'Barkod': b,
            'Eski Fiyat': eski,
            'Yeni Fiyat': yeni
        })
        
    for row in df_indirim.to_dict('records'):
        b = row['BARKOD_CLN']
        if b not in uygulanmayan_barkodlar:
            eski = clean_price(row.get(ind_eski_col))
            yeni = clean_price(row.get(ind_yeni_col))
            
            results.append({
                'Barkod': b,
                'Eski Fiyat': eski,
                'Yeni Fiyat': yeni
            })
            
    df_res = pd.DataFrame(results)
    
    def calc_drop(r):
        e = r['Eski Fiyat']
        y = r['Yeni Fiyat']
        if pd.isna(e) or pd.isna(y) or e == 0:
            return 0.0
        drop = ((e - y) / e) * 100
        return round(drop, 2)
        
    df_res['Fiyat Farkı Yüzdesel (%)'] = df_res.apply(calc_drop, axis=1)
    
    out_path = os.path.join(output_dir, 'Indirim_Uygulanmayan_Fiyat_Kiyas_Raporu.xlsx')
    df_res.to_excel(out_path, index=False)
    print(f"Rapor başarıyla oluşturuldu: {out_path}")

if __name__ == '__main__':
    generate_fiyat_farki_raporu()
