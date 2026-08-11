import pandas as pd
import os

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
    except (TypeError, ValueError):
        return None

def find_column(frame, *terms):
    return next(
        (
            column
            for column in frame.columns
            if all(term in str(column).casefold() for term in terms)
        ),
        None,
    )


def generate_fiyat_farki_raporu(output_dir='Çıktılar', discount_path=None):
    print("Veriler Yükleniyor...")
    canonical_discount = os.path.join('Girdiler', 'Yuklenen', 'discount.xlsx')
    discount_path = discount_path or (canonical_discount if os.path.exists(canonical_discount) else None)
    df_indirim = (
        pd.read_excel(discount_path, dtype=str)
        if discount_path and os.path.exists(discount_path)
        else pd.DataFrame(columns=['BARKOD', 'Eski Fiyat', 'YENİ Fiyat'])
    )
    
    uygulanmayan_path = os.path.join(output_dir, 'Uygulanmayan_Urunler_Raporu.xlsx')
    if not os.path.exists(uygulanmayan_path):
        print("Uygulanmayan_Urunler_Raporu.xlsx Çıktılar klasöründe bulunamadı.")
        return
    df_uygulanmayan = pd.read_excel(uygulanmayan_path, dtype=str)
    
    ind_barkod_col = find_column(df_indirim, 'barkod')
    ind_eski_col = find_column(df_indirim, 'eski', 'fiyat')
    ind_yeni_col = find_column(df_indirim, 'yen', 'fiyat')
    uyg_barkod_col = find_column(df_uygulanmayan, 'barkod')
    uyg_guncel_col = find_column(df_uygulanmayan, 'ncel', 'fiyat')
    uyg_avan_col = find_column(df_uygulanmayan, 'avantajl', 'fiyat')
    if not all((ind_barkod_col, ind_eski_col, ind_yeni_col, uyg_barkod_col, uyg_guncel_col)):
        print("Fiyat farkı raporu için gerekli sütunlar bulunamadı.")
        return
    
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
            yeni = clean_price(row.get(uyg_avan_col)) if uyg_avan_col else None
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
