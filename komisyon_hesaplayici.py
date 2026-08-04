import pandas as pd
import numpy as np

# ==============================================================================
# GEREKLİ DOSYALAR VE AÇIKLAMALARI
# Aşağıdaki değişkenlere kendi dosya isimlerinizi veya yollarınızı girebilirsiniz.
# ==============================================================================
# f_trendyol: "Trendyol Fiyat Güncelleme" şablonu. 
#             - 'Durum' sütunu "İndirim" olan ürünlerin BARKOD'larını temel alır.
# f_avantajli: "Avantajlı Ürün" şablonu.
#             - BARKOD, MÜŞTERİNİN GÖRDÜĞÜ FİYAT, YENİ TSF (FİYAT GÜNCELLE) sütunları gereklidir.
# f_komisyon: "Ürün Komisyon Tarifeleri" şablonu.
#             - Fiyat alt/üst limitleri ve 1., 2., 3., 4. komisyon oranlarını içerir.
# f_guncel: "Güncel Ürünleriniz" şablonu.
#             - Ürünün mevcuttaki Trendyol'da Satılacak Fiyat (KDV Dahil) değerini içerir.
# f_flas: "Flaş Ürünler" şablonu.
#             - '24 Saat Fiyat' sütununu içerir.
# out_file: Hesaplamalar sonucunda oluşturulacak ve kaydedilecek yeni Excel dosyasının adı.
# ==============================================================================

import os
import glob

def calculate_all(karsilamali_config=None):
    base_dir = r'c:\Users\Tasarımcı\Desktop\trendyol'
    input_dir = os.path.join(base_dir, 'Girdiler')
    output_dir = os.path.join(base_dir, 'Çıktılar')
    f_trendyol = os.path.join(input_dir, 'İndirim Uygulanabilecek Ürünler.xlsx')

    # Karşılamalı kampanya konfigürasyonunu kontrol et
    config_json_path = os.path.join(input_dir, 'karsilamali_config.json')
    if not karsilamali_config and os.path.exists(config_json_path):
        try:
            import json
            with open(config_json_path, 'r', encoding='utf-8') as cf:
                karsilamali_config = json.load(cf)
        except:
            pass

    f_karsilamali = None
    satici_indirim = 0.0
    if karsilamali_config:
        f_karsilamali = karsilamali_config.get('file_path')
        if f_karsilamali and not os.path.isabs(f_karsilamali):
            f_karsilamali = os.path.join(base_dir, f_karsilamali)
        toplam_indirim = float(karsilamali_config.get('toplam_indirim', 0))
        trendyol_oran = float(karsilamali_config.get('trendyol_oran', 0))
        satici_indirim = toplam_indirim * (1.0 - (trendyol_oran / 100.0))

    excel_files = glob.glob(os.path.join(input_dir, '*.xlsx'))
    f_avantajli = None
    f_komisyon = None
    f_guncel = None
    f_flas = None
    f_plus = None
    f_plus_ek = None

    for f in excel_files:
        if "Hesaplanmis_Komisyon" in f or "Uygulanmis" in f or "Trendyol Fiyat" in f:
            continue
        try:
            cols = pd.read_excel(f, nrows=0).columns.tolist()
            if 'TARİFE GRUBU' in cols and 'KOMİSYONA ESAS FİYAT' in cols:
                f_komisyon = f
            elif '24 Saat Fiyat' in cols and 'Kampanyalı Ürün' in cols:
                f_flas = f
            elif '1 YILDIZ ÜST FİYAT' in cols and 'YENİ TSF (FİYAT GÜNCELLE)' in cols:
                f_avantajli = f
            elif 'Plus Fiyat Üst Limiti' in cols and 'Plus Komisyon Teklifi' in cols:
                f_plus = f
            elif 'Maksimum Girebileceğin Fiyat' in cols and 'Kampanyalı Satış Fiyatı' in cols:
                if not f_karsilamali or os.path.abspath(f) != os.path.abspath(f_karsilamali):
                    f_plus_ek = f
            elif 'BuyBox Fiyatı' in cols and 'Trendyol\'da Satılacak Fiyat (KDV Dahil)' in cols:
                if 'GüncelÜrünleriniz' in os.path.basename(f):
                    f_guncel = f
        except:
            pass

    if not all([f_avantajli, f_komisyon, f_guncel, f_flas, f_plus, f_plus_ek]):
        return {"success": False, "message": "HATA: Gerekli bazı şablon dosyaları (Avantajlı, Komisyon, Güncel, Flaş, Plus veya Plus Ek İndirim) Girdiler klasöründe bulunamadı!"}

    out_file = os.path.join(output_dir, 'Hesaplanmis_Komisyon_Sonuclari.xlsx')

    try:
        df_tr = pd.read_excel(f_trendyol)
        df_av = pd.read_excel(f_avantajli)
        df_kom = pd.read_excel(f_komisyon)
        df_gun = pd.read_excel(f_guncel)
        df_flas = pd.read_excel(f_flas)
        df_plus = pd.read_excel(f_plus)
        df_plus_ek = pd.read_excel(f_plus_ek)
    except Exception as e:
        return {"success": False, "message": f"Excel dosyaları okunurken hata oluştu: {str(e)}"}

    df_kars = pd.DataFrame()
    if f_karsilamali and os.path.exists(f_karsilamali):
        try:
            df_kars = pd.read_excel(f_karsilamali)
            df_kars['BARKOD_CLN'] = df_kars['Barkod'].astype(str).str.strip()
        except Exception as e:
            print("Karşılamalı dosya okuma hatası:", e)

    df_tr['BARKOD_CLN'] = df_tr['BARKOD'].astype(str).str.strip()
    indirimli = df_tr[df_tr['Durum'].astype(str).str.contains('ndirim', case=False, na=False)]
    tr_barcodes = indirimli['BARKOD_CLN'].unique()
    
    df_gun['BARKOD_CLN'] = df_gun['Barkod'].astype(str).str.strip()
    gun_barcodes = df_gun['BARKOD_CLN'].unique()
    
    barcodes = list(set(tr_barcodes).union(set(gun_barcodes)))

    df_av['BARKOD_CLN'] = df_av['BARKOD'].astype(str).str.strip()
    df_kom['BARKOD_CLN'] = df_kom['BARKOD'].astype(str).str.strip()
    df_flas['BARKOD_CLN'] = df_flas['Barkod'].astype(str).str.strip()
    df_plus['BARKOD_CLN'] = df_plus['Barkod'].astype(str).str.strip()
    df_plus_ek['BARKOD_CLN'] = df_plus_ek['Barkod'].astype(str).str.strip()

    # Komisyon oranını fiyat dilimlerine bakarak getiren yardımcı fonksiyon
    def get_commission_rate(price, row):
        try:
            p = float(price)
            if pd.isna(p):
                return 0.0
        except:
            return 0.0

        r1_alt = float(row['1.Fiyat Alt Limit'])
        r2_ust = float(row['2.Fiyat Üst Limiti'])
        r2_alt = float(row['2.Fiyat Alt Limit'])
        r3_ust = float(row['3.Fiyat Üst Limiti'])
        r3_alt = float(row['3.Fiyat Alt Limit'])
        r4_ust = float(row['4.Fiyat Üst Limiti'])

        eps = 0.001
        
        if p >= r1_alt - eps:
            return float(row['1.KOMİSYON'])
        elif p >= r2_alt - eps and p <= r2_ust + eps:
            return float(row['2.KOMİSYON'])
        elif p >= r3_alt - eps and p <= r3_ust + eps:
            return float(row['3.KOMİSYON'])
        elif p <= r4_ust + eps:
            return float(row['4.KOMİSYON'])
        else:
            if p > r2_ust: return float(row['1.KOMİSYON'])
            elif p > r3_ust: return float(row['2.KOMİSYON'])
            elif p > r4_ust: return float(row['3.KOMİSYON'])
            return float(row['4.KOMİSYON'])


    def to_dict_safe(df, key_col):
        if df.empty: return {}
        # Keep first occurrence of each barcode to simulate iloc[0]
        return df.drop_duplicates(subset=[key_col]).set_index(key_col).to_dict('index')
        
    dict_av = to_dict_safe(df_av, 'BARKOD_CLN')
    dict_flas = to_dict_safe(df_flas, 'BARKOD_CLN')
    dict_kom = to_dict_safe(df_kom, 'BARKOD_CLN')
    dict_gun = to_dict_safe(df_gun, 'BARKOD_CLN')
    dict_ind = to_dict_safe(indirimli, 'BARKOD_CLN')
    dict_plus = to_dict_safe(df_plus, 'BARKOD_CLN')
    dict_plus_ek = to_dict_safe(df_plus_ek, 'BARKOD_CLN')
    dict_kars = to_dict_safe(df_kars, 'BARKOD_CLN')

    results = []
    
    for b in barcodes:
        av_row = dict_av.get(b)
        fl_row = dict_flas.get(b)
        kom_row = dict_kom.get(b)
        gun_row = dict_gun.get(b)
        tr_row = dict_ind.get(b)
        plus_row = dict_plus.get(b)
        plus_ek_row = dict_plus_ek.get(b)
        kars_row = dict_kars.get(b)
        
        match_av = av_row is not None
        match_fl = fl_row is not None
        match_plus = plus_row is not None
        match_plus_ek = plus_ek_row is not None
        match_kars = kars_row is not None
        
        her_ikisi = match_av and match_fl
        
        av_str = 'Eşleşti' if match_av else 'Eşleşme Yok'
        fl_str = 'Eşleşti' if match_fl else 'Eşleşme Yok'
        plus_str = 'Eşleşti' if match_plus else 'Eşleşme Yok'
        plus_ek_str = 'Eşleşti' if match_plus_ek else 'Eşleşme Yok'
        kars_str = 'Eşleşti' if match_kars else 'Eşleşme Yok'
        both_str = 'Evet' if her_ikisi else 'Hayır'
        
        guncel_fiyat = None
        rate_1 = None
        net_1 = None
        
        yeni_tsf = None
        rate_2 = None
        net_2 = None
        
        f_24_fiyat = None
        rate_3 = None
        net_3 = None
        
        plus_fiyat = None
        rate_4 = None
        net_4 = None

        plus_ek_fiyat = None
        rate_5 = None
        net_5 = None
        
        plus_ek_fiyat_5 = None
        rate_5_5 = None
        net_5_5 = None
        
        plus_ek_fiyat_10 = None
        rate_5_10 = None
        net_5_10 = None
        
        plus_ek_fiyat_20 = None
        rate_5_20 = None
        net_5_20 = None

        kars_fiyat = None
        rate_kars = None
        net_kars = None
        
        daha_karli_kampanya = None
        guncel_kom_orani = None
        mevcut_indirim_orani = None
        
        is_indirim = tr_row is not None
        
        if is_indirim:
            try:
                yeni_fiyat = float(tr_row['YENİ Fiyat'])
                eski_fiyat = float(tr_row['Eski Fiyat'])
                guncel_fiyat = eski_fiyat if eski_fiyat > 0 else yeni_fiyat
                if eski_fiyat > 0:
                    mevcut_indirim_orani = round(((eski_fiyat - yeni_fiyat) / eski_fiyat) * 100.0, 2)
            except:
                guncel_fiyat = None
                eski_fiyat = None
            
            rate_1 = None
            if kom_row is not None and guncel_fiyat is not None:
                rate_1 = get_commission_rate(guncel_fiyat, kom_row)
            
            if (pd.isna(rate_1) or rate_1 is None) and gun_row is not None:
                try: rate_1 = float(gun_row['Komisyon Oranı'])
                except: pass
                
            if rate_1 is not None and not pd.isna(rate_1) and guncel_fiyat is not None:
                kom_tutar = guncel_fiyat * (rate_1 / 100.0)
                net_1 = round(guncel_fiyat - kom_tutar, 2)
        else:
            if gun_row is not None:
                try:
                    guncel_fiyat = float(gun_row['Trendyol\'da Satılacak Fiyat (KDV Dahil)'])
                except:
                    guncel_fiyat = None
                
                try:
                    piyasa_fiyat = float(gun_row['Piyasa Satış Fiyatı (KDV Dahil)'])
                    if piyasa_fiyat > 0 and guncel_fiyat is not None and piyasa_fiyat > guncel_fiyat:
                        mevcut_indirim_orani = round(((piyasa_fiyat - guncel_fiyat) / piyasa_fiyat) * 100.0, 2)
                except:
                    pass
                
                rate_1 = None
                if guncel_fiyat is not None and kom_row is not None:
                    rate_1 = get_commission_rate(guncel_fiyat, kom_row)
                
                if (pd.isna(rate_1) or rate_1 is None):
                    try: rate_1 = float(gun_row['Komisyon Oranı'])
                    except: pass
                    
                if rate_1 is not None and not pd.isna(rate_1) and guncel_fiyat is not None:
                    net_1 = round(guncel_fiyat - (guncel_fiyat * (rate_1 / 100.0)), 2)

        if match_av:
            try: yeni_tsf = float(av_row['YENİ TSF (FİYAT GÜNCELLE)'])
            except: yeni_tsf = None
                
        if match_fl:
            try: f_24_fiyat = float(fl_row['24 Saat Fiyat'])
            except: f_24_fiyat = None
            
        if match_plus:
            try: plus_fiyat = float(plus_row['Plus Fiyat Üst Limiti'])
            except: plus_fiyat = None

        if match_plus_ek:
            try: plus_ek_fiyat = float(plus_ek_row['Maksimum Girebileceğin Fiyat'])
            except: plus_ek_fiyat = None

        if yeni_tsf is not None:
            rate_2 = None
            if kom_row is not None:
                rate_2 = get_commission_rate(yeni_tsf, kom_row)
            if (pd.isna(rate_2) or rate_2 is None) and gun_row is not None:
                try: rate_2 = float(gun_row['Komisyon Oranı'])
                except: pass
            if rate_2 is not None and not pd.isna(rate_2):
                net_2 = round(yeni_tsf - (yeni_tsf * (rate_2 / 100.0)), 2)
                
        if f_24_fiyat is not None:
            rate_3 = None
            if kom_row is not None:
                rate_3 = get_commission_rate(f_24_fiyat, kom_row)
            if (pd.isna(rate_3) or rate_3 is None) and gun_row is not None:
                try: rate_3 = float(gun_row['Komisyon Oranı'])
                except: pass
            if rate_3 is not None and not pd.isna(rate_3):
                net_3 = round(f_24_fiyat - (f_24_fiyat * (rate_3 / 100.0)), 2)
                
        if plus_fiyat is not None:
            try:
                rate_str = str(plus_row['Plus Komisyon Teklifi']).replace(',', '.')
                rate_4 = float(rate_str)
            except:
                rate_4 = None
            
            if (pd.isna(rate_4) or rate_4 is None) and kom_row is not None:
                rate_4 = get_commission_rate(plus_fiyat, kom_row)
            if (pd.isna(rate_4) or rate_4 is None) and gun_row is not None:
                try: rate_4 = float(gun_row['Komisyon Oranı'])
                except: pass
                
            if rate_4 is not None and not pd.isna(rate_4):
                net_4 = round(plus_fiyat - (plus_fiyat * (rate_4 / 100.0)), 2)

        if plus_ek_fiyat is not None:
            rate_5 = None
            if kom_row is not None:
                rate_5 = get_commission_rate(plus_ek_fiyat, kom_row)
            if (pd.isna(rate_5) or rate_5 is None) and gun_row is not None:
                try: rate_5 = float(gun_row['Komisyon Oranı'])
                except: pass
                
            if rate_5 is not None and not pd.isna(rate_5):
                net_5 = round(plus_ek_fiyat - (plus_ek_fiyat * (rate_5 / 100.0)), 2)
            else:
                net_5 = None

            # Plus Ek İndirim %5
            plus_ek_fiyat_5 = plus_ek_fiyat
            rate_5_5 = rate_5
            if rate_5_5 is not None:
                indirimli_fiyat_5 = plus_ek_fiyat * 0.95
                net_5_5 = round(indirimli_fiyat_5 - (plus_ek_fiyat * (rate_5_5 / 100.0)), 2)
            else:
                net_5_5 = None

            # Plus Ek İndirim %10
            plus_ek_fiyat_10 = plus_ek_fiyat
            rate_5_10 = rate_5
            if rate_5_10 is not None:
                indirimli_fiyat_10 = plus_ek_fiyat * 0.90
                net_5_10 = round(indirimli_fiyat_10 - (plus_ek_fiyat * (rate_5_10 / 100.0)), 2)
            else:
                net_5_10 = None

            # Plus Ek İndirim %20
            plus_ek_fiyat_20 = plus_ek_fiyat
            rate_5_20 = rate_5
            if rate_5_20 is not None:
                indirimli_fiyat_20 = plus_ek_fiyat * 0.80
                net_5_20 = round(indirimli_fiyat_20 - (plus_ek_fiyat * (rate_5_20 / 100.0)), 2)
            else:
                net_5_20 = None

        if match_kars:
            try: kars_fiyat = float(kars_row['Maksimum Girebileceğin Fiyat'])
            except: kars_fiyat = None
            
        if kars_fiyat is not None:
            rate_kars = None
            if kom_row is not None:
                rate_kars = get_commission_rate(kars_fiyat, kom_row)
            if (pd.isna(rate_kars) or rate_kars is None) and gun_row is not None:
                try: rate_kars = float(gun_row['Komisyon Oranı'])
                except: pass
            if rate_kars is not None and not pd.isna(rate_kars):
                net_kars = round(kars_fiyat - (kars_fiyat * (rate_kars / 100.0)) - satici_indirim, 2)

        karlilik_farki_yuzde = ''
        n1 = net_1 if net_1 is not None else 0
        n2 = net_2 if net_2 is not None else -9999
        n3 = net_3 if net_3 is not None else -9999
        n4 = net_4 if net_4 is not None else -9999
        n5 = net_5 if net_5 is not None else -9999
        
        av_gecerli = False
        fl_gecerli = False
        plus_gecerli = False
        plus_ek_gecerli = False
        
        if n2 > 0:
            if match_av and yeni_tsf is not None:
                av_gecerli = True
            else:
                av_gecerli = (n2 >= n1)
                
        if n3 > 0:
            if match_av and yeni_tsf is not None:
                if f_24_fiyat is not None and f_24_fiyat >= yeni_tsf:
                    fl_gecerli = True
                else:
                    fl_gecerli = (n3 >= n1)
            else:
                fl_gecerli = (n3 >= n1)
                
        if n4 > 0:
            plus_gecerli = (n4 >= n1)
            
        valid_options = []
        if av_gecerli: valid_options.append(('Avantajlı Ürün', n2))
        if fl_gecerli: valid_options.append(('Flaş Ürün', n3))
        if plus_gecerli: valid_options.append(('Plus Ürün', n4))
        
        if not valid_options:
            daha_karli_kampanya = 'Hiçbiri'
        else:
            valid_options.sort(key=lambda x: x[1], reverse=True)
            daha_karli_kampanya = valid_options[0][0]

        if daha_karli_kampanya != 'Hiçbiri':
            hedef_n = valid_options[0][1]

        guncel_fiyat_display = guncel_fiyat
        if is_indirim and eski_fiyat is not None:
            guncel_fiyat_display = eski_fiyat

        results.append({
            'Barkod': b,
            'Plus Ek İndirim Eşleşme Durumu': plus_ek_str,
            'Plus Eşleşme Durumu': plus_str,
            'Avantajlı Ürün Eşleşme Durumu': av_str,
            'Flaş Ürün Eşleşme Durumu': fl_str,
            'Hem Avantajlı Hem Flaş': both_str,
            'İndirim Uygulanabilir': 'Evet' if is_indirim else 'Hayır',
            'Mevcut İndirim Oranı (%)': mevcut_indirim_orani,
            'Hangisi Daha Karlı?': daha_karli_kampanya,
            'Karlılık Farkı (%)': karlilik_farki_yuzde,
            'Güncel Ürün Fiyatı (TL)': guncel_fiyat_display,
            'Güncel Ürün Komisyon (%)': rate_1,
            'Güncel Ürün Kalan Net (TL)': net_1,
            'Avantajlı Ürün Fiyatı (YENİ TSF) (TL)': yeni_tsf,
            'Avantajlı Ürün Komisyon (%)': rate_2,
            'Avantajlı Ürün Kalan Net (TL)': net_2,
            'Flaş Ürün 24 Saat Fiyatı (TL)': f_24_fiyat,
            'Flaş Ürün Komisyon (%)': rate_3,
            'Flaş Ürün Kalan Net (TL)': net_3,
            'Plus Fiyatı (TL)': plus_fiyat,
            'Plus Komisyon (%)': rate_4,
            'Plus Net (TL)': net_4,
            'Plus Ek Fiyatı (TL)': plus_ek_fiyat,
            'Plus Ek Komisyon (%)': rate_5,
            'Plus Ek Net (TL)': net_5,
            'Plus Ek Fiyatı %5 (TL)': plus_ek_fiyat_5,
            'Plus Ek Net %5 (TL)': net_5_5,
            'Plus Ek Fiyatı %10 (TL)': plus_ek_fiyat_10,
            'Plus Ek Net %10 (TL)': net_5_10,
            'Plus Ek Fiyatı %20 (TL)': plus_ek_fiyat_20,
            'Plus Ek Net %20 (TL)': net_5_20,
            'Karşılamalı Kampanya Eşleşme Durumu': kars_str,
            'Karşılamalı Kampanya Fiyatı (TL)': kars_fiyat,
            'Karşılamalı Kampanya Komisyon (%)': rate_kars,
            'Karşılamalı Kampanya Kalan Net (TL)': net_kars
        })

    res_df = pd.DataFrame(results)
    res_df.to_excel(out_file, index=False)
    return {"success": True, "message": f'Başarıyla hesaplandı ve kaydedildi: {os.path.basename(out_file)}'}

if __name__ == "__main__":
    print(calculate_all())
