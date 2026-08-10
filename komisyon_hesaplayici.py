import os
import math
import pandas as pd
from pathlib import Path

ROUNDING_EPSILON = 1e-9

CAMPAIGN_LABELS = {
    "Avantajlı": "Avantajlı Ürün",
    "Flaş": "Flaş Ürün",
    "Plus": "Plus Ürün",
}


def to_float(val):
    if val is None or pd.isna(val):
        return None
    try:
        number = float(val) if isinstance(val, (int, float)) else float(str(val).strip().replace(',', '.'))
        return number if math.isfinite(number) else None
    except ValueError:
        return None


def get_commission_rate(price, row):
    try:
        p = to_float(price)
        r1_alt = to_float(row['1.Fiyat Alt Limit'])
        r2_ust = to_float(row['2.Fiyat Üst Limiti'])
        r2_alt = to_float(row['2.Fiyat Alt Limit'])
        r3_ust = to_float(row['3.Fiyat Üst Limiti'])
        r3_alt = to_float(row['3.Fiyat Alt Limit'])
        r4_ust = to_float(row['4.Fiyat Üst Limiti'])
    except (KeyError, TypeError, ValueError):
        return None

    if p is None or r1_alt is None:
        return None

    eps = 0.001
    
    try:
        if p >= r1_alt - eps:
            return to_float(row['1.KOMİSYON'])
        if r2_alt is not None and r2_ust is not None and r2_alt - eps <= p <= r2_ust + eps:
            return to_float(row['2.KOMİSYON'])
        if r3_alt is not None and r3_ust is not None and r3_alt - eps <= p <= r3_ust + eps:
            return to_float(row['3.KOMİSYON'])
        if r4_ust is not None and p <= r4_ust + eps:
            return to_float(row['4.KOMİSYON'])
        if r2_ust is not None and p > r2_ust:
            return to_float(row['1.KOMİSYON'])
        if r3_ust is not None and p > r3_ust:
            return to_float(row['2.KOMİSYON'])
        if r4_ust is not None and p > r4_ust:
            return to_float(row['3.KOMİSYON'])
        return to_float(row['4.KOMİSYON'])
    except (KeyError, TypeError, ValueError):
        return None


def selectable_campaigns(current_net, candidates):
    return [
        candidate
        for candidate in candidates
        if not (
            len(candidate) > 4
            and candidate[4]
            and (current_net is None or candidate[1] <= current_net)
        )
    ]


ALLOWED_FOR_RECOMMENDATION = {'Avantajlı', 'Flaş', 'Plus'}
MAIN_CAMPAIGN_KEYS = {'Avantajlı', 'Flaş', 'Plus'}


def choose_campaigns_smart(current_net, candidates):
    """
    candidates: (campaign_key, net_price, eff_price, rate, used_current_price=False, is_muhasebe=False)
    
    Önerilen Kampanya (Ana):
      Dip fiyatı geçebilen Ana kampanyalar ('Avantajlı', 'Flaş', 'Plus') arasından 
      fiyatı dip fiyata en yakın olanı seçilir.
      
    Önerilen Ekstra Kampanya:
      Dip fiyatı geçebilen Ekstra kampanyalar (Karşılamalı veya Plus Ek İndirim) arasından 
      fiyatı dip fiyata en yakın olanı seçilir.
    """
    selectable = selectable_campaigns(current_net, candidates)
    if not selectable:
        return 'Hiçbiri', 'Hiçbiri', '', 'Hiçbiri', 'Hiçbiri'

    main_selectable = [c for c in selectable if c[0] in MAIN_CAMPAIGN_KEYS]
    if main_selectable:
        best_main = min(main_selectable, key=lambda x: (x[2], -x[1]))[0]
    else:
        best_main = 'Hiçbiri'

    extra_selectable = [c for c in selectable if c[0] not in MAIN_CAMPAIGN_KEYS]
    if extra_selectable:
        best_extra = min(extra_selectable, key=lambda x: (x[2], -x[1]))[0]
    else:
        best_extra = 'Hiçbiri'

    applicable = []
    for c in selectable:
        group = 'Plus Ek İndirim' if c[0].startswith('Plus Ek İndirim %') else c[0]
        if group not in applicable:
            applicable.append(group)

    return (
        best_main, 
        CAMPAIGN_LABELS.get(best_main, best_main), 
        ', '.join(applicable), 
        best_extra, 
        CAMPAIGN_LABELS.get(best_extra, best_extra)
    )


choose_campaigns = choose_campaigns_smart


def build_discount_fields(is_eligible, current_price, dip_price, market_price):
    def discount(upper_price, lower_price):
        try:
            upper = to_float(upper_price)
            lower = to_float(lower_price)
        except (TypeError, ValueError):
            return None, None
        if upper is None or lower is None or pd.isna(upper) or pd.isna(lower) or upper <= 0 or lower > upper:
            return None, None
        amount = round(upper - lower, 2)
        return amount, round((amount / upper) * 100, 2)

    # Dip fiyat indirimli ürünlerde veya muhasebe fiyat listelerinde varsa is_eligible True kabul edilir
    available_amount, available_percent = (
        discount(current_price, dip_price) if is_eligible else (None, None)
    )
    current_amount, current_percent = discount(market_price, current_price)

    return {
        'Uygulanabilecek İndirim (TL)': available_amount,
        'Uygulanabilecek İndirim (%)': available_percent,
        'Mevcut İndirim (TL)': current_amount,
        'Mevcut İndirim (%)': current_percent,
        'Düşülebilecek Dip Fiyat (TL)': to_float(dip_price) if is_eligible and dip_price is not None and not pd.isna(dip_price) else None,
    }


def calculate_all(input_files, counter_files=None, plus_extra_files=None, coupon_files=None, karsilamali_config=None, output_dir=None, user_selections=None):
    required = ('commission', 'current')
    missing = [key for key in required if not input_files.get(key)]
    if missing:
        return {"success": False, "message": "Zorunlu girdi dosyaları eksik."}

    output_dir = output_dir or os.path.join(os.path.dirname(os.path.abspath(__file__)), 'Çıktılar')
    os.makedirs(output_dir, exist_ok=True)
    out_file = os.path.join(output_dir, 'Kampanya_Hesaplama_Sonuclari.xlsx')

    # Parse multi-counter files configuration if provided
    counter_items = []
    if counter_files and isinstance(counter_files, list):
        for idx, item in enumerate(counter_files):
            if item.get('enabled', True) is False:
                continue
            try:
                c_df = pd.read_excel(item['path']) if isinstance(item.get('path'), (str, Path)) else item.get('df', pd.DataFrame())
                if not c_df.empty and 'Barkod' in c_df.columns:
                    c_df['BARKOD_CLN'] = c_df['Barkod'].astype(str).str.strip()
                    item_dict = c_df.drop_duplicates(subset=['BARKOD_CLN']).set_index('BARKOD_CLN').to_dict('index')
                    
                    min_p = float(item.get('min_price', 0))
                    disc_amt = float(item.get('discount_amount', 0))
                    tr_pct = float(item.get('trendyol_percent', 0))
                    label = item.get('label') or f"Karşılamalı ({int(min_p) if min_p.is_integer() else min_p} TL Üzeri / {int(disc_amt) if disc_amt.is_integer() else disc_amt} TL İndirim)"
                    
                    counter_items.append({
                        'id': item.get('id') or f"counter_{idx+1}",
                        'label': label,
                        'min_price': min_p,
                        'discount_amount': disc_amt,
                        'trendyol_percent': tr_pct,
                        'dict': item_dict
                    })
            except Exception: pass
    elif input_files.get('counter'):
        try:
            c_df = pd.read_excel(input_files['counter'])
            if not c_df.empty and 'Barkod' in c_df.columns:
                c_df['BARKOD_CLN'] = c_df['Barkod'].astype(str).str.strip()
                item_dict = c_df.drop_duplicates(subset=['BARKOD_CLN']).set_index('BARKOD_CLN').to_dict('index')
                toplam_ind = float(karsilamali_config.get('toplam_indirim', 0)) if karsilamali_config else 0
                tr_oran = float(karsilamali_config.get('trendyol_oran', 0)) if karsilamali_config else 0
                counter_items.append({
                    'id': 'counter',
                    'label': 'Karşılamalı Kampanya',
                    'min_price': 0.0,
                    'discount_amount': toplam_ind,
                    'trendyol_percent': tr_oran,
                    'dict': item_dict
                })
        except Exception: pass

    # Parse multi plus_extra files configuration if provided
    plus_extra_items = []
    if plus_extra_files and isinstance(plus_extra_files, list):
        for idx, item in enumerate(plus_extra_files):
            if item.get('enabled', True) is False:
                continue
            try:
                pe_df = pd.read_excel(item['path']) if isinstance(item.get('path'), (str, Path)) else item.get('df', pd.DataFrame())
                if not pe_df.empty and 'Barkod' in pe_df.columns:
                    pe_df['BARKOD_CLN'] = pe_df['Barkod'].astype(str).str.strip()
                    item_dict = pe_df.drop_duplicates(subset=['BARKOD_CLN']).set_index('BARKOD_CLN').to_dict('index')
                    rate = float(item.get('rate', 0))
                    label = item.get('label') or (f"Plus Ek İndirim %{int(rate) if rate.is_integer() else rate}" if rate > 0 else f"Plus Ek İndirim #{idx+1}")
                    plus_extra_items.append({
                        'id': item.get('id') or f"plus_extra_{idx+1}",
                        'label': label,
                        'rate': rate,
                        'dict': item_dict
                    })
            except Exception: pass
    elif input_files.get('plus_extra'):
        try:
            pe_df = pd.read_excel(input_files['plus_extra'])
            if not pe_df.empty and 'Barkod' in pe_df.columns:
                pe_df['BARKOD_CLN'] = pe_df['Barkod'].astype(str).str.strip()
                item_dict = pe_df.drop_duplicates(subset=['BARKOD_CLN']).set_index('BARKOD_CLN').to_dict('index')
                for rate in [5.0, 10.0, 20.0]:
                    plus_extra_items.append({
                        'id': f"plus_extra_{int(rate)}",
                        'label': f"Plus Ek İndirim %{int(rate)}",
                        'rate': rate,
                        'dict': item_dict
                    })
        except Exception: pass

    # Parse multi coupon files configuration if provided
    coupon_items = []
    if coupon_files and isinstance(coupon_files, list):
        for idx, item in enumerate(coupon_files):
            if item.get('enabled', True) is False:
                continue
            try:
                cp_df = pd.read_excel(item['path']) if isinstance(item.get('path'), (str, Path)) else item.get('df', pd.DataFrame())
                if not cp_df.empty and 'Barkod' in cp_df.columns:
                    cp_df['BARKOD_CLN'] = cp_df['Barkod'].astype(str).str.strip()
                    item_dict = cp_df.drop_duplicates(subset=['BARKOD_CLN']).set_index('BARKOD_CLN').to_dict('index')
                    
                    min_p = float(item.get('min_price', 0))
                    disc_amt = float(item.get('discount_amount', 0))
                    tr_pct = float(item.get('trendyol_percent', 0))
                    min_p_str = f"{int(min_p)}" if min_p.is_integer() else f"{min_p}"
                    disc_str = f"{int(disc_amt)}" if disc_amt.is_integer() else f"{disc_amt}"
                    label = item.get('label') or f"{min_p_str} TL Üzerine {disc_str} TL Kupon - Trendyol Plus Müşterilerine Özel"
                    
                    coupon_items.append({
                        'id': item.get('id') or f"coupon_{idx+1}",
                        'label': label,
                        'min_price': min_p,
                        'discount_amount': disc_amt,
                        'trendyol_percent': tr_pct,
                        'dict': item_dict
                    })
            except Exception: pass

    try:
        df_tr = pd.read_excel(input_files['discount'])
        df_kom = pd.read_excel(input_files['commission'])
        df_gun = pd.read_excel(input_files['current'])
        df_av = pd.read_excel(input_files['advantage']) if input_files.get('advantage') else pd.DataFrame()
        df_flas = pd.read_excel(input_files['flash']) if input_files.get('flash') else pd.DataFrame()
        df_plus = pd.read_excel(input_files['plus']) if input_files.get('plus') else pd.DataFrame()
        df_plus_ek = pd.read_excel(input_files['plus_extra']) if input_files.get('plus_extra') else pd.DataFrame()
        df_muh_av = pd.read_excel(input_files['muhasebe_avantaj']) if input_files.get('muhasebe_avantaj') else pd.DataFrame()
        df_muh_flas = pd.read_excel(input_files['muhasebe_flas']) if input_files.get('muhasebe_flas') else pd.DataFrame()
        df_muh_plus = pd.read_excel(input_files['muhasebe_plus']) if input_files.get('muhasebe_plus') else pd.DataFrame()
    except Exception as e:
        return {"success": False, "message": f"Excel dosyaları okunurken hata oluştu: {str(e)}"}

    df_tr['BARKOD_CLN'] = df_tr['BARKOD'].astype(str).str.strip()
    if 'Durum' in df_tr.columns:
        indirimli = df_tr[df_tr['Durum'].astype(str).str.contains('ndirim', case=False, na=False)]
    else:
        indirimli = df_tr
    tr_barcodes = indirimli['BARKOD_CLN'].unique()

    df_gun['BARKOD_CLN'] = df_gun['Barkod'].astype(str).str.strip()
    gun_barcodes = df_gun['BARKOD_CLN'].unique()
    df_kom['BARKOD_CLN'] = df_kom['BARKOD'].astype(str).str.strip()
    barcodes = sorted(set(tr_barcodes).union(set(gun_barcodes)))

    for dataframe, barcode_column in (
        (df_av, 'BARKOD'),
        (df_flas, 'Barkod'),
        (df_plus, 'Barkod'),
        (df_plus_ek, 'Barkod'),
        (df_muh_av, 'BARKOD'),
        (df_muh_flas, 'Barkod'),
        (df_muh_plus, 'Barkod'),
    ):
        if not dataframe.empty and barcode_column in dataframe.columns:
            dataframe['BARKOD_CLN'] = dataframe[barcode_column].astype(str).str.strip()

    def to_dict_safe(df, key_col):
        if df.empty or key_col not in df.columns: return {}
        return df.drop_duplicates(subset=[key_col]).set_index(key_col).to_dict('index')

    dict_av = to_dict_safe(df_av, 'BARKOD_CLN')
    dict_flas = to_dict_safe(df_flas, 'BARKOD_CLN')
    dict_kom = to_dict_safe(df_kom, 'BARKOD_CLN')
    dict_gun = to_dict_safe(df_gun, 'BARKOD_CLN')
    dict_ind = to_dict_safe(indirimli, 'BARKOD_CLN')
    dict_plus = to_dict_safe(df_plus, 'BARKOD_CLN')
    dict_plus_ek = to_dict_safe(df_plus_ek, 'BARKOD_CLN')
    dict_muh_av = to_dict_safe(df_muh_av, 'BARKOD_CLN')
    dict_muh_flas = to_dict_safe(df_muh_flas, 'BARKOD_CLN')
    dict_muh_plus = to_dict_safe(df_muh_plus, 'BARKOD_CLN')

    results = []

    for b in barcodes:
        av_row = dict_av.get(b)
        fl_row = dict_flas.get(b)
        kom_row = dict_kom.get(b)
        gun_row = dict_gun.get(b)
        tr_row = dict_ind.get(b)
        plus_row = dict_plus.get(b)
        plus_ek_row = dict_plus_ek.get(b)

        muh_av_row = dict_muh_av.get(b)
        muh_flas_row = dict_muh_flas.get(b)
        muh_plus_row = dict_muh_plus.get(b)

        stok_val = None
        for row_src in (gun_row, kom_row, fl_row, plus_row, av_row):
            if row_src is not None:
                for s_col in ['Ürün Stok Adedi', 'STOK', 'Stok', 'Mevcut Stok', 'Stok Adedi', 'Stok Miktarı']:
                    if s_col in row_src:
                        val = to_float(row_src[s_col])
                        if val is not None:
                            stok_val = int(val) if val.is_integer() else val
                            break
                if stok_val is not None:
                    break

        match_av = (av_row is not None) or (muh_av_row is not None)
        match_fl = (fl_row is not None) or (muh_flas_row is not None)
        match_plus = (plus_row is not None) or (muh_plus_row is not None)
        match_plus_ek = plus_ek_row is not None

        is_indirim = tr_row is not None
        yeni_fiyat = to_float(tr_row['YENİ Fiyat']) if is_indirim and 'YENİ Fiyat' in tr_row else None
        eski_fiyat = to_float(tr_row['Eski Fiyat']) if is_indirim and 'Eski Fiyat' in tr_row else None
        
        piyasa_fiyat = None
        guncel_fiyat = None
        if gun_row is not None:
            try: piyasa_fiyat = to_float(gun_row['Piyasa Satış Fiyatı (KDV Dahil)'])
            except: pass
            try: guncel_fiyat = to_float(gun_row["Trendyol'da Satılacak Fiyat (KDV Dahil)"])
            except: pass

        if is_indirim and eski_fiyat and eski_fiyat > 0:
            guncel_fiyat_calc = eski_fiyat
        else:
            guncel_fiyat_calc = guncel_fiyat

        # Güncel net hesabı
        rate_1 = None
        net_1 = None
        if guncel_fiyat_calc and guncel_fiyat_calc > 0:
            rate_1 = get_commission_rate(guncel_fiyat_calc, kom_row) if kom_row else None
            if rate_1 is None and gun_row:
                try: rate_1 = to_float(gun_row['Komisyon Oranı'])
                except: pass
            if rate_1 is not None:
                net_1 = round(guncel_fiyat_calc - (guncel_fiyat_calc * (rate_1 / 100.0)), 2)

        # Dip Fiyat tespiti (SADECE Satıcı İndirim Listesi ve Muhasebe Dip Limitleri)
        common_dip = yeni_fiyat if (is_indirim and yeni_fiyat and yeni_fiyat > 0) else None

        av_dip = None
        av_dip_source = None
        if muh_av_row is not None:
            for col_name in ['YENİ TSF (FİYAT GÜNCELLE)', '1 YILDIZ ÜST FİYAT']:
                if col_name in muh_av_row:
                    val = to_float(muh_av_row[col_name])
                    if val and val > 0: 
                        av_dip = val
                        av_dip_source = 'Avantajlı Muhasebe'
                        break

        flas_dip = None
        flas_dip_source = None
        if muh_flas_row is not None:
            for col_name in ['Senin Belirlediğin Flaş Fiyatı', '24 Saat Fiyat']:
                if col_name in muh_flas_row:
                    val = to_float(muh_flas_row[col_name])
                    if val and val > 0: 
                        flas_dip = val
                        flas_dip_source = 'Flaş Muhasebe'
                        break

        plus_dip = None
        plus_dip_source = None
        if muh_plus_row is not None:
            for col_name in ['Plus Fiyat Üst Limiti', 'Plus Fiyat Seçimi']:
                if col_name in muh_plus_row:
                    val = to_float(muh_plus_row[col_name])
                    if val and val > 0: 
                        plus_dip = val
                        plus_dip_source = 'Plus Muhasebe'
                        break

        dip_details = []
        if common_dip is not None:
            dip_details.append({'type': 'İndirim Listesi', 'price': common_dip})
        if plus_dip is not None and plus_dip_source:
            dip_details.append({'type': plus_dip_source, 'price': plus_dip})
        if flas_dip is not None and flas_dip_source:
            dip_details.append({'type': flas_dip_source, 'price': flas_dip})
        if av_dip is not None and av_dip_source:
            dip_details.append({'type': av_dip_source, 'price': av_dip})

        av_threshold = min([d for d in [common_dip, av_dip] if d is not None], default=None)
        flas_threshold = min([d for d in [common_dip, flas_dip] if d is not None], default=None)
        plus_threshold = min([d for d in [common_dip, plus_dip] if d is not None], default=None)
        common_threshold = common_dip

        all_dips = [d['price'] for d in dip_details]
        min_dip_val = min(all_dips) if all_dips else None

        # Ürünün katılabileceği kampanyaların tespiti (eligible_campaigns)
        eligible_campaigns = ['Hiçbiri']
        smart_candidates = []

        # 1. Avantajlı Kampanya
        yeni_tsf = None
        yeni_tsf_is_fallback = False
        rate_2 = None
        net_2 = None
        if match_av:
            if muh_av_row is not None:
                for col_name in ['YENİ TSF (FİYAT GÜNCELLE)', '1 YILDIZ ÜST FİYAT', 'TRENDYOL SATIŞ FİYATI']:
                    if col_name in muh_av_row:
                        val = to_float(muh_av_row[col_name])
                        if val and val > 0: yeni_tsf = val; break
            if yeni_tsf is None and av_row is not None:
                for col_name in ['YENİ TSF (FİYAT GÜNCELLE)', '1 YILDIZ ÜST FİYAT']:
                    if col_name in av_row:
                        val = to_float(av_row[col_name])
                        if val and val > 0: yeni_tsf = val; break

            if yeni_tsf is None and guncel_fiyat_calc and guncel_fiyat_calc > 0:
                yeni_tsf = guncel_fiyat_calc
                yeni_tsf_is_fallback = True

            if yeni_tsf and yeni_tsf > 0:
                rate_2 = get_commission_rate(yeni_tsf, kom_row) if kom_row else None
                if rate_2 is None and gun_row:
                    try: rate_2 = to_float(gun_row['Komisyon Oranı'])
                    except: pass
                if rate_2 is not None:
                    net_2 = round(yeni_tsf - (yeni_tsf * (rate_2 / 100.0)), 2)

                if av_threshold is None or yeni_tsf >= av_threshold - 0.01:
                    eligible_campaigns.append('Avantajlı')
                    if net_2 is not None:
                        smart_candidates.append(('Avantajlı', net_2, yeni_tsf, rate_2, yeni_tsf_is_fallback, muh_av_row is not None))

        # 2. Flaş Kampanya
        f_24_fiyat = None
        f_24_fiyat_is_fallback = False
        rate_3 = None
        net_3 = None
        if match_fl:
            if muh_flas_row is not None:
                for col_name in ['Senin Belirlediğin Flaş Fiyatı', '24 Saat Fiyat', '3 Saat Fiyat', 'Mevcut Fiyat']:
                    if col_name in muh_flas_row:
                        val = to_float(muh_flas_row[col_name])
                        if val and val > 0: f_24_fiyat = val; break
            if f_24_fiyat is None and fl_row is not None:
                for col_name in ['24 Saat Fiyat', '3 Saat Fiyat']:
                    if col_name in fl_row:
                        val = to_float(fl_row[col_name])
                        if val and val > 0: f_24_fiyat = val; break

            if f_24_fiyat is None and guncel_fiyat_calc and guncel_fiyat_calc > 0:
                f_24_fiyat = guncel_fiyat_calc
                f_24_fiyat_is_fallback = True

            if f_24_fiyat and f_24_fiyat > 0:
                rate_3 = get_commission_rate(f_24_fiyat, kom_row) if kom_row else None
                if rate_3 is None and gun_row:
                    try: rate_3 = to_float(gun_row['Komisyon Oranı'])
                    except: pass
                if rate_3 is not None:
                    net_3 = round(f_24_fiyat - (f_24_fiyat * (rate_3 / 100.0)), 2)

                if flas_threshold is None or f_24_fiyat >= flas_threshold - 0.01:
                    eligible_campaigns.append('Flaş')
                    if net_3 is not None:
                        smart_candidates.append(('Flaş', net_3, f_24_fiyat, rate_3, f_24_fiyat_is_fallback, muh_flas_row is not None))

        # 3. Plus Kampanya
        plus_fiyat = None
        plus_fiyat_is_fallback = False
        rate_4 = None
        net_4 = None
        if match_plus:
            if muh_plus_row is not None:
                for col_name in ['Plus Fiyat Üst Limiti', 'Güncel TSF']:
                    if col_name in muh_plus_row:
                        val = to_float(muh_plus_row[col_name])
                        if val and val > 0: plus_fiyat = val; break
            if plus_fiyat is None and plus_row is not None:
                try: plus_fiyat = to_float(plus_row['Plus Fiyat Üst Limiti'])
                except: pass

            if plus_fiyat is None and guncel_fiyat_calc and guncel_fiyat_calc > 0:
                plus_fiyat = guncel_fiyat_calc
                plus_fiyat_is_fallback = True

            if plus_fiyat and plus_fiyat > 0:
                if plus_row and 'Plus Komisyon Teklifi' in plus_row:
                    try: rate_4 = to_float(str(plus_row['Plus Komisyon Teklifi']).replace(',', '.'))
                    except: pass
                if rate_4 is None and kom_row:
                    rate_4 = get_commission_rate(plus_fiyat, kom_row)
                if rate_4 is None and gun_row:
                    try: rate_4 = to_float(gun_row['Komisyon Oranı'])
                    except: pass
                if rate_4 is not None:
                    net_4 = round(plus_fiyat - (plus_fiyat * (rate_4 / 100.0)), 2)

                if plus_threshold is None or plus_fiyat >= plus_threshold - 0.01:
                    eligible_campaigns.append('Plus')
                    if net_4 is not None:
                        smart_candidates.append(('Plus', net_4, plus_fiyat, rate_4, plus_fiyat_is_fallback, muh_plus_row is not None))


        # 4. Plus Ek İndirim Kampanyaları (Çoklu Yükleme Desteği)
        if plus_extra_items:
            for pe_item in plus_extra_items:
                pe_dict = pe_item['dict']
                pe_row = pe_dict.get(b)
                if pe_row is not None:
                    pe_price = None
                    pe_price_is_fallback = False
                    for pe_col in ['Maksimum Girebileceğin Fiyat', 'Kampanyalı Satış Fiyatı']:
                        if pe_col in pe_row:
                            val = to_float(pe_row[pe_col])
                            if val and val > 0: pe_price = val; break
                    if pe_price is None and guncel_fiyat_calc and guncel_fiyat_calc > 0:
                        pe_price = guncel_fiyat_calc
                        pe_price_is_fallback = True
                    
                    if pe_price and pe_price > 0:
                        pe_rate = get_commission_rate(pe_price, kom_row) if kom_row else None
                        if pe_rate is None and gun_row:
                            try: pe_rate = to_float(gun_row['Komisyon Oranı'])
                            except: pass
                        if pe_rate is not None:
                            disc_ratio = (1.0 - (pe_item['rate'] / 100.0)) if pe_item['rate'] > 0 else 1.0
                            calc_pe_price = round(pe_price * disc_ratio, 2)
                            pe_net = round(calc_pe_price - (pe_price * (pe_rate / 100.0)), 2)
                            if common_threshold is None or calc_pe_price >= common_threshold - 0.01:
                                eligible_campaigns.append(pe_item['label'])
                                smart_candidates.append((pe_item['label'], pe_net, calc_pe_price, pe_rate, pe_price_is_fallback))

        # 4. Plus Ek İndirim (%5, %10, %20)
        plus_ek_fiyat = None
        plus_ek_fiyat_is_fallback = False
        rate_5 = None
        net_5 = None
        plus_ek_fiyat_5 = None
        net_5_5 = None
        plus_ek_fiyat_10 = None
        net_5_10 = None
        plus_ek_fiyat_20 = None
        net_5_20 = None
        if match_plus_ek and plus_ek_row is not None:
            try: plus_ek_fiyat = to_float(plus_ek_row['Maksimum Girebileceğin Fiyat'])
            except: pass
            if plus_ek_fiyat is None and guncel_fiyat_calc and guncel_fiyat_calc > 0:
                plus_ek_fiyat = guncel_fiyat_calc
                plus_ek_fiyat_is_fallback = True
            if plus_ek_fiyat and plus_ek_fiyat > 0:
                rate_5 = get_commission_rate(plus_ek_fiyat, kom_row) if kom_row else None
                if rate_5 is None and gun_row:
                    try: rate_5 = to_float(gun_row['Komisyon Oranı'])
                    except: pass
                if rate_5 is not None:
                    net_5 = round(plus_ek_fiyat - (plus_ek_fiyat * (rate_5 / 100.0)), 2)
                    plus_ek_fiyat_5 = round(plus_ek_fiyat * 0.95, 2)
                    net_5_5 = round(plus_ek_fiyat_5 - (plus_ek_fiyat * (rate_5 / 100.0)), 2)
                    plus_ek_fiyat_10 = round(plus_ek_fiyat * 0.90, 2)
                    net_5_10 = round(plus_ek_fiyat_10 - (plus_ek_fiyat * (rate_5 / 100.0)), 2)
                    plus_ek_fiyat_20 = round(plus_ek_fiyat * 0.80, 2)
                    net_5_20 = round(plus_ek_fiyat_20 - (plus_ek_fiyat * (rate_5 / 100.0)), 2)
                    
                    for candidate in (
                        ('Plus Ek İndirim %5', net_5_5, plus_ek_fiyat_5, rate_5, plus_ek_fiyat_is_fallback),
                        ('Plus Ek İndirim %10', net_5_10, plus_ek_fiyat_10, rate_5, plus_ek_fiyat_is_fallback),
                        ('Plus Ek İndirim %20', net_5_20, plus_ek_fiyat_20, rate_5, plus_ek_fiyat_is_fallback),
                    ):
                        if common_threshold is None or candidate[2] >= common_threshold - 0.01:
                            if candidate[0] not in eligible_campaigns:
                                eligible_campaigns.append(candidate[0])
                            smart_candidates.append(candidate)

        # 5. Karşılamalı Kampanyalar (Çoklu)
        counter_evaluations = {}
        for c_item in counter_items:
            c_dict = c_item['dict']
            c_row = c_dict.get(b)
            if c_row is not None:
                c_price = None
                c_price_is_fallback = False
                for c_col in ['Maksimum Girebileceğin Fiyat', 'Kampanyalı Satış Fiyatı']:
                    if c_col in c_row:
                        val = to_float(c_row[c_col])
                        if val and val > 0: c_price = val; break
                if c_price is None and guncel_fiyat_calc and guncel_fiyat_calc > 0:
                    c_price = guncel_fiyat_calc
                    c_price_is_fallback = True
                
                # Minimum tutar kontrolü (örn: 300 TL üzeri)
                if c_price and c_price > 0:
                    eligible_campaigns.append(c_item['label'])
                    if common_threshold is None or c_price >= common_threshold - 0.01:
                        c_rate = get_commission_rate(c_price, kom_row) if kom_row else None
                        if c_rate is None and gun_row:
                            try: c_rate = to_float(gun_row['Komisyon Oranı'])
                            except: pass
                        if c_rate is not None:
                            disc_type = c_item.get('discount_type') or c_item.get('discount_unit') or 'TL'
                            disc_val = c_item.get('discount_amount', 0.0)
                            if disc_type == '%':
                                total_discount = round(c_price * (disc_val / 100.0), 2)
                            else:
                                total_discount = disc_val
                            
                            seller_disc = round(total_discount * (1.0 - (c_item['trendyol_percent'] / 100.0)), 2)
                            c_net = round(c_price - (c_price * (c_rate / 100.0)) - seller_disc, 2)
                            counter_evaluations[c_item['label']] = {
                                'price': c_price,
                                'rate': c_rate,
                                'net': c_net,
                                'seller_disc': seller_disc,
                                'disc_type': disc_type,
                                'disc_val': disc_val,
                                'trendyol_percent': c_item.get('trendyol_percent', 0.0),
                            }
                            smart_candidates.append((c_item['label'], c_net, c_price, c_rate, c_price_is_fallback))

        if coupon_items:
            for cp_item in coupon_items:
                cp_dict = cp_item['dict']
                cp_row = cp_dict.get(b)
                if cp_row is not None:
                    cp_price = None
                    cp_price_is_fallback = False
                    for cp_col in ['Mevcut Satış Fiyatı', 'Maksimum Girebileceğin Fiyat', 'Kampanyalı Satış Fiyatı']:
                        if cp_col in cp_row:
                            val = to_float(cp_row[cp_col])
                            if val and val > 0: cp_price = val; break
                    if cp_price is None and guncel_fiyat_calc and guncel_fiyat_calc > 0:
                        cp_price = guncel_fiyat_calc
                        cp_price_is_fallback = True
                    
                    if cp_price and cp_price > 0:
                        eligible_campaigns.append(cp_item['label'])
                        if common_threshold is None or cp_price >= common_threshold - 0.01:
                            cp_rate = get_commission_rate(cp_price, kom_row) if kom_row else None
                            if cp_rate is None and gun_row:
                                try: cp_rate = to_float(gun_row['Komisyon Oranı'])
                                except: pass
                            if cp_rate is not None:
                                disc_type = 'TL'
                                disc_val = cp_item.get('discount_amount', 0.0)
                                total_discount = disc_val
                                seller_disc = round(total_discount * (1.0 - (cp_item['trendyol_percent'] / 100.0)), 2)
                                cp_net = round(cp_price - (cp_price * (cp_rate / 100.0)) - seller_disc, 2)
                                counter_evaluations[cp_item['label']] = {
                                    'price': cp_price,
                                    'rate': cp_rate,
                                    'net': cp_net,
                                    'seller_disc': seller_disc,
                                    'disc_type': disc_type,
                                    'disc_val': disc_val,
                                    'trendyol_percent': cp_item.get('trendyol_percent', 0.0),
                                }
                                smart_candidates.append((cp_item['label'], cp_net, cp_price, cp_rate, cp_price_is_fallback))

        selectable = selectable_campaigns(net_1, smart_candidates)

        all_matching_main_campaigns = ['Hiçbiri']
        if match_av and yeni_tsf and (not yeni_tsf_is_fallback or (net_2 is not None and net_2 > net_1)):
            if 'Avantajlı' not in all_matching_main_campaigns: all_matching_main_campaigns.append('Avantajlı')
        if match_fl and f_24_fiyat and (not f_24_fiyat_is_fallback or (net_3 is not None and net_3 > net_1)):
            if 'Flaş' not in all_matching_main_campaigns: all_matching_main_campaigns.append('Flaş')
        if match_plus and plus_fiyat and (not plus_fiyat_is_fallback or (net_4 is not None and net_4 > net_1)):
            if 'Plus' not in all_matching_main_campaigns: all_matching_main_campaigns.append('Plus')

        eligible_main_campaigns = ['Hiçbiri'] + [c[0] for c in selectable if c[0] in MAIN_CAMPAIGN_KEYS]

        all_matching_extra_campaigns = ['Hiçbiri']
        if match_plus_ek and plus_ek_fiyat and (not plus_ek_fiyat_is_fallback or (net_5 is not None and net_5 > net_1)):
            for c_name in ('Plus Ek İndirim %5', 'Plus Ek İndirim %10', 'Plus Ek İndirim %20'):
                if c_name not in all_matching_extra_campaigns:
                    all_matching_extra_campaigns.append(c_name)
        if plus_extra_items:
            for pe_item in plus_extra_items:
                if pe_item['dict'].get(b) is not None:
                    if pe_item['label'] not in all_matching_extra_campaigns:
                        all_matching_extra_campaigns.append(pe_item['label'])
        if coupon_items:
            for cp_item in coupon_items:
                if cp_item['dict'].get(b) is not None:
                    if cp_item['label'] not in all_matching_extra_campaigns:
                        all_matching_extra_campaigns.append(cp_item['label'])
        for c_item in counter_items:
            if c_item['dict'].get(b) is not None:
                if c_item['label'] not in all_matching_extra_campaigns:
                    all_matching_extra_campaigns.append(c_item['label'])

        eligible_extra_campaigns = ['Hiçbiri'] + [c[0] for c in selectable if c[0] not in MAIN_CAMPAIGN_KEYS]

        all_matching_campaigns = list(all_matching_main_campaigns)
        for c_name in all_matching_extra_campaigns:
            if c_name not in all_matching_campaigns:
                all_matching_campaigns.append(c_name)

        eligible_campaigns = list(eligible_main_campaigns)
        for c_name in eligible_extra_campaigns:
            if c_name not in eligible_campaigns:
                eligible_campaigns.append(c_name)

        rec_res = choose_campaigns_smart(net_1, selectable)
        if len(rec_res) == 5:
            _rec_kampanya, _rec_kampanya_display, uygulanabilir_kampanyalar, _rec_extra, _rec_extra_display = rec_res
        else:
            _rec_kampanya, _rec_kampanya_display, uygulanabilir_kampanyalar = rec_res
            _rec_extra = 'Hiçbiri'

        if user_selections is not None and isinstance(user_selections, dict) and b in user_selections:
            saved_val = user_selections[b]
            if isinstance(saved_val, dict):
                saved_main = saved_val.get('main', 'Hiçbiri')
                saved_extra = saved_val.get('extra', 'Hiçbiri')
                ilk_main = saved_main if (saved_main == 'Hiçbiri' or saved_main in all_matching_main_campaigns) else 'Hiçbiri'
                ilk_extra = saved_extra if (saved_extra == 'Hiçbiri' or saved_extra in all_matching_extra_campaigns) else 'Hiçbiri'
            else:
                saved_str = str(saved_val)
                if saved_str in all_matching_main_campaigns:
                    ilk_main = saved_str
                    ilk_extra = 'Hiçbiri'
                elif saved_str in all_matching_extra_campaigns:
                    ilk_main = 'Hiçbiri'
                    ilk_extra = saved_str
                else:
                    ilk_main = 'Hiçbiri'
                    ilk_extra = 'Hiçbiri'
        else:
            ilk_main = 'Hiçbiri'
            ilk_extra = 'Hiçbiri'

        guncel_fiyat_display = guncel_fiyat_calc
        discount_fields = build_discount_fields(
            is_indirim or (len(dip_details) > 0),
            guncel_fiyat_display,
            min_dip_val,
            piyasa_fiyat,
        )

        mevcut_indirim_orani = None
        if piyasa_fiyat and guncel_fiyat_display and piyasa_fiyat > guncel_fiyat_display:
            mevcut_indirim_orani = round(((piyasa_fiyat - guncel_fiyat_display) / piyasa_fiyat) * 100.0, 2)

        results.append({
            'Barkod': b,
            'Stok Adedi': stok_val,
            'dip_details': dip_details,
            'Plus Ek İndirim Eşleşme Durumu': 'Eşleşti' if match_plus_ek else 'Eşleşme Yok',
            'Plus Eşleşme Durumu': 'Eşleşti' if match_plus else 'Eşleşme Yok',
            'Avantajlı Ürün Eşleşme Durumu': 'Eşleşti' if match_av else 'Eşleşme Yok',
            'Flaş Ürün Eşleşme Durumu': 'Eşleşti' if match_fl else 'Eşleşme Yok',
            'Hem Avantajlı Hem Flaş': 'Evet' if (match_av and match_fl) else 'Hayır',
            'İndirim Uygulanabilir': 'Evet' if (common_dip is not None or len(dip_details) > 0) else 'Hayır',
            'Mevcut İndirim Oranı (%)': mevcut_indirim_orani,
            'Uygulanabilir Kampanyalar': uygulanabilir_kampanyalar,
            'Önerilen Kampanya': _rec_kampanya,
            'Önerilen Ekstra Kampanya': _rec_extra,
            'İlk Kampanya Seçimi': ilk_main,
            'İlk Ekstra Kampanya Seçimi': ilk_extra,
            'Hangisi Daha Karlı?': 'Hiçbiri',
            'Karlılık Farkı (%)': '',
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
            'Uygulanabilecek İndirim (TL)': discount_fields['Uygulanabilecek İndirim (TL)'],
            'Uygulanabilecek İndirim (%)': discount_fields['Uygulanabilecek İndirim (%)'],
            'Mevcut İndirim (TL)': discount_fields['Mevcut İndirim (TL)'],
            'Mevcut İndirim (%)': discount_fields['Mevcut İndirim (%)'],
            'Düşülebilecek Dip Fiyat (TL)': min_dip_val,
            'eligible_main_campaigns': eligible_main_campaigns,
            'all_matching_main_campaigns': all_matching_main_campaigns,
            'eligible_extra_campaigns': eligible_extra_campaigns,
            'all_matching_extra_campaigns': all_matching_extra_campaigns,
            'eligible_campaigns': eligible_campaigns,
            'all_matching_campaigns': all_matching_campaigns,
            'counter_evaluations': counter_evaluations,
        })

    # Save to Excel
    out_file = os.path.join(output_dir, 'Kampanya_Hesaplama_Sonuclari.xlsx') if output_dir else 'Kampanya_Hesaplama_Sonuclari.xlsx'
    try:
        out_df = pd.DataFrame(results)
        excel_cols = [c for c in out_df.columns if c not in ('eligible_campaigns', 'counter_evaluations', 'eligible_main_campaigns', 'all_matching_main_campaigns', 'eligible_extra_campaigns', 'all_matching_extra_campaigns')]
        out_df[excel_cols].to_excel(out_file, index=False)
    except Exception as e:
        print("Excel kaydetme uyarısı:", e)

    return {"success": True, "output_path": out_file, "results": results, "counter_items": counter_items}
