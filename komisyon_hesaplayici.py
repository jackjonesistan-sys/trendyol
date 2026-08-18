import os
import re
import math
from datetime import date, datetime
import pandas as pd
from pathlib import Path

from input_files import (
    build_campaign_label,
    choose_plus_tariff_label,
    find_plus_period_columns,
    normalize_campaign_config,
    normalize_recommendation_rule,
)

ROUNDING_EPSILON = 1e-9

CAMPAIGN_LABELS = {
    "Avantajlı": "Avantajlı Ürün",
    "Flaş": "Flaş Ürün",
    "Plus": "Plus Ürün",
    "Komisyon Tarifesi": "Komisyon Tarifesi",
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


def find_commission_tariff_period_label(kom_row):
    if not kom_row or not isinstance(kom_row, dict):
        return "7 Günlük Fiyat"
    for col in kom_row.keys():
        m = re.search(r"Tarih\s*aral[ıi][ğg][ıi]\s*\((\d+)\s*G[üu]n\)", str(col), re.IGNORECASE)
        if m:
            days = m.group(1)
            return f"{days} Günlük Fiyat"
    tarife_val = kom_row.get("Tarife Seçimi")
    if tarife_val and not pd.isna(tarife_val) and str(tarife_val).strip() != "":
        return str(tarife_val).strip()
    return "7 Günlük Fiyat"


def evaluate_commission_tariff_bracket(kom_row, dip_price=None, dip_net=None, guncel_fiyat=None):
    if not kom_row or not isinstance(kom_row, dict):
        return None
    if dip_price is None or dip_net is None:
        return None

    has_tier = any(
        to_float(kom_row.get(col)) is not None and to_float(kom_row.get(col)) > 0
        for col in ("2.Fiyat Üst Limiti", "3.Fiyat Üst Limiti", "4.Fiyat Üst Limiti")
    )
    if not has_tier:
        return None

    brackets = [
        {
            "kademe_no": 4,
            "kademe_adi": "4. Kademe",
            "price": to_float(kom_row.get("4.Fiyat Üst Limiti")),
            "rate": to_float(kom_row.get("4.KOMİSYON")),
        },
        {
            "kademe_no": 3,
            "kademe_adi": "3. Kademe",
            "price": to_float(kom_row.get("3.Fiyat Üst Limiti")),
            "rate": to_float(kom_row.get("3.KOMİSYON")),
        },
        {
            "kademe_no": 2,
            "kademe_adi": "2. Kademe",
            "price": to_float(kom_row.get("2.Fiyat Üst Limiti")),
            "rate": to_float(kom_row.get("2.KOMİSYON")),
        },
        {
            "kademe_no": 1,
            "kademe_adi": "1. Kademe",
            "price": to_float(kom_row.get("1.Fiyat Alt Limit")),
            "rate": to_float(kom_row.get("1.KOMİSYON")),
        },
    ]

    valid_candidates = []
    for b in brackets:
        p = b["price"]
        r = b["rate"]
        if p is not None and p > 0 and r is not None:
            net = round(p - (p * (r / 100.0)), 2)
            b["net"] = net
            passes_dip_net = (net >= dip_net - 0.01)
            b["eligible"] = bool(passes_dip_net and net > 0)
            valid_candidates.append(b)

    if not valid_candidates:
        return None

    eligible_candidates = [b for b in valid_candidates if b["eligible"]]
    tariff_selection = find_commission_tariff_period_label(kom_row)

    if eligible_candidates:
        best = min(eligible_candidates, key=lambda x: (x["price"], -x["net"]))
        return {
            **best,
            "tariff_selection": tariff_selection,
            "has_eligible": True,
        }
    else:
        best = min(valid_candidates, key=lambda x: (x["price"], -x["net"]))
        return {
            **best,
            "tariff_selection": tariff_selection,
            "has_eligible": False,
        }


def fallback_candidate_is_selectable(current_net, candidate_net, used_current_price):
    return not used_current_price or (
        current_net is not None
        and candidate_net is not None
        and candidate_net > current_net
    )


def selectable_campaigns(current_net, candidates):
    return [
        candidate
        for candidate in candidates
        if fallback_candidate_is_selectable(
            current_net,
            candidate[1],
            len(candidate) > 4 and candidate[4],
        )
    ]


ALLOWED_FOR_RECOMMENDATION = {'Avantajlı', 'Flaş', 'Plus', 'Komisyon Tarifesi'}
MAIN_CAMPAIGN_KEYS = {'Avantajlı', 'Flaş', 'Plus', 'Komisyon Tarifesi'}


def choose_campaigns_smart(current_net, candidates, recommendation_rule=None):
    """
    candidates: (campaign_key, net_price, eff_price, rate, used_current_price=False, is_muhasebe=False)
    
    Ana kampanyayı etkin kural sırasına, aksi halde kalan nete göre seçer.
    Ekstra kampanya her zaman en yüksek kalan nete göre seçilir.
    """
    rule = normalize_recommendation_rule(recommendation_rule)
    selectable = selectable_campaigns(current_net, candidates)
    if not selectable:
        return 'Hiçbiri', 'Hiçbiri', '', 'Hiçbiri', 'Hiçbiri'

    main_selectable = [c for c in selectable if c[0] in MAIN_CAMPAIGN_KEYS]
    if main_selectable:
        if rule["enabled"]:
            selectable_main_names = {candidate[0] for candidate in main_selectable}
            priority_order = list(rule["priority"])
            for name in MAIN_CAMPAIGN_KEYS:
                if name not in priority_order:
                    priority_order.append(name)
            best_main = next(
                (name for name in priority_order if name in selectable_main_names),
                min(main_selectable, key=lambda x: (-x[1], -x[2], x[0]))[0]
            )
        else:
            best_main = min(
                main_selectable, key=lambda x: (-x[1], -x[2], x[0])
            )[0]
    else:
        best_main = 'Hiçbiri'

    extra_selectable = [c for c in selectable if c[0] not in MAIN_CAMPAIGN_KEYS]
    if extra_selectable:
        best_extra = min(extra_selectable, key=lambda x: (-x[1], -x[2], x[0]))[0]
    else:
        best_extra = 'Hiçbiri'

    applicable = []
    for c in selectable:
        group = 'Plus Ek İndirim' if c[0].startswith('Plus Ek İndirim') else c[0]
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


def build_extra_evaluation(base_price, commission_rate, config):
    price = to_float(base_price)
    rate = to_float(commission_rate)
    if price is None or rate is None:
        return None
    discount_value = config["discount_amount"]
    total_discount = (
        round(price * (discount_value / 100.0), 2)
        if config["discount_type"] == "%"
        else discount_value
    )
    customer_price = round(price - total_discount, 2)
    seller_discount = round(
        total_discount * (1.0 - (config["trendyol_percent"] / 100.0)), 2
    )
    return {
        "customer_price": customer_price,
        "price": price,
        "rate": rate,
        "net": round(price - (price * (rate / 100.0)) - seller_discount, 2),
        "seller_disc": seller_discount,
        "min_price": config["min_price"],
        "disc_type": config["discount_type"],
        "disc_val": discount_value,
        "trendyol_percent": config["trendyol_percent"],
    }


FLASH_PERIOD_COLUMNS = (
    (
        "24 Saat",
        "24 Saat Fiyat",
        "24 Saat Flaş Başlangıç Tarihi",
        "24 Saat Flaş Bitiş Tarihi",
    ),
    (
        "3 Saat",
        "3 Saat Fiyat",
        "3 Saat Flaş Başlangıç Tarihi",
        "3 Saat Flaş Bitiş Tarihi",
    ),
)


def normalize_flash_interval_value(value):
    if value is None or pd.isna(value):
        return None
    if isinstance(value, (datetime, date, pd.Timestamp)):
        return pd.Timestamp(value).strftime("%Y-%m-%d %H:%M:%S")
    text = str(value).strip()
    if not text:
        return None
    for date_format in (
        "%d/%m/%Y %H:%M",
        "%d/%m/%Y %H:%M:%S",
        "%Y-%m-%d %H:%M",
        "%Y-%m-%d %H:%M:%S",
    ):
        try:
            return datetime.strptime(text, date_format).strftime("%Y-%m-%d %H:%M:%S")
        except ValueError:
            continue
    return text


def flash_interval_candidates(rows, origin):
    candidates = []
    for row in rows:
        fixed_price = to_float(row.get("Senin Belirlediğin Flaş Fiyatı"))
        row_candidates = []
        for period, price_column, start_column, end_column in FLASH_PERIOD_COLUMNS:
            period_price = to_float(row.get(price_column))
            start = normalize_flash_interval_value(row.get(start_column))
            end = normalize_flash_interval_value(row.get(end_column))
            interval_exists = period_price is not None or start is not None or end is not None
            if not interval_exists:
                continue
            price = fixed_price if fixed_price is not None and fixed_price > 0 else period_price
            source = "Senin Belirlediğin Flaş Fiyatı" if price == fixed_price else price_column
            if (price is None or price <= 0) and origin == "muhasebe":
                price = to_float(row.get("Mevcut Fiyat"))
                source = "Mevcut Fiyat"
            if price is not None and price > 0:
                row_candidates.append({
                    "period": period,
                    "start": start,
                    "end": end,
                    "price": price,
                    "source": source,
                    "origin": origin,
                    "used_current_price": False,
                })

        legacy_price = fixed_price
        legacy_source = "Senin Belirlediğin Flaş Fiyatı"
        if (legacy_price is None or legacy_price <= 0) and origin == "muhasebe":
            legacy_price = to_float(row.get("Mevcut Fiyat"))
            legacy_source = "Mevcut Fiyat"
        if not row_candidates and legacy_price is not None and legacy_price > 0:
            row_candidates.append({
                "period": "24 Saat",
                "start": None,
                "end": None,
                "price": legacy_price,
                "source": legacy_source,
                "origin": origin,
                "used_current_price": False,
            })
        candidates.extend(row_candidates)
    return candidates


def merge_flash_intervals(campaign_intervals, accounting_intervals):
    if not campaign_intervals:
        return list(accounting_intervals)

    used_accounting = set()
    period_counts = {
        period: sum(item["period"] == period for item in campaign_intervals)
        for period in {item["period"] for item in campaign_intervals}
    }
    merged = []
    for campaign_interval in campaign_intervals:
        key = tuple(campaign_interval[field] for field in ("period", "start", "end"))
        exact_matches = [
            index
            for index, item in enumerate(accounting_intervals)
            if index not in used_accounting
            and tuple(item[field] for field in ("period", "start", "end")) == key
        ]
        match_index = exact_matches[0] if len(exact_matches) == 1 else None
        if not exact_matches and period_counts[campaign_interval["period"]] == 1:
            generic_matches = [
                index
                for index, item in enumerate(accounting_intervals)
                if index not in used_accounting
                if item["period"] == campaign_interval["period"]
                and item["start"] is None
                and item["end"] is None
            ]
            match_index = generic_matches[0] if len(generic_matches) == 1 else None
        if match_index is None:
            merged.append(campaign_interval)
        else:
            used_accounting.add(match_index)
            merged.append({
                **accounting_intervals[match_index],
                "period": campaign_interval["period"],
                "start": campaign_interval["start"],
                "end": campaign_interval["end"],
            })
    return merged


def calculate_all(input_files, counter_files=None, plus_extra_files=None, coupon_files=None, net_discount_config=None, karsilamali_config=None, output_dir=None, user_selections=None, recommendation_rule=None):
    recommendation_rule = normalize_recommendation_rule(recommendation_rule)
    required = ('discount', 'commission', 'current')
    missing = [key for key in required if not input_files.get(key)]
    if missing:
        return {"success": False, "message": "Zorunlu girdi dosyaları eksik."}

    output_dir = output_dir or os.path.join(os.path.dirname(os.path.abspath(__file__)), 'Çıktılar')
    os.makedirs(output_dir, exist_ok=True)
    out_file = os.path.join(output_dir, 'Kampanya_Hesaplama_Sonuclari.xlsx')

    # Parse net_discount configuration if provided
    net_discount_item = None
    if net_discount_config and isinstance(net_discount_config, dict):
        nd_cfg = normalize_campaign_config(net_discount_config, "net_discount")
        if nd_cfg.get('enabled', True) is not False and nd_cfg.get('discount_amount', 0) > 0:
            nd_path = nd_cfg.get('path') or nd_cfg.get('stored_path') or input_files.get('net_discount')
            if nd_path and os.path.exists(nd_path):
                try:
                    nd_df = pd.read_excel(nd_path)
                    if not nd_df.empty and 'Barkod' in nd_df.columns:
                        nd_df = nd_df.assign(BARKOD_CLN=nd_df['Barkod'].astype(str).str.strip())
                        item_dict = nd_df.drop_duplicates(subset=['BARKOD_CLN']).set_index('BARKOD_CLN').to_dict('index')
                        net_discount_item = {
                            'id': nd_cfg.get('id') or 'net_discount_1',
                            'label': build_campaign_label(nd_cfg, "net_discount"),
                            'min_price': nd_cfg.get('min_price', 0.0),
                            'discount_amount': nd_cfg.get('discount_amount', 0.0),
                            'discount_type': nd_cfg.get('discount_type', '%'),
                            'min_basket_price': nd_cfg.get('min_basket_price', 0.0),
                            'order_limit': nd_cfg.get('order_limit', 0),
                            'trendyol_percent': 0.0,
                            'dict': item_dict
                        }
                except Exception: pass
            if net_discount_item is None:
                net_discount_item = {
                    'id': nd_cfg.get('id') or 'net_discount_1',
                    'label': build_campaign_label(nd_cfg, "net_discount"),
                    'min_price': nd_cfg.get('min_price', 0.0),
                    'discount_amount': nd_cfg.get('discount_amount', 0.0),
                    'discount_type': nd_cfg.get('discount_type', '%'),
                    'min_basket_price': nd_cfg.get('min_basket_price', 0.0),
                    'order_limit': nd_cfg.get('order_limit', 0),
                    'trendyol_percent': 0.0,
                    'dict': None
                }
    elif input_files.get('net_discount') and os.path.exists(input_files['net_discount']):
        try:
            nd_df = pd.read_excel(input_files['net_discount'])
            if not nd_df.empty and 'Barkod' in nd_df.columns:
                nd_df = nd_df.assign(BARKOD_CLN=nd_df['Barkod'].astype(str).str.strip())
                item_dict = nd_df.drop_duplicates(subset=['BARKOD_CLN']).set_index('BARKOD_CLN').to_dict('index')
                net_discount_item = {
                    'id': 'net_discount_1',
                    'label': 'Net İndirim',
                    'min_price': 0.0,
                    'discount_amount': 0.0,
                    'discount_type': '%',
                    'min_basket_price': 0.0,
                    'order_limit': 0,
                    'trendyol_percent': 0.0,
                    'dict': item_dict
                }
        except Exception: pass

    # Parse multi-counter files configuration if provided
    counter_items = []
    if counter_files and isinstance(counter_files, list):
        for idx, raw_item in enumerate(counter_files):
            item = normalize_campaign_config(raw_item, "counter")
            if item.get('enabled', True) is False:
                continue
            try:
                item_path = item.get('path') or item.get('stored_path')
                c_df = pd.read_excel(item_path) if isinstance(item_path, (str, Path)) else item.get('df', pd.DataFrame())
                if not c_df.empty and 'Barkod' in c_df.columns:
                    c_df = c_df.assign(BARKOD_CLN=c_df['Barkod'].astype(str).str.strip())
                    item_dict = c_df.drop_duplicates(subset=['BARKOD_CLN']).set_index('BARKOD_CLN').to_dict('index')
                    counter_items.append({
                        'id': item.get('id') or f"counter_{idx+1}",
                        'label': build_campaign_label(item, "counter", idx),
                        'min_price': item['min_price'],
                        'discount_amount': item['discount_amount'],
                        'discount_type': item['discount_type'],
                        'trendyol_percent': item['trendyol_percent'],
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
                    'discount_type': 'TL',
                    'trendyol_percent': tr_oran,
                    'dict': item_dict
                })
        except Exception: pass

    # Parse multi plus_extra files configuration if provided
    plus_extra_items = []
    if plus_extra_files and isinstance(plus_extra_files, list):
        for idx, raw_item in enumerate(plus_extra_files):
            item = normalize_campaign_config(raw_item, "plus_extra")
            if item.get('enabled', True) is False:
                continue
            try:
                item_path = item.get('path') or item.get('stored_path')
                pe_df = pd.read_excel(item_path) if isinstance(item_path, (str, Path)) else item.get('df', pd.DataFrame())
                if not pe_df.empty and 'Barkod' in pe_df.columns:
                    pe_df = pe_df.assign(BARKOD_CLN=pe_df['Barkod'].astype(str).str.strip())
                    item_dict = pe_df.drop_duplicates(subset=['BARKOD_CLN']).set_index('BARKOD_CLN').to_dict('index')
                    plus_extra_items.append({
                        'id': item.get('id') or f"plus_extra_{idx+1}",
                        'label': build_campaign_label(item, "plus_extra", idx),
                        'min_price': item['min_price'],
                        'discount_amount': item['discount_amount'],
                        'discount_type': item['discount_type'],
                        'trendyol_percent': item['trendyol_percent'],
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
                    config = normalize_campaign_config({'rate': rate}, "plus_extra")
                    plus_extra_items.append({
                        'id': f"plus_extra_{int(rate)}",
                        'label': build_campaign_label(config, "plus_extra"),
                        'min_price': config['min_price'],
                        'discount_amount': config['discount_amount'],
                        'discount_type': config['discount_type'],
                        'trendyol_percent': config['trendyol_percent'],
                        'dict': item_dict
                    })
        except Exception: pass

    # Parse multi coupon files configuration if provided
    coupon_items = []
    if coupon_files and isinstance(coupon_files, list):
        for idx, raw_item in enumerate(coupon_files):
            item = normalize_campaign_config(raw_item, "coupon")
            if item.get('enabled', True) is False:
                continue
            try:
                item_path = item.get('path') or item.get('stored_path')
                cp_df = pd.read_excel(item_path) if isinstance(item_path, (str, Path)) else item.get('df', pd.DataFrame())
                if not cp_df.empty and 'Barkod' in cp_df.columns:
                    cp_df = cp_df.assign(BARKOD_CLN=cp_df['Barkod'].astype(str).str.strip())
                    item_dict = cp_df.drop_duplicates(subset=['BARKOD_CLN']).set_index('BARKOD_CLN').to_dict('index')
                    coupon_items.append({
                        'id': item.get('id') or f"coupon_{idx+1}",
                        'label': build_campaign_label(item, "coupon", idx),
                        'min_price': item['min_price'],
                        'discount_amount': item['discount_amount'],
                        'discount_type': item['discount_type'],
                        'trendyol_percent': item['trendyol_percent'],
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

    plus_periods = find_plus_period_columns(df_plus.columns)

    df_tr['BARKOD_CLN'] = df_tr['BARKOD'].astype(str).str.strip()
    if 'Durum' in df_tr.columns:
        indirimli = df_tr[df_tr['Durum'].astype(str).str.contains('ndirim', case=False, na=False)]
    else:
        indirimli = df_tr
    tr_barcodes = indirimli['BARKOD_CLN'].unique()

    df_gun['BARKOD_CLN'] = df_gun['Barkod'].astype(str).str.strip()
    gun_barcodes = df_gun['BARKOD_CLN'].unique()
    kom_b_col = 'BARKOD' if 'BARKOD' in df_kom.columns else ('Barkod' if 'Barkod' in df_kom.columns else df_kom.columns[0])
    df_kom['BARKOD_CLN'] = df_kom[kom_b_col].astype(str).str.strip()
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

    def to_rows_safe(df, key_col):
        if df.empty or key_col not in df.columns:
            return {}
        return {
            str(key).strip(): group.to_dict('records')
            for key, group in df.groupby(key_col, sort=False)
        }

    dict_av = to_dict_safe(df_av, 'BARKOD_CLN')
    dict_kom = to_dict_safe(df_kom, 'BARKOD_CLN')
    dict_gun = to_dict_safe(df_gun, 'BARKOD_CLN')
    dict_ind = to_dict_safe(indirimli, 'BARKOD_CLN')
    dict_plus = to_dict_safe(df_plus, 'BARKOD_CLN')
    dict_plus_ek = to_dict_safe(df_plus_ek, 'BARKOD_CLN')
    dict_muh_av = to_dict_safe(df_muh_av, 'BARKOD_CLN')
    dict_muh_plus = to_dict_safe(df_muh_plus, 'BARKOD_CLN')
    flash_rows_by_barcode = to_rows_safe(df_flas, 'BARKOD_CLN')
    accounting_flash_rows_by_barcode = to_rows_safe(df_muh_flas, 'BARKOD_CLN')

    results = []

    for b in barcodes:
        fl_rows = flash_rows_by_barcode.get(b, [])
        muh_flas_rows = accounting_flash_rows_by_barcode.get(b, [])
        av_row = dict_av.get(b)
        fl_row = fl_rows[0] if fl_rows else None
        kom_row = dict_kom.get(b)
        gun_row = dict_gun.get(b)
        tr_row = dict_ind.get(b)
        plus_row = dict_plus.get(b)
        plus_ek_row = dict_plus_ek.get(b)

        muh_av_row = dict_muh_av.get(b)
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
        match_fl = bool(fl_rows or muh_flas_rows)
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
            for col_name in [
                'YENİ TSF (FİYAT GÜNCELLE)',
                '1 YILDIZ ÜST FİYAT',
                'TRENDYOL SATIŞ FİYATI',
            ]:
                if col_name in muh_av_row:
                    val = to_float(muh_av_row[col_name])
                    if val and val > 0: 
                        av_dip = val
                        av_dip_source = 'Avantajlı Muhasebe'
                        break

        campaign_flash_intervals = flash_interval_candidates(fl_rows, "kampanya")
        accounting_flash_intervals = flash_interval_candidates(muh_flas_rows, "muhasebe")

        flas_dip = None
        flas_dip_source = None
        flash_floor_prices = [
            interval["price"]
            for interval in accounting_flash_intervals
            if interval["source"] != "Mevcut Fiyat"
        ]
        if flash_floor_prices:
            flas_dip = min(flash_floor_prices)
            flas_dip_source = 'Flaş Muhasebe'

        plus_dip = None
        plus_dip_source = None
        if muh_plus_row is not None:
            for col_name in ['Plus Fiyat Seçimi', 'Plus Fiyat Üst Limiti']:
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

        all_dips = [d['price'] for d in dip_details]
        min_dip_val = min(all_dips) if all_dips else None
        common_threshold = min_dip_val

        dip_rate = get_commission_rate(min_dip_val, kom_row) if (min_dip_val and kom_row) else None
        if dip_rate is None and min_dip_val and gun_row:
            try: dip_rate = to_float(gun_row.get('Komisyon Oranı'))
            except Exception: pass
        dip_net = round(min_dip_val - (min_dip_val * (dip_rate / 100.0)), 2) if (min_dip_val and dip_rate is not None) else None
        # 0. Komisyon Tarifesi Optimizasyonu (4 Kademeli)
        kom_tarife_fiyat = None
        kom_tarife_oran = None
        kom_tarife_net = None
        kom_tarife_kademe = None
        kom_tarife_secimi = None
        kom_has_eligible = False
        if kom_row is not None:
            kom_eval = evaluate_commission_tariff_bracket(kom_row, min_dip_val, dip_net, guncel_fiyat_calc)
            if kom_eval is not None:
                kom_tarife_fiyat = kom_eval["price"]
                kom_tarife_oran = kom_eval["rate"]
                kom_tarife_net = kom_eval["net"]
                kom_tarife_kademe = kom_eval["kademe_adi"]
                kom_tarife_secimi = kom_eval["tariff_selection"]
                kom_has_eligible = kom_eval["has_eligible"]

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

                # Eğer Komisyon Tarifesinde daha avantajlı/karlı bir kademe varsa (ör. dip netinden veya mevcut netten daha iyi)
                if (
                    kom_has_eligible
                    and kom_tarife_fiyat is not None
                    and kom_tarife_net is not None
                    and kom_tarife_oran is not None
                ):
                    if (
                        (kom_tarife_fiyat <= yeni_tsf + 0.01 and (net_2 is None or kom_tarife_net >= net_2 - 0.01))
                        or (net_2 is not None and kom_tarife_net > net_2)
                    ):
                        yeni_tsf = kom_tarife_fiyat
                        rate_2 = kom_tarife_oran
                        net_2 = kom_tarife_net

                av_floor_check = (
                    kom_tarife_fiyat
                    if (kom_has_eligible and kom_tarife_fiyat is not None and (av_threshold is None or kom_tarife_fiyat < av_threshold))
                    else av_threshold
                )
                if (
                    (av_floor_check is not None or yeni_tsf_is_fallback)
                    and (av_floor_check is None or yeni_tsf >= av_floor_check - 0.01)
                ):
                    eligible_campaigns.append('Avantajlı')
                    if net_2 is not None:
                        smart_candidates.append(('Avantajlı', net_2, yeni_tsf, rate_2, yeni_tsf_is_fallback, muh_av_row is not None))

        # 2. Flaş Kampanya
        f_24_fiyat = None
        f_24_fiyat_is_fallback = False
        rate_3 = None
        net_3 = None
        flash_evaluations = []
        flas_has_eligible_interval = False
        if match_fl:
            flash_intervals = merge_flash_intervals(
                campaign_flash_intervals,
                accounting_flash_intervals,
            )
            if not flash_intervals and guncel_fiyat_calc and guncel_fiyat_calc > 0:
                flash_intervals = [{
                    "period": "24 Saat",
                    "start": None,
                    "end": None,
                    "price": guncel_fiyat_calc,
                    "source": "Mevcut Fiyat",
                    "origin": "güncel",
                    "used_current_price": True,
                }]

            for interval in flash_intervals:
                interval_rate = get_commission_rate(interval["price"], kom_row) if kom_row else None
                if interval_rate is None and gun_row:
                    interval_rate = to_float(gun_row.get('Komisyon Oranı'))
                interval_net = (
                    round(interval["price"] - (interval["price"] * (interval_rate / 100.0)), 2)
                    if interval_rate is not None
                    else None
                )
                interval_eligible = (
                    interval_rate is not None
                    and (flas_threshold is not None or interval["used_current_price"])
                    and (
                        flas_threshold is None
                        or interval["price"] >= flas_threshold - 0.01
                    )
                )
                flash_evaluations.append({
                    **interval,
                    "rate": interval_rate,
                    "net": interval_net,
                    "eligible": interval_eligible,
                })

            eligible_flash_evaluations = [
                evaluation for evaluation in flash_evaluations
                if evaluation["eligible"] and evaluation["net"] is not None
            ]
            valid_flash_evaluations = [
                evaluation for evaluation in flash_evaluations
                if evaluation["net"] is not None
            ]
            scalar_pool = eligible_flash_evaluations or valid_flash_evaluations
            if scalar_pool:
                selected_flash = min(scalar_pool, key=lambda evaluation: evaluation["net"])
                f_24_fiyat = selected_flash["price"]
                f_24_fiyat_is_fallback = selected_flash["used_current_price"]
                rate_3 = selected_flash["rate"]
                net_3 = selected_flash["net"]
            flas_has_eligible_interval = bool(eligible_flash_evaluations)
            if flas_has_eligible_interval:
                eligible_campaigns.append('Flaş')
                smart_candidates.append((
                    'Flaş',
                    net_3,
                    f_24_fiyat,
                    rate_3,
                    f_24_fiyat_is_fallback,
                    any(item["origin"] == "muhasebe" for item in eligible_flash_evaluations),
                ))

        # 3. Plus Kampanya
        plus_fiyat = None
        plus_fiyat_is_fallback = False
        rate_4 = None
        net_4 = None
        plus_tariff_label = None
        plus_audit_fields = {}
        if match_plus:
            plus_upper_limit = to_float(plus_row.get('Plus Fiyat Üst Limiti')) if plus_row else None
            if plus_upper_limit is None and muh_plus_row:
                plus_upper_limit = to_float(muh_plus_row.get('Plus Fiyat Üst Limiti'))
            for price_row in (muh_plus_row, plus_row):
                if price_row is not None:
                    value = to_float(price_row.get('Plus Fiyat Seçimi'))
                    if value is not None and value > 0:
                        plus_fiyat = value
                        break
            if plus_fiyat is None:
                for price_row in (muh_plus_row, plus_row):
                    if price_row is None:
                        continue
                    for col_name in ('Plus Fiyat Üst Limiti', 'Güncel TSF'):
                        value = to_float(price_row.get(col_name))
                        if value is not None and value > 0:
                            plus_fiyat = value
                            break
                    if plus_fiyat is not None:
                        break

            if plus_fiyat is None and guncel_fiyat_calc and guncel_fiyat_calc > 0:
                plus_fiyat = guncel_fiyat_calc
                plus_fiyat_is_fallback = True

            if plus_fiyat and plus_fiyat > 0:
                row_periods = plus_periods if plus_row is not None and plus_periods else [{
                    'days': 7,
                    'date_position': None,
                    'offer_position': None,
                    'date_column': None,
                    'offer_column': 'Plus Komisyon Teklifi' if plus_row and 'Plus Komisyon Teklifi' in plus_row else None,
                }]
                period_evaluations = []
                eligible_periods = []
                price_passes_floor = (
                    (plus_threshold is not None or plus_fiyat_is_fallback)
                    and (plus_threshold is None or plus_fiyat >= plus_threshold - 0.01)
                )
                price_within_upper_limit = (
                    plus_upper_limit is None or plus_fiyat <= plus_upper_limit + 0.01
                )
                for period in row_periods:
                    period_rate = (
                        to_float(plus_row.get(period['offer_column']))
                        if plus_row is not None and period['offer_column'] is not None
                        else None
                    )
                    if period['date_column'] is None and period_rate is None and kom_row:
                        period_rate = get_commission_rate(plus_fiyat, kom_row)
                    if period['date_column'] is None and period_rate is None and gun_row:
                        period_rate = to_float(gun_row.get('Komisyon Oranı'))
                    if period_rate is not None and not 0 <= period_rate <= 100:
                        period_rate = None
                    period_net = (
                        round(plus_fiyat - (plus_fiyat * (period_rate / 100.0)), 2)
                        if period_rate is not None
                        else None
                    )
                    period_evaluations.append((period, period_rate, period_net))
                    plus_audit_fields[f"Plus Komisyon ({period['days']} Gün) (%)"] = period_rate
                    plus_audit_fields[f"Plus Net ({period['days']} Gün) (TL)"] = period_net

                    date_value = plus_row.get(period['date_column']) if plus_row is not None and period['date_column'] is not None else True
                    date_exists = date_value is True or (
                        date_value is not None
                        and not pd.isna(date_value)
                        and str(date_value).strip() != ''
                    )
                    if (
                        date_exists
                        and period_rate is not None
                        and price_passes_floor
                        and price_within_upper_limit
                        and fallback_candidate_is_selectable(net_1, period_net, plus_fiyat_is_fallback)
                    ):
                        eligible_periods.append((period, period_rate, period_net))

                if len(period_evaluations) == 1:
                    rate_4, net_4 = period_evaluations[0][1:]
                elif eligible_periods:
                    rate_4 = max(period[1] for period in eligible_periods)
                    net_4 = min(period[2] for period in eligible_periods)

                plus_tariff_label = choose_plus_tariff_label(
                    plus_row or {},
                    row_periods,
                    [period[0]['days'] for period in eligible_periods],
                )
                if plus_tariff_label is not None:
                    eligible_campaigns.append('Plus')
                    if net_4 is not None:
                        smart_candidates.append(('Plus', net_4, plus_fiyat, rate_4, plus_fiyat_is_fallback, muh_plus_row is not None))

        # 3.5 Komisyon Tarifesi Kampanyası (4 Kademeli Optimizasyon)
        if kom_has_eligible:
            eligible_campaigns.append("Komisyon Tarifesi")
            if kom_tarife_net is not None:
                smart_candidates.append(("Komisyon Tarifesi", kom_tarife_net, kom_tarife_fiyat, kom_tarife_oran, False, False))

        counter_evaluations = {}
        qualified_extra_labels = set()

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
                    
                    if pe_price and pe_price >= pe_item['min_price'] - 0.01:
                        pe_rate = get_commission_rate(pe_price, kom_row) if kom_row else None
                        if pe_rate is None and gun_row:
                            try: pe_rate = to_float(gun_row['Komisyon Oranı'])
                            except: pass
                        evaluation = build_extra_evaluation(pe_price, pe_rate, pe_item)
                        if (
                            evaluation
                            and evaluation['customer_price'] >= 0
                            and (
                                common_threshold is None
                                or evaluation['customer_price'] >= common_threshold - 0.01
                            )
                        ):
                            label = pe_item['label']
                            qualified_extra_labels.add(label)
                            counter_evaluations[label] = evaluation
                            eligible_campaigns.append(label)
                            smart_candidates.append((
                                label,
                                evaluation['net'],
                                evaluation['customer_price'],
                                evaluation['rate'],
                                pe_price_is_fallback,
                            ))

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
                
                if c_price and c_price >= c_item['min_price'] - 0.01:
                    c_rate = get_commission_rate(c_price, kom_row) if kom_row else None
                    if c_rate is None and gun_row:
                        try: c_rate = to_float(gun_row['Komisyon Oranı'])
                        except: pass
                    evaluation = build_extra_evaluation(c_price, c_rate, c_item)
                    if (
                        evaluation
                        and evaluation['customer_price'] >= 0
                        and (
                            common_threshold is None
                            or evaluation['customer_price'] >= common_threshold - 0.01
                        )
                    ):
                        label = c_item['label']
                        qualified_extra_labels.add(label)
                        counter_evaluations[label] = evaluation
                        eligible_campaigns.append(label)
                        smart_candidates.append((
                            label,
                            evaluation['net'],
                            evaluation['customer_price'],
                            evaluation['rate'],
                            c_price_is_fallback,
                        ))

        # 6. Kupon Kampanyaları (Çoklu)
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
                    
                    if cp_price and cp_price >= cp_item['min_price'] - 0.01:
                        cp_rate = get_commission_rate(cp_price, kom_row) if kom_row else None
                        if cp_rate is None and gun_row:
                            try: cp_rate = to_float(gun_row['Komisyon Oranı'])
                            except: pass
                        evaluation = build_extra_evaluation(cp_price, cp_rate, cp_item)
                        if (
                            evaluation
                            and evaluation['customer_price'] >= 0
                            and (
                                common_threshold is None
                                or evaluation['customer_price'] >= common_threshold - 0.01
                            )
                        ):
                            label = cp_item['label']
                            qualified_extra_labels.add(label)
                            counter_evaluations[label] = evaluation
                            eligible_campaigns.append(label)
                            smart_candidates.append((
                                label,
                                evaluation['net'],
                                evaluation['customer_price'],
                                evaluation['rate'],
                                cp_price_is_fallback,
                            ))

        # 7. Net İndirim Kampanyası
        if net_discount_item:
            nd_dict = net_discount_item['dict']
            nd_row = nd_dict.get(b) if nd_dict is not None else {}
            if nd_row is not None:
                nd_price = None
                nd_price_is_fallback = False
                for nd_col in ['Güncel Satış Fiyatı', 'Mevcut Satış Fiyatı', 'Maksimum Girebileceğin Fiyat', 'Kampanyalı Satış Fiyatı']:
                    if nd_col in nd_row:
                        val = to_float(nd_row[nd_col])
                        if val and val > 0: nd_price = val; break
                if nd_price is None and guncel_fiyat_calc and guncel_fiyat_calc > 0:
                    nd_price = guncel_fiyat_calc
                    nd_price_is_fallback = False

                if nd_price and nd_price >= net_discount_item['min_price'] - 0.01:
                    nd_rate = get_commission_rate(nd_price, kom_row) if kom_row else None
                    if nd_rate is None and gun_row:
                        try: nd_rate = to_float(gun_row['Komisyon Oranı'])
                        except: pass
                    evaluation = build_extra_evaluation(nd_price, nd_rate, net_discount_item)
                    if (
                        evaluation
                        and evaluation['customer_price'] >= 0
                        and (
                            common_threshold is None
                            or evaluation['customer_price'] >= common_threshold - 0.01
                        )
                    ):
                        label = net_discount_item['label']
                        qualified_extra_labels.add(label)
                        counter_evaluations[label] = evaluation
                        eligible_campaigns.append(label)
                        smart_candidates.append((
                            label,
                            evaluation['net'],
                            evaluation['customer_price'],
                            evaluation['rate'],
                            nd_price_is_fallback,
                        ))

        selectable = selectable_campaigns(net_1, smart_candidates)

        all_matching_main_campaigns = ['Hiçbiri']
        if (
            match_av
            and yeni_tsf
            and (av_threshold is not None or yeni_tsf_is_fallback)
            and fallback_candidate_is_selectable(net_1, net_2, yeni_tsf_is_fallback)
        ):
            if 'Avantajlı' not in all_matching_main_campaigns: all_matching_main_campaigns.append('Avantajlı')
        if (
            match_fl
            and f_24_fiyat
            and flas_has_eligible_interval
            and fallback_candidate_is_selectable(net_1, net_3, f_24_fiyat_is_fallback)
        ):
            if 'Flaş' not in all_matching_main_campaigns: all_matching_main_campaigns.append('Flaş')
        if (
            match_plus
            and plus_fiyat
            and plus_tariff_label is not None
        ):
            if 'Plus' not in all_matching_main_campaigns: all_matching_main_campaigns.append('Plus')
        if kom_row is not None and kom_tarife_fiyat and kom_has_eligible:
            if 'Komisyon Tarifesi' not in all_matching_main_campaigns: all_matching_main_campaigns.append('Komisyon Tarifesi')

        eligible_main_campaigns = ['Hiçbiri'] + [c[0] for c in selectable if c[0] in MAIN_CAMPAIGN_KEYS]

        all_matching_extra_campaigns = ['Hiçbiri']
        if match_plus_ek and plus_ek_fiyat and fallback_candidate_is_selectable(net_1, net_5, plus_ek_fiyat_is_fallback):
            for c_name in ('Plus Ek İndirim %5', 'Plus Ek İndirim %10', 'Plus Ek İndirim %20'):
                if c_name not in all_matching_extra_campaigns:
                    all_matching_extra_campaigns.append(c_name)
        if plus_extra_items:
            for pe_item in plus_extra_items:
                if pe_item['label'] in qualified_extra_labels:
                    if pe_item['label'] not in all_matching_extra_campaigns:
                        all_matching_extra_campaigns.append(pe_item['label'])
        if coupon_items:
            for cp_item in coupon_items:
                if cp_item['label'] in qualified_extra_labels:
                    if cp_item['label'] not in all_matching_extra_campaigns:
                        all_matching_extra_campaigns.append(cp_item['label'])
        if net_discount_item and net_discount_item['label'] in qualified_extra_labels:
            if net_discount_item['label'] not in all_matching_extra_campaigns:
                all_matching_extra_campaigns.append(net_discount_item['label'])
        for c_item in counter_items:
            if c_item['label'] in qualified_extra_labels:
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

        rec_res = choose_campaigns_smart(net_1, selectable, recommendation_rule)
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

        effective_dip_fiyat = min([d for d in [min_dip_val, kom_tarife_fiyat] if d is not None]) if (kom_has_eligible and kom_tarife_fiyat is not None) else min_dip_val
        if kom_has_eligible and kom_tarife_fiyat is not None and (min_dip_val is None or kom_tarife_fiyat < min_dip_val):
            dip_details.append({'type': 'Komisyon Tarifesi', 'price': kom_tarife_fiyat})

        guncel_fiyat_display = guncel_fiyat_calc
        discount_fields = build_discount_fields(
            is_indirim or (len(dip_details) > 0),
            guncel_fiyat_display,
            effective_dip_fiyat,
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
            'Komisyon Tarifesi Eşleşme Durumu': 'Eşleşti' if kom_row is not None else 'Eşleşme Yok',
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
            'flash_evaluations': flash_evaluations,
            'Plus Fiyatı (TL)': plus_fiyat,
            'Plus Tarife Seçimi': plus_tariff_label,
            'Plus Komisyon (%)': rate_4,
            'Plus Net (TL)': net_4,
            **plus_audit_fields,
            'Komisyon Tarifesi Fiyatı (TL)': kom_tarife_fiyat,
            'Komisyon Tarifesi Komisyon (%)': kom_tarife_oran,
            'Komisyon Tarifesi Net (TL)': kom_tarife_net,
            'Komisyon Tarifesi Kademe': kom_tarife_kademe,
            'Komisyon Tarifesi Seçimi': kom_tarife_secimi,
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
            'Düşülebilecek Dip Fiyat (TL)': min([d for d in [min_dip_val, kom_tarife_fiyat] if d is not None]) if (kom_has_eligible and kom_tarife_fiyat is not None) else min_dip_val,
            'campaign_floor_prices': {
                'Avantajlı': min([d for d in [av_threshold, yeni_tsf] if d is not None]) if (av_threshold is not None or yeni_tsf is not None) else min_dip_val,
                'Flaş': flas_threshold if flas_threshold is not None else min_dip_val,
                'Plus': plus_threshold if plus_threshold is not None else min_dip_val,
                'Komisyon Tarifesi': kom_tarife_fiyat if (kom_has_eligible and kom_tarife_fiyat is not None) else min_dip_val,
            },
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
        out_df.to_excel(out_file, index=False)
    except Exception as e:
        print("Excel kaydetme uyarısı:", e)

    return {"success": True, "output_path": out_file, "results": results, "counter_items": counter_items}
