import math

def clean_nans(obj):
    """Recursively replaces NaN, Infinity, -Infinity values with None for standard JSON serialization."""
    if isinstance(obj, float):
        if math.isnan(obj) or math.isinf(obj):
            return None
        return obj
    elif isinstance(obj, dict):
        return {k: clean_nans(v) for k, v in obj.items()}
    elif isinstance(obj, list):
        return [clean_nans(v) for v in obj]
    try:
        if pd.isna(obj):
            return None
    except Exception:
        pass
    return obj

import ast
import os
import json
import math
import re
from io import BytesIO
import openpyxl
import pandas as pd
from datetime import datetime, timedelta
from flask import Flask, render_template, jsonify, request, send_from_directory
from input_files import (
    INPUT_SPECS,
    InputValidationError,
    load_upload_set,
    load_upload_status,
    save_upload_set,
    load_counter_configs,
    load_plus_extra_configs,
    load_recommendation_rule,
    save_single_file_expiries,
    load_user_selections,
    build_campaign_label,
    normalize_campaign_config,
    normalize_campaign_configs,
    normalize_recommendation_rule,
    save_recommendation_rule,
    save_user_selections,
)
from xlsx_postprocess import fix_xlsx_for_trendyol

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 240 * 1024 * 1024

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
INPUT_DIR = os.path.join(BASE_DIR, "Girdiler")
OUTPUT_DIR = os.path.join(BASE_DIR, "Çıktılar")
UPLOAD_DIR = os.path.join(INPUT_DIR, "Yuklenen")
INPUT_MANIFEST = os.path.join(INPUT_DIR, "yuklenen_girdiler.json")

os.makedirs(INPUT_DIR, exist_ok=True)
os.makedirs(OUTPUT_DIR, exist_ok=True)

F_HESAP = os.path.join(OUTPUT_DIR, "Kampanya_Hesaplama_Sonuclari.xlsx")

REPORT_COLUMNS = [
    "Barkod",
    "Güncel Fiyat (TL)",
    "Güncel Net",
    "Güncel Komisyon",
    "Avantajlı Fiyat (TL)",
    "Avantajlı Net",
    "Flaş Fiyat (TL)",
    "Flaş Net",
    "Plus Fiyat (TL)",
    "Plus Net",
    "Plus Ek İndirim Fiyat (TL)",
    "Plus Ek İndirim Net",
    "Karşılamalı Kampanya Fiyat (TL)",
    "Karşılamalı Kampanya Net",
    "Uygulanan Kampanya",
    "Ekstra Kampanya",
    "Hangisi Karlı?",
    "Düşülebilecek Dip Fiyat (TL)",
    "Uygulanan Kampanya Fiyat",
    "Uygulanan Kampanya Net",
    "Uygulanan Kampanya Komisyon",
    "Uygulanabilecek İndirim (TL)",
    "Uygulanabilecek İndirim (%)",
    "Uygulanan İndirim (TL)",
    "Uygulanan İndirim (%)",
    "Ekstra Uygulanabilir İndirim (TL)",
    "Ekstra Uygulanabilir İndirim (%)",
]
ROUNDING_EPSILON = 1e-9

CAMPAIGN_LABELS = {
    "Avantajlı": "Avantajlı Ürün",
    "Flaş": "Flaş Ürün",
    "Plus": "Plus Ürün",
}
VALID_TARGET_TYPES = {
    "Hepsi",
    "Avantajlı",
    "Flaş",
    "Plus",
    "Plus Ek İndirim",
    "Karşılamalı Kampanya",
}
MAIN_SELECTIONS = {"Hiçbiri", "Avantajlı", "Flaş", "Plus"}


def as_number(value):
    try:
        number = float(value)
    except (TypeError, ValueError):
        return None
    return number if math.isfinite(number) else None


def parse_campaign_configs_json(raw_value, campaign_type):
    if not raw_value:
        return []
    try:
        configs = json.loads(raw_value)
    except (TypeError, ValueError) as exc:
        raise InputValidationError("Kampanya yapılandırması okunamadı.") from exc
    return normalize_campaign_configs(configs, campaign_type)


def parse_recommendation_rule_json(raw_value):
    try:
        rule = json.loads(raw_value)
    except (TypeError, ValueError) as exc:
        raise InputValidationError("Öneri kuralı okunamadı.") from exc
    if rule is None:
        raise InputValidationError("Öneri kuralı boş olamaz.")
    return normalize_recommendation_rule(rule)


def parse_persisted_collection(value, expected_type):
    if isinstance(value, expected_type):
        return value
    if isinstance(value, str):
        try:
            value = ast.literal_eval(value)
        except (SyntaxError, ValueError):
            pass
    return value if isinstance(value, expected_type) else expected_type()


def restore_persisted_collections(frame):
    return frame.assign(**{
        column: frame[column].map(
            lambda value, kind=expected_type: parse_persisted_collection(value, kind)
        )
        for column, expected_type in (
            ("eligible_main_campaigns", list),
            ("all_matching_main_campaigns", list),
            ("eligible_extra_campaigns", list),
            ("all_matching_extra_campaigns", list),
            ("eligible_campaigns", list),
            ("all_matching_campaigns", list),
            ("counter_evaluations", dict),
            ("dip_details", list),
        )
        if column in frame.columns
    })


def normalize_selection(selection):
    if isinstance(selection, str):
        value = selection.strip()
        if not value:
            return None
        return (value, "Hiçbiri") if value in MAIN_SELECTIONS else ("Hiçbiri", value)
    if not isinstance(selection, dict) or set(selection) - {"main", "extra"}:
        return None
    main = selection.get("main", "Hiçbiri")
    extra = selection.get("extra", "Hiçbiri")
    if (
        not isinstance(main, str)
        or not isinstance(extra, str)
        or main not in MAIN_SELECTIONS
        or not extra.strip()
    ):
        return None
    return main, extra.strip()


def selection_payload_is_valid(selections):
    return isinstance(selections, dict) and all(
        isinstance(barcode, str)
        and bool(barcode.strip())
        and normalize_selection(selection) is not None
        for barcode, selection in selections.items()
    )


def round2(value):
    return math.floor((float(value) * 100) + 0.5 + ROUNDING_EPSILON) / 100


def discounted_price(value, rate):
    price = as_number(value)
    return round2(price * (1 - rate / 100)) if price is not None else None


def plus_extra_export_filename(config, index=0):
    config = normalize_campaign_config(config, "plus_extra")
    format_number = lambda value: (
        str(int(value)) if float(value).is_integer() else str(float(value))
    )
    discount = format_number(config["discount_amount"])
    if (
        config["discount_type"] == "%"
        and config["min_price"] == 0
        and config["trendyol_percent"] == 0
    ):
        return f"Trendyol_Plus_Musterilerine_Ozel_Ek_%{discount}_Indirim.xlsx"

    min_price = format_number(config["min_price"])
    trendyol = format_number(config["trendyol_percent"])
    discount_part = (
        f"%{discount}" if config["discount_type"] == "%" else f"{discount}_TL"
    )
    return (
        "Trendyol_Plus_Musterilerine_Ozel_Ek_"
        f"{min_price}_TL_Uzeri_{discount_part}_Indirim_"
        f"%{trendyol}_Trendyol_Karsilamali.xlsx"
    )


def discount_between(current_price, campaign_price):
    current = as_number(current_price)
    campaign = as_number(campaign_price)
    if current is None or campaign is None or current <= 0 or campaign <= 0:
        return None, None
    amount = round2(max(current - campaign, 0))
    return amount, round2((amount / current) * 100)


def selected_campaign_values(row):
    main_sel = row.get("userSelection", "Hiçbiri") or "Hiçbiri"
    extra_sel = row.get("userExtraSelection", "Hiçbiri") or "Hiçbiri"

    if main_sel == "Avantajlı":
        base_price = as_number(row.get("Avantajlı Ürün Fiyatı (YENİ TSF) (TL)"))
        base_net = as_number(row.get("Avantajlı Ürün Kalan Net (TL)"))
        base_comm = as_number(row.get("Avantajlı Ürün Komisyon (%)"))
    elif main_sel == "Flaş":
        base_price = as_number(row.get("Flaş Ürün 24 Saat Fiyatı (TL)"))
        base_net = as_number(row.get("Flaş Ürün Kalan Net (TL)"))
        base_comm = as_number(row.get("Flaş Ürün Komisyon (%)"))
    elif main_sel == "Plus":
        base_price = as_number(row.get("Plus Fiyatı (TL)"))
        base_net = as_number(row.get("Plus Net (TL)"))
        base_comm = as_number(row.get("Plus Komisyon (%)"))
    else:
        base_price = as_number(row.get("Güncel Ürün Fiyatı (TL)"))
        base_net = as_number(row.get("Güncel Ürün Kalan Net (TL)"))
        base_comm = as_number(row.get("Güncel Ürün Komisyon (%)"))

    if base_price is None:
        base_price = as_number(row.get("Güncel Ürün Fiyatı (TL)"))
    if base_comm is None:
        base_comm = as_number(row.get("Güncel Ürün Komisyon (%)"))

    if extra_sel == "Hiçbiri" or not extra_sel:
        return base_price, base_net, base_comm

    counter_evals = row.get("counter_evaluations", {})
    if isinstance(counter_evals, dict) and extra_sel in counter_evals:
        c_info = counter_evals[extra_sel]
        c_price = as_number(c_info.get("price"))
        if c_price is None:
            c_price = base_price
        disc_type = c_info.get("disc_type", "%")
        disc_val = as_number(c_info.get("disc_val", 0))
        trendyol_percent = as_number(c_info.get("trendyol_percent", 0))
        disc_val = 0 if disc_val is None else disc_val
        trendyol_percent = 0 if trendyol_percent is None else trendyol_percent

        if (
            main_sel != "Hiçbiri"
            and base_price is not None
            and c_price is not None
            and base_price < c_price
        ):
            if disc_type == "%":
                tot_disc = round2(base_price * (disc_val / 100.0))
            else:
                tot_disc = disc_val
            seller_disc = round2(tot_disc * (1.0 - (trendyol_percent / 100.0)))
            final_price = round2(base_price - tot_disc)
            final_comm = 0 if base_comm is None else base_comm
            final_net = round2(
                base_price - (base_price * (final_comm / 100.0)) - seller_disc
            )
            return final_price, final_net, final_comm

        final_price = as_number(c_info.get("customer_price"))
        if final_price is None and c_price is not None:
            total_discount = (
                round2(c_price * (disc_val / 100.0))
                if disc_type == "%"
                else disc_val
            )
            final_price = round2(c_price - total_discount)
        final_comm = as_number(c_info.get("rate"))
        if final_comm is None:
            final_comm = 0 if base_comm is None else base_comm
        final_net = as_number(c_info.get("net"))
        if final_net is None and c_price is not None:
            seller_disc = as_number(c_info.get("seller_disc", 0)) or 0
            final_net = round2(
                c_price - (c_price * (final_comm / 100.0)) - seller_disc
            )
        return final_price, final_net, final_comm

    if extra_sel.startswith("Plus Ek İndirim %"):
        try:
            rate = float(extra_sel.rsplit("%", 1)[-1])
            rate_key = str(int(rate)) if rate.is_integer() else str(rate)
            stored_price = as_number(row.get(f"Plus Ek Fiyatı %{rate_key} (TL)"))
            stored_net = as_number(row.get(f"Plus Ek Net %{rate_key} (TL)"))
            final_comm = as_number(row.get("Plus Ek Komisyon (%)"))
            if final_comm is None:
                final_comm = 0 if base_comm is None else base_comm
            if stored_price is not None and stored_net is not None:
                return stored_price, stored_net, final_comm
            total_discount = round2(base_price * (rate / 100.0))
            final_price = round2(base_price - total_discount)
            final_net = round2(
                base_price - (base_price * (final_comm / 100.0)) - total_discount
            )
            return final_price, final_net, final_comm
        except Exception:
            pass

    return base_price, base_net, base_comm


def build_report_row(row):
    main_sel = row.get("userSelection", "Hiçbiri") or "Hiçbiri"
    extra_sel = row.get("userExtraSelection", "Hiçbiri") or "Hiçbiri"
    current_price = as_number(row.get("Güncel Ürün Fiyatı (TL)"))
    campaign_price, campaign_net, campaign_commission = selected_campaign_values(row)
    applied_amount, applied_percent = discount_between(current_price, campaign_price)
    
    dip_price = as_number(row.get("Düşülebilecek Dip Fiyat (TL)"))
    available_amount, available_percent = discount_between(current_price, dip_price)

    extra_amount = None
    extra_percent = None
    if available_amount is not None and applied_amount is not None:
        extra_amount = round2(max(available_amount - applied_amount, 0))
        extra_percent = round2((extra_amount / current_price) * 100) if current_price else None

    plus_extra_price = None
    plus_extra_net = None
    counter_evals = row.get("counter_evaluations", {})
    selected_evaluation = (
        counter_evals.get(extra_sel)
        if isinstance(counter_evals, dict)
        else None
    )
    if extra_sel.startswith("Plus Ek İndirim") and isinstance(selected_evaluation, dict):
        plus_extra_price = campaign_price
        plus_extra_net = campaign_net
    elif extra_sel.startswith("Plus Ek İndirim %"):
        try:
            plus_rate = int(extra_sel.rsplit("%", 1)[-1])
            plus_extra_price = as_number(row.get(f"Plus Ek Fiyatı %{plus_rate} (TL)"))
            plus_extra_net = as_number(row.get(f"Plus Ek Net %{plus_rate} (TL)"))
        except Exception: pass

    counter_price = None
    counter_net = None
    if isinstance(counter_evals, dict) and extra_sel in counter_evals:
        counter_price = campaign_price
        counter_net = campaign_net

    return {
        "Barkod": row.get("Barkod"),
        "Güncel Fiyat (TL)": current_price,
        "Güncel Net": as_number(row.get("Güncel Ürün Kalan Net (TL)")),
        "Güncel Komisyon": as_number(row.get("Güncel Ürün Komisyon (%)")),
        "Avantajlı Fiyat (TL)": as_number(row.get("Avantajlı Ürün Fiyatı (YENİ TSF) (TL)")),
        "Avantajlı Net": as_number(row.get("Avantajlı Ürün Kalan Net (TL)")),
        "Flaş Fiyat (TL)": as_number(row.get("Flaş Ürün 24 Saat Fiyatı (TL)")),
        "Flaş Net": as_number(row.get("Flaş Ürün Kalan Net (TL)")),
        "Plus Fiyat (TL)": as_number(row.get("Plus Fiyatı (TL)")),
        "Plus Net": as_number(row.get("Plus Net (TL)")),
        "Plus Ek İndirim Fiyat (TL)": plus_extra_price,
        "Plus Ek İndirim Net": plus_extra_net,
        "Karşılamalı Kampanya Fiyat (TL)": counter_price,
        "Karşılamalı Kampanya Net": counter_net,
        "Uygulanan Kampanya": CAMPAIGN_LABELS.get(main_sel, main_sel),
        "Ekstra Kampanya": CAMPAIGN_LABELS.get(extra_sel, extra_sel),
        "Hangisi Karlı?": row.get("Hangisi Daha Karlı?"),
        "Düşülebilecek Dip Fiyat (TL)": dip_price,
        "Uygulanan Kampanya Fiyat": campaign_price,
        "Uygulanan Kampanya Net": campaign_net,
        "Uygulanan Kampanya Komisyon": campaign_commission,
        "Uygulanabilecek İndirim (TL)": available_amount,
        "Uygulanabilecek İndirim (%)": available_percent,
        "Uygulanan İndirim (TL)": applied_amount,
        "Uygulanan İndirim (%)": applied_percent,
        "Ekstra Uygulanabilir İndirim (TL)": extra_amount,
        "Ekstra Uygulanabilir İndirim (%)": extra_percent,
    }


def normalize_visible_columns(requested):
    if requested is None:
        return REPORT_COLUMNS.copy()
    if not isinstance(requested, list):
        return REPORT_COLUMNS.copy()
    requested_set = {column for column in requested if isinstance(column, str)}
    return [column for column in REPORT_COLUMNS if column in requested_set]


def campaign_selection_is_applicable(selection, applicable_value):
    if selection == "Hiçbiri":
        return True
    if not isinstance(selection, str):
        return False
    applicable = {
        item.strip()
        for item in str(applicable_value or "").split(",")
        if item.strip()
    }
    if selection in applicable:
        return True
    if selection.startswith("Plus Ek İndirim"):
        campaign = "Plus Ek İndirim"
    elif selection.startswith("Karşılamalı"):
        campaign = "Karşılamalı Kampanya"
    elif "Kupon" in selection:
        campaign = "Kupon"
    else:
        campaign = selection
    return campaign in applicable


def row_selection_is_applicable(row, main_selection, extra_selection):
    main_options = row.get("eligible_main_campaigns")
    extra_options = row.get("eligible_extra_campaigns")
    main_ok = (
        main_selection in main_options
        if isinstance(main_options, list) and main_options
        else campaign_selection_is_applicable(
            main_selection, row.get("Uygulanabilir Kampanyalar")
        )
    )
    extra_ok = (
        extra_selection in extra_options
        if isinstance(extra_options, list) and extra_options
        else campaign_selection_is_applicable(
            extra_selection, row.get("Uygulanabilir Kampanyalar")
        )
    )
    return main_ok and extra_ok


def calculation_result_is_current():
    try:
        return os.path.isfile(F_HESAP) and os.path.getsize(F_HESAP) > 0
    except OSError:
        return False


def processing_error(label):
    app.logger.exception("%s işlenirken hata", label)
    return jsonify({
        "success": False,
        "message": f"{label} işlenemedi; girdiyi kontrol edip yeniden deneyin.",
    }), 500


def build_report_dataframe(table_data, requested_columns):
    rows = [build_report_row(row) for row in table_data if row.get("Barkod")]
    columns = normalize_visible_columns(requested_columns)
    return pd.DataFrame(rows, columns=REPORT_COLUMNS)[columns]


def write_report_excel(dataframe, output_path):
    dataframe.to_excel(output_path, index=False)
    workbook = openpyxl.load_workbook(output_path)
    sheet = workbook.active
    sheet.freeze_panes = "A2"
    sheet.auto_filter.ref = sheet.dimensions
    for index, column in enumerate(dataframe.columns, 1):
        letter = openpyxl.utils.get_column_letter(index)
        sheet.column_dimensions[letter].width = max(len(str(column)) + 2, 12)
    workbook.save(output_path)


@app.route("/api/download/<folder>/<filename>")
def download_file(folder, filename):
    if not re.fullmatch(r"\d{4}-\d{2}-\d{2}_\d{2}-\d{2}-\d{2}", folder):
        return jsonify({"success": False, "message": "Çıktı klasörü geçersiz."}), 404
    folder_dir = os.path.join(OUTPUT_DIR, folder)
    return send_from_directory(folder_dir, filename, as_attachment=True)

def safe_keep_rows(ws, keep_row_indices):
    original_max_row = ws.max_row
    keep_rows = sorted({
        row for row in keep_row_indices
        if isinstance(row, int) and 2 <= row <= original_max_row
    })
    if not keep_rows:
        if ws.max_row > 1:
            ws.delete_rows(2, ws.max_row - 1)
        return

    max_column = ws.max_column
    last_column = openpyxl.utils.get_column_letter(max_column)
    for target_row, source_row in enumerate(keep_rows, 2):
        if target_row != source_row:
            ws.move_range(
                f"A{source_row}:{last_column}{source_row}",
                rows=target_row - source_row,
            )

    first_unused_row = len(keep_rows) + 2
    if first_unused_row <= original_max_row:
        ws.delete_rows(first_unused_row, original_max_row - first_unused_row + 1)

    for r in range(2, ws.max_row + 1):
        for c in range(1, max_column + 1):
            cell = ws.cell(r, c)
            val = cell.value
            if isinstance(val, str) and (val.startswith('=') or val.startswith('==')):
                new_val = re.sub(r'([a-zA-Z]+)\d+\b', r'\g<1>' + str(r), val)
                cell.value = new_val


def header_index(ws, name):
    wanted = str(name).strip().casefold()
    for column in range(1, ws.max_column + 1):
        if str(ws.cell(1, column).value or "").strip().casefold() == wanted:
            return column
    return None


def ensure_header(ws, name):
    existing = header_index(ws, name)
    if existing:
        return existing
    column = ws.max_column + 1
    ws.cell(1, column).value = name
    return column


def load_merged_campaign_workbook(standard_path, accounting_path, barcode_header):
    paths = [path for path in (standard_path, accounting_path) if path and os.path.exists(path)]
    if not paths:
        return None

    workbook = openpyxl.load_workbook(paths[0])
    target = workbook.active
    target_barcode_column = header_index(target, barcode_header)
    if not target_barcode_column:
        return workbook

    existing_barcodes = {
        str(target.cell(row, target_barcode_column).value or "").strip()
        for row in range(2, target.max_row + 1)
    }
    target_columns = {
        str(target.cell(1, column).value or "").strip().casefold(): column
        for column in range(1, target.max_column + 1)
    }
    for source_path in paths[1:]:
        source = openpyxl.load_workbook(source_path, data_only=False).active
        source_barcode_column = header_index(source, barcode_header)
        if not source_barcode_column:
            continue
        source_columns = {
            str(source.cell(1, column).value or "").strip().casefold(): column
            for column in range(1, source.max_column + 1)
        }
        shared_columns = {
            target_column: source_columns[name]
            for name, target_column in target_columns.items()
            if name in source_columns
        }
        for row in range(2, source.max_row + 1):
            barcode = str(source.cell(row, source_barcode_column).value or "").strip()
            if not barcode or barcode in existing_barcodes:
                continue
            target_row = target.max_row + 1
            for target_column, source_column in shared_columns.items():
                target.cell(target_row, target_column).value = source.cell(
                    row, source_column
                ).value
            target.cell(target_row, target_barcode_column).value = barcode
            existing_barcodes.add(barcode)
    return workbook


def clone_workbook(workbook):
    buffer = BytesIO()
    workbook.save(buffer)
    buffer.seek(0)
    return openpyxl.load_workbook(buffer)


def shrink_data_validations(ws):
    max_row = ws.max_row
    if max_row < 2:
        max_row = 2
    try:
        if hasattr(ws, 'data_validations') and hasattr(ws.data_validations, 'dataValidation'):
            for dv in ws.data_validations.dataValidation:
                if not dv.sqref:
                    continue
                ranges = str(dv.sqref).split()
                new_ranges = []
                for r in ranges:
                    if ':' in r:
                        start, end = r.split(':')
                        import re
                        match = re.match(r"([a-zA-Z]+)(\d+)", start)
                        if match:
                            col = match.group(1)
                            new_ranges.append(f"{start}:{col}{max_row}")
                        else:
                            new_ranges.append(r)
                    else:
                        new_ranges.append(r)
                dv.sqref = ' '.join(new_ranges)
    except Exception as e:
        print("Data validation shrink error:", e)

@app.route("/")
def index():
    from input_files import load_coupon_configs
    return render_template(
        "index.html",
        report_columns=REPORT_COLUMNS,
        input_specs=INPUT_SPECS,
        uploaded_inputs=load_upload_status(UPLOAD_DIR, INPUT_MANIFEST),
        counter_configs=load_counter_configs(INPUT_MANIFEST),
        plus_extra_configs=load_plus_extra_configs(INPUT_MANIFEST),
        coupon_configs=load_coupon_configs(INPUT_MANIFEST),
        recommendation_rule=load_recommendation_rule(INPUT_MANIFEST),
    )


@app.errorhandler(413)
def upload_too_large(_error):
    return jsonify({"success": False, "message": "Yükleme toplam boyut sınırını aşıyor."}), 413


@app.route("/api/save-expiry", methods=["POST"])
def save_expiry():
    """Tarih bilgilerini hesaplama yapmadan anında kaydeder."""
    data = request.get_json(silent=True) or {}
    try:
        # Tek dosya expiry'leri
        single_expiries = data.get("single_expiries", {})
        if single_expiries and isinstance(single_expiries, dict):
            save_single_file_expiries(INPUT_MANIFEST, single_expiries)

        # Çoklu dosya (counter) expiry'leri
        counter_expiries = data.get("counter_expiries", {})
        if counter_expiries and isinstance(counter_expiries, dict):
            from input_files import load_counter_configs, save_counter_configs
            configs = load_counter_configs(INPUT_MANIFEST)
            for cfg in configs:
                cid = cfg.get("id", "")
                if cid in counter_expiries:
                    cfg["expiry_date"] = counter_expiries[cid]
            save_counter_configs(INPUT_MANIFEST, configs)

        # Çoklu dosya (plus_extra) expiry'leri
        plus_extra_expiries = data.get("plus_extra_expiries", {})
        if plus_extra_expiries and isinstance(plus_extra_expiries, dict):
            from input_files import save_plus_extra_configs
            configs = load_plus_extra_configs(INPUT_MANIFEST)
            for cfg in configs:
                cid = cfg.get("id", "")
                if cid in plus_extra_expiries:
                    cfg["expiry_date"] = plus_extra_expiries[cid]
            save_plus_extra_configs(INPUT_MANIFEST, configs)

        # Çoklu dosya (coupon) expiry'leri
        coupon_expiries = data.get("coupon_expiries", {})
        if coupon_expiries and isinstance(coupon_expiries, dict):
            from input_files import load_coupon_configs, save_coupon_configs
            configs = load_coupon_configs(INPUT_MANIFEST)
            for cfg in configs:
                cid = cfg.get("id", "")
                if cid in coupon_expiries:
                    cfg["expiry_date"] = coupon_expiries[cid]
            save_coupon_configs(INPUT_MANIFEST, configs)

        return jsonify({"success": True})
    except Exception:
        app.logger.exception("Tarih kaydedilirken hata")
        return jsonify({"success": False, "message": "Tarih kaydedilemedi."}), 500


@app.route("/api/toggle-campaign-enabled", methods=["POST"])
def toggle_campaign_enabled():
    """Kampanya kartının Aktif/Pasif durumunu anında kaydedip saklar."""
    data = request.get_json(silent=True) or {}
    try:
        camp_type = data.get("type")  # 'counter', 'plus_extra', or 'coupon'
        item_id = data.get("id")
        enabled = bool(data.get("enabled", True))

        if camp_type == "counter":
            from input_files import load_counter_configs, save_counter_configs
            configs = load_counter_configs(INPUT_MANIFEST)
            for item in configs:
                if item.get("id") == item_id:
                    item["enabled"] = enabled
            save_counter_configs(INPUT_MANIFEST, configs)
        elif camp_type == "plus_extra":
            from input_files import load_plus_extra_configs, save_plus_extra_configs
            configs = load_plus_extra_configs(INPUT_MANIFEST)
            for item in configs:
                if item.get("id") == item_id:
                    item["enabled"] = enabled
            save_plus_extra_configs(INPUT_MANIFEST, configs)
        elif camp_type == "coupon":
            from input_files import load_coupon_configs, save_coupon_configs
            configs = load_coupon_configs(INPUT_MANIFEST)
            for item in configs:
                if item.get("id") == item_id:
                    item["enabled"] = enabled
            save_coupon_configs(INPUT_MANIFEST, configs)

        return jsonify({"success": True})
    except Exception as e:
        app.logger.exception("Kampanya durumu kaydedilirken hata")
        return jsonify({"success": False, "message": str(e)}), 500


@app.route("/api/remove-counter-file", methods=["POST"])
def remove_counter_file():
    try:
        data = request.get_json(silent=True) or {}
        item_id = data.get("id")
        item_path = data.get("path")
        
        counter_configs = load_counter_configs(INPUT_MANIFEST)
        new_configs = []
        for item in counter_configs:
            p = item.get("path") or item.get("stored_path")
            if (item_id and item.get("id") == item_id) or (item_path and p == item_path):
                if p and os.path.exists(p):
                    try: os.remove(p)
                    except Exception: pass
            else:
                new_configs.append(item)
        save_counter_configs(INPUT_MANIFEST, new_configs)
        return jsonify({"success": True, "counter_configs": new_configs})
    except Exception as e:
        return jsonify({"success": False, "message": str(e)}), 500


@app.route("/api/remove-plus-extra-file", methods=["POST"])
def remove_plus_extra_file():
    try:
        data = request.get_json(silent=True) or {}
        item_id = data.get("id")
        item_path = data.get("path")
        
        plus_extra_configs = load_plus_extra_configs(INPUT_MANIFEST)
        new_configs = []
        for item in plus_extra_configs:
            p = item.get("path") or item.get("stored_path")
            if (item_id and item.get("id") == item_id) or (item_path and p == item_path):
                if p and os.path.exists(p):
                    try: os.remove(p)
                    except Exception: pass
            else:
                new_configs.append(item)
        save_plus_extra_configs(INPUT_MANIFEST, new_configs)
        return jsonify({"success": True, "plus_extra_configs": new_configs})
    except Exception as e:
        return jsonify({"success": False, "message": str(e)}), 500


@app.route("/api/remove-coupon-file", methods=["POST"])
def remove_coupon_file():
    try:
        data = request.get_json(silent=True) or {}
        item_id = data.get("id")
        item_path = data.get("path")
        
        from input_files import load_coupon_configs, save_coupon_configs
        coupon_configs = load_coupon_configs(INPUT_MANIFEST)
        new_configs = []
        for item in coupon_configs:
            p = item.get("path") or item.get("stored_path")
            if (item_id and item.get("id") == item_id) or (item_path and p == item_path):
                if p and os.path.exists(p):
                    try: os.remove(p)
                    except Exception: pass
            else:
                new_configs.append(item)
        save_coupon_configs(INPUT_MANIFEST, new_configs)
        return jsonify({"success": True, "coupon_configs": new_configs})
    except Exception as e:
        return jsonify({"success": False, "message": str(e)}), 500


@app.route("/api/data")
def get_data():
    if not calculation_result_is_current():
        return jsonify({"needs_calculation": True, "message": "Lütfen önce 'Verileri Güncelle' butonuna basarak hesaplamaları başlatın."}), 200
        
    df = restore_persisted_collections(pd.read_excel(F_HESAP))
    records = df.to_dict(orient="records")
    user_selections = load_user_selections(INPUT_MANIFEST)
    if user_selections and isinstance(user_selections, dict):
        for rec in records:
            b = str(rec.get("Barkod", "")).strip()
            if b in user_selections:
                val = user_selections[b]
                if isinstance(val, dict):
                    rec["İlk Kampanya Seçimi"] = val.get("main", "Hiçbiri")
                    rec["İlk Ekstra Kampanya Seçimi"] = val.get("extra", "Hiçbiri")
                else:
                    rec["İlk Kampanya Seçimi"] = str(val)
    cleaned_records = clean_nans(records)
    return jsonify(cleaned_records), 200

@app.route("/api/save-selections", methods=["POST"])
def save_selections_endpoint():
    try:
        data = request.get_json(silent=True)
        selections = data.get("selections") if isinstance(data, dict) else None
        if not selection_payload_is_valid(selections):
            return jsonify({"success": False, "message": "Geçersiz veri biçimi."}), 400
        normalized_selections = {}
        for barcode, selection in selections.items():
            main, extra = normalize_selection(selection)
            normalized_selections[barcode] = {"main": main, "extra": extra}
        save_user_selections(INPUT_MANIFEST, normalized_selections)
        return jsonify({"success": True, "saved_count": len(normalized_selections)})
    except Exception:
        app.logger.exception("Seçimler kaydedilirken hata")
        return jsonify({"success": False, "message": "Seçimler kaydedilemedi."}), 500


@app.route("/api/recommendation-rule", methods=["POST"])
def save_recommendation_rule_endpoint():
    try:
        payload = request.get_json(silent=True)
        if payload is None:
            raise InputValidationError("Öneri kuralı boş olamaz.")
        if isinstance(payload, dict) and "recommendation_rule" in payload:
            if set(payload) != {"recommendation_rule"}:
                raise InputValidationError("Öneri kuralı isteği geçersiz.")
            payload = payload["recommendation_rule"]
        rule = save_recommendation_rule(INPUT_MANIFEST, payload)
        return jsonify({"success": True, "recommendation_rule": rule})
    except InputValidationError as exc:
        return jsonify({"success": False, "message": str(exc)}), 400
    except Exception:
        app.logger.exception("Öneri kuralı kaydedilirken hata")
        return jsonify({"success": False, "message": "Öneri kuralı kaydedilemedi."}), 500

@app.route("/api/calculate", methods=["POST"])
def calculate():
    try:
        from komisyon_hesaplayici import calculate_all
        from input_files import save_counter_configs, INPUT_SPECS

        counter_configs = parse_campaign_configs_json(
            request.form.get("counter_configs_json"), "counter"
        )
        plus_extra_configs = parse_campaign_configs_json(
            request.form.get("plus_extra_configs_json"), "plus_extra"
        )
        coupon_configs = parse_campaign_configs_json(
            request.form.get("coupon_configs_json"), "coupon"
        )
        recommendation_rule_raw = request.form.get("recommendation_rule_json")
        recommendation_rule = (
            load_recommendation_rule(INPUT_MANIFEST)
            if recommendation_rule_raw is None
            else parse_recommendation_rule_json(recommendation_rule_raw)
        )

        standard_files = {
            k: v for k, v in request.files.items() 
            if k in INPUT_SPECS and v and v.filename
        }

        input_files = save_upload_set(standard_files, UPLOAD_DIR, INPUT_MANIFEST)
        missing_base_inputs = [
            INPUT_SPECS[key]["label"]
            for key in ("discount", "commission", "current")
            if not input_files.get(key)
        ]
        if missing_base_inputs:
            raise InputValidationError(
                f"Zorunlu girdiler eksik: {', '.join(missing_base_inputs)}"
            )
        if recommendation_rule_raw is not None:
            save_recommendation_rule(INPUT_MANIFEST, recommendation_rule)

        single_expiries_raw = request.form.get("single_expiries_json")
        if single_expiries_raw:
            try:
                single_expiries = json.loads(single_expiries_raw)
                save_single_file_expiries(INPUT_MANIFEST, single_expiries)
            except Exception as exp_err:
                print("Single expiries parse error:", exp_err)

        counter_files = []
        counter_dir = os.path.join(UPLOAD_DIR, "counter_files")
        os.makedirs(counter_dir, exist_ok=True)

        for idx, item in enumerate(counter_configs):
            file_key = f"counter_file_{idx}"
            file_obj = request.files.get(file_key)
            stored_name = f"counter_{idx+1}.xlsx"
            target_path = os.path.join(counter_dir, stored_name)
            stored_path = item.get("stored_path") or item.get("path")
            original_name = item.get("original_name")

            if file_obj and file_obj.filename:
                file_obj.save(target_path)
                stored_path = target_path
                original_name = file_obj.filename
            elif os.path.exists(target_path):
                stored_path = target_path

            counter_files.append({
                **item,
                "id": item.get("id", f"counter_{idx+1}"),
                "label": build_campaign_label(item, "counter", idx),
                "filename": item.get("filename") or original_name or stored_name,
                "original_name": original_name,
                "path": stored_path,
                "stored_path": stored_path,
                "expiry_date": item.get("expiry_date", ""),
                "enabled": item.get("enabled", True) is not False,
            })

        save_counter_configs(INPUT_MANIFEST, counter_files)

        from input_files import save_plus_extra_configs, load_plus_extra_configs

        plus_extra_files = []
        plus_extra_dir = os.path.join(UPLOAD_DIR, "plus_extra_files")
        os.makedirs(plus_extra_dir, exist_ok=True)

        for idx, item in enumerate(plus_extra_configs):
            file_key = f"plus_extra_file_{idx}"
            file_obj = request.files.get(file_key)
            stored_name = f"plus_extra_{idx+1}.xlsx"
            target_path = os.path.join(plus_extra_dir, stored_name)
            stored_path = item.get("stored_path") or item.get("path")
            original_name = item.get("original_name")

            if file_obj and file_obj.filename:
                file_obj.save(target_path)
                stored_path = target_path
                original_name = file_obj.filename
            elif os.path.exists(target_path):
                stored_path = target_path

            plus_extra_files.append({
                **item,
                "id": item.get("id", f"plus_extra_{idx+1}"),
                "label": build_campaign_label(item, "plus_extra", idx),
                "filename": item.get("filename") or original_name or stored_name,
                "original_name": original_name,
                "path": stored_path,
                "stored_path": stored_path,
                "expiry_date": item.get("expiry_date", ""),
                "enabled": item.get("enabled", True) is not False,
            })

        save_plus_extra_configs(INPUT_MANIFEST, plus_extra_files)

        from input_files import save_coupon_configs, load_coupon_configs

        coupon_files = []
        coupon_dir = os.path.join(UPLOAD_DIR, "coupon_files")
        os.makedirs(coupon_dir, exist_ok=True)

        for idx, item in enumerate(coupon_configs):
            file_key = f"coupon_file_{idx}"
            file_obj = request.files.get(file_key)
            stored_name = f"coupon_{idx+1}.xlsx"
            target_path = os.path.join(coupon_dir, stored_name)
            stored_path = item.get("stored_path") or item.get("path")
            original_name = item.get("original_name")

            if file_obj and file_obj.filename:
                file_obj.save(target_path)
                stored_path = target_path
                original_name = file_obj.filename
            elif os.path.exists(target_path):
                stored_path = target_path

            coupon_files.append({
                **item,
                "id": item.get("id", f"coupon_{idx+1}"),
                "label": build_campaign_label(item, "coupon", idx),
                "filename": item.get("filename") or original_name or stored_name,
                "original_name": original_name,
                "path": stored_path,
                "stored_path": stored_path,
                "expiry_date": item.get("expiry_date", ""),
                "enabled": item.get("enabled", True) is not False,
            })

        save_coupon_configs(INPUT_MANIFEST, coupon_files)


        toplam_indirim = float(request.form.get("toplam_indirim", 0) or 0)
        trendyol_oran = float(request.form.get("trendyol_oran", 0) or 0)
        min_sepet = float(request.form.get("min_sepet", 0) or 0)
        karsilamali_config = {
            "min_sepet": min_sepet,
            "toplam_indirim": toplam_indirim,
            "trendyol_oran": trendyol_oran,
        }

        # Reset saved user selections on new calculation so all products start fresh as 'Hiçbiri'
        save_user_selections(INPUT_MANIFEST, {})
        user_selections = {}
        result = calculate_all(input_files, counter_files=counter_files, plus_extra_files=plus_extra_files, coupon_files=coupon_files, karsilamali_config=karsilamali_config, output_dir=OUTPUT_DIR, user_selections=user_selections, recommendation_rule=recommendation_rule)
        if result.get("success"):
            result["uploads"] = load_upload_status(UPLOAD_DIR, INPUT_MANIFEST)
            result["counter_configs"] = load_counter_configs(INPUT_MANIFEST)
            result["plus_extra_configs"] = load_plus_extra_configs(INPUT_MANIFEST)
            result["coupon_configs"] = load_coupon_configs(INPUT_MANIFEST)
            result["recommendation_rule"] = recommendation_rule
            pd.DataFrame(result["results"]).to_excel(F_HESAP, index=False)
            result = clean_nans(result)
        return jsonify(result), (200 if result.get("success") else 500)
    except InputValidationError as e:
        return jsonify({"success": False, "message": str(e)}), 400
    except ValueError:
        return jsonify({"success": False, "message": "Sayısal kampanya değerlerini kontrol edin."}), 400
    except Exception:
        app.logger.exception("Hesaplama sırasında hata")
        return jsonify({"success": False, "message": "Hesaplama sırasında beklenmeyen bir hata oluştu."}), 500


@app.route("/api/apply", methods=["POST"])
def apply_campaign():
    data = request.get_json(silent=True) or {}
    selections = data.get("selections", {})
    if not selection_payload_is_valid(selections):
        return jsonify({"success": False, "message": "Kampanya seçimleri geçersiz."}), 400
    visible_columns = data.get("visibleColumns")
    target_type = data.get("target_type", "Hepsi")
    if target_type not in VALID_TARGET_TYPES:
        return jsonify({"success": False, "message": "Çıktı türü geçersiz."}), 400
    try:
        input_files = load_upload_set(UPLOAD_DIR, INPUT_MANIFEST)
    except InputValidationError as e:
        return jsonify({"success": False, "message": str(e)}), 400

    if not calculation_result_is_current():
        return jsonify({
            "success": False,
            "message": "Önce yüklenen girdilerle hesaplama yapın.",
        }), 400
    try:
        raw_rows = restore_persisted_collections(
            pd.read_excel(F_HESAP)
        ).to_dict(orient="records")
    except Exception:
        app.logger.exception("Hesap sonucu okunamadı")
        return jsonify({
            "success": False,
            "message": "Hesap sonucu okunamadı; girdileri yeniden hesaplayın.",
        }), 500

    ignore_zero_stock = bool(data.get("ignore_zero_stock", True))

    table_data = []
    for row in raw_rows:
        barcode = str(row.get("Barkod", "")).strip()
        stok = as_number(row.get("Stok Adedi"))
        if ignore_zero_stock and stok is not None and stok == 0:
            continue

        selection = selections.get(barcode, {
            "main": row.get("İlk Kampanya Seçimi", "Hiçbiri") or "Hiçbiri",
            "extra": row.get("İlk Ekstra Kampanya Seçimi", "Hiçbiri") or "Hiçbiri",
        })
        normalized = normalize_selection(selection)
        if not normalized:
            return jsonify({
                "success": False,
                "message": "Kampanya seçimleri geçersiz.",
            }), 400
        main_selection, extra_selection = normalized
        if not row_selection_is_applicable(row, main_selection, extra_selection):
            return jsonify({
                "success": False,
                "message": "Seçilen kampanya ürün için uygulanabilir değil.",
            }), 400
        row["userSelection"] = main_selection
        row["userExtraSelection"] = extra_selection
        table_data.append(row)

    F_AVAN = input_files.get("advantage")
    F_FLAS = input_files.get("flash")
    F_PLUS = input_files.get("plus")
    F_PLUS_EK = input_files.get("plus_extra")
    F_KARS = input_files.get("counter")
    F_MUH_AVAN = input_files.get("muhasebe_avantaj")
    F_MUH_FLAS = input_files.get("muhasebe_flas")
    F_MUH_PLUS = input_files.get("muhasebe_plus")
    try:
        plus_extra_configs = load_plus_extra_configs(INPUT_MANIFEST)
    except InputValidationError as exc:
        return jsonify({"success": False, "message": str(exc)}), 400
    target_inputs = {
        "Avantajlı": F_AVAN or F_MUH_AVAN,
        "Flaş": F_FLAS or F_MUH_FLAS,
        "Plus": F_PLUS or F_MUH_PLUS,
        "Plus Ek İndirim": F_PLUS_EK or any(
            config.get("path") or config.get("stored_path")
            for config in plus_extra_configs
        ),
        "Karşılamalı Kampanya": F_KARS,
    }
    if target_type != "Hepsi" and target_type in target_inputs and not target_inputs[target_type]:
        return jsonify({
            "success": False,
            "message": f"{target_type} girdisi bu hesaplamada yüklenmedi.",
        }), 400
        
    # Her işlem için Tarih_Saat adında alt klasör oluştur
    timestamp_folder = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
    run_output_dir = os.path.join(OUTPUT_DIR, timestamp_folder)
    os.makedirs(run_output_dir, exist_ok=True)
    generated_files = []

    # Map barcodes to their table_data row dict for easy lookup of match statuses & recommendations
    row_by_barcode = {}
    for row in table_data:
        b_key = str(row.get("Barkod", "")).strip()
        if b_key:
            row_by_barcode[b_key] = row

    def get_selection(b_key):
        row = row_by_barcode.get(b_key, {})
        return (
            row.get("userSelection", "Hiçbiri") or "Hiçbiri",
            row.get("userExtraSelection", "Hiçbiri") or "Hiçbiri",
        )

    # 1. Process Avantajlı
    if target_type in ['Hepsi', 'Avantajlı'] and (F_AVAN or F_MUH_AVAN):
        try:
            wb_av = load_merged_campaign_workbook(F_AVAN, F_MUH_AVAN, "BARKOD")
            ws_av = wb_av.active
            b_idx_av = header_index(ws_av, 'BARKOD')
            tsf_idx_av = ensure_header(ws_av, 'YENİ TSF (FİYAT GÜNCELLE)')
            tarife_idx_av = ensure_header(ws_av, 'Tarife Sonuna Kadar Uygula')
        
            if b_idx_av and tsf_idx_av:
                keep_rows = []
                for r in range(2, ws_av.max_row + 1):
                    b_val = ws_av.cell(r, b_idx_av).value
                    if not b_val: continue
                    b_val_str = str(b_val).strip()
                    main_sel, extra_sel = get_selection(b_val_str)
                    row_info = row_by_barcode.get(b_val_str, {})
                    should_keep = (main_sel == "Avantajlı")

                    if should_keep:
                        selected_price = row_info.get('Avantajlı Ürün Fiyatı (YENİ TSF) (TL)')
                        if selected_price is not None and not pd.isna(selected_price):
                            ws_av.cell(r, tsf_idx_av).value = float(selected_price)
                        if tarife_idx_av:
                            ws_av.cell(r, tarife_idx_av).value = "Evet"
                        keep_rows.append(r)
                
                if keep_rows:
                    safe_keep_rows(ws_av, keep_rows)
                    out_name = os.path.join(run_output_dir, "Avantajli_Urun_Etiketleri.xlsx")
                    shrink_data_validations(ws_av)
                    wb_av.save(out_name)
                    fix_xlsx_for_trendyol(out_name)
                    generated_files.append(os.path.join(timestamp_folder, "Avantajli_Urun_Etiketleri.xlsx"))
        except Exception:
            return processing_error("Avantajlı dosya")

    # 2. Process Flaş (Grouped by Date)
    if target_type in ['Hepsi', 'Flaş'] and (F_FLAS or F_MUH_FLAS):
        try:
            wb_fl = load_merged_campaign_workbook(F_FLAS, F_MUH_FLAS, "Barkod")
            ws_fl = wb_fl.active
            b_idx_fl = header_index(ws_fl, 'Barkod')
            fiyat_24_idx = ensure_header(ws_fl, '24 Saat Fiyat')
            guncel_fiyat_idx = ensure_header(ws_fl, 'Güncellenecek Fiyat')
            baslangic_idx = ensure_header(ws_fl, '24 Saat Flaş Başlangıç Tarihi')
        
            if b_idx_fl and fiyat_24_idx and guncel_fiyat_idx and baslangic_idx:
                date_groups = {} # {date_str: [row_indices]}

                for r in range(2, ws_fl.max_row + 1):
                    b_val = ws_fl.cell(r, b_idx_fl).value
                    if not b_val: continue
                
                    b_val_str = str(b_val).strip()
                    main_sel, extra_sel = get_selection(b_val_str)
                    should_keep = (main_sel == "Flaş")
                
                    if should_keep:
                        date_val = str(ws_fl.cell(r, baslangic_idx).value or "").strip()
                        date_key = (
                            date_val.split()[0].replace('/', '_').replace('-', '_')
                            if date_val
                            else "Genel"
                        )
                        if date_key not in date_groups:
                            date_groups[date_key] = []
                        date_groups[date_key].append(r)
            
                for date_key, keep_rows in date_groups.items():
                    if keep_rows:
                        wb_copy = clone_workbook(wb_fl)
                        ws_copy = wb_copy.active
                        for r in keep_rows:
                            barcode = str(ws_copy.cell(r, b_idx_fl).value or '').strip()
                            selected_price = row_by_barcode.get(barcode, {}).get('Flaş Ürün 24 Saat Fiyatı (TL)')
                            if selected_price is not None and not pd.isna(selected_price):
                                ws_copy.cell(r, fiyat_24_idx).value = float(selected_price)
                            ws_copy.cell(r, guncel_fiyat_idx).value = "24 Saat"
                        safe_keep_rows(ws_copy, keep_rows)
                        out_name = os.path.join(run_output_dir, f"Flas_Urunler_{date_key}.xlsx")
                        shrink_data_validations(ws_copy)
                        wb_copy.save(out_name)
                        fix_xlsx_for_trendyol(out_name)
                        generated_files.append(os.path.join(timestamp_folder, f"Flas_Urunler_{date_key}.xlsx"))
                
        except Exception:
            return processing_error("Flaş dosya")

    # 3. Process Plus (Grouped by Date Interval)
    if target_type in ['Hepsi', 'Plus']:
        if F_PLUS or F_MUH_PLUS:
            try:
                wb_plus = load_merged_campaign_workbook(F_PLUS, F_MUH_PLUS, "Barkod")
                ws_plus = wb_plus.active
                header_plus = [ws_plus.cell(1, c).value for c in range(1, ws_plus.max_column + 1)]
                b_idx_plus = header_index(ws_plus, 'Barkod')
                fiyat_secim_idx = ensure_header(ws_plus, 'Plus Fiyat Seçimi')
                tarife_secim_idx = ensure_header(ws_plus, 'Tarife Seçimi')
                ust_limit_idx = ensure_header(ws_plus, 'Plus Fiyat Üst Limiti')
                header_plus = [ws_plus.cell(1, c).value for c in range(1, ws_plus.max_column + 1)]
            
                tarih_idx_plus = None
                gun_sayisi = 7
                for idx_c, col in enumerate(header_plus):
                    if col and "Tarih Aralığı" in str(col):
                        tarih_idx_plus = idx_c + 1
                        import re
                        match = re.search(r'\((\d+)\s*Gün\)', str(col))
                        if match:
                            gun_sayisi = int(match.group(1))
                        break
                    
                if b_idx_plus and fiyat_secim_idx and tarife_secim_idx and ust_limit_idx:
                    date_groups = {}  # {date_key: [row_indices]}
                    import re

                    for r in range(2, ws_plus.max_row + 1):
                        b_val = ws_plus.cell(r, b_idx_plus).value
                        if not b_val: continue
                        b_val_str = str(b_val).strip()
                        main_sel, extra_sel = get_selection(b_val_str)
                        should_keep = (main_sel == "Plus")

                        if should_keep:
                            date_key = "Genel"
                            if tarih_idx_plus:
                                date_val = str(ws_plus.cell(r, tarih_idx_plus).value or "").strip()
                                if date_val:
                                    tr_map = str.maketrans("çğıöşüÇĞİÖŞÜ", "cgiosuCGIOSU")
                                    clean_d = date_val.translate(tr_map)
                                    date_key = re.sub(r'[^\w\.\-]', '_', clean_d)
                                    date_key = re.sub(r'_+', '_', date_key).strip('_')

                            if not date_key:
                                date_key = "Genel"

                            if date_key not in date_groups:
                                date_groups[date_key] = []
                            date_groups[date_key].append(r)

                    for date_key, keep_rows in date_groups.items():
                        if keep_rows:
                            wb_copy = clone_workbook(wb_plus)
                            ws_copy = wb_copy.active
                            for r in keep_rows:
                                ust_lim = ws_copy.cell(r, ust_limit_idx).value
                                barcode = str(ws_copy.cell(r, b_idx_plus).value or '').strip()
                                row_info = row_by_barcode.get(barcode, {})
                                selected_price = row_info.get('Plus Fiyatı (TL)')
                                ws_copy.cell(r, fiyat_secim_idx).value = selected_price if selected_price is not None and not pd.isna(selected_price) else ust_lim
                                ws_copy.cell(r, tarife_secim_idx).value = f"{gun_sayisi} Günlük Fiyat"

                            safe_keep_rows(ws_copy, keep_rows)
                            if date_key == "Genel":
                                file_name = "Plus_Komisyon_Tarifeleri.xlsx"
                            else:
                                file_name = f"Plus_Komisyon_Tarifeleri_{date_key}.xlsx"

                            out_name = os.path.join(run_output_dir, file_name)
                            shrink_data_validations(ws_copy)
                            wb_copy.save(out_name)
                            fix_xlsx_for_trendyol(out_name)
                            generated_files.append(os.path.join(timestamp_folder, file_name))
            except Exception:
                return processing_error("Plus dosya")

    # 4. Process Plus Ek İndirim (Çoklu Dosya Desteği)
    if target_type in ['Hepsi', 'Plus Ek İndirim']:
        try:
            plus_extra_configs = load_plus_extra_configs(INPUT_MANIFEST)

            for idx, pe_item in enumerate(plus_extra_configs):
                pe_item = normalize_campaign_config(pe_item, "plus_extra")
                pe_path = pe_item.get('stored_path') or pe_item.get('path')
                if pe_path and os.path.exists(pe_path):
                    wb_pe = openpyxl.load_workbook(pe_path)
                    ws_pe = wb_pe.active
                    header_pe = [ws_pe.cell(1, c).value for c in range(1, ws_pe.max_column + 1)]
                    b_idx_pe = header_pe.index('Barkod') + 1 if 'Barkod' in header_pe else None
                    fiyat_idx_pe = header_pe.index('Kampanyalı Satış Fiyatı') + 1 if 'Kampanyalı Satış Fiyatı' in header_pe else None
                    max_fiyat_idx_pe = header_pe.index('Maksimum Girebileceğin Fiyat') + 1 if 'Maksimum Girebileceğin Fiyat' in header_pe else None
                    
                    if b_idx_pe and fiyat_idx_pe and max_fiyat_idx_pe:
                        c_label = build_campaign_label(pe_item, "plus_extra", idx)
                        keep_rows = []
                        for r in range(2, ws_pe.max_row + 1):
                            b_val = ws_pe.cell(r, b_idx_pe).value
                            if not b_val: continue

                            b_val_str = str(b_val).strip()
                            main_sel, extra_sel = get_selection(b_val_str)
                            should_keep = extra_sel == c_label
                            if should_keep:
                                max_f = ws_pe.cell(r, max_fiyat_idx_pe).value
                                row_info = row_by_barcode.get(b_val_str, {})
                                evaluation = row_info.get("counter_evaluations", {}).get(c_label, {})
                                campaign_price = evaluation.get("customer_price")
                                if campaign_price is None:
                                    max_f_num = as_number(max_f)
                                    if max_f_num is not None:
                                        discount_amount = pe_item["discount_amount"]
                                        total_discount = round2(max_f_num * (discount_amount / 100.0)) if pe_item["discount_type"] == "%" else discount_amount
                                        campaign_price = round2(max_f_num - total_discount)
                                if campaign_price is not None:
                                    ws_pe.cell(r, fiyat_idx_pe).value = float(campaign_price)
                                keep_rows.append(r)

                        if keep_rows:
                            safe_keep_rows(ws_pe, keep_rows)
                            file_name = plus_extra_export_filename(pe_item, idx)
                            out_name = os.path.join(run_output_dir, file_name)
                            shrink_data_validations(ws_pe)
                            wb_pe.save(out_name)
                            fix_xlsx_for_trendyol(out_name)
                            generated_files.append(os.path.join(timestamp_folder, file_name))
        except Exception as e:
            print("Plus Ek İndirim export error:", e)

    # 5. Process Karşılamalı Kampanya (Çoklu Dosya Desteği)
    if target_type in ['Hepsi', 'Karşılamalı Kampanya']:
        try:
            from input_files import load_counter_configs
            counter_configs = load_counter_configs(INPUT_MANIFEST)
            
            for idx, c_item in enumerate(counter_configs):
                c_path = c_item.get('stored_path') or c_item.get('path')
                if c_path and os.path.exists(c_path):
                    wb_kars = openpyxl.load_workbook(c_path)
                    ws_kars = wb_kars.active
                    header_kars = [ws_kars.cell(1, c).value for c in range(1, ws_kars.max_column + 1)]
                    b_idx_kars = header_kars.index('Barkod') + 1 if 'Barkod' in header_kars else None
                    fiyat_idx_kars = header_kars.index('Kampanyalı Satış Fiyatı') + 1 if 'Kampanyalı Satış Fiyatı' in header_kars else None
                    max_fiyat_idx_kars = header_kars.index('Maksimum Girebileceğin Fiyat') + 1 if 'Maksimum Girebileceğin Fiyat' in header_kars else None
                    
                    if b_idx_kars and fiyat_idx_kars and max_fiyat_idx_kars:
                        c_label = c_item.get('label') or f"Karşılamalı #{idx+1}"
                        keep_rows = []
                        for r in range(2, ws_kars.max_row + 1):
                            b_val = ws_kars.cell(r, b_idx_kars).value
                            if not b_val: continue
                            
                            b_val_str = str(b_val).strip()
                            main_sel, extra_sel = get_selection(b_val_str)
                            should_keep = (extra_sel == c_label) or (target_type == "Karşılamalı Kampanya" and extra_sel.startswith("Karşılamalı"))

                            if should_keep:
                                max_fiyat_val = ws_kars.cell(r, max_fiyat_idx_kars).value
                                row_info = row_by_barcode.get(b_val_str, {})
                                evaluation = row_info.get('counter_evaluations', {}).get(c_label, {})
                                campaign_price = evaluation.get('price') or max_fiyat_val
                                if campaign_price:
                                    ws_kars.cell(r, fiyat_idx_kars).value = float(campaign_price)
                                keep_rows.append(r)
                                
                        if keep_rows:
                            safe_keep_rows(ws_kars, keep_rows)
                            
                            def format_num_clean(val):
                                if val is None or val == "": return None
                                try:
                                    n = float(val)
                                    return f"{int(n)}" if n.is_integer() else f"{n}"
                                except (ValueError, TypeError):
                                    return str(val).strip()

                            min_p = format_num_clean(c_item.get('min_price'))
                            disc = format_num_clean(c_item.get('discount_amount'))
                            tp = format_num_clean(c_item.get('trendyol_percent'))
                            disc_type = c_item.get('discount_type') or c_item.get('discount_unit') or 'TL'

                            if not min_p or not disc:
                                import re
                                m_pct = re.search(r'(\d+(?:[\.,]\d+)?)\s*TL\s*Üzeri\s*/\s*%\s*(\d+(?:[\.,]\d+)?)', c_label, re.IGNORECASE)
                                if not m_pct:
                                    m_pct = re.search(r'(\d+(?:[\.,]\d+)?)\s*TL\s*Üzeri\s*/\s*(\d+(?:[\.,]\d+)?)\s*%', c_label, re.IGNORECASE)
                                if m_pct:
                                    min_p = min_p or format_num_clean(m_pct.group(1))
                                    disc = disc or format_num_clean(m_pct.group(2))
                                    disc_type = '%'
                                else:
                                    m = re.search(r'(\d+(?:[\.,]\d+)?)\s*TL\s*Üzeri\s*/\s*(\d+(?:[\.,]\d+)?)\s*TL', c_label, re.IGNORECASE)
                                    if m:
                                        min_p = min_p or format_num_clean(m.group(1))
                                        disc = disc or format_num_clean(m.group(2))

                            if min_p and disc:
                                disc_part = f"%{disc}" if disc_type == '%' else f"{disc}_TL"
                                if tp and tp != '0':
                                    out_filename = f"{min_p}_TL_Uzeri_{disc_part}_Indirim_%{tp}_Trendyol_Karsilamali.xlsx"
                                else:
                                    out_filename = f"{min_p}_TL_Uzeri_{disc_part}_Indirim_Trendyol_Karsilamali.xlsx"
                            else:
                                tr_map = str.maketrans("çğıöşüÇĞİÖŞÜ", "cgiosuCGIOSU")
                                clean_l = c_label.translate(tr_map)
                                safe_l = re.sub(r'[^\w%]', '_', clean_l)
                                safe_l = re.sub(r'_+', '_', safe_l).strip('_')
                                out_filename = f"Karsilamali_{safe_l}.xlsx"

                            out_name = os.path.join(run_output_dir, out_filename)
                            shrink_data_validations(ws_kars)
                            wb_kars.save(out_name)
                            fix_xlsx_for_trendyol(out_name)
                            generated_files.append(os.path.join(timestamp_folder, out_filename))
        except Exception:
            return processing_error("Karşılamalı Kampanya dosyası")

    # 6. Process Kupon (Çoklu Dosya Desteği)
    if target_type in ['Hepsi', 'Kupon', 'Plus Kupon', 'Karşılamalı Kampanya']:
        try:
            from input_files import load_coupon_configs
            coupon_configs = load_coupon_configs(INPUT_MANIFEST)
            
            for idx, cp_item in enumerate(coupon_configs):
                cp_path = cp_item.get('stored_path') or cp_item.get('path')
                if cp_path and os.path.exists(cp_path):
                    wb_cp = openpyxl.load_workbook(cp_path)
                    ws_cp = wb_cp.active
                    header_cp = [ws_cp.cell(1, c).value for c in range(1, ws_cp.max_column + 1)]
                    b_idx_cp = header_cp.index('Barkod') + 1 if 'Barkod' in header_cp else None
                    secim_idx_cp = header_cp.index('Eklenecek Ürünleri Seçiniz') + 1 if 'Eklenecek Ürünleri Seçiniz' in header_cp else None
                    
                    if b_idx_cp and secim_idx_cp:
                        cp_label = cp_item.get('label') or f"Kupon #{idx+1}"
                        keep_rows = []
                        for r in range(2, ws_cp.max_row + 1):
                            b_val = ws_cp.cell(r, b_idx_cp).value
                            if not b_val: continue
                            
                            b_val_str = str(b_val).strip()
                            main_sel, extra_sel = get_selection(b_val_str)
                            should_keep = (extra_sel == cp_label) or (target_type in ["Kupon", "Plus Kupon"] and extra_sel.startswith(cp_label.split()[0]))

                            if should_keep:
                                ws_cp.cell(r, secim_idx_cp).value = "Seçildi"
                                keep_rows.append(r)
                                
                        if keep_rows:
                            safe_keep_rows(ws_cp, keep_rows)
                            
                            def format_num_clean(val):
                                if val is None or val == "": return None
                                try:
                                    n = float(val)
                                    return f"{int(n)}" if n.is_integer() else f"{n}"
                                except (ValueError, TypeError):
                                    return str(val).strip()

                            min_p = format_num_clean(cp_item.get('min_price'))
                            disc = format_num_clean(cp_item.get('discount_amount'))
                            tp = format_num_clean(cp_item.get('trendyol_percent'))
                            disc_type = cp_item.get('discount_type') or 'TL'

                            if min_p and disc:
                                disc_part = f"%{disc}" if disc_type == '%' else f"{disc}_TL"
                                if tp and tp != '0':
                                    out_filename = f"{min_p}_TL_Uzerine_{disc_part}_Kupon_%{tp}_Trendyol_Plus.xlsx"
                                else:
                                    out_filename = f"{min_p}_TL_Uzerine_{disc_part}_Kupon_Trendyol_Plus.xlsx"
                            else:
                                tr_map = str.maketrans("çğıöşüÇĞİÖŞÜ", "cgiosuCGIOSU")
                                clean_l = cp_label.translate(tr_map)
                                safe_l = re.sub(r'[^\w%]', '_', clean_l)
                                safe_l = re.sub(r'_+', '_', safe_l).strip('_')
                                out_filename = f"Kupon_{safe_l}.xlsx"

                            out_name = os.path.join(run_output_dir, out_filename)
                            shrink_data_validations(ws_cp)
                            wb_cp.save(out_name)
                            fix_xlsx_for_trendyol(out_name)
                            generated_files.append(os.path.join(timestamp_folder, out_filename))
        except Exception as e:
            print("Kupon export error:", e)
            return processing_error("Kupon Kampanyası dosyası")

    # Ekstra Rapor ve Uygulanmayanlar Excel Çıktıları
    try:
        df_all = pd.DataFrame(table_data)
        if not df_all.empty:
            if 'checked' in df_all.columns:
                df_all = df_all.drop(columns=['checked'])
            if 'userSelection' not in df_all.columns:
                df_all['userSelection'] = 'Hiçbiri'
            if 'userExtraSelection' not in df_all.columns:
                df_all['userExtraSelection'] = 'Hiçbiri'
            selection_columns = {
                'userSelection': 'Uygulanan Kampanya Seçimi',
                'userExtraSelection': 'Uygulanan Ekstra Kampanya Seçimi',
            }
            
            # 1. Uygulanmayanlar
            df_unapplied = df_all[
                (df_all['userSelection'] == 'Hiçbiri')
                & (df_all['userExtraSelection'] == 'Hiçbiri')
            ].copy()
            df_unapplied.rename(columns=selection_columns, inplace=True)
            out_unapplied = os.path.join(run_output_dir, "Uygulanmayan_Urunler_Raporu.xlsx")
            df_unapplied.to_excel(out_unapplied, index=False)
            fix_xlsx_for_trendyol(out_unapplied)
            generated_files.append(os.path.join(timestamp_folder, "Uygulanmayan_Urunler_Raporu.xlsx"))
            
            # 2. Tüm Rapor
            df_all.rename(columns=selection_columns, inplace=True)
            out_report = os.path.join(run_output_dir, "Kampanya_Genel_Raporu.xlsx")
            df_all.to_excel(out_report, index=False)
            fix_xlsx_for_trendyol(out_report)
            generated_files.append(os.path.join(timestamp_folder, "Kampanya_Genel_Raporu.xlsx"))
            
            # 3. Sayfada seçilen sütunlarla aynı sıradaki özet rapor
            df_summary = build_report_dataframe(table_data, visible_columns)
            out_summary = os.path.join(run_output_dir, "Kampanya_Ozet_Raporu.xlsx")
            write_report_excel(df_summary, out_summary)
            fix_xlsx_for_trendyol(out_summary)
            generated_files.append(os.path.join(timestamp_folder, "Kampanya_Ozet_Raporu.xlsx"))
            
            try:
                from fiyat_farki_analiz_script import generate_fiyat_farki_raporu
                generate_fiyat_farki_raporu(run_output_dir, input_files.get("discount"))
                out_kiyas = os.path.join(run_output_dir, "Indirim_Uygulanmayan_Fiyat_Kiyas_Raporu.xlsx")
                if os.path.exists(out_kiyas):
                    fix_xlsx_for_trendyol(out_kiyas)
                generated_files.append(os.path.join(timestamp_folder, "Indirim_Uygulanmayan_Fiyat_Kiyas_Raporu.xlsx"))
            except Exception as e:
                print("Kıyas Raporu hatası:", str(e))
                pass
    except Exception:
        return processing_error("Rapor")
        
    files_str = "\n".join([f"- {f}" for f in generated_files])
    return jsonify({
        "success": True, 
        "message": f"Tarihlere göre {len(generated_files)} dosya başarıyla oluşturuldu!\n\nOluşturulan Dosyalar:\n{files_str}",
        "generated_files": generated_files,
        "timestamp_folder": timestamp_folder
    })

if __name__ == "__main__":
    app.run(port=5114, debug=True)
