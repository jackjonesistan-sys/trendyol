import os
import json
import math
import re
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
    "Uygulanan Kampanya",
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
VALID_SELECTIONS = {
    "Hiçbiri",
    "Avantajlı",
    "Flaş",
    "Plus",
    "Plus Ek İndirim %5",
    "Plus Ek İndirim %10",
    "Plus Ek İndirim %20",
    "Karşılamalı Kampanya",
}
VALID_TARGET_TYPES = {
    "Hepsi",
    "Avantajlı",
    "Flaş",
    "Plus",
    "Plus Ek İndirim",
    "Karşılamalı Kampanya",
}


def as_number(value):
    try:
        number = float(value)
    except (TypeError, ValueError):
        return None
    return None if pd.isna(number) else number


def round2(value):
    return math.floor((float(value) * 100) + 0.5 + ROUNDING_EPSILON) / 100


def discounted_price(value, rate):
    price = as_number(value)
    return round2(price * (1 - rate / 100)) if price is not None else None


def discount_between(current_price, campaign_price):
    current = as_number(current_price)
    campaign = as_number(campaign_price)
    if current is None or campaign is None or current <= 0 or campaign <= 0:
        return None, None
    amount = round2(max(current - campaign, 0))
    return amount, round2((amount / current) * 100)


def selected_campaign_values(row):
    selection = row.get("userSelection", "Hiçbiri") or "Hiçbiri"
    if selection == "Hiçbiri":
        return (
            as_number(row.get("Güncel Ürün Fiyatı (TL)")),
            as_number(row.get("Güncel Ürün Kalan Net (TL)")),
            as_number(row.get("Güncel Ürün Komisyon (%)")),
        )
    if selection.startswith("Plus Ek İndirim %"):
        rate = int(selection.rsplit("%", 1)[-1])
        base_price = as_number(row.get(f"Plus Ek Fiyatı %{rate} (TL)"))
        campaign_price = discounted_price(base_price, rate)
        return (
            campaign_price,
            as_number(row.get(f"Plus Ek Net %{rate} (TL)")),
            as_number(row.get("Plus Ek Komisyon (%)")),
        )

    counter_evals = row.get("counter_evaluations", {})
    if isinstance(counter_evals, dict) and selection in counter_evals:
        c_info = counter_evals[selection]
        return (
            as_number(c_info.get("price")),
            as_number(c_info.get("net")),
            as_number(c_info.get("rate")),
        )

    fields = {
        "Avantajlı": (
            "Avantajlı Ürün Fiyatı (YENİ TSF) (TL)",
            "Avantajlı Ürün Kalan Net (TL)",
            "Avantajlı Ürün Komisyon (%)",
        ),
        "Flaş": (
            "Flaş Ürün 24 Saat Fiyatı (TL)",
            "Flaş Ürün Kalan Net (TL)",
            "Flaş Ürün Komisyon (%)",
        ),
        "Plus": ("Plus Fiyatı (TL)", "Plus Net (TL)", "Plus Komisyon (%)"),
        "Karşılamalı Kampanya": (
            "Karşılamalı Kampanya Fiyatı (TL)",
            "Karşılamalı Kampanya Kalan Net (TL)",
            "Karşılamalı Kampanya Komisyon (%)",
        ),
    }
    selected = fields.get(selection)
    if selected is None:
        return None, None, None
    price_field, net_field, commission_field = selected
    return (
        as_number(row.get(price_field)),
        as_number(row.get(net_field)),
        as_number(row.get(commission_field)),
    )


def build_report_row(row):
    selection = row.get("userSelection", "Hiçbiri") or "Hiçbiri"
    is_discount_eligible = row.get("İndirim Uygulanabilir") == "Evet"
    current_price = as_number(row.get("Güncel Ürün Fiyatı (TL)"))
    campaign_price, campaign_net, campaign_commission = selected_campaign_values(row)
    applied_amount, applied_percent = discount_between(current_price, campaign_price)
    dip_price = (
        as_number(row.get("Düşülebilecek Dip Fiyat (TL)"))
        if is_discount_eligible
        else None
    )
    available_amount, available_percent = (
        discount_between(current_price, dip_price)
        if is_discount_eligible
        else (None, None)
    )
    extra_amount = None
    extra_percent = None
    if available_amount is not None and applied_amount is not None:
        extra_amount = round2(max(available_amount - applied_amount, 0))
        extra_percent = round2((extra_amount / current_price) * 100)

    if selection.startswith("Plus Ek İndirim %"):
        plus_rate = int(selection.rsplit("%", 1)[-1])
        plus_base_price = as_number(row.get(f"Plus Ek Fiyatı %{plus_rate} (TL)"))
        plus_extra_price = discounted_price(plus_base_price, plus_rate)
        plus_extra_net = as_number(row.get(f"Plus Ek Net %{plus_rate} (TL)"))
    else:
        plus_extra_price = None
        plus_extra_net = None

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
        "Uygulanan Kampanya": CAMPAIGN_LABELS.get(selection, selection),
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
    campaign = "Plus Ek İndirim" if selection.startswith("Plus Ek İndirim %") else selection
    applicable = {
        item.strip()
        for item in str(applicable_value or "").split(",")
        if item.strip()
    }
    return campaign in applicable


def calculation_result_is_current():
    try:
        return (
            os.path.isfile(F_HESAP)
            and os.path.isfile(INPUT_MANIFEST)
            and os.path.getmtime(F_HESAP) >= os.path.getmtime(INPUT_MANIFEST)
        )
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
    import re
    if not keep_row_indices:
        if ws.max_row > 1:
            ws.delete_rows(2, ws.max_row - 1)
        return

    rows_to_delete = []
    for r in range(2, ws.max_row + 1):
        if r not in keep_row_indices:
            rows_to_delete.append(r)
    
    blocks = []
    if rows_to_delete:
        start = rows_to_delete[0]
        end = start
        for r in rows_to_delete[1:]:
            if r == end + 1:
                end = r
            else:
                blocks.append((start, end - start + 1))
                start = r
                end = r
        blocks.append((start, end - start + 1))
        
    for start, amount in reversed(blocks):
        ws.delete_rows(start, amount)
        
    for r in range(2, ws.max_row + 1):
        for c in range(1, ws.max_column + 1):
            cell = ws.cell(r, c)
            val = cell.value
            if isinstance(val, str) and (val.startswith('=') or val.startswith('==')):
                new_val = re.sub(r'([a-zA-Z]+)\d+\b', r'\g<1>' + str(r), val)
                cell.value = new_val

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
    return render_template(
        "index.html",
        report_columns=REPORT_COLUMNS,
        input_specs=INPUT_SPECS,
        uploaded_inputs=load_upload_status(UPLOAD_DIR, INPUT_MANIFEST),
    )


@app.errorhandler(413)
def upload_too_large(_error):
    return jsonify({"success": False, "message": "Yükleme toplam boyut sınırını aşıyor."}), 413

@app.route("/api/data")
def get_data():
    if not calculation_result_is_current():
        return jsonify({"needs_calculation": True, "message": "Lütfen önce 'Verileri Güncelle' butonuna basarak hesaplamaları başlatın."}), 200
        
    df = pd.read_excel(F_HESAP)
    return app.response_class(
        response=df.to_json(orient="records"),
        status=200,
        mimetype='application/json'
    )

@app.route("/api/calculate", methods=["POST"])
def calculate():
    try:
        from komisyon_hesaplayici import calculate_all
        from input_files import save_counter_configs, INPUT_SPECS

        standard_files = {
            k: v for k, v in request.files.items() 
            if k in INPUT_SPECS and v and v.filename
        }

        input_files = save_upload_set(standard_files, UPLOAD_DIR, INPUT_MANIFEST)

        counter_configs_raw = request.form.get("counter_configs_json")
        counter_configs = json.loads(counter_configs_raw) if counter_configs_raw else []

        counter_files = []
        counter_dir = os.path.join(UPLOAD_DIR, "counter_files")
        os.makedirs(counter_dir, exist_ok=True)

        for idx, item in enumerate(counter_configs):
            file_key = f"counter_file_{idx}"
            file_obj = request.files.get(file_key)
            stored_name = f"counter_{idx+1}.xlsx"
            target_path = os.path.join(counter_dir, stored_name)

            if file_obj and file_obj.filename:
                file_obj.save(target_path)
                item["stored_path"] = target_path
                item["original_name"] = file_obj.filename
            else:
                if os.path.exists(target_path):
                    item["stored_path"] = target_path

            min_p = float(item.get("min_price", 0))
            disc_amt = float(item.get("discount_amount", 0))
            tr_pct = float(item.get("trendyol_percent", 0))
            label = item.get("label") or (f"Karşılamalı ({int(min_p) if min_p.is_integer() else min_p} TL Üzeri / {int(disc_amt) if disc_amt.is_integer() else disc_amt} TL İndirim)" if min_p > 0 else f"Karşılamalı #{idx+1}")

            counter_files.append({
                "id": item.get("id", f"counter_{idx+1}"),
                "label": label,
                "path": item.get("stored_path"),
                "min_price": min_p,
                "discount_amount": disc_amt,
                "trendyol_percent": tr_pct,
            })

        save_counter_configs(INPUT_MANIFEST, counter_files)

        toplam_indirim = float(request.form.get("toplam_indirim", 0) or 0)
        trendyol_oran = float(request.form.get("trendyol_oran", 0) or 0)
        min_sepet = float(request.form.get("min_sepet", 0) or 0)
        karsilamali_config = {
            "min_sepet": min_sepet,
            "toplam_indirim": toplam_indirim,
            "trendyol_oran": trendyol_oran,
        }

        result = calculate_all(input_files, counter_files=counter_files, karsilamali_config=karsilamali_config, output_dir=OUTPUT_DIR)
        if result.get("success"):
            result["uploads"] = load_upload_status(UPLOAD_DIR, INPUT_MANIFEST)
            # F_HESAP kaydı
            try:
                pd.DataFrame(result["results"]).to_excel(F_HESAP, index=False)
            except Exception: pass
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
    if not isinstance(selections, dict) or any(
        not isinstance(value, str) or value not in VALID_SELECTIONS
        for value in selections.values()
    ):
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
        raw_rows = pd.read_excel(F_HESAP).to_dict(orient="records")
    except Exception:
        app.logger.exception("Hesap sonucu okunamadı")
        return jsonify({
            "success": False,
            "message": "Hesap sonucu okunamadı; girdileri yeniden hesaplayın.",
        }), 500

    table_data = []
    for row in raw_rows:
        barcode = str(row.get("Barkod", "")).strip()
        selection = selections.get(
            barcode, row.get("İlk Kampanya Seçimi", "Hiçbiri")
        )
        if not campaign_selection_is_applicable(
            selection, row.get("Uygulanabilir Kampanyalar")
        ):
            return jsonify({
                "success": False,
                "message": "Seçilen kampanya ürün için uygulanabilir değil.",
            }), 400
        row["userSelection"] = selection
        table_data.append(row)

    F_AVAN = input_files.get("advantage")
    F_FLAS = input_files.get("flash")
    F_PLUS = input_files.get("plus")
    F_PLUS_EK = input_files.get("plus_extra")
    F_KARS = input_files.get("counter")
    target_inputs = {
        "Avantajlı": F_AVAN,
        "Flaş": F_FLAS,
        "Plus": F_PLUS,
        "Plus Ek İndirim": F_PLUS_EK,
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
            
    # 1. Process Avantajlı
    if target_type in ['Hepsi', 'Avantajlı'] and F_AVAN:
        try:
            wb_av = openpyxl.load_workbook(F_AVAN)
            ws_av = wb_av.active
            header_av = [ws_av.cell(1, c).value for c in range(1, ws_av.max_column + 1)]
            b_idx_av = header_av.index('BARKOD') + 1 if 'BARKOD' in header_av else None
            tsf_idx_av = header_av.index('YENİ TSF (FİYAT GÜNCELLE)') + 1 if 'YENİ TSF (FİYAT GÜNCELLE)' in header_av else None
        
            if b_idx_av and tsf_idx_av:
                keep_rows = []
                for r in range(2, ws_av.max_row + 1):
                    b_val = ws_av.cell(r, b_idx_av).value
                    if not b_val: continue
                    b_val_str = str(b_val).strip()
                    sel = selections.get(b_val_str, "Hiçbiri")
                    row_info = row_by_barcode.get(b_val_str, {})
                    should_keep = sel == "Avantajlı"

                    if should_keep:
                        selected_price = row_info.get('Avantajlı Ürün Fiyatı (YENİ TSF) (TL)')
                        if selected_price is not None and not pd.isna(selected_price):
                            ws_av.cell(r, tsf_idx_av).value = float(selected_price)
                        keep_rows.append(r)
                
                if keep_rows:
                    safe_keep_rows(ws_av, keep_rows)
                    out_name = os.path.join(run_output_dir, "Avantajlı Ürün.xlsx")
                    shrink_data_validations(ws_av)
                    wb_av.save(out_name)
                    fix_xlsx_for_trendyol(out_name)
                    generated_files.append(os.path.join(timestamp_folder, "Avantajlı Ürün.xlsx"))
        except Exception:
            return processing_error("Avantajlı dosya")

    # 2. Process Flaş (Grouped by Date)
    if target_type in ['Hepsi', 'Flaş'] and F_FLAS:
        try:
            wb_fl = openpyxl.load_workbook(F_FLAS)
            ws_fl = wb_fl.active
            header_fl = [ws_fl.cell(1, c).value for c in range(1, ws_fl.max_column + 1)]
            b_idx_fl = header_fl.index('Barkod') + 1 if 'Barkod' in header_fl else None
            guncel_fiyat_idx = header_fl.index('Güncellenecek Fiyat') + 1 if 'Güncellenecek Fiyat' in header_fl else None
            baslangic_idx = header_fl.index('24 Saat Flaş Başlangıç Tarihi') + 1 if '24 Saat Flaş Başlangıç Tarihi' in header_fl else None
        
            if b_idx_fl and guncel_fiyat_idx and baslangic_idx:
                date_groups = {} # {date_str: [row_indices]}

                for r in range(2, ws_fl.max_row + 1):
                    b_val = ws_fl.cell(r, b_idx_fl).value
                    if not b_val: continue
                
                    b_val_str = str(b_val).strip()
                    sel = selections.get(b_val_str, "Hiçbiri")
                    should_keep = sel == "Flaş"
                
                    if should_keep:
                        date_val = str(ws_fl.cell(r, baslangic_idx).value or "").strip()
                        date_key = date_val.split()[0].replace('/', '_').replace('-', '_')
                        if not date_key or date_key == 'None':
                            date_key = "Genel"
                        if date_key not in date_groups:
                            date_groups[date_key] = []
                        date_groups[date_key].append(r)
            
                for date_key, keep_rows in date_groups.items():
                    if keep_rows:
                        wb_copy = openpyxl.load_workbook(F_FLAS)
                        ws_copy = wb_copy.active
                        for r in keep_rows:
                            ws_copy.cell(r, guncel_fiyat_idx).value = "24 Saat"
                        safe_keep_rows(ws_copy, keep_rows)
                        out_name = os.path.join(run_output_dir, f"Flas_Urun_{date_key}.xlsx")
                        shrink_data_validations(ws_copy)
                        wb_copy.save(out_name)
                        fix_xlsx_for_trendyol(out_name)
                        generated_files.append(os.path.join(timestamp_folder, f"Flas_Urun_{date_key}.xlsx"))
                
        except Exception:
            return processing_error("Flaş dosya")

    # 3. Process Plus
    if target_type in ['Hepsi', 'Plus']:
        if F_PLUS:
            try:
                wb_plus = openpyxl.load_workbook(F_PLUS)
                ws_plus = wb_plus.active
                header_plus = [ws_plus.cell(1, c).value for c in range(1, ws_plus.max_column + 1)]
                b_idx_plus = header_plus.index('Barkod') + 1 if 'Barkod' in header_plus else None
                fiyat_secim_idx = header_plus.index('Plus Fiyat Seçimi') + 1 if 'Plus Fiyat Seçimi' in header_plus else None
                tarife_secim_idx = header_plus.index('Tarife Seçimi') + 1 if 'Tarife Seçimi' in header_plus else None
                ust_limit_idx = header_plus.index('Plus Fiyat Üst Limiti') + 1 if 'Plus Fiyat Üst Limiti' in header_plus else None
            
                gun_sayisi = 7
                for col in header_plus:
                    if col and "Tarih Aralığı" in str(col):
                        import re
                        match = re.search(r'\((\d+)\s*Gün\)', str(col))
                        if match:
                            gun_sayisi = int(match.group(1))
                        break
                    
                if b_idx_plus and fiyat_secim_idx and tarife_secim_idx and ust_limit_idx:
                    keep_rows = []
                    for r in range(2, ws_plus.max_row + 1):
                        b_val = ws_plus.cell(r, b_idx_plus).value
                        if not b_val: continue
                        b_val_str = str(b_val).strip()
                        sel = selections.get(b_val_str, "Hiçbiri")
                        should_keep = sel == "Plus"

                        if should_keep:
                            ust_lim = ws_plus.cell(r, ust_limit_idx).value
                            ws_plus.cell(r, fiyat_secim_idx).value = ust_lim
                            ws_plus.cell(r, tarife_secim_idx).value = f"{gun_sayisi} Günlük Fiyat"
                            keep_rows.append(r)
                
                    if keep_rows:
                        safe_keep_rows(ws_plus, keep_rows)
                        out_name = os.path.join(run_output_dir, "Plus_Urun.xlsx")
                        shrink_data_validations(ws_plus)
                        wb_plus.save(out_name)
                        fix_xlsx_for_trendyol(out_name)
                        generated_files.append(os.path.join(timestamp_folder, "Plus_Urun.xlsx"))
            except Exception:
                return processing_error("Plus dosya")

    # 4. Process Plus Ek İndirim
    if target_type in ['Hepsi', 'Plus Ek İndirim']:
        if F_PLUS_EK:
            try:
                for rate in [5, 10, 20]:
                    wb_pe = openpyxl.load_workbook(F_PLUS_EK)
                    ws_pe = wb_pe.active
                    header_pe = [ws_pe.cell(1, c).value for c in range(1, ws_pe.max_column + 1)]
                    b_idx_pe = header_pe.index('Barkod') + 1 if 'Barkod' in header_pe else None
                    fiyat_idx_pe = header_pe.index('Kampanyalı Satış Fiyatı') + 1 if 'Kampanyalı Satış Fiyatı' in header_pe else None
                    max_fiyat_idx_pe = header_pe.index('Maksimum Girebileceğin Fiyat') + 1 if 'Maksimum Girebileceğin Fiyat' in header_pe else None
                
                    if b_idx_pe and fiyat_idx_pe and max_fiyat_idx_pe:
                        target_sel = f"Plus Ek İndirim %{rate}"
                        keep_rows = []
                        for r in range(2, ws_pe.max_row + 1):
                            b_val = ws_pe.cell(r, b_idx_pe).value
                            if not b_val: continue

                            b_val_str = str(b_val).strip()
                            sel = selections.get(b_val_str, "Hiçbiri")
                            should_keep = sel == target_sel

                            if should_keep:
                                campaign_price = discounted_price(
                                    ws_pe.cell(r, max_fiyat_idx_pe).value, rate
                                )
                                if campaign_price is not None:
                                    ws_pe.cell(r, fiyat_idx_pe).value = campaign_price
                                keep_rows.append(r)

                        if keep_rows:
                            safe_keep_rows(ws_pe, keep_rows)
                            out_name = os.path.join(run_output_dir, f"Plus_Ek_Indirim_{rate}.xlsx")
                            shrink_data_validations(ws_pe)
                            wb_pe.save(out_name)
                            fix_xlsx_for_trendyol(out_name)
                            generated_files.append(os.path.join(timestamp_folder, f"Plus_Ek_Indirim_{rate}.xlsx"))
            except Exception:
                return processing_error("Plus Ek İndirim dosyası")

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
                            sel = selections.get(b_val_str, "Hiçbiri")
                            should_keep = (sel == c_label) or (target_type == "Karşılamalı Kampanya" and sel.startswith("Karşılamalı"))

                            if should_keep:
                                max_fiyat_val = ws_kars.cell(r, max_fiyat_idx_kars).value
                                if max_fiyat_val:
                                    ws_kars.cell(r, fiyat_idx_kars).value = max_fiyat_val
                                keep_rows.append(r)
                                
                        if keep_rows:
                            safe_keep_rows(ws_kars, keep_rows)
                            safe_label = re.sub(r'[^\w\-_]', '_', c_label)
                            out_filename = f"Karsilamali_{safe_label}.xlsx"
                            out_name = os.path.join(run_output_dir, out_filename)
                            shrink_data_validations(ws_kars)
                            wb_kars.save(out_name)
                            fix_xlsx_for_trendyol(out_name)
                            generated_files.append(os.path.join(timestamp_folder, out_filename))
        except Exception:
            return processing_error("Karşılamalı Kampanya dosyası")
    # Ekstra Rapor ve Uygulanmayanlar Excel Çıktıları
    try:
        df_all = pd.DataFrame(table_data)
        if not df_all.empty:
            if 'checked' in df_all.columns:
                df_all = df_all.drop(columns=['checked'])
            if 'userSelection' not in df_all.columns:
                df_all['userSelection'] = 'Hiçbiri'
            
            # 1. Uygulanmayanlar
            df_unapplied = df_all[df_all['userSelection'] == 'Hiçbiri'].copy()
            df_unapplied.rename(columns={'userSelection': 'Uygulanan Kampanya Seçimi'}, inplace=True)
            out_unapplied = os.path.join(run_output_dir, "Uygulanmayan_Urunler_Raporu.xlsx")
            df_unapplied.to_excel(out_unapplied, index=False)
            fix_xlsx_for_trendyol(out_unapplied)
            generated_files.append(os.path.join(timestamp_folder, "Uygulanmayan_Urunler_Raporu.xlsx"))
            
            # 2. Tüm Rapor
            df_all.rename(columns={'userSelection': 'Uygulanan Kampanya Seçimi'}, inplace=True)
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
                generate_fiyat_farki_raporu(run_output_dir)
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
