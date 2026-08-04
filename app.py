import os
import glob
import openpyxl
import pandas as pd
from datetime import datetime, timedelta
from flask import Flask, render_template, jsonify, request, send_from_directory
from xlsx_postprocess import fix_xlsx_for_trendyol

app = Flask(__name__)

BASE_DIR = r"c:\Users\Tasarımcı\Desktop\trendyol"
INPUT_DIR = os.path.join(BASE_DIR, "Girdiler")
OUTPUT_DIR = os.path.join(BASE_DIR, "Çıktılar")

os.makedirs(INPUT_DIR, exist_ok=True)
os.makedirs(OUTPUT_DIR, exist_ok=True)

F_HESAP = os.path.join(OUTPUT_DIR, "Hesaplanmis_Komisyon_Sonuclari.xlsx")

@app.route("/api/download/<folder>/<filename>")
def download_file(folder, filename):
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

def find_files():
    excel_files = glob.glob(os.path.join(INPUT_DIR, '*.xlsx'))
    f_avan = None
    f_flas = None
    f_plus = None
    f_plus_ek = None
    for f in excel_files:
        if "Uygulanmis" in f or "Hesaplanmis" in f: continue
        try:
            cols = pd.read_excel(f, nrows=0).columns.tolist()
            if '1 YILDIZ ÜST FİYAT' in cols and 'YENİ TSF (FİYAT GÜNCELLE)' in cols:
                f_avan = f
            elif '24 Saat Fiyat' in cols and 'Kampanyalı Ürün' in cols:
                f_flas = f
            elif 'Plus Fiyat Üst Limiti' in cols and 'Plus Komisyon Teklifi' in cols:
                f_plus = f
            elif 'Maksimum Girebileceğin Fiyat' in cols and 'Kampanyalı Satış Fiyatı' in cols:
                f_plus_ek = f
        except: pass
    return f_avan, f_flas, f_plus, f_plus_ek


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
    return render_template("index.html")

@app.route("/api/data")
def get_data():
    if not os.path.exists(F_HESAP):
        return jsonify({"needs_calculation": True, "message": "Lütfen önce 'Verileri Güncelle' butonuna basarak hesaplamaları başlatın."}), 200
        
    df = pd.read_excel(F_HESAP)
    return app.response_class(
        response=df.to_json(orient="records"),
        status=200,
        mimetype='application/json'
    )

@app.route("/api/save_karsilamali_config", methods=["POST"])
def save_karsilamali_config():
    try:
        import json
        req_json = request.get_json(silent=True) or {}
        min_sepet = request.form.get("min_sepet") or req_json.get("min_sepet", 0)
        toplam_indirim = request.form.get("toplam_indirim") or req_json.get("toplam_indirim", 0)
        trendyol_oran = request.form.get("trendyol_oran") or req_json.get("trendyol_oran", 0)
        
        file_path = None
        if 'file' in request.files:
            f = request.files['file']
            if f.filename:
                file_path = os.path.join(INPUT_DIR, f.filename)
                f.save(file_path)
                
        if not file_path:
            config_json_path = os.path.join(INPUT_DIR, 'karsilamali_config.json')
            if os.path.exists(config_json_path):
                with open(config_json_path, 'r', encoding='utf-8') as cf:
                    old_cfg = json.load(cf)
                    file_path = old_cfg.get('file_path')

        config = {
            "min_sepet": float(min_sepet) if min_sepet else 0,
            "toplam_indirim": float(toplam_indirim) if toplam_indirim else 0,
            "trendyol_oran": float(trendyol_oran) if trendyol_oran else 0,
            "file_path": file_path
        }
        
        with open(os.path.join(INPUT_DIR, 'karsilamali_config.json'), 'w', encoding='utf-8') as cf:
            json.dump(config, cf, ensure_ascii=False, indent=2)
            
        return jsonify({"success": True, "config": config})
    except Exception as e:
        return jsonify({"success": False, "message": str(e)}), 500

@app.route("/api/calculate", methods=["POST"])
def calculate():
    try:
        from komisyon_hesaplayici import calculate_all
        req_data = request.get_json(silent=True) or {}
        karsilamali_config = req_data.get("karsilamali_config")
        result = calculate_all(karsilamali_config)
        return jsonify(result), (200 if result.get("success") else 500)
    except Exception as e:
        return jsonify({"success": False, "message": f"Hesaplama sırasında hata: {str(e)}"}), 500

@app.route("/api/apply", methods=["POST"])
def apply_campaign():
    data = request.get_json(silent=True) or {}
    selections = data.get("selections", {})
    table_data = data.get("tableData", [])
    if not table_data and os.path.exists(F_HESAP):
        try:
            df_h = pd.read_excel(F_HESAP)
            raw_rows = df_h.to_dict(orient="records")
            table_data = []
            for r in raw_rows:
                b_val = str(r.get("Barkod", "")).strip()
                r["userSelection"] = selections.get(b_val, "Hiçbiri")
                table_data.append(r)
        except Exception as e:
            print("F_HESAP okuma hatası:", e)
            table_data = []

    if not selections and table_data:
        selections = {str(row.get("Barkod", "")).strip(): row.get("userSelection", "Hiçbiri") for row in table_data if row.get("Barkod")}

    target_type = data.get("target_type", "Hepsi")
    
    F_AVAN, F_FLAS, F_PLUS, F_PLUS_EK = find_files()
    if not F_AVAN or not F_FLAS:
        return jsonify({"success": False, "message": "Avantajlı veya Flaş ürün şablonu klasörde bulunamadı!"}), 400
        
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
    if target_type in ['Hepsi', 'Avantajlı']:
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
                    rec = row_info.get("Hangisi Daha Karlı?", "")

                    should_keep = False
                    if sel == "Avantajlı":
                        should_keep = True
                    elif sel == "Hiçbiri" and target_type in ["Hepsi", "Avantajlı"]:
                        if rec in ["Avantajlı Ürün", "Sadece Avantajlı Var"]:
                            should_keep = True

                    if should_keep:
                        keep_rows.append(r)
                
                if keep_rows:
                    safe_keep_rows(ws_av, keep_rows)
                    out_name = os.path.join(run_output_dir, "Avantajlı Ürün.xlsx")
                    shrink_data_validations(ws_av)
                    wb_av.save(out_name)
                    fix_xlsx_for_trendyol(out_name)
                    generated_files.append(os.path.join(timestamp_folder, "Avantajlı Ürün.xlsx"))
        except Exception as e:
            return jsonify({"success": False, "message": f"Avantajlı dosya işlenirken hata: {str(e)}"}), 500

    # 2. Process Flaş (Grouped by Date)
    if target_type in ['Hepsi', 'Flaş']:
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
                    row_info = row_by_barcode.get(b_val_str, {})
                    rec = row_info.get("Hangisi Daha Karlı?", "")

                    should_keep = False
                    if sel == "Flaş":
                        should_keep = True
                    elif sel == "Hiçbiri" and target_type in ["Hepsi", "Flaş"]:
                        if rec in ["Flaş Ürün", "Sadece Flaş Var"]:
                            should_keep = True
                
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
                
        except Exception as e:
            return jsonify({"success": False, "message": f"Flaş dosya işlenirken hata: {str(e)}"}), 500

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
                        row_info = row_by_barcode.get(b_val_str, {})
                        rec = row_info.get("Hangisi Daha Karlı?", "")

                        should_keep = False
                        if sel == "Plus":
                            should_keep = True
                        elif sel == "Hiçbiri" and target_type in ["Hepsi", "Plus"]:
                            if rec in ["Plus Ürün", "Sadece Plus Var"]:
                                should_keep = True

                        if should_keep:
                            ust_lim = ws_plus.cell(r, ust_limit_idx).value
                            ws_plus.cell(r, fiyat_secim_idx).value = ust_lim
                            ws_plus.cell(r, tarife_secim_idx).value = f"{gun_sayisi} Günlük Fiyat"
                            keep_rows.append(r)
                
                    if keep_rows:
                        safe_keep_rows(ws_plus, keep_rows)
                        out_name = os.path.join(run_output_dir, "Plus_Urun_Uygulanmis.xlsx")
                        shrink_data_validations(ws_plus)
                        wb_plus.save(out_name)
                        fix_xlsx_for_trendyol(out_name)
                        generated_files.append(os.path.join(timestamp_folder, "Plus_Urun_Uygulanmis.xlsx"))
            except Exception as e:
                return jsonify({"success": False, "message": f"Plus dosya işlenirken hata: {str(e)}"}), 500

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
                            row_info = row_by_barcode.get(b_val_str, {})
                            rec = row_info.get("Hangisi Daha Karlı?", "")
                        
                            should_keep = False
                            if sel == target_sel:
                                should_keep = True
                            elif sel == "Hiçbiri" and target_type in ["Hepsi", "Plus Ek İndirim"]:
                                if rec == target_sel or rec == f"Plus Ek İndirim %{rate} Var":
                                    should_keep = True

                            if should_keep:
                                max_fiyat_val = ws_pe.cell(r, max_fiyat_idx_pe).value
                                if max_fiyat_val:
                                    ws_pe.cell(r, fiyat_idx_pe).value = max_fiyat_val
                                keep_rows.append(r)
                            
                        if keep_rows:
                            safe_keep_rows(ws_pe, keep_rows)
                            out_name = os.path.join(run_output_dir, f"Plus_Ek_Indirim_{rate}_Uygulanmis.xlsx")
                            shrink_data_validations(ws_pe)
                            wb_pe.save(out_name)
                            fix_xlsx_for_trendyol(out_name)
                            generated_files.append(os.path.join(timestamp_folder, f"Plus_Ek_Indirim_{rate}_Uygulanmis.xlsx"))
            except Exception as e:
                return jsonify({"success": False, "message": f"Plus Ek İndirim dosya işlenirken hata: {str(e)}"}), 500

    # 5. Process Karşılamalı Kampanya
    if target_type in ['Hepsi', 'Karşılamalı Kampanya']:
        try:
            f_karsilamali = None
            config_json_path = os.path.join(INPUT_DIR, 'karsilamali_config.json')
            if os.path.exists(config_json_path):
                import json
                with open(config_json_path, 'r', encoding='utf-8') as cf:
                    cfg = json.load(cf)
                    f_karsilamali = cfg.get('file_path')
            
            if not f_karsilamali or not os.path.exists(f_karsilamali):
                for f in glob.glob(os.path.join(INPUT_DIR, '*.xlsx')):
                    if "2000-tl-uzeri" in f or "karsilamali" in f.lower() or "karşılamalı" in f.lower():
                        f_karsilamali = f
                        break
                        
            if f_karsilamali and os.path.exists(f_karsilamali):
                wb_kars = openpyxl.load_workbook(f_karsilamali)
                ws_kars = wb_kars.active
                header_kars = [ws_kars.cell(1, c).value for c in range(1, ws_kars.max_column + 1)]
                b_idx_kars = header_kars.index('Barkod') + 1 if 'Barkod' in header_kars else None
                fiyat_idx_kars = header_kars.index('Kampanyalı Satış Fiyatı') + 1 if 'Kampanyalı Satış Fiyatı' in header_kars else None
                max_fiyat_idx_kars = header_kars.index('Maksimum Girebileceğin Fiyat') + 1 if 'Maksimum Girebileceğin Fiyat' in header_kars else None
                
                if b_idx_kars and fiyat_idx_kars and max_fiyat_idx_kars:
                    keep_rows = []
                    for r in range(2, ws_kars.max_row + 1):
                        b_val = ws_kars.cell(r, b_idx_kars).value
                        if not b_val: continue
                        
                        b_val_str = str(b_val).strip()
                        sel = selections.get(b_val_str, "Hiçbiri")
                        row_info = row_by_barcode.get(b_val_str, {})
                        rec = row_info.get("Hangisi Daha Karlı?", "")

                        should_keep = False
                        if sel == "Karşılamalı Kampanya":
                            should_keep = True
                        elif sel == "Hiçbiri" and target_type in ["Hepsi", "Karşılamalı Kampanya"]:
                            if rec == "Karşılamalı Kampanya":
                                should_keep = True

                        if should_keep:
                            max_fiyat_val = ws_kars.cell(r, max_fiyat_idx_kars).value
                            if max_fiyat_val:
                                ws_kars.cell(r, fiyat_idx_kars).value = max_fiyat_val
                            keep_rows.append(r)
                            
                    if keep_rows:
                        safe_keep_rows(ws_kars, keep_rows)
                        out_name = os.path.join(run_output_dir, "Karsilamali_Indirim_Uygulanmis.xlsx")
                        shrink_data_validations(ws_kars)
                        wb_kars.save(out_name)
                        fix_xlsx_for_trendyol(out_name)
                        generated_files.append(os.path.join(timestamp_folder, "Karsilamali_Indirim_Uygulanmis.xlsx"))
        except Exception as e:
            print("Karşılamalı Kampanya dosya işleme hatası:", e)

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
            
            # 3. Özet Rapor
            summary_rows = []
            for row in table_data:
                barkod = row.get('Barkod')
                if not barkod: continue
                
                guncel_fiyat = row.get('Güncel Ürün Fiyatı (TL)')
                guncel_komisyon = row.get('Güncel Ürün Komisyon (%)')
                guncel_net = row.get('Güncel Ürün Kalan Net (TL)')
                
                user_sel = row.get('userSelection', 'Hiçbiri')
                
                uygulanan_fiyat = None
                uygulanan_komisyon = None
                uygulanan_net = None
                
                if user_sel == 'Avantajlı':
                    uygulanan_fiyat = row.get('Avantajlı Ürün Fiyatı (YENİ TSF) (TL)')
                    uygulanan_komisyon = row.get('Avantajlı Ürün Komisyon (%)')
                    uygulanan_net = row.get('Avantajlı Ürün Kalan Net (TL)')
                elif user_sel == 'Flaş':
                    uygulanan_fiyat = row.get('Flaş Ürün 24 Saat Fiyatı (TL)')
                    uygulanan_komisyon = row.get('Flaş Ürün Komisyon (%)')
                    uygulanan_net = row.get('Flaş Ürün Kalan Net (TL)')
                elif user_sel == 'Plus':
                    uygulanan_fiyat = row.get('Plus Fiyatı (TL)')
                    uygulanan_komisyon = row.get('Plus Komisyon (%)')
                    uygulanan_net = row.get('Plus Net (TL)')
                elif user_sel == 'Plus Ek İndirim %5':
                    uygulanan_fiyat = row.get('Plus Ek Fiyatı %5 (TL)')
                    uygulanan_komisyon = row.get('Plus Ek Komisyon (%)')
                    uygulanan_net = row.get('Plus Ek Net %5 (TL)')
                elif user_sel == 'Plus Ek İndirim %10':
                    uygulanan_fiyat = row.get('Plus Ek Fiyatı %10 (TL)')
                    uygulanan_komisyon = row.get('Plus Ek Komisyon (%)')
                    uygulanan_net = row.get('Plus Ek Net %10 (TL)')
                elif user_sel == 'Plus Ek İndirim %20':
                    uygulanan_fiyat = row.get('Plus Ek Fiyatı %20 (TL)')
                    uygulanan_komisyon = row.get('Plus Ek Komisyon (%)')
                    uygulanan_net = row.get('Plus Ek Net %20 (TL)')
                elif user_sel == 'Karşılamalı Kampanya':
                    uygulanan_fiyat = row.get('Karşılamalı Kampanya Fiyatı (TL)')
                    uygulanan_komisyon = row.get('Karşılamalı Kampanya Komisyon (%)')
                    uygulanan_net = row.get('Karşılamalı Kampanya Kalan Net (TL)')
                else: # Hiçbiri
                    uygulanan_fiyat = guncel_fiyat
                    uygulanan_komisyon = guncel_komisyon
                    uygulanan_net = guncel_net
                    
                # Sadece indirim uygulanabilir ürünlerde (is_indirim=True) değer göster
                is_indirim = row.get('İndirim Uygulanabilir') == 'Evet'
                
                toplam_indirim = row.get('Mevcut İndirim Oranı (%)')
                toplam_indirim_val = None
                toplam_indirim_str = '-'
                if is_indirim and toplam_indirim is not None and str(toplam_indirim).strip() not in ['', '-', 'None']:
                    try:
                        v = float(toplam_indirim)
                        if not pd.isna(v) and v > 0:
                            toplam_indirim_val = v
                            toplam_indirim_str = f"%{v:.2f}"
                    except:
                        pass

                guncel_fiyat_val = float(guncel_fiyat) if (guncel_fiyat and str(guncel_fiyat) != '-') else 0.0
                uygulanan_fiyat_val = float(uygulanan_fiyat) if (uygulanan_fiyat and str(uygulanan_fiyat) != '-') else 0.0

                uygulanan_indirim_val = None
                uygulanan_indirim_str = '-'
                if guncel_fiyat_val > 0 and user_sel != 'Hiçbiri' and uygulanan_fiyat_val > 0:
                    uygulanan_indirim_val = ((guncel_fiyat_val - uygulanan_fiyat_val) / guncel_fiyat_val) * 100.0
                    uygulanan_indirim_str = f"%{uygulanan_indirim_val:.2f}"

                ekstra_indirim_str = '-'
                if is_indirim and toplam_indirim_val is not None:
                    # Kampanya seçilmemişse tüm kalan indirim potansiyeli gösterilir
                    current_applied = uygulanan_indirim_val if (uygulanan_indirim_val is not None) else 0.0
                    rem_diff = round(toplam_indirim_val - current_applied, 2)
                    # Sadece pozitif fark varsa göster (negatif = zaten fazla uygulanmış)
                    if rem_diff > 0:
                        ekstra_indirim_str = f"+%{rem_diff:.2f}"

                summary_rows.append({
                    'Barkod': barkod,
                    'Güncel Fiyat': guncel_fiyat,
                    'Güncel Komisyon': guncel_komisyon,
                    'Güncel Net': guncel_net,
                    'Uygulanan Kampanya': user_sel,
                    'Uygulanan Fiyat': uygulanan_fiyat,
                    'Uygulanan Komisyon': uygulanan_komisyon,
                    'Uygulanan Net': uygulanan_net,
                    'Toplam Uygulanabilecek İndirim (%)': toplam_indirim_str,
                    'Uygulanan İndirim (%)': uygulanan_indirim_str,
                    'Ekstra Uygulanabilir İndirim (%)': ekstra_indirim_str
                })

            df_summary = pd.DataFrame(summary_rows)
            out_summary = os.path.join(run_output_dir, "Kampanya_Ozet_Raporu.xlsx")
            df_summary.to_excel(out_summary, index=False)
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
    except Exception as e:
        print("Rapor oluşturma hatası:", str(e))
        pass
        
    files_str = "\n".join([f"- {f}" for f in generated_files])
    return jsonify({
        "success": True, 
        "message": f"Tarihlere göre {len(generated_files)} dosya başarıyla oluşturuldu!\n\nOluşturulan Dosyalar:\n{files_str}",
        "generated_files": generated_files,
        "timestamp_folder": timestamp_folder
    })

if __name__ == "__main__":
    app.run(port=5114, debug=True)
