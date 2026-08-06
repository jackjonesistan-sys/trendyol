import json
import os
import tempfile
import zipfile
from datetime import datetime
from pathlib import Path

import pandas as pd


INPUT_SPECS = {
    "discount": {
        "label": "İndirim Uygulanabilecek Ürünler",
        "required": False,
        "filename": "discount.xlsx",
        "columns": {"BARKOD", "Eski Fiyat", "YENİ Fiyat", "Durum"},
    },
    "commission": {
        "label": "Ürün Komisyon Tarifeleri",
        "required": True,
        "filename": "commission.xlsx",
        "columns": {
            "BARKOD",
            "1.Fiyat Alt Limit",
            "2.Fiyat Üst Limiti",
            "2.Fiyat Alt Limit",
            "3.Fiyat Üst Limiti",
            "3.Fiyat Alt Limit",
            "4.Fiyat Üst Limiti",
            "1.KOMİSYON",
            "2.KOMİSYON",
            "3.KOMİSYON",
            "4.KOMİSYON",
            "KOMİSYONA ESAS FİYAT",
            "TARİFE GRUBU",
        },
    },
    "current": {
        "label": "Ürün Listesi",
        "required": True,
        "filename": "current.xlsx",
        "columns": {
            "Barkod",
            "Komisyon Oranı",
            "Piyasa Satış Fiyatı (KDV Dahil)",
            "Trendyol'da Satılacak Fiyat (KDV Dahil)",
        },
    },
    "advantage": {
        "label": "Avantajlı Ürün Etiketleri",
        "required": False,
        "filename": "advantage.xlsx",
        "columns": {"BARKOD", "1 YILDIZ ÜST FİYAT", "YENİ TSF (FİYAT GÜNCELLE)"},
    },
    "flash": {
        "label": "Flaş Ürünler",
        "required": False,
        "filename": "flash.xlsx",
        "columns": {
            "Barkod",
            "24 Saat Fiyat",
            "Kampanyalı Ürün",
            "Güncellenecek Fiyat",
            "24 Saat Flaş Başlangıç Tarihi",
        },
    },
    "plus": {
        "label": "Plus Komisyon Tarifeleri",
        "required": False,
        "filename": "plus.xlsx",
        "columns": {
            "Barkod",
            "Plus Fiyat Üst Limiti",
            "Plus Komisyon Teklifi",
            "Plus Fiyat Seçimi",
            "Tarife Seçimi",
        },
    },
    "muhasebe_avantaj": {
        "label": "Muhasebe – Avantajlı Ürün Etiketleri",
        "required": False,
        "filename": "muhasebe_avantaj.xlsx",
        "columns": {"BARKOD", "YENİ TSF (FİYAT GÜNCELLE)"},
    },
    "muhasebe_flas": {
        "label": "Muhasebe – Flaş Ürünler",
        "required": False,
        "filename": "muhasebe_flas.xlsx",
        "columns": {"Barkod", "Senin Belirlediğin Flaş Fiyatı"},
    },
    "muhasebe_plus": {
        "label": "Muhasebe – Plus Komisyon Tarifeleri",
        "required": False,
        "filename": "muhasebe_plus.xlsx",
        "columns": {"Barkod", "Plus Fiyat Üst Limiti"},
    },
}

MAX_UPLOAD_BYTES = 30 * 1024 * 1024
MAX_UNCOMPRESSED_BYTES = 150 * 1024 * 1024
MAX_ZIP_ENTRIES = 5000


class InputValidationError(ValueError):
    pass


def _read_manifest(manifest_path, required=False):
    manifest_path = Path(manifest_path)
    if not manifest_path.exists():
        if required:
            raise InputValidationError("Önce girdi dosyalarını yükleyip hesaplama yapın.")
        return {"files": {}}
    try:
        manifest = json.loads(manifest_path.read_text(encoding="utf-8"))
    except (OSError, ValueError) as exc:
        raise InputValidationError("Yüklenen girdi kaydı okunamadı.") from exc
    if not isinstance(manifest, dict) or not isinstance(manifest.get("files"), dict):
        raise InputValidationError("Yüklenen girdi kaydı okunamadı.")
    return manifest


def _valid_manifest_entries(upload_dir, manifest):
    upload_dir = Path(upload_dir).resolve()
    valid = {}
    for key, item in manifest.get("files", {}).items():
        if (
            key not in INPUT_SPECS
            or not isinstance(item, dict)
            or item.get("stored_name") != INPUT_SPECS[key]["filename"]
        ):
            continue
        path = (upload_dir / item["stored_name"]).resolve()
        if path.parent == upload_dir and path.is_file():
            valid[key] = (path, item.copy())
    return valid


def validate_workbook(kind, path):
    spec = INPUT_SPECS.get(kind)
    if spec is None:
        raise InputValidationError("Bilinmeyen girdi türü.")

    path = Path(path)
    if path.suffix.lower() != ".xlsx":
        raise InputValidationError(f"{spec['label']}: geçerli bir .xlsx dosyası seçin.")
    if path.stat().st_size > MAX_UPLOAD_BYTES:
        raise InputValidationError(f"{spec['label']}: dosya 30 MB sınırını aşıyor.")
    if not zipfile.is_zipfile(path):
        raise InputValidationError(f"{spec['label']}: geçerli bir .xlsx dosyası seçin.")

    with zipfile.ZipFile(path) as workbook_zip:
        entries = workbook_zip.infolist()
        if len(entries) > MAX_ZIP_ENTRIES or sum(entry.file_size for entry in entries) > MAX_UNCOMPRESSED_BYTES:
            raise InputValidationError(f"{spec['label']}: sıkıştırılmış dosya güvenli boyut sınırını aşıyor.")

    try:
        columns = {str(column).strip() for column in pd.read_excel(path, nrows=0).columns}
    except Exception as exc:
        raise InputValidationError(f"{spec['label']}: Excel dosyası okunamadı.") from exc

    missing = sorted(spec["columns"] - columns)
    if missing:
        raise InputValidationError(
            f"{spec['label']}: eksik sütunlar: {', '.join(missing)}"
        )
    return columns


def save_upload_set(files, upload_dir, manifest_path):
    provided = {key: upload for key, upload in files.items() if upload and upload.filename}
    unknown = sorted(set(provided) - set(INPUT_SPECS))
    if unknown:
        raise InputValidationError("Bilinmeyen girdi alanı gönderildi.")

    upload_dir = Path(upload_dir).resolve()
    manifest_path = Path(manifest_path)
    existing = _valid_manifest_entries(
        upload_dir, _read_manifest(manifest_path)
    )
    missing = [
        spec["label"]
        for key, spec in INPUT_SPECS.items()
        if spec["required"] and key not in provided and key not in existing
    ]
    if missing:
        raise InputValidationError(f"Zorunlu girdiler eksik: {', '.join(missing)}")

    upload_dir.mkdir(parents=True, exist_ok=True)
    manifest_path.parent.mkdir(parents=True, exist_ok=True)
    if not provided:
        return {key: str(path) for key, (path, _item) in existing.items()}

    staged = {}
    temp_manifest_path = None

    try:
        for key, upload in provided.items():
            if not str(upload.filename).lower().endswith(".xlsx"):
                raise InputValidationError(f"{INPUT_SPECS[key]['label']}: yalnızca .xlsx yüklenebilir.")
            with tempfile.NamedTemporaryFile(dir=upload_dir, suffix=".xlsx", delete=False) as temp_file:
                temp_path = Path(temp_file.name)
            staged[key] = temp_path
            upload.save(temp_path)
            validate_workbook(key, temp_path)

        saved = {key: str(path) for key, (path, _item) in existing.items()}
        manifest_files = {}
        for key, (path, item) in existing.items():
            uploaded_at = item.get("uploaded_at")
            if not uploaded_at:
                uploaded_at = datetime.fromtimestamp(path.stat().st_mtime).astimezone().isoformat(timespec="seconds")
            entry = {
                "stored_name": path.name,
                "original_name": item.get("original_name") or path.name,
                "uploaded_at": uploaded_at,
            }
            # Varsa expiry_date de koru - hesapla yapilinca kaybolmasin
            if item.get("expiry_date"):
                entry["expiry_date"] = item["expiry_date"]
            manifest_files[key] = entry

        uploaded_at = datetime.now().astimezone().isoformat(timespec="seconds")
        for key, temp_path in staged.items():
            target = upload_dir / INPUT_SPECS[key]["filename"]
            os.replace(temp_path, target)
            saved[key] = str(target)
            original_name = os.path.basename(str(provided[key].filename).replace("\\", "/"))
            manifest_files[key] = {
                "stored_name": target.name,
                "original_name": original_name,
                "uploaded_at": uploaded_at,
            }

        # Mevcut manifest'teki diğer anahtarları (counter_configs, plus_extra_configs vb.) koru
        existing_manifest = _read_manifest(manifest_path)
        existing_manifest["files"] = manifest_files
        with tempfile.NamedTemporaryFile(
            mode="w", encoding="utf-8", dir=manifest_path.parent, delete=False
        ) as temp_manifest:
            json.dump(existing_manifest, temp_manifest, ensure_ascii=False, indent=2)
            temp_manifest_path = Path(temp_manifest.name)
        os.replace(temp_manifest_path, manifest_path)
        return saved
    finally:
        for temp_path in staged.values():
            if temp_path.exists():
                temp_path.unlink()
        if temp_manifest_path and temp_manifest_path.exists():
            temp_manifest_path.unlink()


import re

def parse_counter_filename(filename):
    pattern = r'(\d+)\s*[-_]?tl[-_]?uzeri\s*[-_]?(\d+)\s*[-_]?tl[-_]?indirim\s*[-_]?(\d+)\s*[-_]?trendyol'
    m = re.search(pattern, str(filename), re.IGNORECASE)
    if m:
        try:
            return float(m.group(1)), float(m.group(2)), float(m.group(3))
        except (ValueError, TypeError):
            pass
    return 0.0, 0.0, 0.0


def save_counter_configs(manifest_path, counter_configs):
    manifest_path = Path(manifest_path)
    manifest = _read_manifest(manifest_path)
    manifest["counter_configs"] = counter_configs
    with tempfile.NamedTemporaryFile(
        mode="w", encoding="utf-8", dir=manifest_path.parent, delete=False
    ) as temp_manifest:
        json.dump(manifest, temp_manifest, ensure_ascii=False, indent=2)
        temp_manifest_path = Path(temp_manifest.name)
    os.replace(temp_manifest_path, manifest_path)


def load_counter_configs(manifest_path):
    manifest = _read_manifest(manifest_path)
    return manifest.get("counter_configs", [])


def parse_plus_extra_filename(filename):
    pattern = r'%?\s*(\d+)\s*%?'
    m = re.search(pattern, str(filename))
    if m:
        try:
            return float(m.group(1))
        except (ValueError, TypeError):
            pass
    return 0.0


def save_plus_extra_configs(manifest_path, plus_extra_configs):
    manifest_path = Path(manifest_path)
    manifest = _read_manifest(manifest_path)
    manifest["plus_extra_configs"] = plus_extra_configs
    with tempfile.NamedTemporaryFile(
        mode="w", encoding="utf-8", dir=manifest_path.parent, delete=False
    ) as temp_manifest:
        json.dump(manifest, temp_manifest, ensure_ascii=False, indent=2)
        temp_manifest_path = Path(temp_manifest.name)
    os.replace(temp_manifest_path, manifest_path)


def load_plus_extra_configs(manifest_path):
    manifest = _read_manifest(manifest_path)
    return manifest.get("plus_extra_configs", [])


def load_upload_set(upload_dir, manifest_path):
    upload_dir = Path(upload_dir).resolve()
    entries = _valid_manifest_entries(
        upload_dir, _read_manifest(manifest_path, required=True)
    )
    return {key: str(path) for key, (path, _item) in entries.items()}


def load_upload_status(upload_dir, manifest_path):
    entries = _valid_manifest_entries(
        upload_dir, _read_manifest(manifest_path)
    )
    status = {}
    for key, (path, item) in entries.items():
        uploaded_at = item.get("uploaded_at")
        try:
            uploaded_datetime = datetime.fromisoformat(uploaded_at) if uploaded_at else None
        except (TypeError, ValueError):
            uploaded_datetime = None
        if uploaded_datetime is None:
            uploaded_datetime = datetime.fromtimestamp(path.stat().st_mtime).astimezone()
            uploaded_at = uploaded_datetime.isoformat(timespec="seconds")
        status[key] = {
            "original_name": item.get("original_name") or path.name,
            "uploaded_at": uploaded_at,
            "uploaded_at_display": uploaded_datetime.astimezone().strftime("%d.%m.%Y %H:%M"),
            "expiry_date": item.get("expiry_date", ""),
        }
    return status


def save_single_file_expiries(manifest_path, expiries_dict):
    if not isinstance(expiries_dict, dict):
        return
    manifest_path = Path(manifest_path)
    manifest = _read_manifest(manifest_path)
    files = manifest.get("files", {})
    updated = False
    for k, expiry in expiries_dict.items():
        if k in files:
            files[k]["expiry_date"] = str(expiry or "")
            updated = True
    if updated:
        manifest["files"] = files
        with tempfile.NamedTemporaryFile(
            mode="w", encoding="utf-8", dir=manifest_path.parent, delete=False
        ) as temp_manifest:
            json.dump(manifest, temp_manifest, ensure_ascii=False, indent=2)
            temp_manifest_path = Path(temp_manifest.name)
        os.replace(temp_manifest_path, manifest_path)


def load_user_selections(manifest_path):
    manifest = _read_manifest(manifest_path)
    return manifest.get("user_selections", {})


def save_user_selections(manifest_path, selections_dict):
    if not isinstance(selections_dict, dict):
        return
    manifest_path = Path(manifest_path)
    manifest = _read_manifest(manifest_path)
    manifest["user_selections"] = selections_dict
    with tempfile.NamedTemporaryFile(
        mode="w", encoding="utf-8", dir=manifest_path.parent, delete=False
    ) as temp_manifest:
        json.dump(manifest, temp_manifest, ensure_ascii=False, indent=2)
        temp_manifest_path = Path(temp_manifest.name)
    os.replace(temp_manifest_path, manifest_path)

