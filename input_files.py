import json
import math
import os
import re
import tempfile
import threading
import zipfile
from datetime import datetime
from pathlib import Path

import pandas as pd


INPUT_SPECS = {
    "discount": {
        "label": "İndirim Uygulanabilecek Ürünler",
        "required": True,
        "filename": "discount.xlsx",
        "columns": {"BARKOD", "Eski Fiyat", "YENİ Fiyat"},
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

# ponytail: This app has one manifest and one process; use an inter-process lock
# if it is ever deployed with multiple worker processes.
_MANIFEST_WRITE_LOCK = threading.Lock()


class InputValidationError(ValueError):
    pass


MAIN_CAMPAIGN_PRIORITY = ("Avantajlı", "Flaş", "Plus")

_PLUS_PERIOD_COLUMN = re.compile(r"Tarih Aralığı \((\d+) Gün\)", re.IGNORECASE)
_PLUS_OFFER_COLUMN = re.compile(r"Plus Komisyon Teklifi(?:\.\d+)?", re.IGNORECASE)


def find_plus_period_columns(columns):
    column_keys = list(columns)
    periods = []
    for date_position, date_column in enumerate(column_keys[:-1]):
        match = _PLUS_PERIOD_COLUMN.fullmatch(str(date_column).strip())
        offer_position = date_position + 1
        offer_column = column_keys[offer_position]
        if not match or not _PLUS_OFFER_COLUMN.fullmatch(str(offer_column).strip()):
            continue
        periods.append({
            "days": int(match.group(1)),
            "date_position": date_position,
            "offer_position": offer_position,
            "date_column": date_column,
            "offer_column": offer_column,
        })
    return periods


def _plus_cell_value(row, column):
    value = row.get(column) if hasattr(row, "get") else None
    return None if value is None or pd.isna(value) or str(value).strip() == "" else value


def choose_plus_tariff_label(row, periods, eligible_days):
    periods = list(periods)
    eligible_days = set(eligible_days)
    eligible = [period for period in periods if period["days"] in eligible_days]
    if not eligible:
        return None

    if len(eligible) == len(periods):
        days = sum(period["days"] for period in periods)
        return _plus_cell_value(row, f"{days} Gün Tarih Aralığı") or f"{days} Günlük Fiyat"

    if len(eligible) == 1:
        period = eligible[0]
        days = period["days"]
        helper_value = _plus_cell_value(row, f"{days} Gün Tarih Aralığı")
        if helper_value is not None:
            return helper_value
        date_value = _plus_cell_value(row, period["date_column"])
        return f"{days} Günlük Fiyat ({date_value})" if date_value is not None else None

    return None


def normalize_recommendation_rule(rule=None):
    if rule is None:
        return {"enabled": True, "priority": list(MAIN_CAMPAIGN_PRIORITY)}
    if not isinstance(rule, dict) or set(rule) != {"enabled", "priority"}:
        raise InputValidationError("Öneri kuralı enabled ve priority alanlarını içermelidir.")
    if not isinstance(rule["enabled"], bool):
        raise InputValidationError("Öneri kuralı enabled alanı doğru/yanlış olmalıdır.")
    priority = rule["priority"]
    if (
        not isinstance(priority, list)
        or len(priority) != len(MAIN_CAMPAIGN_PRIORITY)
        or any(not isinstance(name, str) for name in priority)
        or len(set(priority)) != len(MAIN_CAMPAIGN_PRIORITY)
        or set(priority) != set(MAIN_CAMPAIGN_PRIORITY)
    ):
        raise InputValidationError(
            "Öneri önceliği Avantajlı, Flaş ve Plus kampanyalarını birer kez içermelidir."
        )
    return {"enabled": rule["enabled"], "priority": list(priority)}


CAMPAIGN_NAMES = {
    "counter": "Karşılamalı Kampanya",
    "plus_extra": "Plus Ek İndirim",
    "coupon": "Kupon",
}


def _campaign_number(value, field, campaign_name):
    if value is None or value == "":
        return 0.0
    if isinstance(value, bool):
        raise InputValidationError(f"{campaign_name} {field} sonlu bir sayı olmalıdır.")
    try:
        number = float(str(value).strip().replace(",", "."))
    except (TypeError, ValueError) as exc:
        raise InputValidationError(
            f"{campaign_name} {field} sonlu bir sayı olmalıdır."
        ) from exc
    if not math.isfinite(number):
        raise InputValidationError(f"{campaign_name} {field} sonlu bir sayı olmalıdır.")
    return number


def normalize_campaign_config(config, campaign_type):
    campaign_name = CAMPAIGN_NAMES.get(campaign_type)
    if campaign_name is None or not isinstance(config, dict):
        raise InputValidationError("Kampanya yapılandırması geçersiz.")

    legacy_plus = (
        campaign_type == "plus_extra"
        and "discount_amount" not in config
        and "rate" in config
    )
    discount_type = config.get("discount_type") or config.get("discount_unit")
    if discount_type is None:
        discount_type = "%" if campaign_type == "plus_extra" else "TL"
    discount_type = str(discount_type).strip()
    if discount_type not in {"TL", "%"}:
        raise InputValidationError(
            f"{campaign_name} indirim tipi yalnızca TL veya % olabilir."
        )

    min_price = _campaign_number(config.get("min_price", 0), "alt limiti", campaign_name)
    discount_amount = _campaign_number(
        config.get("rate", 0) if legacy_plus else config.get("discount_amount", 0),
        "indirim tutarı",
        campaign_name,
    )
    trendyol_percent = _campaign_number(
        config.get("trendyol_percent", 0), "Trendyol katkı oranı", campaign_name
    )
    if min_price < 0:
        raise InputValidationError(f"{campaign_name} alt limiti negatif olamaz.")
    if discount_amount < 0:
        raise InputValidationError(f"{campaign_name} indirim tutarı negatif olamaz.")
    if not 0 <= trendyol_percent <= 100:
        raise InputValidationError(
            f"{campaign_name} Trendyol katkı oranı 0 ile 100 arasında olmalıdır."
        )
    if discount_type == "%" and discount_amount > 100:
        raise InputValidationError(
            f"{campaign_name} yüzde indirimi 100'ü aşamaz."
        )

    normalized = {
        **config,
        "min_price": min_price,
        "discount_amount": discount_amount,
        "discount_type": discount_type,
        "trendyol_percent": trendyol_percent,
    }
    if campaign_type == "plus_extra" and discount_type == "%":
        normalized["rate"] = discount_amount
    return normalized


def normalize_campaign_configs(configs, campaign_type):
    if not isinstance(configs, list):
        raise InputValidationError("Kampanya yapılandırmaları liste olmalıdır.")
    return [normalize_campaign_config(config, campaign_type) for config in configs]


def _format_campaign_number(value):
    number = float(value)
    return str(int(number)) if number.is_integer() else str(number)


def build_campaign_label(config, campaign_type, index=0):
    config = normalize_campaign_config(config, campaign_type)
    explicit_label = config.get("label")
    if isinstance(explicit_label, str) and explicit_label.strip():
        return explicit_label.strip()

    min_price = _format_campaign_number(config["min_price"])
    discount = _format_campaign_number(config["discount_amount"])
    trendyol = _format_campaign_number(config["trendyol_percent"])
    discount_part = (
        f"%{discount} İndirim"
        if config["discount_type"] == "%"
        else f"{discount} TL İndirim"
    )
    trendyol_part = f"%{trendyol} Trendyol Karşılamalı"

    if campaign_type == "plus_extra":
        if (
            config["discount_type"] == "%"
            and config["min_price"] == 0
            and config["trendyol_percent"] == 0
        ):
            return f"Plus Ek İndirim %{discount}"
        return (
            f"Plus Ek İndirim ({min_price} TL Üzeri / {discount_part} / "
            f"{trendyol_part})"
        )

    if campaign_type == "counter":
        if (
            config["min_price"] == 0
            and config["discount_amount"] == 0
            and config["trendyol_percent"] == 0
        ):
            return f"Karşılamalı #{index + 1}"
        parts = [f"{min_price} TL Üzeri", discount_part]
        if config["trendyol_percent"]:
            parts.append(trendyol_part)
        return f"Karşılamalı ({' / '.join(parts)})"

    coupon_discount = (
        f"%{discount}" if config["discount_type"] == "%" else f"{discount} TL"
    )
    label = (
        f"{min_price} TL Üzerine {coupon_discount} Kupon - "
        "Trendyol Plus Müşterilerine Özel"
    )
    return (
        f"{label} (%{trendyol} Trendyol Karşılamalı)"
        if config["trendyol_percent"]
        else label
    )


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


def _write_manifest_atomic(manifest_path, manifest):
    manifest_path = Path(manifest_path)
    manifest_path.parent.mkdir(parents=True, exist_ok=True)
    temp_manifest_path = None
    try:
        with tempfile.NamedTemporaryFile(
            mode="w", encoding="utf-8", dir=manifest_path.parent, delete=False
        ) as temp_manifest:
            json.dump(manifest, temp_manifest, ensure_ascii=False, indent=2)
            temp_manifest_path = Path(temp_manifest.name)
        os.replace(temp_manifest_path, manifest_path)
    finally:
        if temp_manifest_path and temp_manifest_path.exists():
            temp_manifest_path.unlink()


def _mutate_manifest_atomic(manifest_path, transform):
    with _MANIFEST_WRITE_LOCK:
        manifest = _read_manifest(manifest_path)
        updated = transform(manifest)
        if updated != manifest:
            _write_manifest_atomic(manifest_path, updated)
        return updated


def load_recommendation_rule(manifest_path):
    manifest = _read_manifest(manifest_path)
    return normalize_recommendation_rule(manifest.get("recommendation_rule"))


def save_recommendation_rule(manifest_path, rule):
    normalized = normalize_recommendation_rule(rule)
    _mutate_manifest_atomic(
        manifest_path,
        lambda manifest: {**manifest, "recommendation_rule": normalized},
    )
    return normalized


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
        _mutate_manifest_atomic(
            manifest_path,
            lambda manifest: {**manifest, "files": manifest_files},
        )
        return saved
    finally:
        for temp_path in staged.values():
            if temp_path.exists():
                temp_path.unlink()


import re

def parse_counter_filename(filename):
    return _parse_discount_filename(filename, "indirim")


def _save_campaign_configs(manifest_path, key, campaign_type, configs):
    normalized = normalize_campaign_configs(configs, campaign_type)
    _mutate_manifest_atomic(
        manifest_path,
        lambda manifest: {**manifest, key: normalized},
    )


def _load_campaign_configs(manifest_path, key, campaign_type):
    manifest_path = Path(manifest_path)
    manifest = _read_manifest(manifest_path)
    configs = manifest.get(key, [])
    normalized_configs = normalize_campaign_configs(configs, campaign_type)
    valid_configs = []
    changed = normalized_configs != configs
    for item in normalized_configs:
        path = item.get("path") or item.get("stored_path")
        if path and os.path.exists(path):
            valid_configs.append(item)
        else:
            changed = True
    if changed:
        _save_campaign_configs(manifest_path, key, campaign_type, valid_configs)
    return valid_configs


def save_counter_configs(manifest_path, counter_configs):
    _save_campaign_configs(
        manifest_path, "counter_configs", "counter", counter_configs
    )


def load_counter_configs(manifest_path):
    return _load_campaign_configs(manifest_path, "counter_configs", "counter")


def parse_plus_extra_filename(filename):
    return _parse_discount_filename(filename, "indirim", bare_percent=True)


def save_plus_extra_configs(manifest_path, plus_extra_configs):
    _save_campaign_configs(
        manifest_path, "plus_extra_configs", "plus_extra", plus_extra_configs
    )


def load_plus_extra_configs(manifest_path):
    return _load_campaign_configs(
        manifest_path, "plus_extra_configs", "plus_extra"
    )


def parse_coupon_filename(filename):
    return _parse_discount_filename(filename, "kupon")


def _parse_discount_filename(filename, discount_word, bare_percent=False):
    fn_str = str(filename)
    min_p = 0.0
    number_value = r'\d+(?:[\.,]\d+)?'
    number = rf'({number_value})'
    separator = r'[\s_-]*'
    m_min = re.search(
        number + separator + r'tl' + separator + r'(?:uzeri(?:ne)?|üzeri(?:ne)?)',
        fn_str,
        re.IGNORECASE,
    )
    if m_min:
        try: min_p = float(m_min.group(1).replace(',', '.'))
        except (ValueError, TypeError): pass

    disc = 0.0
    disc_type = 'TL'
    m_disc_pct = re.search(
        rf'(?:(?:%|y[uü]zde){separator}({number_value})|'
        rf'({number_value}){separator}%)'
        + separator
        + rf'(?:ek{separator})?{discount_word}',
        fn_str,
        re.IGNORECASE,
    )
    if m_disc_pct:
        disc_type = '%'
        raw_discount = m_disc_pct.group(1) or m_disc_pct.group(2)
        try: disc = float(raw_discount.replace(',', '.'))
        except (ValueError, TypeError): pass
    else:
        m_disc_bare = re.search(
            rf'ek{separator}({number_value}){separator}{discount_word}',
            fn_str,
            re.IGNORECASE,
        ) if bare_percent else None
        if m_disc_bare:
            disc_type = '%'
            try: disc = float(m_disc_bare.group(1).replace(',', '.'))
            except (ValueError, TypeError): pass
        else:
            m_disc_tl = re.search(
                number + separator + r'tl' + separator + discount_word,
                fn_str,
                re.IGNORECASE,
            )
            if m_disc_tl:
                try: disc = float(m_disc_tl.group(1).replace(',', '.'))
                except (ValueError, TypeError): pass

    trendyol_p = 0.0
    m_tr = re.search(
        r'%?' + separator + number + separator + r'%?' + separator + r'trendyol',
        fn_str,
        re.IGNORECASE,
    )
    if m_tr:
        try: trendyol_p = float(m_tr.group(1).replace(',', '.'))
        except (ValueError, TypeError): pass

    return min_p, disc, trendyol_p, disc_type


def save_coupon_configs(manifest_path, coupon_configs):
    _save_campaign_configs(manifest_path, "coupon_configs", "coupon", coupon_configs)


def load_coupon_configs(manifest_path):
    return _load_campaign_configs(manifest_path, "coupon_configs", "coupon")


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

    def update_expiries(manifest):
        files = manifest.get("files", {})
        updated_files = {
            key: (
                {**item, "expiry_date": str(expiries_dict[key] or "")}
                if key in expiries_dict
                else item
            )
            for key, item in files.items()
        }
        return {**manifest, "files": updated_files}

    _mutate_manifest_atomic(manifest_path, update_expiries)


def load_user_selections(manifest_path):
    manifest = _read_manifest(manifest_path)
    return manifest.get("user_selections", {})


def save_user_selections(manifest_path, selections_dict):
    if not isinstance(selections_dict, dict):
        return
    selections = dict(selections_dict)
    _mutate_manifest_atomic(
        manifest_path,
        lambda manifest: {**manifest, "user_selections": selections},
    )

