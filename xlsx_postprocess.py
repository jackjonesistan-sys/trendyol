# -*- coding: utf-8 -*-
"""
xlsx_postprocess.py

openpyxl'nin oluşturduğu XLSX dosyalarını Trendyol'un backend parser'ı ile
uyumlu hale getiren post-processing modülü.

Sorun: openpyxl her zaman string hücreleri "inlineStr" formatında yazar
(t="inlineStr" + <is><t>...</t></is>) ve sharedStrings.xml oluşturmaz.
Trendyol'un parser'ı ise sadece "shared strings" formatını kabul eder
(t="s" + <v>INDEX</v> ve xl/sharedStrings.xml referansı).

Bu modül:
1. Tüm inlineStr hücrelerini shared string referanslarına dönüştürür
2. xl/sharedStrings.xml dosyası oluşturur
3. [Content_Types].xml ve xl/_rels/workbook.xml.rels günceller
4. Formül cached değerlerini hesaplayıp yazar
5. Geçersiz boş hücre XML'lerini temizler
"""

import os
import re
import zipfile
import xml.etree.ElementTree as ET

# OOXML Namespace sabitleri
NS_SS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
NS_CT = "http://schemas.openxmlformats.org/package/2006/content-types"
NS_RELS = "http://schemas.openxmlformats.org/package/2006/relationships"
NS_RELS_DOC = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"


def _collect_cell_values(sheet_root):
    """Tüm hücrelerin değerlerini topla (formül hesaplaması için)."""
    cell_vals = {}  # (row_num, col_letter) -> value
    for row_el in sheet_root.findall(f".//{{{NS_SS}}}row"):
        r = int(row_el.get("r"))
        for c_el in row_el.findall(f"{{{NS_SS}}}c"):
            ref = c_el.get("r")
            col = re.match(r"([A-Z]+)", ref).group(1)
            t = c_el.get("t", "")

            v_el = c_el.find(f"{{{NS_SS}}}v")
            is_el = c_el.find(f"{{{NS_SS}}}is")

            if t == "inlineStr" and is_el is not None:
                t_el = is_el.find(f"{{{NS_SS}}}t")
                cell_vals[(r, col)] = t_el.text if t_el is not None else ""
            elif v_el is not None and v_el.text:
                try:
                    cell_vals[(r, col)] = float(v_el.text)
                except (ValueError, TypeError):
                    cell_vals[(r, col)] = v_el.text
    return cell_vals


def _compute_formula_cached_value(formula_text, row, cell_vals):
    """Plus şablonundaki bilinen formül kalıpları için cached değer hesapla."""
    if not formula_text:
        return None

    upper_formula = formula_text.upper()

    # IF(...SEARCH("3 Günlük", T{r})..., IF(S{r}<=M{r}, O{r}, "-")...)
    if "SEARCH" in upper_formula:
        search_refs = re.findall(
            r'SEARCH\s*\(\s*"((?:""|[^"])*)"\s*,\s*\$?([A-Z]+)\$?\d+\s*\)',
            formula_text,
            flags=re.IGNORECASE,
        )
        value_refs = re.search(
            r'IF\s*\(\s*\$?([A-Z]+)\$?\d+\s*<=\s*\$?([A-Z]+)\$?\d+\s*,'
            r'\s*\$?([A-Z]+)\$?\d+\s*,\s*"-"\s*\)',
            formula_text,
            flags=re.IGNORECASE,
        )
        if not search_refs or value_refs is None:
            return None

        price_col, upper_limit_col, offer_col = (
            column.upper() for column in value_refs.groups()
        )
        price = cell_vals.get((row, price_col))
        upper_limit = cell_vals.get((row, upper_limit_col))
        offer = cell_vals.get((row, offer_col))
        if price is None or price == "":
            return "-"

        tariff_matches = False
        for phrase, tariff_col in search_refs:
            tariff = cell_vals.get((row, tariff_col.upper()))
            if tariff not in (None, "") and (
                phrase.replace('""', '"').casefold() in str(tariff).casefold()
            ):
                tariff_matches = True
                break
        if not tariff_matches or offer is None or offer == "":
            return "-"

        try:
            return offer if float(price) <= float(upper_limit) else "-"
        except (ValueError, TypeError):
            return "-"

    # IF(ISBLANK(S{r}), "", "Hayır")
    cancel_ref = re.search(
        r'IF\s*\(\s*ISBLANK\s*\(\s*\$?([A-Z]+)\$?\d+\s*\)\s*,'
        r'\s*""\s*,\s*"HAYIR"\s*\)\s*$',
        formula_text,
        flags=re.IGNORECASE,
    )
    if cancel_ref is not None:
        price = cell_vals.get((row, cancel_ref.group(1).upper()))
        return "" if price is None or price == "" else "Hay\u0131r"

    # IF(AND(ISBLANK(...), ...), "", "<tarife>")
    if "ISBLANK" in upper_formula and "AND" in upper_formula:
        blank_refs = re.findall(
            r'ISBLANK\s*\(\s*\$?([A-Z]+)\$?\d+\s*\)',
            formula_text,
            flags=re.IGNORECASE,
        )
        result_text = re.search(r',\s*""\s*,\s*"([^"]*)"\s*\)\s*$', formula_text)
        if blank_refs and result_text is not None:
            all_blank = all(
                cell_vals.get((row, column.upper())) in (None, "")
                for column in blank_refs
            )
            return "" if all_blank else result_text.group(1)

    return None


def fix_xlsx_for_trendyol(filepath):
    """
    openpyxl'nin kaydettiği XLSX dosyasını Trendyol uyumlu formata çevirir.

    - inlineStr → shared strings
    - sharedStrings.xml oluşturur
    - Content_Types ve rels günceller
    - Formül cached değerlerini hesaplar
    - Boş hücreleri temizler
    """

    # ── 1. ZIP içeriğini oku ──
    file_contents = {}
    with zipfile.ZipFile(filepath, "r") as zin:
        for name in zin.namelist():
            file_contents[name] = zin.read(name)

    # Zaten işlenmiş dosyayı tekrar işleme (idempotency)
    if "xl/sharedStrings.xml" in file_contents:
        return

    # ── 2. sheet1.xml'i parse et ──
    ET.register_namespace("", NS_SS)
    sheet_root = ET.fromstring(file_contents["xl/worksheets/sheet1.xml"])

    # ── 3. Hücre değerlerini topla ──
    cell_vals = _collect_cell_values(sheet_root)

    # ── 4. Shared strings tablosu ──
    shared_strings = []
    ss_map = {}  # text -> index

    def get_ss_idx(text):
        text = text or ""
        if text not in ss_map:
            ss_map[text] = len(shared_strings)
            shared_strings.append(text)
        return ss_map[text]

    # ── 5. Hücreleri dönüştür ──
    for row_el in sheet_root.findall(f".//{{{NS_SS}}}row"):
        r = int(row_el.get("r"))

        for c_el in row_el.findall(f"{{{NS_SS}}}c"):
            ref = c_el.get("r")
            t = c_el.get("t", "")
            v_el = c_el.find(f"{{{NS_SS}}}v")
            is_el = c_el.find(f"{{{NS_SS}}}is")
            f_el = c_el.find(f"{{{NS_SS}}}f")

            # ─ Case A: inlineStr → shared string ─
            if t == "inlineStr" and is_el is not None:
                t_el = is_el.find(f"{{{NS_SS}}}t")
                text = t_el.text if t_el is not None else ""
                idx = get_ss_idx(text)

                c_el.set("t", "s")
                c_el.remove(is_el)
                v_new = ET.SubElement(c_el, f"{{{NS_SS}}}v")
                v_new.text = str(idx)
                continue

            # ─ Case B: inlineStr etiketli ama is yok (bozuk) ─
            if t == "inlineStr" and is_el is None:
                idx = get_ss_idx("")
                c_el.set("t", "s")
                if v_el is None:
                    v_el = ET.SubElement(c_el, f"{{{NS_SS}}}v")
                v_el.text = str(idx)
                continue

            # ─ Case C: Boş hücre t="n" ─
            if t == "n" and (v_el is None or not v_el.text):
                if "t" in c_el.attrib:
                    del c_el.attrib["t"]
                if v_el is not None:
                    c_el.remove(v_el)
                continue

            # ─ Case D: Sayısal hücre t="n" (gereksiz, kaldır) ─
            if t == "n" and v_el is not None and v_el.text:
                del c_el.attrib["t"]

            # ─ Case E: Formül hücresi ─
            if f_el is not None:
                if v_el is None:
                    v_el = ET.SubElement(c_el, f"{{{NS_SS}}}v")

                if not v_el.text:
                    cached = _compute_formula_cached_value(
                        f_el.text, r, cell_vals
                    )
                    if cached is not None:
                        if isinstance(cached, str):
                            v_el.text = cached
                            c_el.set("t", "str")
                        else:
                            v_el.text = str(cached)
                            if "t" in c_el.attrib:
                                del c_el.attrib["t"]

    # ── 6. sharedStrings.xml oluştur ──
    sst_root = ET.Element("sst")
    sst_root.set("xmlns", NS_SS)
    sst_root.set("count", str(len(shared_strings)))
    sst_root.set("uniqueCount", str(len(shared_strings)))

    for s in shared_strings:
        si = ET.SubElement(sst_root, "si")
        t_el = ET.SubElement(si, "t")
        t_el.text = s
        # Boşluklu string'ler için xml:space="preserve"
        if s and (s != s.strip()):
            t_el.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")

    shared_strings_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
        + ET.tostring(sst_root, encoding="unicode", xml_declaration=False)
    )

    # ── 7. [Content_Types].xml güncelle ──
    ET.register_namespace("", NS_CT)
    ct_root = ET.fromstring(file_contents["[Content_Types].xml"])

    has_ss_override = False
    for child in ct_root:
        if child.get("PartName") == "/xl/sharedStrings.xml":
            has_ss_override = True
            break

    if not has_ss_override:
        ET.SubElement(
            ct_root,
            f"{{{NS_CT}}}Override",
            {
                "PartName": "/xl/sharedStrings.xml",
                "ContentType": "application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml",
            },
        )

    ct_xml_out = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
        + ET.tostring(ct_root, encoding="unicode", xml_declaration=False)
    )

    # ── 8. xl/_rels/workbook.xml.rels güncelle ──
    ET.register_namespace("", NS_RELS)
    rels_root = ET.fromstring(file_contents["xl/_rels/workbook.xml.rels"])

    has_ss_rel = False
    max_rid = 0
    for rel in rels_root:
        target = rel.get("Target", "")
        if "sharedStrings" in target:
            has_ss_rel = True
        rid = rel.get("Id", "")
        if rid.startswith("rId"):
            try:
                max_rid = max(max_rid, int(rid[3:]))
            except ValueError:
                pass

    if not has_ss_rel:
        ET.SubElement(
            rels_root,
            f"{{{NS_RELS}}}Relationship",
            {
                "Id": f"rId{max_rid + 1}",
                "Type": f"{NS_RELS_DOC}/sharedStrings",
                "Target": "sharedStrings.xml",
            },
        )

    rels_xml_out = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
        + ET.tostring(rels_root, encoding="unicode", xml_declaration=False)
    )

    # ── 9. Sheet XML serialize ──
    ET.register_namespace("", NS_SS)
    sheet_xml_out = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
        + ET.tostring(sheet_root, encoding="unicode", xml_declaration=False)
    )

    # ── 10. Yeni ZIP yaz ──
    tmp_path = filepath + ".tmp"
    with zipfile.ZipFile(tmp_path, "w", zipfile.ZIP_DEFLATED) as zout:
        for name in file_contents:
            if name == "xl/worksheets/sheet1.xml":
                zout.writestr(name, sheet_xml_out.encode("utf-8"))
            elif name == "[Content_Types].xml":
                zout.writestr(name, ct_xml_out.encode("utf-8"))
            elif name == "xl/_rels/workbook.xml.rels":
                zout.writestr(name, rels_xml_out.encode("utf-8"))
            else:
                zout.writestr(name, file_contents[name])
        # sharedStrings.xml ekle
        zout.writestr("xl/sharedStrings.xml", shared_strings_xml.encode("utf-8"))

    os.replace(tmp_path, filepath)
