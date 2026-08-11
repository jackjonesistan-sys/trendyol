import json
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

import openpyxl
import pandas as pd

import app


def write_workbook(path, headers, rows):
    workbook = openpyxl.Workbook()
    sheet = workbook.active
    sheet.append(headers)
    for row in rows:
        sheet.append(row)
    workbook.save(path)


class ApplyExportTests(unittest.TestCase):
    def setUp(self):
        app.app.config["TESTING"] = True
        self.client = app.app.test_client()

    def apply_from_temp(self, root, result_rows, input_files, payload):
        output_dir = root / "output"
        output_dir.mkdir()
        result_path = root / "result.xlsx"
        pd.DataFrame(result_rows).to_excel(result_path, index=False)
        manifest_path = root / "manifest.json"
        if not manifest_path.exists():
            manifest_path.write_text('{"files": {}}', encoding="utf-8")

        with (
            patch.object(app, "OUTPUT_DIR", str(output_dir)),
            patch.object(app, "F_HESAP", str(result_path)),
            patch.object(app, "INPUT_MANIFEST", str(manifest_path)),
            patch.object(app, "load_upload_set", return_value=input_files),
            patch.object(app, "fix_xlsx_for_trendyol", side_effect=lambda _path: None),
            patch(
                "fiyat_farki_analiz_script.generate_fiyat_farki_raporu",
                side_effect=lambda *_args, **_kwargs: None,
            ),
        ):
            response = self.client.post("/api/apply", json=payload)
        return response, output_dir

    def test_accounting_only_rows_are_exported_for_all_main_campaigns(self):
        cases = (
            {
                "campaign": "Avantajlı",
                "input_key": "muhasebe_avantaj",
                "headers": ["BARKOD", "YENİ TSF (FİYAT GÜNCELLE)"],
                "source_row": ["A-1", 80],
                "result_price": {"Avantajlı Ürün Fiyatı (YENİ TSF) (TL)": 75},
                "filename": "Avantajli_Urun_Etiketleri.xlsx",
                "price_header": "YENİ TSF (FİYAT GÜNCELLE)",
                "price": 75,
            },
            {
                "campaign": "Flaş",
                "input_key": "muhasebe_flas",
                "headers": ["Barkod", "Senin Belirlediğin Flaş Fiyatı"],
                "source_row": ["F-1", 85],
                "result_price": {"Flaş Ürün 24 Saat Fiyatı (TL)": 82},
                "filename": "Flas_Urunler_Genel.xlsx",
                "price_header": "24 Saat Fiyat",
                "price": 82,
            },
            {
                "campaign": "Plus",
                "input_key": "muhasebe_plus",
                "headers": ["Barkod", "Plus Fiyat Üst Limiti"],
                "source_row": ["P-1", 90],
                "result_price": {"Plus Fiyatı (TL)": 88},
                "filename": "Plus_Komisyon_Tarifeleri.xlsx",
                "price_header": "Plus Fiyat Seçimi",
                "price": 88,
            },
        )

        for case in cases:
            with self.subTest(campaign=case["campaign"]), tempfile.TemporaryDirectory() as temp_dir:
                root = Path(temp_dir)
                source = root / f'{case["input_key"]}.xlsx'
                write_workbook(source, case["headers"], [case["source_row"]])
                barcode = case["source_row"][0]
                row = {
                    "Barkod": barcode,
                    "Stok Adedi": 1,
                    "İlk Kampanya Seçimi": "Hiçbiri",
                    "İlk Ekstra Kampanya Seçimi": "Hiçbiri",
                    "Uygulanabilir Kampanyalar": case["campaign"],
                    "eligible_main_campaigns": ["Hiçbiri", case["campaign"]],
                    "eligible_extra_campaigns": ["Hiçbiri"],
                    **case["result_price"],
                }

                response, output_dir = self.apply_from_temp(
                    root,
                    [row],
                    {case["input_key"]: str(source)},
                    {
                        "target_type": case["campaign"],
                        "selections": {
                            barcode: {"main": case["campaign"], "extra": "Hiçbiri"}
                        },
                    },
                )

                self.assertEqual(response.status_code, 200, response.get_json())
                run_dir = output_dir / response.get_json()["timestamp_folder"]
                exported = openpyxl.load_workbook(run_dir / case["filename"]).active
                headers = [cell.value for cell in exported[1]]
                barcode_column = next(
                    index for index, value in enumerate(headers, 1)
                    if str(value).strip().casefold() == "barkod"
                )
                price_column = headers.index(case["price_header"]) + 1
                self.assertEqual(exported.cell(2, barcode_column).value, barcode)
                self.assertEqual(exported.cell(2, price_column).value, case["price"])

    def test_standard_template_wins_duplicates_and_accounting_only_row_is_appended(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            root = Path(temp_dir)
            standard = root / "advantage.xlsx"
            accounting = root / "accounting.xlsx"
            headers = [
                "BARKOD",
                "YENİ TSF (FİYAT GÜNCELLE)",
                "Tarife Sonuna Kadar Uygula",
                "Kaynak",
            ]
            write_workbook(standard, headers, [["DUP", 90, None, "standart"]])
            write_workbook(
                accounting,
                headers,
                [["DUP", 70, None, "muhasebe"], ["ONLY", 75, None, "muhasebe"]],
            )
            rows = [
                {
                    "Barkod": barcode,
                    "Stok Adedi": 1,
                    "Uygulanabilir Kampanyalar": "Avantajlı",
                    "eligible_main_campaigns": ["Hiçbiri", "Avantajlı"],
                    "eligible_extra_campaigns": ["Hiçbiri"],
                    "Avantajlı Ürün Fiyatı (YENİ TSF) (TL)": price,
                }
                for barcode, price in (("DUP", 81), ("ONLY", 72))
            ]

            response, output_dir = self.apply_from_temp(
                root,
                rows,
                {
                    "advantage": str(standard),
                    "muhasebe_avantaj": str(accounting),
                },
                {
                    "target_type": "Avantajlı",
                    "selections": {
                        barcode: {"main": "Avantajlı", "extra": "Hiçbiri"}
                        for barcode in ("DUP", "ONLY")
                    },
                },
            )

            self.assertEqual(response.status_code, 200, response.get_json())
            run_dir = output_dir / response.get_json()["timestamp_folder"]
            sheet = openpyxl.load_workbook(run_dir / "Avantajli_Urun_Etiketleri.xlsx").active
            exported = {
                sheet.cell(row, 1).value: sheet.cell(row, 4).value
                for row in range(2, sheet.max_row + 1)
            }
            self.assertEqual(exported, {"DUP": "standart", "ONLY": "muhasebe"})

    def test_dict_selection_preserves_extra_export_and_excludes_it_from_unapplied(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            root = Path(temp_dir)
            plus_extra = root / "plus-extra.xlsx"
            label = "Plus Ek İndirim %10"
            write_workbook(
                plus_extra,
                ["Barkod", "Maksimum Girebileceğin Fiyat", "Kampanyalı Satış Fiyatı"],
                [["E-1", 100, None]],
            )
            (root / "manifest.json").write_text(
                json.dumps({
                    "files": {},
                    "plus_extra_configs": [
                        {"path": str(plus_extra), "label": label, "rate": 10}
                    ],
                }),
                encoding="utf-8",
            )
            row = {
                "Barkod": "E-1",
                "Stok Adedi": 1,
                "İlk Kampanya Seçimi": "Hiçbiri",
                "İlk Ekstra Kampanya Seçimi": "Hiçbiri",
                "Uygulanabilir Kampanyalar": "Plus Ek İndirim",
                "eligible_main_campaigns": ["Hiçbiri"],
                "eligible_extra_campaigns": ["Hiçbiri", label],
                "Güncel Ürün Fiyatı (TL)": 100,
                "Güncel Ürün Kalan Net (TL)": 90,
                "Güncel Ürün Komisyon (%)": 10,
                "Plus Ek Fiyatı %10 (TL)": 90,
                "Plus Ek Net %10 (TL)": 81,
            }

            response, output_dir = self.apply_from_temp(
                root,
                [row],
                {},
                {
                    "target_type": "Hepsi",
                    "selections": {"E-1": {"main": "Hiçbiri", "extra": label}},
                },
            )

            self.assertEqual(response.status_code, 200, response.get_json())
            run_dir = output_dir / response.get_json()["timestamp_folder"]
            extra_export = run_dir / "Trendyol_Plus_Musterilerine_Ozel_Ek_%10_Indirim.xlsx"
            self.assertTrue(extra_export.exists())
            general = pd.read_excel(run_dir / "Kampanya_Genel_Raporu.xlsx")
            self.assertEqual(general.loc[0, "Uygulanan Kampanya Seçimi"], "Hiçbiri")
            self.assertEqual(general.loc[0, "Uygulanan Ekstra Kampanya Seçimi"], label)
            unapplied = pd.read_excel(run_dir / "Uygulanmayan_Urunler_Raporu.xlsx")
            self.assertTrue(unapplied.empty)

    def test_plus_extra_only_target_exports_evaluated_customer_price_and_config_name(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            root = Path(temp_dir)
            plus_extra = root / "plus-extra.xlsx"
            label = (
                "Plus Ek İndirim (300 TL Üzeri / 50 TL İndirim / "
                "%60 Trendyol Karşılamalı)"
            )
            write_workbook(
                plus_extra,
                ["Barkod", "Maksimum Girebileceğin Fiyat", "Kampanyalı Satış Fiyatı"],
                [["E-1", 400, None]],
            )
            (root / "manifest.json").write_text(
                json.dumps({
                    "files": {},
                    "plus_extra_configs": [{
                        "path": str(plus_extra),
                        "label": label,
                        "min_price": 300,
                        "discount_amount": 50,
                        "discount_type": "TL",
                        "trendyol_percent": 60,
                    }],
                }),
                encoding="utf-8",
            )
            row = {
                "Barkod": "E-1",
                "Stok Adedi": 1,
                "İlk Kampanya Seçimi": "Hiçbiri",
                "İlk Ekstra Kampanya Seçimi": "Hiçbiri",
                "Uygulanabilir Kampanyalar": "Plus Ek İndirim",
                "eligible_main_campaigns": ["Hiçbiri"],
                "eligible_extra_campaigns": ["Hiçbiri", label],
                "counter_evaluations": {
                    label: {
                        "price": 400,
                        "customer_price": 350,
                        "net": 330,
                        "rate": 10,
                        "seller_disc": 20,
                        "min_price": 300,
                        "disc_type": "TL",
                        "disc_val": 50,
                        "trendyol_percent": 60,
                    }
                },
            }

            response, output_dir = self.apply_from_temp(
                root,
                [row],
                {},
                {
                    "target_type": "Plus Ek İndirim",
                    "selections": {"E-1": {"main": "Hiçbiri", "extra": label}},
                },
            )

            self.assertEqual(response.status_code, 200, response.get_json())
            generated = response.get_json()["generated_files"]
            extra_name = next(Path(name).name for name in generated if "Trendyol_Plus" in name)
            self.assertIn("300_TL_Uzeri", extra_name)
            self.assertIn("50_TL_Indirim", extra_name)
            self.assertIn("%60_Trendyol_Karsilamali", extra_name)
            exported = openpyxl.load_workbook(
                output_dir / response.get_json()["timestamp_folder"] / extra_name
            ).active
            self.assertEqual(exported["C2"].value, 350)

    def test_percent_coupon_export_filename_keeps_discount_type(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            root = Path(temp_dir)
            coupon = root / "coupon.xlsx"
            label = "300 TL Üzerine %25 Kupon"
            write_workbook(
                coupon,
                ["Barkod", "Eklenecek Ürünleri Seçiniz"],
                [["C-1", None]],
            )
            (root / "manifest.json").write_text(
                json.dumps({
                    "files": {},
                    "coupon_configs": [{
                        "path": str(coupon),
                        "label": label,
                        "min_price": 300,
                        "discount_amount": 25,
                        "discount_type": "%",
                        "trendyol_percent": 60,
                    }],
                }),
                encoding="utf-8",
            )
            row = {
                "Barkod": "C-1",
                "Stok Adedi": 1,
                "Uygulanabilir Kampanyalar": label,
                "eligible_main_campaigns": ["Hiçbiri"],
                "eligible_extra_campaigns": ["Hiçbiri", label],
            }

            response, _output_dir = self.apply_from_temp(
                root,
                [row],
                {},
                {
                    "target_type": "Hepsi",
                    "selections": {"C-1": {"main": "Hiçbiri", "extra": label}},
                },
            )

            self.assertEqual(response.status_code, 200, response.get_json())
            coupon_name = next(
                Path(name).name
                for name in response.get_json()["generated_files"]
                if "Kupon" in name
            )
            self.assertIn("_%25_Kupon_", coupon_name)

    def test_selection_payload_and_number_validation_fail_closed(self):
        self.assertFalse(
            app.selection_payload_is_valid({"A1": {"main": "Bilinmeyen", "extra": "Hiçbiri"}})
        )
        self.assertFalse(
            app.selection_payload_is_valid({"A1": {"main": "Avantajlı", "unexpected": "x"}})
        )
        for value in ("nan", float("nan"), "inf", float("inf"), "-inf"):
            with self.subTest(value=value):
                self.assertIsNone(app.as_number(value))

    def test_calculate_rejects_non_finite_campaign_config_with_validation_error(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            root = Path(temp_dir)
            manifest = root / "manifest.json"
            manifest.write_text('{"files": {}}', encoding="utf-8")
            upload = root / "uploads"
            output = root / "output"
            upload.mkdir()
            output.mkdir()
            with (
                patch.object(app, "UPLOAD_DIR", str(upload)),
                patch.object(app, "OUTPUT_DIR", str(output)),
                patch.object(app, "INPUT_MANIFEST", str(manifest)),
                patch.object(app, "save_upload_set", return_value={
                    "discount": "discount.xlsx",
                    "commission": "commission.xlsx",
                    "current": "current.xlsx",
                }),
                patch(
                    "komisyon_hesaplayici.calculate_all",
                    return_value={"success": False, "message": "should not run"},
                ),
            ):
                response = self.client.post(
                    "/api/calculate",
                    data={
                        "plus_extra_configs_json": json.dumps(
                            [{"rate": float("nan")}]
                        )
                    },
                )

        self.assertEqual(response.status_code, 400)
        self.assertIn("sonlu", response.get_json()["message"])

    def test_calculate_rejects_missing_discount_even_if_loader_returns_other_base_inputs(self):
        with patch.object(
            app,
            "save_upload_set",
            return_value={"commission": "commission.xlsx", "current": "current.xlsx"},
        ):
            response = self.client.post("/api/calculate", data={})

        self.assertEqual(response.status_code, 400)
        self.assertIn("İndirim", response.get_json()["message"])


if __name__ == "__main__":
    unittest.main()
