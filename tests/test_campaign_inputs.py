import json
import tempfile
import unittest
from datetime import datetime
from pathlib import Path

import pandas as pd


class FakeUpload:
    def __init__(self, source, filename):
        self.source = Path(source)
        self.filename = filename

    def save(self, target):
        Path(target).write_bytes(self.source.read_bytes())


def write_xlsx(path, columns, row=None):
    pd.DataFrame([row or {column: None for column in columns}], columns=columns).to_excel(
        path, index=False
    )


class CampaignInputTests(unittest.TestCase):
    def test_non_finite_values_are_not_valid_numbers(self):
        from komisyon_hesaplayici import to_float

        for value in ("nan", float("nan"), "inf", float("inf"), "-inf"):
            with self.subTest(value=value):
                self.assertIsNone(to_float(value))

    def test_current_products_are_validated_by_columns_not_filename(self):
        from input_files import validate_workbook

        with tempfile.TemporaryDirectory() as temp_dir:
            path = Path(temp_dir) / "istedigim-herhangi-bir-ad.xlsx"
            write_xlsx(
                path,
                [
                    "Barkod",
                    "Komisyon Oranı",
                    "Piyasa Satış Fiyatı (KDV Dahil)",
                    "Trendyol'da Satılacak Fiyat (KDV Dahil)",
                ],
            )

            validate_workbook("current", path)

    def test_upload_set_requires_only_the_three_base_inputs(self):
        from input_files import InputValidationError, load_upload_status, save_upload_set

        with tempfile.TemporaryDirectory() as temp_dir:
            root = Path(temp_dir)
            source = root / "source"
            upload = root / "upload"
            source.mkdir()
            files = {}
            definitions = {
                "discount": ["BARKOD", "Eski Fiyat", "YENİ Fiyat", "Durum"],
                "commission": [
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
                ],
                "current": [
                    "Barkod",
                    "Komisyon Oranı",
                    "Piyasa Satış Fiyatı (KDV Dahil)",
                    "Trendyol'da Satılacak Fiyat (KDV Dahil)",
                ],
            }
            for key, columns in definitions.items():
                path = source / f"{key}.xlsx"
                write_xlsx(path, columns)
                files[key] = FakeUpload(path, "../guvenilmeyen-ad.xlsx")

            manifest_path = root / "manifest.json"
            saved = save_upload_set(files, upload, manifest_path)

            self.assertEqual(set(saved), {"discount", "commission", "current"})
            self.assertTrue(all(Path(path).resolve().parent == Path(upload).resolve() for path in saved.values()))
            self.assertFalse((root / "guvenilmeyen-ad.xlsx").exists())

            first_manifest = json.loads(manifest_path.read_text(encoding="utf-8"))
            discount_metadata = first_manifest["files"]["discount"].copy()
            datetime.fromisoformat(discount_metadata["uploaded_at"])

            replacement = source / "replacement-current.xlsx"
            write_xlsx(replacement, definitions["current"])
            saved = save_upload_set(
                {"current": FakeUpload(replacement, "guncel-yeni.xlsx")},
                upload,
                manifest_path,
            )

            self.assertEqual(set(saved), {"discount", "commission", "current"})
            second_manifest = json.loads(manifest_path.read_text(encoding="utf-8"))
            self.assertEqual(second_manifest["files"]["discount"], discount_metadata)
            self.assertEqual(
                second_manifest["files"]["current"]["original_name"],
                "guncel-yeni.xlsx",
            )
            datetime.fromisoformat(second_manifest["files"]["current"]["uploaded_at"])
            status = load_upload_status(upload, manifest_path)
            self.assertEqual(status["current"]["original_name"], "guncel-yeni.xlsx")
            self.assertRegex(status["current"]["uploaded_at_display"], r"^\d{2}\.\d{2}\.\d{4} \d{2}:\d{2}$")

            with self.assertRaises(InputValidationError):
                save_upload_set(
                    {"discount": files["discount"]},
                    root / "fresh-upload",
                    root / "fresh-manifest.json",
                )

    def test_optional_campaigns_drive_initial_selection_and_best_net_override(self):
        from komisyon_hesaplayici import choose_campaigns_smart

        candidates = [
            ("Avantajlı", 92, 100, 8),
            ("Flaş", 105, 120, 12.5),
            ("Plus", 101, 110, 8.2),
            ("Plus Ek İndirim %5", 103, 108, 4.6),
            ("Karşılamalı Kampanya", 104, 115, 9.5),
        ]
        initial, recommended, applicable = choose_campaigns_smart(100, candidates)

        self.assertEqual(initial, "Flaş")
        self.assertEqual(recommended, "Flaş Ürün")
        self.assertEqual(
            applicable,
            "Avantajlı, Flaş, Plus, Plus Ek İndirim, Karşılamalı Kampanya",
        )

        initial, recommended, applicable = choose_campaigns_smart(110, candidates)

        self.assertEqual((initial, recommended), ("Flaş", "Flaş Ürün"))
        self.assertIn("Avantajlı", applicable)

        initial, recommended, applicable = choose_campaigns_smart(
            110,
            [("Plus", 105, 100, 5, True)],
        )
        self.assertEqual((initial, recommended, applicable), ("Hiçbiri", "Hiçbiri", ""))

        initial, recommended, applicable = choose_campaigns_smart(
            110,
            [("Plus", 115, 100, 5, True)],
        )
        self.assertEqual((initial, recommended, applicable), ("Plus", "Plus Ürün", "Plus"))

        initial, recommended, applicable = choose_campaigns_smart(
            None,
            [("Plus", 115, 100, 5, True)],
        )
        self.assertEqual((initial, recommended, applicable), ("Hiçbiri", "Hiçbiri", ""))


class CalculatorInputTests(unittest.TestCase):
    def test_calculation_succeeds_with_required_inputs_and_plus_extra(self):
        from komisyon_hesaplayici import calculate_all

        with tempfile.TemporaryDirectory() as temp_dir:
            root = Path(temp_dir)
            output = root / "output"
            discount = root / "discount.xlsx"
            commission = root / "commission.xlsx"
            current = root / "current-without-name-rule.xlsx"

            pd.DataFrame(
                [{"BARKOD": "A1", "Eski Fiyat": 100, "YENİ Fiyat": 80, "Durum": "İndirim"}]
            ).to_excel(discount, index=False)
            pd.DataFrame(
                [
                    {
                        "BARKOD": "A1",
                        "1.Fiyat Alt Limit": None,
                        "2.Fiyat Üst Limiti": 89.99,
                        "2.Fiyat Alt Limit": 70,
                        "3.Fiyat Üst Limiti": 69.99,
                        "3.Fiyat Alt Limit": 50,
                        "4.Fiyat Üst Limiti": 49.99,
                        "1.KOMİSYON": 10,
                        "2.KOMİSYON": 10,
                        "3.KOMİSYON": 10,
                        "4.KOMİSYON": 10,
                        "KOMİSYONA ESAS FİYAT": 100,
                        "TARİFE GRUBU": "T",
                    }
                ]
            ).to_excel(commission, index=False)
            pd.DataFrame(
                [
                    {
                        "Barkod": "A1",
                        "Komisyon Oranı": 10,
                        "Piyasa Satış Fiyatı (KDV Dahil)": 120,
                        "Trendyol'da Satılacak Fiyat (KDV Dahil)": 100,
                    }
                ]
            ).to_excel(current, index=False)

            result = calculate_all(
                {"discount": discount, "commission": commission, "current": current},
                output_dir=output,
            )

            self.assertTrue(result["success"], result)
            report = pd.read_excel(output / "Kampanya_Hesaplama_Sonuclari.xlsx")
            self.assertEqual(report.loc[0, "İlk Kampanya Seçimi"], "Hiçbiri")
            self.assertTrue(pd.isna(report.loc[0, "Uygulanabilir Kampanyalar"]))

            plus_extra = root / "plus-extra.xlsx"
            pd.DataFrame(
                [{
                    "Barkod": "A1",
                    "Maksimum Girebileceğin Fiyat": 100,
                    "Kampanyalı Satış Fiyatı": None,
                }]
            ).to_excel(plus_extra, index=False)
            result = calculate_all(
                {
                    "discount": discount,
                    "commission": commission,
                    "current": current,
                    "plus_extra": plus_extra,
                },
                output_dir=output,
            )

            row = result["results"][0]
            self.assertIn("Plus Ek İndirim", row["Uygulanabilir Kampanyalar"])
            self.assertIn("Plus Ek İndirim %5", row["eligible_campaigns"])

            counter = root / "counter.xlsx"
            pd.DataFrame(
                [{
                    "Barkod": "A1",
                    "Maksimum Girebileceğin Fiyat": 95,
                    "Kampanyalı Satış Fiyatı": None,
                }]
            ).to_excel(counter, index=False)
            result = calculate_all(
                {"discount": discount, "commission": commission, "current": current},
                counter_files=[{
                    "path": counter,
                    "label": "Karşılamalı Test",
                    "min_price": 0,
                    "discount_amount": 10,
                    "trendyol_percent": 50,
                }],
                output_dir=output,
            )

            evaluation = result["results"][0]["counter_evaluations"]["Karşılamalı Test"]
            self.assertEqual(set(evaluation), {"price", "rate", "net", "seller_disc"})

    def test_missing_campaign_price_uses_current_price_only_when_net_improves(self):
        from komisyon_hesaplayici import calculate_all

        with tempfile.TemporaryDirectory() as temp_dir:
            root = Path(temp_dir)
            discount = root / "discount.xlsx"
            commission = root / "commission.xlsx"
            current = root / "current.xlsx"
            plus = root / "plus.xlsx"

            pd.DataFrame([
                {"BARKOD": "A1", "Eski Fiyat": None, "YENİ Fiyat": None, "Durum": "Yok"},
                {"BARKOD": "A2", "Eski Fiyat": None, "YENİ Fiyat": None, "Durum": "Yok"},
                {"BARKOD": "A3", "Eski Fiyat": None, "YENİ Fiyat": None, "Durum": "Yok"},
            ]).to_excel(discount, index=False)
            pd.DataFrame([
                {
                    "BARKOD": barcode,
                    "1.Fiyat Alt Limit": 0,
                    "2.Fiyat Üst Limiti": None,
                    "2.Fiyat Alt Limit": None,
                    "3.Fiyat Üst Limiti": None,
                    "3.Fiyat Alt Limit": None,
                    "4.Fiyat Üst Limiti": None,
                    "1.KOMİSYON": 20,
                    "2.KOMİSYON": 20,
                    "3.KOMİSYON": 20,
                    "4.KOMİSYON": 20,
                }
                for barcode in ("A1", "A2", "A3")
            ]).to_excel(commission, index=False)
            pd.DataFrame([
                {
                    "Barkod": barcode,
                    "Komisyon Oranı": 20,
                    "Piyasa Satış Fiyatı (KDV Dahil)": 100,
                    "Trendyol'da Satılacak Fiyat (KDV Dahil)": 100,
                }
                for barcode in ("A1", "A2", "A3")
            ]).to_excel(current, index=False)
            pd.DataFrame([
                {
                    "Barkod": "A1",
                    "Plus Fiyat Üst Limiti": None,
                    "Plus Komisyon Teklifi": 10,
                    "Plus Fiyat Seçimi": None,
                    "Tarife Seçimi": None,
                },
                {
                    "Barkod": "A2",
                    "Plus Fiyat Üst Limiti": None,
                    "Plus Komisyon Teklifi": 20,
                    "Plus Fiyat Seçimi": None,
                    "Tarife Seçimi": None,
                },
                {
                    "Barkod": "A3",
                    "Plus Fiyat Üst Limiti": None,
                    "Plus Komisyon Teklifi": "-inf",
                    "Plus Fiyat Seçimi": None,
                    "Tarife Seçimi": None,
                },
            ]).to_excel(plus, index=False)

            result = calculate_all(
                {
                    "discount": discount,
                    "commission": commission,
                    "current": current,
                    "plus": plus,
                },
                output_dir=root / "output",
            )
            rows = {row["Barkod"]: row for row in result["results"]}

            self.assertEqual(rows["A1"]["Plus Fiyatı (TL)"], 100)
            self.assertEqual(rows["A1"]["İlk Kampanya Seçimi"], "Hiçbiri")
            self.assertEqual(rows["A1"]["Önerilen Kampanya"], "Plus")
            self.assertIn("Plus", rows["A1"]["eligible_campaigns"])
            self.assertEqual(rows["A2"]["İlk Kampanya Seçimi"], "Hiçbiri")
            self.assertEqual(rows["A2"]["eligible_campaigns"], ["Hiçbiri"])
            self.assertEqual(rows["A3"]["İlk Kampanya Seçimi"], "Hiçbiri")
            self.assertEqual(rows["A3"]["eligible_campaigns"], ["Hiçbiri"])

            plus_extra = root / "plus-extra.xlsx"
            pd.DataFrame([
                {"BARKOD": "A1", "Eski Fiyat": 100, "YENİ Fiyat": 90, "Durum": "İndirim"},
            ]).to_excel(discount, index=False)
            pd.DataFrame([
                {
                    "Barkod": "A1",
                    "Maksimum Girebileceğin Fiyat": 90,
                    "Kampanyalı Satış Fiyatı": None,
                },
            ]).to_excel(plus_extra, index=False)

            result = calculate_all(
                {
                    "discount": discount,
                    "commission": commission,
                    "current": current,
                    "plus_extra": plus_extra,
                },
                output_dir=root / "output-plus-extra",
            )
            row = result["results"][0]
            self.assertEqual(row["İlk Kampanya Seçimi"], "Hiçbiri")
            self.assertEqual(row["eligible_campaigns"], ["Hiçbiri"])


if __name__ == "__main__":
    unittest.main()
