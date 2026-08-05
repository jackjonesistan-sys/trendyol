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


class CalculatorInputTests(unittest.TestCase):
    def test_calculation_succeeds_with_required_inputs_only(self):
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


if __name__ == "__main__":
    unittest.main()
