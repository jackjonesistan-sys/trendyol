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


def write_calculation_base(root, barcodes, current_price=200, current_rate=0):
    paths = {name: root / f"{name}.xlsx" for name in ("discount", "commission", "current")}
    pd.DataFrame([
        {"BARKOD": barcode, "Eski Fiyat": current_price, "YENİ Fiyat": 80, "Durum": "İndirim"}
        for barcode in barcodes
    ]).to_excel(paths["discount"], index=False)
    pd.DataFrame([
        {
            "BARKOD": barcode,
            "1.Fiyat Alt Limit": 0,
            "2.Fiyat Üst Limiti": None,
            "2.Fiyat Alt Limit": None,
            "3.Fiyat Üst Limiti": None,
            "3.Fiyat Alt Limit": None,
            "4.Fiyat Üst Limiti": None,
            "1.KOMİSYON": current_rate,
            "2.KOMİSYON": current_rate,
            "3.KOMİSYON": current_rate,
            "4.KOMİSYON": current_rate,
        }
        for barcode in barcodes
    ]).to_excel(paths["commission"], index=False)
    pd.DataFrame([
        {
            "Barkod": barcode,
            "Komisyon Oranı": current_rate,
            "Piyasa Satış Fiyatı (KDV Dahil)": current_price,
            "Trendyol'da Satılacak Fiyat (KDV Dahil)": current_price,
        }
        for barcode in barcodes
    ]).to_excel(paths["current"], index=False)
    return paths


class CampaignInputTests(unittest.TestCase):
    def test_non_finite_values_are_not_valid_numbers(self):
        from komisyon_hesaplayici import to_float

        for value in ("nan", float("nan"), "inf", float("inf"), "-inf"):
            with self.subTest(value=value):
                self.assertIsNone(to_float(value))

    def test_campaign_configs_share_validation_and_legacy_plus_defaults(self):
        from input_files import InputValidationError, normalize_campaign_config

        legacy = normalize_campaign_config({"rate": 10}, "plus_extra")
        self.assertEqual(
            {key: legacy[key] for key in (
                "min_price",
                "discount_amount",
                "discount_type",
                "trendyol_percent",
            )},
            {
                "min_price": 0.0,
                "discount_amount": 10.0,
                "discount_type": "%",
                "trendyol_percent": 0.0,
            },
        )

        invalid_configs = (
            {"min_price": -1},
            {"discount_amount": -1},
            {"discount_type": "EUR"},
            {"trendyol_percent": -1},
            {"trendyol_percent": 101},
            {"discount_type": "%", "discount_amount": 101},
            {"min_price": float("nan")},
            {"discount_amount": float("inf")},
            {"trendyol_percent": float("-inf")},
        )
        for campaign_type in ("counter", "plus_extra", "coupon"):
            for config in invalid_configs:
                with self.subTest(campaign_type=campaign_type, config=config):
                    with self.assertRaises(InputValidationError):
                        normalize_campaign_config(config, campaign_type)

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
                "discount": ["BARKOD", "Eski Fiyat", "YENİ Fiyat"],
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

        highest_net_rule = {
            "enabled": False,
            "priority": ["Avantajlı", "Flaş", "Plus"],
        }
        candidates = [
            ("Avantajlı", 92, 100, 8),
            ("Flaş", 105, 120, 12.5),
            ("Plus", 101, 110, 8.2),
            ("Plus Ek İndirim %5", 103, 108, 4.6),
            ("Karşılamalı Kampanya", 104, 115, 9.5),
        ]
        res = choose_campaigns_smart(100, candidates, highest_net_rule)
        initial, recommended, applicable = res[0], res[1], res[2]

        self.assertEqual(initial, "Flaş")
        self.assertEqual(recommended, "Flaş Ürün")
        self.assertEqual(res[3], "Karşılamalı Kampanya")
        self.assertEqual(
            applicable,
            "Avantajlı, Flaş, Plus, Plus Ek İndirim, Karşılamalı Kampanya",
        )

        res = choose_campaigns_smart(110, candidates, highest_net_rule)
        initial, recommended, applicable = res[0], res[1], res[2]

        self.assertEqual((initial, recommended), ("Flaş", "Flaş Ürün"))
        self.assertIn("Avantajlı", applicable)

        tied_candidates = [
            ("Plus", 105, 110, 5),
            ("Flaş", 105, 120, 10),
            ("Avantajlı", 105, 120, 10),
            ("Ekstra A", 103, 105, 2),
            ("Ekstra B", 103, 110, 2),
        ]
        first = choose_campaigns_smart(100, tied_candidates, highest_net_rule)
        reversed_order = choose_campaigns_smart(
            100, list(reversed(tied_candidates)), highest_net_rule
        )
        self.assertEqual(first[0], "Avantajlı")
        self.assertEqual(first[3], "Ekstra B")
        self.assertEqual((first[0], first[3]), (reversed_order[0], reversed_order[3]))

        res = choose_campaigns_smart(
            110,
            [("Plus", 105, 100, 5, True)],
            highest_net_rule,
        )
        initial, recommended, applicable = res[0], res[1], res[2]
        self.assertEqual((initial, recommended, applicable), ("Hiçbiri", "Hiçbiri", ""))

        res = choose_campaigns_smart(
            110,
            [("Plus", 115, 100, 5, True)],
            highest_net_rule,
        )
        initial, recommended, applicable = res[0], res[1], res[2]
        self.assertEqual((initial, recommended, applicable), ("Plus", "Plus Ürün", "Plus"))

        res = choose_campaigns_smart(
            None,
            [("Plus", 115, 100, 5, True)],
            highest_net_rule,
        )
        initial, recommended, applicable = res[0], res[1], res[2]
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
            self.assertEqual(report.loc[0, "Uygulanabilir Kampanyalar"], "Komisyon Tarifesi")

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
            self.assertTrue({"price", "rate", "net", "seller_disc"}.issubset(set(evaluation)))

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

    def test_template_prices_without_real_floor_are_not_recommended(self):
        from komisyon_hesaplayici import calculate_all

        with tempfile.TemporaryDirectory() as temp_dir:
            root = Path(temp_dir)
            discount = root / "discount.xlsx"
            commission = root / "commission.xlsx"
            current = root / "current.xlsx"
            advantage = root / "advantage.xlsx"
            flash = root / "flash.xlsx"
            plus = root / "plus.xlsx"

            pd.DataFrame([{"BARKOD": "A1", "Eski Fiyat": None, "YENİ Fiyat": None}]).to_excel(discount, index=False)
            pd.DataFrame([{
                "BARKOD": "A1",
                "1.Fiyat Alt Limit": 0,
                "1.KOMİSYON": 10,
                "2.KOMİSYON": 10,
                "3.KOMİSYON": 10,
                "4.KOMİSYON": 10,
            }]).to_excel(commission, index=False)
            pd.DataFrame([{
                "Barkod": "A1",
                "Komisyon Oranı": 10,
                "Piyasa Satış Fiyatı (KDV Dahil)": 120,
                "Trendyol'da Satılacak Fiyat (KDV Dahil)": 100,
            }]).to_excel(current, index=False)
            pd.DataFrame([{
                "BARKOD": "A1",
                "YENİ TSF (FİYAT GÜNCELLE)": 90,
                "1 YILDIZ ÜST FİYAT": 90,
            }]).to_excel(advantage, index=False)
            pd.DataFrame([{
                "Barkod": "A1",
                "24 Saat Fiyat": 89,
                "3 Saat Fiyat": 89,
            }]).to_excel(flash, index=False)
            pd.DataFrame([{
                "Barkod": "A1",
                "Plus Fiyat Üst Limiti": 88,
                "Plus Komisyon Teklifi": 5,
            }]).to_excel(plus, index=False)

            result = calculate_all(
                {
                    "discount": discount,
                    "commission": commission,
                    "current": current,
                    "advantage": advantage,
                    "flash": flash,
                    "plus": plus,
                },
                output_dir=root / "output",
            )

            row = result["results"][0]
            self.assertEqual(row["Düşülebilecek Dip Fiyat (TL)"], None)
            self.assertEqual(row["Önerilen Kampanya"], "Hiçbiri")
            self.assertEqual(row["eligible_main_campaigns"], ["Hiçbiri"])
            self.assertEqual(row["all_matching_main_campaigns"], ["Hiçbiri"])

            pd.DataFrame([{
                "BARKOD": "A1", "Eski Fiyat": 100, "YENİ Fiyat": 80, "Durum": "İndirim"
            }]).to_excel(discount, index=False)
            with_common_floor = calculate_all(
                {
                    "discount": discount,
                    "commission": commission,
                    "current": current,
                    "advantage": advantage,
                    "flash": flash,
                    "plus": plus,
                },
                output_dir=root / "output-common-floor",
            )["results"][0]
            self.assertEqual(
                set(with_common_floor["eligible_main_campaigns"]),
                {"Hiçbiri", "Avantajlı", "Flaş", "Plus"},
            )

            pd.DataFrame([{"BARKOD": "A1", "Eski Fiyat": None, "YENİ Fiyat": None}]).to_excel(discount, index=False)
            accounting_advantage = root / "accounting-advantage.xlsx"
            pd.DataFrame([{
                "BARKOD": "A1",
                "TRENDYOL SATIŞ FİYATI": 90,
            }]).to_excel(accounting_advantage, index=False)
            with_accounting_floor = calculate_all(
                {
                    "discount": discount,
                    "commission": commission,
                    "current": current,
                    "advantage": advantage,
                    "flash": flash,
                    "plus": plus,
                    "muhasebe_avantaj": accounting_advantage,
                },
                output_dir=root / "output-accounting-floor",
            )["results"][0]
            self.assertIn("Avantajlı", with_accounting_floor["eligible_main_campaigns"])
            self.assertEqual(with_accounting_floor["Düşülebilecek Dip Fiyat (TL)"], 90)

            plus_extra = root / "plus-extra-below-accounting-floor.xlsx"
            pd.DataFrame([{
                "Barkod": "A1",
                "Maksimum Girebileceğin Fiyat": 85,
            }]).to_excel(plus_extra, index=False)
            with_unsafe_extra = calculate_all(
                {
                    "discount": discount,
                    "commission": commission,
                    "current": current,
                    "muhasebe_avantaj": accounting_advantage,
                },
                plus_extra_files=[{
                    "path": plus_extra,
                    "label": "Plus Ek İndirim %10",
                    "rate": 10,
                }],
                output_dir=root / "output-extra-floor",
            )["results"][0]
            self.assertEqual(with_unsafe_extra["eligible_extra_campaigns"], ["Hiçbiri"])

    def test_counter_and_coupon_minimum_prices_are_enforced_at_boundary(self):
        from komisyon_hesaplayici import calculate_all

        with tempfile.TemporaryDirectory() as temp_dir:
            root = Path(temp_dir)
            discount = root / "discount.xlsx"
            commission = root / "commission.xlsx"
            current = root / "current.xlsx"
            counter = root / "counter.xlsx"
            coupon = root / "coupon.xlsx"

            pd.DataFrame([
                {"BARKOD": "LOW", "Eski Fiyat": 99, "YENİ Fiyat": 50, "Durum": "İndirim"},
                {"BARKOD": "EDGE", "Eski Fiyat": 100, "YENİ Fiyat": 50, "Durum": "İndirim"},
            ]).to_excel(discount, index=False)
            pd.DataFrame([
                {
                    "BARKOD": barcode,
                    "1.Fiyat Alt Limit": 0,
                    "1.KOMİSYON": 10,
                    "2.KOMİSYON": 10,
                    "3.KOMİSYON": 10,
                    "4.KOMİSYON": 10,
                }
                for barcode in ("LOW", "EDGE")
            ]).to_excel(commission, index=False)
            pd.DataFrame([
                {
                    "Barkod": barcode,
                    "Komisyon Oranı": 10,
                    "Piyasa Satış Fiyatı (KDV Dahil)": price,
                    "Trendyol'da Satılacak Fiyat (KDV Dahil)": price,
                }
                for barcode, price in (("LOW", 99), ("EDGE", 100))
            ]).to_excel(current, index=False)
            campaign_rows = [
                {"Barkod": "LOW", "Maksimum Girebileceğin Fiyat": 99},
                {"Barkod": "EDGE", "Maksimum Girebileceğin Fiyat": 100},
            ]
            pd.DataFrame(campaign_rows).to_excel(counter, index=False)
            pd.DataFrame(campaign_rows).to_excel(coupon, index=False)

            result = calculate_all(
                {"discount": discount, "commission": commission, "current": current},
                counter_files=[{
                    "path": counter,
                    "label": "Karşılamalı Minimum",
                    "min_price": 100,
                    "discount_amount": 10,
                    "trendyol_percent": 100,
                }],
                coupon_files=[{
                    "path": coupon,
                    "label": "Kupon Minimum",
                    "min_price": 100,
                    "discount_amount": 10,
                    "trendyol_percent": 100,
                }],
                output_dir=root / "output",
            )
            rows = {row["Barkod"]: row for row in result["results"]}

            self.assertEqual(rows["LOW"]["eligible_extra_campaigns"], ["Hiçbiri"])
            self.assertEqual(rows["LOW"]["counter_evaluations"], {})
            self.assertEqual(rows["LOW"]["Önerilen Ekstra Kampanya"], "Hiçbiri")
            self.assertIn("Karşılamalı Minimum", rows["EDGE"]["eligible_extra_campaigns"])
            self.assertIn("Kupon Minimum", rows["EDGE"]["eligible_extra_campaigns"])

    def test_missing_current_net_does_not_compare_fallback_net_to_none(self):
        from komisyon_hesaplayici import calculate_all

        with tempfile.TemporaryDirectory() as temp_dir:
            root = Path(temp_dir)
            discount = root / "discount.xlsx"
            commission = root / "commission.xlsx"
            current = root / "current.xlsx"
            plus = root / "plus.xlsx"

            pd.DataFrame([{"BARKOD": "A1", "Eski Fiyat": None, "YENİ Fiyat": None}]).to_excel(discount, index=False)
            pd.DataFrame([{
                "BARKOD": "A1",
                "1.Fiyat Alt Limit": 0,
                "1.KOMİSYON": None,
                "2.KOMİSYON": None,
                "3.KOMİSYON": None,
                "4.KOMİSYON": None,
            }]).to_excel(commission, index=False)
            pd.DataFrame([{
                "Barkod": "A1",
                "Komisyon Oranı": None,
                "Piyasa Satış Fiyatı (KDV Dahil)": 100,
                "Trendyol'da Satılacak Fiyat (KDV Dahil)": 100,
            }]).to_excel(current, index=False)
            pd.DataFrame([{
                "Barkod": "A1",
                "Plus Fiyat Üst Limiti": None,
                "Plus Komisyon Teklifi": 10,
            }]).to_excel(plus, index=False)

            result = calculate_all(
                {
                    "discount": discount,
                    "commission": commission,
                    "current": current,
                    "plus": plus,
                },
                output_dir=root / "output",
            )

            self.assertTrue(result["success"])
            self.assertEqual(result["results"][0]["eligible_main_campaigns"], ["Hiçbiri"])

    def test_parse_counter_filename_percentage_mode(self):
        from input_files import parse_counter_filename

        min_p, disc, tr_p, disc_type = parse_counter_filename("2000_TL_Uzeri_%10_Indirim_%30_Trendyol_Karsilamali.xlsx")
        self.assertEqual(min_p, 2000.0)
        self.assertEqual(disc, 10.0)
        self.assertEqual(tr_p, 30.0)
        self.assertEqual(disc_type, "%")

        min_p2, disc2, tr_p2, disc_type2 = parse_counter_filename("500_TL_Uzeri_40_TL_Indirim_%30_Trendyol_Karsilamali.xlsx")
        self.assertEqual(min_p2, 500.0)
        self.assertEqual(disc2, 40.0)
        self.assertEqual(tr_p2, 30.0)
        self.assertEqual(disc_type2, "TL")

    def test_parse_coupon_filename(self):
        from input_files import parse_coupon_filename

        min_p, disc, tr_p, disc_type = parse_coupon_filename("750-tl-uzerine-100-tl-kupon-trendyol-plus-musterilerine-ozel-_2026-08-10_17-12_tr-TR_part_1.xlsx")
        self.assertEqual(min_p, 750.0)
        self.assertEqual(disc, 100.0)
        self.assertEqual(tr_p, 0.0)
        self.assertEqual(disc_type, "TL")

        percent = parse_coupon_filename(
            "300_TL_Uzerine_%10_Kupon_%60_Trendyol_Karsilamali.xlsx"
        )
        self.assertEqual(percent, (300.0, 10.0, 60.0, "%"))

    def test_parse_plus_extra_filename_uses_named_fields_not_first_number(self):
        from input_files import parse_plus_extra_filename

        parsed = parse_plus_extra_filename(
            "Trendyol Plus Müşterilerine Özel 300 TL Üzerine Ek %10 İndirim "
            "- %60 Trendyol Karşılamalı.xlsx"
        )

        self.assertEqual(parsed, (300.0, 10.0, 60.0, "%"))
        self.assertEqual(
            parse_plus_extra_filename(
                "trendyol-plus-musterilerine-ozel-ek-5-indirim_2026.xlsx"
            ),
            (0.0, 5.0, 0.0, "%"),
        )
        self.assertEqual(
            parse_plus_extra_filename(
                "Trendyol Plus Müşterilerine Özel Ek yüzde10 İndirim.xlsx"
            ),
            (0.0, 10.0, 0.0, "%"),
        )

    def test_extra_campaigns_share_customer_price_and_seller_net_contract(self):
        from komisyon_hesaplayici import calculate_all

        with tempfile.TemporaryDirectory() as temp_dir:
            root = Path(temp_dir)
            discount = root / "discount.xlsx"
            commission = root / "commission.xlsx"
            current = root / "current.xlsx"
            campaign = root / "campaign.xlsx"

            pd.DataFrame([{
                "BARKOD": "A1",
                "Eski Fiyat": 400,
                "YENİ Fiyat": 300,
                "Durum": "İndirim",
            }]).to_excel(discount, index=False)
            pd.DataFrame([{
                "BARKOD": "A1",
                "1.Fiyat Alt Limit": 0,
                "2.Fiyat Üst Limiti": None,
                "2.Fiyat Alt Limit": None,
                "3.Fiyat Üst Limiti": None,
                "3.Fiyat Alt Limit": None,
                "4.Fiyat Üst Limiti": None,
                "1.KOMİSYON": 10,
                "2.KOMİSYON": 10,
                "3.KOMİSYON": 10,
                "4.KOMİSYON": 10,
                "KOMİSYONA ESAS FİYAT": 400,
                "TARİFE GRUBU": "T",
            }]).to_excel(commission, index=False)
            pd.DataFrame([{
                "Barkod": "A1",
                "Komisyon Oranı": 10,
                "Piyasa Satış Fiyatı (KDV Dahil)": 400,
                "Trendyol'da Satılacak Fiyat (KDV Dahil)": 400,
            }]).to_excel(current, index=False)
            pd.DataFrame([{
                "Barkod": "A1",
                "Maksimum Girebileceğin Fiyat": 400,
                "Kampanyalı Satış Fiyatı": None,
            }]).to_excel(campaign, index=False)

            counter_label = "Karşılamalı Yüzde"
            coupon_label = "Yüzde Kupon"
            advanced_plus_label = (
                "Plus Ek İndirim (300 TL Üzeri / %25 İndirim / "
                "%60 Trendyol Karşılamalı)"
            )
            legacy_plus_label = "Plus Ek İndirim %10"
            below_dip_label = "Plus Ek İndirim Dip Altı"
            result = calculate_all(
                {"discount": discount, "commission": commission, "current": current},
                counter_files=[{
                    "path": campaign,
                    "label": counter_label,
                    "min_price": 300,
                    "discount_type": "%",
                    "discount_amount": 10,
                    "trendyol_percent": 60,
                }],
                coupon_files=[{
                    "path": campaign,
                    "label": coupon_label,
                    "min_price": 300,
                    "discount_type": "%",
                    "discount_amount": 25,
                    "trendyol_percent": 50,
                }],
                plus_extra_files=[
                    {
                        "path": campaign,
                        "min_price": 300,
                        "discount_type": "%",
                        "discount_amount": 25,
                        "trendyol_percent": 60,
                    },
                    {"path": campaign, "rate": 10},
                    {
                        "path": campaign,
                        "label": below_dip_label,
                        "discount_type": "%",
                        "discount_amount": 30,
                    },
                ],
                output_dir=root / "output",
            )

            row = result["results"][0]
            evaluations = row["counter_evaluations"]
            expected = {
                counter_label: (360, 344, 16, "%", 10, 60),
                coupon_label: (300, 310, 50, "%", 25, 50),
                advanced_plus_label: (300, 320, 40, "%", 25, 60),
                legacy_plus_label: (360, 320, 40, "%", 10, 0),
            }
            for label, values in expected.items():
                with self.subTest(label=label):
                    evaluation = evaluations[label]
                    self.assertEqual(evaluation["price"], 400)
                    self.assertEqual(evaluation["customer_price"], values[0])
                    self.assertEqual(evaluation["net"], values[1])
                    self.assertEqual(evaluation["seller_disc"], values[2])
                    self.assertEqual(evaluation["min_price"], 300 if label not in (legacy_plus_label,) else 0)
                    self.assertEqual(evaluation["disc_type"], values[3])
                    self.assertEqual(evaluation["disc_val"], values[4])
                    self.assertEqual(evaluation["trendyol_percent"], values[5])

            self.assertNotIn(below_dip_label, evaluations)
            self.assertNotIn(below_dip_label, row["eligible_extra_campaigns"])

    def test_disabled_campaign_files_are_ignored_in_calculation(self):
        from komisyon_hesaplayici import calculate_all

        with tempfile.TemporaryDirectory() as temp_dir:
            root = Path(temp_dir)
            commission = root / "commission.xlsx"
            current = root / "current.xlsx"
            counter = root / "counter.xlsx"
            discount = root / "discount.xlsx"

            pd.DataFrame([
                {"BARKOD": "A1", "Eski Fiyat": 100, "YENİ Fiyat": 90, "Durum": "İndirim"},
            ]).to_excel(discount, index=False)

            pd.DataFrame([
                {
                    "BARKOD": "A1",
                    "1.Fiyat Alt Limit": 0,
                    "2.Fiyat Üst Limiti": 100,
                    "2.Fiyat Alt Limit": 100,
                    "3.Fiyat Üst Limiti": 200,
                    "3.Fiyat Alt Limit": 200,
                    "4.Fiyat Üst Limiti": 300,
                    "1.KOMİSYON": 15,
                    "2.KOMİSYON": 15,
                    "3.KOMİSYON": 15,
                    "4.KOMİSYON": 15,
                    "KOMİSYONA ESAS FİYAT": 100,
                    "TARİFE GRUBU": "Grup 1",
                }
            ]).to_excel(commission, index=False)

            pd.DataFrame([
                {
                    "Barkod": "A1",
                    "Komisyon Oranı": 15,
                    "Piyasa Satış Fiyatı (KDV Dahil)": 150,
                    "Trendyol'da Satılacak Fiyat (KDV Dahil)": 100,
                }
            ]).to_excel(current, index=False)

            pd.DataFrame([
                {"Barkod": "A1", "Maksimum Girebileceğin Fiyat": 90}
            ]).to_excel(counter, index=False)

            res_disabled = calculate_all(
                {"discount": discount, "commission": commission, "current": current},
                counter_files=[{"path": counter, "min_price": 50, "discount_amount": 10, "enabled": False}],
                output_dir=root / "output-disabled",
            )
            self.assertEqual(res_disabled["results"][0]["all_matching_extra_campaigns"], ["Hiçbiri"])


class PlusPeriodCalculationTests(unittest.TestCase):
    def test_flash_merge_follows_export_template_interval_identity(self):
        from komisyon_hesaplayici import merge_flash_intervals

        def interval(start, price, origin="kampanya"):
            return {
                "period": "24 Saat",
                "start": start,
                "end": start.replace("00:00:00", "23:59:00") if start else None,
                "price": price,
                "source": "24 Saat Fiyat",
                "origin": origin,
                "used_current_price": False,
            }

        standard_17 = interval("2026-08-17 00:00:00", 980)
        standard_18 = interval("2026-08-18 00:00:00", 970)
        accounting_18 = interval("2026-08-18 00:00:00", 955, "muhasebe")

        self.assertEqual(
            merge_flash_intervals([standard_17], [accounting_18]),
            [standard_17],
        )

        exact_accounting = interval("2026-08-17 00:00:00", 950, "muhasebe")
        self.assertEqual(
            merge_flash_intervals([standard_17], [exact_accounting]),
            [exact_accounting],
        )
        second_exact = {**exact_accounting, "price": 940}
        self.assertEqual(
            merge_flash_intervals([standard_17], [exact_accounting, second_exact]),
            [standard_17],
        )

        generic_accounting = interval(None, 945, "muhasebe")
        generic_overlay = merge_flash_intervals([standard_17], [generic_accounting])
        self.assertEqual(
            (generic_overlay[0]["start"], generic_overlay[0]["end"], generic_overlay[0]["price"]),
            (standard_17["start"], standard_17["end"], 945),
        )
        self.assertEqual(
            merge_flash_intervals([standard_17, standard_18], [generic_accounting]),
            [standard_17, standard_18],
        )
        second_generic = {**generic_accounting, "price": 940}
        self.assertEqual(
            merge_flash_intervals([standard_17], [generic_accounting, second_generic]),
            [standard_17],
        )

        accounting_only = [exact_accounting, accounting_18]
        self.assertEqual(merge_flash_intervals([], accounting_only), accounting_only)

    def test_period_helpers_preserve_duplicate_offer_positions_and_workbook_labels(self):
        from input_files import choose_plus_tariff_label, find_plus_period_columns

        with tempfile.TemporaryDirectory() as temp_dir:
            workbook = Path(temp_dir) / "plus.xlsx"
            columns = [
                "Barkod",
                "Tarih Aralığı (2 Gün)",
                "Plus Komisyon Teklifi",
                "Tarih Aralığı (5 Gün)",
                "Plus Komisyon Teklifi",
                "2 Gün Tarih Aralığı",
                "5 Gün Tarih Aralığı",
                "7 Gün Tarih Aralığı",
            ]
            pd.DataFrame([[
                "A1", "1-3 Eylül", 10, "3-8 Eylül", 20,
                "2 Günlük Fiyat (1-3 Eylül)",
                "5 Günlük Fiyat (3-8 Eylül)",
                "Birleşik 7 Günlük Fiyat",
            ]], columns=columns).to_excel(workbook, index=False)
            row = pd.read_excel(workbook).iloc[0]

        periods = find_plus_period_columns(row.index)
        self.assertEqual(
            [
                (
                    period["days"],
                    period["date_position"],
                    period["offer_position"],
                    period["date_column"],
                    period["offer_column"],
                )
                for period in periods
            ],
            [
                (2, 1, 2, "Tarih Aralığı (2 Gün)", "Plus Komisyon Teklifi"),
                (5, 3, 4, "Tarih Aralığı (5 Gün)", "Plus Komisyon Teklifi.1"),
            ],
        )
        self.assertEqual(
            choose_plus_tariff_label(row, periods, [2, 5]),
            "Birleşik 7 Günlük Fiyat",
        )
        self.assertEqual(
            choose_plus_tariff_label(row, periods, [5]),
            "5 Günlük Fiyat (3-8 Eylül)",
        )
        fallback_row = row.drop(labels=["5 Gün Tarih Aralığı", "7 Gün Tarih Aralığı"])
        self.assertEqual(choose_plus_tariff_label(fallback_row, periods, [2, 5]), "7 Günlük Fiyat")
        self.assertEqual(
            choose_plus_tariff_label(fallback_row, periods, [5]),
            "5 Günlük Fiyat (3-8 Eylül)",
        )
        self.assertIsNone(choose_plus_tariff_label(row, periods, []))

    def test_two_period_plus_selects_both_first_second_or_none(self):
        from komisyon_hesaplayici import calculate_all

        with tempfile.TemporaryDirectory() as temp_dir:
            root = Path(temp_dir)
            barcodes = ("BOTH", "FIRST", "SECOND", "NONE")
            inputs = write_calculation_base(root, barcodes)
            plus = root / "plus.xlsx"
            columns = [
                "Barkod",
                "Plus Fiyat Üst Limiti",
                "Tarih Aralığı (3 Gün)",
                "Plus Komisyon Teklifi",
                "Tarih Aralığı (4 Gün)",
                "Plus Komisyon Teklifi",
                "Plus Fiyat Seçimi",
                "Tarife Seçimi",
                "3 Gün Tarih Aralığı",
                "4 Gün Tarih Aralığı",
                "7 Gün Tarih Aralığı",
            ]
            rows = [
                ["BOTH", 100, "11-14 Ağustos", 10, "14-18 Ağustos", 20, None, None,
                 "3 Günlük Fiyat (11-14 Ağustos)", "4 Günlük Fiyat (14-18 Ağustos)", "7 Günlük Fiyat"],
                ["FIRST", 100, "11-14 Ağustos", 11, None, 19, None, None,
                 "3 Günlük Fiyat (11-14 Ağustos)", "4 Günlük Fiyat (14-18 Ağustos)", "7 Günlük Fiyat"],
                ["SECOND", 100, None, 18, "14-18 Ağustos", 12, None, None,
                 "3 Günlük Fiyat (11-14 Ağustos)", "4 Günlük Fiyat (14-18 Ağustos)", "7 Günlük Fiyat"],
                ["NONE", 100, "11-14 Ağustos", -1, "14-18 Ağustos", 101, None, None,
                 "3 Günlük Fiyat (11-14 Ağustos)", "4 Günlük Fiyat (14-18 Ağustos)", "7 Günlük Fiyat"],
            ]
            pd.DataFrame(rows, columns=columns).to_excel(plus, index=False)
            inputs["plus"] = plus

            result = calculate_all(inputs, output_dir=root / "output")
            by_barcode = {row["Barkod"]: row for row in result["results"]}

        both = by_barcode["BOTH"]
        self.assertEqual(both["Plus Tarife Seçimi"], "7 Günlük Fiyat")
        self.assertEqual((both["Plus Komisyon (%)"], both["Plus Net (TL)"]), (20, 80))
        self.assertEqual((both["Plus Komisyon (3 Gün) (%)"], both["Plus Net (3 Gün) (TL)"]), (10, 90))
        self.assertEqual((both["Plus Komisyon (4 Gün) (%)"], both["Plus Net (4 Gün) (TL)"]), (20, 80))
        self.assertEqual(both["eligible_main_campaigns"].count("Plus"), 1)
        self.assertEqual(both["Önerilen Kampanya"], "Plus")

        first = by_barcode["FIRST"]
        self.assertEqual(first["Plus Tarife Seçimi"], "3 Günlük Fiyat (11-14 Ağustos)")
        self.assertEqual((first["Plus Komisyon (%)"], first["Plus Net (TL)"]), (11, 89))
        self.assertEqual((first["Plus Komisyon (4 Gün) (%)"], first["Plus Net (4 Gün) (TL)"]), (19, 81))

        second = by_barcode["SECOND"]
        self.assertEqual(second["Plus Tarife Seçimi"], "4 Günlük Fiyat (14-18 Ağustos)")
        self.assertEqual((second["Plus Komisyon (%)"], second["Plus Net (TL)"]), (12, 88))
        self.assertEqual((second["Plus Komisyon (3 Gün) (%)"], second["Plus Net (3 Gün) (TL)"]), (18, 82))

        none = by_barcode["NONE"]
        self.assertIsNone(none["Plus Tarife Seçimi"])
        self.assertNotIn("Plus", none["eligible_main_campaigns"])
        self.assertEqual((none["Plus Komisyon (%)"], none["Plus Net (TL)"]), (None, None))
        for days in (3, 4):
            self.assertIn(f"Plus Komisyon ({days} Gün) (%)", none)
            self.assertIn(f"Plus Net ({days} Gün) (TL)", none)

    def test_variable_periods_and_price_above_upper_limit(self):
        from komisyon_hesaplayici import calculate_all

        with tempfile.TemporaryDirectory() as temp_dir:
            root = Path(temp_dir)
            inputs = write_calculation_base(root, ("VARIABLE", "TOO_HIGH"))
            plus = root / "plus.xlsx"
            columns = [
                "Barkod", "Plus Fiyat Üst Limiti",
                "Tarih Aralığı (2 Gün)", "Plus Komisyon Teklifi",
                "Tarih Aralığı (5 Gün)", "Plus Komisyon Teklifi",
                "7 Gün Tarih Aralığı",
            ]
            pd.DataFrame([
                ["VARIABLE", 100, "1-3 Eylül", 7, "3-8 Eylül", 13, "Özel 7 Günlük Fiyat"],
                ["TOO_HIGH", 100, "1-3 Eylül", 7, "3-8 Eylül", 13, "Özel 7 Günlük Fiyat"],
            ], columns=columns).to_excel(plus, index=False)
            accounting = root / "accounting-plus.xlsx"
            pd.DataFrame([
                {"Barkod": "TOO_HIGH", "Plus Fiyat Üst Limiti": 110},
            ]).to_excel(accounting, index=False)
            inputs.update({"plus": plus, "muhasebe_plus": accounting})

            rows = {
                row["Barkod"]: row
                for row in calculate_all(inputs, output_dir=root / "output")["results"]
            }

        variable = rows["VARIABLE"]
        self.assertEqual(variable["Plus Tarife Seçimi"], "Özel 7 Günlük Fiyat")
        self.assertEqual(variable["Plus Komisyon (2 Gün) (%)"], 7)
        self.assertEqual(variable["Plus Komisyon (5 Gün) (%)"], 13)
        self.assertEqual((variable["Plus Komisyon (%)"], variable["Plus Net (TL)"]), (13, 87))

        too_high = rows["TOO_HIGH"]
        self.assertEqual(too_high["Plus Fiyatı (TL)"], 110)
        self.assertIsNone(too_high["Plus Tarife Seçimi"])
        self.assertNotIn("Plus", too_high["eligible_main_campaigns"])

    def test_old_single_period_and_accounting_only_plus_remain_supported(self):
        from komisyon_hesaplayici import calculate_all

        with tempfile.TemporaryDirectory() as temp_dir:
            root = Path(temp_dir)
            inputs = write_calculation_base(root, ("OLD", "ACCOUNTING"), current_rate=15)
            plus = root / "plus.xlsx"
            pd.DataFrame([{
                "Barkod": "OLD",
                "Plus Fiyat Üst Limiti": 100,
                "Tarih Aralığı (7 Gün)": "4-11 Ağustos",
                "Plus Komisyon Teklifi": 9,
                "Plus Fiyat Seçimi": None,
                "Tarife Seçimi": None,
            }]).to_excel(plus, index=False)
            accounting = root / "accounting-plus.xlsx"
            pd.DataFrame([{
                "Barkod": "ACCOUNTING",
                "Plus Fiyat Üst Limiti": 100,
            }]).to_excel(accounting, index=False)
            inputs.update({"plus": plus, "muhasebe_plus": accounting})

            rows = {
                row["Barkod"]: row
                for row in calculate_all(inputs, output_dir=root / "output")["results"]
            }

        old = rows["OLD"]
        self.assertEqual(old["Plus Tarife Seçimi"], "7 Günlük Fiyat")
        self.assertEqual((old["Plus Komisyon (%)"], old["Plus Net (TL)"]), (9, 91))
        self.assertEqual(old["Plus Komisyon (7 Gün) (%)"], 9)
        self.assertIn("Plus", old["eligible_main_campaigns"])

        accounting_only = rows["ACCOUNTING"]
        self.assertEqual(accounting_only["Plus Tarife Seçimi"], "7 Günlük Fiyat")
        self.assertEqual(
            (accounting_only["Plus Komisyon (%)"], accounting_only["Plus Net (TL)"]),
            (15, 85),
        )
        self.assertEqual(accounting_only["Plus Komisyon (7 Gün) (%)"], 15)
        self.assertIn("Plus", accounting_only["eligible_main_campaigns"])

    def test_flash_prices_are_evaluated_per_interval_with_exact_price_and_floor(self):
        from komisyon_hesaplayici import calculate_all

        with tempfile.TemporaryDirectory() as temp_dir:
            root = Path(temp_dir)
            barcodes = (
                "MULTI", "FIXED", "FLOOR", "ALL_BELOW", "ACCOUNTING", "LEGACY",
            )
            inputs = write_calculation_base(root, barcodes, current_price=1200)
            pd.DataFrame([
                {
                    "BARKOD": barcode,
                    "1.Fiyat Alt Limit": 965,
                    "2.Fiyat Üst Limiti": 964.99,
                    "2.Fiyat Alt Limit": 0,
                    "3.Fiyat Üst Limiti": None,
                    "3.Fiyat Alt Limit": None,
                    "4.Fiyat Üst Limiti": None,
                    "1.KOMİSYON": 10,
                    "2.KOMİSYON": 20,
                    "3.KOMİSYON": 30,
                    "4.KOMİSYON": 40,
                }
                for barcode in barcodes
            ]).to_excel(inputs["commission"], index=False)
            pd.DataFrame([
                {
                    "Barkod": "MULTI",
                    "24 Saat Fiyat": 968.81,
                    "24 Saat Flaş Başlangıç Tarihi": "10/08/2026 00:00",
                    "24 Saat Flaş Bitiş Tarihi": "10/08/2026 23:59",
                },
                {
                    "Barkod": "MULTI",
                    "24 Saat Fiyat": 960.64,
                    "24 Saat Flaş Başlangıç Tarihi": "12/08/2026 00:00",
                    "24 Saat Flaş Bitiş Tarihi": "12/08/2026 23:59",
                },
                {
                    "Barkod": "FIXED",
                    "24 Saat Fiyat": 980,
                    "3 Saat Fiyat": 970,
                    "Senin Belirlediğin Flaş Fiyatı": 945,
                    "24 Saat Flaş Başlangıç Tarihi": "13/08/2026 00:00",
                    "24 Saat Flaş Bitiş Tarihi": "13/08/2026 23:59",
                    "3 Saat Flaş Başlangıç Tarihi": "13/08/2026 20:00",
                    "3 Saat Flaş Bitiş Tarihi": "13/08/2026 22:59",
                },
                {
                    "Barkod": "FLOOR",
                    "24 Saat Fiyat": 1000,
                    "24 Saat Flaş Başlangıç Tarihi": "14/08/2026 00:00",
                    "24 Saat Flaş Bitiş Tarihi": "14/08/2026 23:59",
                },
                {
                    "Barkod": "FLOOR",
                    "24 Saat Fiyat": 70,
                    "24 Saat Flaş Başlangıç Tarihi": "15/08/2026 00:00",
                    "24 Saat Flaş Bitiş Tarihi": "15/08/2026 23:59",
                },
                {
                    "Barkod": "ALL_BELOW",
                    "24 Saat Fiyat": 70,
                    "24 Saat Flaş Başlangıç Tarihi": "16/08/2026 00:00",
                    "24 Saat Flaş Bitiş Tarihi": "16/08/2026 23:59",
                },
                {
                    "Barkod": "ACCOUNTING",
                    "24 Saat Fiyat": 980,
                    "24 Saat Flaş Başlangıç Tarihi": "17/08/2026 00:00",
                    "24 Saat Flaş Bitiş Tarihi": "17/08/2026 23:59",
                },
                {
                    "Barkod": "ACCOUNTING",
                    "24 Saat Fiyat": 970,
                    "24 Saat Flaş Başlangıç Tarihi": "18/08/2026 00:00",
                    "24 Saat Flaş Bitiş Tarihi": "18/08/2026 23:59",
                },
                {"Barkod": "LEGACY"},
            ]).to_excel(root / "flash.xlsx", index=False)
            pd.DataFrame([{
                "Barkod": "ACCOUNTING",
                "24 Saat Fiyat": 955,
                "Senin Belirlediğin Flaş Fiyatı": 955,
                "24 Saat Flaş Başlangıç Tarihi": "18/08/2026 00:00",
                "24 Saat Flaş Bitiş Tarihi": "18/08/2026 23:59",
            }]).to_excel(root / "accounting-flash.xlsx", index=False)
            inputs.update({
                "flash": root / "flash.xlsx",
                "muhasebe_flas": root / "accounting-flash.xlsx",
            })

            rows = {
                row["Barkod"]: row
                for row in calculate_all(inputs, output_dir=root / "output")["results"]
            }

        multi = rows["MULTI"]
        self.assertEqual(
            [(item["start"], item["price"], item["rate"], item["net"])
             for item in multi["flash_evaluations"]],
            [
                ("2026-08-10 00:00:00", 968.81, 10, 871.93),
                ("2026-08-12 00:00:00", 960.64, 20, 768.51),
            ],
        )
        self.assertEqual(
            (multi["Flaş Ürün 24 Saat Fiyatı (TL)"], multi["Flaş Ürün Komisyon (%)"]),
            (960.64, 20),
        )

        fixed = rows["FIXED"]
        self.assertEqual(
            [(item["period"], item["price"], item["source"])
             for item in fixed["flash_evaluations"]],
            [
                ("24 Saat", 945, "Senin Belirlediğin Flaş Fiyatı"),
                ("3 Saat", 945, "Senin Belirlediğin Flaş Fiyatı"),
            ],
        )

        floor = rows["FLOOR"]
        self.assertEqual(
            [(item["price"], item["eligible"]) for item in floor["flash_evaluations"]],
            [(1000, True), (70, False)],
        )
        self.assertEqual(floor["Flaş Ürün 24 Saat Fiyatı (TL)"], 1000)
        self.assertIn("Flaş", floor["eligible_main_campaigns"])
        self.assertNotIn("Flaş", rows["ALL_BELOW"]["eligible_main_campaigns"])
        self.assertEqual(rows["ALL_BELOW"]["Flaş Ürün 24 Saat Fiyatı (TL)"], 70)

        legacy = rows["LEGACY"]
        self.assertEqual(legacy["Flaş Ürün 24 Saat Fiyatı (TL)"], 1200)
        self.assertEqual(legacy["flash_evaluations"][0]["source"], "Mevcut Fiyat")
        self.assertNotIn("Flaş", legacy["eligible_main_campaigns"])

        accounting = rows["ACCOUNTING"]
        self.assertEqual(
            [(item["start"], item["price"], item["origin"], item["source"])
             for item in accounting["flash_evaluations"]],
            [
                ("2026-08-17 00:00:00", 980, "kampanya", "24 Saat Fiyat"),
                (
                    "2026-08-18 00:00:00",
                    955,
                    "muhasebe",
                    "Senin Belirlediğin Flaş Fiyatı",
                ),
            ],
        )

    def test_undated_accounting_flash_price_is_ignored_for_multiple_standard_intervals(self):
        from komisyon_hesaplayici import calculate_all

        with tempfile.TemporaryDirectory() as temp_dir:
            root = Path(temp_dir)
            inputs = write_calculation_base(root, ("FIXED",), current_price=200, current_rate=10)
            pd.DataFrame([
                {
                    "Barkod": "FIXED",
                    "24 Saat Fiyat": 120,
                    "24 Saat Flaş Başlangıç Tarihi": "10/08/2026 00:00",
                    "24 Saat Flaş Bitiş Tarihi": "10/08/2026 23:59",
                },
                {
                    "Barkod": "FIXED",
                    "24 Saat Fiyat": 110,
                    "24 Saat Flaş Başlangıç Tarihi": "12/08/2026 00:00",
                    "24 Saat Flaş Bitiş Tarihi": "12/08/2026 23:59",
                },
            ]).to_excel(root / "flash.xlsx", index=False)
            pd.DataFrame([{
                "Barkod": "FIXED",
                "Senin Belirlediğin Flaş Fiyatı": 95,
            }]).to_excel(root / "accounting-flash.xlsx", index=False)
            inputs.update({
                "flash": root / "flash.xlsx",
                "muhasebe_flas": root / "accounting-flash.xlsx",
            })

            row = calculate_all(inputs, output_dir=root / "output")["results"][0]

        self.assertEqual(
            [(item["start"], item["price"], item["origin"], item["source"])
             for item in row["flash_evaluations"]],
            [
                ("2026-08-10 00:00:00", 120, "kampanya", "24 Saat Fiyat"),
                ("2026-08-12 00:00:00", 110, "kampanya", "24 Saat Fiyat"),
            ],
        )
        self.assertEqual(row["Flaş Ürün 24 Saat Fiyatı (TL)"], 110)

    def test_plus_fixed_selection_precedes_upper_limit_for_standard_and_accounting(self):
        from komisyon_hesaplayici import calculate_all

        with tempfile.TemporaryDirectory() as temp_dir:
            root = Path(temp_dir)
            inputs = write_calculation_base(root, ("STANDARD", "ACCOUNTING"))
            plus = root / "plus.xlsx"
            pd.DataFrame([
                {
                    "Barkod": "STANDARD",
                    "Plus Fiyat Seçimi": 90,
                    "Plus Fiyat Üst Limiti": 100,
                    "Tarih Aralığı (7 Gün)": "11-18 Ağustos",
                    "Plus Komisyon Teklifi": 10,
                },
                {
                    "Barkod": "ACCOUNTING",
                    "Plus Fiyat Seçimi": None,
                    "Plus Fiyat Üst Limiti": 100,
                    "Tarih Aralığı (7 Gün)": "11-18 Ağustos",
                    "Plus Komisyon Teklifi": 10,
                },
            ]).to_excel(plus, index=False)
            accounting = root / "accounting-plus.xlsx"
            pd.DataFrame([{
                "Barkod": "ACCOUNTING",
                "Plus Fiyat Seçimi": 85,
                "Plus Fiyat Üst Limiti": 95,
            }]).to_excel(accounting, index=False)
            inputs.update({"plus": plus, "muhasebe_plus": accounting})

            rows = {
                row["Barkod"]: row
                for row in calculate_all(inputs, output_dir=root / "output")["results"]
            }

        self.assertEqual(rows["STANDARD"]["Plus Fiyatı (TL)"], 90)
        self.assertEqual(rows["ACCOUNTING"]["Plus Fiyatı (TL)"], 85)


if __name__ == "__main__":
    unittest.main()
