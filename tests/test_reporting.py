import os
import sys
import tempfile
import types
import unittest
from pathlib import Path

import pandas as pd

try:
    import flask  # noqa: F401
except ImportError:
    flask_stub = types.ModuleType("flask")

    class FlaskStub:
        def __init__(self, _name):
            self.config = {}

        def route(self, *_args, **_kwargs):
            return lambda function: function

        def errorhandler(self, *_args, **_kwargs):
            return lambda function: function

    flask_stub.Flask = FlaskStub
    flask_stub.render_template = lambda *args, **kwargs: None
    flask_stub.jsonify = lambda value: value
    flask_stub.request = types.SimpleNamespace()
    flask_stub.send_from_directory = lambda *args, **kwargs: None
    sys.modules["flask"] = flask_stub

import app


class ReportingTests(unittest.TestCase):
    def test_report_columns_are_the_page_and_excel_contract(self):
        self.assertEqual(
            app.REPORT_COLUMNS,
            [
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
            ],
        )

    def test_discount_fields_follow_current_selected_and_dip_prices(self):
        row = {
            "Barkod": "A1",
            "userSelection": "Avantajlı",
            "İndirim Uygulanabilir": "Evet",
            "Güncel Ürün Fiyatı (TL)": 100,
            "Güncel Ürün Kalan Net (TL)": 90,
            "Güncel Ürün Komisyon (%)": 10,
            "Avantajlı Ürün Fiyatı (YENİ TSF) (TL)": 90,
            "Avantajlı Ürün Kalan Net (TL)": 81,
            "Avantajlı Ürün Komisyon (%)": 10,
            "Düşülebilecek Dip Fiyat (TL)": 80,
        }

        report = app.build_report_row(row)

        self.assertEqual(report["Uygulanabilecek İndirim (TL)"], 20)
        self.assertEqual(report["Uygulanabilecek İndirim (%)"], 20)
        self.assertEqual(report["Uygulanan İndirim (TL)"], 10)
        self.assertEqual(report["Uygulanan İndirim (%)"], 10)
        self.assertEqual(report["Ekstra Uygulanabilir İndirim (TL)"], 10)
        self.assertEqual(report["Ekstra Uygulanabilir İndirim (%)"], 10)

    def test_plus_extra_uses_customer_price_after_selected_rate(self):
        row = {
            "Barkod": "A1",
            "userExtraSelection": "Plus Ek İndirim %10",
            "İndirim Uygulanabilir": "Hayır",
            "Güncel Ürün Fiyatı (TL)": 100,
            "Güncel Ürün Kalan Net (TL)": 90,
            "Güncel Ürün Komisyon (%)": 10,
            "Plus Ek Fiyatı %10 (TL)": 90,
            "Plus Ek Net %10 (TL)": 80,
            "Plus Ek Komisyon (%)": 10,
        }
        report = app.build_report_row(row)

        self.assertEqual(report["Plus Ek İndirim Fiyat (TL)"], 90)
        self.assertEqual(report["Plus Ek İndirim Net"], 80)
        self.assertEqual(report["Uygulanan Kampanya Fiyat"], 90)
        self.assertEqual(app.selected_campaign_values(row), (90, 80, 10))
        self.assertEqual(app.discounted_price(100, 10), 90)

    def test_advanced_plus_extra_uses_persisted_evaluation_for_report_values(self):
        label = (
            "Plus Ek İndirim (300 TL Üzeri / 50 TL İndirim / "
            "%60 Trendyol Karşılamalı)"
        )
        row = {
            "Barkod": "A1",
            "userExtraSelection": label,
            "Güncel Ürün Fiyatı (TL)": 400,
            "Güncel Ürün Kalan Net (TL)": 360,
            "Güncel Ürün Komisyon (%)": 10,
            "counter_evaluations": {
                label: {
                    "price": 400,
                    "customer_price": 350,
                    "rate": 10,
                    "net": 330,
                    "seller_disc": 20,
                    "disc_type": "TL",
                    "disc_val": 50,
                    "trendyol_percent": 60,
                }
            },
        }

        self.assertEqual(app.selected_campaign_values(row), (350, 330, 10))
        report = app.build_report_row(row)
        self.assertEqual(report["Plus Ek İndirim Fiyat (TL)"], 350)
        self.assertEqual(report["Plus Ek İndirim Net"], 330)
        self.assertEqual(report["Uygulanan Kampanya Fiyat"], 350)
        self.assertEqual(report["Uygulanan Kampanya Net"], 330)

    def test_main_plus_extra_recomposition_keeps_trendyol_funded_discount_in_net(self):
        label = "Plus Ek İndirim Gelişmiş"
        row = {
            "userSelection": "Plus",
            "userExtraSelection": label,
            "Plus Fiyatı (TL)": 400,
            "Plus Net (TL)": 360,
            "Plus Komisyon (%)": 10,
            "Güncel Ürün Fiyatı (TL)": 500,
            "Güncel Ürün Komisyon (%)": 10,
            "counter_evaluations": {
                label: {
                    "price": 500,
                    "customer_price": 450,
                    "rate": 10,
                    "net": 430,
                    "seller_disc": 20,
                    "disc_type": "%",
                    "disc_val": 10,
                    "trendyol_percent": 60,
                }
            },
        }

        self.assertEqual(app.selected_campaign_values(row), (360, 344, 10))

    def test_visible_columns_keep_contract_order(self):
        requested = ["Hangisi Karlı?", "Barkod", "Güncel Net"]
        self.assertEqual(
            app.normalize_visible_columns(requested),
            ["Barkod", "Güncel Net", "Hangisi Karlı?"],
        )

    def test_campaign_selection_must_be_in_server_calculated_applicable_set(self):
        self.assertTrue(
            app.campaign_selection_is_applicable(
                "Plus Ek İndirim %10", "Avantajlı, Plus Ek İndirim"
            )
        )
        self.assertFalse(
            app.campaign_selection_is_applicable("Flaş", "Avantajlı, Plus")
        )
        coupon = "500 TL Üzerine 50 TL Kupon - Trendyol Plus Müşterilerine Özel"
        self.assertTrue(app.campaign_selection_is_applicable(coupon, coupon))

    def test_excel_persisted_campaign_collections_restore_their_types(self):
        self.assertEqual(
            app.parse_persisted_collection("['Hiçbiri', 'Avantajlı']", list),
            ["Hiçbiri", "Avantajlı"],
        )
        self.assertEqual(
            app.parse_persisted_collection(
                "{'Karşılamalı': {'price': 90, 'net': 81}}", dict
            ),
            {"Karşılamalı": {"price": 90, "net": 81}},
        )
        self.assertEqual(app.parse_persisted_collection(float("nan"), list), [])

    def test_dynamic_counter_selection_payload_keeps_string_validation(self):
        self.assertTrue(
            app.selection_payload_is_valid(
                {"A1": "Karşılamalı (300 TL Üzeri / 50 TL İndirim)"}
            )
        )
        self.assertFalse(app.selection_payload_is_valid({"A1": 50}))

    def test_result_must_not_predate_the_uploaded_input_manifest(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            root = Path(temp_dir)
            manifest = root / "manifest.json"
            result = root / "result.xlsx"
            result.write_bytes(b"data")
            old_result = app.F_HESAP
            app.F_HESAP = str(result)
            try:
                self.assertTrue(app.calculation_result_is_current())
                result.unlink()
                self.assertFalse(app.calculation_result_is_current())
            finally:
                app.F_HESAP = old_result

    def test_price_difference_report_does_not_crash_on_unrelated_indirim_file(self):
        from fiyat_farki_analiz_script import generate_fiyat_farki_raporu

        cwd = os.getcwd()
        with tempfile.TemporaryDirectory() as temp_dir:
            root = Path(temp_dir)
            (root / "Girdiler").mkdir()
            out = root / "out"
            out.mkdir()
            pd.DataFrame([{"Barkod": "A1", "Başka Kolon": 1}]).to_excel(
                root / "Girdiler" / "yanlis-indirim.xlsx", index=False
            )
            pd.DataFrame([{
                "Barkod": "A1",
                "Güncel Ürün Fiyatı (TL)": 100,
                "Avantajlı Ürün Fiyatı (YENİ TSF) (TL)": 90,
            }]).to_excel(out / "Uygulanmayan_Urunler_Raporu.xlsx", index=False)
            os.chdir(root)
            try:
                generate_fiyat_farki_raporu(out)
            finally:
                os.chdir(cwd)

            self.assertTrue((out / "Indirim_Uygulanmayan_Fiyat_Kiyas_Raporu.xlsx").exists())


if __name__ == "__main__":
    unittest.main()
