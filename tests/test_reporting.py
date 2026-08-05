import os
import sys
import tempfile
import types
import unittest
from pathlib import Path


if "flask" not in sys.modules:
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
        report = app.build_report_row(
            {
                "Barkod": "A1",
                "userSelection": "Plus Ek İndirim %10",
                "İndirim Uygulanabilir": "Hayır",
                "Güncel Ürün Fiyatı (TL)": 100,
                "Güncel Ürün Kalan Net (TL)": 90,
                "Güncel Ürün Komisyon (%)": 10,
                "Plus Ek Fiyatı %10 (TL)": 100,
                "Plus Ek Net %10 (TL)": 80,
                "Plus Ek Komisyon (%)": 10,
            }
        )

        self.assertEqual(report["Plus Ek İndirim Fiyat (TL)"], 90)
        self.assertEqual(report["Uygulanan Kampanya Fiyat"], 90)
        self.assertEqual(app.discounted_price(100, 10), 90)

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

    def test_result_must_not_predate_the_uploaded_input_manifest(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            root = Path(temp_dir)
            manifest = root / "manifest.json"
            result = root / "result.xlsx"
            result.touch()
            manifest.touch()
            old_result, old_manifest = app.F_HESAP, app.INPUT_MANIFEST
            app.F_HESAP, app.INPUT_MANIFEST = str(result), str(manifest)
            try:
                os.utime(result, (3, 3))
                os.utime(manifest, (2, 2))
                self.assertTrue(app.calculation_result_is_current())
                os.utime(manifest, (4, 4))
                self.assertFalse(app.calculation_result_is_current())
            finally:
                app.F_HESAP, app.INPUT_MANIFEST = old_result, old_manifest


if __name__ == "__main__":
    unittest.main()
