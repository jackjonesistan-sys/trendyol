import json
import os
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch
import pandas as pd

import app
from input_files import load_upload_status, save_single_file_enabled, save_upload_set


class SingleFileToggleTests(unittest.TestCase):
    def setUp(self):
        app.app.config["TESTING"] = True
        self.client = app.app.test_client()

    def test_single_file_toggle_endpoint_and_calculation_exemption(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            root = Path(temp_dir)
            manifest_path = root / "manifest.json"
            upload_dir = root / "uploads"
            output_dir = root / "output"
            hesap_path = output_dir / "Kampanya_Hesaplama_Sonuclari.xlsx"
            upload_dir.mkdir(parents=True, exist_ok=True)
            output_dir.mkdir(parents=True, exist_ok=True)

            discount_p = upload_dir / "discount.xlsx"
            commission_p = upload_dir / "commission.xlsx"
            current_p = upload_dir / "current.xlsx"
            advantage_p = upload_dir / "advantage.xlsx"

            pd.DataFrame([{"BARKOD": "B1", "Eski Fiyat": 100, "YENİ Fiyat": 80, "Durum": "İndirim"}]).to_excel(discount_p, index=False)
            pd.DataFrame([{
                "Barkod": "B1",
                "1.Fiyat Alt Limit": 95,
                "1.KOMİSYON": 15,
                "2.Fiyat Üst Limiti": 90,
                "2.KOMİSYON": 12,
                "3.Fiyat Üst Limiti": 85,
                "3.KOMİSYON": 10,
                "4.Fiyat Üst Limiti": 80,
                "4.KOMİSYON": 8,
                "Tarife Seçimi": "7 Günlük Fiyat"
            }]).to_excel(commission_p, index=False)
            pd.DataFrame([{
                "Barkod": "B1",
                "Komisyon Oranı": 20,
                "Piyasa Satış Fiyatı (KDV Dahil)": 150,
                "Trendyol'da Satılacak Fiyat (KDV Dahil)": 100,
            }]).to_excel(current_p, index=False)
            pd.DataFrame([{
                "BARKOD": "B1",
                "1 YILDIZ ÜST FİYAT": 100,
                "YENİ TSF (FİYAT GÜNCELLE)": 90,
            }]).to_excel(advantage_p, index=False)

            manifest_path.write_text(
                json.dumps({
                    "files": {
                        "discount": {"stored_name": "discount.xlsx", "original_name": "discount.xlsx", "enabled": True},
                        "commission": {"stored_name": "commission.xlsx", "original_name": "commission.xlsx", "enabled": True},
                        "current": {"stored_name": "current.xlsx", "original_name": "current.xlsx", "enabled": True},
                        "advantage": {"stored_name": "advantage.xlsx", "original_name": "advantage.xlsx", "enabled": True},
                    },
                    "user_selections": {},
                }),
                encoding="utf-8",
            )

            with (
                patch.object(app, "UPLOAD_DIR", str(upload_dir)),
                patch.object(app, "OUTPUT_DIR", str(output_dir)),
                patch.object(app, "F_HESAP", str(hesap_path)),
                patch.object(app, "INPUT_MANIFEST", str(manifest_path)),
            ):
                # 1. Check toggle endpoint
                res = self.client.post(
                    "/api/toggle-campaign-enabled",
                    json={"type": "single_file", "id": "advantage", "enabled": False},
                )
                self.assertEqual(res.status_code, 200)
                status = load_upload_status(upload_dir, manifest_path)
                self.assertFalse(status["advantage"]["enabled"])

                # 2. Run calculate with advantage disabled
                res_calc = self.client.post("/api/calculate")
                self.assertEqual(res_calc.status_code, 200)
                data = res_calc.get_json()
                results = data["results"]
                row_b1 = next((r for r in results if r["Barkod"] == "B1"), None)
                self.assertIsNotNone(row_b1)
                # Because advantage is disabled, it should NOT be in eligible_campaigns
                self.assertNotIn("Avantajlı", row_b1.get("eligible_campaigns", []))

                # 3. Re-enable advantage
                res_toggle_on = self.client.post(
                    "/api/toggle-campaign-enabled",
                    json={"type": "single_file", "id": "advantage", "enabled": True},
                )
                self.assertEqual(res_toggle_on.status_code, 200)
                res_calc_2 = self.client.post("/api/calculate")
                self.assertEqual(res_calc_2.status_code, 200)
                results_2 = res_calc_2.get_json()["results"]
                row_b1_2 = next((r for r in results_2 if r["Barkod"] == "B1"), None)
                self.assertIn("Avantajlı", row_b1_2.get("eligible_campaigns", []))


if __name__ == "__main__":
    unittest.main()
