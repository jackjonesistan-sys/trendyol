import json
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

import app


DEFAULT_RULE = {
    "enabled": True,
    "priority": ["Avantajlı", "Flaş", "Plus"],
}


class RecommendationRuleTests(unittest.TestCase):
    def setUp(self):
        app.app.config["TESTING"] = True
        self.client = app.app.test_client()

    def test_rule_normalization_is_strict_and_has_a_safe_default(self):
        from input_files import InputValidationError, normalize_recommendation_rule

        self.assertEqual(normalize_recommendation_rule(), DEFAULT_RULE)
        self.assertEqual(
            normalize_recommendation_rule({
                "enabled": False,
                "priority": ["Plus", "Avantajlı", "Flaş"],
            }),
            {
                "enabled": False,
                "priority": ["Plus", "Avantajlı", "Flaş"],
            },
        )

        invalid_rules = (
            {},
            {"enabled": True},
            {"priority": DEFAULT_RULE["priority"]},
            {"enabled": 1, "priority": DEFAULT_RULE["priority"]},
            {"enabled": True, "priority": "Avantajlı,Flaş,Plus"},
            {"enabled": True, "priority": ["Avantajlı", "Flaş"]},
            {"enabled": True, "priority": ["Avantajlı", "Flaş", "Flaş"]},
            {"enabled": True, "priority": ["Avantajlı", "Flaş", "Bilinmeyen"]},
            {**DEFAULT_RULE, "unexpected": True},
        )
        for rule in invalid_rules:
            with self.subTest(rule=rule), self.assertRaises(InputValidationError):
                normalize_recommendation_rule(rule)

    def test_rule_save_load_preserves_manifest_and_invalid_rule_does_not_mutate(self):
        from input_files import (
            InputValidationError,
            load_recommendation_rule,
            save_recommendation_rule,
        )

        with tempfile.TemporaryDirectory() as temp_dir:
            manifest = Path(temp_dir) / "manifest.json"
            manifest.write_text(
                json.dumps({"files": {}, "user_selections": {"A1": "Plus"}}),
                encoding="utf-8",
            )
            rule = {
                "enabled": True,
                "priority": ["Plus", "Flaş", "Avantajlı"],
            }

            self.assertEqual(save_recommendation_rule(manifest, rule), rule)
            self.assertEqual(load_recommendation_rule(manifest), rule)
            saved = json.loads(manifest.read_text(encoding="utf-8"))
            self.assertEqual(saved["user_selections"], {"A1": "Plus"})
            before_invalid = manifest.read_bytes()

            with self.assertRaises(InputValidationError):
                save_recommendation_rule(
                    manifest,
                    {"enabled": True, "priority": ["Plus", "Plus", "Flaş"]},
                )
            self.assertEqual(manifest.read_bytes(), before_invalid)

    def test_priority_rule_chooses_first_selectable_main_and_keeps_extra_best_net(self):
        from komisyon_hesaplayici import choose_campaigns_smart

        candidates = [
            ("Avantajlı", 91, 100, 9),
            ("Flaş", 108, 120, 10),
            ("Plus", 104, 110, 5),
            ("Ekstra A", 103, 105, 2),
            ("Ekstra B", 106, 110, 2),
        ]
        rule = {
            "enabled": True,
            "priority": ["Plus", "Avantajlı", "Flaş"],
        }

        selected = choose_campaigns_smart(100, candidates, rule)
        self.assertEqual(selected[0], "Plus")
        self.assertEqual(selected[3], "Ekstra B")

        plus_below_current_net = [
            candidate if candidate[0] != "Plus" else (*candidate, True)
            for candidate in candidates
        ]
        selected = choose_campaigns_smart(105, plus_below_current_net, rule)
        self.assertEqual(selected[0], "Avantajlı")

        selected = choose_campaigns_smart(
            100,
            candidates,
            {"enabled": False, "priority": rule["priority"]},
        )
        self.assertEqual(selected[0], "Flaş")
        self.assertEqual(choose_campaigns_smart(100, candidates)[0], "Avantajlı")

    def test_rule_endpoint_persists_valid_rule_and_rejects_invalid_without_mutation(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            manifest = Path(temp_dir) / "manifest.json"
            manifest.write_text('{"files": {}, "marker": 1}', encoding="utf-8")
            rule = {
                "enabled": True,
                "priority": ["Flaş", "Plus", "Avantajlı"],
            }
            with patch.object(app, "INPUT_MANIFEST", str(manifest)):
                response = self.client.post("/api/recommendation-rule", json=rule)
                self.assertEqual(response.status_code, 200, response.get_json())
                self.assertEqual(response.get_json()["recommendation_rule"], rule)
                self.assertEqual(
                    json.loads(manifest.read_text(encoding="utf-8"))["marker"], 1
                )

                before_invalid = manifest.read_bytes()
                response = self.client.post(
                    "/api/recommendation-rule",
                    json={
                        "enabled": True,
                        "priority": ["Flaş", "Flaş", "Avantajlı"],
                    },
                )

            self.assertEqual(response.status_code, 400)
            self.assertEqual(manifest.read_bytes(), before_invalid)

            with patch.object(app, "INPUT_MANIFEST", str(manifest)):
                response = self.client.post("/api/recommendation-rule")
            self.assertEqual(response.status_code, 400)
            self.assertEqual(manifest.read_bytes(), before_invalid)

    def test_save_selections_rejects_invalid_values_without_mutating_manifest(self):
        invalid_selections = (
            {"A1": {"main": "Bilinmeyen", "extra": "Hiçbiri"}},
            {"A1": {"main": "Avantajlı", "extra": ""}},
            {"A1": 42},
            {"A1": ["Plus"]},
        )
        with tempfile.TemporaryDirectory() as temp_dir:
            manifest = Path(temp_dir) / "manifest.json"
            manifest.write_text(
                json.dumps({
                    "files": {},
                    "user_selections": {
                        "OLD": {"main": "Plus", "extra": "Hiçbiri"},
                    },
                }),
                encoding="utf-8",
            )
            original = manifest.read_bytes()

            with patch.object(app, "INPUT_MANIFEST", str(manifest)):
                for selections in invalid_selections:
                    with self.subTest(selections=selections):
                        response = self.client.post(
                            "/api/save-selections",
                            json={"selections": selections},
                        )
                        self.assertEqual(response.status_code, 400)
                        self.assertEqual(manifest.read_bytes(), original)

    def test_save_selections_persists_canonical_values_and_accepts_legacy_strings(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            manifest = Path(temp_dir) / "manifest.json"
            result = Path(temp_dir) / "missing-result.xlsx"
            manifest.write_text('{"files": {}, "marker": 1}', encoding="utf-8")

            with (
                patch.object(app, "INPUT_MANIFEST", str(manifest)),
                patch.object(app, "F_HESAP", str(result)),
            ):
                response = self.client.post(
                    "/api/save-selections",
                    json={
                        "selections": {
                            "A1": "Plus",
                            "A2": "  Plus Ek Kampanya  ",
                            "A3": {"main": "Flaş", "extra": "  Hiçbiri  "},
                        },
                    },
                )

            self.assertEqual(response.status_code, 200, response.get_json())
            self.assertEqual(response.get_json()["saved_count"], 3)
            saved = json.loads(manifest.read_text(encoding="utf-8"))
            self.assertEqual(saved["marker"], 1)
            self.assertEqual(
                saved["user_selections"],
                {
                    "A1": {"main": "Plus", "extra": "Hiçbiri"},
                    "A2": {"main": "Hiçbiri", "extra": "Plus Ek Kampanya"},
                    "A3": {"main": "Flaş", "extra": "Hiçbiri"},
                },
            )

    def test_present_null_calculation_rule_is_not_treated_as_absent(self):
        from input_files import InputValidationError

        with self.assertRaises(InputValidationError):
            app.parse_recommendation_rule_json("null")

    def test_calculate_persists_supplied_rule_and_passes_it_to_calculator(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            root = Path(temp_dir)
            manifest = root / "manifest.json"
            manifest.write_text('{"files": {}}', encoding="utf-8")
            uploads = root / "uploads"
            output = root / "output"
            uploads.mkdir()
            output.mkdir()
            rule = {
                "enabled": False,
                "priority": ["Plus", "Flaş", "Avantajlı"],
            }
            with (
                patch.object(app, "INPUT_MANIFEST", str(manifest)),
                patch.object(app, "UPLOAD_DIR", str(uploads)),
                patch.object(app, "OUTPUT_DIR", str(output)),
                patch.object(app, "F_HESAP", str(output / "result.xlsx")),
                patch.object(
                    app,
                    "save_upload_set",
                    return_value={
                        "discount": "discount.xlsx",
                        "commission": "commission.xlsx",
                        "current": "current.xlsx",
                    },
                ),
                patch.object(app, "load_upload_status", return_value={}),
                patch(
                    "komisyon_hesaplayici.calculate_all",
                    return_value={"success": True, "results": []},
                ) as calculate_all,
            ):
                response = self.client.post(
                    "/api/calculate",
                    data={"recommendation_rule_json": json.dumps(rule)},
                )

            self.assertEqual(response.status_code, 200, response.get_json())
            self.assertEqual(
                calculate_all.call_args.kwargs["recommendation_rule"], rule
            )
            self.assertEqual(
                json.loads(manifest.read_text(encoding="utf-8"))[
                    "recommendation_rule"
                ],
                rule,
            )


if __name__ == "__main__":
    unittest.main()
