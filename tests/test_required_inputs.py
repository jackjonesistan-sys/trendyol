import tempfile
import unittest
from pathlib import Path

from input_files import InputValidationError, INPUT_SPECS, save_upload_set


class RequiredInputTests(unittest.TestCase):
    def test_discount_is_a_required_base_input(self):
        self.assertTrue(INPUT_SPECS["discount"]["required"])

        class Upload:
            filename = "placeholder.xlsx"

        with tempfile.TemporaryDirectory() as temp_dir:
            root = Path(temp_dir)
            with self.assertRaisesRegex(InputValidationError, "İndirim Uygulanabilecek Ürünler"):
                save_upload_set(
                    {"commission": Upload(), "current": Upload()},
                    root / "uploads",
                    root / "manifest.json",
                )

    def test_calculation_core_requires_discount_too(self):
        from komisyon_hesaplayici import calculate_all

        result = calculate_all({"commission": "unused.xlsx", "current": "unused.xlsx"})
        self.assertFalse(result["success"])
        self.assertIn("Zorunlu", result["message"])


if __name__ == "__main__":
    unittest.main()
