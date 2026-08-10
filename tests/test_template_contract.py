import unittest
from pathlib import Path


class TemplateContractTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.template = (Path(__file__).parents[1] / "templates" / "index.html").read_text(
            encoding="utf-8"
        )
import unittest
from pathlib import Path


class TemplateContractTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.template = (Path(__file__).parents[1] / "templates" / "index.html").read_text(
            encoding="utf-8"
        )

    def test_uploaded_file_status_is_visible_for_each_input(self):
        self.assertIn('id="upload_status_{{ key }}"', self.template)
        self.assertIn("uploaded_at_display", self.template)

    def test_campaign_and_profit_filters_are_restored(self):
        self.assertIn(
            "const filterColumnNames = ['Uygulanan Kampanya', 'Ekstra Kampanya', 'Hangisi Karlı?'];",
            self.template,
        )
        self.assertIn('data-report-column="{{ column }}"', self.template)
        self.assertIn("th.dataset.reportColumn === columnName", self.template)

    def test_recommendation_button_applies_the_best_campaign_candidate(self):
        self.assertIn(
            "const rec = row['Önerilen Kampanya'] || 'Hiçbiri';",
            self.template,
        )
        self.assertIn("let selectedCount = 0;", self.template)
        self.assertIn("ürüne önerilen", self.template)

    def test_template_data_and_missing_numbers_are_safe(self):
        self.assertIn('id="report-columns-data" type="application/json"', self.template)
        self.assertIn(
            "JSON.parse(document.getElementById('report-columns-data').textContent)",
            self.template,
        )
        self.assertNotIn("const reportColumns = {{", self.template)
        self.assertIn(
            "if (value === null || value === undefined || value === '') return null;",
            self.template,
        )

    def test_campaign_actions_have_one_implementation(self):
        self.assertEqual(self.template.count("function updateSelection("), 1)
        self.assertEqual(self.template.count("function applyBulkAction("), 1)
        self.assertIn(
            "const counter = row.counter_evaluations?.[extraSel];", self.template
        )


if __name__ == "__main__":
    unittest.main()
