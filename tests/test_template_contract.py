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
            "const filterColumnNames = ['Uygulanan Kampanya', 'Hangisi Karlı?'];",
            self.template,
        )
        self.assertIn('data-report-column="{{ column }}"', self.template)
        self.assertIn("th.dataset.reportColumn === columnName", self.template)

    def test_recommendation_button_applies_the_best_campaign_candidate(self):
        self.assertIn(
            "row.userSelection = row['İlk Kampanya Seçimi'] || 'Hiçbiri';",
            self.template,
        )
        self.assertIn("let selectedCount = 0;", self.template)
        self.assertIn("${selectedCount} ürüne en yüksek netli kampanya seçimi uygulandı.", self.template)


if __name__ == "__main__":
    unittest.main()
