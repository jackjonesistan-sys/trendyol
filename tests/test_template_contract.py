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

    def test_theme_toggle_is_persistent_and_uses_daisyui(self):
        self.assertIn('id="themeToggle"', self.template)
        self.assertIn("class=\"theme-controller toggle toggle-sm toggle-primary\"", self.template)
        self.assertIn("localStorage.setItem('theme', theme)", self.template)
        self.assertIn("document.documentElement.dataset.theme", self.template)
        self.assertIn("window.matchMedia('(prefers-color-scheme: dark)')", self.template)

    def test_ui_accessibility_and_mobile_layout_contract(self):
        self.assertIn('<meta name="description"', self.template)
        self.assertIn('rel="icon" href="data:image/svg+xml,', self.template)
        self.assertIn('role="status"', self.template)
        self.assertIn('aria-live="polite"', self.template)
        self.assertIn('aria-busy="true"', self.template)
        self.assertIn('id="toastContainer" aria-live="polite"', self.template)
        self.assertIn('<caption class="sr-only">', self.template)
        self.assertIn('aria-label="Tüm ürünleri seç"', self.template)
        self.assertIn('aria-label="Ürünü seç"', self.template)
        self.assertIn('aria-label="Ana kampanya toplu seçimi"', self.template)
        self.assertIn('aria-label="Ekstra kampanya toplu seçimi"', self.template)
        self.assertIn('aria-label="{{ spec.label }} Excel dosyası"', self.template)
        self.assertIn('aria-label="${escapeHtml(row.Barkod || \'Ürün\')} ana kampanya seçimi"', self.template)
        self.assertIn('aria-label="Ürün veya barkod ara"', self.template)
        self.assertIn('aria-label="Minimum fiyat"', self.template)
        self.assertIn('aria-label="Maksimum fiyat"', self.template)
        self.assertIn('[data-theme="dark"] .text-base-content\\/40', self.template)
        self.assertIn('focus-visible:outline', self.template)
        self.assertIn('max-w-[1800px] mx-auto', self.template)
        self.assertIn('href="#excel-girdileri"', self.template)
        self.assertIn('href="#kampanya-tablosu"', self.template)

    def test_base_inputs_text_names_three_required_files(self):
        self.assertIn("zorunludur. İndirim Uygulanabilecek Ürünler de zorunludur.", self.template)
        self.assertIn("* Zorunlu 3 Girdi", self.template)

    def test_campaign_actions_have_one_implementation(self):
        self.assertEqual(self.template.count("function updateSelection("), 1)
        self.assertEqual(self.template.count("function applyBulkAction("), 1)
        self.assertEqual(self.template.count("function autoSelect("), 1)
        self.assertIn(
            "const counter = row.counter_evaluations?.[extraSel];", self.template
        )
        self.assertIn("main: row.userSelection || 'Hiçbiri'", self.template)
        self.assertIn("extra: row.userExtraSelection || 'Hiçbiri'", self.template)

    def test_dark_mode_uses_semantic_datatable_colors(self):
        self.assertIn("--campaign-advantage-bg: color-mix", self.template)
        self.assertIn("table.dataTable tbody tr:hover", self.template)
        self.assertIn("var(--color-success)", self.template)
        self.assertIn("var(--color-info)", self.template)
        self.assertIn("var(--color-warning)", self.template)

    def test_all_dynamic_discounts_support_tl_and_percent_configs(self):
        self.assertIn("Ek İndirim (TL / % · Çoklu Yükleme)", self.template)
        self.assertIn("function dynamicCampaignLabel(type, item, index)", self.template)
        self.assertIn("yüzde[\\s_-]*", self.template)
        self.assertIn("const fallbackRate = !percentMatch && !tlMatch", self.template)
        self.assertIn("discount_type: item.discount_type || item.discount_unit || '%'", self.template)
        self.assertIn("discount_amount: parseFloat(item.discount_amount ?? item.rate) || 0", self.template)
        self.assertIn("min_price: parseFloat(item.min_price) || 0", self.template)
        self.assertIn("trendyol_percent: parseFloat(item.trendyol_percent) || 0", self.template)
        self.assertIn("updatePlusExtraParam('${item.id}', 'discount_type', this.value)", self.template)
        self.assertIn("updateCouponParam('${item.id}', 'discount_type', this.value)", self.template)
        self.assertIn("Ek İndirim Tutarı / Oranı", self.template)
        self.assertIn("Kupon Tutarı / Oranı", self.template)
        self.assertIn("Trendyol Karşılama (%)", self.template)
        self.assertIn("discount_type: item.discount_type || '%'", self.template)
        self.assertIn("discount_type: item.discount_type || 'TL'", self.template)
        self.assertIn("Array.isArray(data.plus_extra_configs)", self.template)
        self.assertIn(":not(#input_coupon_multi)", self.template)

    def test_dynamic_campaign_selection_uses_customer_price_and_seller_net(self):
        self.assertIn("asNumber(counter.customer_price)", self.template)
        self.assertIn("asNumber(counter.net)", self.template)
        self.assertIn(
            "basePrice - (basePrice * (commRate / 100)) - sellerDisc",
            self.template,
        )

    def test_dynamic_discount_campaigns_are_bulk_extra_options(self):
        option = '<option value="${escapeHtml(cName)}" data-dynamic="true">'
        self.assertIn("selectExtra.append(`" + option, self.template)
        self.assertNotIn("select.append(`" + option, self.template)


if __name__ == "__main__":
    unittest.main()
