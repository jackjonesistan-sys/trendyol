# Trendyol Kampanya Yönetimi - Kapsamlı Sistem Spesifikasyonu ve Mimari Dökümü

Bu doküman, "Trendyol Kampanya Yönetimi" web uygulamasını sıfırdan inşa edecek bir yapay zeka (AI Agent) veya yazılımcı için **kusursuzlaştırılmış, eksiksiz bir rehber** niteliği taşımaktadır. Başka bir AI'ın sadece bu dosyayı okuyarak (hiçbir koda bakmadan) sistemi birebir kodlayabilmesi için tasarlanmıştır.

---

## 1. Dizin Hiyerarşisi (Directory Structure)
Sistem dizini tamamen şu şekilde kurgulanmalıdır:
```text
C:\Users\Tasarımcı\Desktop\trendyol\
├── app.py                      # Flask sunucusu ve API endpointleri
├── komisyon_hesaplayici.py     # Veri çekme ve Pandas/OpenPyXL hesaplama algoritmaları
├── templates/
│   └── index.html              # Frontend arayüzü (DaisyUI, DataTables, Kural Motoru)
├── Girdiler/                   # (Klasör) Kullanıcının yüklediği ham Trendyol Excel'leri
├── Çıktılar/                   # (Klasör) Sistemin ürettiği nihai yüklemeye hazır Excel'ler
└── docs/
    └── sistem_spesifikasyonu.md
```

---

## 2. API Endpoint Sözleşmeleri (Contract for AI)

Uygulamanın Backend-Frontend iletişimi %100 JSON tabanlı AJAX istekleriyle (Fetch) yapılmalıdır. AI bu endpointleri birebir kodlamalıdır:

### `GET /api/data`
*   **Amaç:** `Çıktılar/Hesaplanmis_Komisyon_Sonuclari.xlsx` dosyasını okuyup JSON'a çevirir. Eğer bu dosya yoksa `{"needs_calculation": true, "message": "Hesaplama gerekiyor"}` döner.
*   **Örnek Response (tableData element):**
    ```json
    [
      {
        "Barkod": "ST103-4070",
        "Güncel Ürün Fiyatı (TL)": 749,
        "Güncel Ürün Kalan Net (TL)": 434.22,
        "Avantajlı Ürün Fiyatı (YENİ TSF) (TL)": 567.91,
        "Avantajlı Ürün Kalan Net (TL)": 453.19,
        "Flaş Ürün 24 Saat Fiyatı (TL)": "-",
        "Flaş Ürün Kalan Net (TL)": "-",
        "Plus Fiyatı (TL)": 597.51,
        "Plus Net (TL)": 486.97,
        "Mevcut İndirim Oranı (%)": 5.2,
        "Plus Ek Fiyatı %5 (TL)": 749,
        "Plus Ek Net %5 (TL)": 584.22,
        "Plus Ek Fiyatı %10 (TL)": 749,
        "Plus Ek Net %10 (TL)": 584.22,
        "Plus Ek Fiyatı %20 (TL)": 749,
        "Plus Ek Net %20 (TL)": 584.22,
        "Avantajlı Ürün Eşleşme Durumu": "Eşleşti",
        "Flaş Ürün Eşleşme Durumu": "Bulunamadı",
        "Plus Eşleşme Durumu": "Eşleşti",
        "Plus Ek İndirim Eşleşme Durumu": "Eşleşti",
        "Hangisi Daha Karlı?": "Plus Ürün"
      }
    ]
    ```

### `POST /api/calculate`
*   **Amaç:** `komisyon_hesaplayici.py` dosyasındaki ana fonksiyonu asenkron çalıştırıp, girdi dosyalarını okuyup yukarıdaki JSON yapısını üreten `.xlsx` çıktısını yaratmaktır.
*   **Payload:** Boş.
*   **Response:** `{"success": true}` veya `{"success": false, "message": "Error..."}`

### `POST /api/apply_campaigns`
*   **Amaç:** Frontend'den gelen seçilmiş ürünler listesine göre `Girdiler/` klasöründeki şablonları klonlayıp satırları silerek `Çıktılar/` klasörüne yazmak.
*   **Örnek Payload:**
    ```json
    {
      "target_type": "Flaş", 
      "avanStartDate": "2026-07-24", 
      "avanEndDate": "2026-07-28",
      "tableData": [
        { "Barkod": "ST103", "userSelection": "Flaş" },
        { "Barkod": "ST104", "userSelection": "Hiçbiri" }
      ]
    }
    ```
*   **Not:** `target_type` "Hepsi" ise tüm kampanyalar (Avantajlı, Flaş, vb.) sırayla üretilir.

---

## 3. Veri Birleştirme (Merging) ve DataFrame Logiği (AI için Formüller)

`komisyon_hesaplayici.py` içindeki mantık şu şekilde inşa edilmelidir:

### A. Havuz Oluşturma (Union)
```python
# Tüm girdiler pandas ile okunur (pd.read_excel).
# 'Barkod' veya 'BARKOD' veya 'Barkodu' sütunları normalize edilir (str.strip, upper).
b1 = df_guncel['BARKOD_CLN'].unique()
b2 = df_indirim['BARKOD_CLN'].unique()
all_barcodes = set(b1).union(set(b2))
# Daha sonra her bir df 'BARKOD_CLN' indexine set edilir (df.set_index('BARKOD_CLN'))
# Döngü tüm all_barcodes içinde döner ve `row = df.loc[barkod]` ile veriler eşleştirilir.
```

### B. Dinamik Kâr Formülü (Çok Kritik!)
Komisyon tarifesine göre net hesaplanırken: `Net = Fiyat - (Fiyat * (Komisyon_Orani / 100))` formülü kullanılır.
*   **Komisyon Oranı Nereden Bulunur?** Sınıflandırılmış komisyon tarifesinde ürün fiyatı aranır. Örn: Fiyat 150 TL ise ve komisyon tarifesinde `2.Fiyat Alt Limit=100`, `2.Fiyat Üst Limiti=200` ise oran `2.KOMİSYON`'dur.

### C. "Hangisi Daha Karlı?" Algoritması
Her barkod için hesaplanan `n1 (Güncel)`, `n2 (Avantajlı)`, `n3 (Flaş)`, `n4 (Plus)` ve `n5_x (Plus Ek)` netleri kıyaslanır.
*   Sıfırdan veya eksi olan netler yoksayılır.
*   Karşılaştırma `n1` üzerinden yapılır. Diğer kampanyalar en az `n1`'e eşit veya ondan büyük olmalıdır.
*   Eğer bir ürün Flaş listesinde varsa, ancak `Flaş Fiyatı < Avantajlı Fiyatı` ise Flaş diskalifiye edilir (Avantajlı kazanır).
*   Maksimum `Net` değere sahip olan etiket (Örn: "Avantajlı Ürün") `Hangisi Daha Karlı?` alanına atanır.

---

## 4. Kullanıcı Arayüzü (Frontend) - DataTables ve Kural Motoru

`templates/index.html` bir AI tarafından yazılırken şunlara dikkat edilmelidir:

### A. Dinamik Fark (%) Sütunu
Bu tablo sütunu JSON'dan okunmaz, `render` fonksiyonu içinde yazılır:
```javascript
let fark = '';
let n1 = parseFloat(row['Güncel Ürün Kalan Net (TL)']);
let n_hedef = parseFloat(row['Seçilen Kampanya Kalan Net (TL)']);
if (sel === 'Hiçbiri') {
    let mevcut = row['Mevcut İndirim Oranı (%)'];
    if (mevcut) fark = parseFloat(mevcut);
} else {
    if (n1 > 0 && n_hedef) {
        fark = ((n_hedef - n1) / n1) * 100;
    }
}
return fark.toFixed(2); // Ekrana bu basılır
```

### B. Kural Motoru (Rule Engine) Implementation
Arayüzde oluşturulan `"Eğer [Avantajlı Net] [Flaş Net]'ten [%5] [daha az] ise -> [Flaş Yap]"` şeklindeki kural UI tarafında `eval` veya basit koşul bloklarıyla test edilir.
```javascript
tableData.forEach(row => {
    let solDeger = row[kural.solSutunIsmi]; // Örn: 'Avantajlı Ürün Kalan Net (TL)'
    let sagDeger = row[kural.sagSutunIsmi]; // Örn: 'Flaş Ürün Kalan Net (TL)'
    let farkYuzde = ((solDeger - sagDeger) / sagDeger) * 100;
    if (farkYuzde < -5) { // %5 daha az durumu
         row.userSelection = "Flaş";
    }
});
```
*Önemli: AI burayı yazarken string'leri float'a çevirmeyi ve `isNaN` kontrollerini eklemeyi asla unutmamalıdır.*

---

## 5. Reverse Iteration ile Excel Modifikasyonu (AI'ın Çözeceği Zorluk)

`app.py` içindeki `/api/apply_campaigns` rotası Excel dosyası oluştururken Pandas KULLANMAMALIDIR. (Format, dropdown listeleri, renkler vs. Pandas ile bozulur).
**Doğru Mimari:**
```python
from openpyxl import load_workbook

wb = load_workbook('Girdiler/Flaş.xlsx')
ws = wb.active
barcode_col = 2 # Örnek: Barkod B sütununda

rows_to_delete = []
for row_idx in range(2, ws.max_row + 1):
    cell_val = str(ws.cell(row=row_idx, column=barcode_col).value).strip()
    if cell_val not in selected_barcodes_for_flas:
        rows_to_delete.append(row_idx)

# AI İÇİN ALTIN KURAL: Satır silme (delete_rows) daima tersten yapılmalıdır!
for row_idx in reversed(rows_to_delete):
    ws.delete_rows(row_idx)

wb.save('Çıktılar/Flas_Urunler.xlsx')
```
*   **Avantajlı Çoğaltma Kuralı:** Eğer JSON payload'ında `avanStartDate` ve `avanEndDate` dolu gelirse, yukarıdaki silinmiş tertemiz dosya, gün bazlı bir while/for döngüsüne sokularak `Avantajlı_24_07_2026.xlsx`, `Avantajlı_25_07_2026.xlsx` şeklinde `shutil.copy` veya `wb.save` kullanılarak ayrı dosyalara çoğaltılmalıdır.
*   **Uygulanmayanlar:** Sadece bilgilendirme amaçlıdır. Barkodları alıp `pd.DataFrame.to_excel` kullanılarak basit bir dosya oluşturulur (Şablon bozulması endişesi yoktur).
