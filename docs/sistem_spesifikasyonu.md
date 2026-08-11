# Kampanya hesaplayıcı sistem spesifikasyonu

Bu belge, 5 Ağustos 2026 tarihindeki kaynak kodunun dışarıdan gözlenebilir sözleşmesini tanımlar. Uygulama akışının açıklaması için [mimari ve işletim belgesine](mimari.md), Next.js hedefi için [aktarım promptuna](nextjs_aktarma_promptu.md) bakın.

## 1. Çalışma zamanı

| Özellik | Değer |
| --- | --- |
| Sunucu | Flask 3.1.3 |
| Veri işleme | pandas 3.0.5 |
| Excel okuma/yazma | openpyxl 3.1.5 |
| Şablon | Jinja2 3.1.6 |
| Dinleme | `127.0.0.1:5114` (`python app.py`) |
| İstek üst sınırı | 240 MB toplam HTTP gövdesi |
| Standart dosya üst sınırı | 30 MB / XLSX |
| XLSX açılmış ZIP üst sınırı | 150 MB |
| XLSX ZIP girdi üst sınırı | 5.000 |

Kurulum ve doğrulama:

```powershell
python -m pip install -r requirements.txt
python -m unittest discover -s tests -v
python app.py
```

## 2. Girdi sözleşmesi

### 2.1 Genel kurallar

- Üç zorunlu tür: `discount`, `commission`, `current`.
- Diğer standart türlerin tamamı opsiyoneldir.
- Dosya adı anlamlı değildir. İstemci dosyayı belirli bir yükleme alanında gönderir; sunucu o alanın türünü zorunlu sütun imzasıyla doğrular.
- Sadece `.xlsx` kabul edilir. XLSX’in geçerli ZIP olması, şifreli/bozuk olmaması, boyut sınırlarını aşmaması ve ilk sayfasının gerekli sütunları içermesi gerekir.
- Sütun adları mevcut kodda tam eşleşmeyle kontrol edilir; büyük/küçük harf ve Türkçe karakter farkları kabul edilmez.
- Barkodlar eşleştirmede baştaki/sondaki boşlukları temizlenmiş metne çevrilir. Aynı barkod bir dosyada birden çoksa ilk satır kullanılır.
- Yeni istekte gönderilmeyen geçerli eski girdiler korunur. Yeni gönderilen tür kendi kalıcı kopyasının üzerine atomik olarak geçer.

### 2.2 Standart yükleme alanları

| Form alanı | Arayüz etiketi | Zorunlu | Kalıcı iç ad | Zorunlu sütunlar |
| --- | --- | --- | --- | --- |
| `discount` | İndirim Uygulanabilecek Ürünler | Evet | `discount.xlsx` | `BARKOD`, `Eski Fiyat`, `YENİ Fiyat`, `Durum` |
| `commission` | Komisyon tarifesi | Evet | `commission.xlsx` | `BARKOD`, `1.Fiyat Alt Limit`, `2.Fiyat Üst Limiti`, `2.Fiyat Alt Limit`, `3.Fiyat Üst Limiti`, `3.Fiyat Alt Limit`, `4.Fiyat Üst Limiti`, `1.KOMİSYON`, `2.KOMİSYON`, `3.KOMİSYON`, `4.KOMİSYON`, `KOMİSYONA ESAS FİYAT`, `TARİFE GRUBU` |
| `current` | Güncel ürünler | Evet | `current.xlsx` | `Barkod`, `Komisyon Oranı`, `Piyasa Satış Fiyatı (KDV Dahil)`, `Trendyol'da Satılacak Fiyat (KDV Dahil)` |
| `advantage` | Avantajlı ürün | Hayır | `advantage.xlsx` | `BARKOD`, `1 YILDIZ ÜST FİYAT`, `YENİ TSF (FİYAT GÜNCELLE)` |
| `flash` | Flaş ürün | Hayır | `flash.xlsx` | `Barkod`, `24 Saat Fiyat`, `Kampanyalı Ürün`, `Güncellenecek Fiyat`, `24 Saat Flaş Başlangıç Tarihi` |
| `plus` | Plus ürün | Hayır | `plus.xlsx` | `Barkod`, `Plus Fiyat Üst Limiti`, `Plus Komisyon Teklifi`, `Plus Fiyat Seçimi`, `Tarife Seçimi` |
| `plus_extra` | Plus ek indirim | Hayır | `plus_extra.xlsx` | `Barkod`, `Maksimum Girebileceğin Fiyat`, `Kampanyalı Satış Fiyatı` |
| `muhasebe_avantaj` | Muhasebe – Avantajlı fiyat listesi | Hayır | `muhasebe_avantaj.xlsx` | `BARKOD`, `YENİ TSF (FİYAT GÜNCELLE)` |
| `muhasebe_flas` | Muhasebe – Flaş fiyat listesi | Hayır | `muhasebe_flas.xlsx` | `Barkod`, `Senin Belirlediğin Flaş Fiyatı` |
| `muhasebe_plus` | Muhasebe – Plus fiyat listesi | Hayır | `muhasebe_plus.xlsx` | `Barkod`, `Plus Fiyat Üst Limiti` |

“Muhasebe” dosyaları çıktı şablonu değildir. Aynı barkod için önerilen kampanya fiyatının öncelikli kaynağıdır. Şablon dosyasındaki fiyat ancak ilgili muhasebe fiyatı yoksa kullanılır.

### 2.3 Çoklu karşılamalı kampanya

Karşılamalı kampanya `INPUT_SPECS` içinde standart tek dosya değildir. `POST /api/calculate` isteğinde şu alanlarla gönderilir:

- `counter_configs_json`: aşağıdaki nesnelerden oluşan JSON dizi;
- `counter_file_0`, `counter_file_1`, ...: aynı sıra numarasındaki yeni Excel dosyaları.

```json
[
  {
    "id": "counter_1",
    "filename": "500-tl-uzeri-40-tl-indirim-40-trendyol.xlsx",
    "min_price": 500,
    "discount_amount": 40,
    "trendyol_percent": 40
  }
]
```

Dosya adındaki `N-tl-uzeri-N-tl-indirim-N-trendyol` deseni yalnızca üç sayısal alanın başlangıç değerlerini çıkarır. Kullanıcı bu değerleri arayüzde değiştirebilir; kararın kaynağı düzenlenmiş sayısal alanlardır, dosya adı değildir.

Hesaplama için etkin satır sözleşmesi `Barkod` ve fiyat kaynağı olarak `Maksimum Girebileceğin Fiyat` veya `Kampanyalı Satış Fiyatı` sütunudur. Trendyol çıktı şablonu üretiminde `Barkod`, `Maksimum Girebileceğin Fiyat` ve `Kampanyalı Satış Fiyatı` sütunlarının üçü de gerekir.

Yeni dosya varsa `Girdiler/Yuklenen/counter_files/counter_N.xlsx` yoluna kopyalanır. Yapılandırma, standart dosya manifestindeki `counter_configs` alanında saklanır.

> Mevcut sürümde bu yüklemeler standart XLSX güvenlik doğrulamasını kullanmaz ve kaydedilmiş kartlar sayfa açılışında yeniden oluşturulmaz. Bu, [bilinen bir sınırdır](mimari.md#bilinen-sınırlar-ve-teknik-borç).

## 3. Kalıcı durum sözleşmesi

`Girdiler/yuklenen_girdiler.json` çalışma zamanı manifestidir. Standart dosya başına en az şu anlamdaki bilgiler tutulur:

```json
{
  "files": {
    "current": {
      "stored_name": "current.xlsx",
      "original_name": "kullanıcının-dosyası.xlsx",
      "uploaded_at": "ISO-8601 zaman damgası"
    }
  },
  "counter_configs": []
}
```

Manifest yolu kullanıcı girdisinden üretilmez. Manifestteki her standart girdi yalnızca katalogdaki sabit iç adı taşıyorsa ve çözümlenen dosya `Girdiler/Yuklenen` dizininin doğrudan altındaysa geçerli sayılır.

`Çıktılar/Kampanya_Hesaplama_Sonuclari.xlsx` son hesaplamanın kalıcı önbelleğidir. Dosya yoksa veya değiştirilme zamanı manifestten eskiyse `/api/data` ve `/api/apply` yeni hesaplama ister.

## 4. Hesaplama kuralları

### 4.1 Ürün evreni ve kaynak önceliği

İşlenen barkodlar şu iki kümenin birleşimidir:

1. `discount` dosyasında `Durum` değeri büyük/küçük harften bağımsız olarak `ndirim` içeren satırlar;
2. `current` dosyasındaki bütün barkodlar.

İndirim uygulanabilir bir ürünün geçerli `Eski Fiyat` değeri sıfırdan büyükse hesaplamadaki güncel fiyat odur. Aksi halde `current` dosyasındaki `Trendyol'da Satılacak Fiyat (KDV Dahil)` kullanılır.

Kampanya fiyatı kaynak sırası:

| Kampanya | Birinci kaynak | Yedek kaynak |
| --- | --- | --- |
| Avantajlı | Muhasebe: `YENİ TSF (FİYAT GÜNCELLE)`, ardından varsa `1 YILDIZ ÜST FİYAT` / `TRENDYOL SATIŞ FİYATI` | Avantajlı şablonu: `YENİ TSF (FİYAT GÜNCELLE)`, ardından `1 YILDIZ ÜST FİYAT` |
| Flaş | Muhasebe: `Senin Belirlediğin Flaş Fiyatı`, ardından `24 Saat Fiyat`, `3 Saat Fiyat`, `Mevcut Fiyat` | Flaş şablonu: `24 Saat Fiyat`, ardından `3 Saat Fiyat` |
| Plus | Muhasebe: `Plus Fiyat Üst Limiti`, ardından varsa `Güncel TSF` | Plus şablonu: `Plus Fiyat Üst Limiti` |
| Plus Ek İndirim | Plus Ek şablonu: `Maksimum Girebileceğin Fiyat` | Yok |
| Karşılamalı | Karşılamalı şablonu: `Maksimum Girebileceğin Fiyat`, ardından `Kampanyalı Satış Fiyatı` | Yok |

### 4.2 Komisyon ve net

Standart kampanya neti:

```text
net = fiyat - (fiyat × komisyon_oranı / 100)
```

Komisyon oranı barkoda ait tarife satırında fiyat dilimine göre seçilir. Tarife kullanılamazsa `current` dosyasındaki `Komisyon Oranı` yedektir. Plus için `Plus Komisyon Teklifi` varsa diğer oranlardan önce kullanılır.

Plus Ek İndirimde `P` maksimum fiyat, `r` müşteri indirimi ve `k` komisyon oranıdır:

```text
müşteri_fiyatı = P × (1 - r / 100)        # r = 5, 10 veya 20
net = müşteri_fiyatı - (P × k / 100)
```

Komisyon tutarı indirimli müşteri fiyatından değil, şablondaki maksimum fiyattan hesaplanır. Üretilen Excel’in `Kampanyalı Satış Fiyatı` hücresine müşteri fiyatı bir kez yazılır; ikinci kez indirim uygulanmaz.

Plus Ek İndirimde her oranın son müşteri fiyatı ayrı değerlendirilir. İndirim/muhasebe girdisinden gerçek bir dip fiyat varsa bu değerin altına düşen oran aday olmaz.

Karşılamalı kampanyada `D` toplam kampanya indirimi, `T` Trendyol karşılama yüzdesidir:

```text
satıcı_katkısı = D × (1 - T / 100)
net = kampanya_fiyatı - (kampanya_fiyatı × komisyon_oranı / 100) - satıcı_katkısı
```

Ürünün hesap güncel fiyatı `min_price` eşiğinin altındaysa ilgili karşılamalı kampanya aday olmaz.

### 4.3 Dip fiyat ve indirim alanları

Dip fiyat adayları yalnızca şunlardır:

- indirim uygulanabilir ürünün `YENİ Fiyat` değeri;
- varsa muhasebe Avantajlı fiyatı;
- varsa muhasebe Flaş fiyatı;
- varsa muhasebe Plus fiyatı.

Pozitif adayların en düşüğü `Düşülebilecek Dip Fiyat (TL)` olur. Aday yoksa hesap güncel fiyatı kullanılır. Alan ve ona bağlı “uygulanabilecek/ekstra uygulanabilir” değerler yalnız `İndirim Uygulanabilir = Evet` satırlarında gösterilir.

Seçilmiş kampanya fiyatı `S`, güncel fiyat `G`, dip fiyat `D` olduğunda:

```text
uygulanabilecek_tutar = max(G - D, 0)
uygulanabilecek_yüzde = uygulanabilecek_tutar / G × 100
uygulanan_tutar       = max(G - S, 0)
uygulanan_yüzde       = uygulanan_tutar / G × 100
ekstra_tutar          = max(uygulanabilecek_tutar - uygulanan_tutar, 0)
ekstra_yüzde          = ekstra_tutar / G × 100
```

`Hiçbiri` seçiminde `S = G` kabul edilir; uygulanan indirim sıfırdır. Geçersiz, boş veya sıfır fiyatlar `0` yapılmaz; sonuç `null`/boş kalır ve arayüzde `-` görünür. Para ve yüzde alanları gösterim/rapor aşamasında iki ondalığa yuvarlanır.

### 4.4 Uygulanabilirlik ve öneri

Hesap motorunda iki ayrı kavram vardır:

- `eligible_campaigns`: ürünün ilgili girdi dosyasında fiyatla bulunmasına göre arayüz açılır listesinde gösterilen tam seçim anahtarları. Her zaman `Hiçbiri` içerir.
- `Uygulanabilir Kampanyalar`: dip fiyat, komisyon ve fallback kârlılık kontrollerini geçen adayların kampanya grupları. Sunucu dışa aktarım öncesi bu alanı doğrular; Plus Ek oranları burada tek `Plus Ek İndirim` grubuna iner.

Akıllı seçim:

1. İndirim veya muhasebe dosyasından pozitif dip fiyat varsa son kampanya fiyatı bu değerin `0,01 TL` toleransla altında olan aday elenir. Gerçek dip girilmemişse güncel fiyat yalnız gösterim için dip olur, kampanya fiyatına alt sınır uygulanmaz.
2. Kampanya/muhasebe dosyasında pozitif fiyat varsa aday “girilmiş fiyatlı” sayılır. Fiyat boşsa güncel fiyat kullanılır ve aday “fallback” olarak işaretlenir.
3. Fallback aday yalnız kendi neti hesaplanan güncel netten kesin yüksek olduğunda uygulanabilir olur. Güncel net hesaplanamıyorsa kârlılık kanıtlanamadığı için fallback aday kullanılmaz.
4. Girilmiş fiyatlı adaylar güncel netten düşük olsalar da kampanyaya alınabilir; uygulanabilir adaylar arasındaki en yüksek netli seçenek ilk seçim olur.
5. Gösterim etiketi Avantajlı/Flaş/Plus için sırasıyla `Avantajlı Ürün`, `Flaş Ürün`, `Plus Ürün` olur.

## 5. HTTP API

### `GET /`

HTML sayfasını döndürür. Şablona `report_columns`, `input_specs` ve `uploaded_inputs` verilir.

### `GET /api/data`

Güncel önbellek yoksa HTTP 200:

```json
{
  "needs_calculation": true,
  "message": "Lütfen önce 'Verileri Güncelle' butonuna basarak hesaplamaları başlatın."
}
```

Güncel önbellek varsa sonuç satırlarından oluşan JSON dizi döner. `eligible_campaigns` JSON dizisi, `counter_evaluations` JSON nesnesi olarak geri yüklenmiş olmalıdır.

### `POST /api/calculate`

İçerik türü `multipart/form-data`dır. Standart dosya alanları, `counter_configs_json` ve `counter_file_N` alanlarını kabul eder.

Başarılı yanıt HTTP 200 ve en az `success`, `results`, `output_path`, `counter_items`, `uploads` alanlarını içerir. Girdi doğrulama veya sayısal alan hatası HTTP 400; toplam gövde sınırı HTTP 413; hesaplama hatası HTTP 500 döner.

### `POST /api/apply`

İçerik türü `application/json`:

```json
{
  "target_type": "Hepsi",
  "selections": {
    "8690000000001": "Avantajlı",
    "8690000000002": "Plus Ek İndirim %10",
    "8690000000003": "Hiçbiri"
  },
  "visibleColumns": [
    "Barkod",
    "Uygulanan Kampanya",
    "Hangisi Karlı?"
  ]
}
```

Geçerli `target_type` değerleri:

- `Hepsi`
- `Avantajlı`
- `Flaş`
- `Plus`
- `Plus Ek İndirim`
- `Karşılamalı Kampanya`

`Karşılamalı Kampanya` değer olarak doğrulamadan geçse de mevcut sürüm bu tekil hedefi, standart katalogda bulunmayan eski `counter` girdisini aradığı için üretim dalına ulaşmadan reddeder. Çoklu karşılamalı dosyalar şu anda yalnız `Hepsi` hedefi üzerinden üretilebilir.

`selections` yalnız metin barkod → metin kampanya eşlemesi olabilir. Sunucu her seçimi hesaplanmış uygulanabilir gruplarla doğrular. Başarılı yanıt üretilen göreli dosya yollarını ve zaman damgalı klasörü döndürür:

```json
{
  "success": true,
  "generated_files": ["2026-08-05_21-30-00/Kampanya_Ozet_Raporu.xlsx"],
  "timestamp_folder": "2026-08-05_21-30-00",
  "message": "..."
}
```

### `GET /api/download/<folder>/<filename>`

`folder` yalnız `YYYY-MM-DD_HH-MM-SS` deseninde olabilir. Dosya ilgili `Çıktılar` alt klasöründen indirilir; başka klasörlere geçişe izin verilmez.

## 6. Tablo ve özet rapor sütun sözleşmesi

Sayfadaki rapor sütunları ve `Kampanya_Ozet_Raporu.xlsx` için kanonik sıra:

1. Barkod
2. Güncel Fiyat (TL)
3. Güncel Net
4. Güncel Komisyon
5. Avantajlı Fiyat (TL)
6. Avantajlı Net
7. Flaş Fiyat (TL)
8. Flaş Net
9. Plus Fiyat (TL)
10. Plus Net
11. Plus Ek İndirim Fiyat (TL)
12. Plus Ek İndirim Net
13. Uygulanan Kampanya
14. Hangisi Karlı?
15. Düşülebilecek Dip Fiyat (TL)
16. Uygulanan Kampanya Fiyat
17. Uygulanan Kampanya Net
18. Uygulanan Kampanya Komisyon
19. Uygulanabilecek İndirim (TL)
20. Uygulanabilecek İndirim (%)
21. Uygulanan İndirim (TL)
22. Uygulanan İndirim (%)
23. Ekstra Uygulanabilir İndirim (TL)
24. Ekstra Uygulanabilir İndirim (%)

Tablonun solundaki seçim kutusu bu sözleşmenin parçası değildir. `visibleColumns` yoksa veya dizi değilse 24 sütunun tamamı kullanılır. Dizi verilirse yalnız tanınan adlar alınır; sıralama her zaman yukarıdaki kanonik sıradır.

Plus Ek İndirim fiyat/net sütunları satırda seçili `%5`, `%10` veya `%20` oranını gösterir. Satırda Plus Ek seçili değilse tarayıcı tablosu `%5` önizlemesini gösterir; sunucunun ürettiği `Kampanya_Ozet_Raporu.xlsx` bu iki hücreyi boş bırakır. Bu, mevcut sürümdeki bilinen sayfa/Excel farkıdır.

## 7. Üretilen dosyalar

Her `/api/apply` çağrısı `Çıktılar/YYYY-MM-DD_HH-MM-SS` klasöründe uygun olan dosyaları üretir.

| Dosya | Koşul ve dönüşüm |
| --- | --- |
| `Avantajlı Ürün.xlsx` | Avantajlı seçilen satırlar bırakılır; `YENİ TSF (FİYAT GÜNCELLE)` hesaplanan Avantajlı fiyatla yazılır. |
| `Flas_Urun_<tarih>.xlsx` | Flaş seçilen satırlar başlangıç tarihine göre ayrılır; `Güncellenecek Fiyat = 24 Saat` yazılır. |
| `Plus_Urun.xlsx` | Plus seçilen satırlar bırakılır; `Plus Fiyat Seçimi` üst limite, `Tarife Seçimi` şablondan bulunan gün sayısına (`7` yedek) göre yazılır. |
| `Plus_Ek_Indirim_5.xlsx` | Yalnız `%5` seçilen satırlar; kampanyalı satış fiyatı maksimum fiyatın `%95`idir. |
| `Plus_Ek_Indirim_10.xlsx` | Yalnız `%10` seçilen satırlar; kampanyalı satış fiyatı maksimum fiyatın `%90`ıdır. |
| `Plus_Ek_Indirim_20.xlsx` | Yalnız `%20` seçilen satırlar; kampanyalı satış fiyatı maksimum fiyatın `%80`idir. |
| `Karsilamali_<etiket>.xlsx` | Mevcut sürümde yalnız `Hepsi` hedefi içinde ulaşılabilir. İlgili dinamik etiket seçilen satırlar bırakılır; `Kampanyalı Satış Fiyatı` maksimum fiyata yazılır. |
| `Uygulanmayan_Urunler_Raporu.xlsx` | Son seçimi `Hiçbiri` olan hesap satırları. |
| `Kampanya_Genel_Raporu.xlsx` | Tüm iç hesap alanları ve son seçim. |
| `Kampanya_Ozet_Raporu.xlsx` | Sayfadaki görünür rapor sütunlarının kanonik sıradaki karşılığı; A2 sabit, otomatik filtreli. |
| `Indirim_Uygulanmayan_Fiyat_Kiyas_Raporu.xlsx` | Yardımcı fiyat farkı raporu; üretilemezse ana çıktı akışı devam eder. |

Boş kampanya seçiminde ilgili şablon dosyası üretilmez. Rapor dosyaları veri varsa üretilir. Trendyol şablonları ve raporlar kaydedildikten sonra `fix_xlsx_for_trendyol` işleminden geçirilir.

## 8. Doğrulama kabul ölçütleri

Aşağıdaki davranışlar otomatik veya fixture tabanlı doğrulamalarla korunmalıdır. Mevcut `unittest` paketi çekirdek hesap/arayüz sözleşmelerinin önemli bölümünü kapsar; aşağıda ayrıca belirtilen boşlukları henüz kapsamaz.

1. Güncel ürün girdisi özgün dosya adına bağlı değildir; doğru sütunlu farklı ad kabul edilir.
2. Sadece `discount`, `commission` ve `current` zorunludur.
3. Yüklenen özgün ad disk yolu olmaz; yol geçişi denemesi kalıcı kökün dışına çıkamaz.
4. Yalnız yeni gönderilen tür değişir; diğer kalıcı girdiler ve zaman bilgileri korunur.
5. En yüksek kampanya neti güncel netten yüksek değilse öneri `Hiçbiri`dir.
6. Plus Ek `%5/%10/%20` ayrı seçimlerdir ve dışa aktarılan fiyat bir kez indirilir.
7. Karşılamalı değerlendirme fiyat, komisyon oranı, net ve satıcı katkısını taşır.
8. İndirim örneği `G=100`, `S=90`, `D=80` için sırasıyla uygulanabilecek `20`, uygulanan `10`, ekstra `10` verir.
9. Özet rapor yalnız görünür sütunları ve her zaman kanonik sırayı kullanır.
10. Sunucu, ürün için akıllı olarak uygulanabilir olmayan seçimi reddeder.
11. Excel önbelleğindeki liste/sözlük alanları JSON yanıtında gerçek koleksiyonlara döner.
12. Boş sayısal değer arayüzde `0.00` değil `-` görünür.
13. Üretilen şablonlar shared strings ve gerekli OOXML ilişkilerini içerir; Trendyol’un kabul ettiği örnek dosyalarla açılıp yüklenebilir.

Mevcut otomatik kapsamdaki önemli boşluklar:

- yalnız `Karşılamalı Kampanya` hedefiyle birden çok counter dosyası üretimi;
- Plus Ek seçili değilken sayfa ile `Kampanya_Ozet_Raporu.xlsx` hücrelerinin eşitliği;
- gerçek kampanya şablonlarında shared strings, formül önbelleği ve veri doğrulama yapısının uçtan uca kontrolü.

Kaynak test komutu:

```powershell
python -m unittest discover -s tests -v
```
