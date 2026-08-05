# Trendyol Kampanya Yönetimi - Mimari

## Bileşenler

- `app.py`: Flask rotaları, yükleme akışı, tablo verisi ve Trendyol şablon çıktıları.
- `input_files.py`: Excel tür/sütun doğrulaması, güvenli saklama ve yükleme manifesti.
- `komisyon_hesaplayici.py`: Barkod birleştirme, komisyon/net, indirim ve kampanya kıyasları.
- `templates/index.html`: Excel yükleme alanları, DataTables görünürlüğü ve interaktif kampanya seçimi.
- `Girdiler/Yuklenen/`: Son başarılı yükleme kümesinin sabit, uygulama kontrollü dosya adları.
- `Çıktılar/`: Hesap sonucu ve tarihli kampanya/rapor çıktıları.

Uygulama kökü `app.py` dosyasının bulunduğu dizindir; makineye veya kullanıcı adına bağlı sabit yol yoktur.

## Girdi akışı

`POST /api/calculate` bir `multipart/form-data` isteğidir. İlk hesaplamadan önce şu üç dosyanın yüklenmiş olması zorunludur:

1. İndirim uygulanabilecek ürünler
2. Komisyon tarifesi
3. Güncel ürünler

Avantajlı, Flaş, Plus, Plus Ek İndirim ve Karşılamalı kampanya dosyaları opsiyoneldir. Yüklenen her tür, aynı türde yeni bir dosya yüklenene kadar hesap ve çıktılarda kullanılmaya devam eder.

Dosya adları sınıflandırma amacıyla kullanılmaz. Özellikle Güncel Ürünler dosyası herhangi bir adla yüklenebilir; tür, seçilen yükleme alanı ve beklenen sütun imzasıyla doğrulanır. Plus Ek İndirim ve Karşılamalı şablonları aynı sütun yapısına sahip olduğundan bunları birbirinden kullanıcının seçtiği yükleme alanı ayırır.

Yüklenen `.xlsx` dosyaları boyut, ZIP yapısı ve zorunlu sütunlar açısından doğrulanır. Özgün dosya adı disk yolu oluşturmak için kullanılmaz. Doğrulanan dosyalar `Girdiler/Yuklenen/` altında sabit adlarla saklanır; özgün ad ve yükleme zamanı `Girdiler/yuklenen_girdiler.json` manifestinde tutulup sayfada gösterilir. Yeni istek yalnız gönderilen türleri değiştirir; klasör taraması yapılmaz.

## Hesaplama

Barkod havuzu Güncel Ürünler ile indirim uygulanabilecek ürünlerin birleşimidir. Kampanya dosyaları barkod ve önerilen fiyatlarıyla bu ürünlere aday kampanyalar ekler.

- Avantajlı fiyat: doluysa `YENİ TSF (FİYAT GÜNCELLE)`, değilse `1 YILDIZ ÜST FİYAT`.
- Flaş fiyat: `24 Saat Fiyat`.
- Plus fiyat: `Plus Fiyat Üst Limiti`.
- Plus Ek fiyat: `Maksimum Girebileceğin Fiyat` üzerinden seçilen `%5/%10/%20` indirim uygulandıktan sonraki fiyat.
- Karşılamalı fiyat: `Maksimum Girebileceğin Fiyat`.
- Net: kampanya fiyatı ve ilgili komisyon/satıcı katkısıyla hesaplanır.

Komisyon tarifesindeki ürün satırı veya fiyat dilimi boşsa Güncel Ürünler dosyasındaki `Komisyon Oranı` yedek değer olarak kullanılır.

Her satırda iki ayrı seçim bulunur:

- `İlk Kampanya Seçimi`: yüklenen kampanya adayları arasındaki en yüksek kalan net. Güncel Net daha yüksek olsa bile yüklenen kampanya başlangıç seçimi olarak görünür.
- `Hangisi Daha Karlı?`: Güncel Net dahil kıyas. Kampanya neti Güncel Net'i geçmiyorsa `Hiçbiri` olur.

`Önerilenleri Seç` düğmesi her ürün için kampanya adayları arasındaki en yüksek netli `İlk Kampanya Seçimi` değerini uygular. `Hangisi Daha Karlı?`, kampanyasız Güncel Net kıyasını ayrıca göstermeye devam eder ve kampanya seçimini silmez. `Uygulanabilir Kampanyalar` alanı, geçerli fiyat/net üreten türleri arayüzde rozet ve seçim seçeneği olarak gösterir.

İndirim alanları fiyat bazlıdır:

- Uygulanabilecek indirim = Güncel Fiyat - Düşülebilecek Dip Fiyat
- Uygulanan indirim = Güncel Fiyat - seçilen kampanya müşteri fiyatı
- Ekstra uygulanabilir indirim = uygulanabilecek indirim - uygulanan indirim (en az sıfır)

Bu dört tutar/yüzde alanı ile dip fiyat yalnız indirim uygulanabilecek ürünlerde oluşur.

## Tablo ve Excel çıktısı

Sayfa tablosu ile `Kampanya_Ozet_Raporu.xlsx`, `app.py` içindeki tek `REPORT_COLUMNS` sırasını kullanır. Excel'e yalnız DataTables sütun görünürlüğünde açık olan rapor sütunları, aynı sırayla yazılır. Varsayılan görünürlük ayrıntılı fiyat/net sütunlarını gizler; kullanıcı Sütun Seçimi menüsünden değiştirebilir.

`POST /api/apply` yalnız manifestte bulunan şablon türlerini işler. Seçili ürün yoksa o kampanya dosyası üretilmez. Şablon biçimleri `openpyxl` ile korunur; raporlar ayrıca tarihli çıktı klasörüne yazılır.
