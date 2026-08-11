# Next.js aktarım ana promptu

Bu prompt, Flask uygulamasındaki iş davranışını bir Next.js projesine taşımak için hazırlanmıştır. Hedef ajan mümkünse bu dosyayla birlikte [sistem spesifikasyonunu](sistem_spesifikasyonu.md), [mimari belgeyi](mimari.md), kaynak kodu ve örnek Excel’leri de görmelidir.

Teknik dayanaklar: güncel Next.js App Router [Route Handler sözleşmesi](https://nextjs.org/docs/app/api-reference/file-conventions/route), [Backend for Frontend rehberi](https://nextjs.org/docs/app/guides/backend-for-frontend) ve [Node.js runtime yapılandırması](https://nextjs.org/docs/app/api-reference/file-conventions/route-segment-config/runtime).

`PROMPT BAŞLANGICI` ile `PROMPT SONU` arasını hedef Next.js projesindeki kodlama ajanına verin. Köşeli parantezli yolları gerçek yollarla değiştirin.

---

## PROMPT BAŞLANGICI

Sen kıdemli bir TypeScript/Next.js geliştiricisisin. Görevin, `[KAYNAK_FLASK_REPO]` içindeki Trendyol kampanya hesaplayıcısının **iş davranışını** `[HEDEF_NEXTJS_REPO]` içindeki mevcut Next.js uygulamasına taşımak. Bu bir satır satır Python çevirisi değil; Excel giriş/çıkış sözleşmeleri, hesap sonuçları, öneri kararı, kalıcılık ve kullanıcı akışı davranış bakımından eşdeğer olmalı.

### Çalışma kuralları

1. Önce hedef repodaki `AGENTS.md`/talimatları, `package.json`, kilit dosyası, Next.js sürümü, App Router yapısı, TypeScript ayarları, test altyapısı, mevcut UI bileşenleri, depolama yaklaşımı ve `git status` çıktısını oku.
2. Kaynak Flask repo erişilebiliyorsa şu dosyaları uçtan uca incele: `app.py`, `input_files.py`, `komisyon_hesaplayici.py`, `xlsx_postprocess.py`, `fiyat_farki_analiz_script.py`, `templates/index.html`, `tests/`, `docs/mimari.md`, `docs/sistem_spesifikasyonu.md`.
3. Kullanıcının hedef repodaki mevcut değişikliklerini koru. İlgisiz dosyaları biçimlendirme, geri alma veya yeniden yazma.
4. Uygulamadan önce kısa ve dosya bazlı bir plan çıkar. Ardından makul varsayımlarla uygula; yalnız depolama/dağıtım seçimi sonucu temelden değiştiriyorsa soru sor.
5. Hedef projenin mevcut kütüphanesi aynı işi güvenle yapıyorsa onu kullan. Yeni paket eklemeden önce standart Node/Next özelliğini ve kurulu bağımlılıkları kontrol et.
6. Excel okuma/yazma yalnız sunucuda çalışsın. İlgili Route Handler’larda `export const runtime = 'nodejs'` kullan; değişebilir veri uçlarını statik önbelleğe alma. Mevcut Next.js sürümünün resmi Route Handler API’sine uy.
7. İş kurallarını Route Handler veya React bileşenlerinin içine gömme. Saf TypeScript fonksiyonlarıyla hesaplama, rapor satırı ve seçim doğrulama katmanı oluştur; HTTP katmanı yalnız doğrulasın, depolamayı çağırsın ve sonuç dönsün.
8. Barkodları metin olarak koru. Boş sayısal hücreleri `0` yapma. Para hesaplarında `NaN`, `Infinity` veya bilimsel gösterimli barkod sızıntısına izin verme.
9. Kimlik bilgisi, sabit kişisel yol veya kaynak Flask repodaki eski mutlak yolları kopyalama. Commit/push/deploy ancak ayrıca istenirse yapılır.

### Hedef mimari

Mevcut hedef yapıya uyarlayarak en küçük anlaşılır ayrımı kur. Aşağıdaki isimler öneridir, zorunlu klasör şablonu değildir:

```text
app/
├── page.tsx
└── api/
    ├── data/route.ts
    ├── calculate/route.ts
    ├── apply/route.ts
    └── download/[folder]/[filename]/route.ts
src/
├── campaign/
│   ├── calculator.ts
│   ├── commission.ts
│   ├── report.ts
│   └── contracts.ts
└── server/
    ├── excel/
    ├── input-validation.ts
    └── storage.ts
```

- `calculator.ts` ve `commission.ts` dosya sistemi, HTTP ve React bilmesin.
- `report.ts` sayfa ile Excel özetinin kullandığı tek rapor satırı dönüşümünü içersin.
- `contracts.ts` girdi türleri, kampanya seçim anahtarları ve kanonik rapor sütunları için tek kaynak olsun.
- Tek bir dağıtım hedefi varsa tek depolama uygulaması kullan; sırf gelecek ihtimali için factory/interface zinciri kurma.
- Hedef self-hosted ve tek kalıcı diskli ise çalışma zamanı dosyalarını proje kaynağından ayrı yazılabilir bir veri dizininde tut. Hedef serverless veya çok örnekliyse süreç dosya sistemine güvenme; hedef repoda zaten kullanılan nesne depolama/veritabanını kullan. Dağıtım biçimi bilinmiyor ve repo cevap vermiyorsa uygulamaya başlamadan bu tek noktayı sor.

### Girdi kataloğu

Üç zorunlu tür:

- `discount`: `BARKOD`, `Eski Fiyat`, `YENİ Fiyat`, `Durum`
- `commission`: `BARKOD`, `1.Fiyat Alt Limit`, `2.Fiyat Üst Limiti`, `2.Fiyat Alt Limit`, `3.Fiyat Üst Limiti`, `3.Fiyat Alt Limit`, `4.Fiyat Üst Limiti`, `1.KOMİSYON`, `2.KOMİSYON`, `3.KOMİSYON`, `4.KOMİSYON`, `KOMİSYONA ESAS FİYAT`, `TARİFE GRUBU`
- `current`: `Barkod`, `Komisyon Oranı`, `Piyasa Satış Fiyatı (KDV Dahil)`, `Trendyol'da Satılacak Fiyat (KDV Dahil)`

Opsiyonel türler:

- `advantage`: `BARKOD`, `1 YILDIZ ÜST FİYAT`, `YENİ TSF (FİYAT GÜNCELLE)`
- `flash`: `Barkod`, `24 Saat Fiyat`, `Kampanyalı Ürün`, `Güncellenecek Fiyat`, `24 Saat Flaş Başlangıç Tarihi`
- `plus`: `Barkod`, `Plus Fiyat Üst Limiti`, `Plus Komisyon Teklifi`, `Plus Fiyat Seçimi`, `Tarife Seçimi`
- `plus_extra`: `Barkod`, `Maksimum Girebileceğin Fiyat`, `Kampanyalı Satış Fiyatı`
- `muhasebe_avantaj`: `BARKOD`, `YENİ TSF (FİYAT GÜNCELLE)`
- `muhasebe_flas`: `Barkod`, `Senin Belirlediğin Flaş Fiyatı`
- `muhasebe_plus`: `Barkod`, `Plus Fiyat Üst Limiti`

Dosya tanıma kuralları:

- Özgün dosya adına göre tür tahmin etme. Girdi türü yükleme alanından gelir ve sütun imzasıyla doğrulanır.
- Özellikle “Güncel ürünler” için hiçbir dosya adı kuralı olmasın.
- Eski `Hesaplanmis_Komisyon`, `Uygulanmis` veya `Trendyol Fiyat` sütun yapılarına uyumluluk kodu ekleme.
- Yalnız `.xlsx` kabul et. Uzantı yanında gerçek ZIP/XLSX yapısını doğrula.
- Standart dosya başına 30 MB, açılmış ZIP toplamında 150 MB ve 5.000 ZIP girdisi sınırını uygula. Self-hosted HTTP akışında mevcut 240 MB toplam istek sözleşmesini açıkça yapılandır. Hosting platformunun istek sınırı bunu karşılamıyorsa sınırı sessizce düşürme; dosyayı doğrudan güvenli nesne depolamaya yükleyip sunucu doğrulamasına alan bir akış kur veya dağıtım kararını kullanıcıya taşı.
- ZIP path traversal, aşırı sıkıştırma, şifreli/bozuk dosya, eksik workbook ve eksik sütunları kullanıcıya anlaşılır 4xx hatasıyla reddet.
- Kullanıcı dosya adını disk/nesne anahtarı olarak doğrudan kullanma. İç anahtarları sunucu üretir; indirme ve depolama yollarında kök dışına çıkışı engelle.

### Kalıcılık

- Yüklenen her tür, yenisi gelene kadar kalıcı olsun.
- Yalnız yeni gönderilen tür değişsin; diğerleri korunmalı.
- Her tür için özgün ad ve yükleme zamanı saklansın, sayfa yenilendiğinde görünsün.
- Üç zorunlu dosya daha önce kalıcı olarak yüklüyse kullanıcı yeniden seçmeden hesaplama yapabilsin.
- Son hesap sonucu, kullanılan girdi sürümlerine/manifest revizyonuna bağlı olsun. Herhangi bir girdi değiştiğinde eski sonuç `needs_calculation` sayılmalı. Yalnız dosya `mtime`ına güvenmek yerine depolama katmanına uygun sürüm/hash veya manifest revizyonu tercih et.
- Yükleme kaydı ile dosya yazımı yarım kalmayacak şekilde atomik/işlemsel davran. Hata halinde önceki geçerli set korunmalı.
- Çoklu karşılamalı kampanya dosyalarını, etiketlerini ve düzenlenmiş parametrelerini de kalıcılaştır; sayfa yenilendiğinde kartları geri yükle. Dosya seçilmeden yapılan sonraki hesaplama mevcut counter yapılandırmasını yanlışlıkla silmesin.

### Çoklu karşılamalı kampanya

İstemci `counter_configs_json` ile birden çok yapılandırma ve gerekirse `counter_file_N` alanları gönderebilsin. Her yapılandırma:

```ts
type CounterConfig = {
  id: string
  filename: string
  min_price: number
  discount_amount: number
  trendyol_percent: number
  label: string
  storedKey?: string
}
```

- Dosya adındaki `N-tl-uzeri-N-tl-indirim-N-trendyol` deseni yalnız form varsayılanlarını doldursun; kullanıcı değerleri düzenleyebilsin.
- `trendyol_percent` 0–100, diğer sayılar 0 veya daha büyük ve sonlu olmalı.
- Dosyada `Barkod`, `Maksimum Girebileceğin Fiyat`, `Kampanyalı Satış Fiyatı` başlıklarını doğrula.
- Standart dosyalardaki bütün XLSX güvenlik kontrollerini counter dosyalarına da uygula.
- Her counter etiketi ayrı bir seçim anahtarıdır ve sonuç satırında `counter_evaluations[label] = { price, rate, net, seller_disc }` olarak bulunur.
- `Hepsi` ve yalnız `Karşılamalı Kampanya` dışa aktarımlarının ikisi de kalıcı counter yapılandırmalarını esas alsın; eski tekil `counter` alanını aramasın.

### Hesaplama davranışı

Kaynak Flask testleriyle birebir karşılaştırılabilir saf fonksiyonlar yaz.

1. Ürün evreni, `discount.Durum` değeri `ndirim` içeren barkodlarla bütün `current` barkodlarının birleşimidir. Aynı dosyadaki mükerrer barkodda ilk satır kullanılır.
2. İndirim uygulanabilir satırda pozitif `Eski Fiyat` varsa hesap güncel fiyatıdır; yoksa güncel ürün satış fiyatını kullan.
3. Komisyonu tarife fiyat diliminden bul; bulunamazsa güncel ürün `Komisyon Oranı` yedeğini kullan. Plus’ta geçerli `Plus Komisyon Teklifi` önceliklidir.
4. Normal net: `fiyat - fiyat × komisyon / 100`.
5. Muhasebe Avantajlı/Flaş/Plus fiyatları aynı kampanyanın şablon fiyatından önceliklidir.
6. Dip fiyat, pozitif `discount.YENİ Fiyat` ile mevcut muhasebe Avantajlı/Flaş/Plus fiyatlarının en düşüğüdür; aday yoksa güncel fiyattır. Dip ve bağlı indirim alanları yalnız indirim uygulanabilir ürünlerde görünür.
7. İndirim/muhasebe kaynaklarından pozitif gerçek dip fiyat varsa hiçbir kampanya son fiyatı bu değerin `0,01 TL` toleransla altına inemez. Gerçek dip yoksa kampanya dosyasındaki pozitif fiyatı alt sınır uygulamadan aday yap. Kampanya fiyatı da boşsa güncel fiyatı kullan ve bu fallback adayı yalnız neti hesaplanmış güncel netten kesin yüksekse uygulanabilir say; güncel net yoksa fallback adayı reddet.
8. Plus Ek `%5`, `%10`, `%20` ayrı adaydır. `P` maksimum fiyat ise müşteri fiyatı `P × (1-r/100)`, net `müşteri fiyatı - P × komisyon/100` olur. Gerçek dip varsa son müşteri fiyatını dip ile karşılaştır. `P` boşsa güncel fiyatı taban al ve fallback kârlılık kuralını uygula. Dışa aktarımda müşteri fiyatına ikinci kez indirim uygulama.
9. Counter’da güncel fiyat `min_price` eşiğini geçmeli. Satıcı katkısı `discount_amount × (1-trendyol_percent/100)`, net `kampanya_fiyatı - komisyon_tutarı - satıcı_katkısı`dır.
10. Girilmiş fiyatlı adaylar ile güncel neti geçen fallback adayları arasından en yüksek netliyi öner. Girilmiş fiyatlı adayın güncel netten düşük olması tek başına `Hiçbiri` nedeni değildir.
11. `eligible_campaigns` ve sunucunun gerçekten uygulanabilir kabul ettiği liste aynı seçim anahtarlarını kullansın. Kaynak sürümdeki UI’da görünüp sunucuda reddedilme tutarsızlığını taşıma. Plus Ek oranlarını doğrularken grup bilgisini kaybetme.
12. Rapor hesapları: `uygulanabilecek = güncel-dip`, `uygulanan = güncel-seçili`, `ekstra = max(uygulanabilecek-uygulanan,0)`; yüzdeler güncel fiyata bölünür. Geçersiz fiyat `null` kalır.
13. İki ondalık davranışını kaynak örneklerle golden test et. Rapor tarafında mevcut yarım-yukarı eşdeğeri `Math.floor(value * 100 + 0.5 + 1e-9) / 100`dür; Python ara net yuvarlamalarıyla ayrışan sınır değerlerde kaynak fixture çıktısını sözleşme kabul et.

### Kanonik rapor sütunları

Aşağıdaki tek sabiti hem tabloyu hem özet Excel’i üretmek için kullan; sırayı değiştirme:

```ts
export const REPORT_COLUMNS = [
  'Barkod',
  'Güncel Fiyat (TL)',
  'Güncel Net',
  'Güncel Komisyon',
  'Avantajlı Fiyat (TL)',
  'Avantajlı Net',
  'Flaş Fiyat (TL)',
  'Flaş Net',
  'Plus Fiyat (TL)',
  'Plus Net',
  'Plus Ek İndirim Fiyat (TL)',
  'Plus Ek İndirim Net',
  'Uygulanan Kampanya',
  'Hangisi Karlı?',
  'Düşülebilecek Dip Fiyat (TL)',
  'Uygulanan Kampanya Fiyat',
  'Uygulanan Kampanya Net',
  'Uygulanan Kampanya Komisyon',
  'Uygulanabilecek İndirim (TL)',
  'Uygulanabilecek İndirim (%)',
  'Uygulanan İndirim (TL)',
  'Uygulanan İndirim (%)',
  'Ekstra Uygulanabilir İndirim (TL)',
  'Ekstra Uygulanabilir İndirim (%)',
] as const
```

- Tablo seçim kutusu bu listeye dahil değildir.
- `Kampanya_Ozet_Raporu.xlsx`, yalnız o anda görünür olan sütunları alsın ama her zaman bu kanonik sıraya dizsin.
- Plus Ek fiyat/net alanı satırda seçili `%5/%10/%20` değerini göstersin; seçim Plus Ek değilse hem sayfa hem özet Excel aynı `%5` önizlemesini kullansın. Kaynak Flask sürümündeki sayfa dolu/Excel boş farkını taşıma.
- `null` sayılar tabloda `-`, gerçek sıfırlar `0.00` gösterilsin.

### Arayüz

Hedef projenin mevcut tasarım sistemini kullan. Kaynaktaki CDN/jQuery/DataTables yığınını kopyalama; React bileşenleriyle eşdeğer davranışı kur.

Kullanıcıya görünen metinler Türkçe olsun. Dosya alanları, filtreler, açılır listeler ve tablo kontrolleri klavye ile kullanılabilsin; ilişkili `label` ve erişilebilir adlar eksik olmasın.

Sayfa şunları içermeli:

- zorunlu/opsiyonel Excel yükleme alanları;
- her alan için kalıcı “yüklendi” durumu, özgün ad ve yükleme tarihi;
- birden çok karşılamalı dosya kartı ve düzenlenebilir üç parametre;
- hesapla düğmesi ve anlaşılır yükleniyor/hata/başarı durumları;
- global arama, sayfalama/sayfa boyutu;
- yalnız indirim uygulanabilir ürün filtresi;
- `Uygulanan Kampanya` ve `Hangisi Karlı?` filtreleri;
- filtreleri temizleme;
- sütun görünürlük menüsü, tümünü seç/temizle;
- satır seçim kutuları, tümünü seç, seçili satırlara toplu kampanya;
- her ürünün gerçekten uygulanabilir seçimlerinden oluşan ürün bazlı açılır liste ve rozetler;
- `Önerilenleri Seç` ve tüm seçimleri temizle;
- `Hepsi`, Avantajlı, Flaş, Plus, Plus Ek İndirim ve Karşılamalı çıktı işlemleri;
- kaynak özellik korunacaksa Avantajlı/Flaş özel kural motoru; kuralların geçersiz kampanya atamasına izin vermemesi gerekir.

Varsayılan gizli sütunlar: Güncel Fiyat/Net/Komisyon; Avantajlı Fiyat/Net; Flaş Fiyat/Net; Plus Fiyat/Net; Plus Ek İndirim Fiyat/Net; Uygulanan Kampanya Fiyat/Net/Komisyon. Diğerleri görünür başlasın.

Kullanıcı seçimi tarayıcıdaki sayfalama/filtreleme sırasında kaybolmasın. “Önerilenleri Seç” her satırı sunucudan gelen `İlk Kampanya Seçimi` değerine getirsin ve gerçekten kampanya seçilen ürün sayısını bildirsin.

### API sözleşmesi

App Router Route Handler’larıyla şu davranışları sağla:

- `GET /api/data`: güncel sonuç dizisi veya `{ needs_calculation: true, message }`.
- `POST /api/calculate`: `request.formData()` ile standart dosyalar, `counter_configs_json` ve `counter_file_N`; başarıda sonuçlar ve güncel upload durumları.
- `POST /api/apply`: `{ target_type, selections, visibleColumns }`; seçimleri sunucuda tekrar doğrula, dosyaları üret ve göreli yolları döndür.
- `GET /api/download/[folder]/[filename]`: yalnız sunucunun ürettiği kayıtlı çıktı anahtarlarını indir. Yerel disk kullanılıyorsa klasörün `YYYY-MM-DD_HH-MM-SS` desenini ve çözümlenen yolun çıktı kökü altında kaldığını doğrula.

Yanıtlar tutarlı `{ success, data?, message?, errors? }` biçiminde olabilir; ancak mevcut istemciyi aşamalı taşıyorsan geçiş boyunca Flask yanıt alanlarını koru. Beklenen girdi hataları 4xx, beklenmeyen sunucu hataları ayrıntı sızdırmayan 500 olmalı; sunucu logunda neden ve bağlam bulunmalı.

### Excel çıktı davranışı

Kaynak şablonların biçimlerini, formüllerini ve veri doğrulamalarını mümkün olduğunca koru. Hedef repoda kurulu ve bunu kanıtlayan bir kütüphane yoksa önce küçük bir fixture deneyi yap; başarılı aracı seç. Gerekirse workbook düzenleme için ExcelJS benzeri bir kütüphane, OOXML ZIP son işlemi için mevcut ZIP kütüphanesi kullan; körlemesine bağımlılık ekleme.

Her uygulama çalışması ayrı bir zaman damgalı çıktı grubu üretmeli:

- `Avantajlı Ürün.xlsx`: yalnız Avantajlı seçilenler, hesaplanan yeni TSF.
- `Flas_Urun_<tarih>.xlsx`: yalnız Flaş seçilenler, başlangıç tarihine göre grup, `Güncellenecek Fiyat = 24 Saat`.
- `Plus_Urun.xlsx`: yalnız Plus seçilenler, fiyat ve tarife seçimi doldurulmuş.
- `Plus_Ek_Indirim_5.xlsx`, `_10.xlsx`, `_20.xlsx`: yalnız ilgili oran seçilenler, doğru son müşteri fiyatı.
- `Karsilamali_<etiket>.xlsx`: yalnız ilgili dinamik counter etiketi seçilenler.
- `Uygulanmayan_Urunler_Raporu.xlsx`.
- `Kampanya_Genel_Raporu.xlsx`.
- `Kampanya_Ozet_Raporu.xlsx`: yalnız görünür kanonik sütunlar, A2 dondurma, otomatik filtre ve okunabilir genişlikler.
- `Indirim_Uygulanmayan_Fiyat_Kiyas_Raporu.xlsx`: **kalıcı olarak yüklenen güncel `discount` girdisini** kullansın; depoda adı “indirim” geçen ilk örnek dosyayı taramasın.

Trendyol uyumluluğu için üretilen XLSX’lerde en az şunları fixture üzerinden doğrula:

- `xl/sharedStrings.xml` ve doğru content type/relation;
- metin hücrelerinin shared-string referansları;
- formül hücrelerinin gerekli önbellek değerleri;
- boş sayısal hücrelerin bozuk `0`/NaN üretmemesi;
- silinen satırlardan sonra formül ve data-validation aralıkları;
- Excel/openpyxl ile yeniden açılabilme ve Trendyol’un kabul ettiği bilinen örnekle yapı eşdeğerliği.

### Güvenlik ve veri bütünlüğü

- Bütün multipart, JSON ve route parametrelerini şema ile doğrula.
- Form alanından gelen kampanya seçimine güvenme; barkodun uygulanabilir seçim setiyle sunucuda kontrol et.
- Formül enjeksiyonu oluşturabilecek rapor metinlerini Excel’e yazarken güvenli metin olarak ele al; yalnız kaynak şablondaki/uygulamanın ürettiği bilinen formülleri formül kabul et.
- İstemciye mutlak yol, stack trace veya dosya sistemi ayrıntısı döndürme.
- Hedef projede kimlik doğrulama varsa API ve indirmelere uygula. Yoksa yeni bir auth sistemi uydurma; uygulamayı yalnız güvenilir/self-hosted kullanım sınırıyla belgeleyip herkese açık dağıtımdan önce auth/CSRF/rate-limit gereksinimini açıkça işaretle.
- Eşzamanlı hesapla/uygula isteklerinin aynı manifest veya çıktı grubunu yarım yazmasını engelle. Hedef tek kullanıcı/tek süreçse en küçük güvenli kilit/atomik yazma çözümünü kullan.

### Kaynaktan taşınmayacak eski/bozuk davranışlar

- `temp_script_3.js`, `temp_script_clean.js`, `prompt.txt`, `girdieski/` ve kişisel mutlak yol içeren `karsilamali_config.json` çalışma zamanı bağımlılığı değildir; hedefe taşıma.
- Ana girdi olarak repo klasörlerini veya dosya adı desenlerini tarama.
- Çoklu counter yapılandırmasını sayfa yenilenince kaybetme.
- Counter dosyalarını doğrulamasız kabul etme.
- Yalnız karşılamalı çıktı alırken var olmayan eski tekil `counter` girdisini arama.
- UI ile sunucuda iki farklı uygulanabilirlik listesi kullanma.
- Plus Ek son müşteri fiyatını dip fiyat kontrolü dışında bırakma.
- Plus Ek seçili olmayan satırda sayfa ile özet Excel’e farklı değer yazma.
- Fiyat kıyas raporunda eski örnek indirim dosyasını kullanma.
- Flask `debug=True`, yerel mutlak yollar veya global tek kullanıcı durumunu üretim varsayımı olarak kopyalama.

### TDD ve kabul testleri

Önce parity testlerini yaz ve en az şu kırmızı testleri gör; sonra en küçük uygulamayla geçir. Hedef repo talimatı daha yüksek değilse değiştirilen iş mantığında en az `%80` coverage sağla.

Unit testleri:

1. Farklı adlı ama doğru sütunlu `current` dosyası kabul edilir.
2. Yalnız üç zorunlu tür zorunludur.
3. Komisyon fiyat sınırları ve güncel oran yedeği.
4. `current net=100`, aday neti `100` ise `Hiçbiri`; `100.01` ise aday.
5. Muhasebe fiyatı şablon fiyatının üzerine yazar.
6. `G=100`, `S=90`, `D=80` → uygulanabilecek `20`, uygulanan `10`, ekstra `10`.
7. İndirim uygulanamayan üründe dip ve uygulanabilecek/ekstra alanları `null`.
8. Plus Ek `%5/%10/%20` fiyat ve netleri; dip altındaki son fiyat aday değil; dışa aktarımda çift indirim yok.
9. Counter eşik, Trendyol katkısı ve net hesabı.
10. Görünür sütunlar istemci sırası farklı olsa bile kanonik sıraya döner.
11. Boş sayısal değer `null`, gerçek sıfır `0` kalır.
12. Dinamik counter etiketli seçim doğrulanır; başka ürünün/uygunsuz kampanyanın seçimi reddedilir.
13. Plus Ek seçili olmayan satırda sayfa rapor satırı ile özet Excel rapor satırı aynı `%5` önizlemesini verir.

Entegrasyon testleri:

1. Yeni yükleme yalnız ilgili türü değiştirir, diğer dosya ve tarihler korunur.
2. Yol geçişi içeren özgün ad depolama kökü dışına çıkamaz.
3. Bozuk/şifreli/aşırı büyük XLSX ve eksik başlık 4xx döner; counter da aynı hattan geçer.
4. Manifest/girdi revizyonu değişince `/api/data` hesaplama ister.
5. Excel’e yazılmış sonuçtaki liste/sözlük alanları API’de gerçek JSON koleksiyonlarıdır; tercihen yeni uygulamada bunları Excel metnine bağımlı bırakma.
6. Sayfa yenilenince standart ve counter yükleme tarihleri/parametreleri geri gelir.
7. `/api/apply` görünür sütunları ve son kullanıcı seçimlerini kullanır.
8. Yalnız `Karşılamalı Kampanya` hedefi yüklü çoklu counter ile çalışır.

E2E testleri:

1. Üç zorunlu ve opsiyonel örnekleri yükle → hesapla → önerilenleri seç → en az bir ürünün güncel netten daha iyi kampanyaya geçtiğini gör.
2. Ürün bazında kampanyayı değiştir; filtrele/sayfala; seçim korunmalı.
3. Sütunları gizle; `Kampanya_Ozet_Raporu.xlsx` yalnız görünür sütunları doğru sırada içermeli.
4. Sayfayı yenile; yükleme tarihleri ve counter kartları kalmalı.
5. Hepsi ve tek kampanya çıktılarında doğru barkodlar/son fiyatlar bulunmalı.

Golden Excel testlerinde kaynak Flask uygulamasını ve Next.js uygulamasını aynı küçük fixture setiyle çalıştır; kanonik rapor hücrelerini ve kampanya şablonlarındaki değiştirilmesi gereken hücreleri karşılaştır. ZIP metadata/tarih gibi anlamsız byte farklarını değil semantik workbook içeriğini, shared strings yapısını ve formül/data-validation sözleşmesini karşılaştır.

### Teslim ve doğrulama

Teslimden önce:

1. Unit, integration ve kritik E2E testlerini çalıştır.
2. Coverage raporunu çalıştır.
3. Hedef repodaki lint, typecheck ve production build komutlarını çalıştır.
4. En az bir gerçek örnek girdi setiyle hesaplama ve Excel üretim smoke testi yap.
5. Üretilen XLSX’leri seçilen Node kütüphanesiyle yeniden aç; mümkünse LibreOffice/Excel veya kaynak openpyxl doğrulamasıyla da kontrol et.
6. `git diff`i incele; ilgisiz dosya, secret, kişisel yol ve üretilen Excel olmadığını doğrula.
7. Mevcut proje dokümanına yalnız gerekli çalıştırma, depolama ve API notlarını ekle.

Son yanıtta sonucu önden söyle; değişen dosyaları, kullanılan depolama kararını, çalışan doğrulama komutlarını ve varsa kalan gerçek sınırları kısa biçimde yaz. Test/build çalışmadıysa “tamamlandı” deme ve kesin hata nedenini belirt.

## PROMPT SONU

---

## Kullanım notu

Hedef repo self-hosted mi serverless mı olduğu bilinmiyorsa prompttaki tek gerekli karar depolama modelidir. Diğer ayrıntılar kaynak sözleşmeden keşfedilebilir. Kaynak örnek Excel’lerde gerçek ticari veri bulunuyorsa bunları uzak servise göndermeden önce anonimleştirin.
