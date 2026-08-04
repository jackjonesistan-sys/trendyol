# Trendyol Kampanya Yönetimi - Mimari ve Çalışma Yapısı

Bu doküman, "Trendyol Kampanya Yönetimi" isimli web uygulamasının tüm mimarisini, arka plan hesaplamalarını, dizin yapısını ve kural motorlarını detaylı bir şekilde açıklamaktadır.

## 1. Dizin ve Dosya Mimarisi

Uygulamanın ana dizini `c:\Users\Tasarımcı\Desktop\trendyol` şeklindedir ve aşağıdaki temel yapılara bölünmüştür:

*   **`app.py`:** Flask tabanlı web sunucusu ve API rotalarını (route) barındıran ana tetikleyici dosyadır. Excel çıktı işlemlerini yönetir.
*   **`komisyon_hesaplayici.py`:** Tüm girdi dosyalarını okuyan, Trendyol'un kademeli komisyon dilimlerini çözen ve kâr-zarar optimizasyonlarını yapan hesaplama motorudur.
*   **`templates/index.html`:** Kullanıcı arayüzünü (UI) barındırır. Tailwind CSS, DaisyUI ve DataTables (jQuery) kullanılarak inşa edilmiştir. Kural motoru ve filtreleme mantıkları burada (frontend JS) yer alır.
*   **`Girdiler/` (Klasör):** Kullanıcının Trendyol panosundan indirdiği *ham* şablon Excel dosyalarını koyduğu yerdir.
*   **`Çıktılar/` (Klasör):** Sistem tarafından işlenen kâr hesaplama tablolarının ve doğrudan Trendyol'a yüklemeye hazır nihai Excel çıktılarının kaydedildiği klasördür.
*   **`docs/` (Klasör):** Mimari dokümantasyonların tutulduğu klasördür.

---

## 2. Girdi Dosyaları ve Görevleri

Sistemin düzgün çalışabilmesi için `Girdiler` klasöründe belirli anahtar kelimelere sahip Excel dosyalarının bulunması zorunludur:

1.  **Komisyon Tarifeleri (`*komisyon*.xlsx` vb.):** Ürünlerin satılacağı fiyata göre (1., 2., 3., 4. dilim alt/üst limitleri) komisyon oranını belirler.
2.  **Güncel Ürünleriniz (`GüncelÜrünleriniz*.xlsx`):** Ürünlerin halihazırdaki `BuyBox Fiyatı` ve varsayılan komisyon bilgilerini içerir. Temel net kâr (Güncel Net) buradan hesaplanır.
3.  **İndirim Uygulanabilecek Ürünler (`İndirim*.xlsx`):** Aktif indirimlerin `Eski Fiyat` ve `Yeni Fiyat` karşılaştırmalarını barındırır. Arayüzdeki **"Mevcut İndirim Oranı (%)"** buradan çekilir.
4.  **Kampanya Şablonları:**
    *   **Avantajlı Ürünler:** İçerisinde `YENİ TSF (FİYAT GÜNCELLE)` sütunu bulunan Trendyol kampanyası şablonudur.
    *   **Flaş Ürünler:** İçerisinde `24 Saat Fiyat` sütunu bulunur.
    *   **Plus Ürünler:** `Plus Fiyat Üst Limiti` barındırır.
    *   **Plus Ek İndirimler:** `Maksimum Girebileceğin Fiyat` barındırır. (Ek indirim %5, %10, %20 olasılıklarına göre hesaplanır).

---

## 3. Hesaplama Motoru (`komisyon_hesaplayici.py`)

Kullanıcı arayüzde **"Verileri Hesapla"** butonuna bastığında API (`/api/calculate`) bu scripti çalıştırır. Çalışma sırası şöyledir:

### A. Verilerin Okunması ve Birleştirilmesi
Tüm Excel dosyalarındaki `BARKOD` (veya `Barkod`) sütunları okunur, içlerindeki boşluklar temizlenir (`BARKOD_CLN`). `Güncel Ürünler` ve `İndirim Uygulanabilecek Ürünler` içerisindeki tüm benzersiz (unique) barkodlar bir havuzda toplanır.

### B. Dinamik Komisyon Bulma (`get_commission_rate`)
Trendyol'un dinamik kuralına göre ürünün satış fiyatı (P), komisyon excelindeki dilimlerle (Range) karşılaştırılır:
*   `P >= 1.Fiyat Alt Limit` -> `1.KOMİSYON` uygulanır.
*   `P` değeri 2. dilim aralığındaysa -> `2.KOMİSYON` uygulanır vs.

*Eğer ürünün özel bir komisyon dilimi yoksa (şablonda bulunamadıysa), Güncel Ürünler dosyasındaki sabit `Komisyon Oranı` devreye girer.*

### C. Karlılık ve Net Hesaplama Matematiği
Her barkod için 5 farklı fiyat ihtimali üzerinden kâr hesaplanır:
1.  **Güncel Net (n1):** `Güncel Fiyat - (Güncel Fiyat * Güncel Komisyon)`
2.  **Avantajlı Net (n2):** `Avantajlı Fiyat (YENİ TSF) - (Avantajlı Fiyat * Avantajlı Komisyon)`
3.  **Flaş Net (n3):** `Flaş 24 Saat Fiyatı - (Flaş Fiyat * Flaş Komisyon)`
4.  **Plus Net (n4):** `Plus Fiyat Üst Limiti - (Plus Fiyat * Plus Komisyon)`
5.  **Plus Ek İndirim Net (n5):**
    *   *Burada üç farklı alt durum hesaplanır:* Maksimum Fiyatın üzerinden **%5**, **%10** ve **%20** indirim yapılarak, yeni düşürülmüş fiyatlara göre komisyonlar tekrar hesaplanır ve netler bulunur.

### D. "Hangisi Daha Karlı?" Mantığı
Hesaplanan `n2`, `n3`, `n4` değerleri, baz alınan `Güncel Net (n1)` ile karşılaştırılır:
*   `n2 >= n1` ise Avantajlı mantıklıdır.
*   Eğer bir ürün hem Avantajlı hem de Flaş ise ve `Flaş Fiyatı >= Avantajlı Fiyatı` şartını da sağlıyorsa Flaş'a öncelik verilir.
*   Tüm geçerli olasılıklar listelenir, en büyük (en kârlı) "Net" değere sahip olan kampanya seçilir.
*   Hiçbiri n1'den büyük değilse sonuç: **Hiçbiri**.

Sonuçlar bir JSON matrisi gibi `Çıktılar/Hesaplanmis_Komisyon_Sonuclari.xlsx`'e kaydedilir.

---

## 4. Frontend (UI) ve Kural Motoru Mimarisi (`index.html`)

Frontend yapısı tamamen asenkron (AJAX - Fetch API) olarak tasarlanmıştır.

### A. DataTables Entegrasyonu
Tablo 16 sütundan oluşur. Arama, sıralama, çift yönlü (alt-üst) sayfalandırma ve dinamik DOM yapısına sahiptir.
**Kısıtlama Filtresi (Mevcut İndirim):** Sadece 'İndirim Uygulanabilecek Ürünler' dosyasında yer alanları görmek için DataTables `.ext.search.push()` yöntemi kullanılarak anlık (gerçek zamanlı) satır filtrelemesi yapılır.

### B. Dinamik Fark (%) Hesabı
Tablodaki *Güncel Fark (%)* sütunu sunucudan sabit gelmez, UI tarafında anlık render edilir. Kullanıcının seçtiği kampanyaya (`userSelection`) göre:
`Fark = ((Seçilen Kampanya Neti - Güncel Net) / Güncel Net) * 100`

**Özel Kural:** Kullanıcı seçimi "Hiçbiri" olarak bırakırsa, *Güncel Fark (%)* sütunu otomatik olarak "Mevcut İndirim Oranı (%)"nı (İndirim şablonundaki eski/yeni fiyata göre hesaplanan oran) gösterir.

### C. Toplu İşlemler Koruması (Validation)
Kullanıcı bir ürünü topluca (Örn: Avantajlı Yap) değiştirmek isterse, JavaScript öncelikle satırın `Avantajlı Ürün Eşleşme Durumu === 'Eşleşti'` olup olmadığına bakar. Eğer ürün Girdiler'deki Avantajlı Excel'inde yoksa, sistem bunu **engeller**.
**İstisna (Bypass):** Sadece "Hangisi Daha Karlı?" verisi `Hiçbiri` olan (hiçbir sistemle eşleşmeyen) yetim ürünlerde bu koruma kaldırılır ve özgür seçime izin verilir.

### D. Dinamik Kural Motoru
Kullanıcıların kendi mantıksal şartlarını yazmasına olanak tanır.
Örn: `[Avantajlı Net] [Güncel Net]'ten [%5] [daha az] ise -> [Flaş Yap]`
Algoritma:
1. Bütün tabloyu döngüye alır (`tableData.forEach`)
2. `n2` (Avantajlı) ve `n3` (Flaş) netlerini alır.
3. Kuraldaki operatörlere göre (`((n2 - n3) / n3) * 100`) yüzdelik formülünü işler.
4. Eşleşirse `row.userSelection = 'Flaş'` olarak atar ve satırı yeniler.

---

## 5. Çıktı Üretim İşlemi (`app.py /api/apply_campaigns`)

Kullanıcı arayüzden **Excel'e Uygula** butonuna bastığında, tablo verisi JSON formatında sunucuya gider.
Sistem, arayüzden gelen çoklu kampanya hedefleme komutunu (`target_type`) okur ("Avantajlı", "Flaş", "Hepsi" vb.). Hangi çıktıların üretilmesi emredildiyse yalnızca onların işlemlerini (I/O) yapar.

**Üretim Mantığı (Row Deletion Pattern):**
1. Girdiler klasöründeki orjinal Trendyol Excel şablonu (örneğin *Avantajlı.xlsx*) `openpyxl` kütüphanesiyle hafızaya alınır.
2. Excel'deki `BARKOD` (B Sütunu veya benzeri) okunur.
3. Okunan barkod, ön yüzden gelen "Seçimler Listesinde" `Avantajlı` olarak işaretlenmiş mi diye bakılır.
4. Eğer işaretlenmemişse (başka bir şeye atanmışsa veya hiçbiri ise) o **satır bellekten (Excel şablonundan) tamamen silinir (`ws.delete_rows`)**.
5. Geriye yalnızca seçilen barkodlar kalır. Fiyat değişiklikleri/girişleri yapılır.
6. İşlenen dosya `Çıktılar/Avantajlı_Urunler_[Tarih].xlsx` adıyla tertemiz bir Trendyol yükleme formatında kaydedilir.

*(Not: Avantajlı Çoğaltma işleminde, eğer kullanıcı takvimden Başlangıç-Bitiş seçmişse, bu tarihler arasındaki her bir gün için ayrı ayrı [Gün_Ay_Yıl] şeklinde birden çok Avantajlı Excel dosyası çoğaltılarak üretilir.)*

Ayrıca, Plus Ek İndirim seçilen ürünler %5, %10 ve %20 seçimlerine göre arka planda kendi oransal net fiyatlarına göre tek bir dosyada veya farklı satırlarda çıktıya dahil edilir.

En son, hiçbir kampanyaya dahil edilmeyen (`userSelection === 'Hiçbiri'`) ürünler toparlanarak `Çıktılar/Uygulanmayan_Urunler_Raporu.xlsx` dosyasına bilgilendirme amacıyla yazılır.
