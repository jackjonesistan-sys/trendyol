# Trendyol Kampanya Yönetimi - Sistem Sözleşmesi

## API

### `POST /api/calculate`

İçerik türü `multipart/form-data` olmalıdır.

Dosya alanları:

| Alan | Tür | Zorunlu | Ayırt edici sütunlar |
|---|---|---:|---|
| `discount` | İndirim uygulanabilecek ürünler | Evet | `BARKOD`, `Eski Fiyat`, `YENİ Fiyat`, `Durum` |
| `commission` | Komisyon tarifesi | Evet | `BARKOD`, fiyat limitleri, `1.KOMİSYON`-`4.KOMİSYON`, `TARİFE GRUBU` |
| `current` | Güncel ürünler | Evet | `Barkod`, `Komisyon Oranı`, piyasa ve Trendyol satış fiyatları |
| `advantage` | Avantajlı ürün | Hayır | `BARKOD`, `1 YILDIZ ÜST FİYAT`, `YENİ TSF (FİYAT GÜNCELLE)` |
| `flash` | Flaş ürün | Hayır | `Barkod`, `24 Saat Fiyat`, `Güncellenecek Fiyat` |
| `plus` | Plus ürün | Hayır | `Barkod`, `Plus Fiyat Üst Limiti`, `Plus Komisyon Teklifi` |
| `plus_extra` | Plus ek indirim | Hayır | `Barkod`, `Maksimum Girebileceğin Fiyat`, `Kampanyalı Satış Fiyatı` |
| `counter` | Karşılamalı kampanya | Hayır | Plus Ek ile aynı; alan adıyla ayrılır |

Karşılamalı sayısal alanları: `min_sepet`, `toplam_indirim`, `trendyol_oran` (`0`-`100`). Başarı yanıtı `success`, `message` ve hesap sonucu yolunu içerir. Doğrulama hataları HTTP 400 döner.

Zorunlu türler ilk kullanımda yüklenir; sonraki isteklerde manifestte geçerli kopyaları varsa tekrar gönderilmeleri gerekmez. Gönderilen tür aynı türün önceki dosyasını değiştirir, gönderilmeyen türler korunur. Başarı yanıtındaki `uploads` nesnesi her türün özgün dosya adını ve ISO/gösterim biçimli yükleme zamanını içerir.

### `GET /api/data`

`Çıktılar/Kampanya_Hesaplama_Sonuclari.xlsx` içeriğini JSON satırları olarak döner. Dosya yoksa HTTP 200 ile `needs_calculation: true` döner.

Hesap satırlarında en az şu kontrol alanları bulunur:

- `İndirim Uygulanabilir`
- kampanya eşleşme durumları
- `Uygulanabilir Kampanyalar`
- `İlk Kampanya Seçimi`
- `Hangisi Daha Karlı?`
- fiyat, komisyon ve kalan net alanları
- `Düşülebilecek Dip Fiyat (TL)` ve indirim alanları

### `POST /api/apply`

JSON gövdesi:

```json
{
  "target_type": "Hepsi",
  "selections": {
    "BARKOD-1": "Avantajlı",
    "BARKOD-2": "Hiçbiri"
  },
  "visibleColumns": ["Barkod", "Uygulanan Kampanya", "Hangisi Karlı?"]
}
```

Geçerli seçimler: `Hiçbiri`, `Avantajlı`, `Flaş`, `Plus`, Plus Ek `%5/%10/%20` ve `Karşılamalı Kampanya`. Geçerli hedefler bunlara ek olarak `Hepsi` ve genel `Plus Ek İndirim` hedefidir. Bilinmeyen değerler HTTP 400 ile reddedilir.

## Rapor sütunu sözleşmesi

Sayfa ve özet Excel'in tam sırası:

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
14. Uygulanan Kampanya Fiyat
15. Uygulanan Kampanya Net
16. Uygulanan Kampanya Komisyon
17. Uygulanabilecek İndirim (TL)
18. Uygulanabilecek İndirim (%)
19. Uygulanan İndirim (TL)
20. Uygulanan İndirim (%)
21. Ekstra Uygulanabilir İndirim (TL)
22. Ekstra Uygulanabilir İndirim (%)
23. Hangisi Karlı?
24. Düşülebilecek Dip Fiyat (TL)

`visibleColumns` bu sözleşmenin alt kümesidir. Sunucu istemci sırasına güvenmez; kanonik sırayı koruyarak yalnız görünür sütunları yazar.
