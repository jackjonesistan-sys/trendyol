# Kampanya hesaplayıcı dokümantasyonu

Bu klasör, 5 Ağustos 2026 tarihindeki çalışma ağacını kaynak kabul eder. Davranış değiştiğinde önce kod ve testler, ardından bu belgeler birlikte güncellenmelidir.

- [Mimari ve işletim](mimari.md): bileşenler, uçtan uca akış, kalıcılık, arayüz ve bilinen sınırlar.
- [Sistem spesifikasyonu](sistem_spesifikasyonu.md): Excel sözleşmeleri, hesap kuralları, API ve çıktı dosyaları.
- [Next.js aktarım promptu](nextjs_aktarma_promptu.md): aynı iş davranışını bir Next.js projesine taşıtmak için doğrudan kullanılabilir ana prompt.

Hızlı yerel doğrulama:

```powershell
python -m pip install -r requirements.txt
python -m unittest discover -s tests -v
python app.py
```

Uygulama varsayılan olarak `http://127.0.0.1:5114` adresinde açılır. `app.py` içindeki `debug=True` yerel geliştirme içindir; uygulama bu haliyle internete açılmamalıdır.
