import xml.etree.ElementTree as ET
import requests

# Ticimax sitemap adresinizi girin
sitemap_url = "https://www.paspasofisi.com/sitemap.xml"
response = requests.get(sitemap_url)

with open("sitemap.xml", "wb") as f:
    f.write(response.content)

tree = ET.parse("sitemap.xml")
root = tree.getroot()

# XML namespace tanımları
namespaces = {
    'ns': 'http://www.sitemaps.org/schemas/sitemap/0.9',
    'image': 'http://www.google.com/schemas/sitemap-image/1.1'
}

image_urls = []

# Sitemap içindeki tüm görsel adreslerini topla
for url in root.findall('ns:url', namespaces):
    for img in url.findall('image:image', namespaces):
        loc = img.find('image:loc', namespaces)
        if loc is not None and loc.text:
            image_urls.append(loc.text)

# Görsel URL'lerini txt dosyasına yazdır
with open("urun_gorselleri.txt", "w", encoding="utf-8") as f:
    for img_url in image_urls:
        f.write(img_url + "\n")

print(f"Toplam {len(image_urls)} adet ürün görsel adresi ayıklandı.")