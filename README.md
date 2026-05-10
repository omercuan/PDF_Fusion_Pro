#  PDF Fusion Pro

Modern, hızlı ve kullanımı kolay PDF araç seti. PyQt6 ile geliştirilmiş masaüstü uygulaması.

## Özellikler

| Özellik | Açıklama |
|---|---|
| 📋 **PDF Birleştir** | Birden fazla PDF'i sürükle-bırak ile tek dosyada birleştir |
| 📝 **PDF → Word** | PDF dosyalarını düzenlenebilir `.docx` formatına dönüştür |
| 📑 **Word → PDF** | `.docx` / `.doc` dosyalarını profesyonel PDF'ye dönüştür |

## Masaüstü Kısayolu Oluşturma (Logolu)

Uygulama **ilk açıldığında** `icon.ico` dosyasını otomatik olarak proje klasörüne oluşturur.

1. `python app.py` komutuyla uygulamayı **bir kez** çalıştır → `icon.ico` oluşur
2. Masaüstüne sağ tıkla → **Yeni > Kısayol**
3. Konum olarak şunu gir (kendi yoluna göre düzenle):
   ```
   python "C:\...\modernize-pdf-conversion-tool\app.py"
   ```
4. Kısayola sağ tıkla → **Özellikler** → **İkon Değiştir**
5. Proje klasöründeki `icon.ico` dosyasını seç → Tamam

## Kurulum

### Gereksinimler
- Python 3.10 veya üzeri
- Microsoft Word (yalnızca Word → PDF özelliği için)

### Adımlar

```bash
# 1. Repoyu klonla
git clone https://github.com/KULLANICI_ADIN/modernize-pdf-conversion-tool.git
cd modernize-pdf-conversion-tool

# 2. Sanal ortam oluştur (önerilir)
python -m venv venv
venv\Scripts\activate      # Windows
# source venv/bin/activate  # macOS/Linux

# 3. Bağımlılıkları yükle
pip install -r requirements.txt

# 4. Uygulamayı çalıştır
python app.py
```

## Bağımlılıklar

```
PyQt6      → Grafik arayüz (GUI)
pypdf      → PDF okuma ve birleştirme
pdf2docx   → PDF'den Word'e dönüşüm
docx2pdf   → Word'den PDF'e dönüşüm (MS Word gerektirir)
```

## Executable Oluşturma (PyInstaller)

Kendi `.exe` dosyanı oluşturmak için:

```bash
pip install pyinstaller
pyinstaller app.spec
```

Çıktı `dist/` klasöründe oluşur.

## Proje Yapısı

```
modernize-pdf-conversion-tool/
├── app.py           # Ana uygulama kaynak kodu
├── app.spec         # PyInstaller build konfigürasyonu
├── icon.ico         # Otomatik oluşur (ilk çalıştırmada) — .gitignore'da
├── requirements.txt # Python bağımlılıkları
└── README.md        # Bu dosya
```

## Notlar

- **Word → PDF** özelliği yalnızca bilgisayarınızda **Microsoft Word** yüklüyse çalışır.
- PDF → Word dönüşümünde karmaşık düzenler tam olarak korunmayabilir; bu `pdf2docx` kütüphanesinin doğal bir sınırlamasıdır.

## Lisans

MIT License — dilediğiniz gibi kullanabilirsiniz.
