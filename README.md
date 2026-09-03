# ac2f pack

CorelDRAW için VBA eklenti paketi.

Vektörlerin çevresindeki çizgilerin toplam uzunluğunu ölçer ve bu ölçüme
dayanarak 3'lü LED modül yerleşimi için gereken adetleri hesaplar.

| Modül | İş |
|---|---|
| `ac2fCore` | Ortak çekirdek: ayarlar, birim/sayı yardımcıları, ölçüm motoru |
| `ac2fLength` | **Uzunluk ölçümü** — toplam kontur uzunluğu |
| `ac2fLedModule` | **3'lü LED modül hesabı** — modül, LED, güç, güç kaynağı adedi |
| `ac2fMenu` | Ana menü, ayarlar, hakkında |

## Makrolar

Makro Yöneticisi'nde `ac2fPack` projesi altında görünürler.

| Makro | Açıklama |
|---|---|
| `ac2fPack` | **Ana menü** — hepsine buradan ulaşılır |
| `ac2fUzunlukOlc` | Seçili vektörlerin toplam kontur uzunluğunu ölçer |
| `ac2fUzunlukEtiketle` | Aynı ölçümü yapıp sonucu sayfaya metin olarak koyar |
| `ac2fLedModulHesapla` | Kayıtlı ayarlarla 3'lü LED modül adedini hesaplar |
| `ac2fLedHizliHesap` | Modül aralığını sorarak tek seferlik hesaplar |
| `ac2fAyarlar` | Ayarları düzenler |
| `ac2fAyarlariSifirla` | Ayarları varsayılana döndürür |
| `ac2fHakkinda` | Sürüm ve içerik bilgisi |

## Kurulum

Kısa yol:

```powershell
powershell -ExecutionPolicy Bypass -File tools\build.ps1
```

Sonra CorelDRAW'da `Alt+F11` → `ac2fPack` projesi → `File > Import File` ile
`build\` altındaki dört `.bas` dosyasını içe aktarın.

Adım adım anlatım ve ekran yolları: **[docs/kurulum.md](docs/kurulum.md)**

## Kullanım

```
Nesneleri seç  →  ac2fPack  →  1 (uzunluk)  veya  3 (LED modül)
```

Ayrıntılı kullanım, ayarların anlamı ve hesap yöntemleri:
**[docs/kullanim.md](docs/kullanim.md)**

## Hesap özeti

Ölçüm, her **alt yolu (kontur)** ayrı ele alır:

- **Kapalı kontur** → `yukarı_yuvarla(uzunluk / aralık)` modül
- **Açık yol** → `aşağı_yuvarla(uzunluk / aralık) + 1` modül (iki uç da modül alır)
- Her kontur en az *"kontur başına en az modül"* kadar modül alır

Toplam modül, düzeltme katsayısıyla çarpılıp yukarı yuvarlanır. Ardından:

```
LED adedi   = modül × modül başına LED (varsayılan 3)
Toplam güç  = modül × modül gücü
Gerekli güç = toplam güç × (1 + güvenlik payı)
Güç kaynağı = yukarı_yuvarla(gerekli güç / kaynak kapasitesi)
```

## Geliştirme

```bash
python3 tools/lint.py     # içe aktarmadan önce statik denetim
sh tools/build.sh         # build/ altına Windows-1254 + CRLF üret
```

Kaynak dosyalar depoda **UTF-8 + LF**, VBE'ye aktarılan kopyalar
**Windows-1254 + CRLF** biçimindedir. Gerekçesi ve projenin nasıl ortaya
çıktığı: **[docs/gelistirme.md](docs/gelistirme.md)**

## Gereksinimler

- CorelDRAW Graphics Suite (VBA desteği kurulu olmalı — X6 ve sonrası)
- Windows

## Sürüm

1.0.0 — bkz. [CHANGELOG.md](CHANGELOG.md)
