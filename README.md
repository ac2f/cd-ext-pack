# ac2f pack

CorelDRAW için VBA eklenti paketi.

Vektörlerin çevresindeki çizgilerin toplam uzunluğunu ölçer, bu ölçüme
dayanarak 3'lü LED modül yerleşimi için gereken adetleri hesaplar, kutu
harf yan bordürünün açınımını çıkarıp kesime hazır derzli şeritler çizer
ve alüminyum kompozit paneller için V derz çizgilerini üretir.

| Modül | İş |
|---|---|
| `ac2fCore` | Ortak çekirdek: ayarlar, birim/sayı yardımcıları, ölçüm motoru |
| `ac2fLength` | **Uzunluk ölçümü** — toplam kontur uzunluğu |
| `ac2fLedModule` | **3'lü LED modül hesabı** — modül, LED, güç, güç kaynağı adedi |
| `ac2fBoxLetter` | **Kutu harf şeridi** — bordür açınımı ve derz yerleşimi |
| `ac2fPanel` | **ACP panel derzi** — alüminyum kompozit V derz yerleşimi |
| `ac2fSettings` | **Ayar sayfası ve profiller** — tüm ayarlar tek yerde |
| `ac2fMenu` | Ana menü, hakkında |

## Makrolar

Makro Yöneticisi'nde `ac2fPack` projesi altında görünürler.
Arayüz dili İngilizcedir; bu belgeler Türkçedir.

| Makro | Açıklama |
|---|---|
| `ac2fPack` | **Ana menü** — hepsine buradan ulaşılır |
| `ac2fMeasureLength` | Seçili vektörlerin toplam kontur uzunluğunu ölçer |
| `ac2fLabelLength` | Aynı ölçümü yapıp sonucu sayfaya metin olarak koyar |
| `ac2fLedModuleCount` | Kayıtlı ayarlarla 3'lü LED modül adedini hesaplar |
| `ac2fLedQuickCount` | Modül aralığını sorarak tek seferlik hesaplar |
| `ac2fBoxLetterStrip` | Kutu harf bordürünün açınımını hesaplar ve derzli şeridi çizer |
| `ac2fBoxLetterStripProfile` | Profil seçip, istenen değeri o seferliğine değiştirip çizer |
| `ac2fBoxLetterReport` | Aynı hesabı yapar, çizim yapmaz |
| `ac2fPanelGroove` | Dörtgen çevresine ACP V derz çizgilerini çizer |
| `ac2fSettings` | **Tüm ayarlar ve profiller — tek sayfa** |
| `ac2fProfiles` | Doğrudan profil yönetimi |
| `ac2fResetAllSettings` | Ayarları varsayılana döndürür |
| `ac2fAbout` | Sürüm ve içerik bilgisi |

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
Nesneleri seç  →  ac2fPack  →  1 (uzunluk) / 3 (LED modül) / 5 (kutu harf şeridi)
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

## ACP panel V derzi

Dörtgeni seçin, `ac2fPanelGroove` çalıştırın ve **`5cm out`** yazın.

100×200 bir dörtgen + 5 cm derz → **110×210** plaka. Dört derz çizgisi
plakayı boydan boya geçer ve kesişimleri tam **100×200** üzerine düşer:

```
        <------------- 110 ------------->
    +---+---------------------------+---+  ^
    |   |            2              |   |  |
    +---o===========================+---+  |     o = baslangic noktasi
    |   #                           #   |  |     2 = derz sirasi
    |   #                           #   |  |
    | 1 #                           # 3 | 210
    |   #                           #   |  |
    |   #                           #   |  |
    +---+===========================o---+  |
    |   |            4              |   |  |
    +---o---------------------------+---+  v
        ^ 1 alttan baslar
```

Saat yönünde, **nesne sırası = kesim sırası**:

| # | Derz | Başlangıç | Boy |
|---|---|---|---|
| 1 | Sol | **Alttan** | 210 (tam yükseklik) |
| 2 | Üst | Soldan | 110 (tam genişlik) |
| 3 | Sağ | Üstten | 210 |
| 4 | Alt | Sağdan | 110 |

Her çizgi adlandırılır (`ac2f groove 1 left` …), böylece sıra Nesne
Yöneticisi'nden görülebilir.

### İki yön

| Giriş | Anlamı | 100×200 + 5 |
|---|---|---|
| `5cm out` | Seçim **bitmiş** ölçü, plaka büyür | plaka 110×210, katlanmış 100×200 |
| `5cm in` | Seçim **plaka**, derzler içeri kaçar | plaka 100×200, katlanmış 90×190 |

Ölçü `5cm`, `50mm` ya da düz `50` (= mm) yazılabilir.

### Renk ve gruplama

Siyah = plaka dış hattı (kesim), mavi = V derz. Yapı:

```
ac2f ACP panel 110x210        (grup)
├── ac2f ACP sheet 110x210    (dörtgen)
└── ac2f ACP grooves          (grup)
    ├── ac2f groove 1 left
    ├── ac2f groove 2 top
    ├── ac2f groove 3 right
    └── ac2f groove 4 bottom
```

## Ayarlar ve profiller

**Tüm ayarlar tek sayfada.** Ana menüden `9` ile açılır; 25 ayarın hepsi
numaralı olarak listelenir:

```
SETTINGS  (profile: Aluminium 2mm)

-- LED MODULE --
 1 Module spacing     100.00 mm
 ...
-- BOX LETTER --
 9 Thickness            2.00 mm
11 Flexibility          0.80
...

N=value   change        ?N   explain
P         profiles      R    reset
Enter     close
```

| Yazın | Ne olur |
|---|---|
| `11=0.8` | 11 numaralı ayarı 0,8 yapar |
| `11=0.8 16=0.1` | Birden çok ayarı tek seferde değiştirir |
| `?11` | 11 numaralı ayarın **ayrıntılı açıklamasını** gösterir |
| `Flexibility=0.9` | Numara yerine ad da kullanılabilir |
| `P` | Profil menüsü |
| `R` | Varsayılanlara dön |

Her ayarın `?N` ile açılan uzun bir açıklaması vardır: ne işe yaradığı,
neyi etkilediği, büyütünce/küçültünce ne olduğu ve tipik değerler.

### Profiller

`P` ile açılır. Bir profil, 22 ayarın tamamının bir ad altında
saklanmasıdır.

| Yazın | Ne olur |
|---|---|
| `S 3mm alu` | Şu anki ayarları bu adla kaydeder |
| `L 3mm alu` | Profili yükler |
| `D 3mm alu` | Profili siler |
| `M` | Malzeme ön ayarı uygular |

### Profille çalıştırma + geçici değişiklik

Ana menü `6` (`ac2fBoxLetterStripProfile`): profili seçersiniz, sonra
istediğiniz değeri **yalnız o çalıştırma için** değiştirirsiniz. Geçici
değiştirilen satırlar `*` ile işaretlenir ve kayıtlı ayarlarınıza
yazılmaz.

## Kutu harf şeridi

Harf konturunu seçip `ac2fBoxLetterStrip` çalıştırın; sağ tarafa her kontur
için düz bir şerit çizilir:

```
mavi = eğri derzi      pembe = köşe derzi      kırmızı = rulo kesim yeri
```

**Açınım boyu** nötr eksenden hesaplanır. Basit kapalı bir eğride toplam
dönüş 2π olduğundan (Steiner), açınım `P ∓ 2π·g` olur — `g`, nötr eksenin
vektörden malzemeye doğru ötelenmesidir. Delik (counter) konturlarında
işaret ters çevrilir.

**Derz aralığı** her segment için ayrı ayrı, üç sınırın en küçüğüdür:

```
s₁ = R × (derz ağzı / derz derinliği) × esneklik oranı    ← derz kapanma sınırı
s₂ = √(8 × R × yüzey toleransı)                            ← düzlük (sehim) sınırı
s₃ = en çok derz aralığı                                   ← tavan
```

Sonuç `en az derz aralığı`na kırpılır. Köşelerde dönüş açısı tek derzin
karşılayabileceğinden büyükse birden çok derz açılır.

**Esneklik oranı** kalibrasyon içindir: ilk işten sonra gerçek sonucunuza
göre büyütüp küçültün. Malzeme ön ayarları (alüminyum/galvaniz/kalın/paslanmaz)
ölçülmüş değer değil, başlangıç noktasıdır.

Ayrıntı ve sınırlar: **[docs/kullanim.md](docs/kullanim.md)**

## Geliştirme

```bash
python3 tools/lint.py     # içe aktarmadan önce statik denetim
python3 tools/skeleton.py # büyük düzenlemelerden sonra kontrol akışı karşılaştırması
sh tools/build.sh         # build/ altına CRLF kopya üret (isteğe bağlı)
```

Kaynak **saf ASCII**'dir; hiçbir Türkçe karakter içermez. Bu yüzden
`.bas` dosyaları her makinede kod sayfası dönüşümü olmadan doğrudan içe
aktarılabilir. `build.sh` yalnız CRLF satır sonu üretir, artık zorunlu
değildir. Ayrıntı: **[docs/gelistirme.md](docs/gelistirme.md)**

## Gereksinimler

- CorelDRAW Graphics Suite (VBA desteği kurulu olmalı — X6 ve sonrası)
- Windows

## Sürüm

1.4.0 — bkz. [CHANGELOG.md](CHANGELOG.md)
