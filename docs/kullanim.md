# Kullanım

## Ana menü

`ac2fPack` makrosunu çalıştırın; işlem numarasını soran bir kutu açılır.

> Arayüz dili İngilizcedir. Bu belgeler Türkçedir; makro adları ve ekran
> metinleri İngilizce yazıldığı gibi verilmiştir.

```
  1  -  Length measurement
  2  -  Length measurement + label on page
  3  -  3-LED module count
  4  -  3-LED module count (ask for spacing)
  5  -  Box letter strip
  6  -  Box letter strip (pick profile, tweak, draw)
  7  -  Box letter report (no drawing)
  8  -  ACP panel V-grooves
  9  -  Settings and profiles
 10  -  About
```

Her makro doğrudan da çağrılabilir; ana menü yalnızca kolaylık içindir.

---

## 1. Uzunluk ölçümü

Nesneleri seçin ve `ac2fMeasureLength` çalıştırın.

```
TOTAL LENGTH
   3.842,17 mm   |   384,22 cm   |   3,842 m

SELECTION
   Objects             : 7
   Outlines (sub-paths): 12
      closed           : 12
      open             : 0

OUTLINE STATISTICS
   Longest             : 812,44 mm
   Shortest            : 96,10 mm
   Average             : 320,18 mm

PER OBJECT
   1. Curve 1 - 812,44 mm (2 outlines)
   ...
```

### Neler ölçülür

- Eğriler, dikdörtgen/elips/çokgen gibi şekiller, metinler
- **Gruplar** — içindeki her nesne ayrı ayrı
- **PowerClip** içerikleri

Eğri olmayan nesneler ölçüm sırasında geçici bir kopya üzerinden eğriye
çevrilir; kopyalar silinir ve belge ilk hâline döndürülür. Çiziminizde
kalıcı bir değişiklik olmaz.

### Ölçülemeyenler

Bitmap'ler, boş metinler ve uzunluğu sıfır olan nesneler sessizce atlanır.

> **Kontur = alt yol.** Bir harfin dışı ve içi (örneğin "O") iki ayrı
> konturdur ve ikisi de toplama girer.

---

## 2. Sayfaya etiket

`ac2fLabelLength` aynı ölçümü yapar, ardından seçimin 8 mm altına
artistik metin olarak sonucu yazar:

```
Total length: 3.842,17 mm  (3,84 m)
```

Etiket normal bir metin nesnesidir; taşıyabilir, biçimlendirebilir,
silebilirsiniz.

---

## 3. 3'lü LED modül hesabı

Nesneleri seçin ve `ac2fLedModuleCount` çalıştırın.

```
RESULT
   Modules             : 42
   LEDs                : 126  (3 LEDs per module)

POWER
   Total power         : 30,24 W
   With 20% margin     : 36,29 W
   Power supplies      : 1 x 60 W

MEASUREMENT
   Raw outline length  : 3.842,17 mm
   Length used         : 3.842,17 mm
   Outlines            : 12

SETTINGS USED
   Module spacing      : 100,00 mm
   Method              : Perimeter based (full outline)
   Correction factor   : 1,00
   Module power        : 0,72 W
```

`ac2fLedQuickCount` aynı işi yapar ama modül aralığını o seferlik sorar ve
kayıtlı ayarı değiştirmez. Farklı aralıkları hızlıca denemek için kullanışlıdır.

### Hesap nasıl yapılır

Her kontur **ayrı ayrı** hesaplanır — böylece kısa parçalar da modülsüz kalmaz:

| Kontur türü | Modül adedi |
|---|---|
| Kapalı (halka) | `yukarı_yuvarla(uzunluk / aralık)` |
| Açık (uçlu yol) | `aşağı_yuvarla(uzunluk / aralık) + 1` |

Açık yolda iki uca da modül konduğu için bir fazla çıkar. Her kontur en az
*"kontur başına en az modül"* kadar modül alır.

Konturların toplamı düzeltme katsayısıyla çarpılıp yukarı yuvarlanır:

```
LED adedi   = modül × modül başına LED
Toplam güç  = modül × modül gücü
Gerekli güç = toplam güç × (1 + güvenlik payı / 100)
Güç kaynağı = yukarı_yuvarla(gerekli güç / kaynak kapasitesi)
```

### Hesap yöntemleri

**1 — Çevre bazlı (varsayılan).** Her konturun tam uzunluğu kullanılır.
Modüller çizginin *üzerine* dizildiğinde doğrudur: kanal harflerin yan
duvarına yapıştırılan çevre aydınlatması, ışıklı şerit güzergâhı, neon
flex boyu gibi işler.

**2 — Orta hat tahmini.** Her konturun **yarısı** kullanılır. Harf, iç ve dış
olmak üzere iki çizgiyle çizilmiş kapalı bir şerit ise (içi boş kontur çifti)
çevre gerçek yolun yaklaşık iki katını verir; bu yöntem bunu düzeltir.

> Yöntem 2 bir **tahmindir**. Şerit genişliği her yerde eşitse iyi sonuç verir,
> değişken kalınlıkta sapar. İlk işinizde çıkan sayıyı sahadaki gerçek
> adetle karşılaştırıp farkı **düzeltme katsayısına** yazın; sonraki hesaplar
> tutar.

---

## 4. Kutu harf şeridi

Harf konturlarını seçin ve `ac2fBoxLetterStrip` çalıştırın. Çizimin sağ
tarafına, her kapalı kontur için bir düz şerit çizilir.

```
RESULT
   Strips              : 3
   Total developed     : 1.842,66 mm   |   184,27 cm   |   1,843 m
   Total grooves       : 96  (8 corner)
   Smallest radius     : 12,40 mm
   Coil required       : ~1 x 3.000 mm (20 mm joint)

STRIPS
   1. Curve 1 #1
      developed 1.204,18 mm (raw 1.211,22)  grooves 58  min R 18,7
   2. Curve 1 #2 [hole]
      developed  638,48 mm (raw  631,44)  grooves 38  min R 12,4
```

Çizimdeki renkler:

| Renk | Anlamı |
|---|---|
| Siyah | Şerit dış hattı (kesim) |
| Mavi | Eğri derzi (`blue = curve groove`) |
| Pembe | Köşe derzi (`pink = corner groove`) |
| Kırmızı | Rulo boyu aşıldığında kesim/ek yeri (`red = cut here`) |

`ac2fBoxLetterReport` aynı hesabı yapar ama hiçbir şey çizmez — ayar denemek
için hızlıdır.

> Şerit yalnızca **kapalı** konturlardan çıkarılır. Açık yollar atlanır.

### Açınım boyu nereden geliyor

Sac büküldüğünde uzunluğu koruyan çizgi nötr eksendir; malzemenin dış yüzü
uzar, iç yüzü kısalır. Vektörünüz şeridin hangi yüzünü temsil ediyorsa,
nötr eksen ondan `g` kadar ötededir:

| Referans yüzey | g |
|---|---|
| Vektör = şeridin dış yüzü *(varsayılan)* | `(1 − K) × kalınlık` |
| Vektör = şeridin iç yüzü | `−K × kalınlık` |
| Vektör = nötr eksen | `0` |

Basit kapalı bir eğride toplam dönüş her zaman 2π olduğu için açınım boyu

```
açınım = kontur boyu − 2π × g        (dış kontur)
açınım = kontur boyu + 2π × g        (delik / counter)
```

Delik konturları otomatik bulunur (sınırlayıcı kutusu bir başkasının içinde
kalan alt yol deliktir) ve raporda `[hole]` diye işaretlenir.

**Büyüklük hissi:** 1 mm alüminyumda düzeltme ≈ 3,5 mm; 3 mm'de ≈ 10,6 mm.
Kontur boyundan bağımsızdır — yalnız kalınlığa bağlıdır.

### Derz aralığı nereden geliyor

Her segment için üç sınırın en küçüğü alınır:

| Sınır | Formül | Neyi engeller |
|---|---|---|
| Derz kapanması | `R × (derz ağzı / derz derinliği) × esneklik` | Derz kapanmadan malzemenin sıkışması |
| Yüzey düzlüğü | `√(8 × R × yüzey toleransı)` | Derzler arası düz yüzün göze çarpması |
| Tavan | `en çok derz aralığı` | Çok seyrek derz |

Sonuç `en az derz aralığı`na kırpılır. Köşelerde (dönüş > köşe eşiği) dönüş
açısı tek derzin karşılayabileceğinden büyükse birden çok derz açılır.

`R` her segmentin yarıçapıdır; düz segmentlere derz açılmaz.

### Esneklik oranını kalibre etmek

Ön ayarlar (alüminyum, galvaniz, kalın alüminyum, paslanmaz) **ölçülmüş
değer değildir** — başlangıç noktasıdır. Doğru yol:

1. Bir harf için şeridi çıkarın ve gerçekten bükün.
2. Derzler kapanmıyor, malzeme sıkışıyorsa → esneklik oranını **küçültün**
   (daha sık derz).
3. Gereğinden fazla derz varsa, kontur zaten rahat dönüyorsa → **büyütün**.
4. Yüzeyde köşeli izler görünüyorsa esnekliğe dokunmayın, **yüzey
   toleransını** küçültün.

Bulduğunuz değer o malzeme + kalınlık için kalıcıdır.

### Bu modülün sınırları

- **Segmentler dairesel yay kabul edilir.** Gerçek bezier yaylarda yarıçap
  hatası %0,05 mertebesindedir — ihmal edilebilir. Ancak tek bir segment
  içinde dönüm noktası varsa (S kıvrımı) model onu tek yönlü yay sanar ve o
  segmentte derzi seyrek koyar. Böyle bir yerde düğüm ekleyip segmenti
  ikiye bölmek sorunu çözer.
- **Toplam açınım boyu her koşulda kesindir**; yukarıdaki durum yalnız ara
  derz konumlarını etkiler (ölçülen en kötü sapma 0,06 mm).
- **Şerit tek parça çizilir.** Rulo boyu aşılıyorsa kesim yerleri kırmızı
  işaretlenir; parçalar ayrı ayrı çizilmez.
- **Derz kesiti çizilmez** — konumu çizilir. V ağzının açısı ve derinliği
  makinenizin/bıçağınızın işidir.
- Çok yoğun işlerde şerit başına 5.000, iş başına 200 şerit sınırı vardır.

---

## 5. ACP panel V derzi

Alüminyum kompozit kutu/harf panellerinde, CAM programında V bıçağıyla
açılacak derz çizgilerini üretir.

Dörtgeni seçin, `ac2fPanelGroove` çalıştırın, tek satır yazın:

```
5cm out
```

### Yerleşim

100×200 bir dörtgen, 5 cm derz, `out`:

```
        <------------- 110 ------------->
    +---+---------------------------+---+  ^
    |   |             2             |   |  |
    +---o===========================+---+  |    o = baslangic noktasi
    |   #                           #   |  |    # ve = : derz cizgisi
    |   #                           #   |  |
    | 1 #                           # 3 | 210
    |   #                           #   |  |
    |   #                           #   |  |
    +---+===========================o---+  |
    |   |             4             |   |  |
    +---o---------------------------+---+  v
        ^ 1 numarali derz ALTTAN baslar
```

Plaka **110×210** olur. Dört çizgi plakayı boydan boya geçer;
kesişimleri tam **100×200** üzerine, yani asıl ölçüye düşer.

### Sıra ve başlangıç noktaları

Saat yönünde. **Nesne sırası = kesim sırası** — çizgiler bu sırayla
oluşturulur ve adlandırılır.

| # | Ad | Konum | Başlangıç | Boy |
|---|---|---|---|---|
| 1 | `ac2f groove 1 left` | Sol, kenardan 5 içeride | **Alttan** yukarı | 210 |
| 2 | `ac2f groove 2 top` | Üst, kenardan 5 aşağıda | **Soldan** sağa | 110 |
| 3 | `ac2f groove 3 right` | Sağ, kenardan 5 içeride | **Üstten** aşağı | 210 |
| 4 | `ac2f groove 4 bottom` | Alt, kenardan 5 yukarıda | **Sağdan** sola | 110 |

Başlangıç noktası, CAM programının bıçağı daldıracağı uçtur; çizgi bu
yönde oluşturulur.

### İki yön

| Giriş | Seçim ne demek | 100×200 + 5 sonucu |
|---|---|---|
| `5cm out` | **Bitmiş** ölçü — plaka her yandan bir derz büyür | plaka 110×210, katlanınca 100×200 |
| `5cm in` | **Plaka** ölçüsü — derzler içeri kaçar | plaka 100×200, katlanınca 90×190 |

Ölçüyü `5cm`, `50mm` ya da düz `50` (= mm) yazabilirsiniz. Yön
yazmazsanız son kullandığınız yön geçerli olur. Yazdığınız değerler ayar
sayfasına kaydedilir.

### Renkler ve gruplama

| Renk | Anlamı |
|---|---|
| Siyah | Plaka dış hattı (kesim) |
| Mavi | V derz çizgisi |

```
ac2f ACP panel 110x210        (dis grup)
├── ac2f ACP sheet 110x210    (plaka dortgeni)
└── ac2f ACP grooves          (derz grubu)
    ├── ac2f groove 1 left
    ├── ac2f groove 2 top
    ├── ac2f groove 3 right
    └── ac2f groove 4 bottom
```

### Bilinmesi gerekenler

- **Eksene paralel sınırlayıcı kutu kullanılır.** Döndürülmüş bir
  dörtgen seçerseniz onun sınırlayıcı kutusu alınır, döndürülmüş hâli
  değil. Önce döndürmeyi sıfırlayın.
- **Kaynak dörtgeniniz silinmez** (ayar 25 ile değiştirilebilir).
  `in` modunda yeni plaka dış hattı tam sizin dörtgeninizin üstüne
  düşer; CAM'in aynı yeri iki kez kesmemesi için ya kaynağı silin ya da
  ayarı `0` yapın.
- Birden çok dörtgen seçebilirsiniz; her biri ayrı panel olur.
- `in` modunda derz, kenar uzunluğunun yarısından büyük olamaz; öyle bir
  dörtgen atlanır ve raporda belirtilir.

---

## 6. Ayarlar — tek sayfa

Ana menüden `9` (`ac2fSettings`). Paketin **bütün** ayarları tek listede:

```
SETTINGS  (profile: Aluminium 2mm)

-- LED MODULE --
 1 Module spacing     100.00 mm
 2 LEDs per module         3
 3 Module power         0.72 W
 4 Power supply           60 W
 5 Safety margin          20 %
 6 Min per outline         1
 7 Count method            1
 8 Correction factor    1.00
-- BOX LETTER --
 9 Thickness            2.00 mm
10 Strip height           80 mm
11 Flexibility          0.80
12 Reference face          1
13 K factor             0.42
14 Groove depth ratio   0.75
15 Max groove mouth     1.20 mm
16 Surface tolerance    0.15 mm
17 Min groove spacing    3.0 mm
18 Max groove spacing   60.0 mm
19 Corner threshold     5.00 deg
20 Coil length          3000 mm
21 Joint allowance        20 mm
22 Strip gap              10 mm
-- ACP PANEL --
23 Fold size            50.00 mm
24 Fold direction           1
25 Keep source              1

N=value   change        ?N   explain
P         profiles      R    reset
Enter     close
```

### Komutlar

| Yazın | Ne olur |
|---|---|
| `11=0.8` | 11 numaralı ayarı 0,8 yapar |
| `11=0.8 16=0.1` | Birden çok ayarı tek seferde değiştirir |
| `11=0.8; 16=0.1` | Noktalı virgül de olur |
| `?11` | 11 numaralı ayarın **ayrıntılı açıklaması** |
| `Flexibility=0.9` | Numara yerine ayar adı |
| `?Flexibility` | Ada göre açıklama |
| `P` | Profil menüsü |
| `R` | Tüm ayarları varsayılana döndür (profiller korunur) |
| Boş + Enter | Kapat |

Hem `0.8` hem `0,8` kabul edilir. Değerler ayarın alt/üst sınırına
kırpılır, tam sayı olması gereken ayarlar yuvarlanır.

### `?N` — ayrıntılı açıklama

Her ayarın uzun bir açıklaması vardır: ne işe yaradığı, hangi hesabı
etkilediği, büyütünce/küçültünce ne olduğu ve tipik değerler. Örnek:

```
16.  Surface tolerance  (mm)
----------------------------------------------

How far the flat between two grooves may sit off
the true curve.

Grooving turns a curve into a chain of short
straight facets. The gap at the middle of a facet
is about s squared / (8 x R), so demanding a
smaller gap forces the grooves closer together.

This is the setting that controls how round the
finished letter looks. Lower it when you can see
flats on the curves; 0.1 mm or less for close
viewing, 0.3 mm is fine for signs read from
across a street.

----------------------------------------------
Now      : 0,15 mm
Default  : 0.15 mm
Range    : 0.01 and up
```

> **Fare üzerine gelince çıkan ipucu (tooltip) yok.** Bunun için UserForm
> gerekiyor; `.frm` yanında ikili bir `.frx` ister ve o ikili bu depoda
> güvenle üretilemiyor. Açıklamalar bunun yerine `?N` ile tam metin
> olarak verilir.

---

## 7. Profiller

Ayar sayfasında `P`, ya da doğrudan `ac2fProfiles`.

Bir profil, **22 ayarın tamamının** bir ad altında saklanmasıdır. Farklı
malzeme ve kalınlıklarla çalışıyorsanız her biri için bir profil tutun.

```
PROFILES

   1mm galvaniz
   2mm alu kutu harf   <- loaded
   3mm alu vitrin

S name    save the current settings under that name
L name    load a profile
D name    delete a profile
M         apply a material preset
Enter     back
```

| Yazın | Ne olur |
|---|---|
| `S 3mm alu` | Şu anki ayarları bu adla kaydeder |
| `S` | Adı ayrıca sorar |
| `L 3mm alu` | Profili yükler |
| `L` | Profil listesinden seçtirir |
| `D 3mm alu` | Profili siler |
| `M` | Malzeme ön ayarı uygular (esneklik, derz derinliği, K) |

Profiller kayıt defterinde saklanır:

```
HKCU\Software\VB and VBA Program Settings\ac2fPack\Profiles
```

> Profil adı serbesttir; `L` veya `S` ile başlayan adlar (`Letters3mm`
> gibi) sorun çıkarmaz — komut sayılması için harften sonra boşluk
> gerekir.

Bir ayar ileride paketten çıkarılırsa eski profiller yine yüklenir;
tanınmayan anahtarlar sessizce atlanır.

---

## 8. Profille çalıştırma ve geçici değişiklik

Ana menü `6` (`ac2fBoxLetterStripProfile`). Üç adım:

1. **Profil seç** — `L 2mm alu`, ya da boş Enter ile mevcut ayarlarla devam.
2. **Bu seferliğine değiştir** — ayar sayfası açılır, `9=2.5` gibi
   yazarsınız. Değiştirilen satırlar `*` ile işaretlenir.
3. **Enter** — şerit çizilir.

```
RUN SETTINGS - changes apply to this run only

-- BOX LETTER --
 9 Thickness            2.50 mm *
11 Flexibility          0.80
...

N=value   change        ?N   explain
Enter     run
* = changed for this run only
```

Geçici değerler **kayıtlı ayarlarınıza yazılmaz** ve çizim biter bitmez
silinir. Aynı profille tek bir kalınlığı deneyip görmek için budur.

Kalıcı olmasını istiyorsanız ayar sayfasından (`9`) değiştirin, sonra
`P` → `S <ad>` ile profile kaydedin.

---

## Bilinmesi gerekenler

- Ölçüm **her zaman milimetre** üzerinden yapılır; belgenizin birimi ne
  olursa olsun sonuç doğrudur ve işlem sonunda birim geri alınır.
- Kontur uzunluğu, çizginin **kalınlığını dikkate almaz** — yolun kendi
  uzunluğudur.
- Çok sayıda nesne seçildiğinde ölçüm birkaç saniye sürebilir; ekran
  tazeleme bu sırada kapatılır.
- Nesne kırılımı listesinde en fazla 15 satır gösterilir; toplam her zaman
  seçimin tamamını kapsar.
