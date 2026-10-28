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
  9  -  Centerline (reduce to a single line)
 10  -  Nesting (true shape, sheets window)
 11  -  Nesting report (nothing moves)
 12  -  Bending order: apply it to the objects
 13  -  Bending order: number them on the page
 14  -  Bending order report (nothing changes)
 15  -  Send to plotter
 16  -  Plotter test shape (40 x 80 L)
 17  -  Write HPGL to a file
 18  -  Normalise an existing .plt and send
 19  -  Settings and profiles
 20  -  About
 21  -  Customer job: export, folders, price  (J)
 22  -  Nesting again, last settings
 23  -  Start points   24 - again   25 - markers
 26  -  Mark selection as offcut sheets
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

- **Bu modül segmentleri dairesel yay kabul eder** (yalnız uzunluk ve
  kiriş okunur). Gerçek bezier yaylarda yarıçap hatası %0,05 mertebesindedir
  — ihmal edilebilir. Ancak tek bir segment içinde dönüm noktası varsa
  (S kıvrımı) model onu tek yönlü yay sanar ve o segmentte derzi seyrek
  koyar. Böyle bir yerde düğüm ekleyip segmenti ikiye bölmek sorunu çözer.
  (Plotter, bükme sırası ve orta hat bu modeli **kullanmaz**; onlar eğrinin
  gerçek noktalarını okur.)
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

## 6. Orta hat — tek çizgiye düşürme

Pleksi üzerine 6 mm bıçakla neon LED şerit kanalı açarken, dolu
tasarımın tam ortasından geçen tek bir çizgiye ihtiyaç duyarsınız.
Bu araç onu üretir.

Tasarımı seçin, `ac2fCenterline` çalıştırın, `J` ya da `S` yazın.

### İki mod

| Mod | Ne yapar | Ne zaman |
|---|---|---|
| `J` **birleşik** | Tüm seçim **tek bölge** sayılır. Birbirine değen ya da bindiren harfler tek sürekli çizgi olur. | El yazısı, bitişik yazılmış kelimeler |
| `S` **ayrı ayrı** | Her nesne kendi başına düşürülür, her biri kendi grubunu alır. | Ayrık harfler, ayrı ayrı işlenecek parçalar |

Doğrudan çağırmak için: `ac2fCenterlineJoined`, `ac2fCenterlineSeparate`.

> Birleşik modda harflerin gerçekten **değmesi** gerekir. Aralarında
> boşluk varsa çizgiler de ayrı çıkar — olmayan bağ uydurulmaz.

### Sonuç

Çizgiler **kırmızı** çizilir ve tek grupta toplanır
(`ac2f centerline`). Her kol tek bir polyline'dır, CAM'de ALONG
verilmeye hazır.

```
RESULT
   Mode                : joined - one region
   Centre lines        : 1
   Paths               : 3
   Total length        : 1.284,60 mm   |   128,46 cm   |   1,285 m
   Mean stroke width   : 22,10 mm
```

**Mean stroke width** ölçülen ortalama harf kalınlığıdır (alan ÷ çizgi
boyu). Bıçak ve şerit seçiminde işinize yarar; ayrıca budama ve uç
uzatma bu değere göre ölçeklenir.

### Ayarlar (26-30)

| Ayar | Varsayılan | Ne yapar |
|---|---|---|
| Resolution | 1 mm | Izgara hücresi. Küçültmek doğruluğu artırır ama **süreyi hızla** büyütür |
| Simplify tol | 0,3 mm | Düğüm azaltma toleransı |
| Smoothing | 2 | Zikzak giderme geçişi |
| Min branch xW | 1 | Bundan kısa dallar atılır (kalınlık katı) |
| Extend ends | 1 | Uçları harfin kenarına kadar uzat |

`?26` … `?30` ile her birinin ayrıntılı açıklaması açılır.

### Doğruluk

Orta hattı bilinen şekillerde ölçüldü:

| Test | Sonuç | Gerçek | Hata |
|---|---|---|---|
| Dalgalı şerit | 470,8 mm | 471,0 mm | %0,0 |
| T kavşağı (3 dal) | 279,3 mm | 280,0 mm | %0,2 |
| İki ayrı harf | 161,0 mm | 160,0 mm | %0,6 |

Uç konumu hatası 1,8 mm.

### Bilinmesi gerekenler

- **Süre çözünürlükle hızla artar.** 1 mm'de tipik bir kelime birkaç
  saniye. 0,5 mm'ye inmek ızgarayı dörde katlar ve inceltme turunu
  yaklaşık ikiye — dakikalara çıkabilir. Önce 1 mm ile bakın.
- **Çok büyük işler reddedilir.** 4 milyon hücreyi aşarsa makro
  çözünürlüğü artırmanızı söyler.
- **Uçlar yarım kalınlık geri yenir**, bu inceltmenin doğasıdır;
  `Extend ends` bunu telafi eder. Kapatırsanız çizgi harfin ucuna
  varmaz ve şeridin son parçası yanmaz.
- **Tek segment içinde dönüm noktası** (S kıvrımı) varsa kontur açma
  modeli onu tek yönlü yay sanar; o bölgede hafif sapma olur. O noktaya
  düğüm eklemek çözer.
- Bitmap ve OLE nesneleri atlanır. Yazıyı önce eğriye çevirmeniz
  gerekmez; makro gerekiyorsa geçici kopya üzerinden çevirir.

---

## 7. Nesting

Parçaları seçin, `ac2fNest` (menü `10`) çalıştırın. Akış:

```
ayar penceresi  ->  Preview  ->  önizleme penceresi  ->  Apply
                                     |
                                     +-> Change settings: ayar penceresine geri
                                     +-> Cancel: hiçbir şey değişmez
```

1. **Ayar penceresi**: solda **tabaka listesi**, sağda iki sekmede
   **ayarlar** (*Main* sık kullanılanlar, *More* geri kalanı).
   - **Preview** (Enter): yerleşimi hesaplar ve önizlemeyi açar.
   - **Nest now**: önizlemesiz, hemen yerleştirir.
   - **Cancel**: çıkar.
2. **Önizleme penceresi**: kullanılan her tabaka ve yerleşen her parça
   gerçek şekliyle çizilir (delikler boş, çizili artık kendi şekliyle,
   kenar payı içi açık renk). Sağdaki listeden **tek bir tabaka** seçince
   o tabaka büyütülür; altında rapor (parça sayısı, doluluk, kalan boy).
   Bu sırada belgede **hiçbir şey taşınmamıştır**.
   - **Apply**: parçaları çizildiği gibi yerleştirir (`Ctrl+Z` geri alır).
   - **Change settings**: belgeyi olduğu gibi bırakır, ayar penceresine
     döner; değiştirip yeniden *Preview* deyin.
   - **Cancel**: hiçbir şey değişmez.

`ac2fNestReport` (menü `11`) aynı pencereleri açar. `ac2fNestAgain` (menü
`22`) pencere açmadan son ayarlarla hemen yerleştirir.

Parçalar **gerçek konturlarıyla** yerleşir, sınırlayıcı kutularıyla
değil: içbükey parçalar birbirine geçer (L, T, harfler), istenirse küçük
parçalar büyüklerin deliklerine girer, çizili bir artık tabaka olarak
kullanılabilir.

### Tabaka listesi

| Sütun | Anlamı |
|---|---|
| Use | İşaretli tabakalar kullanılır; işaretini kaldırdığınız dışarıda kalır |
| Sheet | Ad (raporda görünür) |
| Width / Height | mm |
| Qty | Adet. **`0` = gerektiği kadar** |

- **Add a sheet**: sık kullanılan ölçüler (3050 x 2050, 3050 x 1500,
  2440 x 1220 ... ve sayfa ölçüsü) listeden seçilip `Add` ile eklenir;
  ölçü sonra listede değiştirilebilir.
- Tabakalar **yukarıdan aşağı** sırayla kullanılır (`Up` / `Down`). Artıkları
  üste koyarsanız önce onlar dolar.
- `Qty for all` + `Apply` bütün liste tabakalarına aynı adedi yazar.
- Liste kaydedilir; bir sonraki işte aynen gelir.
- Müşteri işinde stoktan seçilen tabaka (iş menüsü) listede yoksa en üste
  eklenir. Liste boşsa sayfa ölçüsü kullanılır.
- **Çizili artık**: adı `offcut`, `sheet`, `artik`, `plaka`, `tabaka` ya
  da `levha` ile başlayan (veya böyle adlı bir katmanda duran) kapalı bir
  şekli parçalarla birlikte seçin. Mavi satır olarak listede görünür,
  **olduğu yerde** ve **kendi şeklinde** kullanılır (içindeki delikler
  dahil). Menü `26` seçili şekilleri `offcut 1`, `offcut 2` ... diye
  adlandırır.
- Önceki çalıştırmanın çizdiği tabaka çerçeveleri (`ac2f sheet N`) ne
  parça ne tabaka sayılır; her şeyi seçip yeniden nest edebilirsiniz.

### Ayarlar

Penceredeki her ayarın üzerine gelince açıklaması çıkar. Aynı ayarlar ayar
sayfasında `#5` (NEST) altındadır.

| Ayar | Varsayılan | Ne yapar |
|---|---|---|
| Part gap | 3 mm | Parçaların **gerçek konturları** arasındaki en az boşluk |
| Margin L / R / T / B | 10 mm | Tabakanın her kenarından ayrı pay |
| Offcut edge | 5 mm | Çizili artığın kenarından pay |
| Rotation | 90° | `0` olduğu gibi, `180` yalnız yarım tur, `90`, `45`, `30`, `15` |
| Parts may be mirrored | kapalı | Ayna görüntüsü de denenir (baskılı/filmli malzemede kapalı tutun) |
| Copies of each part | 1 | Her parça bu kadar kez yerleşir; kopyalar yer bulursa oluşur |
| Shapes inside a shape move with it | açık | Bir şeklin içindeki şekil onunla tek parça olur (gruplamaya gerek yok) |
| Small parts may go into holes | kapalı | İç konturlar kesilip düşüyorsa (harf içleri, çerçeveler) küçük parçalar oraya girer |
| Start corner | sol alt | Yerleşimin başladığı köşe (makinenin sıfırına göre seçin) |
| Fill direction | satır | Satır: kullanılan yükseklik kısa kalır; sütun: kullanılan genişlik |
| Placement | otomatik | En alçak üst kenar / sol-alt; otomatik ikisini de dener |
| Precision | 1 mm | Izgara kare boyu. İnce = sıkı ama yavaş. Boşluk her durumda korunur |
| Optimise for | 10 s | Bu süre boyunca başka sıralar denenir, en iyisi kalır. `0` = tek hızlı geçiş |
| Group the result | each part | Yok / her parça / her tabaka / hepsi |
| Draw sheet outlines | açık | Kullanılan her tabakaya gri çerçeve |
| Set start points afterwards | kapalı | Yerleşen parçalara başlangıç noktası ayarlarını uygular (bölüm 17) |
| Space between sheets | 20 mm | Sayfadaki tabakaların arası (yalnız görünüm) |

**Hazır ayarlar**: alttaki `Preset` kutusuna bir ad yazıp `Save` — o anki
bütün ayarlar o adla saklanır (`Laser 3mm`, `CNC 6mm` ...). `Load` geri
yükler, `Delete` siler.

**İç / dış**: `Shapes inside a shape move with it` açıkken, ayrı ayrı
çizilmiş bir "O"nun iki halkası, bir çerçeve ile içindeki yazı ya da bir
panel ile delikleri tek parça gibi taşınır. `Small parts may go into
holes` iç konturun **kesilip düştüğü** durumlar içindir; iç şekiller kazıma
veya baskıysa kapalı tutun, çünkü oradaki alan hâlâ malzemedir.

### Nasıl yerleştirir

1. Her parça kare ızgaraya dökülür. Bir kare, parçanın yarım boşluk kadar
   büyütülmüş hâlinin herhangi bir noktası oraya düşebiliyorsa dolu
   sayılır, yani karesi çakışmayan iki parça arasında **en az bir boşluk**
   vardır.
2. İzin verilen her dönüş (ve ayna) CorelDRAW'da gerçekten yapılıp kontur
   yeniden okunur; aynı parçanın kopyaları aynı ızgarayı paylaşır.
3. Her parça için en alçak (sonra en soldaki) boş yer aranır. Delikler ve
   girintiler de bu aramada bulunur.
4. Birkaç sabit sıra (büyükten küçüğe, uzun kenar, yükseklik, genişlik,
   döndürmesiz) ve iki strateji denenir, sonra süre dolana kadar en iyi
   sıra rastgele değiştirilerek tekrar denenir. **En az parça dışarıda,
   sonra en az malzeme** olan yerleşim kalır (son tabaka yalnız en üstteki
   parçaya kadar sayılır).

### Sonuç

```
NESTED
   Parts             : 46 of 48 placed
   Did not fit       : 2  (left where they were)
   Sheets used       : 3
      1  3050 x 2050                31 parts   fill 78.4 %
      2  3050 x 2050                12 parts   fill 64.0 %
      3  Offcut 1: offcut 1 (850... 3 parts   fill 55.1 %   used 1210 mm of the height
   Fill (all sheets) : 72.3 %  (the last sheet up to its highest part)
   Layouts tried     : 57 in 10.1 s; best: lowest top edge

SETTINGS USED
   Gap 4.00 mm, margins L R T B 10 10 10 10 mm
   Rotation 90 deg, mirror off, precision 1.00 mm
   Inner shapes with outer on, holes used off, from the bottom left in rows
```

**Fill** gerçek parça alanının kullanılabilir alana oranıdır. **used**
son tabakada gerçekten harcanan boydur; artığın ne kadar kaldığını
söyler.

### Bilinmesi gerekenler

- Sığmayan parça **yerinde bırakılır** ve raporda sayılır.
- `Ctrl+Z` her şeyi geri alır (tek komut grubu). Rapor modu belgeyi
  hiç değiştirmez.
- Çok büyük bir iş belleğe sığsın diye hassasiyet kendiliğinden
  kalınlaşabilir; rapor bunu yazar.
- Konturlar eğri boyunca noktalarla okunur. Bu okuma olmayan çok eski bir
  CorelDRAW'da her eğri parçası, o boydaki **her** eğriyi içine alan bir
  kutuyla kaplanır: parçalar biraz gevşek yerleşir ama asla çakışmaz
  (rapor söyler).
- Pencereler Windows'la gelen PowerShell 5.1 ile açılır (iş menüsündeki
  liste penceresi gibi; Windows 10/11'de hazırdır). İş menüsündeki *List
  window* ayarı `0` ise ya da PowerShell engelliyse kısa bir metin sorusu
  çıkar: `Enter` nest, `R` rapor, `S` ayarlar, `T 3050x2050x0;
  2440x1220x2` tabakaları yazar. Önizleme açılamazsa rapor mesaj
  kutusunda gösterilir.
- **Kilitli** nesneler (Arrange > Lock) taşınmaz: yerinde kalır, raporda
  `Locked` satırında sayılır.
- Sınırlar: 2000 parça, 200 satır tabaka, aynı anda 60 tabaka.

---

## 8. Plotter'a gönderme

Seçimi seçin, `ac2fPlotSend` çalıştırın. Tek soru gelir:

```
Plotter address as host:port

   192.168.1.100:9100    raw socket, the usual case
   192.168.1.100:23      a telnet port
```

Onaylayın, iş gider. **Export etmeniz ve dosya düzeltmeniz gerekmez.**

### Yön: nasıl görüyorsanız öyle

**Hiçbir şey döndürülmez.** Belge X'i plotter X'i, belge Y'si plotter
Y'si olur. Arada döndürecek bir export süzgeci yok — iş ekranda durduğu
gibi kesilir.

Yine de dönük çıkıyorsa sebep dosya değil **makinedir**: çoğu kesicide X
ekseni malzeme besleme yönünde uzar, bu yüzden geniş bir iş ruloya enine
düşer.

| `Rotate` (ayar 61) | Sonuç |
|---|---|
| `0` *(varsayılan)* | Ekranda gördüğünüz gibi |
| `90` | Saat yönünün tersine çeyrek tur |
| `180` | Ters çevrilir |
| `270` | Saat yönüne çeyrek tur |

90 ve 270'te en–boy takas olur (200×300 → 300×200). Kenar payı
döndürmeden **sonra** uygulanır, dolayısıyla iş her durumda köşeye
oturur. 90'ın katı olmayan bir değer en yakın çeyreğe yuvarlanır.

### Dönük, ters ya da aynalı geliyorsa: yön sihirbazı

`ac2fPlotCalibrate` (ana menü `27`) `Rotate` ve `Mirror` değerini **tek
kesimle** bulur. Test L'si hiçbir döndürme olmadan gider, çıkan parçaya
(makinenin önü altta) bakıp sekiz resimden numarasını yazarsınız; ayarlar
buna göre tersine kurulur, L yeniden gider ve "doğru" demenizle biter. Yanlış
resmi seçseniz bile doğrulama kesimi ayarı bir kez daha düzeltir (en çok
4 kesim). Yarıda bırakırsanız önceki ayarınız geri yüklenir.

Gönderim raporunda iki yeni bölüm vardır:

- **Design (CorelDRAW)** / **Check**: dosyadaki boyut, CorelDRAW'ın aynı
  seçim için bildirdiği boyutla karşılaştırılır. Fark çıkarsa dış çizgi
  kalınlığı, bir efekt ya da sayfa dışında kalan bir parça aranır.
- **Curve tolerance**: dosya, çizimden bu kadardan fazla sapmaz. Eğriler
  CorelDRAW'dan gerçek noktalarıyla okunur (bkz. README › Eğri örnekleme);
  eskiden dairesel yay *tahmin edilir* ve S kıvrımlı ya da el yazısı
  şekiller milimetrelerce bozulurdu.

### Ara dosya var mı

**CorelDRAW'ın PLT export'u hiç kullanılmaz.** Modülde tek bir `Export`
çağrısı yok; HPGL doğrudan geometriden yazılır. Önce dışa aktarıp sonra
göndermek diye bir adım yoktur.

Gönderirken baytlar, gönderici sürece devretmek için geçici bir dosyadan
geçer. **Gönderim başarılı olunca o dosya silinir** — raporda
`File: none kept (streamed)` yazar. Başarısız olursa dosya bilerek
bırakılır ki elle gönderebilesiniz.

### Kesim sırası

Nesneler **Object Manager'daki sırayla**, en üstteki önce gider; yolda
hiçbir şey yeniden sıralanmaz. Seçim (`ActiveSelectionRange`) bu sırayı
garanti etmediği için her seçili nesne sayfa yığınında kimliğiyle
(`StaticID`) bulunur. Bir grubun içinden tek tek seçilmiş nesneler
sayfa yığınında yoktur; geri kalanların ardından, seçim sırasıyla
gider ve rapor kaç tane olduğunu söyler.

| `Send order` (ayar 66) | Sıra |
|---|---|
| `0` *(varsayılan)* | Object Manager, üstten başlayarak |
| `1` | Object Manager, alttan başlayarak |
| `2` | Bükme sırası, o anda hesaplanır (bkz. **9. Bükme sırası**) |

Bükme sırasıyla kesmenin olağan yolu: önce ana menü `12` (sırayı
nesnelere uygular, Object Manager'ı yeniden dizer), sonra gönder.

### İş bitince kafa nerede bekler

| `Park` (ayar 67) | Sonuç |
|---|---|
| `1` *(varsayılan)* | Bıçak kalkar, kafa **işin en üst noktasına** gider |
| `0` | Orijine döner (`PU0,0;`) |
| `2` | Son kesimin bittiği yerde kalır |

"En üst", işin **ekranda göründüğü hâliyle** sol üst köşesidir. Nokta
işle aynı `Rotate` ve `Mirror`'dan geçer, yani makine işi hangi yöne
çevirirse çevirsin işin tepesinde kalır (sekiz dönüş/aynalama
birleşiminin hepsinde doğrulandı; kenar payını da hiç etkilemez). Bazı
makineler `SP0`'dan sonra kendi orijinine döner; `Preamble = 0` `SP0`'ı
göndermez.

### Neden düzeltme adımı yok

Export edilmiş bir `.plt`'de geometri, sayfanın verdiği koordinatlarda
kalır — sık sık negatif. Plotter bunu kabul etmez, dosyanın
ayrıştırılıp kaydırılması gerekir.

Burada kaydırma **HPGL yazılırken, geometrinin kendisinden** uygulanır:
en küçük X ve en küçük Y, kenar payına oturtulur. Ayrıştırılacak metin
yok, düzeltilecek dosya yok.

```
SENT
   Plotter             : 192.168.1.100:9100
   Bytes               : 48.312
   Paths               : 27
   Points              : 3.914

PLACEMENT
   Size                : 412,60 x 280,15 mm
   Margin              : 5,00 mm  (200 units)
   X range             : 200 .. 16704
   Y range             : 200 .. 11406
   Curve tolerance     : 0,05 mm

Coordinates are already shifted; nothing to fix.
```

HPGL birimi milimetrede **40**'tır (inçte 1016), yani 5 mm = 200 birim.

### Bağlantı nasıl kuruluyor

VBA'nın kendi socket'i yok. Baytlar kısa bir PowerShell betiği üzerinden
`System.Net.Sockets.TcpClient` ile çıkar. PowerShell desteklenen her
Windows'ta bulunur, bu yüzden `Declare` bildirimi, 32/64 bit uyumu ya da
OCX kaydı gerekmez.

Betik bir günlük yazar, makro onu geri okur. Bağlantı kurulamazsa
gerçek sebebi görürsünüz:

```
SEND FAILED
   ...
REASON
   No answer from 192.168.1.100:9100 within 5000 ms

The file is written, so you can send it by hand.
```

> Bu **ham TCP akışıdır** — 9100 ya da bir telnet portunda dinleyen
> plotter'ların beklediği şey. Telnet protokolü anlaşması (IAC kaçırma)
> yapılmaz; plotter'lar anlaşma istemediği için bu doğru davranıştır.

### Dönük geliyorsa

Paket işi **birebir** gönderir: belge X'i plotter X'i, belge Y'si
plotter Y'si. Arada döndüren bir şey yok, `Rotate` varsayılanı 0.

Yine de dönük çıkıyorsa sebep **makinedir**: çoğu kesicide X ekseni
malzeme besleme yönünde uzar, bu yüzden geniş bir iş ruloya enine
düşer. Telafisi `Rotate` ayarıdır (42) — **270** ya da **90**. Hangisi
olduğunu test L'si bir seferde söyler.

Döndürme **aynalanmış** bir ekseni düzeltemez; onun için ayrı bir
`Mirror` ayarı (43) var ve döndürmeden sonra uygulanır.

### Gezinip kesmiyorsa

Bu belirti nettir: `PD` koordinatları hareket olarak çalışıyor ama bıçak
inmiyor. `ac2fPlotTest` (ana menü 16) bunu bir dakikada teşhis eder — gerçek işin
kullandığı **aynı** yazıcı, gönderici ve yönlendirme adımından
40 × 80 mm'lik bir **L** gönderir:

```
   |
   |
   |___     uzun bacak SOLDA, ayak SAĞA, 80 mm uzun yön
```

| Sonuç | Anlamı | Çözüm |
|---|---|---|
| **Doğru şekil, temiz kesti** | Her şey yolunda | — |
| **Yan yatmış** (80 mm enine) | Makine eksenleri takas ediyor | `Rotate = 270`; ters olursa `90` |
| **Ayna görüntüsü** (ayak solda) | Makine bir ekseni aynalıyor | `Mirror = 1`; baş aşağıysa `2` |
| **Doğru şekil ama sadece gezdi** | Makine hiç kesmiyor | Bıçak kuvveti, derinlik, takım; sonra `Preamble = 2` |
| **Hiç kıpırdamadı** | Ulaşmadı | Adres, port, kablo |

Test, `Rotate` ve `Mirror` ayarlarını uygular — yani testte doğru çıkan
değer gerçek işte de doğrudur. Kare yerine L olmasının sebebi: kare
hangi yöne çevrilse aynı görünür, aynalansa yine karedir.

**`PD pairs` neden önemli:** eskiden komut başına 40 koordinat çifti
yazılıyordu — 180 noktalı bir daire için **402 karakterlik** satırlar.
Giriş tamponu veya ayrıştırıcısı bunu kaldıramayan bir kesici satırın
sonunu düşürür; sonraki `PD` bozuk komutun parçası olarak yutulur,
kaleme inme emri ulaşmaz ve makine şekilleri bıçak yukarıda dolaşır.
`1` ile her komut 12 karakter olur, dosya yalnız %38 büyür.

**`Preamble` neden önemli:** `IN;` cihazı açılış varsayılanlarına
döndürür. Pek çok kesicide bu, panelden ayarladığınız **bıçak kuvvetini
ve takımı da siler**. `2` yaparsanız `IN;` gönderilmez, panel ayarları
olduğu gibi kalır.

**`Send delay`:** 0'da bırakın. Yalnız tamponu küçük, akış denetimi
olmayan ethernet–seri dönüştürücüler içindir; bayt düşerse aynı belirti
görülür. İlk iki ayar işe yaramadıysa 20 deneyin.

### Diğer iki makro

| Makro | Ne yapar |
|---|---|
| `ac2fPlotSave` | HPGL'i dosyaya yazar, göndermez. Çıktıyı incelemek için. |
| `ac2fPlotFixSend` | Var olan bir `.plt`'yi normalize eder, sorar, isterseniz gönderir. |

`ac2fPlotFixSend` eski betiğinizin işini yapar. Bir farkla: **çizim alanı
yalnız çizim komutlarından** hesaplanır — her `PD`, ve ardından `PD`
gelen `PU`'lar.

Sondaki `PU0,0;` kalem park komutudur, çizim değildir. Hesaba katılırsa:

| Dosya | Tüm PU/PD | Yalnız çizim |
|---|---|---|
| Çizim negatif, park var | doğru | doğru |
| **Çizim pozitif, park var** | **yanlış** | doğru |
| Çizim pozitif, park yok | doğru | doğru |

Ortadaki satır önemli: çizim tamamen pozitifse min 0,0 çıkar,
normalizasyon sabit bir kaydırmaya döner, iş kenara oturmaz ve malzeme
boşa gider.

### Ayarlar (59-67)

| # | Ayar | Varsayılan | Ne yapar |
|---|---|---|---|
| 40 | Margin | 5 mm | İşin plotter orijininden uzaklığı |
| 41 | Curve tol | 0,05 mm | Eğrilerin kaç düz adıma bölüneceği (bükme de bunu kullanır) |
| 42 | Rotate | 0 | Çeyrek tur. **0 = gördüğünüz gibi** |
| 43 | Mirror | 0 | `1` sol-sağ, `2` alt-üst, `3` ikisi. Ayna görüntüsü gelirse |
| 44 | PD pairs | 1 | Komut başına koordinat çifti. Kesmiyorsa buradan başlayın |
| 45 | Preamble | 1 | `2` = `IN;` gönderme (panel ayarları korunur) |
| 46 | Send delay | 0 | KB başına duraklama. Seri dönüştürücü için |
| 47 | Send order | 0 | `0` Object Manager üstten, `1` alttan, `2` bükme sırası |
| 48 | Park | 1 | İş bitince kafa: `1` işin tepesi, `0` orijin, `2` olduğu yer |

Bükme sırası ayarları **49-55**, `BEND` grubunda — bkz. 9. bölüm.

Plotter adresi **ayar sayfasında değil** — gönderirken sorulur ve
hatırlanır. Bir makine adresi, geometri parametresi değil.

### Bilinmesi gerekenler

- Tek kalem kullanılır (`SP1`). Renk/katman başına kalem ataması yok.
- Sayfa sonu (`PG;`) gönderilmez — beklemediğiniz bir malzeme ilerlemesi
  olmaz. Gerekiyorsa plotter panelinden verin.
- Yol sırası **Object Manager sırasıdır** (ayar 66). Kalem yolunu
  kısaltan bir optimizasyon yapılmaz.
- Bitmap ve yol taşımayan nesneler atlanır.
- Gönderim başarısız olsa bile **dosya yazılmıştır**; raporda yolu
  vardır, elle gönderebilirsiniz.

---

## 9. Bükme sırası

Kutu harf bükerken sıra, makineyi kaç kez kurduğunuzu belirler. Olduğu
gibi alınırsa `NOBOBOXB` sekiz kurulumdur; sıraya konunca dörttür.

| Ana menü | Makro | Ne yapar |
|---|---|---|
| `12` | `ac2fBendOrder` | Sırayı **nesnelere uygular** (istif sırası + isim) |
| `13` | `ac2fBendPreview` | Sayfada numaralandırır, başka hiçbir şeye dokunmaz |
| `14` | `ac2fBendReport` | Sırayı listeler, hiçbir şeyi değiştirmez |

**Plotter'a göndermek:** `ac2fPlotSend` nesneleri **Object Manager'daki
sırayla**, en üstteki önce gönderir ve yolda hiçbir şeyi yeniden
sıralamaz. Bükme sırasıyla kesmek için önce `12` ile sırayı uygulayın —
Object Manager yeniden dizilir — sonra gönderin. Ya da `Send order`
(ayar 66) = `2` ile plotter sırayı kendisi hesaplasın.

**Algoritma her çalıştırmada sorulur.** `12`, `13` ve `14` önce 1/2/3
seçimini ister; son seçtiğiniz varsayılan gelir ve hatırlanır.

### Nasıl çalışıyor

**1. İçine alma.** Her kapalı kontur ötekilere karşı sınanır: ilk
noktası içine düşüyor mu, ve öteki daha büyük mü? Hiçbir şeyin içinde
olmayan kontur bir **dış**, bir dışın içindeki **göbek** (A'nın ortası),
göbeğin içindeki yine dıştır. Bir dış ve kapsadığı her şey **bir
harftir**.

**2. Çalışma sırası.** Harfler soldan sağa dizilir. X'te üst üste binen
harfler bir **sütun** sayılır ve sütun içinde yukarıdan aşağı gidilir —
"soldan sağa, önce üst" bir sayfada satır satır duran parçalar için
budur.

**3. Önce aynılar.** Bir harfe gelindiğinde ondan sonraki **aynı
şekilli** bütün harfler hemen arkasına çekilir.

### Dört algoritma — `Order` (ayar 68)

| | Sıra |
|---|---|
| `0` | Çizimdeki hâliyle |
| `1` *(varsayılan)* | Soldan sağa, önce üst |
| `2` | Aynısı, ama **önce düz harfler**, sonra eğriler |
| `3` | 2 ile aynı, ama **göbekler en sona** |

`2`'de düz–eğri ayrımı şudur: **içinde hiç eğri olmayan** harfler önce
(N, A, K, X, E, M, V, W, Z, L, T, I, H, F), **içinde eğri olan** harfler
sonra (O, B, P, S, C, G, D, R, U, Q, J). Rulolu bir bükme makinesinde
düz işin tamamı, rulolar hiç devreye girmeden biter.

`NOBOBOXB` için sonuç (doğrulandı):

```
1 ->  N  O O. O O. O O.  B B.1 B.2  B B.1 B.2  B B.1 B.2  X
2 ->  N  X  O O. O O. O O.  B B.1 B.2  B B.1 B.2  B B.1 B.2
3 ->  N  X  O O O  B B B  |  O. O. O.  B.1 B.2  B.1 B.2  B.1 B.2
```

`.` ile bitenler göbek. Kurulum sayısı: **S1 = N, S2 = O ×3,
S3 = B ×3, S4 = X.**

### `ac2fBendOrder` nesnelere ne yapar

- **İstif sırası.** İlk bükülecek nesne en üste alınır (`Stack order`,
  ayar 73 ile alta da alınabilir). Çoğu program dışarı yazarken istif
  sırasını kullanır, ama hangi uçtan başladıkları değişir; iş ters
  çıkarsa bu ayarı çevirin.
- **İsim.** Nesne `[007 S03] eski adı` olur: **yedinci bükülecek, üçüncü
  kurulum.** Aynı `S` numarasını taşıyanlar aynı şekildir ve peş peşe
  bükülür. İsimler Object Manager'da görünür ve istif sırasının
  dayanmadığı pek çok işlemden sağ çıkar. Tekrar çalıştırınca eski etiket
  sökülür, üst üste binmez. `Name objects` (55) ile kapatılabilir.
- **Hiçbir nesne sayfada oynatılmaz.** Tek `Undo` hepsini geri alır.

Bir sınır: `Order 3` göbekleri sona atar, ama bir göbek **dışıyla aynı
nesnenin içindeyse** istif sırasıyla ayrılamaz (istif nesneleri sıralar,
konturları değil). O durumda `ac2fPlotSend` kullanın — orada konturlar
tek tek yazıldığı için ayrım tam olarak uygulanır.

### Aynı sayılma ölçütü

İki harf; kontur sayısı, uzunluk, kapsadığı alan ve kutu ölçüsü
`Match tol` (53) içinde uyuşuyorsa aynıdır. En ve boy yer değişebilir,
yani çeyrek tur döndürülmüş kopya da aynı sayılır. **Aynalanmış kopya
sayılmaz** — ters yöne bükülür, kendi kurulumunu ister; işaretli alanın
işareti bunu yakalar.

**Gruplama sırayı ezer.** `Group same` (50) açıkken soldan sağa sıra
bozulabilir; zaten amacı budur. Sayfadaki sıra kurulum sayısından daha
önemliyse `0` yapın.

### Ayarlar (68-74)

| # | Ayar | Varsayılan | Ne işe yarar |
|---|---|---|---|
| 49 | `Order` | `1` | Yukarıdaki dört algoritma |
| 50 | `Group same` | `1` | Aynı harfleri peş peşe bük |
| 51 | `Column overlap` | 5 mm | Bir sütun sayılmak için gereken X bindirmesi |
| 52 | `Curve limit` | 0,5 mm | Bir harfin "eğri" sayılması için gereken şişkinlik |
| 53 | `Match tol` | 0,5 mm | İki harfin aynı sayılması için gereken yakınlık |
| 54 | `Stack order` | `0` | İlk harf üstte mi altta mı |
| 55 | `Name objects` | `1` | Sırayı nesne adlarına da yaz |

Eğri toleransı ayrı değil: konturları düz adımlara bölen tek ayar
`Curve tol` (41), plotter grubunda.

Sıra **ekranda gördüğünüz çizim üzerinden**, `Rotate` ve `Mirror`'dan
önce hesaplanır; yani "sol" her zaman ekranınızın solu.

---

## 10. Ayarlar — tek sayfa, gruplara bölünmüş

Ana menüden `19` (`ac2fSettings`). Paketin **bütün** ayarları tek listede:

```
SETTINGS  [Aluminium 2mm]

[NEST]  page 5/6
31 Marg left           10 32 Marg right          10
33 Marg top            10 34 Marg bottom         10
35 Part gap             3 36 Rot step            90
37 Sheet W           1220 38 Sheet H           2440
39 Sheet gap           20

1 LED  2 BOX  3 ACP  4 CENTER  5 NEST  6 PLOT

N=val  ?N=help  #n=page  P=profiles  R=reset  Enter=close
```

55 ayar tek `InputBox` istemine (≈1024 karakter) hiçbir düzende
sığmadığı için sayfa **gruba göre bölündü**. Ama **ayar numaraları
geneldir**: `36=180` yazmak hangi sayfada olursanız olun çalışır. Yani
hâlâ tek menü — soru zinciri değil.

### Komutlar

| Yazın | Ne olur |
|---|---|
| `11=0.8` | 11 numaralı ayarı 0,8 yapar |
| `11=0.8 16=0.1` | Birden çok ayarı tek seferde değiştirir |
| `11=0.8; 16=0.1` | Noktalı virgül de olur |
| `?11` | 11 numaralı ayarın **ayrıntılı açıklaması** |
| `Flexibility=0.9` | Numara yerine ayar adı |
| `?Flexibility` | Ada göre açıklama |
| `#5` | 5 numaralı gruba geç |
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

## 11. Profiller

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

## 12. Profille çalıştırma ve geçici değişiklik

Ana menü `6` (`ac2fBoxLetterStripProfile`). Üç adım:

1. **Profil seç** — `L 2mm alu`, ya da boş Enter ile mevcut ayarlarla devam.
2. **Bu seferliğine değiştir** — ayar sayfası açılır, `9=2.5` gibi
   yazarsınız. Değiştirilen satırlar `*` ile işaretlenir.
3. **Enter** — şerit çizilir.

```
RUN SETTINGS - changes apply to this run only

[BOX]
 9 Thickness            2.50 mm *
11 Flexibility          0.80
...

N=value   change        ?N   explain
Enter     run
* = changed for this run only
```

Geçici değerler **kayıtlı ayarlarınıza yazılmaz** ve çizim biter bitmez
silinir. Aynı profille tek bir kalınlığı deneyip görmek için budur.

Kalıcı olmasını istiyorsanız ayar sayfasından (`16`) değiştirin, sonra
`P` → `S <ad>` ile profile kaydedin.

---

## 13. Müşteri işi — StockTrack ile otomatik adlandırma

Müşterinin dosyasını açtığınızda onu StockTrack'teki müşteriye bağlar,
stoktaki plakayı seçtirir; dışa aktarılan DXF'e ve takım yolu
klasörlerine adı bu ikisinden verir. Hiçbir adı elle yazmazsınız.

Ana menü `21` (ya da `J`) iş menüsünü açar. Hızlı çalışmak için
`ac2fJobExport` ve `ac2fJobFolders` makrolarını doğrudan araç çubuğuna
koyun (bkz. [kurulum](kurulum.md#araç-çubuğuna-kısayol-koymak)).

| İş menüsü | Makro | Ne yapar |
|---|---|---|
| `1` | `ac2fJobLink` | Dosyayı müşteriye bağlar |
| `2` | `ac2fJobMaterial` | İşin plakasını stoktan seçer |
| `3` | `ac2fJobExport` | DXF'i müşteri klasörüne otomatik adla yazar |
| `4` | `ac2fJobFolders` | Takım yolu klasörlerini açar ya da istifler |
| `5` | `ac2fJobUsage` | Teklif kesildikten sonra tabakadan kullanılan parçayı işaretler (`u1350x1000`) |
| `6` | `ac2fJobOpen` | Müşteri klasörünü Gezgin'de açar |
| `7` | `ac2fJobSetup` | Sunucu, anahtar, gün klasörü, ayraç |
| `8` | `ac2fJobHistory` | Müşterinin eski işleri |
| `9` | `ac2fCheck` | Export öncesi kontrol ([14](#14-export-öncesi-kontrol)) |
| `10` | `ac2fCutTime` | Süre ile fiyat ([15](#15-süre-ile-fiyat)) |
| `11` | `ac2fCutArea` | m² ile fiyat ([16](#16-m-ile-fiyat)) |
| `12` | `ac2fJobDay` | İşin gün klasörü: bugün ya da önceki bir gün (seçimi değiştirir) |

### Klasör düzeni

```
<gün klasörü>\<müşteri>\exp01--pleksi.dxf
                        exp02--dekota.dxf
                        001-pleksi-1350x2000--3pcs--2800um-akpolimer-dokme-beyaz\001
                                                                                \002
                                                                                \003
                                                                                \u1350x1000
                        002-pleksi-1350x2000--1pcs--2800um-akpolimer-dokme-siyah\001
                        003-aluminyum-kompozit-1500x5000--1pcs--4000um-alucobond-gumus\001\p001
                                                                                          \p002
```

- **Gün klasörü**, CDR Auto Renamer'ın her sabah açtığı klasördür
  (`17-10-2026--cumartesi`). Eklenti onun `config.json` dosyasını okur
  (`%LOCALAPPDATA%\CdrAutoRenamer\config.json`): ana dizin, tarih biçimi
  ve ayraç oradan gelir, ikinci bir ayar gerekmez. Bugünün klasörü yoksa
  aynı adla açılır.
- **Müşteri klasörü**: firma adı varsa o, yoksa kişi adı; **küçük harf,
  boşluksuz**, Türkçe harfler sadeleşir, harf/rakam dışı her şey tek `-`
  olur: *Öztürk Reklam Ltd. Şti.* → `ozturk-reklam-ltd-sti`. Bu sürümden
  önce açılmış `Ozturk Reklam Ltd. Sti` gibi klasörler bulunmaya devam
  eder; iş eski klasörde sürerse o klasöre yazılır, yeni klasörler yeni
  kuralla açılır. İş daha önceki bir günde başladıysa nereye yazılacağı
  sorulur (aşağıda).
- **Export**: `exp` + sıra numarası + ayraç + malzeme türü. Numara o
  klasördeki son `exp`'nin bir fazlasıdır.
- **Takım yolu klasörü**: `NNN-tür-ebat--Nadet--kalınlıkµm-marka-varyant-renk--F<gün>.<sahip>.<tür>-<renk>`.
  Kalınlık mikron yazılır (2,8 mm → `2800um`).
- **Belirteç** (sondaki `F21-10-2026.merkez.pleksi-beyaz`): projenin gün
  klasörünün tarihi, malzemenin sahibi (müşterinin kendi malında müşteri
  klasörü, dükkânın stoğunda StockTrack'teki **klasör kodu**), tür-renk.
  Klasör başka bilgisayara ya da CNC'ye kopyalansa da kimin, hangi günün
  işi olduğu adından okunur. İstif belirteci yok sayar (ilk açıldığı gibi
  kalır) ama sahibi farklı tabakalar ayrı klasöre gider.
- **Ebat**: klasör adında her zaman plakanın stoktaki **tam** ölçüsü
  yazar. Plakanın ne kadarının kullanıldığı teklif kesilince belli olur;
  o zaman `5` klasörün içine `u1350x1000` gibi bir dosya koyar (aşağıda).
- Bütün adlar **ASCII**'dir: Türkçe harfler katlanır (`Öztürk` →
  `Ozturk`). CAM programlarının çoğu Türkçe karakterli yolda sorun çıkarır.

### İş başka bir günde başladıysa

Teklif 22'sinde verildi, iş onaylandı ve bugün 25'i. Export, takım yolu
klasörleri ya da klasörü aç çalışınca şu sorulur:

```
This job already has a folder from an earlier day:
   22-10-2026--persembe\ozturk-reklam-ltd-sti
   2 export(s), 1 toolpath folder(s), last exp02--pleksi.dxf

Today is 25-10-2026--pazar.

Yes       carry on in 22-10-2026--persembe
No        start a new folder under today
Cancel    stop

The answer is kept with the job. To change it later:
Customer job menu > 12 (Day folder).
```

- **Evet** → export numarası (`exp03`) ve istifler 22'sindeki klasörden
  devam eder.
- **Hayır** → 25'inin altında yeni müşteri klasörü açılır, numaralar
  baştan başlar.
- Cevap işle birlikte saklanır; **aynı gün** bir daha sorulmaz. Ertesi
  gün yeniden sorulur.
- **Cevabı değiştirmek için** iş menüsü `12` (`ac2fJobDay`): bugün ve
  müşterinin klasörü olan önceki günler listelenir (her birinde kaç export
  ve takım yolu klasörü olduğu, şu anki gün `<- now` ile işaretli);
  seçilen gün işe kaydedilir. Aynı gün içinde, bugüne geri dönmek için de
  çalışır. Dosyayı kaydetmeyi unutmayın. Menünün özetindeki **Folder**
  satırı işin şu anki klasörünü gösterir.
- Dosya henüz bir güne bağlı değilse müşterinin son **30 günde**
  (`Look back`) açılmış klasörüne bakılır. Bulunursa yine sorulur;
  bulunmazsa bugünün altına yazılır.
- Başka bir müşteri bağlanınca günü de sıfırlanır.
- Kullanılan parça (`5`) da önce işin kendi gününe bakar.

### İstifleme

Tür, ebat, kalınlık, marka, varyant ve renk **aynıysa** sayfa aynı
klasöre eklenir: `--2pcs--` → `--3pcs--` olur ve `003` açılır. Bunlardan
biri farklıysa (başka renk, başka marka, başka ölçüde plaka) sıradaki
numarayla yeni klasör açılır. Bir kısmı kullanılacak plaka da tam
plakalarla aynı klasöre istiflenir. Her seferinde ne yapılacağı önce
gösterilir ve onay istenir.

Dikey ebatı **`Long sheet`** (4000 mm) ayarından uzun plakalarda her
`001`, `002` klasörünün içine `p001` ve `p002` açılır. İki katından uzun
plakada `p003` de gelir.

### Tipik akış

```
dosyayı aç → 3 (Export) → müşteri: "ozt" + Enter → plaka: 2 + Enter → DXF penceresi: OK
```

Export, işte eksik olanı kendisi sorar; yeni bir dosyada tek makro
yeter. Sonra takım yolları için `4`:

1. Parçaları **Nesting** ile plakaya dizdiyseniz (plaka seçilince
   Nesting'in plaka ölçüsü de ona ayarlanır), eklenti sayfadaki plaka
   çerçevelerinden üzerinde parça olanları sayar.
2. Nesting yoksa seçim için bir plaka önerir.
3. Öneri kutuda hazır yazılıdır: Enter kabul eder, ya da düzeltirsiniz.

| Yazın | Anlamı |
|---|---|
| `3` | işin plakasından 3 adet |
| `1 2050x3050` | başka ölçüde 1 plaka |
| `2, 1 2050x3050` | ikisi birden |
| `M` | önce başka plaka seç |

Teklif kesildikten sonra `5` (aşağıda).

### Eski işler

İş menüsü `8` (`ac2fJobHistory`): bağlı müşterinin son bir yılda klasörü
olan günleri, yeniden eskiye listeler.

```
OLD JOBS  -  ozturk-reklam-ltd-sti
  1  25-10-2026--pazar       2 export, 1 folder
  2  22-10-2026--persembe    3 export, 2 folder, 1 cdr
```

- `2` → o klasörü Gezgin'de açar.
- `2 O` → klasördeki CorelDRAW dosyasını açar ("geçen ayki işin
  aynısı").
- `2 C` → bu işi o klasörden devam ettirir: export ve takım yolu
  klasörleri bugün oraya gider.

### Kullanılan parça: `u` dosyası

Teklifi StockTrack'te kestikten sonra iş menüsü `5` (`ac2fJobUsage`):

1. Müşterinin son teklifleri listelenir (yeniden eskiye); Enter en
   yenisini alır.
2. Teklifin tabaka kalemleri okunur: hangi plakadan ne ölçü
   kullanıldığı.
3. Her kalem için takım yolu klasörü bulunur (tür, kalınlık, marka,
   varyant, renk aynı; müşterinin klasöründe, teklif gününden iki hafta
   geriye kadar). Plakanın **yalnız bir kısmı** kullanıldıysa o klasörün
   içine o parçanın adını taşıyan bir dosya yazılır:

```
001-pleksi-1350x2000--1pcs--2800um-akpolimer-dokme-beyaz\u1350x1000
```

1350 × 2000 pleksinin 1000 × 1350'si kullanıldı demektir. Önce
plakanın enine denk gelen kenar yazılır (`u1350x1000`). Tam kullanılan
plakaya dosya konmaz. Dosyanın içinde teklif no, tarih ve adet yazar.

Ne yazılacağı önce gösterilir. Aynı dosya zaten varsa dokunulmaz, yani
iki kez çalıştırmak bir şey bozmaz. Teklif kaleminde ölçü girilmemişse
StockTrack plakanın tamamını yazar ve dosya konmaz; parça satıyor ya da
işliyorsanız kalemde en × boyu girin.

### Teklif işe dönüşünce: StockTrack Runner

StockTrack'te teklif *Satış/İşlemeye Dönüştür* ile işe döndüğünde
klasörleri dükkân bilgisayarındaki
[stock-track-runner](https://github.com/ac2f/stock-track-runner) açar;
Corel'i açmanız gerekmez:

- müşteri klasörü (teklif günü ile dönüşüm günü arasında müşterinin bir
  klasörü varsa o, yoksa dönüşüm gününün altında yeni),
- teklifteki her kesilecek tabaka için takım yolu klasörü, aynı adlarla,
  aynı belirteçle ve aynı istifleme kuralıyla
  (`001-pleksi-1220x2440--2pcs--...--F21-10-2026.merkez.pleksi-beyaz\001`,
  `\002`, uzun tabakada `p001`/`p002`),
- tabakanın bir kısmı kullanıldıysa `u1220x900` dosyası,
- **malzeme satışında** (teklifte kesim kalemi olmayan satış) takım yolu
  klasörü değil `malzeme-F21-10-2026.merkez.pleksi-beyaz` klasörü ve
  içinde satılan her ölçü için `m1220x2440-2800um-akpolimer-dokme-2adet`
  gibi bir dosya. Eklenti bu klasörlere dokunmaz.

Runner'ın klasör açtığı işler için **Toolpath folders**'ı (iş menüsü `4`)
ayrıca çalıştırmayın, tabakalar ikinci kez istiflenir. Eklenti müşteri
klasöründe runner'ın `.stocktrack` işaretini görürse onay penceresinde
hangi teklifler için açıldığını söyler. **Export DXF** ve **Used parts**
eskisi gibi kullanılır. Kurulum ve ayarlar runner deposunun README'sinde.

### Müşteri bağlama

Bir **liste penceresi** açılır: son seçtiğiniz müşteriler ve StockTrack'te
en son işlem gören müşteriler. Üstteki kutuya adın, firma adının ya da
telefonun bir parçasını yazdıkça liste daralır (Türkçe harfe duyarsız:
`ozturk` → *Öztürk*). Müşteriye çift tıklayın ya da oklarla gelip
**Enter**'a basın.

- Listede kalmadıysa **Enter** (ya da *Search StockTrack*, F3) bütün
  müşterilerde arar. Bulunanların **hepsi listelenir — tek sonuç çıksa
  bile** kendiliğinden bağlanmaz, siz seçersiniz.
- *Use as typed* (ya da başına `=` koyup Enter: `=Mehmet Usta`)
  StockTrack'e bakmadan, yazıldığı gibi kullanır.
- Pencere PowerShell ile çizilir (Windows'la gelir). PowerShell
  engelliyse ya da ayarlarda **List window** `0` ise eski numaralı soru
  açılır; orada da tek sonuç artık listelenip sorulur.
- Dosya zaten bir `<gün>\<müşteri>\` klasöründen açıldıysa (eski
  numaralı soruda) o müşterinin adı hazır gelir; Enter yeter.
- Bağlantı **dosyanın içine** yazılır (belge özelliği): dosyayı kaydedip
  sonra açtığınızda müşteri ve plaka yerinde durur. Kaydedilmemiş
  dosyalar için kayıt defterinde de bir kopyası tutulur.

### Plaka seçme

Müşteri bağlamadaki gibi bir liste penceresi açılır. Sırası:
**müşterinin kendi malzemesi** (StockTrack'te ona ait konsinye stok,
`[musterinin]` işaretli), son seçtikleriniz, sonra işletme stoğu.
Kutuya yazdıkça liste daralır; listede kalmayınca **Enter** (ya da
*Search stock*, F3) stokta arar: `pleksi beyaz 3`, `dokme 2,8`,
`1220x2440`, renk kodu ya da stok kodu. Bulunanların hepsi listelenir,
tek sonuç da seçilmeden kullanılmaz.

Teklif seçerken (*Used parts*) ve **eski işlerde** de aynı pencere
çıkar; eski işlerde *Open its CorelDRAW file* ve *Carry this job on here*
düğmeleri vardır.

StockTrack'te olmayan bir plaka için (kutuya yazıp *Use as typed* ya da
Enter):

```
=pleksi;2.8;akpolimer;dokme;beyaz;1220x2440
```

Yalnız tür zorunludur: `=dekota` da olur (takım yolu klasöründe ebatı
sorar).

### Kurulum

1. StockTrack'te **Ayarlar › Corel entegrasyonu** → *Anahtar oluştur*.
   Sunucu adresi ve anahtar orada yazar. **Her dükkânın kendi anahtarı
   vardır:** eklentiye hangi dükkânın anahtarını yapıştırırsanız müşteri,
   stok ve teklif listeleri o dükkândan gelir. Anahtarı üretmek için o
   dükkânın ayarlarını değiştirme izni gerekir.
2. Corel'de iş menüsü `7` (`ac2fJobSetup`):

```
1=http://192.168.1.50/api/v1          (dükkândaki sunucu)
1=https://stok.ornek.com/api/v1       (uzak sunucu)
2=<anahtar>
5=192.168.1.9                         (dükkândaki YEREL sunucu, isteğe bağlı)
T          bağlantıyı dener
L          yalnız yerel sunucuyu dener
```

3. Gün klasörü CDR Auto Renamer'dan gelir. O program yoksa ya da başka
   bir yer istiyorsanız `3=D:\Isler` yazın; `3=` (boş) geri alır.
4. Ayraç varsayılan `--`; `4=` ile değişir.

**Yerel sunucu (5):** dükkânda StockTrack'in yerel kopyası (replika) çalışıyorsa
adresini `5=` ile yazın. Ana sunucu cevap vermezse (internet yok, sunucu
kapalı, önündeki ağ geçidi 502/503/504 diyor) eklenti yerel sunucuya sorar;
başlıkta `(local server)` yazar. Geçişten sonra 5 dakika boyunca önce yerel
sunucuya sorulur (her listede ana sunucunun zaman aşımı beklenmez), sonra
yine ana sunucu denenir. Ana sunucu gerçek bir hata verirse (anahtar
geçersiz, sürüm eski) bu yerel sunucuyla örtülmez. Anahtar aynıdır: yerel
sunucu dükkânın ayarlarını merkezden alır. `5=` (boş) yerel sunucuyu kaldırır.
Runner da bu adresi kullanır.

StockTrack'e (yerel sunucu dahil) ulaşılamazsa eklenti çalışmayı sürdürür:

- StockTrack'ten gelen her liste bilgisayarda saklanır
  (`%LOCALAPPDATA%\ac2f\stocktrack`, sunucu ve anahtar başına ayrı). Sunucu
  cevap vermezse (kapalı, yeniden başlıyor, internet yok) aynı liste
  buradan gelir; başlıkta `(offline, saved 17.10.2026 09:30)` yazar.
- Daha önce hiç aranmamış bir müşteri ya da plaka araması, o güne dek
  görülen bütün müşteriler / işletme plakaları içinde yapılır (en fazla
  1000'er satır). Müşterinin kendi plakaları yalnız o müşterinin
  kaydedilmiş listesinden gelir — çevrimdışıyken başka müşteriye
  önerilmez.
- Hiçbiri yoksa müşteri adını ve plakayı yazarak devam edersiniz.

### Ayarlar (75-81)

| # | Ayar | Varsayılan | Ne işe yarar |
|---|---|---|---|
| 56 | `Long sheet` | 4000 mm | Bundan uzun plakaya `p001`, `p002` |
| 57 | `Export dialog` | `1` | DXF export penceresi gösterilsin mi |
| 58 | `Open folder` | `0` | İşten sonra müşteri klasörünü aç |
| 59 | `Nest sheet` | `1` | Plaka seçilince Nesting plaka ölçüsünü de ayarla |
| 60 | `Look back` | 30 gün | Yeni işte müşterinin eski klasörü bu kadar geriye aranır; `0` = hep bugün |
| 61 | `Check first` | `1` | Her export'tan önce kontrol ([14](#14-export-öncesi-kontrol)) |

`Export dialog = 0` yapınca DXF, Corel'in son kullandığı ayarlarla
pencere açmadan yazılır.

### Bilinmesi gerekenler

- StockTrack'ten **yalnız okunur**; stok düşülmez, iş kaydı açılmaz.
- `u` dosyası teklifin kendi müşterisinin klasörüne gider; dosyayı
  başka bir müşteriye bağladıysanız önce doğru müşteriyi bağlayın.
- Bir istif klasörünü elle yeniden adlandırırsanız eklenti onu
  tanımayabilir; adı bozulmamış klasörler istiflenir.
- Klasör Gezgin'de ya da CAM programında açıksa yeniden adlandırma
  (`2pcs` → `3pcs`) başarısız olabilir; eklenti bunu raporlar ve o
  klasöre dokunmaz.

## 14. Export öncesi kontrol

`ac2fCheck` (iş menüsü `9`) seçimi, seçim yoksa sayfayı kontrol eder ve
sorunlu nesneleri istersen seçer.

| Kesimi bozar | Neden |
|---|---|
| Eğriye çevrilmemiş yazı | DXF'e yazı olarak gider, CAM kesemez (`Ctrl+Q`) |
| Bitmap / OLE | Kesilecek bir şey yok |
| Çift çizgi | Üst üste iki nesne: aynı çizgi iki kez kesilir |
| Plakadan taşan | Nest plakasının kenarını aşan parça; nest yoksa işin plakasından büyük iş |
| Plaka çerçevesi | Nesting'in plaka dikdörtgeni export'a karışmış: plakanın etrafı da kesilir |

| Bakmaya değer | |
|---|---|
| Açık kontur | Kazıma için doğru, parça kesimi için hata |
| Kırıntı | 0,5 mm'den küçük nesne, kaçak nokta |
| PowerClip | İçerik göründüğü gibi export edilmeyebilir |

**Export DXF her seferinde kontrol eder** (`Check first`, ayar 80).
Yalnız *kesimi bozan* bir şey varsa durur:

```
Yes       export anyway
No        stop and select them
Cancel    stop
```

Açık kontur, kırıntı ve PowerClip export'u hiç durdurmaz.

Çift çizgi, kutusu ve uzunluğu 0,05 mm içinde aynı olan iki nesnedir;
gruplar ve PowerClip içerikleri de taranır.

## 15. Süre ile fiyat

Seçimi yapın, `ac2fCutTime` (iş menüsü `10`):

```
CUT TIME  -  ozturk-reklam-ltd-sti
48 outline(s), 2 pass(es)  (pleksi 5 mm at 3 mm a pass)

Cutting   12.450 mm x2 at 25 mm/s       16 min 36 s
Travel    6.320 mm at 120 mm/s          53 s
Plunges   48 x2 x 1.5 s                 2 min 24 s
Total                                   19 min 53 s

Billed    30 min  (minimum 30)
Price     30 min x 1.500,00 TRY/h = 750,00 TRY
```

- **Kesim** = kontur boyu × paso ÷ kesim hızı.
- **Boşta gezme**, kabaca: işin sol alt köşesinden en yakın kontura,
  onu kesip bir sonraki en yakına, en sonda köşeye dönüş. Açık kontura
  yakın ucundan girilir. CAM programının yolu genelde biraz daha kısadır;
  tahmin güvenli taraftadır. 2000 konturdan sonra alandan tahmin edilir.
- **Dalma** = kontur × paso × dalma süresi.
- **Paso** = işin plakasının kalınlığı ÷ paso derinliği (yukarı yuvarlanır;
  derinlik 0 ise tek paso).
- **Fiyat** = faturalanan dakika ÷ 60 × saatlik ücret; faturalanan dakika
  asgari süreden az olmaz.

Değerler StockTrack'ten gelir (*Ayarlar › Kesim fiyatlandırması*), son
okunan saklanır, StockTrack'e ulaşılamasa da çalışır. O iş için
değiştirmek:

```
f=30        kesim hızı (mm/sn)       r=150   boşta gezme hızı
p=1         dalma süresi (sn)        n=2     paso sayısı
h=1800      saatlik ücret            m=20    asgari dakika
```

Birden fazlası tek satırda: `f=30 n=1`.

## 16. m² ile fiyat

Seçimi yapın, `ac2fCutArea` (iş menüsü `11`):

```
AREA PRICE  -  ozturk-reklam-ltd-sti

1  Selection box     1180 x 870 mm         1,027 m2  <- used
2  Each object's box 6 object(s)           0,640 m2
3  Nest sheets       2 x 1220x2440         5,954 m2

Billed    1,027 m2
Price     1,027 m2 x 450,00 TRY = 462,15 TRY
```

| Alan | Ne zaman |
|---|---|
| `1` Seçim kutusu | Tek pano, tabela |
| `2` Nesnelerin kutuları | Ayrı parçalar, harfler: her biri kendi kutusuyla |
| `3` Nest plakaları | İşin gerçekten kaç plaka aldığı (sayfada nest çerçevesi varsa) |

Sayfada nest çerçevesi varsa `3`, yoksa `1` seçili gelir; rakam yazınca
değişir. **Fiyat** = max(alan × adet, asgari m²) × m² ücreti. O iş için:
`a=450` (m² ücreti), `m=0.5` (asgari m²), `x=4` (adet).

## 17. Başlangıç noktaları

Lazer ya da router'ın her parçaya **nereden** girdiğini ayarlar: her
yolun ilk düğümü seçilen köşeye, kenar ortasına ya da merkeze taşınır.
İstenirse kesim yönü de düzenlenir.

Seçin, `ac2fStartPoints` (menü `23`) çalıştırın. Konum sayısal tuş
takımı düzenindedir:

```
        7 ----- 8 ----- 9
        |               |
        4       5       6
        |               |
        1 ----- 2 ----- 3
```

Sayıdan sonra harflerle seçenek değişir (hatırlanır):

| Harf | Anlamı |
|---|---|
| `E` | Her konturun **kendi** kutusu (delik kendi köşesinden başlar) |
| `O` | Her nesnenin kutusu |
| `S` | Bütün seçimin kutusu |
| `X` | Tam nokta: gerekirse oraya düğüm eklenir (şekil değişmez) |
| `N` | En yakın mevcut düğüm, ekleme yok |
| `K` | Yön olduğu gibi |
| `W` | Dış saat yönünde, delikler ters |
| `C` | Dış saat yönünün tersine, delikler saat yönünde |
| `A` / `Z` | Hepsi saat yönünde / hepsi tersine |
| `M` / `Q` | Sonra işaret çiz / çizme |

Örnek: `8 S W` = seçimin üst ortası, dış saat yönünde. `7` tek başına =
her konturun sol üstü.

- `ac2fStartPointsAgain` (menü `24`) son ayarla sormadan uygular.
- `ac2fStartPointsShow` (menü `25`) her başlangıca küçük kırmızı bir
  daire ve yönü gösteren kısa bir çizgi koyar (`ac2f start points`
  katmanı); tekrar çalıştırınca siler.
- Dörtgen, elips ve çokgenler önce eğriye çevrilir (düğümleri yoktur);
  metin atlanır. Rapor kaç yolun değiştiğini söyler.
- Açık yollar (çizgiler) diğer ucu konuma daha yakınsa ters çevrilir.
- Yön değiştirmeyi desteklemeyen bir CorelDRAW sürümünde başlangıç yine
  taşınır, rapor yönün değişmediğini yazar.
- Nesting penceresindeki `Set start points afterwards` bu ayarları
  yerleşen parçalara uygular.
- Ayarlar sayfada `#6` (START) altındadır. `Ctrl+Z` geri alır.

## Bilinmesi gerekenler

- Ölçüm **her zaman milimetre** üzerinden yapılır; belgenizin birimi ne
  olursa olsun sonuç doğrudur ve işlem sonunda birim geri alınır.
- Kontur uzunluğu, çizginin **kalınlığını dikkate almaz** — yolun kendi
  uzunluğudur.
- Çok sayıda nesne seçildiğinde ölçüm birkaç saniye sürebilir; ekran
  tazeleme bu sırada kapatılır.
- Nesne kırılımı listesinde en fazla 15 satır gösterilir; toplam her zaman
  seçimin tamamını kapsar.
