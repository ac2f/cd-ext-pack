# ac2f pack

CorelDRAW için VBA eklenti paketi.

Vektörlerin çevresindeki çizgilerin toplam uzunluğunu ölçer, bu ölçüme
dayanarak 3'lü LED modül yerleşimi için gereken adetleri hesaplar, kutu
harf yan bordürünün açınımını çıkarıp kesime hazır derzli şeritler çizer
alüminyum kompozit paneller için V derz çizgilerini üretir ve neon LED
şerit kanalı için dolu tasarımı tek çizgiye düşürür.

| Modül | İş |
|---|---|
| `ac2fCore` | Ortak çekirdek: ayarlar, birim/sayı yardımcıları, ölçüm motoru |
| `ac2fLength` | **Uzunluk ölçümü** — toplam kontur uzunluğu |
| `ac2fLedModule` | **3'lü LED modül hesabı** — modül, LED, güç, güç kaynağı adedi |
| `ac2fBoxLetter` | **Kutu harf şeridi** — bordür açınımı ve derz yerleşimi |
| `ac2fPanel` | **ACP panel derzi** — alüminyum kompozit V derz yerleşimi |
| `ac2fCenterline` | **Orta hat** — dolu tasarımı tek çizgiye düşürür |
| `ac2fNest` | **Nesting** — parçaları plakalara yerleştirir |
| `ac2fPlot` | **Plotter** — HPGL üretip ağdan doğrudan gönderir |
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
| `ac2fCenterline` | Tek çizgiye düşürür (birleşik / ayrı ayrı sorar) |
| `ac2fCenterlineJoined` | Tüm seçimi tek bölge sayar — birleşik el yazısı |
| `ac2fCenterlineSeparate` | Her nesneyi ayrı ayrı düşürür |
| `ac2fNest` | Seçili parçaları plakalara yerleştirir |
| `ac2fNestReport` | Hesaplar ve raporlar, hiçbir şeyi oynatmaz |
| `ac2fPlotSend` | Seçimi HPGL olarak plotter'a gönderir |
| `ac2fPlotOrderPreview` | Seçimi kesim sırasına göre sayfada numaralandırır |
| `ac2fPlotTest` | 40 × 80 mm test L'si gönderir — makine mi dosya mı, ayırır |
| `ac2fPlotSave` | HPGL'i dosyaya yazar (incelemek için) |
| `ac2fPlotFixSend` | Var olan bir `.plt`'yi normalize edip gönderir |
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

## Orta hat (tek çizgiye düşürme)

Pleksi üzerine 6 mm bıçakla neon şerit kanalı açmak için, dolu tasarımı
merkezden geçen tek çizgiye indirir.

Tasarımı seçin, `ac2fCenterline` çalıştırın ve **`J`** ya da **`S`** yazın:

| Seçenek | Ne yapar |
|---|---|
| `J` **birleşik** | Tüm seçim tek bölge sayılır; **birbirine değen harfler bağlı kalır** — el yazısı için |
| `S` **ayrı ayrı** | Her nesne kendi başına düşürülür |

Doğrudan `ac2fCenterlineJoined` / `ac2fCenterlineSeparate` de çağrılabilir.

### Nasıl çalışıyor

1. Konturlar poligona açılır (dairesel yay modeli — bezier kontrol
   noktasına gerek yok)
2. **Şekil başına** çift-tek dolgu (delikler doğru çıkar), şekiller
   **birleştirilir** (değen harfler tek bölge olur)
3. Zhang-Suen inceltmesiyle tek piksel kalınlığa iner
4. Fazlalık merdiven pikselleri atılır, zincirler izlenir, kısa
   çıkıntılar budanır
5. Serbest uçlar uzatılır — inceltme uçları yarım kalınlık geri yer
6. Zikzak yumuşatılır, Douglas-Peucker ile sadeleştirilir
7. Her zincir **tek polyline** olarak, kırmızı, gruplu çizilir

### Ölçülen doğruluk

Orta hattı bilinen şekiller üzerinde:

| Test | Sonuç | Gerçek | Hata |
|---|---|---|---|
| Dalgalı şerit | 470,8 mm | 471,0 mm | %0,0 |
| T kavşağı (3 dal) | 279,3 mm | 280,0 mm | %0,2 |
| İki ayrı harf | 161,0 mm | 160,0 mm | %0,6 |

Uç konumu hatası 1,8 mm (uç uzatma olmadan 12,1 mm).

## Nesting

Parçaları seçin, `ac2fNest` çalıştırın. **Her kenardan ayrı boşluk**,
**parçalar arası boşluk** ve **döndürme adımı** ayarlanabilir (ayar
sayfası, `#5` NEST).

| Ayar | Varsayılan | Ne yapar |
|---|---|---|
| Marg left / right / top / bottom | 10 mm | Her kenardan ayrı pay |
| Part gap | 3 mm | Parçalar arası boşluk |
| Rot step | 90° | Denenecek dönüş adımı. `0` = olduğu gibi |
| Sheet W / H | 0 | Plaka ölçüsü. `0` = sayfa ölçüsü |
| Sheet gap | 20 mm | Birden çok plaka çizilirken arası |

`Rot step`: `0` hiç döndürmez, `90` dört yönü dener, `180` yalnız yarım
tur (desenli/hasırlı malzeme için), `15` en ince adım.

### Algoritma ve bir uyarı

Sınırlayıcı kutular üzerinde **skyline bottom-left**. Parçalar en uzun
kenara göre sıralanır; her parça için tüm açılar ve tüm skyline
basamakları denenir ve **skyline'ı en alçak bırakan** yerleşim seçilir.

Bu skor seçimi kritik. Bariz görünen "en düşük y" seçimi döndürmeyi
**zararlı** hâle getiriyor — beş parça setinde ölçtüm, ikisinde %8 ve
%18 malzeme kaybettirdi. Skyline yüksekliğine göre skorlayınca aynı
setler %8, %7, %18, %5 ve %13 kazandı, hiçbiri gerilemedi:

| Set | En düşük y | **Skyline yüksekliği** |
|---|---|---|
| Eşit kutular | −%8,3 | **+%8,1** |
| Karışık | +%1,6 | **+%7,2** |
| Uzun şeritler | +%5,8 | **+%17,6** |
| Dar-uzun | −%18,5 | **+%5,0** |
| 450×160 | +%7,6 | **+%13,5** |

Testlerin kapsamadığı bir parça karışımında yine de ters tepebilir, o
yüzden döndürme açıkken iş **iki kez** yerleştirilir — döndürmeli ve
döndürmesiz — ve iyi olan tutulur. **Döndürme asla malzeme kaybettirmez.**

> Parçalar **sınırlayıcı kutuyla** yerleştirilir, gerçek konturla değil.
> İçbükey parçalar birbirine geçmez.

Sığmayan parçalar yerinde bırakılır ve raporda belirtilir. `Ctrl+Z`
her şeyi geri alır.

## Plotter'a gönderme

Seçimi seçin, `ac2fPlotSend` çalıştırın, adresi onaylayın. **Export yok,
düzeltilecek dosya yok.**

```
192.168.1.100:9100
```

### Yön: nasıl görüyorsanız öyle

**Hiçbir şey döndürülmez.** Belge X'i plotter X'i, belge Y'si plotter
Y'si olur; iş ekranda durduğu gibi kesilir. Arada döndürecek bir export
süzgeci yoktur.

Yine de dönük çıkıyorsa sebep makinedir: çoğu kesicide X ekseni malzeme
besleme yönünde uzar, bu yüzden geniş bir iş ruloya enine düşer. `Rotate`
ayarı (42) bunu telafi eder, **varsayılanı 0** — yani döndürme yok.
90 saat yönünün tersine, 270 saat yönüne çevirir; en–boy takas olur ve
kenar payı sonradan uygulandığı için iş yine köşeye oturur.

### Kesim sırası: hangi kontur önce gider

Kutu harf şeridi bükerken sıra, makineyi kaç kez kurduğunuzu belirler.
Paket konturları önce **harflere** ayırır, sonra harfleri bir iş sırasına
dizer.

**1. İçine alma.** Her kapalı kontur diğerlerine karşı sınanır: ilk
noktası içine düşüyor mu, ve öteki daha büyük mü? Hiçbir şeyin içinde
olmayan kontur bir **dış**, bir dışın içindeki kontur onun **göbeği**
(A'nın ortası), göbeğin içindeki yine bir dıştır. Bir dış ve kapsadığı
her şey **bir harftir**.

**2. Çalışma sırası.** Harfler soldan sağa dizilir. X'te üst üste binen
harfler bir **sütun** sayılır ve sütun içinde yukarıdan aşağı gidilir —
"soldan sağa, önce üst" bir sayfada satır satır duran parçalar için
budur.

**3. Önce aynılar.** Bir harfe gelindiğinde, ondan sonraki **aynı
şekilli** bütün harfler hemen arkasına çekilir. `NOBOBOXB` kelimesindeki
üç `O`, araya `B` ve `X` girmeden peş peşe bükülür: üç kurulum yerine
bir kurulum.

| `Order` (ayar 44) | Sıra |
|---|---|
| `0` | Belgedeki hâliyle |
| `1` *(varsayılan)* | Soldan sağa, önce üst |
| `2` | Aynısı, ama **önce ovalliği olmayan** harfler, sonra eğriler |
| `3` | 2 ile aynı, ama **göbekler en sona** |

`NOBOBOXB` için sonuç (doğrulandı):

```
1 ->  N  O O. O O. O O.  B B.1 B.2  B B.1 B.2  B B.1 B.2  X
2 ->  N  X  O O. O O. O O.  B B.1 B.2  B B.1 B.2  B B.1 B.2
3 ->  N  X  O O O  B B B  |  O. O. O.  B.1 B.2  B.1 B.2  B.1 B.2
```

`.` ile bitenler göbek. `2`, rulolu bir bükme makinesi içindir: düz işin
tamamı, rulolar hiç devreye girmeden biter. `3` daha ileri gider,
hiçbir dış bir göbeği beklemez; göbekler de harflerle **aynı sırayı**
korur.

**Aynı sayılma ölçütü.** İki harf; kontur sayısı, uzunluk, kapsadığı
alan ve kutu ölçüsü `Match tol` (48) içinde uyuşuyorsa aynıdır. En ve
boy yer değişebilir, yani çeyrek tur döndürülmüş bir kopya da aynı
sayılır. **Aynalanmış kopya sayılmaz** — ters yöne bükülür, kendi
kurulumunu ister; alanın işareti bunu yakalar.

**Gruplama sırayı ezer.** `Group same` (45) açıkken soldan sağa sıra
bozulabilir; zaten amacı budur. Sayfadaki sıra kurulum sayısından daha
önemliyse `0` yapın.

| Ayar | Varsayılan | Ne işe yarar |
|---|---|---|
| `Order` (44) | `1` | Yukarıdaki dört algoritma |
| `Group same` (45) | `1` | Aynı harfleri peş peşe bük |
| `Column overlap` (46) | 5 mm | Bir sütun sayılmak için gereken X bindirmesi |
| `Curve limit` (47) | 0,5 mm | Bir harfin "eğri" sayılması için gereken şişkinlik |
| `Match tol` (48) | 0,5 mm | İki harfin aynı sayılması için gereken yakınlık |

Sıra **ekranda gördüğünüz çizim üzerinden**, `Rotate` ve `Mirror`'dan
önce hesaplanır; yani "sol" her zaman ekranınızın solu.

**Göndermeden önce görün.** Ana menü `13` (`ac2fPlotOrderPreview`)
seçimi sayfada kesim sırasına göre numaralandırır. Tek `Undo` hepsini
siler.

### Neden düzeltme adımı yok

Export edilmiş bir `.plt`'nin olağan sorunu, geometrinin sayfanın verdiği
koordinatlarda (sık sık negatif) kalması ve plotter'ın kabul etmesi için
ayrıştırılıp kaydırılması gerekmesidir. Burada kaydırma **HPGL yazılırken
geometriden** uygulanıyor — ayrıştırılacak bir şey yok, düzeltilecek bir
şey yok.

HPGL birimi milimetrede 40'tır (inçte 1016).

### Ara dosya var mı

**CorelDRAW'ın PLT export'u hiç kullanılmaz** — modülde tek bir `Export`
çağrısı yoktur. HPGL doğrudan geometriden yazılır.

Gönderirken baytlar, gönderici sürece devretmek için geçici bir dosyadan
geçer ve **gönderim başarılı olunca dosya silinir**; geriye bir şey
kalmaz. Gönderim başarısız olursa dosya bilerek bırakılır, raporda yolu
vardır, elle gönderebilirsiniz. Dosyayı kalıcı istiyorsanız
`ac2fPlotSave` var.

### Gönderim yolu

VBA'nın kendi socket'i yok. Baytlar kısa bir PowerShell betiği üzerinden
`System.Net.Sockets.TcpClient` ile çıkıyor. PowerShell desteklenen her
Windows'ta var; bu sayede `Declare`, 32/64 bit sorunu ve OCX kaydı
gerekmiyor. Betik bir günlük yazıyor, makro onu geri okuyor — bağlantı
reddedilirse gerçek sebebi görüyorsunuz.

Bu **ham TCP akışı**; 9100 ya da bir telnet portunda dinleyen
plotter'ların beklediği şey. Telnet protokolü anlaşması yapılmaz.

### Dönük ya da kesmiyorsa: önce test L'si

`ac2fPlotTest` (ana menü 14) gerçek işin kullandığı **aynı** yazıcı,
gönderici ve yönlendirme adımından 40 × 80 mm'lik bir **L** gönderir:

```
   |
   |
   |___     uzun bacak SOLDA, ayak SAĞA, 80 mm uzun yön
```

| Sonuç | Anlamı | Ne yapmalı |
|---|---|---|
| Doğru şekil, temiz kesti | Her şey yolunda | — |
| **Yan yatmış** | Makine eksenleri takas ediyor | `Rotate = 270`, ters olursa `90` |
| **Ayna görüntüsü** | Makine bir ekseni aynalıyor | `Mirror = 1` (ya da baş aşağıysa `2`) |
| Doğru şekil ama sadece gezdi | Makine hiç kesmiyor | Bıçak kuvveti/derinliği; `Preamble = 2` |
| Kıpırdamadı | Ulaşmadı | Adres, port, kablo |

Kare neden olmaz: hangi yöne çevrilse aynı görünür, aynalansa yine
karedir. L dört dönüşü ve üç aynalamayı birbirinden ayırır.

İki ayar bu iş için:

- **`PD pairs`** — komut başına koordinat çifti. `1` ile her komut bir
  düzine karakter olur. Eskiden 40'tı ve 180 noktalı bir daire için
  **402 karakterlik** satırlar üretiyordu; bunu kaldıramayan bir kesici
  satırın sonunu düşürür, sonraki `PD` yutulur, bıçak hiç inmez.
- **`Preamble`** — `2` yaparsanız `IN;` gönderilmez. `IN` cihazı açılış
  varsayılanlarına döndürür ve pek çok kesicide panelden ayarladığınız
  **bıçak kuvvetini de siler**.

### Diğer iki makro

| Makro | Ne yapar |
|---|---|
| `ac2fPlotSave` | HPGL'i dosyaya yazar, göndermez |
| `ac2fPlotFixSend` | Var olan bir `.plt`'yi normalize eder, isterseniz gönderir |

`ac2fPlotFixSend`, eski Python betiğinizin işini yapar — ama bir farkla:
**çizim alanı yalnız çizim komutlarından** hesaplanır (her `PD`, ve
ardından `PD` gelen `PU`'lar). Sondaki `PU0,0;` kalem park komutudur,
çizim değildir. Hesaba katılırsa, çizimi tamamen pozitif koordinatlarda
olan bir dosyada min 0,0 çıkar ve normalizasyon sabit bir kaydırmaya
döner — iş kenarda kalmaz, malzeme boşa gider.

## Ayarlar ve profiller

**Tüm ayarlar tek sayfada.** Ana menüden `17` ile açılır; 51 ayar
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

1.9.0 — bkz. [CHANGELOG.md](CHANGELOG.md)
