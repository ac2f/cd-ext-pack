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
| `ac2fNest` | **Nesting** — parçaları gerçek konturlarıyla tabakalara yerleştirir (pencere, tabaka listesi, adetler) |
| `ac2fNestCore` | Nesting motoru — saf hesap, CorelDRAW dışında sınanır |
| `ac2fStart` | **Başlangıç noktaları** — yolların ilk düğümünü köşeye / kenar ortasına / merkeze taşır, yönü ayarlar |
| `ac2fBending` | **Bükme sırası** — konturları harflere ayırır, sıraya dizer |
| `ac2fPlot` | **Plotter** — HPGL üretip ağdan doğrudan gönderir |
| `ac2fJob` | **Müşteri işi** — StockTrack müşterisi ve stoğuyla export ve takım yolu klasörlerini adlandırır |
| `ac2fCheck` | **Export öncesi kontrol** — yazı, bitmap, çift çizgi, plakadan taşan parça |
| `ac2fCutTime` | **Süre ile fiyat** — kesim + boşta gezme + dalma süresi, saatlik ücretle fiyat |
| `ac2fCutArea` | **m² ile fiyat** — seçim / parça / nest plakası alanı, m² ücretiyle fiyat |
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
| `ac2fPanelGroove` | Dörtgen çevresine ACP V derz çizgilerini çizer; panel plakadan uzunsa **plakalara böler** |
| `ac2fPanelSplit` | Aynısı, ama plaka/derz payı sorularını her zaman sorar (ana menü `28`) |
| `ac2fCenterline` | Tek çizgiye düşürür (birleşik / ayrı ayrı sorar) |
| `ac2fCenterlineJoined` | Tüm seçimi tek bölge sayar — birleşik el yazısı |
| `ac2fCenterlineSeparate` | Her nesneyi ayrı ayrı düşürür |
| `ac2fNest` | Nesting penceresi (tabakalar, adetler, ayarlar, hazır ayarlar), sonra yerleştirir |
| `ac2fNestReport` | Aynı pencere; hesaplar ve raporlar, hiçbir şeyi oynatmaz |
| `ac2fNestAgain` | Pencere açmadan son ayarlarla yerleştirir |
| `ac2fNestMarkSheets` | Seçili şekilleri çizili artık tabaka olarak adlandırır |
| `ac2fStartPoints` | Başlangıç noktası: konum (7 8 9 / 4 5 6 / 1 2 3), kutu, tam nokta, yön |
| `ac2fStartPointsAgain` | Son başlangıç noktası ayarıyla sormadan uygular |
| `ac2fStartPointsShow` | Başlangıç noktalarına işaret koyar / kaldırır |
| `ac2fBendOrder` | Seçili nesneleri bükme sırasına dizer (istif + isim) |
| `ac2fBendPreview` | Bükme sırasını sayfada numaralandırır |
| `ac2fBendReport` | Bükme sırasını listeler, hiçbir şeyi değiştirmez |
| `ac2fPlotSend` | Seçimi HPGL olarak plotter'a gönderir |
| `ac2fPlotTest` | 40 × 80 mm test L'si gönderir — makine mi dosya mı, ayırır |
| `ac2fPlotSave` | HPGL'i dosyaya yazar (incelemek için) |
| `ac2fPlotFixSend` | Var olan bir `.plt`'yi normalize edip gönderir |
| `ac2fJob` | **Müşteri işi menüsü** (ana menü `21` / `J`) |
| `ac2fJobLink` | Dosyayı StockTrack'teki müşteriye bağlar |
| `ac2fJobMaterial` | İşin plakasını stoktan seçer (müşterinin kendi malı önce) |
| `ac2fJobExport` | DXF'i `<gün>\<müşteri>\exp01--pleksi.dxf` olarak yazar |
| `ac2fJobFolders` | Takım yolu klasörlerini açar, aynı plakaları istifler |
| `ac2fJobUsage` | Teklif kesilince tabakadan kullanılan parçayı `u1350x1000` dosyasıyla işaretler |
| `ac2fJobOpen` | Müşteri klasörünü Gezgin'de açar |
| `ac2fJobSetup` | StockTrack adresi, anahtar, gün klasörü, ayraç |
| `ac2fJobHistory` | Müşterinin eski işleri: klasörü ya da CorelDRAW dosyasını açar |
| `ac2fJobDay` | İşin gün klasörü: bugün ya da önceki bir gün; export'un sorduğu seçimi sonradan değiştirir |
| `ac2fCheck` | Seçimi (ya da sayfayı) export öncesi kontrol eder |
| `ac2fCutTime` | Seçimin kesim süresi ve saatlik ücretle fiyatı |
| `ac2fCutArea` | Seçimin m² alanı ve m² ücretiyle fiyatı |
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
`build\` altındaki on yedi `.bas` dosyasını içe aktarın.

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

### Uzun panel: plakalara bölme

4 m'lik tabela, 5 cm kat, 3,20 m plaka? Açınım (4,10 m + birleşimler)
plakadan uzunsa makro **plakaları ve derz payını sorar**, bölme
seçenekleri sunar ve her plaka için ayrı panel çizer:

```
5cm out          kat
3200 j30         plaka 3200, birleşim derz payı 30   (3200 2440 j20: iki boy)
1                1. seçenek: 3200 + 960   (2: 2080 + 2080)  ya da kendi boylarınız: 3000 1160
```

```
plaka 1   kesim 3200 = 50 (kat) + 3120 yüz + 30 (derz)
plaka 2   kesim  960 = 30 (derz) + 880 yüz + 50 (kat)
                                    -------------------
                                    yüzler toplamı 4000
```

Baştaki ve sondaki plakada üç kenar **kat** (50), birleşim kenarı **derz
payı** (30); ortadaki plakada iki paralel kenar 50, birleşim kenarları 30.
Derz payları ve katlar düşülünce **net ölçü** bitmiş ölçüye eşit çıkar ve
raporda yazılır. Ayrıntı, giriş biçimleri ve çizim düzeni:
[docs/kullanim.md › Uzun panel](docs/kullanim.md#uzun-panel-plakalara-bölme).

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

1. Konturlar poligona açılır (eğrinin **gerçek** noktaları okunur, bkz.
   [Eğri örnekleme](#eğri-örnekleme))
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

Parçaları seçin, `ac2fNest` çalıştırın. Parçalar **gerçek konturlarıyla**
yerleşir, sınırlayıcı kutularıyla değil: içbükey parçalar birbirine
geçer, istenirse küçük parçalar büyüklerin deliklerine girer, çizili bir
artık (her şekilde) tabaka olarak kullanılır.

Önce **ayar penceresi** açılır; **Preview** yerleşimi hesaplayıp
**önizleme penceresinde** çizer: tabakalar ve parçalar gerçek şekilleriyle,
yanında rapor. Belgede henüz hiçbir şey taşınmamıştır. **Apply**
yerleştirir, **Change settings** geri döner, **Cancel** hiçbir şeyi
değiştirmez. **Nest now** önizlemesiz, hemen yerleştirir.

Ayar penceresinde:

- **Tabaka listesi** — her satır kullan/kullanma, ad, en, boy, **adet**
  (`0` = gerektiği kadar). Yukarıdan aşağı sırayla kullanılır; `Qty for all`
  hepsine aynı adedi yazar. Müşteri işinin stoktan aldığı tabaka listede
  yoksa en üste eklenir; parçalarla birlikte seçilen `offcut ...` adlı
  şekiller olduğu yerde, kendi şekliyle tabaka olur. Sık kullanılan
  ölçüler (3050 x 2050, 2440 x 1220 ...) tek tıkla eklenir.
- **Ayarlar** (iki sekme: *Main*, *More*) — parça aralığı, dört kenar payı, artık kenar payı, dönüş
  adımı, ayna, kopya adedi, *iç şekiller dıştakiyle gitsin*, *küçük
  parçalar deliklere girsin*, başlangıç köşesi, satır/sütun yönü,
  yerleştirme stratejisi, hassasiyet, optimizasyon süresi, sonuçtaki
  gruplama (yok / parça / tabaka / hepsi), tabaka çerçevesi, ardından
  başlangıç noktası. Her ayarın üzerine gelince açıklaması çıkar.
- **Hazır ayarlar** — `Laser 3mm`, `CNC 6mm` gibi adlarla kaydet / yükle.

`ac2fNestReport` aynı pencereleri açar. `ac2fNestAgain` pencere açmadan
son ayarlarla hemen yerleştirir.

### Algoritma

Her parça kare ızgaraya dökülür; bir kare, parçanın **yarım boşluk kadar
büyütülmüş** hâlinin herhangi bir noktası oraya düşebiliyorsa dolu
sayılır. Karesi çakışmayan iki parça arasında bu yüzden en az bir boşluk
vardır — şekil ne olursa olsun. Tabaka satırları "buradan itibaren kaç
kare boş / dolu" diye tutulur, böylece bir aralığın boş olup olmadığı tek
okumayla, dolu çıkarsa nereye atlanacağı da aynı okumayla bulunur.

Her parça, izin verilen her dönüşte en alçak (sonra en soldaki) boş yere
konur ve üst kenarı en alçakta kalan dönüş seçilir. Birkaç sabit sıra ve
iki strateji denendikten sonra, süre dolana kadar en iyi sıra rastgele
değiştirilip tekrar denenir; **en az parça dışarıda, sonra en az
malzeme** olan yerleşim kalır.

Dönüşler CorelDRAW'da gerçekten yapılıp kontur yeniden okunur ve parça
sonunda sınırlayıcı kutusuyla taşınır; nesnenin dönme merkezi nerede
olursa olsun sonuç aynıdır. Aynı parçanın kopyaları aynı ızgarayı
paylaşır ve neredeyse bedavaya yerleşir.

Motor (`ac2fNestCore`) CorelDRAW nesnesine dokunmaz. `tools/test_nest.py`
onu Python'a çevirip gerçek geometriyle sınar (boşluk, kenar payı,
delikler, çizili tabaka, adetler, iç içe geçen L'ler kutu yerleşimine göre
%19 daha az malzeme). `tools/test_corel.py` ise `ac2fNest` ve `ac2fStart`'ı
bir CorelDRAW taklidi üzerinde uçtan uca çalıştırır.

## Başlangıç noktaları

Lazer ya da router'ın her parçaya nereden girdiğini ayarlar. Seçin,
`ac2fStartPoints` çalıştırın; konum sayısal tuş takımı gibidir
(`7` sol üst, `8` üst orta, `9` sağ üst, `4` / `6` kenar ortaları, `5`
merkez, `1` / `2` / `3` alt). Konum her konturun kendi kutusuna (`E`),
nesnenin kutusuna (`O`) ya da bütün seçime (`S`) göre ölçülür; `X` ile
tam noktaya düğüm eklenir (şekil değişmez), `N` en yakın düğümü kullanır.
Yön: `W` dış saat yönünde + delikler ters, `C` tersi, `A` / `Z` hepsi bir
yöne, `K` olduğu gibi. `M` işaret çizer. Örnek: `8 S W`.

## Bükme sırası (`ac2fBending`)

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

## Eğri örnekleme

Plotter, bükme sırası, kesim süresi ve orta hat, eğrileri düz adımlara
bölerken **eğrinin gerçek noktalarını** CorelDRAW'dan okur
(`Segment.GetPointPositionAt`). Hiçbir şey uzunluk ve kirişten *tahmin
edilmez*: S kıvrımı, ilmek, sivri uç, el yazısı ya da script yazı tipi
tasarımdaki gibi çıkar.

- Her adım, eğriyi çeyrek, orta ve dörtte-üç noktalarında yoklar; eğri
  `Curve tol` (ayar 41, varsayılan 0,05 mm) uzaklığın yarısından fazla
  sapıyorsa adım ikiye bölünür. Her segment en az 4 adımla ve yaklaşık
  8 mm'de bir adımla başlar; adımdan küçük bir ayrıntı yoklamalar arasında
  gizlenemez.
- İki bağımsız denetim vardır: her nokta segmentin iki ucuna yetişebilir
  olmalı (`|AP| + |PB| ≤ uzunluk`; gerçek bir eğrinin hiçbir noktası bunu
  aşamaz) ve bitmiş çokgenin boyu CorelDRAW'ın bildirdiği segment boyuna
  eşit olmalı. Atlanmış bir ilmek çokgeni kısa bırakır; segment 4 kat sık
  yeniden okunur.
- CorelDRAW nokta vermiyorsa ya da saçma yanıt veriyorsa (`438`, ya da
  uca yetişmeyen nokta) segment, **bir kopya üzerinde** `Segment.AddNodeAt`
  ile, parçalar kirişlerine `Curve tol`'un çeyreği içine girene dek
  ikiye bölünür. Kanıtlanmış sınır: uzunluğu `L`, kirişi `c` olan bir
  eğri kirişten en çok `√(L² − c²) / 2` uzaklaşır.
- Sınama: `python3 tools/test_plot.py` dosyadaki HPGL'i geri okur ve
  gerçek Bezier eğrileriyle karşılaştırır (S, ilmek, sivri uç, 60+ rastgele
  karalama). En kötü sapma 0,055 mm; eski dairesel yay tahmini aynı
  şekillerde **3 ile 67 mm** arasında saparak "el yazısı gibi" bozuk
  çıktı veriyordu.

## Plotter'a gönderme

Seçimi seçin, `ac2fPlotSend` çalıştırın, adresi onaylayın. **Export yok,
düzeltilecek dosya yok.**

```
192.168.1.100:9100
```

### Yön: nasıl görüyorsanız öyle

**Ayar gerekmez: kesim, ekranda gördüğünüz gibi çıkar.** Bu plotter her işi
kendiliğinden **90 derece sola** çevirip işliyor; dosya bu yüzden gönderilmeden
önce **90 derece sağa** çevrilir (`Rotate` varsayılanı **270**), ikisi birbirini
götürür ve iş malzemeye ekranda durduğu gibi oturur. Genişlik ve yükseklik
yer değiştirmez; kenar payı dönmeden sonra uygulandığı için iş yine köşeye
oturur.

Başka bir makinede farklı bir dönüş gerekirse `Rotate` ayarı (42) vardır:
`0` işi çevirmeden gönderir, 90 saat yönünün tersine, 270 saat yönüne çevirir.

### Kesim sırası ve bitiş

Nesneler **Object Manager'daki sırayla**, en üstteki önce gider; yolda
hiçbir şey yeniden sıralanmaz (`Send order`, ayar 66). Bükme sırasıyla
kesmek için önce [Bükme sırası](#bükme-sırası-ac2fbending) ile Object
Manager'ı dizin, sonra gönderin.

İş bitince bıçak kalkar ve kafa **işin en üst noktasında** bekler
(`Park`, ayar 67; `0` orijin, `2` olduğu yer).

### Dönük ya da aynalı geliyorsa: yön sihirbazı (ana menü 27)

Paket döndürmüyor; makinenin X ekseni besleme yönünde olabilir. Deneme
yanılma yerine **`ac2fPlotCalibrate`** (ana menü `27`) tek kesimle bulur:

1. Test L'si **hiçbir döndürme ve aynalama olmadan** gider; makinenin tek
   başına ne yaptığını gösterir.
2. Çıkan parçaya bakıp sekiz resimden numarasını yazarsınız. Resimler
   L'nin sekiz yönüdür ve kodda geometriden üretilip yazıyla tarif edilir
   ("köşe sol-altta, uzun kol YUKARI, kısa kol SAĞA"). Parçaya makinenin
   önü (size yakın kenar) altta kalacak biçimde bakın; tabakayı ters
   çevirmeyin.
3. `Rotate` ve `Mirror` buna göre **tersine** ayarlanır; L bu ayarlarla
   yeniden gider.
4. "Doğru" derseniz bitti (`Rotate`/`Mirror` kalıcı). Değilse gördüğünüz
   şekil ayarı bir kez daha düzeltir; en çok 4 kesim.

Yarıda bırakırsanız ya da gönderim başarısız olursa önceki ayarınız geri
yüklenir. Bu plotter için sihirbaza gerek yoktur (varsayılan ayar zaten
doğrudur); başka bir makine için vardır.

Elle yapmak isterseniz adresi yazarken sonuna yön ekleyin, kalıcı olarak
hatırlanır:

```
192.168.1.200:8080 r270
```

Ters yöne döndüyse `r90`. Aynalı geldiyse `m1`. Geri almak için `r0 m0`.

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

`ac2fPlotTest` (ana menü 16) gerçek işin kullandığı **aynı** yazıcı,
gönderici ve yönlendirme adımından 40 × 80 mm'lik bir **L** gönderir:

```
   |
   |
   |___     uzun bacak SOLDA, ayak SAĞA, 80 mm uzun yön
```

| Sonuç | Anlamı | Ne yapmalı |
|---|---|---|
| Doğru şekil, temiz kesti | Her şey yolunda | — |
| **Yan yatmış, ters ya da ayna görüntüsü** | Makine eksenleri takas ediyor ya da bir ekseni aynalıyor | **Yön sihirbazı** (ana menü `27`); elle: `Rotate = 270`, ters olursa `90`, `Mirror = 1`/`2` |
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

## Müşteri işi (StockTrack)

Müşterinin dosyasını açın, `ac2fJobExport` çalıştırın: müşteriyi
(StockTrack'te arayarak) ve plakayı (stoktan, müşterinin kendi malı
önce) sorar, DXF'i günün klasörüne otomatik adla yazar.

```
17-10-2026--cumartesi\ozturk-reklam-ltd-sti\exp01--pleksi.dxf
                                             001-pleksi-1350x2000--2pcs--2800um-akpolimer-dokme-beyaz\001
                                                                                                     \002
                                                                                                     \u1350x1000
                                             002-pleksi-1350x2000--1pcs--2800um-akpolimer-dokme-siyah\001
```

- Gün klasörü **CDR Auto Renamer**'ın ayarından okunur; aynı klasöre
  yazılır. İş önceki bir günde başladıysa (teklif 22'sinde, onay
  25'inde) o günün klasöründen mi devam edileceği, bugün yeni klasör mü
  açılacağı sorulur.
- Aynı tür/ebat/kalınlık/marka/varyant/renk **istiflenir**: `--2pcs--` →
  `--3pcs--`, `003` eklenir. Farklı olan yeni numaralı klasöre gider.
- Klasör adında plakanın tam ölçüsü yazar. Teklif kesildikten sonra
  `ac2fJobUsage`, plakanın yalnız bir kısmı kullanıldıysa klasöre
  `u1350x1000` adlı bir dosya koyar.
- 4000 mm'den uzun plakalarda her sayfa klasörüne `p001`, `p002`.
- Klasör adları küçük harf ve boşluksuzdur (`ozturk-reklam-ltd-sti`);
  eski adla açılmış klasörler bulunmaya devam eder.
- Teklif StockTrack'te **işe dönüştürülünce** takım yolu klasörlerini
  dükkân bilgisayarındaki [stock-track-runner](https://github.com/ac2f/stock-track-runner)
  kendiliğinden açar (aynı adlarla). O işler için iş menüsü `4`'ü ayrıca
  çalıştırmayın; eklenti runner'ın işaretini görürse uyarır.
- Müşteri, plaka, teklif ve eski iş bir **liste penceresinde** seçilir:
  yazdıkça daralır, StockTrack'te bulunanların hepsi listelenir (tek sonuç
  da sorulur), çift tıkla ya da Enter ile seçilir.
- `ac2fJobFolders` sayfa sayısını Nesting'in plaka çerçevelerinden okur.

Kurulum: StockTrack › Ayarlar › Corel entegrasyonu'ndan anahtar alın,
iş menüsü `7` ile yapıştırın. Ayrıntı:
**[docs/kullanim.md](docs/kullanim.md#13-müşteri-işi--stocktrack-ile-otomatik-adlandırma)**

## Kontrol ve fiyat

- **Export öncesi kontrol** (`ac2fCheck`, iş menüsü `9`): eğriye
  çevrilmemiş yazı, bitmap, üst üste iki çizgi (iki kez kesilir), nest
  plakasının kenarından taşan parça, export'a karışan plaka çerçevesi.
  Export DXF bunu her seferinde kendisi yapar ve yalnız kesimi bozan
  hatada sorar.
- **Süre ile fiyat** (`ac2fCutTime`, `10`): kesim boyu × paso ÷ kesim
  hızı + konturlar arası boşta gezme ÷ boşta hızı + kontur × dalma
  süresi; saatlik ücret ve "iş en az X dakikadan başlar".
- **m² ile fiyat** (`ac2fCutArea`, `11`): seçim kutusu, parçaların
  kutuları ya da nest plakaları; m² ücreti ve asgari m².
- Ücretler, asgariler ve makine hızları StockTrack'te (*Ayarlar › Kesim
  fiyatlandırması*); teklifte süreyle fiyat da aynı değerleri kullanır.

## Ayarlar ve profiller

**Tüm ayarlar tek sayfada.** Ana menüden `19` ile açılır; 55 ayar
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
python3 tools/lint.py          # içe aktarmadan önce statik denetim + VBA derleyici kuralları
python3 tools/test_nest.py     # nesting motoru, gerçek geometriyle
python3 tools/test_corel.py    # nesting + başlangıç noktası, CorelDRAW taklidinde uçtan uca
python3 tools/test_plot.py     # plotter: HPGL'i geri okuyup gerçek Bezier eğrileriyle karşılaştırır, yön sihirbazı
python3 tools/test_panel.py    # ACP panel: tek panel + plakalara bölme (kat/derz payı, net ölçü, seçenekler)
python3 tools/test_windows.py  # nesting pencereleri, PowerShell 7 ile (AC2F_PWSH)
python3 tools/ps2vba.py        # nest_window.ps1 / nest_preview.ps1 değişince ac2fNest.bas'a gömer
python3 tools/skeleton.py      # büyük düzenlemelerden sonra kontrol akışı karşılaştırması
sh tools/build.sh              # build/ altına CRLF kopya üret (isteğe bağlı)
```

Kaynak **saf ASCII**'dir; hiçbir Türkçe karakter içermez. Bu yüzden
`.bas` dosyaları her makinede kod sayfası dönüşümü olmadan doğrudan içe
aktarılabilir; klonda ve GitHub ZIP'inde CRLF satır sonuyla gelir.
Ayrıntı: **[docs/gelistirme.md](docs/gelistirme.md)**

## Gereksinimler

- CorelDRAW Graphics Suite (VBA desteği kurulu olmalı — X6 ve sonrası)
- Windows

## Sürüm

1.22.0 — bkz. [CHANGELOG.md](CHANGELOG.md)
