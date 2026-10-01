# Değişiklik Günlüğü

Bu proje [Semantic Versioning](https://semver.org/lang/tr/) kullanır.

## [1.8.0] - 2026-10-01

### Sorun: makine yolları geziyor ama kesmiyor

Belirti çok şey söylüyor: `PD` koordinatları **hareket olarak**
çalışıyor, ama bıçak inmiyor. Yani `PD` belirteci cihaza ulaşmıyor ya da
bıçak kuvveti sıfırlanmış. İki olası sebep bulundu ve ikisi de giderildi.

**1. Çok uzun `PD` komutları.** Üretilen dosyada komut başına 40
koordinat çifti vardı; 180 noktalı bir daire için **402 karakterlik**
satırlar. Giriş tamponu ya da komut ayrıştırıcısı bunu kaldıramayan bir
kesici satırın sonunu düşürür; sonraki `PD` bozuk komutun parçası olarak
yutulur, kaleme inme emri hiç ulaşmaz ve makine şekilleri bıçak yukarıda
dolaşır — **tam olarak görülen belirti.**

**2. `IN;` panel ayarlarını siliyor.** `IN` cihazı açılış
varsayılanlarına döndürür; pek çok kesicide bu, panelden ayarlanan bıçak
kuvvetini ve takımı da siler.

### Eklendi

- **`ac2fPlotTest`** (ana menü 13) — gerçek işin kullandığı **aynı
  yazıcı ve göndericiden** 50 mm'lik bir kare gönderir. Üç farklı sorunu
  bir dakikada ayırır:

  | Sonuç | Anlamı |
  |---|---|
  | Temiz kesti | Makine ve bağlantı iyi; sorun gerçek işin dosyasında → `PD pairs = 1` |
  | Gezdi ama kesmedi | Makine hiç kesmiyor → bıçak kuvveti/derinliği, ya da `Preamble = 2` |
  | Hiç kıpırdamadı | Ulaşmadı → adres, port, kablo |

- Üç yeni ayar:
  - **`PD pairs`** (varsayılan **1**) — komut başına koordinat çifti.
    1 ile her komut bir düzine karakter; hiçbir kesici zorlanmaz.
    Dosya %38 büyüyor, ki bu önemsiz.
  - **`Preamble`** (varsayılan 1) — `1 = IN;SP1;PA;`, `2 = SP1;PA;`
    (IN yok, panel ayarları korunur), `0 = hiçbiri`.
  - **`Send delay`** (varsayılan 0) — kilobayt başına duraklama. Yalnız
    tamponu küçük, akış denetimi olmayan ethernet–seri dönüştürücüler
    için; bayt düşerse aynı belirti görülür.

### Değişti

- **Varsayılan `PD pairs` 40'tan 1'e indirildi.** Dosya biraz büyüyor,
  uyumluluk belirgin şekilde artıyor.

## [1.7.1] - 2026-09-29

### Düzeltildi

**Derleme hatası: `Syntax error`.** VBA'nın ayrılmış sözcükleri değişken
adı olarak kullanılmıştı. Derleyici hangi belirtecin suçlu olduğunu
söylemediği için satırı işaretleyip "Syntax error" demekle yetiniyor.

| Nerede | Eski | Yeni |
|---|---|---|
| `ac2fPlot`, `ac2fCenterline` | `sgn` | `turnDir` |
| `ac2fCenterline` | `width` | `strokeW` |
| `ac2fCenterline` | `base`, `spc` | `rowBase`, `spCount` |
| `ac2fCore`, `ac2fBoxLetter`, `ac2fCenterline` | `mid` | `mp` |

Ayrıca **`ac2fSettings`'te gizli bir hata**: `ac2fProfileLoad` içinde
`val` adlı yerel değişken `Val()` kütüphane fonksiyonunu gölgeliyordu,
yani aynı yordamdaki `Val(val)` çağrısı bozuktu. `sVal` yapıldı.

### Eklendi

- `tools/lint.py` artık **VBA ayrılmış sözcüklerini ve kütüphane
  fonksiyonu gölgelemesini** denetliyor. Bildirim örüntüsü
  (`<ad> As <tip>`) taranır; `For Output As #1` gibi `Open` deyimleri ve
  kendi ad alanı olan `Type` alanları hariç tutulur. Bu sınıftaki
  hatalar bir daha CorelDRAW'a kadar gitmez.

## [1.7.0] - 2026-09-29

### Eklendi

- **`ac2fNest`** — parçaları plakalara yerleştirir (ana menü 10).
  Her kenardan ayrı pay, parçalar arası boşluk ve döndürme adımı
  ayarlanabilir. `ac2fNestReport` hesaplayıp raporlar, hiçbir şeyi
  oynatmaz.
- Dokuz yeni ayar: dört kenar payı, parça aralığı, dönme adımı,
  plaka en/boy, plakalar arası boşluk.

### Algoritma

Sınırlayıcı kutular üzerinde skyline bottom-left. Parçalar en uzun
kenara göre sıralanır; her parça için tüm açılar ve skyline basamakları
denenir, **skyline'ı en alçak bırakan** yerleşim seçilir.

Skor seçimi ölçülerek yapıldı. "En düşük y" seçimi döndürmeyi zararlı
hâle getiriyordu:

| Set | En düşük y | Skyline yüksekliği |
|---|---|---|
| Eşit kutular | −%8,3 | **+%8,1** |
| Karışık | +%1,6 | **+%7,2** |
| Uzun şeritler | +%5,8 | **+%17,6** |
| Dar-uzun | −%18,5 | **+%5,0** |
| 450×160 | +%7,6 | **+%13,5** |

Testlerin kapsamadığı bir karışımda yine ters tepebileceği için,
döndürme açıkken iş iki kez yerleştirilir (döndürmeli ve döndürmesiz) ve
iyi olan tutulur. Döndürme asla malzeme kaybettirmez.

Beş sette çakışma ve sınır denetimi yapıldı: temiz.

### Değişti

- **Ayar sayfası gruba göre sayfalandı.** 42 ayar tek `InputBox`
  istemine (≈1024 karakter) hiçbir düzende sığmıyordu. Artık bir grup
  gösteriliyor, `#n` ile geçiliyor; **ayar numaraları genel** olduğu için
  `N=value` her sayfadan çalışır — hâlâ tek menü, soru zinciri değil.
  En büyük sayfa 524 karakter.
- Değer sütunu 7 karaktere çıktı ve sayfada kompakt biçim kullanılıyor
  (binlik ayırıcı yok, tam sayıda ondalık yok). Önceden `1220.00` sütuna
  sığmayıp `1220.0` diye kırpılıyordu.
- `ac2fAbout` artık tüm ayarları dökmüyor — 42 satır mesaj kutusuna da
  sığmıyor. Grup adları ve sayıları gösteriliyor.
- `tools/lint.py` artık en büyük **sayfayı** ölçüyor, tüm listeyi değil.

## [1.6.1] - 2026-09-28

### Netleştirme

Sorulan iki nokta koda bakılarak doğrulandı:

- **Döndürme yoktu ve yok.** `ac2fPlot` belge X'ini plotter X'ine, belge
  Y'sini plotter Y'sine eşler; arada döndürecek bir süzgeç yok.
- **CorelDRAW PLT export'u kullanılmıyor.** Modülde tek bir `Export`
  çağrısı yok; HPGL doğrudan geometriden yazılıyor. Önce dışa aktarıp
  sonra gönderme adımı hiç olmadı.

### Eklendi

- **`Rotate` ayarı (33)** — çeyrek tur, varsayılan **0 = gördüğünüz
  gibi**. İş yine de dönük çıkıyorsa sebep makinedir: çoğu kesicide X
  ekseni malzeme besleme yönünde uzar. 90 / 180 / 270 bunu telafi eder.
  Döndürme, kenar payı uygulanmadan **önce** yapılır, böylece iş her
  yönde köşeye oturur. 90'ın katı olmayan değer en yakın çeyreğe
  yuvarlanır.

### Düzeltildi

- Gönderim başarılı olduğunda **geçici dosya siliniyor**. Önceden
  `%TEMP%\ac2f_plot.plt` geride kalıyordu. Başarısızlıkta bilerek
  bırakılır, elle gönderilebilsin diye.

### Doğrulama

200×300 mm, sol altı işaretli bir iş dört yönde de ölçüldü:

| Rotate | Boyut | Çizim min | İşaret |
|---|---|---|---|
| 0 | 200 × 300 | (200, 200) | sol-alt |
| 90 | 300 × 200 | (200, 200) | sağ-alt |
| 180 | 200 × 300 | (200, 200) | sağ-üst |
| 270 | 300 × 200 | (200, 200) | sol-üst |

Dört yönde de kenar payı korunuyor.

Ayar sayfası: grup adları kısaltıldı ve başlık sıkıştırıldı
(`SETTINGS [profil]`), 33 ayarla 960 karakter.

## [1.6.0] - 2026-09-28

### Eklendi

- **`ac2fPlot`** — seçimi HPGL olarak üretip plotter'a ağdan doğrudan
  gönderir (`ac2fPlotSend`, ana menü 10). Export adımı ve sonradan
  düzeltilecek dosya yok.

  Kaydırma **HPGL yazılırken geometriden** uygulanır, dolayısıyla
  ayrıştırılacak metin yoktur. Eğriler kendi yarıçaplarına göre
  adımlanır (sehim `s²/(8R)`), böylece hata her yerde tolerans altında
  kalır.

- `ac2fPlotSave` — HPGL'i dosyaya yazar, göndermez.
- `ac2fPlotFixSend` — var olan bir `.plt`'yi normalize edip gönderir.
- İki yeni ayar: kenar payı (mm), eğri toleransı (mm). Plotter adresi
  ayar sayfasında değil; gönderirken sorulur ve hatırlanır.
- `ac2fCore`: paylaşılan geometri yardımcıları `ac2fSolveTheta`,
  `ac2fAtan2`, `ac2fWrapAngle`.

### Gönderim yolu

VBA'nın socket'i yok. Baytlar kısa bir PowerShell betiği üzerinden
`System.Net.Sockets.TcpClient` ile gider — `Declare`, 32/64 bit sorunu ve
OCX kaydı gerekmez. Betik günlük yazar, makro geri okur; bağlantı
reddedilirse gerçek sebep gösterilir. Ham TCP akışıdır, telnet protokolü
anlaşması yapılmaz.

### Kaynak betiğe göre bir düzeltme

`.plt` normalize ederken çizim alanı **yalnız çizim komutlarından**
hesaplanır: her `PD`, ve ardından `PD` gelen `PU`'lar. Sondaki `PU0,0;`
kalem park komutudur. Ölçüldü:

| Dosya | Tüm PU/PD | Yalnız çizim |
|---|---|---|
| Çizim negatif, park var | doğru | doğru |
| **Çizim pozitif, park var** | **yanlış** — iş 4200'de kalır | doğru — 200'e oturur |
| Çizim pozitif, park yok | doğru | doğru |

Yani hata yalnız çizim tamamen pozitif koordinatlardayken ortaya çıkar;
o durumda normalizasyon sabit bir kaydırmaya dönüşür ve malzeme boşa
gider.

### Değişti

- Ayar sayfası çerçevesi sıkıştırıldı (`[GRUP]`, tek satır altlık, uzun
  profil adı kırpılır) — plotter grubuna yer açmak için. 32 ayarla sayfa
  970 karakter.
- `tools/lint.py` sembol taramasında artık dize içeriğine bakmıyor;
  `"ac2f_plot.plt"` gibi dosya adları tanımsız sembol sanılıyordu.

### Doğrulama

| Ölçüt | Sonuç |
|---|---|
| Eğri toleransı 0,5 / 0,1 / 0,05 / 0,01 mm | ölçülen sehim hepsinde tolerans altında |
| 200×300 mm kare, 5 mm kenar payı | çizim min tam (200,200), max (8200,12200) |
| Negatif koordinatlı girdi | çıktıda hepsi pozitif, min (200,200) |
| Park komutu tuzağı | üç dosya biçiminde ayrı ayrı ölçüldü |

## [1.5.0] - 2026-09-16

### Eklendi

- **`ac2fCenterline`** — dolu tasarımı tek çizgiye düşürür
  (`ac2fCenterline`, ana menü 9). Pleksi üzerine 6 mm bıçakla neon LED
  şerit kanalı açmak için.

  İki mod, istendiği gibi ayrı opsiyonlar olarak:
  - `ac2fCenterlineJoined` — tüm seçim tek bölge; **değen harfler bağlı
    kalır**, el yazısı tek sürekli çizgi olur
  - `ac2fCenterlineSeparate` — her nesne kendi başına

  Boru hattı: kontur açma (dairesel yay modeli) → şekil başına çift-tek
  dolgu + şekiller arası birleşim → Zhang-Suen inceltme → fazlalık
  merdiven pikseli temizliği → zincir izleme → çıkıntı budama → uç
  uzatma → yumuşatma → Douglas-Peucker. Her zincir tek polyline olarak
  kırmızı çizilir ve gruplanır.

- Beş yeni ayar: çözünürlük, sadeleştirme toleransı, yumuşatma,
  en az dal, uçları uzat.

### Değişti

- **Ayar sayfası iki sütuna geçti.** 30 ayar tek sütunda okunabilir
  hiçbir genişlikte 1024 karakterlik `InputBox` sınırına sığmıyordu.
  Izgaradan birim sütunu kaldırıldı (birim `?N` ve raporlarda duruyor);
  mm dışı birimler etikete taşındı. Sayfa 952 karakter.
- `tools/lint.py` sayfa ölçümü yeni düzene göre güncellendi.

### Doğrulama

Algoritma önce Python'da prototiplendi ve **orta hattı bilinen**
şekillerde ölçüldü; VBA sürümü aynı işlem sırasıyla yeniden sınandı:

| Test | Sonuç | Gerçek | Hata |
|---|---|---|---|
| Dalgalı şerit (1 yol) | 470,8 mm | 471,0 mm | %0,0 |
| T kavşağı (3 dal) | 279,3 mm | 280,0 mm | %0,2 |
| İki ayrı harf (2 yol) | 161,0 mm | 160,0 mm | %0,6 |
| Yay örnekleme, tam çember | — | — | 0,000000000 mm |
| Yay örnekleme, gerçek bezier | — | — | %0,036 |

Yol boyunca bulunan üç hata:

- İki örtüşen poligonu **tek** çift-tek taramasına sokmak, örtüşen
  bölgeyi delik yapıyordu — şekil başına dolgu alıp OR'lamak gerekti.
  Bu olmadan birleşik mod çalışmıyordu.
- Merdiven basamaklarındaki fazlalık pikseller sahte kavşak üretip
  tek şeridi 305 parçaya bölüyordu.
- Uç uzatma yönünü son pikselden almak, düz uç kapağının köşesine
  sapıyordu (12,1 mm hata). Yönü bir şerit kalınlığı boyunca
  ortalamak hatayı 1,8 mm'ye indirdi.

## [1.4.0] - 2026-09-10

### Eklendi

- **`ac2fPanel`** — alüminyum kompozit (ACP) paneller için V derz
  yerleşimi (`ac2fPanelGroove`, ana menü 8).

  Seçilen dörtgenden düz plakayı ve dört derz çizgisini üretir. Çizgiler
  plakayı boydan boya geçer, **saat yönünde** ve **kesim sırasıyla**
  oluşturulur; her birinin başlangıç noktası operatörün beklediği uçtadır:

  | # | Derz | Başlangıç | Boy |
  |---|---|---|---|
  | 1 | Sol | alttan | tam yükseklik |
  | 2 | Üst | soldan | tam genişlik |
  | 3 | Sağ | üstten | tam yükseklik |
  | 4 | Alt | sağdan | tam genişlik |

  Çizgiler tek tek adlandırılır, böylece sıra Nesne Yöneticisi'nde
  görünür.

- İki yön: `out` (seçim bitmiş ölçü, plaka her yandan bir derz büyür)
  ve `in` (seçim plaka, derzler içeri kaçar). Ölçü `5cm`, `50mm` ya da
  düz sayı (mm) olarak girilebilir.
- Renk kodu: siyah = plaka dış hattı, mavi = V derz. Derz çizgileri bir
  grupta, o grup da plaka dörtgeniyle birlikte dış bir grupta toplanır.
- Üç yeni ayar: derz ölçüsü, yön, kaynak dörtgeni koru/sil.
- `tools/lint.py` artık **ayar sayfasının uzunluğunu** ölçüyor ve VBA'nın
  ~1024 karakterlik `InputBox` sınırı aşılırsa hata veriyor. 25 ayarla
  sayfa 942 karakter; etiket sütunu 18'den 16'ya indirildi (18'de 1007
  karakter çıkıyordu, sınıra fazla yakın).

### Doğrulama

Yerleşim Python'a taşınıp ölçüldü:

| Ölçüt | Sonuç |
|---|---|
| 100×200 + 5 `out` → plaka | 110×210 |
| Derz kesişimleri | tam **100×200** (asıl ölçü) |
| Başlangıç noktalarının dönüş yönü | −360° = **saat yönü** |
| Çizgi boyları | 210 / 110 / 210 / 110 |
| `in` modu, 100×200 + 5 | plaka 100×200, katlanmış 90×190 |

## [1.3.0] - 2026-09-10

### Eklendi

- **`ac2fSettings`** — paketin tüm ayarları artık **tek sayfada**
  (`ac2fSettings`, ana menü 8). 22 ayar numaralı listelenir; `N=value`
  ile değiştirilir, `?N` ile ayrıntılı açıklaması okunur. Birden çok
  ayar tek satırda değiştirilebilir (`11=0.8 16=0.1`), numara yerine
  ayar adı da kullanılabilir.
- **Ayrıntılı açıklamalar.** Her ayar için ne işe yaradığını, neyi
  etkilediğini, büyütüp küçültünce ne olacağını ve tipik değerleri
  anlatan uzun metin. `?N` ile açılır.
- **Profiller** — 22 ayarın tamamı bir ad altında kaydedilir, yüklenir,
  silinir (`P` komutu ya da `ac2fProfiles`). Profil, kayıt defterinde tek
  bir `key=value|...` dizesi olarak saklanır (22 ayar için ~570 karakter).
- **`ac2fBoxLetterStripProfile`** — profili seçip, istenen değeri
  **yalnız o çalıştırma için** değiştirip şeridi çizer. Geçici değerler
  kayıtlı ayarlara yazılmaz, sayfada `*` ile işaretlenir.
- **Geçici değer katmanı** (`ac2fCore`) — `ac2fSetOverride` /
  `ac2fClearOverrides`. Her okuma `ac2fGetNum` üzerinden gittiği için
  geçici değer tüm modüllerce görülür.
- Malzeme ön ayarları profil menüsüne taşındı (`M`).

### Değişti

- **`ac2fLedSettings` ve `ac2fBoxLetterSettings` kaldırıldı.** Yerlerini
  tek `ac2fSettings` sayfası aldı; sıralı soru zinciri yok.
- Ana menü yeniden düzenlendi: 6 profille çizim, 7 rapor, 8 ayarlar.
- `ac2fAbout` artık tüm ayarları ve etkin profili tek listede gösterir.

### Bilinen sınır

**Fare üzerine gelince açıklama (tooltip) yok.** Bunun için UserForm
gerekir; `.frm` dosyası yanında ikili bir `.frx` ister ve bu ikili metin
tabanlı bir depoda güvenle üretilip doğrulanamaz — bozuk bir `.frx` içe
aktarmada çöker. Açıklamalar bunun yerine `?N` ile, tam metin olarak
verilir.

### Doğrulama

| Ölçüt | Sonuç |
|---|---|
| Ayar sayfası uzunluğu | 866 karakter (VBA InputBox sınırı ~1024) |
| Profil dizesi | 22 ayar için ~570 karakter |
| Geometri/hesap yordamları | kontrol akışı değişmedi |
| ASCII dışı karakter | yok |

Düzeltilen üç kullanılabilirlik hatası: profil seçicide boş girişte
sonsuz özyineleme; `L`/`S` ile **başlayan** profil adlarının komut
sanılıp kırpılması (`Letters3mm` → `etters3mm`); profil adı sorulurken
İptal ile boş adın ayırt edilmemesi.

## [1.2.0] - 2026-09-10

### Kırıcı değişiklik

Arayüz tamamen İngilizceye çevrildi ve **makro adları değişti**. Araç
çubuğuna eklediğiniz kısayollar bozulur, yeniden atanmalıdır. Beş modülün
tamamı yeniden içe aktarılmalıdır — geçiş tablosu:
[`docs/kurulum.md`](docs/kurulum.md#sürüm-110dan-120ya--kırıcı-değişiklik)

`ac2fPack` (ana menü) adı değişmedi.

**Ayarlar korunur.** Kayıt defteri anahtarlarının değerleri bilerek
değiştirilmedi; bunlar kullanıcıya görünmeyen iç tanımlayıcılardır ve
yeniden adlandırmak kayıtlı her ayarı sessizce sıfırlardı.

### Değişti

- Tüm diyalog, rapor ve ayar metinleri İngilizce. Kod yorumları da
  İngilizceye çevrildi.
- Kaynak artık **saf ASCII**. Windows-1254 dönüşümü zorunlu olmaktan
  çıktı; `.bas` dosyaları her makinede doğrudan içe aktarılabilir.
  `tools/build.{ps1,sh}` yalnız CRLF satır sonu üretir, kolaylık içindir.
- Belgeler Türkçe kalmaya devam ediyor.

### Eklendi

- `tools/skeleton.py` — her yordamın kontrol akışı anahtar sözcüklerini
  sırayla basar. Büyük çaplı yeniden adlandırma ve çeviride mantık
  kaybını yakalamak için `diff` ile karşılaştırılır.

### Doğrulama

Bu sürüm beş modülün tamamının yeniden yazılmasıydı, bu yüzden çeviri
mekanik olarak doğrulandı:

| Ölçüt | Sonuç |
|---|---|
| Yordam sayısı | 62 → 62 |
| Kontrol akışı iskeleti | 62 yordamın tamamı **birebir aynı** |
| Sayısal sabitler | 395 literalin tamamı aynı |
| ASCII dışı karakter | yok |

## [1.1.0] - 2026-09-10

### Eklendi

- **`ac2fBoxLetter`** — kutu harf yan bordürü açınımı ve derz yerleşimi
  (`ac2fKutuHarfSerit`, `ac2fKutuHarfRapor`, `ac2fKutuHarfAyarlar`).
  Malzeme kalınlığı, K faktörü ve esneklik oranına göre açınım boyunu
  çıkarır, eğriliğe göre derz konumlarını hesaplar ve kesime hazır düz
  şeritleri renk kodlu olarak sayfaya çizer (mavi = eğri derzi,
  pembe = köşe derzi, kırmızı = rulo kesim yeri). Delik konturları
  ayrı işaretlenir ve açınım işareti ters çevrilir.
- Malzeme ön ayarları: alüminyum (ince/kalın), galvaniz, paslanmaz.
  Bunlar ölçülmüş değil, kalibrasyon başlangıcıdır.
- Ana menüye 5-8 numaralı girdiler; `ac2fHakkinda` artık her iki ayar
  grubunu da gösterir.

### Değişti

- `tools/lint.py` yeniden yazıldı: VBA satır devamlarını (` _`) birleştirir
  ve yorum ayıklaması artık dize farkındalıdır (`"3'lü"` yorum başlatmaz).
  Satır devamı işareti **boşluk + alt çizgi** olarak doğru tanınır; aksi
  hâlde `CAPTION_` gibi alt çizgiyle biten tanımlayıcılar devam sanılıyordu.

### Geometri notu

Kutu harf modülü bezier kontrol noktası okumaz. Her segment dairesel yay
kabul edilir ve kiriş/yay oranından dönüş açısı çözülür; böylece yalnızca
düğüm konumu ile segment uzunluğu yeter. Doğrulama sonuçları:

| Ölçüt | Sonuç |
|---|---|
| Dairesel yayda yarıçap | tam |
| Gerçek kübik bezier yayda yarıçap | %0,05 hata |
| Toplam açınım boyu | tam (Steiner); segment işaretinden bağımsız |
| Ara derz konumu sapması | en kötü 0,06 mm (3 mm kalınlık, derin loblu kontur) |

## [1.0.0] - 2026-09-03

İlk sürüm. Paket `ac2f pack` adıyla, `ac2f` ön ekli modül ve makro
isimlendirmesiyle yeniden kuruldu.

### Eklendi

- **`ac2fCore`** — ortak çekirdek: ayar yönetimi (kayıt defteri), yerelden
  bağımsız sayı ayrıştırma, birim biçimleme ve geometri ölçüm motoru.
  Gruplar ve PowerClip içerikleri özyinelemeli taranır; eğri olmayan
  nesneler geçici kopya üzerinden eğriye çevrilerek ölçülür.
- **`ac2fLength`** — vektörlerin çevresindeki çizgilerin toplam uzunluğunu
  ölçer (`ac2fUzunlukOlc`), sonucu sayfaya metin olarak ekleyebilir
  (`ac2fUzunlukEtiketle`).
- **`ac2fLedModule`** — 3'lü LED modül yerleşimi için adet hesabı
  (`ac2fLedModulHesapla`, `ac2fLedHizliHesap`). Modül, LED, toplam güç,
  güvenlik paylı güç ve güç kaynağı adedini çıkarır.
- **`ac2fMenu`** — ana menü (`ac2fPack`), ayarlar (`ac2fAyarlar`),
  hakkında (`ac2fHakkinda`) ve ayar sıfırlama (`ac2fAyarlariSifirla`).
- `tools/build.ps1` ve `tools/build.sh` — kaynağı VBE'nin beklediği
  Windows-1254 + CRLF biçimine çeviren yapı betikleri.
- `tools/lint.py` — içe aktarmadan önce çalıştırılabilecek statik denetleyici.

### Notlar

- Bu sürüm sıfırdan yazılmıştır. Kaynak alınan `gdg_measureIt_2023.gms`
  dosyasının içeriği sıkıştırılmış/korumalı olduğundan kodu çıkarılamadı;
  yalnızca davranışı (kontur uzunluğu ölçme) örnek alındı.
  Ayrıntı için [`docs/gelistirme.md`](docs/gelistirme.md).

### Planlanan

- Modül konumlarının sayfaya daire olarak çizilmesi (yerleşim önizlemesi).
- Alan tabanlı hesap yöntemi (dolu yüzeyli kutu harfler için).
- UserForm tabanlı tek pencerelik arayüz.
- Kutu harf şeridinin rulo boyuna göre ayrı parçalar hâlinde çizilmesi
  (şimdilik tek parça çizilip kesim yeri işaretleniyor).
