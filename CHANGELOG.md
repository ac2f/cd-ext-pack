# Değişiklik Günlüğü

Bu proje [Semantic Versioning](https://semver.org/lang/tr/) kullanır.

## [1.24.0] - 2026-10-29

### Eklendi: ACP panel plakalara bölme (`ac2fPanelGroove`, `ac2fPanelSplit`)

Tabela plakadan uzunsa (4 m'lik tabela, 5 cm kat, 3,20 m plaka) panel artık
birleşimli birkaç plakaya bölünür. Açınım **plaka uzunluğundan** büyükse
`ac2fPanelGroove` kendisi sorar; her seferinde sormasını isteyen
`ac2fPanelSplit` (ana menü `28`) çalıştırır.

- **Plaka boyları ve derz payı** tek satırda girilir (`3200 j30`,
  `3200 2440 j20`, `3.2m j3cm`); sonraki sefer hazır gelir. `0` = bölme.
- **Seçenekler sunulur**: en az plakayla başlayıp en az fire verenler öne
  gelir (ör. `3200 + 960` ve eşit yüzlü `2080 + 2080`). Bir numara seçilir
  ya da her plakadan **kesilecek boy** elle yazılır (`3200 960`,
  `3050:3000 3200:*`, `*` = kalanı bu plaka alsın). Toplam tutmazsa fark
  söylenip yeniden sorulur.
- **Pay düzeni**: baştaki ve sondaki plakada üç kenar kat (ör. 50), birleşim
  kenarı derz payı (ör. 30); ortadaki plakada iki paralel kenar kat, diğer
  iki kenar derz payı. Katlar ve derz payları düşülünce **net ölçü bitmiş
  ölçüye eşit** çıkar ve raporda yazılır (`faces add up to 4000 mm`).
  `in` yönünde seçim toplam açınımdır, net ölçü ondan hesaplanır.
- **Çizim**: her plaka kendi paneli (dikdörtgen + dört derz, eski sıra ve
  başlangıç noktalarıyla), `ac2f ACP panel 1 of 2 3200x1100`. Plakalar
  üst üste binmez: geniş panelde kaynağın altına, dik panelde sağına dizilir.
  Bölme uzun kenar boyunca yapılır. 100 mm'den dar son plaka olmasın diye
  önerilerde son plakaya en az 100 mm yüz bırakılır.
- Tüm sorular, hiçbir şey çizilmeden **önce** sorulur: iptal belgeyi
  değiştirmez. Her soru ve rapor 1000 karakterin altında kalır (VBA'nın
  kutu sınırı).
- Plakaya sığan panelde davranış **eskisi gibidir** (tek soru, tek panel).

### Düzeltildi

- Panel makrosunda hata olursa gösterilen ileti artık gerçek hatadır;
  eskiden `EndCommandGroup` çağrısının hatası asıl hatanın üstüne yazılıyordu.
- Nesne adlarındaki ölçüler binlik ayracı olmadan yazılır
  (`ac2f ACP panel 4100x1100`; eskiden yerel ayara göre `4.100x1.100`).
- Ana menü kısaldı (`8` satırına `28` eklendi, iki satır kısaltıldı) — InputBox
  sınırı için.

### Sınama

- `tools/test_panel.py` (yeni, 100+ denetim): tek panel eskisi gibi; 4 m /
  kat 50 / derz payı 30 / plaka 3200 örneği `3200 + 960`, her kenarın payı,
  yüzler toplamı = 4000; üç plaka, `in` yönü, dik panel, özel boylar ve hatalı
  girişler, iptal, birim geri yükleme, soru uzunlukları.
- `tools/vbapy.py`: boş dize (Cancel) bir `String` değişkene atanınca artık
  boş kalıyor; harness iptali gerçeğe uygun taklit ediyor.

## [1.23.1] - 2026-10-28

### Düzeltildi: plotter işi 90 derece sola çevirip kesiyordu

Bu plotter her işi kendiliğinden 90 derece sola çeviriyor. `Rotate`
varsayılanı artık **270** (dosya 90 derece sağa çevrilerek gönderilir):
ikisi birbirini götürür ve kesim ekranda gördüğünüz gibi çıkar. Ayar,
sihirbaz ya da adres satırına yön yazmak gerekmez.

- Gönderim raporu, adres sorusu ve test L'si artık yön sihirbazından söz
  etmez; "yön denetlenmedi" uyarısı kalktı. Sihirbaz (`ana menü 27`) başka
  bir makine için duruyor, kimseye dayatılmıyor.
- Hız ve kuvvet bu sürümde **değişmedi**: dosyaya hız (`VS`) ya da kuvvet
  (`FS`) komutu yazılmıyor, makinenin kendi ayarı geçerli.
- Test: `tools/test_plot.py` makinenin 90 sola çevirmesini taklit eder ve
  varsayılan ayarla L, 300 x 100 dikdörtgen ve el yazısı eğrinin ekrandaki
  gibi çıktığını doğrular.

## [1.23.0] - 2026-10-28

### Düzeltildi: plotter eğrileri bozuk kesiyordu ("el yazısı gibi")

`ac2fBending`, plotter dosyasının noktalarını üreten düzleştirme kodu,
her segmenti **dairesel yay** sanıyordu: yay açısını uzunluk ve kirişten
çözüyor, yayın hangi yana şiştiğini komşu kirişlerin yönünden *tahmin*
ediyordu. Gerçek bir bezier daire dilimi değildir. S kıvrımı, ilmek,
sivri uç, tutamaç boyları birbirinden çok farklı el yazısı ya da script
yazı tipleri yanlış tarafa şişiyor ya da biçim değiştiriyordu. Ölçüm
(gerçek Bezier eğrilerine karşı): **3 ile 67 mm sapma**.

Artık eğrinin **gerçek noktaları** CorelDRAW'dan okunur
(`Segment.GetPointPositionAt`) ve `Curve tol` içinde kalana dek yoklanıp
bölünür. En kötü ölçülen sapma **0,055 mm**. Ayrıntı: README › Eğri
örnekleme. Aynı düzeltme bükme sırasını, kesim süresini ve orta hattı da
düzeltir (hepsi aynı geometriyi kullanıyordu).

- Yedek yol: CorelDRAW nokta vermiyorsa ya da saçma yanıt veriyorsa
  `Segment.AddNodeAt` ile bir **kopya** parçalara bölünür (kanıtlanmış
  elips sınırıyla); sonuç yine tolerans içindedir.
- **Dosyada sıfır uzunluklu `PD` kalmaz**: kapalı konturların son düğümü
  başlangıçla aynıydı ve bir yinelenen nokta yazılıyordu.
- **Orta hat**: açık alt yolun son düğümü artık eklenir.

### Eklendi: yön sihirbazı (`ac2fPlotCalibrate`, ana menü 27)

Makine işi 90 derece dönük ya da aynalı çiziyorsa deneme yanılma yok:
test L'si hiçbir dönüşüm olmadan gider, çıkan parçanın resmini seçersiniz,
`Rotate` ve `Mirror` tersine kurulur, L yeniden gider ve doğrulanır.
Sekiz olası makine yönünün hepsinde en çok 2 kesimde doğru ayara varır
(testlenmiş); yanlış resim seçilse bile doğrulama kesimi düzeltir. İptal ya
da gönderim hatasında önceki ayar geri yüklenir.

### Eklendi: gönderim raporunda tasarım denetimi

Dosyadaki boyut, CorelDRAW'ın aynı seçim için bildirdiği boyutla
karşılaştırılır (`Design (CorelDRAW)` / `Check`). Yön sihirbazı çalışmadıysa
rapor bunu hatırlatır.

### Düzeltildi: `ac2fCenterline` derlenmiyordu

İki modül düzeyi değişken (`m_pi`, `m_used`) yordamların ortasında
bildirilmişti; VBA bunu "Only comments may appear after End Sub, End
Function, or End Property" ile reddeder. Bildirimler üste taşındı.
`tools/lint.py` artık bu hatayı yakalıyor (ilk yordamdan sonraki her
modül düzeyi bildirim hata).

### Araçlar

- `tools/test_plot.py` (yeni): HPGL'i geri okuyup gerçek Bezier eğrileriyle
  karşılaştırır; döndürme/aynalama (16 birleşim), yön sihirbazı (8 makine),
  yedek yollar, tolerans ayarı, rastgele karalamalar.
- `tools/cdrmock.py`: `Collection`, `Page.Shapes`, `FileLen`, yay uzunluğuna
  göre ofset (CorelDRAW gibi), çöp/yok yanıt kipleri, ağ gönderimini
  yakalama.

## [1.22.0] - 2026-10-23

### Takım yolu klasöründe belirteç

**Toolpath folders** (iş menüsü `4`) klasör adının sonuna bir belirteç
ekler; StockTrack Runner da aynısını yapar:

```
001-pleksi-1220x2440--2pcs--2800um-akpolimer-dokme-beyaz--F21-10-2026.merkez.pleksi-beyaz
```

- `F21-10-2026`: projenin bulunduğu gün klasörünün tarihi (iş önceki bir
  günde başladıysa o gün),
- `merkez`: malzemenin sahibi — müşterinin kendi malıysa müşteri klasörü
  (`ozturk-reklam`), dükkânın stoğuysa dükkânın **klasör kodu**
  (StockTrack › Ayarlar › Corel entegrasyonu; eklenti bağlantı denemesinde
  ve klasör açarken StockTrack'ten okur, çevrimdışıyken son bilineni),
- `pleksi-beyaz`: tür-renk.

İstifleme belirteci yok sayar (klasör ilk açıldığı belirteçle kalır,
belirteci olmayan eski klasöre istiflerken eklenir) ama **sahibi farklı
tabakayı aynı klasöre koymaz**. Klasör CNC'ye kopyalansa da kimin, hangi
günün işi olduğu adından okunur. Runner'ın açtığı `malzeme-F…` klasörleri
(kesimsiz malzeme satışı) takım yolu klasörü sayılmaz. Used parts
belirteçli klasörleri de bulur.

Test aracı `vbapy` artık `Like`, `InStrRev` ve `Format$(3, "000")`'ı da
çalıştırıyor (iş modülünün klasör adı yordamları ilk kez testte).

## [1.21.0] - 2026-10-22

### Düzeltildi: nesting makroları derlenmiyordu

`ac2fNest` modülü ayar ekranını `ac2fSettings` diye çağırıyordu. O yordam
kendi modülüyle aynı adı taşıdığı için VBA bu adı **modül** sanıyor ve
`ac2fNest` modülünün tamamı **"Compile error: Expected variable or
procedure, not module"** ile derlenmiyordu; nesting makrolarının hiçbiri
çalışmıyordu. Aynı hata ana menüde (`10` nesting, `9` orta hat, `19`
ayarlar, `21` iş) ve iş menüsünde (`9` kontrol, `10` süre, `11` m²) de
vardı. Bu çağrılar artık `Modül.Yordam` biçiminde (`ac2fSettings.ac2fSettings`).

Aynı turda bulunanlar:

- **Birim**: sayfa ölçüsü ve çizili artıkların ölçüsü belge birimiyle
  okunuyordu (inç belgede "Page 8 x 11"). Artık her şey mm'ye geçildikten
  sonra okunuyor.
- **Kilitli nesneler** CorelDRAW taşımadığı için işi yarıda kesiyordu;
  artık yerinde bırakılıyor ve raporda `Locked` satırında sayılıyor.
  Yalnız kilitli nesne seçildiyse ne yapılacağı söyleniyor.
- **Gruplama** seçim yerine `ShapeRange` ile: Optimization açıkken
  CorelDRAW seçimi her zaman güncel tutmuyor.
- **Ayar tablosu** (`ac2fBuildTable`, 53 KB) dört yordama bölündü; VBA
  derlenmiş hâli 64 KB'ı aşan yordamı "Procedure too large" ile reddeder.
- **Satır sonları**: `.bas` dosyaları çalışma kopyasına ve GitHub ZIP'ine
  CRLF olarak çıkıyor (VBE içe aktarması bunu bekler).

### Eklendi: önizleme penceresi ve daha pratik ayar penceresi

`ac2fNest` artık şu akışla çalışıyor: **ayarlar → Preview → önizleme →
Apply** (ya da *Change settings* ile geri dönüp tekrar).

- **Önizleme penceresi**: kullanılan her tabaka ve yerleşen her parça
  gerçek şekliyle çizilir (delikler boş, çizili artık kendi şekliyle).
  Solda çizim, sağda tabaka listesi (hepsi ya da tek tabaka) ve rapor.
  `Apply` yerleştirir; `Change settings` belgeyi olduğu gibi bırakıp ayar
  penceresine döner; `Cancel` hiçbir şeyi değiştirmez. Çizim, uygulanınca
  parçaların gideceği yerle birebir aynıdır.
- **Ayar penceresi**: ayarlar iki sekmede (*Main*: aralık, kenar payları,
  dönüş, kopya, süre, iç/dış, delikler; *More*: geri kalanı). Sık
  kullanılan tabaka ölçüleri (3050 x 2050, 2440 x 1220 ... ve sayfa ölçüsü)
  tek tıkla eklenir. Düğmeler: **Preview** (Enter), **Nest now** (önizlemesiz),
  **Cancel**.
- Önizlemeden sonra ayrıca mesaj kutusu çıkmaz; rapor penceredeydi.

### Geliştirme

- **tools/vbacheck.py**: gerçek VBA derleyicisinin yakalayıp test
  derleyicisinin kaçırdığı hatalar — modül adıyla aynı yordamı çağırmak,
  ByRef tür uyuşmazlığı, `Name(a, b)` ile Sub çağrısı, büyük/küçük harf
  farkıyla aynı ad, CorelDRAW nesnesinde olmayan üye, Set'siz nesne
  ataması, tanımsız etiket, 1023 karakterlik satır, 24'ten fazla satır
  devamı. `tools/lint.py` bunu da çalıştırır.
- **tools/test_windows.py** + **tools/winstub.cs**: iki pencere betiği
  Linux'taki PowerShell 7'de, WinForms/GDI+ taklidiyle gerçekten
  çalıştırılır; sınama düğmelere basar, kutulara yazar, çıkışı denetler.
  Windows PowerShell 5.1'de olmayan yapılar ve `$PID` gibi otomatik
  değişkenlere atama da aranır.
- tools/test_corel.py: önizleme (Apply, Change settings + Cancel),
  kilitli nesne ve gerçek pencerelerle uçtan uca senaryolar.
- tools/ps2vba.py iki betiği de gömer, ~8 KB'lık yordamlara böler.

## [1.20.0] - 2026-10-21

### Eklendi: yerel sunucuya geçiş (iş menüsü 7, `5=`)

Dükkânda StockTrack'in yerel kopyası (replika) çalışıyorsa adresi Job
setup'ta `5=` ile yazılır. Ana sunucu cevap vermezse (ya da önündeki ağ
geçidi 502/503/504 derse) müşteri, malzeme ve teklif listeleri yerel
sunucudan gelir; başlıkta `(local server)` yazar. Geçişten sonra 5 dakika
önce yerel sunucuya sorulur. Gerçek hatalar (anahtar geçersiz ...) yerel
sunucuyla örtülmez; ikisi de yoksa eski çevrimdışı kopya kullanılır. `L`
yalnız yerel sunucuyu dener. Runner aynı ayarı okur.

- Test derleyicisine (tools/vbapy.py) `Now`, `Date`, `TimeSerial`,
  `DateSerial`, `Day/Month/Year`, `Weekday`, `Choose`, `ChrW$` ve `Date`
  tipi; sahte Corel'e (tools/cdrmock.py) MSXML HTTP nesnesi eklendi.
  tools/test_corel.py'de yeni senaryo: `t_job_local_server`.

## [1.19.0] - 2026-10-21

### Değişti: nesting artık gerçek şekille (`ac2fNest`, yeni `ac2fNestCore`)

Parçalar sınırlayıcı kutularıyla değil **gerçek konturlarıyla**
yerleşiyor. Her parça kare ızgaraya dökülüyor; bir kare, parçanın yarım
boşluk kadar büyütülmüş hâli oraya düşebiliyorsa dolu sayılıyor, böylece
karesi çakışmayan iki parça arasında boşluk her şekilde korunuyor.
İçbükey parçalar birbirine geçiyor (on L parçada kutu yerleşimine göre
%19 daha az malzeme), istenirse küçük parçalar büyüklerin deliklerine
giriyor, çizili artık her şekliyle tabaka oluyor.

- **Pencere** (PowerShell): tabaka listesi (kullan / ad / en / boy /
  adet, `0` = gerektiği kadar, sıra = kullanım sırası, hepsine aynı
  adet), bütün ayarlar açıklamalı, **hazır ayarlar** (kaydet / yükle /
  sil). Açılamazsa kısa metin sorusu.
- **Yeni ayarlar**: hassasiyet, optimizasyon süresi, iç şekiller
  dıştakiyle gitsin, küçük parçalar deliklere girsin, sonuç gruplaması
  (yok / parça / tabaka / hepsi), ayna, başlangıç köşesi, satır/sütun
  yönü, strateji, kopya adedi, tabaka çerçevesi, artık kenar payı,
  ardından başlangıç noktası.
- **Optimizasyon**: birkaç sabit sıra ve iki strateji, sonra süre
  dolana kadar rastgele iyileştirme; en az parça dışarıda, sonra en az
  malzeme olan yerleşim kalır.
- Dönüşler CorelDRAW'da gerçekten yapılıp kontur yeniden okunuyor, parça
  sınırlayıcı kutusuyla taşınıyor; dönme merkezi nerede olursa olsun
  sonuç doğru. Aynı parçalar tanınıp ızgarayı paylaşıyor.
- Eğri boyunca nokta okuması olmayan bir sürümde her eğri parçası, o
  boydaki her eğriyi içine alan bir kutuyla kaplanıyor (gevşek ama asla
  çakışmaz).
- `ac2fNestAgain` (menü 22): pencere açmadan son ayarlarla.
  `ac2fNestMarkSheets` (menü 26): seçimi çizili artık olarak adlandırır.
- Tabaka çerçeveleri parça gruplarına girmiyor, eskisi gibi
  `ac2f sheets` grubunda; müşteri işi "hepsi" grubunun içindeki parçaları
  da görüyor.

### Eklendi: başlangıç noktaları (`ac2fStart`, menü 23-25)

Her yolun ilk düğümü sayısal tuş takımı düzeninde seçilen yere taşınıyor
(`7 8 9 / 4 5 6 / 1 2 3`): her konturun kendi kutusuna, nesnenin
kutusuna ya da bütün seçime göre. Tam nokta modunda gerekirse oraya
düğüm ekleniyor (şekil değişmez). Kesim yönü: dış saat yönünde + delikler
ters, tersi, hepsi bir yöne ya da olduğu gibi. Açık yollar yakın uçtan
başlıyor. İşaretler (`ac2f start points` katmanı) çizilip kaldırılabiliyor.
Nesting sonrası otomatik uygulanabiliyor.

### Düzeltildi

- **Hakkında** kutusu MsgBox'ın ~1024 karakter sınırını aşıp kesiliyordu;
  her özellik tek satıra indi.

### Geliştirme

- `tools/vbapy.py`: VBA'yı Python'a çeviren küçük derleyici (tür
  dönüşümleri, sınır denetimli diziler, ByRef, `On Error`, nesne üyeleri).
- `tools/cdrmock.py`: CorelDRAW nesne modelinin sınama taklidi.
- `tools/test_nest.py`, `tools/test_corel.py`: motor ve uçtan uca sınamalar.
- `tools/nest_window.ps1` + `tools/ps2vba.py`: pencere betiği tek kaynakta.

## [1.18.0] - 2026-10-20

### Eklendi: işin gün klasörünü değiştirme (`ac2fJobDay`, iş menüsü 12)

Export "önceki günde devam mı, bugün yeni klasör mü" diye sorduğunda
verilen cevap işle saklanıyor ve aynı gün bir daha sorulmuyordu; yanlış
cevap verilince (özellikle **bugün** seçilince) aynı gün geri dönmenin
yolu yoktu. İş menüsü `12` bugünü ve müşterinin klasörü olan önceki
günleri (export ve klasör sayılarıyla, şu anki gün işaretli) listeler;
seçilen gün işe kaydedilir. Soru kutusu, export bildirimi ve menü özeti
bunu hatırlatır.

## [1.17.1] - 2026-10-18

### Eklendi: StockTrack Runner uyarısı

Teklif işe dönüştürülünce klasörleri artık dükkân bilgisayarındaki
[stock-track-runner](https://github.com/ac2f/stock-track-runner) açabiliyor
(aynı adlarla). Runner müşteri klasörüne `.stocktrack\<teklif>.done`
işareti bırakıyor; **Toolpath folders** (iş menüsü 4) bu işareti görürse
onay penceresinde hangi teklifler için klasörlerin zaten açıldığını
söylüyor, aynı işin tabakaları ikinci kez istiflenmesin.

## [1.17.0] - 2026-10-18

### Değişti: klasör adları küçük harf ve boşluksuz (`ac2fJob`)

Müşteri klasörü artık `Ozturk Reklam Ltd. Sti` değil
`ozturk-reklam-ltd-sti`: küçük harf, Türkçe harfler sadeleşmiş, harf ve
rakam dışındaki her dizi tek `-`. StockTrack aynı kuralla gönderiyor;
`=ad` ile yazılan ve son seçilenler listesinden gelen adlar da aynı
kurala uyuyor. Ayraç ayarı artık boşluk kabul etmiyor.

Eski adla açılmış klasörler bulunmaya devam ediyor (aynı okunan ad
eşleşir): önceki günde başlayan iş, eski işler listesi ve kullanılan
parça işareti o klasörleri görür; yalnız yeni açılan klasörler yeni adı
alır.

### Eklendi: liste penceresi

Müşteri bağlarken, plaka seçerken, teklif seçerken (*Used parts*) ve
eski işlerde numara yazılan kutu yerine bir pencere açılıyor: üstteki
kutuya yazdıkça liste daralıyor (Türkçe harfe duyarsız), çift tıklama ya
da oklar + Enter ile seçiliyor.

- StockTrack'te aramanın **bütün sonuçları listeleniyor; tek sonuç da
  kendiliğinden seçilmiyor.** Listede kalmayınca Enter (ya da F3) aramayı
  sunucuda yapıyor; *Use as typed* yazılanı olduğu gibi kullanıyor.
- Eski işlerde *Open its CorelDRAW file* ve *Carry this job on here*
  düğmeleri var.
- Pencereyi PowerShell çiziyor (Windows'la gelir; Plot da kullanıyor).
  PowerShell engelliyse eski numaralı soru kendiliğinden açılıyor; orada
  da tek sonuç artık listelenip soruluyor. Ayarlarda **List window** `0`
  her zaman numaralı soruyu kullanır.

## [1.16.0] - 2026-10-17

### Eklendi: StockTrack'e ulaşılamazken son listeler (`ac2fJob`)

StockTrack uzak bir sunucuya taşınınca bağlantı kopabilir. Eklenti artık
StockTrack'ten gelen her listeyi bilgisayarda saklıyor
(`%LOCALAPPDATA%\ac2f\stocktrack`, sunucu + anahtar başına ayrı).

- Sunucu cevap vermezse (ya da önündeki ağ geçidi 502/503/504 dönerse)
  aynı istek kayıtlı kopyadan cevaplanıyor; listelerin başlığında
  `(offline, saved <tarih>)` yazıyor.
- Daha önce aranmamış müşteri/plaka araması, o güne dek görülen bütün
  müşteriler ve işletme plakaları (en fazla 1000'er) içinde yapılıyor.
  Müşterinin kendi plakaları ortak listeye girmiyor; çevrimdışıyken başka
  müşteriye önerilmesin.

### Değişti: dükkân başına anahtar

StockTrack çok dükkânlı oldu: her dükkânın kendi Corel anahtarı var ve
anahtar dükkânı belirliyor. Eklentide değişiklik gerekmiyor; bilgisayarın
çalıştığı dükkânın anahtarını girmek yeterli.

## [1.15.0] - 2026-10-17

### Eklendi: export öncesi kontrol (`ac2fCheck`)

Seçimi (ya da sayfayı) DXF'e gitmeden önce kontrol ediyor.

- **Kesimi bozar:** eğriye çevrilmemiş yazı, bitmap/OLE, üst üste iki
  çizgi (kutu ve uzunluk 0,05 mm içinde), nest plakasının kenarından
  taşan parça ya da işin plakasından büyük iş, export'a karışan Nesting
  plaka çerçevesi.
- **Bakmaya değer:** açık kontur, 0,5 mm'den küçük kırıntı, PowerClip.
- Export DXF her seferinde kontrol ediyor (`Check first`, ayar 61);
  yalnız kesimi bozan hatada "yine de export / dur ve seç / vazgeç"
  diye soruyor. Sorunlu nesneler tek adımda seçilebiliyor.

### Eklendi: süre ile fiyat (`ac2fCutTime`)

Kesim boyu ve konturların başlangıç/bitiş noktaları plotter'ın
kullandığı düzleştirmeden (`ac2fBending`). Süre = kesim boyu × paso ÷
kesim hızı + boşta gezme ÷ boşta hızı + kontur × paso × dalma süresi.
Boşta gezme, işin sol alt köşesinden en yakın komşu turuyla
hesaplanıyor (2000 konturdan sonra alan formülüyle). Paso = plaka
kalınlığı ÷ paso derinliği. Fiyat = faturalanan dk (asgari dk'dan az
değil) ÷ 60 × saatlik ücret.

### Eklendi: m² ile fiyat (`ac2fCutArea`)

Üç alan: seçim kutusu, nesnelerin kutuları toplamı, nest plakaları.
Fiyat = max(alan × adet, asgari m²) × m² ücreti.

Her iki modül ücretleri, asgarileri ve makine hızlarını StockTrack'ten
(`/corel/pricing`) okuyor, son okunanı saklıyor (çevrimdışı çalışır);
her değer o çalıştırma için değiştirilebiliyor (`f=30 n=2`, `x=4`).

### Eklendi: eski işler (`ac2fJobHistory`)

Müşterinin son bir yılda klasörü olan günleri, yeniden eskiye, export /
klasör / cdr sayısıyla listeliyor. Numara klasörü açıyor; `3 O`
klasördeki CorelDRAW dosyasını açıyor, `3 C` bu işi o klasörden devam
ettiriyor.

İş menüsü: `8` eski işler, `9` kontrol, `10` süre ile fiyat, `11` m²
ile fiyat. Gereksinim: StockTrack'in kesim fiyatlandırmalı sürümü.

## [1.14.0] - 2026-10-17

### Eklendi: iş başka bir günde başladıysa sorar

Teklif 22'sinde verilip 25'inde onaylanan iş için export, takım yolu
klasörleri ve klasörü aç artık soruyor: 22'sinin klasöründen devam mı
(Evet), bugünün altında yeni klasör mü (Hayır), vazgeç mi (İptal). Soru
eski klasörde kaç export ve takım yolu klasörü olduğunu da gösteriyor.

- İşin günü dosyadaki bağlantıyla saklanıyor; aynı gün bir daha
  sorulmuyor, ertesi gün yeniden soruluyor.
- Henüz güne bağlanmamış dosyada müşterinin son `Look back` (30) günde
  açılmış klasörüne bakılıyor.
- Başka müşteri bağlanınca gün sıfırlanıyor.
- Kullanılan parça (`ac2fJobUsage`) önce işin kendi gününe bakıyor.

Yeni ayar: `Look back` (60). Eski bağlantı kayıtları olduğu gibi okunuyor.

## [1.13.0] - 2026-10-17

### Değişti: ebatta `X` yok, kullanılan parça tekliften

Klasör adındaki ebat artık hep plakanın stoktaki tam ölçüsü; `X1220x900`
gibi kesik ölçü yazılmıyor. Bir kısmı kullanılacak plaka tam plakalarla
aynı klasöre istifleniyor.

Plakanın ne kadarının kullanıldığı teklif kesilince belli oluyor:

- **Yeni `ac2fJobUsage`** (iş menüsü `5`): müşterinin teklifini
  StockTrack'ten okuyor, her tabaka kaleminin takım yolu klasörünü
  buluyor ve plakanın yalnız bir kısmı kullanıldıysa klasöre
  `u1350x1000` adlı bir dosya yazıyor (1350 × 2000 plakanın
  1000 × 1350'si; plakanın eni önce). Tam kullanılan plakaya dosya
  konmuyor, var olan dosyaya dokunulmuyor.
- Klasör, teklif gününden iki hafta geriye kadar müşterinin gün
  klasörlerinde aranıyor.
- `Round up` ayarı (60) kaldırıldı; yalnız `X` için vardı.
- İş menüsü: `5` kullanılan parça, `6` klasörü aç, `7` kurulum.

Gereksinim: StockTrack'in `/corel/quotes` uçlarını içeren sürümü.

## [1.12.0] - 2026-10-17

### Eklendi: müşteri işi (`ac2fJob`)

Müşterinin dosyası StockTrack'teki müşteriye bağlanıyor, plaka stoktan
seçiliyor; DXF export'u ve takım yolu klasörleri bu ikisinden
adlandırılıyor.

- **Bağlama** (`ac2fJobLink`): son 9 müşteri ya da StockTrack'te arama
  (Türkçe harfe duyarsız). Bağlantı belge özelliği olarak dosyanın
  içinde, kaydedilmemiş dosyalar için kayıt defterinde de tutuluyor.
  `<gün>\<müşteri>\` klasöründen açılan dosya o müşteriyi öneriyor.
- **Plaka** (`ac2fJobMaterial`): müşterinin kendi (konsinye) malzemesi
  önce, sonra son seçilenler ve işletme stoğu. Seçilen plaka Nesting'in
  plaka ölçüsünü de ayarlıyor (`Nest sheet`).
- **Export** (`ac2fJobExport`): `<gün>\<müşteri>\exp01--pleksi.dxf`,
  numara klasördeki son export'un bir fazlası. Eksik müşteri ya da
  plaka varsa önce onu soruyor.
- **Takım yolu klasörleri** (`ac2fJobFolders`):
  `001-pleksi-1220x2440--2pcs--2800um-akpolimer-dokme-beyaz\001 ...`.
  Aynı plakalar istifleniyor (`2pcs` → `3pcs`, `003` ekleniyor); kesik
  plaka `X1220x900` olarak ayrı klasöre gidiyor; `Long sheet`'ten
  (4000 mm) uzun plakalarda her sayfa klasörüne `p001`, `p002`. Plan
  Nesting'in plaka çerçevelerinden ya da seçimin ölçüsünden öneriliyor.
- **Gün klasörü** CDR Auto Renamer'ın `config.json`'ından okunuyor
  (ana dizin, tarih biçimi, ayraç); yoksa aynı adla açılıyor.
- StockTrack'e ulaşılamazsa müşteri ve plaka yazılarak devam ediliyor.

Yeni ayar grubu `JOB` (56-60): `Long sheet`, `Export dialog`,
`Open folder`, `Nest sheet`, `Round up`. Sunucu adresi, anahtar, gün
klasörü ve ayraç metin olduğu için iş menüsü `6`'da (`ac2fJobSetup`).

Ana menüye `21` (ya da `J`) eklendi.

Gereksinim: StockTrack'in Corel uçlarını içeren sürümü
(`/api/v1/corel/...`, Ayarlar › Corel entegrasyonu).

## [1.11.0] - 2026-10-16

### Değişti: plotter Object Manager sırasıyla gönderir

`ac2fPlotSend` artık nesneleri **Object Manager'da göründükleri
sırayla**, en üstteki önce gönderiyor ve yolda hiçbir şeyi yeniden
sıralamıyor. Önceki sürüm bükme sırasını kendiliğinden uyguluyordu.

`ActiveSelectionRange` bu sırayı garanti etmediği için her seçili nesne
sayfa yığınında (`ActivePage.Shapes`, önden arkaya) `StaticID` ile
bulunuyor. Grup içinden tek tek seçilmiş nesneler sayfa yığınında
olmadığından geri kalanların ardından, seçim sırasıyla gidiyor; rapor
kaç tane olduğunu söylüyor. Grup içindeki nesneler de Object Manager
yönünde alınıyor.

Yeni `Send order` ayarı (47): `0` üstten (varsayılan), `1` alttan,
`2` bükme sırası o anda hesaplanır. Bükme sırasıyla kesmenin olağan
yolu: önce `ac2fBendOrder` (Object Manager'ı yeniden dizer), sonra
gönder.

### Eklendi: iş bitince kafa işin tepesinde bekler

Son konturdan sonra bıçak kalkıyor ve kafa **işin en üst noktasına**
gidiyor — ekranda göründüğü hâliyle sol üst köşe. Nokta işle aynı
`Rotate` ve `Mirror`'dan geçtiği için makine işi hangi yöne çevirirse
çevirsin tepede kalıyor; dört dönüşün ve aynalamanın sekiz birleşiminde
doğrulandı, kenar payını hiç etkilemiyor. `Park` ayarı (48): `1` tepe
(varsayılan), `0` orijin (eski davranış, `PU0,0;`), `2` olduğu yer.
Test L'si de aynı park adımından geçiyor.

### Eklendi: yön, adres satırından

Adres kutusu şu anki `Rotate`/`Mirror` değerini gösteriyor ve adresin
sonuna yazılan `r270`, `r90`, `m1`, `r0 m0` gibi ekleri kalıcı ayar
olarak kaydediyor. Yalnız ek yazılırsa kayıtlı adres kullanılıyor.

Paket hâlâ hiçbir şey döndürmüyor — 90 derece makinenin ekseninden
geliyor. Bu ek, düzeltmeyi ayar sayfasına girmeden tek gönderimde
yapmak için.

### Eklendi: bükme algoritması her çalıştırmada sorulur

`ac2fBendOrder`, `ac2fBendPreview` ve `ac2fBendReport` önce 1/2/3
seçimini istiyor; son seçim varsayılan olarak geliyor ve hatırlanıyor.

### Değişti

- Bükme ayarları **49-55** (iki plotter ayarı eklendiği için).
- `Order 3` uyarısı: aynı nesnedeki göbek için öneri artık harfi
  ayırmak ya da `Send order = 2`.

## [1.10.0] - 2026-10-04

### Düzeltildi: ayar dosyası içe aktarılamıyordu

`ac2fSettings.bas`, **"Too many line continuations"** ile reddediliyordu.
VBA bir mantıksal satırı en çok **25 satır birleştirmeden** kurar;
`Order` ayarının açıklaması 27 tanesini kullanıyordu.

Çare iki parçalı: açıklamanın devamını yazan `ac2fHelpMore` yardımcısı,
ve `tools/lint.py` içinde kalıcı bir denetim — sınır 20'ye çekildi,
böylece ileride bir cümle eklemek için pay kalıyor. Aynı taramada ana
menünün `InputBox` uzunluğu da denetleniyor (şu an 797 / ~1024).

### Değişti: bükme sırası kendi modülünde

Sıralama `ac2fPlot`'tan çıkıp **`ac2fBending`**'e taşındı. Plotter'a
özel bir davranış değil; aynı sıra hem makineye gönderirken hem de
çizimin kendisinde kullanılıyor.

Modül aynı zamanda **düzleştirilmiş geometrinin tutulduğu yer** oldu:
`ac2fPlot` HPGL'i onun tamponundan yazıyor. Bir toplama, bir sıralama,
iki tüketici.

Ayar grubu `BEND` (47-53). Plotter ayarları **40-46**'ya geri döndü,
yani `Rotate` yine **42**, `Mirror` **43**.

### Eklendi: `ac2fBendOrder` — sırayı nesnelere uygular

Önizlemeden ayrı olarak sırayı **çizimin kendisine** yazar:

- **İstif sırası.** İlk bükülecek nesne en üste (`Stack order`, 52 ile
  alta). Çoğu program dışarı yazarken istif sırasını kullanır ama hangi
  uçtan başladıkları değişir; ayar bunun içindir.
- **İsim.** Nesne `[007 S03] eski adı` olur: yedinci bükülecek, üçüncü
  kurulum. Aynı `S` numarası aynı şekil demektir. Tekrar çalıştırınca
  eski etiket sökülür, üst üste binmez. `Name objects` (53).
- Sayfada hiçbir nesne oynatılmaz; tek `Undo` geri alır.

Bir sınır açıkça belirtiliyor: `Order 3` göbekleri sona atar, ama bir
göbek dışıyla **aynı nesnenin** içindeyse istif sırasıyla ayrılamaz —
istif nesneleri sıralar, konturları değil. O ayrım `ac2fPlotSend`
yolunda tam olarak uygulanıyor.

### Eklendi: `ac2fBendReport`

Sırayı tablo hâlinde listeler (numara, kurulum, dış/göbek, ölçü,
uzunluk), hiçbir şeyi değiştirmez.

### Netleştirme: "düz" ve "eğri" ne demek

`Order 2` ayrımı **içinde hiç eğri olmayan** harfler (N, A, K, X, E, M,
V, W, Z, L, T) ile **içinde eğri olan** harfler (O, B, P, S, C, G, D, R,
U) arasındadır. Eğik kenar eğri sayılmaz: A'nın ve K'nin yan çizgileri
düzdür, o yüzden ikisi de ilk gruba girer. Ölçülen şey, segmentin
kirişinden en çok saptığı miktardır (`Curve limit`, 50).

### Değişti

- Ana menü: bükme `12`-`14`, plotter `15`-`18`, ayarlar `19`,
  hakkında `20`.
- `ac2fPlotOrderPreview` -> `ac2fBendPreview`.
- Sıralama ayarlarının kayıt anahtarları `PL...` yerine `BN...`; o beş
  ayar bir kereliğine varsayılanına döner.
- Konturları düz adımlara bölen tolerans tek: `Curve tol` (41). Bükme
  modülü de onu okur.

## [1.9.0] - 2026-10-03

> Bu sürüm içe aktarılamıyor (satır birleştirme sınırı). 1.10.0 kullanın.

### Eklendi: kesim sırası ve gruplama (`ac2fPlot`)

Kutu harf şeridi bükerken sıra, makinenin kaç kez kurulduğunu belirler.
Konturlar artık önce **harflere** ayrılıyor, sonra harfler bir iş
sırasına diziliyor.

**İçine alma.** Her kapalı kontur ötekilere karşı sınanıyor: ilk noktası
içine düşüyor mu, ve öteki daha büyük mü (kesişen kutu ön elemesi, sonra
tek–çift ışın testi). Hiçbir şeyin içinde olmayan kontur bir *dış*, bir
dışın içindeki *göbek*, göbeğin içindeki yine dış. Bir dış ve kapsadığı
her şey bir harf.

**Üç algoritma** (`Order`, ayar 44):

| Değer | Sıra |
|---|---|
| `0` | Belgedeki hâliyle |
| `1` *(varsayılan)* | Soldan sağa, önce üst |
| `2` | Aynısı, ama önce ovalliği olmayan harfler, sonra eğriler |
| `3` | 2 ile aynı, ama göbekler en sona (kendi aralarında aynı sıra) |

**Önce aynılar** (`Group same`, ayar 45). Bir harfe gelindiğinde ondan
sonraki aynı şekilli bütün harfler hemen arkasına çekiliyor. `NOBOBOXB`
için üç `O` peş peşe bükülüyor — sekiz harf, **dört kurulum**.

Doğrulanan çıktı:

```
1 ->  N  O O. O O. O O.  B B.1 B.2  B B.1 B.2  B B.1 B.2  X
2 ->  N  X  O O. O O. O O.  B B.1 B.2  B B.1 B.2  B B.1 B.2
3 ->  N  X  O O O  B B B  |  O. O. O.  B.1 B.2  B.1 B.2  B.1 B.2
```

**Aynı sayılma ölçütü**: kontur sayısı, uzunluk, kapsanan alan ve kutu
ölçüsü, hepsi `Match tol` (48) içinde. En ve boy yer değişebilir, yani
çeyrek tur döndürülmüş kopya da aynı sayılır. **Aynalanmış kopya
sayılmaz** — ters yöne bükülür ve kendi kurulumunu ister; işaretli
alanın işareti bunu yakalar.

**Sütunlar.** X'te üst üste binen harfler bir sütun sayılıp yukarıdan
aşağı gidiliyor. Bir satırdaki kelime için her harf kendi sütunu olur ve
sıra düz soldan sağa kalır; sayfaya satır satır dizilmiş parçalarda ise
sütun sütun aşağı inilir. Bindirme eşiği `Column overlap` (46).

**Eğrilik ölçüsü** (`Curve limit`, 47): segmentin kirişinden en çok
saptığı miktar. Dönüş açısı bilindiğinde tam sagitta; açı çözücünün
altında kalan sığ yaylarda `Sqr(24 · kiriş · (yay − kiriş)) / 8`
yaklaşımı (θ ≤ 0,1 rad'da hata %0,01). Eğri olarak çizilmiş ama düz
duran bir segment sıfır ölçer ve düz sayılır — bükme makinesinin
umursadığı şey budur.

Sıra **ekranda gördüğünüz çizim üzerinden**, `Rotate` ve `Mirror`'dan
önce hesaplanıyor; "sol" her zaman ekranın solu. Noktalar hiç
taşınmıyor, yalnız kontur dizini yeniden sıralanıyor; eksik bir
permütasyon çıkarsa sıralama hiç uygulanmıyor.

### Eklendi: `ac2fPlotOrderPreview` (ana menü 13)

Seçimi sayfada kesim sırasına göre numaralandırır — göndermeden önce
sırayı görmek için. Tek `Undo` hepsini siler.

### Değişti

- Ana menü numaraları: önizleme `13`, test `14`, dosyaya yaz `15`,
  `.plt` düzelt `16`, ayarlar `17`, hakkında `18`.
- Plotter raporuna `ORDER` bölümü eklendi: kural, harf sayısı, dış/göbek
  dağılımı ve kurulum sayısı.
- Beş yeni ayar (44-48). Toplam 51.

## [1.8.1] - 2026-10-01

### Netleştirme

`ac2fPlot` işi **birebir** gönderiyor: belge X'i plotter X'i, belge Y'si
plotter Y'si, `Rotate` varsayılanı 0. Kod kontrol edildi, döndürme
yalnızca ayar sıfırdan farklıysa uygulanıyor. 90 derece makineden
geliyor — çoğu kesicide X ekseni malzeme besleme yönünde uzar.

Çözüm `Rotate` ayarı (ayar 42): **270** ya da **90**. Hangisi
olduğunu test aşağıdaki şekille bir seferde söyler.

### Düzeltildi

- **Test şekli artık asimetrik.** 50 mm'lik kare yönlenme hakkında
  hiçbir şey gösteremiyordu: hangi yöne çevrilse aynı görünür,
  aynalansa yine karedir. Yerine **40 × 80 mm bir L** kondu — uzun
  bacak solda, ayak sağa bakar. Dört dönüşün ve üç aynalamanın hepsi
  birbirinden ayırt edilebiliyor (doğrulandı).
- **Test artık `Rotate` ve `Mirror` ayarlarını uyguluyor.** Önceden
  bunları atlıyordu, yani testin söylediği şey gerçek işte olanla
  aynı değildi. Artık ikisi de aynı yönlendirme adımından geçiyor.

### Eklendi

- **`Mirror` ayarı** (0 yok, 1 sol-sağ, 2 alt-üst, 3 ikisi). Döndürme
  aynalanmış bir ekseni düzeltemez ve bazı makineler bir ekseni
  aynalar; bu yüzden ayrı bir denetim. Döndürmeden **sonra** uygulanır.

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
