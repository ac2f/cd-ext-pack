# Geliştirme

## Projenin çıkış noktası

Paket, `gdg_measureIt_2023.gms` adlı üçüncü taraf CorelDRAW makrosunun
yaptığı işten yola çıkılarak hazırlandı: **vektörlerin çevresindeki
çizgilerin toplam uzunluğunu ölçmek.**

Kaynak dosyanın kodu **çıkarılamadı**. Dosya incelendiğinde:

- Başlık `GMS\x01` — Corel'e özgü, OLE olmayan bir kapsayıcı biçimi
- Yük kısmının entropisi ~7,9 bit/bayt — yani sıkıştırılmış ya da şifreli
- `Attribute`, `VB_Name`, `Sub` gibi hiçbir VBA imi bulunmuyor
- Ne zlib akışı ne de MS-OVBA sıkıştırma kapsayıcısı çözülebiliyor

Bu nedenle dosya korumalıdır ve içindeki VBA kaynağına ulaşılamaz.
Var olan kodu "yeniden adlandırmak" mümkün olmadığından paket **sıfırdan
yazıldı**; kaynak makrodan yalnızca *ne yaptığı* örnek alındı.

Üçüncü taraf ikili dosya, lisans durumu belirsiz olduğu için bu depoya
**dâhil edilmemiştir**.

## İsimlendirme

Her şey `ac2f` ön ekini taşır. VBA'da bütün standart modüllerin `Public`
üyeleri **tek bir küresel ad alanını** paylaşır; ön ek, aynı anda yüklü başka
GMS projeleriyle çakışmayı önler.

| Tür | Kural | Örnek |
|---|---|---|
| Proje (GMS) | `ac2fPack` | `ac2fPack.gms` |
| Modül | `ac2f` + rol | `ac2fLedModule` |
| Genel makro | `ac2f` + İngilizce eylem | `ac2fMeasureLength` |
| Sabit | `AC2F_` | `AC2F_DEF_SPACING` |
| Ayar anahtarı sabiti | `AC2F_K_` | `AC2F_K_SPACING` |
| Tip | `ac2f` + isim | `ac2fLedResult` |

Kod, yorumlar ve arayüz metinleri **İngilizce**dir; bu belgeler Türkçedir.

Ayar anahtarlarının **değerleri** (`"ModulAraligiMM"` gibi) bilerek
değiştirilmedi. Bunlar kullanıcıya hiç görünmeyen iç tanımlayıcılardır;
yeniden adlandırmak, kayıtlı her ayarı sessizce varsayılana döndürürdü.

Makro Yöneticisi'nde yalnızca **parametresiz `Public Sub`** yordamları
listelenir. Bir yordamın kullanıcıya görünmesini istemiyorsanız `Private`
yapın ya da parametre verin.

## Dosya kodlaması

Kaynak **saf ASCII**'dir. VBE'nin `File > Import File` işlevi dosyayı
sistem ANSI kod sayfasıyla okuduğu için, ASCII dışı her karakter makineden
makineye bozulma riski taşır. Bunu tamamen ortadan kaldırmak için arayüz
metinleri ve yorumlar İngilizce ve ASCII tutulur.

Bu kuralı `tools/lint.py` denetler (yorumlar dahil); elle bakmak için:

```bash
LC_ALL=C grep -n '[^ -~\t]' src/*.bas    # hiçbir şey dönmemeli
```

Satır sonu: VBE içe aktarması CRLF bekler. `.gitattributes` 1.21.0'dan
beri `*.bas text eol=crlf` der; depoda LF durur, klona ve GitHub ZIP'ine
CRLF çıkar. `tools/build.ps1` ve `tools/build.sh` aynısını elle yapar
(tek dosya "raw" indirildiğinde gerekir). Araçlar iki satır sonunu da okur.

## Denetim

```bash
python3 tools/lint.py
```

Şunlara bakar:

- `Attribute VB_Name` ilk satırda mı
- Blok dengesi: `Sub`/`Function`/`Type`/`If`/`Select`/`For`/`Do`
- `Attribute X.VB_Description` doğru yordamın hemen altında mı
- VBA küresel ad alanında çakışan yordam adı var mı
- Tanımsız `ac2f*` sembolü çağrılıyor mu

Denetleyici satır devamlarını (` _`) birleştirerek çok satırlı `If ... Then`
bloklarını doğru sayar, ve yorum ayıklaması dize farkındadır — `"3'lü"`
içindeki kesme işareti yorum başlatmaz.

> Satır devamı **boşluk + alt çizgi**dir. Yalnız `_` ile biten bir
> tanımlayıcı (`CAPTION_`) devam işareti değildir; bu ayrım gözetilmezse
> denetleyici sağlam dosyalarda yanlış alarm verir.

### Ayrılmış sözcük denetimi

VBA'nın ayrılmış sözcüğünü değişken adı yapmak derlemede yalnızca
`Syntax error` verir — hangi belirtecin suçlu olduğu söylenmez, o yüzden
CorelDRAW'da bulmak zordur. Denetleyici `<ad> As <tip>` bildirim
örüntüsünü tarayıp iki listeye bakar:

- **Ayrılmış sözcükler** — deyim anahtar sözcükleri ve `Sgn`, `Len`,
  `Mid`, `Int`, `Abs` gibi MS-VBAL ayrılmış üye adları.
- **Gölgelenen kütüphane fonksiyonları** — `Val`, `Format`, `Left`,
  `Space` gibi kodun kendi çağırdığı adlar. Aynı adda bir yerel,
  o yordamdaki çağrıyı sessizce değişkene bağlar.

İki durum hariç tutulur: `Open path For Output As #1` bir bildirim
değildir, ve `Type` alanları kendi ad alanındadır (`Name As String` bir
UDT içinde yasaldır).

### Kontrol akışı karşılaştırması

```bash
python3 tools/skeleton.py src > before.txt
# ... büyük çaplı yeniden adlandırma / çeviri ...
python3 tools/skeleton.py src > after.txt
diff before.txt after.txt
```

`skeleton.py` her yordamın içerdiği kontrol akışı anahtar sözcüklerini
sırayla basar. Tanımlayıcı adı değiştirmek, dizeleri çevirmek ve yorumları
yeniden yazmak bu diziyi **değiştirmemelidir**; değişiyorsa mantık
düşmüş ya da çoğalmış demektir. 1.2.0 İngilizceleştirmesi bu araçla
doğrulandı (62 yordamın tamamı birebir aynı).

VBA derleyicisinin yerini **tutmaz**. Asıl doğrulama CorelDRAW'da
`Debug > Compile ac2fPack` ile yapılır.

### Derleyici kuralları (`tools/vbacheck.py`)

`tools/lint.py` bunu da çalıştırır. `vbapy` kodu Python'a çevirip çalıştırır
ama gerçek VBA derleyicisi daha katıdır ve CorelDRAW'ın tür kütüphanesine
bağlıdır; bu kuralların her biri CorelDRAW'da **modülün tamamını**
derlenmez yapar:

| Kural | VBA'nın verdiği hata |
|---|---|
| Kendi modülüyle aynı adı taşıyan yordamı başka modülden adıyla çağırmak (`ac2fSettings`) | `Expected variable or procedure, not module` — `ac2fSettings.ac2fSettings` yazılmalı |
| ByRef (varsayılan!) parametreye başka türden değişken | `ByRef argument type mismatch` |
| `Ad(a, b)` ile Sub çağrısı (Call'suz) | `Expected: =` |
| Büyük/küçük harf dışında aynı iki ad (`m_pcN` dizi, `m_pcn` sayaç) | `Duplicate declaration in current scope` |
| Erken bağlı CorelDRAW nesnesinde olmayan üye | `Method or data member not found` |
| Nesneye Set'siz atama, nesne olmayana Set | çalışma hatası 91 / `Object required` |
| Tanımsız etiket, yanlış `Exit Sub/Function` | `Label not defined` ... |
| 1023 karakteri aşan satır, 24'ten fazla satır devamı | `Line too long`, `Too many line continuations` |

1.19.0–1.20.0'da ilk kural yüzünden `ac2fNest` hiç derlenmiyordu: test
derleyicisi Python'da adları düz bir tabloda çözdüğü için bunu görmüyordu.
CorelDRAW üyeleri için `vbacheck.py` içinde bilinen üyeler listesi vardır;
listede olmayan bir üye `CHECK` olarak bildirilir ve belgeyle doğrulanır.

## Mimari

```
ac2fMenu ────┬──> ac2fLength ─────┐
             ├──> ac2fLedModule ──┤
             ├──> ac2fBoxLetter ──┤
             ├──> ac2fPanel ──────┤
             ├──> ac2fCenterline ─┤
             ├──> ac2fBending ────┤
             ├──> ac2fPlot ───────┤
             ├──> ac2fNest ───────┼──> ac2fCore     (ac2fNest -> ac2fNestCore,
             ├──> ac2fStart ──────┤                  saf motor; -> ac2fStart)
             ├──> ac2fJob ────────┤   (ac2fJob, ac2fNest'in plaka
             ├──> ac2fCheck ──────┤    ayarlarını da okur/yazar)
             ├──> ac2fCutTime ────┤   (ac2fBending'in düzleştirmesini
             ├──> ac2fCutArea ────┤    kullanır; üçü de fiyatı ve
             └──> ac2fSettings ───┘    işin plakasını ac2fJob'tan alır)
```

`ac2fCore` hiçbir üst modüle bağlı değildir; kullanıcı arayüzü içermez
(yalnızca `ac2fInfo`/`ac2fWarn` sarmalayıcıları). Ölçüm mantığı tek yerdedir,
her iki özellik de onu kullanır.

### Ölçüm akışı

1. `ac2fMeasureSelection` → `ActiveSelectionRange`
2. Belge birimi geçici olarak **mm** yapılır, `Optimization` açılır
3. Bir komut grubu açılır
4. Her şekil için `ac2fCollect`:
   - **Bitmap / OLE** → atlanır (`DisplayCurve` bunlarda çerçeveyi
     döndürüp toplamı şişirebilir)
   - **Grup** → içeriğine özyinele
   - **PowerClip** → içeriğine özyinele
   - `Shape.DisplayCurve` varsa doğrudan kullanılır — belgeye dokunulmaz
   - Yoksa `Duplicate(0, 0)` + `ConvertToCurves`, ölçülür, kopya silinir
5. Her alt yolun uzunluğu ve kapalı/açık bilgisi `ac2fResult` içine yazılır
6. Geçici kopya üretildiyse komut grubu geri alınır, birim ve
   `Optimization` eski hâline döner

> **Dikkat:** `Undo` yalnızca `m_TempShapes > 0` iken çağrılır. Koşulsuz
> çağrılırsa, hiç geçici nesne üretilmediği durumda kullanıcının **kendi
> son işlemini** geri alır.

### Sonuç yapıları

`ac2fResult` alt yol uzunluklarını (`SubLenMM`) ve kapalılık bilgisini
(`SubClosed`) ayrı dizilerde taşır. LED hesabı bunlara ihtiyaç duyar:
adet, toplam uzunluktan değil **her kontur için ayrı** hesaplanır.

Diziler 64/256 kapasiteyle başlar ve dolunca ikiye katlanır.

## Kutu harf geometri modeli (`ac2fBoxLetter`)

Bu modülün tasarımını belirleyen kısıt: **CorelDRAW'dan bezier kontrol
noktalarını güvenilir biçimde okuyamıyoruz.** `Segment.GetPointPositionAt`
gibi çağrıların sürümden sürüme varlığı belirsiz, ve bu depoda VBA
derlenip test edilemiyor. Bu yüzden model yalnızca v1.0.0'da zaten
kullanılan sağlam yüzeye dayanır:

```
SubPath.Length / .Closed / .Segments / .Nodes
Segment.Length / .StartNode / .EndNode  ->  .PositionX / .PositionY
```

### Her segment dairesel yay kabul edilir

Kiriş `c` ve yay `L` bilindiğinde dönüş açısı `t`, şu bağıntıdan çözülür:

```
c / L = 2·sin(t/2) / t
```

Sağ taraf `(0, 2π)` aralığında kesin azalandır, bu yüzden ikiye bölme ile
60 adımda çözülür (`ac2fBLTheta`). Yarıçap `R = L / t`.

Yan fayda: **düzlük geometriden anlaşılır** (`c ≈ L`), `cdrLineSegment`
sabitine bağımlılık yok.

### Açınım: Steiner

Basit kapalı bir eğride toplam dönüş 2π'dir. Nötr eksen `g` kadar
ötelenince açınım `P ∓ 2π·g` olur. Segment segment yürütülür:

```
açınım_segment = L − g·m·θ̂        m = +1 dış kontur, −1 delik
açınım_köşe    =   − g·m·α̂        θ̂, α̂ = konturun kendi yönüne göre
                                    normalleştirilmiş dönüşler
```

`Σθ̂ + Σα̂ = 2π` olduğundan toplam tam çıkar.

### Dönüş işareti ve neden önemsiz olduğu

Segmentin hangi yöne büküldüğü, kiriş yönlerinin komşu düğümlerdeki
dönüşünden kestirilir. Bu bir **kestirimdir** ve derin loblu konturlarda
bazı segmentlerde yanılabilir. Ölçüldü:

| Ölçüt | Sonuç |
|---|---|
| Dairesel yayda yarıçap | tam |
| Kübik bezier yayda yarıçap | %0,05 hata |
| Toplam açınım boyu | **tam** — işaret hatalarından bağımsız |
| Ara derz konumu | en kötü **0,06 mm** (3 mm kalınlık, patolojik kontur) |

Toplam boyun işaretten bağımsız olmasının nedeni: köşe terimleri segment
terimlerini birebir dengeler, çünkü ikisi de aynı teğetlerden türer.

> **Aynı nedenle `m_Tau` bir doğrulama aracı DEĞİLDİR.** Kiriş yönü
> dizisinin toplam dönüşü basit kapalı bir çokgende yapısal olarak 2π'dir,
> yani `m_Tau` işaretler yanlışken de ±360° çıkar. Yalnız yön (saat yönü
> mü değil mi) bilgisi için kullanılır.

Tek gerçek sınır: **bir segment içinde dönüm noktası** (S kıvrımı). Model
onu tek yönlü yay sanar ve o segmentte derzi seyreltir. Kullanıcıya çözümü
belgelendi (düğüm ekleyip segmenti bölmek).

### Derz aralığı

```
s = min( R·(ağız/derinlik)·esneklik ,  √(8·R·tolerans) ,  s_max )
s = max( s , s_min )
```

İkinci terim sehim (sagitta) bağıntısıdır: `h ≈ s²/(8R)`.

### Çizim savunması

Geometri hesabı ile çizim biçimlendirmesi ayrılmıştır: renk, kalınlık ve
yazı boyutu `On Error Resume Next` altında en iyi çaba olarak uygulanır.
Bir biçimlendirme çağrısı desteklenmiyorsa şerit yine doğru çizilir.

## Ayar tablosu ve profiller (`ac2fSettings`)

Her ayar **tek bir tabloda** (`ac2fBuildTable`) durur: anahtar, grup,
etiket, birim, tür, varsayılan, alt/üst sınır ve yardım metni. Ayar
sayfası, profil kaydı ve geçici değer ayrıştırıcısı hep bu tabloyu okur.

**Yeni ayar eklemek = tabloya bir satır eklemek.** Başka hiçbir yere
dokunulmaz; sayfada, profillerde ve `?N` yardımında kendiliğinden çıkar.

### Tek sayfa ve InputBox sınırı

VBA'da `InputBox` istem metni **yaklaşık 1024 karakterle** sınırlıdır.
Sayfa bu yüzden dar biçimlendirilir (etiket 18, değer 6 karakter) ve
kısa açıklamalar sayfaya değil `?N` altına konur. 22 ayarla sayfa
**866 karakter** tutuyor; yeni ayar eklerken bu payı gözetin:

```bash
# sayfa uzunlugunu kabaca olcmek icin
python3 - <<'EOF'
rows = 22 ; print(2 + 22*32 + 160)   # ~ satir sayisi x satir genisligi
EOF
```

### İki sütunlu sayfa

30 ayar tek sütunda okunabilir hiçbir genişlikte sınıra sığmadığı için
sayfa iki sütuna geçti ve ızgaradan birim sütunu kaldırıldı (birim `?N`
ve raporlarda duruyor; mm dışı birimler etikete taşındı). 32 ayara
çıkarken çerçeve de sıkıştırıldı: `[GRUP]`, tek satır altlık, uzun
profil adı kırpılır. Sayfa 970 karakter.

### Sayfa uzunluğu denetlenir

`tools/lint.py` ayar sayfasını `ac2fSheet` ile aynı biçimde yeniden kurup
uzunluğunu ölçer ve 1000 karakteri aşarsa hata verir. Ayar eklerken bu
kendiliğinden yakalanır: 25 ayarla sayfa 942 karakter, etiket sütunu 16.
18'de 1007 çıkıyordu ve sınıra fazla yakındı.

### Geçici değer katmanı

`ac2fCore` içinde küçük bir örtme katmanı vardır:

```
ac2fSetOverride key, value      ' yalnız bellekte
ac2fGetOverride key, v          ' var mı?
ac2fClearOverrides
```

`ac2fGetNum` **önce** örtmeye bakar. Tüm modüller ayarları o fonksiyondan
okuduğu için, geçici bir değer hiçbir modüle ayrıca haber vermeden
geçerli olur. "Profili kullan ama bu seferlik bir değeri değiştir"
özelliği budur.

Örtmeler bir çalıştırmaya aittir: profille çalıştıran makro sonunda
temizler, profilsiz makrolar da başında temizler ki önceki çalıştırmadan
sızıntı olmasın.

### Profil biçimi

Bir profil, kayıt defterinde tek bir dizedir:

```
ac2fPack\Profiles\<ad> = "ModulAraligiMM=100|KHKalinlikMM=2|..."
```

22 ayar için ~570 karakter. Yüklerken tanınmayan anahtarlar sessizce
atlanır, böylece bir ayar paketten çıkarılsa bile eski profiller
yüklenmeye devam eder.

### Neden UserForm yok

Fare üzerine gelince çıkan ipucu (`ControlTipText`) bir UserForm ister.
VBE bir formu `.frm` + **ikili `.frx`** çifti olarak dışa aktarır; `.frm`
içindeki `OleObjectBlob` satırı `.frx`'e işaret eder ve tüm denetim
yerleşimi o ikilinin içindedir.

O ikiliyi metin tabanlı bir depoda elle üretmek MS-OFORMS biçimini
bayt düzeyinde doğru kurmayı gerektirir ve burada derlenip
sınanamaz — bozuk bir `.frx` içe aktarmada çöker. Bu yüzden arayüz
`MsgBox`/`InputBox` üzerine kuruludur ve açıklamalar `?N` ile tam metin
olarak verilir.

Gerçek pencere gereken yerde (nesting'in ayar ve önizleme pencereleri,
iş menüsünün liste penceresi) VBA, Windows'la gelen PowerShell'e WinForms
betiği yazıp çalıştırır ve bitmesini bekler; konuşma iki dosyayladır.
Betikler metindir, depoda durur ve `tools/test_windows.py` ile sınanır.

Form eklenecekse VBE içinde çizilip `.frm` + `.frx` **birlikte**
depoya konmalıdır.

### Sayfalı ayar ekranı

42 ayar tek `InputBox` istemine sığmıyor. Sayfa gruba göre bölündü ama
**ayar numaraları genel** kaldı: `N=value` her sayfadan çalışır, yani
kullanıcı açısından hâlâ tek menü. `tools/lint.py` artık en büyük
sayfayı ölçer.

Sayfadaki değerler `ac2fSheetVal` ile kompakt biçimlenir (binlik ayırıcı
yok, tam sayıda ondalık yok); `ac2fSettingText` tam biçimi `?N` ve
raporlar için korur. Bu ayrım olmadan `1220.00` sütuna sığmayıp
`1220.0` diye kırpılıyordu.

## ACP panel yerleşimi (`ac2fPanel`)

Dörtgenin eksene paralel sınırlayıcı kutusundan plakayı ve dört derz
çizgisini üretir. `out` modunda plaka her yandan bir derz büyür, `in`
modunda plaka seçimin kendisidir.

```
1  sol    x = X0+d      alttan -> ustten   boy H
2  ust    y = Y0+H-d    soldan -> saga     boy W
3  sag    x = X0+W-d    ustten -> asagi    boy H
4  alt    y = Y0+d      sagdan -> sola     boy W
```

Çizgiler plakayı **boydan boya** geçer; kesişimleri katlanmış ölçüyü
verir. Sıra saat yönüdür ve **oluşturma sırası kesim sırasıdır**, bu
yüzden çizgiler tek tek adlandırılır.

Başlangıç noktası `CreateLineSegment`'in ilk koordinat çiftidir; CAM
bıçağı orada daldırır, dolayısıyla çift sırası önemlidir ve
değiştirilmemelidir.

**Sınırlayıcı kutular önce toplanır.** Şekil oluşturmak ve gruplamak
seçimi değiştirdiği için, çizime başlandıktan sonra kaynak `ShapeRange`
üzerinde gezinilemez.

Gruplama `ClearSelection` + `AddToSelection` + `ShapeRange.Group()` ile
yapılır ve başarısız olursa çizgiler sayfada serbest kalır; makro bunu
raporlar, geometri kaybolmaz.

## Orta hat çıkarma (`ac2fCenterline`)

Paketin en yoğun algoritması. Önce Python'da prototiplenip **orta hattı
bilinen** şekillerde ölçüldü, sonra VBA'ya geçirildi ve aynı işlem
sırasıyla yeniden sınandı.

```
1  konturlari poligona ac   (dairesel yay modeli, ac2fBoxLetter ile ayni)
2  sekil basina cift-tek dolgu, sekiller arasi OR
3  Zhang-Suen inceltme
4  fazlalik merdiven pikseli temizligi
5  zincir izleme
6  cikinti budama
7  serbest uc uzatma
8  yumusatma + Douglas-Peucker
9  zincir basina tek polyline
```

### Neden şekil başına dolgu

İki örtüşen poligonu **tek** çift-tek taramasına sokmak, örtüşen bölgeyi
XOR'layıp **delik** yapar. Her şekli kendi içinde çift-tek doldurup
(kendi delikleri doğru çıkar) şekilleri OR'lamak gerekir. Birleşik mod
bu olmadan hiç çalışmıyordu.

### Neden merdiven temizliği

Zhang-Suen çıktısı 8-bağlantılıdır ama fazlalıksız değildir: merdiven
basamaklarındaki pikseller üç komşulu görünür ve izleme onları kavşak
sanar. Tek bir şerit **305 parçaya** bölünüyordu. Komşusu tarafından
kapsanan pikselleri atmak dereceleri 1 ve 2'ye indirir.

### Neden uç yönü kuyruktan alınır

İnceltme serbest uçları yarım kalınlık geri yer. Uzatma yönünü **son
pikselden** almak, düz uç kapağında medyal eksenin köşeye çatallanması
yüzünden yanlış yöne gider: ölçülen hata **12,1 mm**. Yönü bir şerit
kalınlığı boyunca ortalamak hatayı **1,8 mm**'ye indirir.

### Maliyet

İnceltme maliyeti hücre sayısı × tur sayısıdır ve çözünürlükle **dörde
katlanır**; tur sayısı da yaklaşık ikiye çıkar. Bu yüzden çözünürlük bir
ayardır, 1 mm varsayılan, ve 4 milyon hücre üstü iş reddedilir.

### Ölçülen doğruluk

| Test | Sonuç | Gerçek |
|---|---|---|
| Dalgalı şerit | 470,8 mm | 471,0 mm |
| T kavşağı, 3 dal | 279,3 mm | 280,0 mm |
| İki ayrı harf | 161,0 mm | 160,0 mm |
| Yay örnekleme, tam çember | 0,000000000 mm hata | |
| Yay örnekleme, gerçek bezier | %0,036 | |

## Plotter'a gönderme (`ac2fPlot`)

### Kaydırma neden yazarken uygulanır

Export edilip sonra ayrıştırılan bir `.plt`, metin işleme gerektirir ve
komut biçimlerine (boşluk mu virgül mü, `PA` mı `PR` mi, etiket içindeki
`PU` harfleri) bağımlı hâle gelir. Kaydırmayı **yazarken** geometriden
uygulamak bu bağımlılığı tümden kaldırır: ayrıştırılacak metin yoktur.

Var olan dosyaları düzeltme yolu (`ac2fPlotFixSend`) ayrı bir işlevdir ve
metin ayrıştırır; orada çizim alanı kuralı önemlidir (aşağıda).

### Çizim alanı kuralı

Alan **yalnız çizim komutlarından** hesaplanır: her `PD`, ve ardından
`PD` gelen `PU`. Sondaki `PU0,0;` kalem park komutudur. Ölçülen etki:

| Dosya | Tüm PU/PD | Yalnız çizim |
|---|---|---|
| Çizim negatif, park var | doğru | doğru |
| Çizim **pozitif**, park var | yanlış | doğru |
| Çizim pozitif, park yok | doğru | doğru |

Hata yalnız çizim tamamen pozitifken görünür, çünkü ancak o zaman park
komutunun (0,0) koordinatı minimumu aşağı çeker.

### Eğri adımlaması

HPGL'de eğri yoktur. Her yay, sehim bağıntısından (`h ≈ s²/(8R)`) kendi
yarıçapına göre adımlanır: `s = sqrt(8·R·tol)`. Yumuşak eğriler uzun,
sıkı eğriler kısa adım alır ve hata her yerde tolerans altında kalır.
Ölçüldü: 0,5 / 0,1 / 0,05 / 0,01 mm toleranslarda gerçek sehim hep
sınırın altında.

### Bükme sırası (`ac2fBending`)

Modül, **düzleştirilmiş geometriyi tutan yer**dir: `ac2fPlot` HPGL'i
onun tamponundan yazar. Nokta başına fonksiyon çağrısı olmasın diye
tampon (`ac2fGX`, `ac2fGY`, `ac2fGN`, `ac2fGS`, `ac2fGC`, `ac2fGCl`,
`ac2fGSub`) `Public`. Bir toplama, bir sıralama: `ac2fBNOrderNow`
ikinci kez çağrılırsa sıralamayı tekrarlamaz, sadece ne çalıştığını
söyler (`m_ordered`).

Nesne düzeyinde sıralama (`ac2fBendOrder`) için her kontur, geldiği
**yaprak nesneyi** (`m_sshp`) taşır; grup değil, çünkü istifte yer
değiştiren yapraktır. Nesnenin anahtarı, ona ait en erken konturun
sıradaki yeridir. `OrderToFront` yığının tepesine aldığı için listede
geriye doğru yürümek ilk bükülecek nesneyi en üste koyar.

Konturlar **harflere** ayrılıp sıraya dizilir. Üç karar var.

**Harf nedir.** İçine alma ile bulunur: kapalı bir kontur, kendisinden
**kesin olarak daha büyük** (|alan| karşılaştırması) ve kutusu kendi
kutusunu kapsayan bir konturun içinde mi? Eleme geçilirse ilk noktası
için tek–çift ışın testi yapılır. "Kesin olarak daha büyük" koşulu iki
işi birden görür: eşit kopyaların birbirini ebeveyn göstermesini ve
dolayısıyla döngü oluşmasını engeller. Derinlik tek ise göbek, çift ise
dış; kök dışa kadar çıkılarak harf kimliği atanır.

**Sıra.** Harfler `X0`'a göre dizilir, X'te üst üste binenler bir sütun
sayılır (yeni sütun, sol kenar sütunun sağ kenarını `Column overlap`
kadar geçtiğinde başlar), sütun içinde `Y1` azalan sırada gidilir.
Bindirme eşiğiyle çalışmanın nedeni: bir satırdaki kelimede harfler
yan yanadır ve her biri kendi sütunu olur, sayfaya satır satır dizilmiş
parçalarda ise tam binerler ve doğru biçimde sütun oluştururlar.

**Aynılık.** `ac2fPLSameGlyph`: kontur sayısı, uzunluk, kapsanan alan ve
sıralanmış kutu ölçüleri (`min(w,h)`, `max(w,h)` — çeyrek tur dönmüş
kopya da aynı sayılsın diye), hepsi `Match tol` artı büyük harfler için
oransal bir pay içinde. Ek olarak **işaretli alanın işareti** eşleşmek
zorunda: aynalanmış bir kopyanın uzunluğu ve kutusu aynıdır ama işareti
terstir, ve ters yöne bükülür. Bu, eşleştirmeyi kasten temkinli yapar —
yanlış ayırmanın bedeli bir kurulum, yanlış birleştirmenin bedeli hatalı
parçadır.

Hesap **belge koordinatlarında**, `ac2fPLOrient`'ten önce yapılır.
Noktalar taşınmaz; yalnız kontur dizin dizileri (`m_ss`, `m_sc`, `m_scl`
ve üstverileri) yeniden sıralanır. `ac2fBNApplyOrder` tam bir
permütasyon üretemezse sıralama **hiç** uygulanmaz; yarı sıralı bir iş
göndermekten iyidir.

### Eğrilik ölçüsü

Bir segmentin kirişinden sapması. Dönüş açısı `th` biliniyorsa tam
sagitta `r·(1 − cos(th/2))`. Açı çözücünün eşiğinin ((yay−kiriş)/yay ≥
2e-4) altında kalan sığ yaylar için `Sqr(24·kiriş·(yay − kiriş))/8`
yaklaşımı kullanılır. Ölçülen hata: θ = 0,1 rad'da %0,01, θ = 0,3'te
%0,06 — yani tam da kullanıldığı aralıkta kesin. S biçimli bir segment
iki yöne döndüğü için ikisinde de küçük ölçer; orada yalnız kirişe
eklediği uzunluk kalır. Bu bilinen sınırdır ve bir harfin birden çok
segmenti olduğu için pratikte sorun çıkarmaz.

### Yön

Dönüşüm birebirdir: belge X'i HPGL X'i, belge Y'si HPGL Y'si. Corel'in
ve HPGL'in Y ekseni de aynı yöne bakar, dolayısıyla döndürme yoktur.
`Rotate` ayarı yalnızca makine eksenini telafi etmek içindir ve
varsayılanı 0'dır.

Döndürme, uzanım alınmadan **önce** uygulanır; sonra alınsaydı kenar
payı 90 ve 270'te yanlış kenara düşerdi.

### Ara dosya

Modül CorelDRAW export süzgeci kullanmaz. Gönderimdeki geçici dosya
yalnızca PowerShell sürecine devretme aracıdır ve başarıdan sonra
silinir. Tamamen dosyasız bir yol (stdin borusu) mümkündür ama
`WScript.Shell.Exec` konsol penceresini gizleyemez; `Run` ile gizli
çalıştırma + geçici dosya, kısa bir siyah pencere çakmasına yeğlendi.

### Neden PowerShell

VBA'nın socket'i yok. Seçenekler `ws2_32.dll` için `Declare` (32/64 bit
ve `LongPtr` sorunları), `MSWinsock` OCX (kayıtlı olması gerekir) ya da
bir yardımcı süreç. PowerShell `System.Net.Sockets.TcpClient` sunar,
desteklenen her Windows'ta vardır ve hiçbir kayıt gerektirmez.

Betik her çalıştırmada yeniden yazılır, böylece modülle her zaman
uyumludur. Bir günlük dosyası yazar; makro çıkış kodu sıfır değilse
günlüğü okuyup gerçek hatayı gösterir.

> PowerShell parametresi `-Target` adındadır, `-Host` değil: `$Host`
> PowerShell'in otomatik değişkenidir ve ona bağlanma başarısız olur.

### Bilinen tekrar

Yay açma mantığı (kiriş/yay oranından dönüş açısı, yay örnekleme) artık
üç modülde var: `ac2fBoxLetter`, `ac2fCenterline` ve `ac2fPlot`. Ortak
matematik yardımcıları (`ac2fSolveTheta`, `ac2fAtan2`, `ac2fWrapAngle`)
`ac2fCore`'a taşındı ve yeni kod onları kullanıyor; eski iki modül kendi
özel kopyalarını koruyor. Doğrulanmış modülleri aynı turda kurcalamamak
için bilinçli bırakıldı — açma katmanının tek yere çıkarılması bir
takip işidir.

## Nesting (`ac2fNest` + `ac2fNestCore`)

1.19.0'da sınırlayıcı kutu skyline'ı bırakıldı; yerleştirme gerçek
şekille, ızgara üzerinde yapılıyor. İş ikiye bölündü:

- **`ac2fNestCore`** saf hesaptır: hiçbir CorelDRAW nesnesine dokunmaz.
  Bu yüzden `tools/vbapy.py` onu olduğu gibi Python'a çevirip
  `tools/test_nest.py` ile gerçek geometride sınayabiliyor.
- **`ac2fNest`** belgeyle konuşur: seçimi ayırır, konturu okur,
  dönüşleri yapar, sonucu uygular, pencereyi açar.

### Korumacı ızgara

Bir parça kare hücrelere dökülürken bir hücre şu durumda dolu sayılır:

```
dolu = merkezi malzemenin içinde kalan hücreler
     + her kenar çevresindeki m_d yarıçaplı kapsülün değdiği hücreler
m_d  = parça aralığı / 2 + okuma toleransı
```

Parça, kenarları çevresindeki kapsüllerle birleşince tam olarak `m_d`
kadar büyütülmüş hâlini verir (Minkowski toplamı); merkezi içeride olan
ya da kapsüle değen her hücre işaretlendiği için örtü korumacıdır. Hücresi
çakışmayan iki parça arasında en az `2·m_d >= aralık` vardır. Kapsülün
satır bandıyla kesişimi, kenarın `[y0-r, y1+r]` içindeki x aralığı `r`
kadar genişletilerek alınır (köşeler kare kalır, yine korumacı).

Malzeme kesim mantığına göre: kapalı bir konturun **iç içe derinliği**
çiftse malzeme, tekse delik. Bir taramada her konturun içi tek-çift ile
bulunur, sonra o noktayı tutan en derin kontur çift mi diye bakılır;
böylece üst üste binen iki dış şekil XOR'lanıp delik olmaz. Delikler
kapalıyken her kapalı kontur dolu boyanır.

### Satır başına koşular

Tabaka satırı `Integer` dizisidir:

```
run(c) > 0   boş, c'den itibaren bu kadar boş hücre (KESİN)
run(c) < 0   dolu, c'den itibaren EN AZ -run(c) dolu hücre (alt sınır)
```

Parçanın bir satır aralığı `[a, b)` için tek okuma yeter: `run(x+a) >=
b-a` ise sığar; değilse aynı okuma bir sonraki adayı verir (kısa boş
koşunun bittiği dolu hücrenin koşusunun sonu). Kayan pencere her
atlamada bütün satırları yeniden gezmek yerine döngüsel ilerler ve hiç
atlamasız tam tur atınca durur; her atlama x'i büyüttüğü için sonludur.

Dolu hücrelerde **alt sınır** tutmak kritik: kesin "sonraki boş hücreye
uzaklık" tutulduğunda her yerleştirme soldaki bütün dolu bölgeyi yeniden
yazıyordu (bir dikdörtgen sınamasında süre 25 s'den 10 s'ye indi).
Tabaka bir geçiş boyunca yalnız dolduğu için eski bir sınır asla yanlış
olmaz, yalnız bir atlama daha yaptırır. Boş koşular ise kesin kalmak
zorunda; yalnız yerleşen aralığın solundaki boş koşu güncellenir.

### Arama ve not defteri

Her yön (dönüş/ayna) için en alçak, sonra en soldaki boş yer aranır.
Her satırın **en geniş boş koşusu** (`m_mf`) tutulur; parçanın en geniş
tek aralığına sahip üç satırı o satırlara sığmıyorsa satır hiç
denenmeden geçilir.

Not defteri `(yön, tabaka)` başına "son aramanın bittiği yer"dir.
Tabaka yalnız dolduğu için altı bir daha açılmaz; aynı parçanın ikinci
kopyası aramaya oradan başlar. Strateji 1'de (en alçak üst kenar) bir yön
bulunduktan sonra diğer yönler o üst kenarı geçemeyecek satırlarda
aramayı bırakır; bırakılan satırlar da not defterine yazılır.

### Optimizasyon

Bir geçiş bütün parçaları bir sırada yerleştirir. `ac2fNCOptStep` önce
sabit geçişleri yapar (alana göre, uzun kenara göre, yükseklik,
genişlik, ikinci strateji, döndürmesiz), sonra en iyi sırada farklı
sınıftan parçaları (çoğunlukla yakın olanları) takas eder; çağıran süre
bitene kadar çağırır. Skor sözlüksel: yerleşmeyen parça sayısı, sonra
malzeme (son açılan tabaka yalnız en üstteki parçaya kadar), sonra
toplam kullanılan yükseklik. Tabakaların başlangıç hâli tek bir dizi
atamasıyla (`m_run = m_runInit`, yerel memcpy) geri gelir.

### Belge tarafı

- **Dönüşler CorelDRAW'da yapılır.** Her sınıfın her yönü için temsilci
  şekil gerçekten çevrilir (`Flip`, `Rotate`), kontur o hâliyle okunur,
  sınırlayıcı kutusunun köşesi kaydedilir, şekil geri çevrilir. Uygularken
  parça aynı işlemleri görür ve **kutusunun köşesine göre** taşınır:
  dönme merkezi nerede olursa olsun, kopyalar dahil, doğru konum.
- **Başlangıç köşesi ve yön** motora dokunmadan, koordinat çerçevesi
  değiştirilerek yapılır: `x' = a x + b y`, `y' = c x + e y` (eksen
  takası ve işaret çevirme). Motor hep kendi sol altından, kendi
  satırlarıyla doldurur; taşıma vektörü çerçevenin tersiyle (devriği)
  sayfaya döner.
- **Kontur okuma**: `Segment.GetPointPositionAt` ile eğri boyunca
  noktalar. İlk çağrıda uçların düğümlere denk geldiği, her noktada
  `|AP| + |PB| <= L` olduğu denetlenir. Çağrı yoksa ya da tutarsızsa her
  eğri parçası için kiriş + **o boydaki her eğriyi içine alan** kutu
  (`AB` odaklı, büyük ekseni `L` olan elipsin dikdörtgeni) katı kontur
  olarak eklenir; yerleşim gevşer ama çakışmaz.
- **İç / dış**: her nesnenin konturu bir havuza yazılır; bir nesnenin üç
  noktası başka nesnenin kapalı konturunda ve kutusu kesin büyükse
  onun içindedir; en dıştaki kapsayıcıyla gruplanır.
- **Kopya tanıma**: konturların kutuya göre 0,01 mm'ye yuvarlanmış
  koordinatlarından imza; ayna görüntüsü farklı imza verir.
- Bellek: hücre bütçesi 8 milyon (iki `Integer` dizisi); iş sığmazsa
  hassasiyet kalınlaşır, açılabilecek tabaka sayısı sınırlanır.

### Pencereler

İki betik vardır: `tools/nest_window.ps1` (ayarlar) ve
`tools/nest_preview.ps1` (önizleme). `tools/ps2vba.py` ikisini de satır
satır `Print #f` olarak `ac2fNest`'in sonuna gömer, ~8 KB'lık yordamlara
böler (`ac2fNSPsSet1..`, `ac2fNSPsPrev1..`); `ac2fNSWriteScript` istenen
betiği sırayla yazar, `ac2fNSShowWindow` çalıştırıp çıktıyı okur.

**Ayar penceresi** — giriş/çıkış `|` ile ayrılmış UTF-8 satırlar:
`K|anahtar|değer`, `S|kullan|ad|en|boy|adet|kaynak` (`L` liste, `D<n>`
çizili artık), `G|sayfa eni|boyu`, `R|hazır ayar|blob`; çıkışta
`A|PREVIEW`, `A|NEST` ya da `A|CANCEL`. Hazır ayarlar `NestPresets`
bölümünde, hangi düğmeyle kapanırsa kapansın geri yazılır.

**Önizleme** — giriş `Print #` ile doğrudan dosyaya yazılır (megabaytları
bulabilir, VBA'da tek dizede biriktirmek yavaştır), ANSI; betik aynı kod
sayfasıyla okur. Koordinatlar sayfada, milimetrenin onda biri cinsinden
**tam sayıdır** (ondalık ayırıcıya bağlı hiçbir şey yok):

```
T|başlık            R|rapor satırı
S|n|ad|parça|x0|y0|x1|y1|kx0|ky0|kx1|ky1   tabaka ve kenar payı içi
O|n|tür|x y x y ...                        çizili artığın konturu
P|n|parça|sınıf|tür|x y x y ...            yerleşen parçanın bir konturu
```

Her yönün (variant) konturu motor tarafından inceltilip saklanır
(`ac2fNCKeepOutline`, kontur başına en çok 120 nokta). Yerleşen parçanın
konturu, rasterin gittiği hücreye kaydırılır (motor çerçevesi) ve
çerçevenin tersiyle (devriği) sayfaya döner — `ac2fNSApply`'ın taşıma
vektörüyle aynı hesap, bu yüzden çizim uygulanınca parçanın gideceği
yerle birebir aynıdır (`test_corel.py` bunu denetler). Çıkış `A|APPLY`,
`A|SETTINGS` ya da `A|CANCEL`.

**Akış** (`ac2fNSRun`): hesap tek komut grubunda yapılır; önizleme açıkken
belge iç/dış gruplarıyla bekler. `Apply` taşır ve grubu kapatır; diğer iki
düğmede grup kapanıp `Undo` ile her şey geri alınır, *Change settings*
seçimdeki aynı şekillerle ayar penceresini yeniden açar.

## Başlangıç noktaları (`ac2fStart`)

En yakın nokta segment boyunca örneklerle bulunur, eğride altın oran
aramasıyla, düz çizgide analitik olarak inceltilir. Tam nokta modunda
oraya `Segment.AddNodeAt` ile düğüm eklenir. Düğümü başlangıç yapmak için
yol o düğümde kırılır (`Node.BreakApart`) ve uçları yeniden birleştirilir
(`StartNode.JoinWith EndNode`) — CorelDRAW kullanıcılarının elle yaptığı
yöntem. Yön, örneklenmiş konturun işaretli alanından (y yukarı: pozitif
= saat yönünün tersi) ve iç içe derinliğinden bulunur; ters çevirmek için
`SubPath.ReverseDirection` denenir, sürüm desteklemiyorsa sayılıp
raporlanır.

## Sınama araçları

```
tools/vbapy.py         VBA -> Python derleyicisi (alt küme)
tools/cdrmock.py       CorelDRAW nesne modeli taklidi
tools/vbacheck.py      gerçek VBA derleyicisinin kuralları (lint çalıştırır)
tools/test_nest.py     ac2fNestCore, gerçek geometriyle
tools/test_corel.py    ac2fNest + ac2fStart, taklit belgede uçtan uca
tools/test_windows.py  pencere betikleri, PowerShell 7 + tools/winstub.cs
```

`test_windows.py` PowerShell 7'yi (`pwsh`) ister: `AC2F_PWSH` ortam
değişkeni ya da `PATH`. Linux'ta WinForms ve GDI+ olmadığından
`tools/winstub.cs` onların yerini tutar (yalnız betiklerin kullandığı
üyeler, WinForms'un hata verdiği yerde hata vererek: aralık dışı
`SelectedIndex`, bağlı satırı yeniden ekleme, 3 noktadan az çokgen ...).
Sınama betiği değiştirmeden çalıştırır (yalnız GDI+ türleri taklide
yönlendirilir), düğmelere basar, kutulara yazar, çıkışı okur; ayrıca
Windows PowerShell 5.1'de olmayan yapıları (`?:`, `??`, `?.`, `&&`) ve
`$PID`, `$Input` gibi otomatik değişkenlere atamayı arar. `pwsh` varsa
`test_corel.py` da VBA'nın yazdığı gerçek betiklerle uçtan uca çalışır
(`t_real_windows`); yoksa o senaryo atlanır.

`vbapy` VBA anlamını korumaya çalışır: `Long`/`Integer` atamada bankacı
yuvarlaması ve taşma hatası, sınır denetimli diziler (`a(-1)` sondan
okumaz), kısa devre yapmayan `And`/`Or`, ByRef skalerlerin geri
yazılması, `a = b` dizi kopyası, `On Error Resume Next` / `GoTo etiket`
(etiketler arası bölümlerle bir dağıtıcı döngü), nesnede olmayan üyeye
erişimin hata 438 olması, `Option Explicit` (tanımsız ad derleme hatası).
Desteklemediği bir yapı (UDT gibi) o yordamı taslağa çevirir; çağrılırsa
`NotImplementedError` atar.

`cdrmock` gerçek geometri taşır: kübik bezier yollar, kutunun merkezine
göre dönüş/ayna, eğriden ölçülen kutu, `Duplicate`, `Group`, `Ungroup`,
komut grubu + `Undo` (anlık görüntü), kayıt defteri, dosyalar,
`InputBox`/`MsgBox`, `WScript.Shell` (pencere yerine sınamanın verdiği bir
işleyici) ve `ADODB.Stream`. Sınamalar makro bittikten sonra belgedeki
gerçek şekillere bakar: aralık, kenar payı, iç/dış, kopyalar, rapor
modunun belgeyi değiştirmemesi, pencere protokolü, başlangıç noktaları.

Taklit, CorelDRAW'ın davranışı hakkındaki varsayımları da taşır (ör.
`BreakApart` + `JoinWith` kırılan düğümü başlangıç yapar). Bu varsayımlar
gerçek CorelDRAW'da ilk kullanımda doğrulanmalıdır; sınamalar kodun kendi
içinde tutarlı olduğunu gösterir.

## Müşteri işi (`ac2fJob`)

StockTrack'le konuşan tek modül. Üç tasarım kararı:

**Düz metin protokol.** VBA'da JSON ayrıştırıcı yok. StockTrack'in
`/corel/*` uçları sekmeyle ayrılmış, tamamen ASCII satırlar döner; ilk
satır `#ac2f-corel 1`. Modül bu başlığı görmezse yanlış adrese
bağlandığını söyler. İstek `MSXML2.ServerXMLHTTP.6.0` ile gider (zaman
aşımı alır, IE önbelleğinden cevap vermez); yoksa WinHttp.

**Adlandırma kuralı iki yerde, aynı.** Müşteri klasörü ve ad jetonları
StockTrack'te üretilir (`corel-naming.util.ts`, Jest testli); modül
bunları olduğu gibi kullanır. Elle yazılan müşteri/plaka için aynı
kuralın VBA kopyası var (`ac2fJBSafeName`, `ac2fJBToken`, `ac2fJBFold`).
Biri değişirse öteki de değişmeli — yoksa aynı plaka iki klasöre düşer.
Türkçe harfler `ChrW` kodlarıyla katlanır; kaynak ASCII kalır.

**İstif durumu diskte.** Hangi klasörün kaç sayfa olduğu hiçbir yerde
saklanmaz; klasör adından okunur (`ac2fJBSplitGroup`). `--` ile ayrılmış
her yerde `rakam + pcs` aranır, bu yüzden ayraç `-` bile olsa çözülür.
Jeton hiçbir zaman `--` içermez (harf/rakam dışı her dizi tek `-`).
`Dir` iç içe çağrılamadığı için klasörler önce tek seferde listelenir;
varlık denetimi `GetAttr` ile yapılır.

Belge bağlantısı `Document.Properties` ile dosyaya yazılır. Bu üye
**geç bağlanır** (`As Object`): olmadığı bir sürümde erken bağlı çağrı
bütün projeyi derletmezdi. Kayıt defterindeki kopya dosya yolunun
özetiyle anahtarlanır.

## Yeni bir araç eklemek

1. `src/ac2fYeniArac.bas` oluşturun, ilk satır:
   `Attribute VB_Name = "ac2fYeniArac"`
2. `Option Explicit` yazın.
3. Giriş noktasını **parametresiz `Public Sub`** yapın, hemen altına
   `Attribute <YordamAdı>.VB_Description = "ac2f pack: ..."` ekleyin.
   Makro adı ve tüm metinler İngilizce ve ASCII olmalı.
4. Ölçüm gerekiyorsa `ac2fMeasureSelection()` çağırın — tekrar yazmayın.
5. `ac2fMenu.ac2fPack` içindeki `Select Case` bloğuna yeni bir numara ekleyin.
6. `python3 tools/lint.py` çalıştırın.
7. README'deki makro tablosunu ve `CHANGELOG.md` dosyasını güncelleyin.

## Sınırlar

- **Arayüz İngilizce, belgeler Türkçe.** Kaynağın ASCII kalması bilinçli
  bir kısıttır; kodlama sorunlarını tamamen ortadan kaldırır.
- **Windows'a özgü.** `GetSetting`/`SaveSetting` kayıt defterini kullanır.
  macOS CorelDRAW'da ayarlar kalıcı olmaz; hesaplar yine çalışır.
- **UserForm yok**, dolayısıyla fare ipucu da yok. Gerekçesi yukarıda.
- **LED yerleşim çizimi yok.** Paket adedi hesaplar, modülleri sayfaya
  yerleştirmez.
- **Kutu harf şeridi tek parça çizilir.** Rulo boyu aşılıyorsa kesim
  yerleri işaretlenir, parçalar ayrı çizilmez.
