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

Bu kuralı koruyun — yeni metin eklediğinizde doğrulayın:

```bash
LC_ALL=C grep -n '[^ -~\t]' src/*.bas    # hiçbir şey dönmemeli
```

`tools/build.ps1` ve `tools/build.sh` yalnız CRLF satır sonu üretir;
ASCII kaynakla artık zorunlu değildir, kolaylık içindir.

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

## Mimari

```
ac2fMenu ────┬──> ac2fLength ─────┐
             ├──> ac2fLedModule ──┤
             ├──> ac2fBoxLetter ──┤
             ├──> ac2fPanel ──────┤
             ├──> ac2fCenterline ─┼──> ac2fCore
             └──> ac2fSettings ───┘
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
ve raporlarda duruyor; mm dışı birimler etikete taşındı).

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

Form eklenecekse VBE içinde çizilip `.frm` + `.frx` **birlikte**
depoya konmalıdır.

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
