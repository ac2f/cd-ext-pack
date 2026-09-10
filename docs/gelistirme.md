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
             └──> ac2fBoxLetter ──┴──> ac2fCore
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
- **UserForm yok.** Diyaloglar `MsgBox`/`InputBox` ile kurulmuştur. Bunun
  nedeni, `.frm` dosyalarının yanlarında ikili bir `.frx` gerektirmesi ve
  bu ikilinin metin bir depoda güvenle üretilip sürümlenememesidir. Arayüz
  eklenecekse form VBE içinde çizilip `.frm` + `.frx` birlikte eklenmelidir.
- **LED yerleşim çizimi yok.** Paket adedi hesaplar, modülleri sayfaya
  yerleştirmez.
- **Kutu harf şeridi tek parça çizilir.** Rulo boyu aşılıyorsa kesim
  yerleri işaretlenir, parçalar ayrı çizilmez.
