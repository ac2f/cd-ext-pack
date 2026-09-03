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
| Genel makro | `ac2f` + Türkçe eylem | `ac2fUzunlukOlc` |
| Sabit | `AC2F_` | `AC2F_DEF_SPACING` |
| Ayar anahtarı | `AC2F_K_` | `AC2F_K_SPACING` |
| Tip | `ac2f` + isim | `ac2fLedResult` |

Makro Yöneticisi'nde yalnızca **parametresiz `Public Sub`** yordamları
listelenir. Bir yordamın kullanıcıya görünmesini istemiyorsanız `Private`
yapın ya da parametre verin.

## Dosya kodlaması

| Yer | Kodlama | Satır sonu |
|---|---|---|
| `src/*.bas` (depo) | UTF-8 | LF |
| `build/*.bas` (içe aktarılan) | Windows-1254 | CRLF |

VBE'nin `File > Import File` işlevi dosyayı **sistem ANSI kod sayfasıyla**
okur; Türkçe Windows'ta bu 1254'tür. UTF-8 dosyayı doğrudan aktarmak
derleme hatası vermez ama diyaloglardaki Türkçe harfleri bozar.

`tools/build.ps1` (Windows) ve `tools/build.sh` (POSIX) bu dönüşümü yapar.

> Kodu VBE'ye **kopyala-yapıştır** ile taşırsanız dönüşüm gerekmez; pano
> Unicode taşır. Dönüşüm yalnızca dosyadan içe aktarma içindir.

Yeni metin eklerken kullandığınız karakterlerin CP1254'e çevrilebildiğini
doğrulayın:

```bash
iconv -f UTF-8 -t WINDOWS-1254 src/ac2fMenu.bas > /dev/null
```

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

VBA derleyicisinin yerini **tutmaz**. Asıl doğrulama CorelDRAW'da
`Debug > Compile ac2fPack` ile yapılır.

## Mimari

```
ac2fMenu ────┐
             ├──> ac2fLength ─────┐
             └──> ac2fLedModule ──┴──> ac2fCore
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

## Yeni bir araç eklemek

1. `src/ac2fYeniArac.bas` oluşturun, ilk satır:
   `Attribute VB_Name = "ac2fYeniArac"`
2. `Option Explicit` yazın.
3. Giriş noktasını **parametresiz `Public Sub`** yapın, hemen altına
   `Attribute <YordamAdı>.VB_Description = "ac2f pack: ..."` ekleyin
   (açıklama metni ASCII olmalı).
4. Ölçüm gerekiyorsa `ac2fMeasureSelection()` çağırın — tekrar yazmayın.
5. `ac2fMenu.ac2fPack` içindeki `Select Case` bloğuna yeni bir numara ekleyin.
6. `python3 tools/lint.py` çalıştırın.
7. README'deki makro tablosunu ve `CHANGELOG.md` dosyasını güncelleyin.

## Sınırlar

- **Windows'a özgü.** `GetSetting`/`SaveSetting` kayıt defterini kullanır.
  macOS CorelDRAW'da ayarlar kalıcı olmaz; hesaplar yine çalışır.
- **UserForm yok.** Diyaloglar `MsgBox`/`InputBox` ile kurulmuştur. Bunun
  nedeni, `.frm` dosyalarının yanlarında ikili bir `.frx` gerektirmesi ve
  bu ikilinin metin bir depoda güvenle üretilip sürümlenememesidir. Arayüz
  eklenecekse form VBE içinde çizilip `.frm` + `.frx` birlikte eklenmelidir.
- **Yerleşim çizimi yok.** Paket adedi hesaplar, modülleri sayfaya
  yerleştirmez.
