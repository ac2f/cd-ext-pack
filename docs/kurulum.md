# Kurulum

## 0. Ön koşul: VBA desteği

CorelDRAW kurulumunda **Visual Basic for Applications** bileşeni seçili
olmalıdır. `Alt+F11` bir pencere açmıyorsa:

Windows *Ayarlar > Uygulamalar* → CorelDRAW → *Değiştir* → *Modify* →
**Visual Basic for Applications** işaretle → kurulumu tamamla.

## 1. Kaynağı hazırla (isteğe bağlı)

Sürüm 1.2.0'dan itibaren kaynak **saf ASCII**'dir — hiçbir Türkçe karakter
içermez. Bu yüzden `src\` altındaki `.bas` dosyalarını doğrudan içe
aktarabilirsiniz; kod sayfası dönüşümü artık gerekmiyor.

Yine de satır sonlarını CRLF yapmak için betiği çalıştırabilirsiniz:

```powershell
powershell -ExecutionPolicy Bypass -File tools\build.ps1
```

`build\` klasöründe yedi dosya oluşur:

```
ac2fCore.bas
ac2fLength.bas
ac2fLedModule.bas
ac2fBoxLetter.bas
ac2fPanel.bas
ac2fSettings.bas
ac2fMenu.bas
```

> Linux/macOS'ta: `sh tools/build.sh`

## 2. Projeyi oluştur

1. CorelDRAW'ı açın.
2. `Araçlar > Makrolar > Makro Düzenleyici` (veya `Alt+F11`).
3. Açılan VBA düzenleyicisinde `File > New Project` ile yeni bir proje
   oluşturun ve adını **`ac2fPack`** koyun.

   Proje penceresinde projeye sağ tıklayıp
   `ac2fPack Özellikleri` ile *Project Name* alanını `ac2fPack` yapabilirsiniz.

   Proje dosyası şuraya yazılır:

   ```
   %AppData%\Corel\CorelDRAW Graphics Suite <sürüm>\Draw\GMS\ac2fPack.gms
   ```

## 3. Modülleri içe aktar

`ac2fPack` projesi seçiliyken, yedi dosyanın **her biri** için:

`File > Import File...` → `build\ac2*.bas` → **Open**

Sonuçta proje ağacı şöyle görünmeli:

```
ac2fPack
└── Modules
    ├── ac2fCore
    ├── ac2fLength
    ├── ac2fLedModule
    ├── ac2fBoxLetter
    ├── ac2fPanel
    ├── ac2fSettings
    └── ac2fMenu
```

## 4. Derle ve kaydet

1. `Debug > Compile ac2fPack` — hata vermemeli.
2. `File > Save ac2fPack`.

## 5. Çalıştır

`Araçlar > Makrolar > Makroları Çalıştır` → *Makro konumu:* `ac2fPack` →
**`ac2fMenu.ac2fPack`** → **Çalıştır**.

### Araç çubuğuna kısayol koymak

1. `Araçlar > Seçenekler > Özelleştirme > Komutlar`
2. Soldaki açılır listeden **Makrolar**'ı seçin.
3. `ac2fMenu.ac2fPack` komutunu araç çubuğuna sürükleyin.

Aynı yerden `Kısayol Tuşları` sekmesiyle klavye kısayolu da atayabilirsiniz.

## Güncelleme

**Projeyi baştan oluşturmayın.** Var olan `ac2fPack` projesine yalnızca
değişen dosyalar aktarılır.

1. `tools\build.ps1` çalıştırın.
2. **Yeni** modülleri doğrudan içe aktarın (`File > Import File`).
3. **Değişmiş** modüller için önce eskisini silin, sonra yenisini aktarın
   (aşağıdaki uyarıya bakın).
4. `Debug > Compile ac2fPack`, sonra `File > Save ac2fPack`.

Dokunulmayan modülleri olduğu gibi bırakın.

Ayarlarınız kayıt defterinde tutulduğu için güncellemede korunur.

### Önemli: içe aktarma üzerine yazmaz

VBE aynı adlı bir modülü **değiştirmez**, yanına ikinci bir kopya ekler:
`ac2fMenu` dururken `ac2fMenu.bas` aktarılırsa `ac2fMenu1` oluşur. İki
modülde de aynı `Public Sub` adları bulunduğu için derleme şu hatayı verir:

```
Ambiguous name detected: ac2fPack
```

Bu yüzden değişmiş bir modülü aktarmadan **önce** eskisini silin:

> Proje ağacında modüle sağ tıkla → `Remove ac2fMenu` →
> "Do you want to export...?" sorusuna **No**

Zaten `ac2fMenu1` oluştuysa onu silin, sonra eski `ac2fMenu`'yü silip
yeniden aktarın.

### Sürüm 1.3.0'dan 1.4.0'a

| Dosya | Ne yapmalı |
|---|---|
| `ac2fPanel.bas` | **Yeni** — doğrudan içe aktarın |
| `ac2fCore.bas` | **Değişti** — önce silin, sonra aktarın |
| `ac2fSettings.bas` | **Değişti** — önce silin, sonra aktarın |
| `ac2fMenu.bas` | **Değişti** — önce silin, sonra aktarın |
| `ac2fLength`, `ac2fLedModule`, `ac2fBoxLetter` | Değişmedi — dokunmayın |

Yeni makro: `ac2fPanelGroove`. Ayarlarınız korunur.
Ana menü numaraları kaydı: ayarlar artık `9`, hakkında `10`.

### Sürüm 1.2.0'dan 1.3.0'a

Bir modül **eklendi**, dördü **değişti**:

| Dosya | Ne yapmalı |
|---|---|
| `ac2fSettings.bas` | **Yeni** — doğrudan içe aktarın |
| `ac2fCore.bas` | **Değişti** — önce silin, sonra aktarın |
| `ac2fBoxLetter.bas` | **Değişti** — önce silin, sonra aktarın |
| `ac2fLedModule.bas` | **Değişti** — önce silin, sonra aktarın |
| `ac2fMenu.bas` | **Değişti** — önce silin, sonra aktarın |
| `ac2fLength.bas` | Değişmedi — dokunmayın |

Kaldırılan makrolar: `ac2fLedSettings` ve `ac2fBoxLetterSettings`.
Yerlerini tek `ac2fSettings` sayfası aldı. Bu ikisine araç çubuğu
kısayolu atadıysanız kaldırın; yerine `ac2fSettings` koyun.

Eklenen makrolar: `ac2fSettings`, `ac2fProfiles`,
`ac2fBoxLetterStripProfile`.

Ayarlarınız korunur.

### Sürüm 1.1.0'dan 1.2.0'a — kırıcı değişiklik

Arayüz İngilizceye çevrildi ve **makro adları değişti**. Bu yüzden
**beş modülün tamamı** yeniden aktarılmalıdır.

1. Beş eski modülü de silin (her birine sağ tık → `Remove ...` → **No**).
2. `src\` (veya `build\`) altındaki beş `.bas` dosyasını aktarın.
3. `Debug > Compile ac2fPack`, sonra `File > Save ac2fPack`.

**Araç çubuğu kısayollarınız bozulur.** Makro adları değiştiği için
eskiden atadığınız düğmeler artık bulunamayan bir makroyu gösterir;
silip yeniden atamanız gerekir.

| Eski makro adı | Yeni makro adı |
|---|---|
| `ac2fUzunlukOlc` | `ac2fMeasureLength` |
| `ac2fUzunlukEtiketle` | `ac2fLabelLength` |
| `ac2fLedModulHesapla` | `ac2fLedModuleCount` |
| `ac2fLedHizliHesap` | `ac2fLedQuickCount` |
| `ac2fKutuHarfSerit` | `ac2fBoxLetterStrip` |
| `ac2fKutuHarfRapor` | `ac2fBoxLetterReport` |
| `ac2fKutuHarfAyarlar` | `ac2fBoxLetterSettings` |
| `ac2fAyarlar` | `ac2fLedSettings` |
| `ac2fAyarlariSifirla` | `ac2fResetAllSettings` |
| `ac2fHakkinda` | `ac2fAbout` |
| `ac2fPack` | `ac2fPack` *(değişmedi)* |

**Ayarlarınız korunur** — kayıt defteri anahtarları bilerek
değiştirilmedi.

## Sorun giderme

| Belirti | Çözüm |
|---|---|
| `Alt+F11` açılmıyor | VBA bileşeni kurulu değil (bkz. adım 0) |
| Menüde eski Türkçe metinler görünüyor | Eski modüller silinmemiş; bkz. "içe aktarma üzerine yazmaz" |
| `User-defined type not defined` | `ac2fCore` içe aktarılmamış; önce onu ekleyin |
| Makro listede görünmüyor | Makro yalnızca **Public Sub** ve **parametresiz** ise listelenir; `Debug > Compile` ile hata olup olmadığına bakın |
| "Makrolar devre dışı" uyarısı | `Araçlar > Makrolar > Güvenlik` → güven düzeyini düşürün |
