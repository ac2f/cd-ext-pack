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

VBE içe aktarması Windows satır sonu (CRLF) bekler. 1.21.0'dan beri depo
`.bas` dosyalarını klonda ve GitHub ZIP'inde CRLF olarak verir. Tek bir
dosyayı tarayıcıdan "raw" indirdiyseniz LF gelebilir; o zaman ya da emin
olmak için betiği çalıştırın:

```powershell
powershell -ExecutionPolicy Bypass -File tools\build.ps1
```

`build\` klasöründe on sekiz dosya oluşur:

```
ac2fCore.bas
ac2fLength.bas
ac2fLedModule.bas
ac2fBoxLetter.bas
ac2fPanel.bas
ac2fCenterline.bas
ac2fBending.bas
ac2fPlot.bas
ac2fNestCore.bas
ac2fNest.bas
ac2fStart.bas
ac2fResize.bas
ac2fJob.bas
ac2fCheck.bas
ac2fCutTime.bas
ac2fCutArea.bas
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

`ac2fPack` projesi seçiliyken, on sekiz dosyanın **her biri** için:

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
    ├── ac2fCenterline
    ├── ac2fBending
    ├── ac2fPlot
    ├── ac2fNestCore
    ├── ac2fNest
    ├── ac2fStart
    ├── ac2fResize
    ├── ac2fJob
    ├── ac2fCheck
    ├── ac2fCutTime
    ├── ac2fCutArea
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

### Sürüm 1.26.0'dan 1.27.0'a

| Dosya | Ne yapmalı |
|---|---|
| `ac2fResize.bas` | **Yeni** — doğrudan içe aktarın |
| `ac2fMenu.bas` | **Değişti** (menü `29` / `30`, Hakkında satırı) — önce silin, sonra aktarın |
| `ac2fCore.bas` | **Değişti** (sürüm) — önce silin, sonra aktarın |
| Diğerleri | Değişmedi |

Yeni makrolar: `ac2fResizeEach` (menü `29`), `ac2fResizeEachAgain` (menü
`30`). Seçimdeki her nesneyi ayrı ayrı yazdığınız ölçüye getirir. Ayar
sayfasına yeni ayar eklenmedi; numaralar ve kayıtlı ayarlar değişmedi.

### Sürüm 1.25.1'den 1.26.0'a

| Dosya | Ne yapmalı |
|---|---|
| `ac2fPanel.bas` | **Değişti** (net/kesim ölçü girişi, konumla bölme, onay) — önce silin, sonra aktarın |
| `ac2fCore.bas` | **Değişti** (sürüm) — önce silin, sonra aktarın |
| Diğerleri | Değişmedi |

Eski girişler (`3200 960`, `3050:3000 3200:*`) aynen çalışır; artık yazılan
plan çizmeden önce parça parça gösterilip onay sorulur.

### Sürüm 1.25.0'dan 1.25.1'e

| Dosya | Ne yapmalı |
|---|---|
| `ac2fPanel.bas` | **Değişti** (plakalar ayrı gruplanıyor) — önce silin, sonra aktarın |
| `ac2fCore.bas` | **Değişti** (sürüm) — önce silin, sonra aktarın |
| Diğerleri | Değişmedi |

### Sürüm 1.24.0'dan 1.25.0'a

| Dosya | Ne yapmalı |
|---|---|
| `ac2fPanel.bas` | **Değişti** (katmanlar, yerinde bölme) — önce silin, sonra aktarın |
| `ac2fCore.bas` | **Değişti** (sürüm) — önce silin, sonra aktarın |
| Diğerleri | Değişmedi |

Panel artık dış kesimi `ACP Cut`, derzleri `ACP Grooves` katmanına çizer
(katmanlar yoksa açılır). Bölünen panelde plakalar net ölçü üzerinde,
yerinde çizilir; eski gibi ayrı dizmek için plaka satırına `sep` yazın.
Kayıtlı plaka listesi ve derz payı aynen geçerli.

### Sürüm 1.23.1'den 1.24.0'a

| Dosya | Ne yapmalı |
|---|---|
| `ac2fPanel.bas` | **Değişti** (plakalara bölme) — önce silin, sonra aktarın |
| `ac2fMenu.bas` | **Değişti** (ana menü `28`) — önce silin, sonra aktarın |
| `ac2fCore.bas` | **Değişti** (sürüm) — önce silin, sonra aktarın |
| Diğerleri | Değişmedi |

Panel makrosu, açınım plaka uzunluğundan büyükse plakaları ve derz payını
sorar; sığan panelde davranış eskisi gibidir. İlk kullanımda plaka listesi
`3200`, derz payı `30` gelir; sorudan değiştirince hatırlanır. Ayar
numaraları ve kayıtlı ayarlar değişmedi.

### Sürüm 1.23.0'dan 1.23.1'e

| Dosya | Ne yapmalı |
|---|---|
| `ac2fPlot.bas` | **Değişti** (`Rotate` varsayılanı 270) — önce silin, sonra aktarın |
| `ac2fSettings.bas` | **Değişti** (ayar açıklaması) — önce silin, sonra aktarın |
| `ac2fCore.bas` | **Değişti** (sürüm) — önce silin, sonra aktarın |
| Diğerleri | Değişmedi |

Plotter artık işi ekranda gördüğünüz gibi keser; ayar yapmanız gerekmez.
Daha önce `Rotate` değerini elle yazdıysanız o değer geçerli kalır
(ana menü `19` › Rotate ya da adres satırına `r270`).

### Sürüm 1.22.0'dan 1.23.0'a

| Dosya | Ne yapmalı |
|---|---|
| `ac2fCore.bas` | **Değişti** (eğri örnekleme, sürüm) — önce silin, sonra aktarın |
| `ac2fBending.bas` | **Değişti** (gerçek eğri noktaları) — önce silin, sonra aktarın |
| `ac2fPlot.bas` | **Değişti** (yön sihirbazı, tasarım denetimi) — önce silin, sonra aktarın |
| `ac2fCenterline.bas` | **Değişti** (derleme hatası giderildi) — önce silin, sonra aktarın |
| `ac2fMenu.bas` | **Değişti** (menü `27`) — önce silin, sonra aktarın |
| Diğerleri | Değişmedi |

Plotter eğrileri artık tasarımdaki gibi çıkar (eski sürüm S kıvrımlı ve
el yazısı şekilleri milimetrelerce bozuyordu). Bir kez **ana menü `27`**
(yön sihirbazı) çalıştırın: iş 90 derece dönük ya da aynalı çıkıyorsa
`Rotate`/`Mirror` değerini tek kesimle bulur. Ayar numaraları ve kayıtlı
ayarlar değişmedi. Beş modülü de yenileyin, sonra `Debug > Compile ac2fPack`
hatasız bitmeli.

### Sürüm 1.21.0'dan 1.22.0'a

| Dosya | Ne yapmalı |
|---|---|
| `ac2fJob.bas` | **Değişti** — önce silin, sonra aktarın |
| `ac2fCore.bas` | **Değişti** (sürüm) — önce silin, sonra aktarın |
| Diğerleri | Değişmedi |

Takım yolu klasörlerinin adına `--F<gün>.<sahip>.<tür>-<renk>` belirteci
eklenir. StockTrack Runner'ı da 1.3.0'a güncelleyin ki ikisi aynı adları
kursun. Ayar numaraları değişmedi.

### Sürüm 1.20.0'dan 1.21.0'a

| Dosya | Ne yapmalı |
|---|---|
| `ac2fNest.bas` | **Değişti** — önce silin, sonra aktarın |
| `ac2fNestCore.bas` | **Değişti** — önce silin, sonra aktarın |
| `ac2fSettings.bas` | **Değişti** — önce silin, sonra aktarın |
| `ac2fMenu.bas` | **Değişti** — önce silin, sonra aktarın |
| `ac2fJob.bas` | **Değişti** — önce silin, sonra aktarın |
| `ac2fCore.bas` | **Değişti** (sürüm) — önce silin, sonra aktarın |
| Diğerleri | Değişmedi |

Bu sürüm, nesting makrolarının **"Expected variable or procedure, not
module"** derleme hatasını giderir. Altı modülü de yenileyin, sonra
`Debug > Compile ac2fPack` hatasız bitmeli. Ayar numaraları ve kayıtlı
ayarlar değişmedi.

### Sürüm 1.19.0'dan 1.20.0'a

| Dosya | Ne yapmalı |
|---|---|
| `ac2fJob.bas` | **Değişti** — önce silin, sonra aktarın |
| `ac2fCore.bas` | **Değişti** (sürüm) — önce silin, sonra aktarın |
| Diğerleri | Değişmedi |

Job setup'ta (iş menüsü `7`) yeni `5` satırı: dükkândaki yerel StockTrack
sunucusunun adresi. Ana sunucu cevap vermezse listeler oradan gelir. Ayar
numaraları değişmedi.

### Sürüm 1.18.0'dan 1.19.0'a

| Dosya | Ne yapmalı |
|---|---|
| `ac2fNestCore.bas` | **Yeni** — doğrudan içe aktarın |
| `ac2fStart.bas` | **Yeni** — doğrudan içe aktarın |
| `ac2fNest.bas` | **Değişti** — önce silin, sonra aktarın |
| `ac2fSettings.bas` | **Değişti** — önce silin, sonra aktarın |
| `ac2fMenu.bas` | **Değişti** — önce silin, sonra aktarın |
| `ac2fJob.bas` | **Değişti** — önce silin, sonra aktarın |
| `ac2fCore.bas` | **Değişti** (sürüm) — önce silin, sonra aktarın |
| Diğerleri | Değişmedi |

Yeni makrolar: `ac2fNestAgain`, `ac2fNestMarkSheets`, `ac2fStartPoints`,
`ac2fStartPointsAgain`, `ac2fStartPointsShow` (ana menü `22`-`26`).
Nesting ayarlarınız korunur; ayar sayfasında NEST grubundan sonra yeni
START grubu geldiği için PLOT, BEND ve JOB ayarlarının **numaraları 19
kaydı** (ör. `Send order` 47 → 66). Nesting penceresi iş menüsündeki
liste penceresi gibi PowerShell ile açılır; o kapalıysa (`List window`
ayarı 0) kısa metin sorusu çıkar.

### Sürüm 1.15.0'dan 1.16.0'a

Yalnız `ac2fJob` ve `ac2fCore` değişti — ikisini silip yeniden aktarın.
StockTrack çok dükkânlı olduysa: her dükkânın anahtarı ayrıdır; bu
bilgisayarın çalıştığı dükkânın *Ayarlar › Corel entegrasyonu*'ndaki
anahtarı iş menüsü `7` → `2=` ile girin. Eski anahtar ilk dükkânın
anahtarı olarak çalışmaya devam eder.

### Sürüm 1.14.0'dan 1.15.0'a

`ac2fCheck.bas`, `ac2fCutTime.bas`, `ac2fCutArea.bas` **yeni** —
doğrudan aktarın. `ac2fJob`, `ac2fSettings`, `ac2fMenu`, `ac2fCore`
değişti — dördünü silip yeniden aktarın. Yeni ayar `Check first` (61).
StockTrack'i güncelleyip *Ayarlar › Kesim fiyatlandırması*'nı doldurun.

### Sürüm 1.13.0'dan 1.14.0'a

`ac2fJob`, `ac2fSettings` ve `ac2fCore` değişti — üçünü silip yeniden
aktarın. Yeni ayar `Look back` (60); diğer numaralar değişmedi.

### Sürüm 1.12.0'dan 1.13.0'a

`ac2fJob` ve `ac2fSettings` değişti — ikisini silip yeniden aktarın.
StockTrack'i de güncelleyin (teklif uçları). İş menüsünde kurulum artık
`7`. `Round up` ayarı kalktı; diğer ayar numaraları değişmedi.

### Sürüm 1.11.0'dan 1.12.0'a

`ac2fJob.bas` **yeni** — doğrudan aktarın. `ac2fSettings`, `ac2fMenu`,
`ac2fCore` değişti — üçünü silip yeniden aktarın. Ayar sayfasına `JOB`
grubu eklendi (56-60); diğer ayar numaraları değişmedi.

StockTrack bağlantısı için StockTrack'i de güncelleyin, *Ayarlar ›
Corel entegrasyonu*'ndan anahtar oluşturun ve Corel'de ana menü `21` →
`7` (Job setup) ile adresi ve anahtarı girin. Ayrıntı:
[kullanım](kullanim.md#13-müşteri-işi--stocktrack-ile-otomatik-adlandırma).

### Sürüm 1.10.0'dan 1.11.0'a

`ac2fPlot`, `ac2fBending`, `ac2fSettings`, `ac2fCore` değişti — dördünü
silip yeniden aktarın. Plotter'a iki ayar eklendi (`Send order` 47,
`Park` 48), bu yüzden bükme ayarları **49-55**'e kaydı. `Rotate` yine
42, `Mirror` 43.

### Sürüm 1.9.0'dan 1.10.0'a

1.9.0 ayar dosyası VBA'nın **25 satır birleştirme** sınırını aşıyordu
ve "Too many line continuations" ile içe aktarılamıyordu; düzeltildi
(ve `tools/lint.py` artık bunu yakalıyor).

Bükme sırası `ac2fPlot`'tan çıkıp **`ac2fBending`** adlı kendi modülüne
taşındı. `ac2fBending.bas` **yeni** — doğrudan aktarın. `ac2fPlot`,
`ac2fSettings`, `ac2fMenu`, `ac2fCore` değişti — dördünü **silip**
yeniden aktarın.

Sıralama ayarlarının kayıt anahtarları `PL...`'den `BN...`'e döndüğü
için o beş ayar varsayılanına döner; diğer ayarlarınız ve profilleriniz
korunur. Ayar numaraları: plotter yine **40-46** (`Rotate` 42,
`Mirror` 43), bükme **47-53**.

Ana menü: bükme `12`-`14`, plotter `15`-`18`, ayarlar `19`,
hakkında `20`.

### Sürüm 1.8.1'den 1.9.0'a

`ac2fPlot`, `ac2fSettings`, `ac2fMenu`, `ac2fCore` değişti — dördünü
silip yeniden aktarın. Kesim sırası ve gruplama geldi. Bu sürüm içe
aktarılamıyor; 1.10.0'a geçin.

### Sürüm 1.8.0'dan 1.8.1'e

`ac2fPlot`, `ac2fSettings`, `ac2fCore` değişti — üçünü silip yeniden
aktarın. Yeni `Mirror` ayarı, test şekli artık asimetrik bir L.

### Sürüm 1.7.x'ten 1.8.0'a

| Dosya | Ne yapmalı |
|---|---|
| `ac2fPlot.bas` | **Değişti** — önce silin, sonra aktarın |
| `ac2fSettings.bas` | **Değişti** — önce silin, sonra aktarın |
| `ac2fMenu.bas` | **Değişti** — önce silin, sonra aktarın |
| `ac2fCore.bas` | **Değişti** (sürüm) — önce silin, sonra aktarın |
| Diğerleri | Değişmedi |

Yeni makro: `ac2fPlotTest` (ana menü 13). Plotter menü numaraları
kaydı: gönder `12`, test `13`, dosyaya yaz `14`, `.plt` düzelt `15`,
ayarlar `16`, hakkında `17`.

### Sürüm 1.6.x'ten 1.7.0'a

| Dosya | Ne yapmalı |
|---|---|
| `ac2fNest.bas` | **Yeni** — doğrudan içe aktarın |
| `ac2fCore.bas` | **Değişti** — önce silin, sonra aktarın |
| `ac2fSettings.bas` | **Değişti** — önce silin, sonra aktarın |
| `ac2fMenu.bas` | **Değişti** — önce silin, sonra aktarın |
| Diğerleri | Değişmedi |

Yeni makrolar: `ac2fNest`, `ac2fNestReport`. Ayarlarınız korunur.
Ana menü: nesting `10`-`11`, plotter `12`-`14`, ayarlar `15`,
hakkında `16`.

Ayar sayfası artık gruba göre sayfalı — `#n` ile geçilir, ayar
numaraları her sayfadan çalışır.

### Sürüm 1.5.0'dan 1.6.0'a

| Dosya | Ne yapmalı |
|---|---|
| `ac2fPlot.bas` | **Yeni** — doğrudan içe aktarın |
| `ac2fCore.bas` | **Değişti** — önce silin, sonra aktarın |
| `ac2fSettings.bas` | **Değişti** — önce silin, sonra aktarın |
| `ac2fMenu.bas` | **Değişti** — önce silin, sonra aktarın |
| `ac2fLength`, `ac2fLedModule`, `ac2fBoxLetter`, `ac2fPanel`, `ac2fCenterline` | Değişmedi |

Yeni makrolar: `ac2fPlotSend`, `ac2fPlotSave`, `ac2fPlotFixSend`.
Ayarlarınız korunur. Ana menü: plotter `10`-`12`, ayarlar `13`,
hakkında `14`.

**Plotter göndermek için ek gereksinim yok** — PowerShell Windows'ta
zaten var. Kurumsal bir makinede PowerShell betik çalıştırma
kapatılmışsa makro bunu hata olarak bildirir; betik
`-ExecutionPolicy Bypass` ile çağrılır, bu çoğu kısıtlamayı aşar.

### Sürüm 1.4.0'dan 1.5.0'a

| Dosya | Ne yapmalı |
|---|---|
| `ac2fCenterline.bas` | **Yeni** — doğrudan içe aktarın |
| `ac2fCore.bas` | **Değişti** — önce silin, sonra aktarın |
| `ac2fSettings.bas` | **Değişti** — önce silin, sonra aktarın |
| `ac2fMenu.bas` | **Değişti** — önce silin, sonra aktarın |
| `ac2fLength`, `ac2fLedModule`, `ac2fBoxLetter`, `ac2fPanel` | Değişmedi |

Yeni makrolar: `ac2fCenterline`, `ac2fCenterlineJoined`,
`ac2fCenterlineSeparate`. Ayarlarınız korunur.
Ana menü: orta hat `9`, ayarlar `10`, hakkında `11`.

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
| `Expected variable or procedure, not module` | 1.21.0'dan eski bir `ac2fNest`, `ac2fMenu` ya da `ac2fJob` duruyor; bu modülleri silip yeniden aktarın |
| İçe aktarınca kodda `Attribute VB_Name` satırı görünüyor, modülün adı `Module1` | Dosya LF satır sonlu gelmiş (tek dosya "raw" indirildiyse olur); `tools\build.ps1` ile CRLF'e çevirin ya da depoyu ZIP/klon olarak alın |
| Nesting penceresi açılmıyor, metin sorusu geliyor | PowerShell engelli ya da iş menüsündeki *List window* ayarı `0`; Windows 10/11'deki Windows PowerShell 5.1 gerekir |
| Makro listede görünmüyor | Makro yalnızca **Public Sub** ve **parametresiz** ise listelenir; `Debug > Compile` ile hata olup olmadığına bakın |
| "Makrolar devre dışı" uyarısı | `Araçlar > Makrolar > Güvenlik` → güven düzeyini düşürün |
