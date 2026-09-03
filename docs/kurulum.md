# Kurulum

## 0. Ön koşul: VBA desteği

CorelDRAW kurulumunda **Visual Basic for Applications** bileşeni seçili
olmalıdır. `Alt+F11` bir pencere açmıyorsa:

Windows *Ayarlar > Uygulamalar* → CorelDRAW → *Değiştir* → *Modify* →
**Visual Basic for Applications** işaretle → kurulumu tamamla.

## 1. Kaynağı hazırla

`.bas` dosyaları depoda UTF-8 + LF tutulur; VBA'nın içe aktarıcısı ise ANSI
bekler. Türkçe karakterlerin bozulmaması için önce dönüştürün:

```powershell
powershell -ExecutionPolicy Bypass -File tools\build.ps1
```

`build\` klasöründe dört dosya oluşur:

```
ac2fCore.bas
ac2fLength.bas
ac2fLedModule.bas
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

`ac2fPack` projesi seçiliyken, dört dosyanın **her biri** için:

`File > Import File...` → `build\ac2*.bas` → **Open**

Sonuçta proje ağacı şöyle görünmeli:

```
ac2fPack
└── Modules
    ├── ac2fCore
    ├── ac2fLength
    ├── ac2fLedModule
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

Yeni sürüm geldiğinde:

1. `tools\build.ps1` çalıştırın.
2. VBE'de eski modülleri sağ tıklayıp `Remove ac2f...` (dışa aktarma
   sorusuna **No**) ile silin.
3. Yeni `.bas` dosyalarını içe aktarın, derleyip kaydedin.

Ayarlarınız kayıt defterinde tutulduğu için güncellemede korunur.

## Sorun giderme

| Belirti | Çözüm |
|---|---|
| `Alt+F11` açılmıyor | VBA bileşeni kurulu değil (bkz. adım 0) |
| Türkçe harfler bozuk görünüyor | `build\` yerine `src\` içindekiler aktarılmış; adım 1'i uygulayın |
| `User-defined type not defined` | `ac2fCore` içe aktarılmamış; önce onu ekleyin |
| Makro listede görünmüyor | Makro yalnızca **Public Sub** ve **parametresiz** ise listelenir; `Debug > Compile` ile hata olup olmadığına bakın |
| "Makrolar devre dışı" uyarısı | `Araçlar > Makrolar > Güvenlik` → güven düzeyini düşürün |
