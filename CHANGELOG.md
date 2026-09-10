# Değişiklik Günlüğü

Bu proje [Semantic Versioning](https://semver.org/lang/tr/) kullanır.

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
