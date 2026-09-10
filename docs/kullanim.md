# Kullanım

## Ana menü

`ac2fPack` makrosunu çalıştırın; işlem numarasını soran bir kutu açılır.

```
  1  -  Uzunluk ölçümü
  2  -  Uzunluk ölçümü + sayfaya etiket
  3  -  3'lü LED modül hesabı
  4  -  3'lü LED modül hesabı (aralığı sorarak)
  5  -  Kutu harf şeridi (hesapla ve çiz)
  6  -  Kutu harf raporu (çizim yok)
  7  -  Ayarlar - LED modül
  8  -  Ayarlar - Kutu harf
  9  -  Hakkında
```

Her makro doğrudan da çağrılabilir; ana menü yalnızca kolaylık içindir.

---

## 1. Uzunluk ölçümü

Nesneleri seçin ve `ac2fUzunlukOlc` çalıştırın.

```
TOPLAM UZUNLUK
   3.842,17 mm   |   384,22 cm   |   3,842 m

SEÇİM ÖZETİ
   Nesne sayısı        : 7
   Kontur (alt yol)    : 12
      kapalı           : 12
      açık             : 0

KONTUR İSTATİSTİĞİ
   En uzun             : 812,44 mm
   En kısa             : 96,10 mm
   Ortalama            : 320,18 mm

NESNE KIRILIMI
   1. Curve 1 - 812,44 mm (2 kontur)
   ...
```

### Neler ölçülür

- Eğriler, dikdörtgen/elips/çokgen gibi şekiller, metinler
- **Gruplar** — içindeki her nesne ayrı ayrı
- **PowerClip** içerikleri

Eğri olmayan nesneler ölçüm sırasında geçici bir kopya üzerinden eğriye
çevrilir; kopyalar silinir ve belge ilk hâline döndürülür. Çiziminizde
kalıcı bir değişiklik olmaz.

### Ölçülemeyenler

Bitmap'ler, boş metinler ve uzunluğu sıfır olan nesneler sessizce atlanır.

> **Kontur = alt yol.** Bir harfin dışı ve içi (örneğin "O") iki ayrı
> konturdur ve ikisi de toplama girer.

---

## 2. Sayfaya etiket

`ac2fUzunlukEtiketle` aynı ölçümü yapar, ardından seçimin 8 mm altına
artistik metin olarak sonucu yazar:

```
Toplam uzunluk: 3.842,17 mm  (3,84 m)
```

Etiket normal bir metin nesnesidir; taşıyabilir, biçimlendirebilir,
silebilirsiniz.

---

## 3. 3'lü LED modül hesabı

Nesneleri seçin ve `ac2fLedModulHesapla` çalıştırın.

```
SONUÇ
   Modül adedi         : 42 adet
   LED adedi           : 126 adet  (3 LED/modül)

GÜÇ
   Toplam güç          : 30,24 W
   %20 pay ile         : 36,29 W
   Güç kaynağı         : 1 adet x 60 W

ÖLÇÜM
   Ham kontur uzunluğu : 3.842,17 mm
   Hesaba giren uzunluk: 3.842,17 mm
   Kontur sayısı       : 12

KULLANILAN AYARLAR
   Modül aralığı       : 100,00 mm
   Yöntem              : Çevre bazlı (tam kontur)
   Düzeltme katsayısı  : 1,00
   Modül gücü          : 0,72 W
```

`ac2fLedHizliHesap` aynı işi yapar ama modül aralığını o seferlik sorar ve
kayıtlı ayarı değiştirmez. Farklı aralıkları hızlıca denemek için kullanışlıdır.

### Hesap nasıl yapılır

Her kontur **ayrı ayrı** hesaplanır — böylece kısa parçalar da modülsüz kalmaz:

| Kontur türü | Modül adedi |
|---|---|
| Kapalı (halka) | `yukarı_yuvarla(uzunluk / aralık)` |
| Açık (uçlu yol) | `aşağı_yuvarla(uzunluk / aralık) + 1` |

Açık yolda iki uca da modül konduğu için bir fazla çıkar. Her kontur en az
*"kontur başına en az modül"* kadar modül alır.

Konturların toplamı düzeltme katsayısıyla çarpılıp yukarı yuvarlanır:

```
LED adedi   = modül × modül başına LED
Toplam güç  = modül × modül gücü
Gerekli güç = toplam güç × (1 + güvenlik payı / 100)
Güç kaynağı = yukarı_yuvarla(gerekli güç / kaynak kapasitesi)
```

### Hesap yöntemleri

**1 — Çevre bazlı (varsayılan).** Her konturun tam uzunluğu kullanılır.
Modüller çizginin *üzerine* dizildiğinde doğrudur: kanal harflerin yan
duvarına yapıştırılan çevre aydınlatması, ışıklı şerit güzergâhı, neon
flex boyu gibi işler.

**2 — Orta hat tahmini.** Her konturun **yarısı** kullanılır. Harf, iç ve dış
olmak üzere iki çizgiyle çizilmiş kapalı bir şerit ise (içi boş kontur çifti)
çevre gerçek yolun yaklaşık iki katını verir; bu yöntem bunu düzeltir.

> Yöntem 2 bir **tahmindir**. Şerit genişliği her yerde eşitse iyi sonuç verir,
> değişken kalınlıkta sapar. İlk işinizde çıkan sayıyı sahadaki gerçek
> adetle karşılaştırıp farkı **düzeltme katsayısına** yazın; sonraki hesaplar
> tutar.

---

## 4. Kutu harf şeridi

Harf konturlarını seçin ve `ac2fKutuHarfSerit` çalıştırın. Çizimin sağ
tarafına, her kapalı kontur için bir düz şerit çizilir.

```
SONUÇ
   Şerit sayısı        : 3
   Toplam açınım       : 1.842,66 mm   |   184,27 cm   |   1,843 m
   Toplam derz         : 96 adet  (8 köşe)
   En küçük yarıçap    : 12,40 mm
   Rulo ihtiyacı       : ~1 parça x 3.000 mm (20 mm ek payı)

ŞERİTLER
   1. Curve 1 #1
      açınım 1.204,18 mm (ham 1.211,22)  derz 58  min R 18,7
   2. Curve 1 #2 [delik]
      açınım  638,48 mm (ham  631,44)  derz 38  min R 12,4
```

Çizimdeki renkler:

| Renk | Anlamı |
|---|---|
| Siyah | Şerit dış hattı (kesim) |
| Mavi | Eğri derzi |
| Pembe | Köşe derzi |
| Kırmızı | Rulo boyu aşıldığında kesim/ek yeri |

`ac2fKutuHarfRapor` aynı hesabı yapar ama hiçbir şey çizmez — ayar denemek
için hızlıdır.

> Şerit yalnızca **kapalı** konturlardan çıkarılır. Açık yollar atlanır.

### Açınım boyu nereden geliyor

Sac büküldüğünde uzunluğu koruyan çizgi nötr eksendir; malzemenin dış yüzü
uzar, iç yüzü kısalır. Vektörünüz şeridin hangi yüzünü temsil ediyorsa,
nötr eksen ondan `g` kadar ötededir:

| Referans yüzey | g |
|---|---|
| Vektör = şeridin dış yüzü *(varsayılan)* | `(1 − K) × kalınlık` |
| Vektör = şeridin iç yüzü | `−K × kalınlık` |
| Vektör = nötr eksen | `0` |

Basit kapalı bir eğride toplam dönüş her zaman 2π olduğu için açınım boyu

```
açınım = kontur boyu − 2π × g        (dış kontur)
açınım = kontur boyu + 2π × g        (delik / counter)
```

Delik konturları otomatik bulunur (sınırlayıcı kutusu bir başkasının içinde
kalan alt yol deliktir) ve raporda `[delik]` diye işaretlenir.

**Büyüklük hissi:** 1 mm alüminyumda düzeltme ≈ 3,5 mm; 3 mm'de ≈ 10,6 mm.
Kontur boyundan bağımsızdır — yalnız kalınlığa bağlıdır.

### Derz aralığı nereden geliyor

Her segment için üç sınırın en küçüğü alınır:

| Sınır | Formül | Neyi engeller |
|---|---|---|
| Derz kapanması | `R × (derz ağzı / derz derinliği) × esneklik` | Derz kapanmadan malzemenin sıkışması |
| Yüzey düzlüğü | `√(8 × R × yüzey toleransı)` | Derzler arası düz yüzün göze çarpması |
| Tavan | `en çok derz aralığı` | Çok seyrek derz |

Sonuç `en az derz aralığı`na kırpılır. Köşelerde (dönüş > köşe eşiği) dönüş
açısı tek derzin karşılayabileceğinden büyükse birden çok derz açılır.

`R` her segmentin yarıçapıdır; düz segmentlere derz açılmaz.

### Esneklik oranını kalibre etmek

Ön ayarlar (alüminyum, galvaniz, kalın alüminyum, paslanmaz) **ölçülmüş
değer değildir** — başlangıç noktasıdır. Doğru yol:

1. Bir harf için şeridi çıkarın ve gerçekten bükün.
2. Derzler kapanmıyor, malzeme sıkışıyorsa → esneklik oranını **küçültün**
   (daha sık derz).
3. Gereğinden fazla derz varsa, kontur zaten rahat dönüyorsa → **büyütün**.
4. Yüzeyde köşeli izler görünüyorsa esnekliğe dokunmayın, **yüzey
   toleransını** küçültün.

Bulduğunuz değer o malzeme + kalınlık için kalıcıdır.

### Bu modülün sınırları

- **Segmentler dairesel yay kabul edilir.** Gerçek bezier yaylarda yarıçap
  hatası %0,05 mertebesindedir — ihmal edilebilir. Ancak tek bir segment
  içinde dönüm noktası varsa (S kıvrımı) model onu tek yönlü yay sanar ve o
  segmentte derzi seyrek koyar. Böyle bir yerde düğüm ekleyip segmenti
  ikiye bölmek sorunu çözer.
- **Toplam açınım boyu her koşulda kesindir**; yukarıdaki durum yalnız ara
  derz konumlarını etkiler (ölçülen en kötü sapma 0,06 mm).
- **Şerit tek parça çizilir.** Rulo boyu aşılıyorsa kesim yerleri kırmızı
  işaretlenir; parçalar ayrı ayrı çizilmez.
- **Derz kesiti çizilmez** — konumu çizilir. V ağzının açısı ve derinliği
  makinenizin/bıçağınızın işidir.
- Çok yoğun işlerde şerit başına 5.000, iş başına 200 şerit sınırı vardır.

---

## 5. Ayarlar

`ac2fAyarlar` ayarları sırayla sorar. Herhangi bir adımda **İptal** derseniz
o adıma kadar girdikleriniz kaydedilmiş olur, kalanı değişmez.

| Ayar | Varsayılan | Anlamı |
|---|---|---|
| Modül aralığı | 100 mm | İki modül arası mesafe |
| Modül başına LED | 3 | 3'lü modüller için 3 |
| Modül gücü | 0,72 W | Tek modülün çektiği güç |
| Güç kaynağı kapasitesi | 60 W | `0` girilirse kaynak adedi hesaplanmaz |
| Güvenlik payı | %20 | Kaynak seçiminde eklenen pay |
| Kontur başına en az modül | 1 | Kısa parçaların boş kalmaması için |
| Hesap yöntemi | 1 | `1` = çevre, `2` = orta hat |
| Düzeltme katsayısı | 1 | Sonucu ölçekler |

### Kutu harf ayarları (`ac2fKutuHarfAyarlar`)

Önce bir malzeme ön ayarı sorulur, sonra temel değerler. Gelişmiş ayarlar
ayrıca istenir.

| Ayar | Varsayılan | Anlamı |
|---|---|---|
| Malzeme kalınlığı | 1 mm | Sacın kalınlığı |
| Şerit yüksekliği | 80 mm | Harfin derinliği |
| Esneklik oranı | 1 | Büyük = daha az derz |
| Referans yüzey | 1 | Vektör hangi yüz: 1 dış, 2 iç, 3 nötr |
| K faktörü | 0,44 | Nötr eksenin kalınlık içindeki konumu |
| Derz derinlik oranı | 0,7 | Derz derinliği / kalınlık |
| Derz ağzı en çok | 1,2 mm | Kapandığında iz bırakmayan en geniş ağız |
| Yüzey toleransı | 0,15 mm | Derzler arası düz yüzün eğriden sapması |
| Derz aralığı | 3 - 60 mm | Alt ve üst sınır |
| Köşe eşiği | 5° | Üstündeki dönüş köşe sayılır |
| Rulo boyu | 3000 mm | 0 = bölme yapma |
| Ek payı | 20 mm | Parça eklerinde bindirme |
| Şeritler arası boşluk | 10 mm | Çizim yerleşimi |

Ayarlar Windows kayıt defterinde tutulur:

```
HKCU\Software\VB and VBA Program Settings\ac2fPack\Ayarlar
```

Kullanıcıya özeldir, CorelDRAW güncellemelerinden ve paket güncellemelerinden
etkilenmez. `ac2fAyarlariSifirla` hepsini varsayılana döndürür.

### Sayı girişi

Hem `12,5` hem `12.5` kabul edilir. **Binlik ayırıcı kullanmayın** — `1.250`
bin iki yüz elli değil, bir virgül iki yüz elli olarak okunur.

---

## Bilinmesi gerekenler

- Ölçüm **her zaman milimetre** üzerinden yapılır; belgenizin birimi ne
  olursa olsun sonuç doğrudur ve işlem sonunda birim geri alınır.
- Kontur uzunluğu, çizginin **kalınlığını dikkate almaz** — yolun kendi
  uzunluğudur.
- Çok sayıda nesne seçildiğinde ölçüm birkaç saniye sürebilir; ekran
  tazeleme bu sırada kapatılır.
- Nesne kırılımı listesinde en fazla 15 satır gösterilir; toplam her zaman
  seçimin tamamını kapsar.
