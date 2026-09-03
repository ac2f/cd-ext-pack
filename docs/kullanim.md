# Kullanım

## Ana menü

`ac2fPack` makrosunu çalıştırın; işlem numarasını soran bir kutu açılır.

```
  1  -  Uzunluk ölçümü
  2  -  Uzunluk ölçümü + sayfaya etiket
  3  -  3'lü LED modül hesabı
  4  -  3'lü LED modül hesabı (aralığı sorarak)
  5  -  Ayarlar
  6  -  Hakkında
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

## 4. Ayarlar

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
