# Değişiklik Günlüğü

Bu proje [Semantic Versioning](https://semver.org/lang/tr/) kullanır.

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
