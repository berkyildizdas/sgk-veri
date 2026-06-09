# Günlük SGK Veri Güncellemesi — Kurulum

Bu sistem, SGK'nın resmi **EK-4/A "Bedeli Ödenecek İlaçlar Listesi"** dosyasını her gün
otomatik kontrol eder, değişiklik varsa temiz JSON olarak yayımlar ve eczacı uygulaması
bu veriyi günlük indirir. Böylece güncel olmayan veriden kaynaklı ceza riski azalır.

## Mimari

```
SGK EK-4/A (.xlsx)  ──günlük──>  GitHub Actions cron  ──>  sgk-data/*.json (repo)
                                                              │ raw.githubusercontent
                                          Eczacı uygulaması  <┘  günlük indirir + uygular
```

- **Maliyet:** Sıfır (GitHub Actions ücretsiz kotada).
- **Bakım:** SGK format değiştirirse yalnızca `scripts/sgk-veri-cek.mjs` güncellenir; eczacılara yeni exe gerekmez.

## Kurulum adımları

### 1) GitHub deposu
1. GitHub'da bir depo aç (public önerilir; private ise raw erişim için token gerekir).
2. Bu projeyi (en azından `scripts/`, `.github/workflows/` ve `sgk-data/` klasörleri) depoya push'la.

### 2) SGK EK-4/A linkini ayarla
SGK dosya linki GUID'li ve değişkendir. En güvenilir yol, güncel linki **Secret** olarak vermek:
1. Depo > **Settings > Secrets and variables > Actions > New repository secret**
2. `SGK_EK4A_URL` = SGK'daki güncel EK-4/A `.xlsx` doğrudan linki
   - Linki bulmak: sgk.gov.tr > Duyurular > "Bedeli Ödenecek İlaçlar Listesi" duyurusu > ekteki `.xlsx`
3. (Opsiyonel) `SGK_INDEX_URL` = otomatik link bulmak için taranacak duyuru sayfası

> Not: SGK her gün değil, değişiklik tebliğleriyle (~aylık) güncellenir. Cron günlük kontrol eder; çoğu gün "değişiklik yok" döner.

### 3) Workflow'u çalıştır
- **Actions** sekmesi > "SGK Veri Güncelle" > **Run workflow** (ilk veriyi üretmek için elle tetikle).
- Sonrasında her gün 06:00 UTC'de otomatik çalışır.
- Üretilen dosyalar: `sgk-data/version.json`, `sgk-data/ilaclar.json`.

### 4) Uygulamayı veri kaynağına bağla
Uygulamada **Ayarlar > Veri Kaynağı** alanına raw taban adresini gir:
```
https://raw.githubusercontent.com/KULLANICI/REPO/main/sgk-data
```
Uygulama açılışta + günde bir kez bu adresteki `version.json`'u kontrol eder; yeni sürüm
varsa `ilaclar.json`'u indirip yerel olarak saklar.

## Güvenlik ilkesi
- **İlaç listesi (EK-4/A):** Yapısal Excel → güvenle otomatik uygulanır.
- **Kural eşikleri (LDL, HbA1c vb.):** Yasal metin yanlış parse edilirse uygulamanın kendisi
  hata kaynağı olur. Bu yüzden eşik değişiklikleri **incele-sonra-uygula** akışındadır
  (otomatik körlemesine uygulanmaz). Kesin ödeme kararı için **MEDULA** esastır; bu uygulama
  ön kontrol/karar destek amaçlıdır.

## Yerel test
```
npm install xlsx --no-save
# Elle indirdiğin bir EK-4/A dosyasıyla test:
$env:SGK_EK4A_URL="file:///C:/yol/ek4a.xlsx"   # veya gerçek SGK linki
node scripts/sgk-veri-cek.mjs
# Çıktı: sgk-data/ilaclar.json + version.json
```
