# EK-4/D Muafiyet Kuralları Nasıl Güncellenir

`sgk-data/muafiyet-kurallari.json` = uygulamanın katılım payı muafiyeti (EK-4/D)
kararlarında kullandığı kurallar. **Bu dosyayı güncelleyince tüm eczanelerin
uygulaması ertesi gün otomatik alır — yeni exe/dağıtım GEREKMEZ.**

## Ne zaman güncellenir?
Uygulama, SGK EK-4/D listesi değiştiğinde "⚠️ muafiyet listesi güncellenmiş olabilir"
uyarısı gösterir. O zaman:

## Adımlar
1. GitHub'da `sgk-data/muafiyet-kurallari.json` dosyasını aç → kalem (✏️ Edit) ikonu
2. İlgili hastalık grubunu bul ve düzenle. Örnek alanlar:
   - `icd10_kodlari`: muaf ICD-10 kodları listesi
   - `yetkili_branslar`: rapor yazabilecek uzman branşlar
   - `ek4d_kapsaminda`: katılım payından muaf mı (true/false)
   - `uzmanlik_sarti`: uzman raporu zorunlu mu (true/false)
   - `kural_ozeti`, `ozel_notlar`: açıklama metinleri
3. **Commit changes** (yeşil buton) → kaydet
4. Bitti. Uygulama içerik değişikliğini otomatik algılar (`hash` alanını elle
   değiştirmene gerek yok — uygulama hash'i içerikten kendisi hesaplar).

## Önemli
- **Eczacı hiçbir şey yapmaz** — değişiklik merkezi olur, otomatik dağılır.
- **Yapı bozulmamalı:** JSON formatına dikkat (virgüller, tırnaklar). GitHub editör
  hata gösterirse düzeltmeden commit etme.
- Emin değilsen bir yedek tut (eski içeriği bir yere kopyala).
- Bu kurallar **karar-destek** içindir; kesin ödeme kararı MEDULA'dadır.
