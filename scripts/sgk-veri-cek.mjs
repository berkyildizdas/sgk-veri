#!/usr/bin/env node
/**
 * SGK Veri Çekme Pipeline — GitHub Actions cron tarafından günlük çalıştırılır.
 *
 * Görev:
 *  1. SGK'nın resmi EK-4/A "Bedeli Ödenecek İlaçlar Listesi" (.xlsx) dosyasını indir
 *  2. Excel'i temiz JSON'a çevir (barkod, ad, etken madde, ödeme durumu, referans fiyat)
 *  3. Hash ile değişiklik tespit et; değiştiyse version.json + ilaclar.json yayımla
 *
 * GÜVENLİK: Dosya bulunamaz/parse edilemezse ESKİ veriyi KORUR, asla boş yayımlamaz.
 *
 * Yapılandırma (ortam değişkenleri):
 *  - SGK_EK4A_URL   : EK-4/A xlsx dosyasının doğrudan linki (en güvenilir yol)
 *  - SGK_INDEX_URL  : (opsiyonel) linki otomatik bulmak için taranacak SGK sayfası
 *
 * Çıktı: sgk-data/ilaclar.json, sgk-data/version.json
 */

import { readFileSync, writeFileSync, mkdirSync, existsSync } from "node:fs";
import { createHash } from "node:crypto";
import { fileURLToPath } from "node:url";
import { dirname, join } from "node:path";
import * as XLSX from "xlsx";

const __dirname = dirname(fileURLToPath(import.meta.url));
const OUT_DIR = join(__dirname, "..", "sgk-data");

// SGK SUT ekleri sayfası — EK-4/A linkini otomatik bulmak için varsayılan taranacak sayfa
const VARSAYILAN_INDEX =
  "https://www.sgk.gov.tr/Duyuru/Index?page=1";

const UA = { "User-Agent": "Mozilla/5.0 (SGK-Eczane-Veri-Bot)" };

function log(...a) { console.log("[sgk-veri]", ...a); }

/* ─── EK-4/A xlsx linkini bul ─── */
async function ek4aLinkiniBul() {
  // 1) Doğrudan URL verilmişse onu kullan (en güvenilir)
  if (process.env.SGK_EK4A_URL) {
    log("Doğrudan URL kullanılıyor:", process.env.SGK_EK4A_URL);
    return process.env.SGK_EK4A_URL;
  }
  // 2) Index sayfasını tara, "Ek-4/A" veya "Bedeli Ödenecek" geçen .xlsx linkini bul
  const indexUrl = process.env.SGK_INDEX_URL || VARSAYILAN_INDEX;
  log("Index taranıyor:", indexUrl);
  const res = await fetch(indexUrl, { headers: UA });
  if (!res.ok) throw new Error(`Index sayfası alınamadı: ${res.status}`);
  const html = await res.text();
  // .xlsx linklerini ve civarındaki metni eşle
  const adaylar = [...html.matchAll(/href="([^"]+\.xlsx[^"]*)"/gi)].map((m) => m[1]);
  const ek4a = adaylar.find((u) => /ek.?4.?a|bedeli.?odenecek|bedeli%20odenecek/i.test(u)) || adaylar[0];
  if (!ek4a) throw new Error("EK-4/A xlsx linki bulunamadı (SGK_EK4A_URL ortam değişkeni ile elle verin).");
  const tam = ek4a.startsWith("http") ? ek4a : new URL(ek4a, "https://www.sgk.gov.tr").href;
  log("Bulunan link:", tam);
  return tam;
}

/* ─── Excel → ilaç JSON ─── */
function excelParse(buffer) {
  const wb = XLSX.read(buffer, { type: "buffer" });
  const sheet = wb.Sheets[wb.SheetNames[0]];
  const satirlar = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: "" });

  // Başlık satırını bul (barkod / ilaç adı geçen)
  let baslikIdx = satirlar.findIndex((r) =>
    r.some((c) => /barkod/i.test(String(c))) && r.some((c) => /ad|isim/i.test(String(c)))
  );
  if (baslikIdx === -1) baslikIdx = 0;
  const basliklar = satirlar[baslikIdx].map((c) => String(c).toLowerCase().trim());

  const kolon = (anahtarlar) =>
    basliklar.findIndex((b) => anahtarlar.some((a) => b.includes(a)));

  const iBarkod = kolon(["barkod"]);
  const iAd = kolon(["ilaç adı", "ilac adi", "ürün adı", "urun adi", "ad"]);
  const iEtken = kolon(["etkin madde", "etken madde", "etkin"]);
  const iFiyat = kolon(["fiyat", "referans"]);
  const iDurum = kolon(["ödeme", "odeme", "durum", "aktif"]);

  const ilaclar = [];
  for (let i = baslikIdx + 1; i < satirlar.length; i++) {
    const r = satirlar[i];
    const ad = iAd >= 0 ? String(r[iAd] ?? "").trim() : "";
    const barkod = iBarkod >= 0 ? String(r[iBarkod] ?? "").trim() : "";
    if (!ad && !barkod) continue;
    ilaclar.push({
      barkod,
      ad,
      etken_madde: iEtken >= 0 ? String(r[iEtken] ?? "").trim() : "",
      referans_fiyat: iFiyat >= 0 ? String(r[iFiyat] ?? "").trim() : "",
      odeme_durumu: iDurum >= 0 ? String(r[iDurum] ?? "").trim() : "",
    });
  }
  return ilaclar;
}

/* ─── Ana akış ─── */
async function main() {
  if (!existsSync(OUT_DIR)) mkdirSync(OUT_DIR, { recursive: true });
  const versionPath = join(OUT_DIR, "version.json");
  const ilaclarPath = join(OUT_DIR, "ilaclar.json");

  const eskiVersion = existsSync(versionPath)
    ? JSON.parse(readFileSync(versionPath, "utf8"))
    : { tarih: "", hash: "", kayit_sayisi: 0 };

  let link, buffer;
  try {
    link = await ek4aLinkiniBul();
    const res = await fetch(link, { headers: UA });
    if (!res.ok) throw new Error(`xlsx indirilemedi: ${res.status}`);
    buffer = Buffer.from(await res.arrayBuffer());
  } catch (e) {
    log("HATA (indirme):", e.message);
    log("Eski veri korunuyor, yayımlama yapılmadı.");
    process.exit(0); // CI başarısız sayılmasın; sadece bu sefer güncelleme yok
  }

  let ilaclar;
  try {
    ilaclar = excelParse(buffer);
  } catch (e) {
    log("HATA (parse):", e.message);
    log("Eski veri korunuyor.");
    process.exit(0);
  }

  if (!ilaclar.length) {
    log("Parse sonucu 0 kayıt — şüpheli. Eski veri korunuyor.");
    process.exit(0);
  }

  const hash = createHash("sha256").update(JSON.stringify(ilaclar)).digest("hex").slice(0, 16);
  if (hash === eskiVersion.hash) {
    log(`Değişiklik yok (${ilaclar.length} kayıt, hash ${hash}). Güncelleme gerekmiyor.`);
    process.exit(0);
  }

  const bugun = new Date().toISOString().split("T")[0];
  const version = {
    tarih: bugun,
    hash,
    kayit_sayisi: ilaclar.length,
    kaynak_url: link,
    onceki_tarih: eskiVersion.tarih || null,
  };

  writeFileSync(ilaclarPath, JSON.stringify(ilaclar), "utf8");
  writeFileSync(versionPath, JSON.stringify(version, null, 2), "utf8");
  log(`YAYIMLANDI: ${ilaclar.length} ilaç, tarih ${bugun}, hash ${hash}`);
}

main().catch((e) => { console.error("[sgk-veri] KRİTİK:", e); process.exit(1); });
