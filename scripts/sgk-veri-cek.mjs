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

/* Türkçe başlık normalizasyonu (İ/ı/ş/ğ/ç/ö/ü + birleşik nokta) */
function trNorm(s) {
  return String(s)
    .replace(/̇/g, "")
    .replace(/İ/g, "i").replace(/I/g, "i").replace(/ı/g, "i")
    .replace(/Ş/g, "s").replace(/ş/g, "s")
    .replace(/Ğ/g, "g").replace(/ğ/g, "g")
    .replace(/Ü/g, "u").replace(/ü/g, "u")
    .replace(/Ö/g, "o").replace(/ö/g, "o")
    .replace(/Ç/g, "c").replace(/ç/g, "c")
    .toLowerCase().replace(/\s+/g, " ").trim();
}

/* ─── Excel → ilaç JSON (SGK EK-4/A gerçek şeması) ───
   Kolonlar: Kamu No | Güncel Barkod | İlaç Adı | Eski Barkodlar |
   Eşdeğer İlaç Grubu | Terapötik Referans Grubu | Listeye Giriş |
   Aktiflenme | Pasiflenme | İndirim Esas Durumu | fiyat kademeleri...
   NOT: EK-4/A'da etkin madde kolonu yoktur; listede olmak = bedeli ödenir.
   Pasiflenme tarihi doluysa ilaç artık ödenmiyor demektir. */
function excelParse(buffer) {
  const wb = XLSX.read(buffer, { type: "buffer" });
  const sheet = wb.Sheets[wb.SheetNames[0]];
  const satirlar = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: "" });

  // Başlık satırını bul (barkod + ad geçen)
  let baslikIdx = satirlar.findIndex((r) =>
    r.some((c) => /barkod/i.test(String(c))) && r.some((c) => trNorm(c).includes("ad"))
  );
  if (baslikIdx === -1) baslikIdx = 0;
  const H = satirlar[baslikIdx].map(trNorm);
  const kolon = (...keys) => H.findIndex((b) => keys.some((k) => b.includes(k)));

  let iBarkod = H.findIndex((b) => b.includes("guncel barkod"));
  if (iBarkod === -1) iBarkod = H.findIndex((b) => b.includes("barkod") && !b.includes("eski"));
  const iAd = H.findIndex((b) => b.includes("ilac ad") || b.includes("urun ad") || (b.includes("ad") && !b.includes("barkod")));
  const iKamu = kolon("kamu no", "kamu");
  const iEsdeger = kolon("esdeger");
  const iTerapotik = kolon("terapotik");
  const iAktif = kolon("aktiflenme");
  const iPasif = kolon("pasiflenme");

  const ilaclar = [];
  for (let i = baslikIdx + 1; i < satirlar.length; i++) {
    const r = satirlar[i];
    const g = (idx) => (idx >= 0 ? String(r[idx] ?? "").trim() : "");
    const ad = g(iAd);
    const barkod = g(iBarkod);
    if (!ad && !barkod) continue;
    const pasif = g(iPasif);
    ilaclar.push({
      kamu_no: g(iKamu),
      barkod,
      ad,
      esdeger_grubu: g(iEsdeger),
      terapotik_grup: g(iTerapotik),
      aktiflenme_tarihi: g(iAktif),
      pasiflenme_tarihi: pasif,
      odeme_durumu: pasif ? "Pasif (ödenmez)" : "Ödenir (EK-4/A)",
    });
  }
  return ilaclar;
}

/* ─── EK-4/D izleme (PARSE ETMEZ — sadece değişiklik tespiti) ───
   EK-4/D PDF/DOC formatındadır; güvenlik için içeriği parse etmeyiz.
   Sadece dosyanın/sayfanın hash'ini alıp "değişti mi" diye izleriz.
   Değiştiğinde uygulama kullanıcıyı uyarır; kuralları kör uygulamaz. */
async function ek4dIzle(eskiEk4d, bugun) {
  const url = process.env.SGK_EK4D_URL;
  // İzleme URL'i yoksa eski durumu koru
  if (!url) {
    log("EK-4/D izleme URL'i (SGK_EK4D_URL) yok — atlanıyor.");
    return eskiEk4d || null;
  }
  try {
    const res = await fetch(url, { headers: UA });
    if (!res.ok) throw new Error(`EK-4/D alınamadı: ${res.status}`);
    const buf = Buffer.from(await res.arrayBuffer());
    const hash = createHash("sha256").update(buf).digest("hex").slice(0, 16);
    const degisti = !eskiEk4d || eskiEk4d.hash !== hash;
    if (degisti) log(`EK-4/D DEĞİŞTİ (yeni hash ${hash}). Uygulama kullanıcıyı uyaracak.`);
    else log(`EK-4/D değişmedi (hash ${hash}).`);
    return {
      hash,
      kaynak_url: url,
      kontrol_tarihi: bugun,
      degisim_tarihi: degisti ? bugun : (eskiEk4d?.degisim_tarihi || bugun),
    };
  } catch (e) {
    log("EK-4/D izleme hatası:", e.message, "- eski durum korunuyor.");
    return eskiEk4d || null;
  }
}

/* ─── Ana akış ─── */
async function main() {
  if (!existsSync(OUT_DIR)) mkdirSync(OUT_DIR, { recursive: true });
  const versionPath = join(OUT_DIR, "version.json");
  const ilaclarPath = join(OUT_DIR, "ilaclar.json");
  const bugun = new Date().toISOString().split("T")[0];

  const eskiVersion = existsSync(versionPath)
    ? JSON.parse(readFileSync(versionPath, "utf8"))
    : { tarih: "", hash: "", kayit_sayisi: 0, ek4d: null };

  // ── EK-4/A (ilaç listesi) ──
  let link, buffer, ilaclar = null, hash = eskiVersion.hash;
  try {
    link = await ek4aLinkiniBul();
    const res = await fetch(link, { headers: UA });
    if (!res.ok) throw new Error(`xlsx indirilemedi: ${res.status}`);
    buffer = Buffer.from(await res.arrayBuffer());
    ilaclar = excelParse(buffer);
    if (!ilaclar.length) throw new Error("Parse sonucu 0 kayıt — şüpheli");
    hash = createHash("sha256").update(JSON.stringify(ilaclar)).digest("hex").slice(0, 16);
  } catch (e) {
    log("EK-4/A hatası:", e.message, "- eski ilaç listesi korunuyor.");
    ilaclar = null; // yeniden yazma
  }

  // ── EK-4/D (sadece izleme) ──
  const ek4d = await ek4dIzle(eskiVersion.ek4d, bugun);

  const ilacDegisti = ilaclar !== null && hash !== eskiVersion.hash;
  const ek4dDegisti = JSON.stringify(ek4d) !== JSON.stringify(eskiVersion.ek4d || null);

  if (!ilacDegisti && !ek4dDegisti) {
    log("Değişiklik yok (ne EK-4/A ne EK-4/D). Güncelleme gerekmiyor.");
    process.exit(0);
  }

  // EK-4/A değiştiyse ilaç listesini yaz; değişmediyse eski tarih/hash korunur
  const version = {
    tarih: ilacDegisti ? bugun : eskiVersion.tarih,
    hash: ilacDegisti ? hash : eskiVersion.hash,
    kayit_sayisi: ilacDegisti ? ilaclar.length : eskiVersion.kayit_sayisi,
    kaynak_url: ilacDegisti ? link : eskiVersion.kaynak_url,
    onceki_tarih: ilacDegisti ? (eskiVersion.tarih || null) : eskiVersion.onceki_tarih,
    ek4d,
  };

  if (ilacDegisti) {
    writeFileSync(ilaclarPath, JSON.stringify(ilaclar), "utf8");
    log(`EK-4/A YAYIMLANDI: ${ilaclar.length} ilaç, hash ${hash}`);
  }
  writeFileSync(versionPath, JSON.stringify(version, null, 2), "utf8");
  log(`version.json güncellendi (EK-4/A ${ilacDegisti ? "değişti" : "aynı"}, EK-4/D ${ek4dDegisti ? "değişti" : "aynı"}).`);
}

main().catch((e) => { console.error("[sgk-veri] KRİTİK:", e); process.exit(1); });
