#!/usr/bin/env node
/**
 * Ed25519 İMZALA — sgk-data JSON'larını gizli anahtarla mühürler (.sig üretir).
 *
 * Eşi: eczasist uygulamasındaki gömülü açık anahtar (electron/imza.ts) doğrular.
 * Anahtar: IMZA_ANAHTAR_PEM env (GitHub Actions secret) veya .imza-gizli-anahtar.pem.
 *
 * Kullanım:  node scripts/imzala.mjs sgk-data/ilaclar.json   → sgk-data/ilaclar.json.sig
 *
 * NOT: İmza, GİT'TE SAKLANAN (raw.githubusercontent'in sunduğu) baytlara göre üretilmeli.
 * CI'da checkout LF bıraktığı için dosyayı okuyup imzalamak yeterli (autocrlf yok).
 */
import crypto from "node:crypto";
import fs from "node:fs";

/* GitHub/Netlify secret'ı çok satırlı PEM'i bozabilir (\n literali, boşluk, tek satır);
   bu normalizer hepsini geçerli PEM'e geri çevirir. */
function normalizePem(raw) {
  if (!raw) return null;
  let s = String(raw).trim().replace(/^["']|["']$/g, "").replace(/\\n/g, "\n");
  const m = s.match(/-----BEGIN ([A-Z ]+)-----([\s\S]*?)-----END \1-----/);
  if (m) {
    const body = m[2].replace(/\s+/g, "");
    const lines = (body.match(/.{1,64}/g) || [body]).join("\n");
    return `-----BEGIN ${m[1]}-----\n${lines}\n-----END ${m[1]}-----`;
  }
  return s;
}

function gizliAnahtar() {
  const pem = process.env.IMZA_ANAHTAR_PEM;
  if (pem) return crypto.createPrivateKey(normalizePem(pem));
  const yol = process.env.IMZA_ANAHTAR || ".imza-gizli-anahtar.pem";
  if (!fs.existsSync(yol)) {
    console.error(`Gizli anahtar yok: ${yol} (veya IMZA_ANAHTAR_PEM env).`);
    process.exit(2);
  }
  return crypto.createPrivateKey(fs.readFileSync(yol, "utf8"));
}

const dosya = process.argv[2];
if (!dosya) { console.error("Kullanım: node scripts/imzala.mjs <dosya.json>"); process.exit(1); }

const priv = gizliAnahtar();
const data = fs.readFileSync(dosya);
const sig = crypto.sign(null, data, priv).toString("base64");
fs.writeFileSync(dosya + ".sig", sig + "\n");
console.log(`İmzalandı → ${dosya}.sig`);
