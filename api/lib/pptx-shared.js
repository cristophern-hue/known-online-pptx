// ── Brand colors ─────────────────────────────────────────────────────────────
const ORANGE     = "FA5A1E";
const ORANGE2    = "FF912D";
const DARK       = "1A1A2E";
const WHITE      = "FFFFFF";
const LIGHT_BG   = "FFF8F5";
const GRAY_TEXT  = "64748B";
const LIGHT_GRAY = "F1F0EC";
const GREEN      = "3B6D11";
const GREEN_BG   = "EAF3DE";
const RED        = "A32D2D";
const RED_BG     = "FCEBEB";
const AMBER      = "854F0B";
const AMBER_BG   = "FAEEDA";
const BLUE       = "185FA5";
const LIGHT_BLUE = "E6F1FB";

const COLORS = {
  ORANGE, ORANGE2, DARK, WHITE, LIGHT_BG, GRAY_TEXT, LIGHT_GRAY,
  GREEN, GREEN_BG, RED, RED_BG, AMBER, AMBER_BG, BLUE, LIGHT_BLUE,
};

// ── Image fetch ───────────────────────────────────────────────────────────────
async function fetchImageAsBase64(url) {
  try {
    const res = await fetch(url, { signal: AbortSignal.timeout(8000) });
    if (!res.ok) return null;
    const ct  = res.headers.get("content-type") || "image/jpeg";
    const buf = Buffer.from(await res.arrayBuffer());
    const mime = ct.startsWith("image/") ? ct.split(";")[0] : "image/jpeg";
    return `image/${mime.split("/")[1]};base64,${buf.toString("base64")}`;
  } catch { return null; }
}

// ── USD conversion (Cummins) ──────────────────────────────────────────────────
const FALLBACK_RATE = 1050;

async function getUsdRate() {
  try {
    const r = await fetch("https://dolarapi.com/v1/dolares/oficial", { signal: AbortSignal.timeout(4000) });
    if (!r.ok) throw new Error();
    const j    = await r.json();
    const rate = parseFloat(j.venta);
    if (!rate || rate <= 0) throw new Error();
    return { rate, fallback: false };
  } catch {
    return { rate: FALLBACK_RATE, fallback: true };
  }
}

function arsToUsd(str, rate) {
  const c = (str || "0").replace(/\./g, "").replace(",", ".").replace(/[^0-9.]/g, "");
  return (parseFloat(c) || 0) / rate;
}

function fmtUSD(n) {
  return "$ " + n.toLocaleString("es-AR", { minimumFractionDigits: 2, maximumFractionDigits: 2 });
}

const ARS_FIELDS = [
  "META_COSTO", "META_COSTO_PREV",
  "GOOGLE_COSTO", "GOOGLE_COSTO_PREV",
  "INVERSION_TOTAL", "INVERSION_PREV",
  "META_CPC", "META_CPC_PREV",
  "GOOGLE_CPC", "GOOGLE_CPC_PREV",
  "CPC_TOTAL", "CPC_PREV",
  "ECOMMERCE_INGRESOS", "ECOMMERCE_INGRESOS_PREV",
  "ECOMMERCE_TICKET", "ECOMMERCE_TICKET_PREV",
];

function normalizeDataForUSD(DATA, rate) {
  const d = { ...DATA };
  for (const f of ARS_FIELDS) {
    if (d[f]) d[f] = fmtUSD(arsToUsd(d[f], rate));
  }
  if (Array.isArray(d.CAMPANAS)) {
    d.CAMPANAS = d.CAMPANAS.map(c => c.costo ? { ...c, costo: fmtUSD(arsToUsd(c.costo, rate)) } : c);
  }
  if (Array.isArray(d.TOP_ANUNCIOS_META)) {
    d.TOP_ANUNCIOS_META = d.TOP_ANUNCIOS_META.map(a => a.costo ? { ...a, costo: fmtUSD(arsToUsd(a.costo, rate)) } : a);
  }
  return d;
}

// ── Shared slide builders ─────────────────────────────────────────────────────

function buildSlide_Cover(pres, DATA) {
  const s1 = pres.addSlide();
  s1.background = { color: ORANGE };

  s1.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 0.18, h: 5.625, fill: { color: WHITE }, line: { color: WHITE } });
  s1.addShape(pres.shapes.OVAL, { x: 7.8, y: -1.2, w: 4.0, h: 4.0, fill: { color: WHITE, transparency: 88 }, line: { color: WHITE, transparency: 88 } });
  s1.addShape(pres.shapes.OVAL, { x: 8.5, y: -0.4, w: 2.4, h: 2.4, fill: { color: WHITE, transparency: 75 }, line: { color: WHITE, transparency: 75 } });

  s1.addShape(pres.shapes.OVAL, { x: 0.5,  y: 0.45, w: 0.52, h: 0.52, fill: { color: WHITE  }, line: { color: WHITE  } });
  s1.addShape(pres.shapes.OVAL, { x: 0.64, y: 0.59, w: 0.26, h: 0.26, fill: { color: ORANGE }, line: { color: ORANGE } });
  s1.addText("Known Online", { x: 1.15, y: 0.48, w: 3.5, h: 0.45, fontSize: 15, color: WHITE, bold: true, fontFace: "DM Sans", margin: 0 });

  s1.addShape(pres.shapes.RECTANGLE, { x: 0.45, y: 1.5, w: 1.8, h: 0.32, fill: { color: WHITE }, line: { color: WHITE } });
  s1.addText(DATA.CLIENTE_NOMBRE || "CLIENTE", { x: 0.45, y: 1.5, w: 1.8, h: 0.32, fontSize: 10, color: ORANGE, bold: true, fontFace: "DM Sans", align: "center", margin: 0 });

  s1.addText("Reporte\nPaid Media", { x: 0.45, y: 1.95, w: 7, h: 1.5, fontSize: 52, color: WHITE, bold: true, fontFace: "DM Sans", valign: "top" });
  s1.addText(`${DATA.PERIODO_ACTUAL_LABEL || ""} vs. ${DATA.PERIODO_ANTERIOR_LABEL || ""}`, { x: 0.45, y: 3.55, w: 7, h: 0.45, fontSize: 18, color: "FFD4B8", fontFace: "DM Sans" });

  if (DATA.TIPO_CAMBIO_USADO) {
    const tcLabel = DATA.TIPO_CAMBIO_FALLBACK
      ? `* Valores convertidos a USD · TC referencia: $${DATA.TIPO_CAMBIO_USADO.toLocaleString("es-AR")} ARS/USD (fallback — sin conexión a dolarapi.com al momento de generación)`
      : `* Valores convertidos a USD · TC dólar oficial: $${DATA.TIPO_CAMBIO_USADO.toLocaleString("es-AR")} ARS/USD (Fuente: dolarapi.com · Banco Nación Argentina)`;
    s1.addText(tcLabel, { x: 0.45, y: 5.28, w: 9.1, h: 0.22, fontSize: 7.5, color: "999999", fontFace: "DM Sans" });
  }
}

function buildSlide_Recommendations(pres, DATA) {
  const s = pres.addSlide();
  s.background = { color: ORANGE };
  s.addShape(pres.shapes.OVAL, { x:  7.5, y:  3.5, w: 3.5, h: 3.5, fill: { color: WHITE, transparency: 92 }, line: { color: WHITE, transparency: 92 } });
  s.addShape(pres.shapes.OVAL, { x: -1.0, y: -0.5, w: 2.5, h: 2.5, fill: { color: WHITE, transparency: 88 }, line: { color: WHITE, transparency: 88 } });
  const PLACEHOLDERS = [
    "Escribir recomendación aquí.",
    "Escribir recomendación aquí.",
    "Escribir recomendación aquí.",
    "Escribir recomendación aquí.",
    "Escribir recomendación aquí.",
  ];

  const parsed = Array.isArray(DATA.RECOMENDACIONES)
    ? DATA.RECOMENDACIONES
    : (() => { try { return JSON.parse(DATA.RECOMENDACIONES || '[]'); } catch(e) { return []; } })();
  const _recs = parsed.length > 0 ? parsed.map(r => r.titulo || r) : PLACEHOLDERS;

  _recs.slice(0, 5).forEach((texto, i) => {
    const num = String(i + 1).padStart(2, "0");
    const y   = 0.8 + i * 0.92;
    s.addShape(pres.shapes.OVAL, { x: 0.4, y, w: 0.45, h: 0.45, fill: { color: WHITE }, line: { color: WHITE } });
    s.addText(num,   { x: 0.4,  y,     w: 0.45, h: 0.45, fontSize: 11, bold: true, color: ORANGE, fontFace: "DM Sans", align: "center", valign: "middle" });
    s.addText(texto, { x: 1.05, y: y + 0.04, w: 8.5, h: 0.38, fontSize: 14, color: WHITE, fontFace: "DM Sans", valign: "middle" });
  });
}

function buildSlide_Close(pres, DATA) {
  const s = pres.addSlide();
  s.background = { color: ORANGE };
  s.addShape(pres.shapes.OVAL, { x:  6.5, y: -1.5, w: 5.5, h: 5.5, fill: { color: WHITE, transparency: 92 }, line: { color: WHITE, transparency: 92 } });
  s.addShape(pres.shapes.OVAL, { x: -2.0, y:  3.0, w: 4.5, h: 4.5, fill: { color: DARK,  transparency: 88 }, line: { color: DARK,  transparency: 88 } });
  s.addShape(pres.shapes.OVAL, { x: 0.5,  y:  0.5, w: 0.55, h: 0.55, fill: { color: WHITE  }, line: { color: WHITE  } });
  s.addShape(pres.shapes.OVAL, { x: 0.65, y: 0.65, w: 0.28, h: 0.28, fill: { color: ORANGE }, line: { color: ORANGE } });
  s.addText("Known Online", { x: 1.2, y: 0.52, w: 4, h: 0.45, fontSize: 16, bold: true, color: WHITE, fontFace: "DM Sans", margin: 0 });
  s.addText("¡Muchas gracias!", { x: 0.5, y: 1.6, w: 9, h: 1.4, fontSize: 56, bold: true, color: WHITE, fontFace: "DM Sans", align: "center" });
  s.addText("Logramos tu transformación digital", { x: 0.5, y: 3.1, w: 9, h: 0.45, fontSize: 18, color: "FFD4B8", fontFace: "DM Sans", align: "center", italic: true });
  s.addShape(pres.shapes.RECTANGLE, { x: 3.5, y: 3.7, w: 3.0, h: 0.04, fill: { color: WHITE, transparency: 60 }, line: { color: WHITE, transparency: 60 } });
  s.addText(DATA.WEB      || "www.knownonline.com",   { x: 0.5, y: 3.9, w: 9, h: 0.35, fontSize: 14, color: WHITE,    fontFace: "DM Sans", align: "center", bold: true });
  s.addText(DATA.CONTACTO || "ariel@knownonline.com", { x: 0.5, y: 4.3, w: 9, h: 0.3,  fontSize: 12, color: "FFD4B8", fontFace: "DM Sans", align: "center" });
}

module.exports = {
  COLORS,
  fetchImageAsBase64,
  FALLBACK_RATE,
  getUsdRate,
  arsToUsd,
  fmtUSD,
  ARS_FIELDS,
  normalizeDataForUSD,
  buildSlide_Cover,
  buildSlide_Recommendations,
  buildSlide_Close,
};
