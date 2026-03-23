const pptxgen = require("pptxgenjs");

async function fetchImageAsBase64(url) {
  try {
    const res = await fetch(url, { signal: AbortSignal.timeout(8000) });
    if (!res.ok) return null;
    const ct = res.headers.get("content-type") || "image/jpeg";
    const buf = Buffer.from(await res.arrayBuffer());
    const mime = ct.startsWith("image/") ? ct.split(";")[0] : "image/jpeg";
    return `image/${mime.split("/")[1]};base64,${buf.toString("base64")}`;
  } catch { return null; }
}

// ── USD conversion (Cummins) ─────────────────────────────────────────────────
const FALLBACK_RATE = 1050;

async function getUsdRate() {
  try {
    const r = await fetch("https://dolarapi.com/v1/dolares/oficial", { signal: AbortSignal.timeout(4000) });
    if (!r.ok) throw new Error();
    const j = await r.json();
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

module.exports = async function handler(req, res) {
  if (req.method !== "POST") {
    return res.status(405).json({ error: "Method not allowed. Use POST." });
  }

  const { DATA } = req.body || {};
  if (!DATA) {
    return res.status(400).json({ error: "Missing DATA in request body." });
  }

  try {
    const isCummins = (DATA.CLIENTE_NOMBRE || "").toLowerCase().includes("cummins");
    let pptxData = DATA;
    if (isCummins) {
      const override = parseFloat(DATA.TIPO_CAMBIO_OVERRIDE);
      const { rate, fallback } = override > 0
        ? { rate: override, fallback: false }
        : await getUsdRate();
      pptxData = normalizeDataForUSD(DATA, rate);
      pptxData.TIPO_CAMBIO_USADO = rate;
      pptxData.TIPO_CAMBIO_FALLBACK = fallback;
    }
    const base64 = await generatePptx(pptxData);
    const filename = `Reporte_${pptxData.CLIENTE_NOMBRE || "Cliente"}_${pptxData.PERIODO_ACTUAL_LABEL || "Periodo"}.pptx`
      .replace(/\s+/g, "_");
    return res.status(200).json({ pptx: base64, filename });
  } catch (err) {
    console.error(err);
    return res.status(500).json({ error: err.message || "Error generating PPTX." });
  }
};

async function generatePptx(DATA) {
  const pres = new pptxgen();
  pres.layout = "LAYOUT_16x9";
  pres.title = `Reporte Paid Media ${DATA.CLIENTE_NOMBRE || ""} - ${DATA.PERIODO_ACTUAL_LABEL || ""}`;

  // ── Period short labels ───────────────────────────────────────────────────
  const MESES = ["Ene","Feb","Mar","Abr","May","Jun","Jul","Ago","Sep","Oct","Nov","Dic"];
  const mesIdx           = (parseInt(DATA.mes_actual) || 1) - 1;
  const añoActual        = parseInt(DATA.año_actual) || new Date().getFullYear();
  const labelCortoActual   = `${MESES[mesIdx]} '${String(añoActual).slice(2)}`;
  const labelCortoAnterior = `${MESES[mesIdx]} '${String(añoActual - 1).slice(2)}`;

  // ── Brand colors ──────────────────────────────────────────────────────────
  const ORANGE    = "FA5A1E";
  const ORANGE2   = "FF912D";
  const DARK      = "1A1A2E";
  const WHITE     = "FFFFFF";
  const LIGHT_BG  = "FFF8F5";
  const GRAY_TEXT = "64748B";
  const LIGHT_GRAY = "F1F0EC";
  const GREEN     = "3B6D11";
  const GREEN_BG  = "EAF3DE";
  const RED       = "A32D2D";
  const RED_BG    = "FCEBEB";
  const AMBER     = "854F0B";
  const AMBER_BG  = "FAEEDA";
  const BLUE      = "185FA5";
  const LIGHT_BLUE = "E6F1FB";

  // ── Helpers ───────────────────────────────────────────────────────────────
  const parseNum = str => {
    const c = (str || "0").replace(/\./g, "").replace(",", ".").replace(/[^0-9.]/g, "");
    return parseFloat(c) || 0;
  };
  const fmtMoneyCompact = val => {
    const n = parseNum(val);
    if (n >= 1_000_000) return `$${(n / 1_000_000).toFixed(2).replace(".", ",")} M`;
    if (n >= 1_000)     return `$${(n / 1_000).toFixed(1).replace(".", ",")} K`;
    return val || "";
  };

  // ── SLIDE 1 – COVER ───────────────────────────────────────────────────────
  let s1 = pres.addSlide();
  s1.background = { color: DARK };

  s1.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 0.18, h: 5.625, fill: { color: ORANGE }, line: { color: ORANGE } });
  s1.addShape(pres.shapes.OVAL, { x: 7.8, y: -1.2, w: 4.0, h: 4.0, fill: { color: ORANGE, transparency: 88 }, line: { color: ORANGE, transparency: 88 } });
  s1.addShape(pres.shapes.OVAL, { x: 8.5, y: -0.4, w: 2.4, h: 2.4, fill: { color: ORANGE, transparency: 75 }, line: { color: ORANGE, transparency: 75 } });

  s1.addShape(pres.shapes.OVAL, { x: 0.5, y: 0.45, w: 0.52, h: 0.52, fill: { color: ORANGE }, line: { color: ORANGE } });
  s1.addShape(pres.shapes.OVAL, { x: 0.64, y: 0.59, w: 0.26, h: 0.26, fill: { color: WHITE }, line: { color: WHITE } });
  s1.addText("Known Online", { x: 1.15, y: 0.48, w: 3.5, h: 0.45, fontSize: 15, color: ORANGE, bold: true, fontFace: "DM Sans", margin: 0 });

  s1.addShape(pres.shapes.RECTANGLE, { x: 0.45, y: 1.5, w: 1.8, h: 0.32, fill: { color: ORANGE }, line: { color: ORANGE } });
  s1.addText(DATA.CLIENTE_NOMBRE || "CLIENTE", { x: 0.45, y: 1.5, w: 1.8, h: 0.32, fontSize: 10, color: WHITE, bold: true, fontFace: "DM Sans", align: "center", margin: 0 });

  s1.addText("Reporte\nPaid Media", { x: 0.45, y: 1.95, w: 7, h: 1.5, fontSize: 52, color: WHITE, bold: true, fontFace: "Trebuchet MS", valign: "top" });
  s1.addText(`${DATA.PERIODO_ACTUAL_LABEL || ""} vs. ${DATA.PERIODO_ANTERIOR_LABEL || ""}`, { x: 0.45, y: 3.55, w: 7, h: 0.45, fontSize: 18, color: ORANGE2, fontFace: "DM Sans" });

  if (DATA.TIPO_CAMBIO_USADO) {
    const tcLabel = DATA.TIPO_CAMBIO_FALLBACK
      ? `Inversión convertida a USD · TC ref: $${DATA.TIPO_CAMBIO_USADO.toLocaleString("es-AR")} ARS (sin conexión a API — verificar)`
      : `Inversión convertida a USD · TC oficial: $${DATA.TIPO_CAMBIO_USADO.toLocaleString("es-AR")} ARS`;
    s1.addText(tcLabel, { x: 0.45, y: 5.28, w: 9.1, h: 0.22, fontSize: 7.5, color: "999999", fontFace: "DM Sans" });
  }

  // ── SLIDE ATIKA – TABLA KPIs GENERAL (CONDICIONAL) ───────────────────────
  if (DATA.ATIKA_PINTEREST_INV) {
    // Parsea string ARS/numérico ("$ 4.102.650", "7,42", "49 s") a número
    const parseNum = str => {
      const c = (str || "0").replace(/\./g, "").replace(",", ".").replace(/[^0-9.]/g, "");
      return parseFloat(c) || 0;
    };
    const fmtARS    = n => "$ " + Math.round(n).toLocaleString("es-AR");
    const fmtROAS   = n => n.toFixed(2).replace(".", ",") + "x";
    const fmtDelta  = n => (n >= 0 ? "+" : "") + Math.round(n) + "%";
    const calcDelta = (a, p) => p !== 0 ? fmtDelta((a - p) / p * 100) : "%";
    const calcUp    = (a, p) => p !== 0 ? a >= p : true;

    // Totales inversión
    const invMetaA = parseNum(DATA.META_COSTO),        invMetaP = parseNum(DATA.META_COSTO_PREV);
    const invGoogA = parseNum(DATA.GOOGLE_COSTO),      invGoogP = parseNum(DATA.GOOGLE_COSTO_PREV);
    const invPinA  = parseNum(DATA.ATIKA_PINTEREST_INV), invPinP = parseNum(DATA.ATIKA_PINTEREST_INV_PREV);
    const invTotA  = invMetaA + invGoogA + invPinA,    invTotP  = invMetaP + invGoogP + invPinP;

    // Ventas sin canceladas
    const vSinA = parseNum(DATA.VTEX_INGRESOS_ACTUAL), vSinP = parseNum(DATA.VTEX_INGRESOS_ANTERIOR);

    // ROAS calculados
    const vCpcA = parseNum(DATA.ATIKA_VENTAS_CPC),     vCpcP = parseNum(DATA.ATIKA_VENTAS_CPC_PREV);
    const vPinA = parseNum(DATA.ATIKA_VENTAS_PINTEREST), vPinP = parseNum(DATA.ATIKA_VENTAS_PINTEREST_PREV);
    const roasGenA = invGoogA ? vSinA / invGoogA : 0,  roasGenP = invGoogP ? vSinP / invGoogP : 0;
    const roasCpcA = invGoogA ? vCpcA / invGoogA : 0,  roasCpcP = invGoogP ? vCpcP / invGoogP : 0;
    const roasPinA = invPinA  ? vPinA / invPinA  : 0,  roasPinP = invPinP  ? vPinP / invPinP  : 0;

    let sAtika = pres.addSlide();
    sAtika.background = { color: WHITE };
    sAtika.addText("Performance General", { x: 0.4, y: 0.17, w: 9.2, h: 0.42, fontSize: 26, bold: true, color: DARK, fontFace: "Trebuchet MS" });
    sAtika.addText(`${DATA.PERIODO_ACTUAL_LABEL || ""} vs ${DATA.PERIODO_ANTERIOR_LABEL || ""}  ·  Inversión · Tráfico · Tiempo · Ventas · ROAS`, { x: 0.4, y: 0.59, w: 9.2, h: 0.24, fontSize: 10, color: GRAY_TEXT, fontFace: "DM Sans" });

    const atX = [0.4, 4.1, 5.95, 7.8], atW = [3.7, 1.85, 1.85, 1.8];
    const hY = 0.88, hH = 0.30, rH = 0.227;

    sAtika.addShape(pres.shapes.RECTANGLE, { x: 0.4, y: hY, w: 9.2, h: hH, fill: { color: DARK }, line: { color: DARK } });
    [["KPI","left"], [DATA.PERIODO_ACTUAL_LABEL||"Actual","center"], [DATA.PERIODO_ANTERIOR_LABEL||"Anterior","center"], ["Var %","center"]].forEach(([h,align],i) => {
      sAtika.addText(h, { x: atX[i]+0.08, y: hY+0.04, w: atW[i], h: hH-0.08, fontSize: 9.5, bold: true, color: WHITE, fontFace: "DM Sans", valign: "middle", align });
    });

    const atikaRows = [
      { label: "Inversión Meta",                    a: DATA.META_COSTO||"",                   p: DATA.META_COSTO_PREV||"",                d: DATA.META_COSTO_DELTA||"",                up: DATA.META_COSTO_DELTA_UP===true },
      { label: "Inversión Google",                  a: DATA.GOOGLE_COSTO||"",                  p: DATA.GOOGLE_COSTO_PREV||"",              d: DATA.GOOGLE_COSTO_DELTA||"",              up: DATA.GOOGLE_COSTO_DELTA_UP===true },
      { label: "Inversión Pinterest",               a: DATA.ATIKA_PINTEREST_INV||"",           p: DATA.ATIKA_PINTEREST_INV_PREV||"",       d: DATA.ATIKA_PINTEREST_INV_DELTA||"",       up: DATA.ATIKA_PINTEREST_INV_UP===true },
      { label: "Total Inversión", bold: true,       a: fmtARS(invTotA),                        p: fmtARS(invTotP),                         d: calcDelta(invTotA,invTotP),               up: calcUp(invTotA,invTotP) },
      { label: "Tráfico web",                       a: DATA.ATIKA_TRAFICO_TOTAL||"",           p: DATA.ATIKA_TRAFICO_TOTAL_PREV||"",       d: DATA.ATIKA_TRAFICO_TOTAL_DELTA||"",       up: DATA.ATIKA_TRAFICO_TOTAL_UP===true },
      { label: "Tráfico CPC",                       a: DATA.ATIKA_TRAFICO_CPC||"",             p: DATA.ATIKA_TRAFICO_CPC_PREV||"",         d: DATA.ATIKA_TRAFICO_CPC_DELTA||"",         up: DATA.ATIKA_TRAFICO_CPC_UP===true },
      { label: "Tráfico email mkt",                 a: DATA.ATIKA_TRAFICO_EMAIL||"",           p: DATA.ATIKA_TRAFICO_EMAIL_PREV||"",       d: DATA.ATIKA_TRAFICO_EMAIL_DELTA||"",       up: DATA.ATIKA_TRAFICO_EMAIL_UP===true },
      { label: "Tiempo de permanencia Web",         a: DATA.ATIKA_TIEMPO_WEB||"",              p: DATA.ATIKA_TIEMPO_WEB_PREV||"",          d: DATA.ATIKA_TIEMPO_WEB_DELTA||"",          up: DATA.ATIKA_TIEMPO_WEB_UP===true },
      { label: "Tiempo de permanencia CPC",         a: DATA.ATIKA_TIEMPO_CPC||"",              p: DATA.ATIKA_TIEMPO_CPC_PREV||"",          d: DATA.ATIKA_TIEMPO_CPC_DELTA||"",          up: DATA.ATIKA_TIEMPO_CPC_UP===true },
      { label: "Tiempo de permanencia email mkt",   a: DATA.ATIKA_TIEMPO_EMAIL||"",            p: DATA.ATIKA_TIEMPO_EMAIL_PREV||"",        d: DATA.ATIKA_TIEMPO_EMAIL_DELTA||"",        up: DATA.ATIKA_TIEMPO_EMAIL_UP===true },
      { label: "Tiempo de permanencia Orgánico",    a: DATA.ATIKA_TIEMPO_ORGANICO||"",         p: DATA.ATIKA_TIEMPO_ORGANICO_PREV||"",     d: DATA.ATIKA_TIEMPO_ORGANICO_DELTA||"",     up: DATA.ATIKA_TIEMPO_ORGANICO_UP===true },
      { label: "Ventas sitio (con canceladas)",     a: DATA.GA4_INGRESOS||"",                  p: DATA.GA4_INGRESOS_PREV||"",              d: DATA.GA4_INGRESOS_DELTA||"",              up: DATA.GA4_INGRESOS_DELTA_UP===true },
      { label: "Ventas sitio (sin canceladas)",     a: DATA.VTEX_INGRESOS_ACTUAL||"",          p: DATA.VTEX_INGRESOS_ANTERIOR||"",         d: calcDelta(vSinA,vSinP),                   up: calcUp(vSinA,vSinP) },
      { label: "Ventas CPC",                        a: DATA.ATIKA_VENTAS_CPC||"",              p: DATA.ATIKA_VENTAS_CPC_PREV||"",          d: DATA.ATIKA_VENTAS_CPC_DELTA||"",          up: DATA.ATIKA_VENTAS_CPC_UP===true },
      { label: "Ventas email mkt",                  a: DATA.ATIKA_VENTAS_EMAIL||"",            p: DATA.ATIKA_VENTAS_EMAIL_PREV||"",        d: DATA.ATIKA_VENTAS_EMAIL_DELTA||"",        up: DATA.ATIKA_VENTAS_EMAIL_UP===true },
      { label: "Ventas Pinterest",                  a: DATA.ATIKA_VENTAS_PINTEREST||"",        p: DATA.ATIKA_VENTAS_PINTEREST_PREV||"",    d: DATA.ATIKA_VENTAS_PINTEREST_DELTA||"",    up: DATA.ATIKA_VENTAS_PINTEREST_UP===true },
      { label: "ROAS General (VTEX) · inv. Google", a: fmtROAS(roasGenA),                      p: fmtROAS(roasGenP),                       d: calcDelta(roasGenA,roasGenP),             up: calcUp(roasGenA,roasGenP) },
      { label: "ROAS CPC",                          a: fmtROAS(roasCpcA),                      p: fmtROAS(roasCpcP),                       d: calcDelta(roasCpcA,roasCpcP),             up: calcUp(roasCpcA,roasCpcP) },
      { label: "ROAS Pinterest",                    a: fmtROAS(roasPinA),                      p: fmtROAS(roasPinP),                       d: calcDelta(roasPinA,roasPinP),             up: calcUp(roasPinA,roasPinP) },
    ];

    const sepAfter = new Set([3, 6, 10, 15]);
    atikaRows.forEach((row, i) => {
      const ry = hY + hH + i * rH;
      const bg = row.bold ? "EDE8E0" : (i % 2 === 0 ? WHITE : LIGHT_BG);
      sAtika.addShape(pres.shapes.RECTANGLE, { x: 0.4, y: ry, w: 9.2, h: rH, fill: { color: bg }, line: { color: "E8E0D8", width: 0.3 } });
      sAtika.addText(row.label, { x: atX[0]+0.08, y: ry+0.02, w: atW[0]-0.1, h: rH-0.04, fontSize: 8.5, bold: !!row.bold, color: DARK,      fontFace: "DM Sans", valign: "middle" });
      sAtika.addText(row.a,     { x: atX[1],       y: ry+0.02, w: atW[1],     h: rH-0.04, fontSize: 8.5, bold: !!row.bold, color: DARK,      fontFace: "DM Sans", align: "center", valign: "middle" });
      sAtika.addText(row.p,     { x: atX[2],       y: ry+0.02, w: atW[2],     h: rH-0.04, fontSize: 8.5,                   color: GRAY_TEXT, fontFace: "DM Sans", align: "center", valign: "middle" });
      if (row.d) {
        const bw = 1.1, bh = rH - 0.06, bx = atX[3] + (atW[3] - bw) / 2;
        sAtika.addShape(pres.shapes.RECTANGLE, { x: bx, y: ry+0.03, w: bw, h: bh, fill: { color: row.up ? GREEN_BG : RED_BG }, line: { color: row.up ? GREEN_BG : RED_BG } });
        sAtika.addText(row.d,   { x: bx,           y: ry+0.03, w: bw,         h: bh,      fontSize: 8.5, bold: true, color: row.up ? GREEN : RED, fontFace: "DM Sans", align: "center", valign: "middle" });
      }
      if (sepAfter.has(i)) {
        sAtika.addShape(pres.shapes.RECTANGLE, { x: 0.4, y: ry + rH - 0.012, w: 9.2, h: 0.012, fill: { color: "C8BEB5" }, line: { color: "C8BEB5" } });
      }
    });
  }

  // ── SLIDE 3 – GA4 ─────────────────────────────────────────────────────────
  let s7 = pres.addSlide();
  s7.background = { color: WHITE };
  s7.addText("Informe del Sitio", { x: 0.5, y: 0.2, w: 7, h: 0.55, fontSize: 28, bold: true, color: DARK, fontFace: "Trebuchet MS" });
  s7.addText(`GA4  ·  ${DATA.PERIODO_ACTUAL_LABEL || ""} vs ${DATA.PERIODO_ANTERIOR_LABEL || ""}`, { x: 0.5, y: 0.76, w: 7, h: 0.3, fontSize: 13, color: GRAY_TEXT, fontFace: "DM Sans" });

  const ga4Metrics = [
    { icon: "S", label: "Sesiones",              sub: "Sesiones frente al año anterior",          val26: DATA.GA4_SESIONES     || "", val25: DATA.GA4_SESIONES_PREV     || "", delta: DATA.GA4_SESIONES_DELTA     || "", deltaColor: DATA.GA4_SESIONES_DELTA_UP     === true ? GREEN : RED, deltaBg: DATA.GA4_SESIONES_DELTA_UP     === true ? GREEN_BG : RED_BG },
    { icon: "R", label: "Ingresos",               sub: "Revenue GA4 (Purchase)",                   val26: fmtMoneyCompact(DATA.GA4_INGRESOS), val25: DATA.GA4_INGRESOS_PREV     || "", delta: DATA.GA4_INGRESOS_DELTA     || "", deltaColor: DATA.GA4_INGRESOS_DELTA_UP     === true ? GREEN : RED, deltaBg: DATA.GA4_INGRESOS_DELTA_UP     === true ? GREEN_BG : RED_BG },
    { icon: "T", label: "Transacciones",         sub: "Transacciones ecommerce (VTEX/GA4)",        val26: DATA.GA4_TRANSACCIONES || "", val25: DATA.GA4_TRANSACCIONES_PREV || "", delta: DATA.GA4_TRANSACCIONES_DELTA || "", deltaColor: DATA.GA4_TRANSACCIONES_DELTA_UP === true ? GREEN : RED, deltaBg: DATA.GA4_TRANSACCIONES_DELTA_UP === true ? GREEN_BG : RED_BG },
    { icon: "$", label: "Inversión publicitaria", sub: "Total Meta Ads + Google Ads",              val26: fmtMoneyCompact(DATA.INVERSION_TOTAL), val25: DATA.INVERSION_PREV        || "", delta: DATA.INVERSION_DELTA        || "", deltaColor: DATA.INVERSION_DELTA_UP        === true ? GREEN : RED, deltaBg: DATA.INVERSION_DELTA_UP        === true ? GREEN_BG : RED_BG },
    { icon: "%", label: "Tasa de conversión",    sub: "eventCount(purchase) / sesiones",           val26: DATA.GA4_CONV_RATE    || "", val25: DATA.GA4_CONV_RATE_PREV    || "", delta: DATA.GA4_CONV_RATE_DELTA    || "", deltaColor: DATA.GA4_CONV_RATE_DELTA_UP    === true ? GREEN : RED, deltaBg: DATA.GA4_CONV_RATE_DELTA_UP    === true ? GREEN_BG : RED_BG },
    { icon: "T", label: "Ticket promedio",        sub: "Ingreso promedio por compra GA4",            val26: DATA.GA4_TICKET        || "", val25: DATA.GA4_TICKET_PREV        || "", delta: DATA.GA4_TICKET_DELTA        || "", deltaColor: DATA.GA4_TICKET_DELTA_UP        === true ? GREEN : RED, deltaBg: DATA.GA4_TICKET_DELTA_UP        === true ? GREEN_BG : RED_BG },
  ];
  ga4Metrics.forEach((m, i) => {
    const col = i % 3, row = Math.floor(i / 3);
    const x = 0.4 + col * 3.13, y = 1.2 + row * 1.85;
    s7.addShape(pres.shapes.RECTANGLE, { x, y, w: 2.9, h: 1.7, fill: { color: LIGHT_BG }, line: { color: "F0E8E0", width: 0.5 } });
    s7.addShape(pres.shapes.RECTANGLE, { x, y, w: 2.9, h: 0.06, fill: { color: ORANGE }, line: { color: ORANGE } });
    s7.addShape(pres.shapes.OVAL, { x: x + 0.14, y: y + 0.18, w: 0.36, h: 0.36, fill: { color: ORANGE }, line: { color: ORANGE } });
    s7.addText(m.label, { x: x + 0.58, y: y + 0.18, w: 2.2, h: 0.22, fontSize: 11, bold: true, color: DARK, fontFace: "DM Sans" });
    s7.addShape(pres.shapes.RECTANGLE, { x: x + 0.14, y: y + 0.65, w: 2.62, h: 0.02, fill: { color: "E8E0D8" }, line: { color: "E8E0D8" } });
    s7.addText(labelCortoActual, { x: x + 0.14, y: y + 0.75, w: 1.3,  h: 0.18, fontSize: 9,  color: GRAY_TEXT, fontFace: "DM Sans" });
    s7.addText(m.val26,  { x: x + 0.14, y: y + 0.92, w: 1.5,  h: 0.38, fontSize: 22, bold: true, color: DARK, fontFace: "Trebuchet MS" });
    s7.addShape(pres.shapes.RECTANGLE, { x: x + 1.75, y: y + 0.95, w: 0.95, h: 0.28, fill: { color: m.deltaBg }, line: { color: m.deltaBg } });
    s7.addText(m.delta,  { x: x + 1.75, y: y + 0.95, w: 0.95, h: 0.28, fontSize: 11, bold: true, color: m.deltaColor, fontFace: "DM Sans", align: "center" });
    s7.addText(`${labelCortoAnterior}: ${m.val25}`, { x: x + 0.14, y: y + 1.35, w: 2.5, h: 0.2, fontSize: 9, color: GRAY_TEXT, fontFace: "DM Sans" });
  });

  s7.addShape(pres.shapes.RECTANGLE, { x: 0.4, y: 4.95, w: 9.2, h: 0.55, fill: { color: "FFF0EB" }, line: { color: "FA5A1E", width: 0.5 } });
  s7.addShape(pres.shapes.RECTANGLE, { x: 0.4, y: 4.95, w: 0.08, h: 0.55, fill: { color: ORANGE }, line: { color: ORANGE } });
  s7.addText(DATA.GA4_INSIGHT || "", { x: 0.6, y: 4.97, w: 8.9, h: 0.5, fontSize: 10, color: DARK, fontFace: "DM Sans", valign: "middle" });

  // ── SLIDE 3B – ECOMMERCE PLATFORM (OPCIONAL) ─────────────────────────────
  if (DATA.ECOMMERCE_INGRESOS) {
    let sEc = pres.addSlide();
    sEc.background = { color: WHITE };
    sEc.addText("Performance Ecommerce", { x: 0.5, y: 0.2, w: 7, h: 0.55, fontSize: 28, bold: true, color: DARK, fontFace: "Trebuchet MS" });
    sEc.addText(`Plataforma  ·  ${DATA.PERIODO_ACTUAL_LABEL || ""} vs ${DATA.PERIODO_ANTERIOR_LABEL || ""}`, { x: 0.5, y: 0.76, w: 7, h: 0.3, fontSize: 13, color: GRAY_TEXT, fontFace: "DM Sans" });

    const ecMetrics = [
      { icon: "R", label: "Ingresos",       sub: "Revenue de plataforma ecommerce", val: fmtMoneyCompact(DATA.ECOMMERCE_INGRESOS), prev: DATA.ECOMMERCE_INGRESOS_PREV || "", delta: DATA.ECOMMERCE_INGRESOS_DELTA || "", up: DATA.ECOMMERCE_INGRESOS_DELTA_UP === true },
      { icon: "O", label: "Órdenes",        sub: "Transacciones / pedidos",         val: DATA.ECOMMERCE_ORDENES  || "", prev: DATA.ECOMMERCE_ORDENES_PREV  || "", delta: DATA.ECOMMERCE_ORDENES_DELTA  || "", up: DATA.ECOMMERCE_ORDENES_DELTA_UP  === true },
      { icon: "T", label: "Ticket promedio", sub: "Ingreso promedio por orden",      val: DATA.ECOMMERCE_TICKET   || "", prev: DATA.ECOMMERCE_TICKET_PREV   || "", delta: DATA.ECOMMERCE_TICKET_DELTA   || "", up: DATA.ECOMMERCE_TICKET_DELTA_UP   === true },
    ];

    ecMetrics.forEach((m, i) => {
      const x = 0.85 + i * 2.8, y = 1.3;
      sEc.addShape(pres.shapes.RECTANGLE, { x, y, w: 2.6, h: 3.2, fill: { color: LIGHT_BG }, line: { color: "F0E8E0", width: 0.5 } });
      sEc.addShape(pres.shapes.RECTANGLE, { x, y, w: 2.6, h: 0.06, fill: { color: ORANGE }, line: { color: ORANGE } });
      sEc.addShape(pres.shapes.OVAL, { x: x + 0.14, y: y + 0.2, w: 0.45, h: 0.45, fill: { color: ORANGE }, line: { color: ORANGE } });
      sEc.addText(m.label, { x: x + 0.7,  y: y + 0.22, w: 1.8,  h: 0.28, fontSize: 12, bold: true, color: DARK,      fontFace: "DM Sans" });
      sEc.addShape(pres.shapes.RECTANGLE, { x: x + 0.14, y: y + 0.82, w: 2.32, h: 0.02, fill: { color: "E8E0D8" }, line: { color: "E8E0D8" } });
      sEc.addText(DATA.PERIODO_ACTUAL_LABEL || "", { x: x + 0.14, y: y + 0.92, w: 2.0, h: 0.2, fontSize: 9, color: GRAY_TEXT, fontFace: "DM Sans" });
      sEc.addText(m.val,   { x: x + 0.14, y: y + 1.12, w: 2.32, h: 0.65, fontSize: 28, bold: true, color: DARK, fontFace: "Trebuchet MS" });
      sEc.addShape(pres.shapes.RECTANGLE, { x: x + 0.14, y: y + 1.82, w: 2.32, h: 0.02, fill: { color: "E8E0D8" }, line: { color: "E8E0D8" } });
      sEc.addText(`${DATA.PERIODO_ANTERIOR_LABEL || "Período ant."}: ${m.prev}`, { x: x + 0.14, y: y + 1.92, w: 2.32, h: 0.2, fontSize: 9, color: GRAY_TEXT, fontFace: "DM Sans" });
      sEc.addShape(pres.shapes.RECTANGLE, { x: x + 0.5, y: y + 2.25, w: 1.6, h: 0.38, fill: { color: m.up ? GREEN_BG : RED_BG }, line: { color: m.up ? GREEN_BG : RED_BG } });
      sEc.addText(m.delta, { x: x + 0.5, y: y + 2.25, w: 1.6, h: 0.38, fontSize: 16, bold: true, color: m.up ? GREEN : RED, fontFace: "DM Sans", align: "center", valign: "middle" });
    });
  }

  // ── SLIDE 5B – TOP FUENTE / MEDIO (GA4) ──────────────────────────────────
  // fuenteMedio: array of { nombre, sesiones, txns, tc, tc_prev, tc_delta, tc_delta_up, revenue, revenue_prev, revenue_delta, revenue_delta_up }
  const fuenteMedio = DATA.FUENTE_MEDIO || [];
  if (fuenteMedio.length > 0) {
    let sFm = pres.addSlide();
    sFm.background = { color: WHITE };
    sFm.addText("Top 10 Fuente / Medio", { x: 0.5, y: 0.18, w: 7, h: 0.52, fontSize: 28, bold: true, color: DARK, fontFace: "Trebuchet MS" });
    sFm.addText(`GA4  ·  ${DATA.PERIODO_ACTUAL_LABEL || ""} vs ${DATA.PERIODO_ANTERIOR_LABEL || ""}`, { x: 0.5, y: 0.71, w: 7, h: 0.28, fontSize: 13, color: GRAY_TEXT, fontFace: "DM Sans" });

    // Table header
    const fmColW = [2.7, 1.05, 0.9, 1.0, 1.0, 0.85, 1.05, 0.55];
    const fmHeaders = ["Fuente / Medio", "Sesiones", "Txns", `TC% ${DATA.PERIODO_ACTUAL_LABEL || "Actual"}`, `TC% ${DATA.PERIODO_ANTERIOR_LABEL || "Ant."}`, "ΔTC", "Revenue", "ΔRev"];
    const fmY0 = 1.08;

    sFm.addShape(pres.shapes.RECTANGLE, { x: 0.4, y: fmY0, w: 9.2, h: 0.36, fill: { color: DARK }, line: { color: DARK } });
    let fmCx = 0.55;
    fmHeaders.forEach((h, i) => {
      const align = i === 0 ? "left" : "center";
      sFm.addText(h, { x: fmCx, y: fmY0 + 0.02, w: fmColW[i], h: 0.32, fontSize: 9, bold: true, color: WHITE, fontFace: "DM Sans", valign: "middle", align });
      fmCx += fmColW[i];
    });

    // Table rows
    fuenteMedio.slice(0, 10).forEach((row, i) => {
      const ry = fmY0 + 0.36 + i * 0.37;
      const bg = i % 2 === 0 ? WHITE : LIGHT_BG;
      sFm.addShape(pres.shapes.RECTANGLE, { x: 0.4, y: ry, w: 9.2, h: 0.36, fill: { color: bg }, line: { color: "E8E0D8", width: 0.5 } });

      const up = row.tc_delta_up === true;
      const tcColor  = up ? DARK : RED;
      const deltaBg  = up ? GREEN_BG  : RED_BG;
      const deltaTxt = up ? GREEN     : RED;

      let rx = 0.55;
      // Fuente / Medio
      sFm.addText(row.nombre || "", { x: rx, y: ry + 0.07, w: fmColW[0], h: 0.26, fontSize: 9.5, color: DARK, fontFace: "DM Sans" });
      rx += fmColW[0];
      // Sesiones
      sFm.addText(String(row.sesiones ?? ""), { x: rx, y: ry + 0.07, w: fmColW[1], h: 0.26, fontSize: 9.5, color: DARK, fontFace: "DM Sans", align: "center" });
      rx += fmColW[1];
      // Txns
      sFm.addText(String(row.txns ?? ""), { x: rx, y: ry + 0.07, w: fmColW[2], h: 0.26, fontSize: 9.5, color: DARK, fontFace: "DM Sans", align: "center" });
      rx += fmColW[2];
      // TC% actual (rojo si bajó)
      sFm.addText(row.tc || "", { x: rx, y: ry + 0.07, w: fmColW[3], h: 0.26, fontSize: 9.5, bold: !up, color: tcColor, fontFace: "DM Sans", align: "center" });
      rx += fmColW[3];
      // TC% anterior
      sFm.addText(row.tc_prev || "", { x: rx, y: ry + 0.07, w: fmColW[4], h: 0.26, fontSize: 9.5, color: GRAY_TEXT, fontFace: "DM Sans", align: "center" });
      rx += fmColW[4];
      // ΔTC badge
      sFm.addShape(pres.shapes.RECTANGLE, { x: rx + 0.05, y: ry + 0.09, w: 0.72, h: 0.22, fill: { color: deltaBg }, line: { color: deltaBg } });
      sFm.addText(row.tc_delta || "", { x: rx + 0.05, y: ry + 0.09, w: 0.72, h: 0.22, fontSize: 9, bold: true, color: deltaTxt, fontFace: "DM Sans", align: "center", valign: "middle" });
      rx += fmColW[5];
      // Revenue
      sFm.addText(row.revenue || "", { x: rx, y: ry + 0.07, w: fmColW[6], h: 0.26, fontSize: 9.5, color: DARK, fontFace: "DM Sans", align: "right" });
      rx += fmColW[6];
      // ΔRevenue badge
      const revUp = row.revenue_delta_up === true;
      const revDeltaBg  = revUp ? GREEN_BG  : RED_BG;
      const revDeltaTxt = revUp ? GREEN     : RED;
      sFm.addShape(pres.shapes.RECTANGLE, { x: rx + 0.04, y: ry + 0.09, w: 0.48, h: 0.22, fill: { color: revDeltaBg }, line: { color: revDeltaBg } });
      sFm.addText(row.revenue_delta || "", { x: rx + 0.04, y: ry + 0.09, w: 0.48, h: 0.22, fontSize: 8, bold: true, color: revDeltaTxt, fontFace: "DM Sans", align: "center", valign: "middle" });
    });

    // Insight box (opcional)
    const fmTableBottom = fmY0 + 0.36 + Math.min(fuenteMedio.length, 10) * 0.37;
    if (DATA.FUENTE_MEDIO_INSIGHT && fmTableBottom < 5.1) {
      sFm.addShape(pres.shapes.RECTANGLE, { x: 0.4, y: fmTableBottom + 0.06, w: 9.2, h: 0.38, fill: { color: "FFF0EB" }, line: { color: "FA5A1E", width: 0.5 } });
      sFm.addShape(pres.shapes.RECTANGLE, { x: 0.4, y: fmTableBottom + 0.06, w: 0.08, h: 0.38, fill: { color: ORANGE }, line: { color: ORANGE } });
      sFm.addText(DATA.FUENTE_MEDIO_INSIGHT, { x: 0.6, y: fmTableBottom + 0.08, w: 8.9, h: 0.34, fontSize: 9.5, color: DARK, fontFace: "DM Sans", valign: "middle" });
    }
  }

  // ── SLIDE 3C – CANAL AGENTES (CHAIDE – OPCIONAL) ─────────────────────────
  if (DATA.CHAIDE_VENTAS_AGENTES_ACTUAL) {
    // Parsea string ARS ("$12.500.000") a número
    const parseARS = str => parseFloat((str || "0").replace(/[^0-9,]/g, "").replace(",", ".")) || 0;
    const fmtARS   = n   => "$" + Math.round(n).toLocaleString("es-AR");
    const fmtDelta = n   => (n >= 0 ? "+" : "") + n.toFixed(1).replace(".", ",") + "%";

    const agActual = parseARS(DATA.CHAIDE_VENTAS_AGENTES_ACTUAL);
    const agPrev   = parseARS(DATA.CHAIDE_VENTAS_AGENTES_PREV);
    const agDelta  = agPrev !== 0 ? ((agActual - agPrev) / agPrev) * 100 : 0;

    if (!DATA.CHAIDE_VENTAS_AGENTES_DELTA) DATA.CHAIDE_VENTAS_AGENTES_DELTA = fmtDelta(agDelta);
    if (DATA.CHAIDE_VENTAS_AGENTES_UP == null) DATA.CHAIDE_VENTAS_AGENTES_UP = agActual >= agPrev;

    const ga4Actual  = parseARS(DATA.VTEX_INGRESOS_ACTUAL);
    const ga4Prev    = parseARS(DATA.VTEX_INGRESOS_ANTERIOR);
    const consActual = ga4Actual + agActual;
    const consPrev   = ga4Prev + agPrev;
    const consDelta  = consPrev !== 0 ? ((consActual - consPrev) / consPrev) * 100 : 0;

    if (!DATA.CHAIDE_CONSOLIDADO_ACTUAL)   DATA.CHAIDE_CONSOLIDADO_ACTUAL   = fmtARS(consActual);
    if (!DATA.CHAIDE_CONSOLIDADO_PREV)     DATA.CHAIDE_CONSOLIDADO_PREV     = fmtARS(consPrev);
    if (!DATA.CHAIDE_CONSOLIDADO_DELTA)    DATA.CHAIDE_CONSOLIDADO_DELTA    = fmtDelta(consDelta);
    if (DATA.CHAIDE_CONSOLIDADO_DELTA_UP == null) DATA.CHAIDE_CONSOLIDADO_DELTA_UP = consActual >= consPrev;

    let sAg = pres.addSlide();
    sAg.background = { color: WHITE };
    sAg.addText("Canal Agentes", { x: 0.5, y: 0.18, w: 7, h: 0.52, fontSize: 28, bold: true, color: DARK, fontFace: "Trebuchet MS" });
    sAg.addText(`${DATA.PERIODO_ACTUAL_LABEL || ""} vs ${DATA.PERIODO_ANTERIOR_LABEL || ""}`, { x: 0.5, y: 0.71, w: 7, h: 0.28, fontSize: 13, color: GRAY_TEXT, fontFace: "DM Sans" });

    // ── Sección 1: 3 KPI cards ────────────────────────────────────────────
    sAg.addText("Ventas Canal Agentes", { x: 0.5, y: 1.1, w: 9, h: 0.28, fontSize: 11, bold: true, color: DARK, fontFace: "DM Sans" });

    const agKPIs = [
      { icon: "$", label: "Ventas Agentes",  sub: DATA.PERIODO_ACTUAL_LABEL   || "Actual",   val: DATA.CHAIDE_VENTAS_AGENTES_ACTUAL || "", isDelta: false },
      { icon: "$", label: "Ventas Agentes",  sub: DATA.PERIODO_ANTERIOR_LABEL || "Anterior", val: DATA.CHAIDE_VENTAS_AGENTES_PREV   || "", isDelta: false },
      { icon: "Δ", label: "Variación",       sub: "vs período anterior",                     val: DATA.CHAIDE_VENTAS_AGENTES_DELTA  || "", isDelta: true, up: DATA.CHAIDE_VENTAS_AGENTES_UP === true },
    ];
    agKPIs.forEach((k, i) => {
      const x = 0.4 + i * 3.1, y = 1.42;
      const cardBg  = k.isDelta ? (k.up ? GREEN_BG : RED_BG) : LIGHT_BG;
      const cardBdr = k.isDelta ? (k.up ? GREEN_BG : RED_BG) : "F0E8E0";
      const valColor = k.isDelta ? (k.up ? GREEN : RED) : DARK;
      sAg.addShape(pres.shapes.RECTANGLE, { x, y, w: 2.9, h: 1.1, fill: { color: cardBg }, line: { color: cardBdr, width: 0.5 } });
      sAg.addShape(pres.shapes.RECTANGLE, { x, y, w: 2.9, h: 0.06, fill: { color: ORANGE }, line: { color: ORANGE } });
      sAg.addShape(pres.shapes.OVAL, { x: x + 0.14, y: y + 0.14, w: 0.3, h: 0.3, fill: { color: ORANGE }, line: { color: ORANGE } });
      sAg.addText(k.label, { x: x + 0.52, y: y + 0.14, w: 2.2,  h: 0.18, fontSize: 9,  bold: true, color: DARK,     fontFace: "DM Sans" });
      sAg.addText(k.val,   { x: x + 0.14, y: y + 0.52, w: 2.62, h: 0.48, fontSize: 22, bold: true, color: valColor,  fontFace: "Trebuchet MS", align: "center" });
    });

    // ── Sección 2: Consolidado Ecommerce + Agentes ────────────────────────
    sAg.addShape(pres.shapes.RECTANGLE, { x: 0.4, y: 2.68, w: 9.2, h: 0.04, fill: { color: "E8E0D8" }, line: { color: "E8E0D8" } });
    sAg.addText("Consolidado Total · Ecommerce + Agentes", { x: 0.5, y: 2.76, w: 9, h: 0.28, fontSize: 11, bold: true, color: DARK, fontFace: "DM Sans" });

    const consolidadoCols = [
      { label: DATA.PERIODO_ACTUAL_LABEL   || "Actual",   ga4: DATA.VTEX_INGRESOS_ACTUAL   || "N/D", agentes: DATA.CHAIDE_VENTAS_AGENTES_ACTUAL || "", total: DATA.CHAIDE_CONSOLIDADO_ACTUAL || "" },
      { label: DATA.PERIODO_ANTERIOR_LABEL || "Anterior", ga4: DATA.VTEX_INGRESOS_ANTERIOR || "N/D", agentes: DATA.CHAIDE_VENTAS_AGENTES_PREV   || "", total: DATA.CHAIDE_CONSOLIDADO_PREV   || "" },
    ];
    consolidadoCols.forEach((col, i) => {
      const x = 0.4 + i * 4.7, y = 3.08;
      const isActual = i === 0;
      sAg.addShape(pres.shapes.RECTANGLE, { x, y, w: 4.5, h: 1.95, fill: { color: isActual ? LIGHT_BG : "F5F5F5" }, line: { color: isActual ? "F0E8E0" : "E0E0E0", width: 0.5 } });
      sAg.addShape(pres.shapes.RECTANGLE, { x, y, w: 4.5, h: 0.06, fill: { color: isActual ? ORANGE : GRAY_TEXT }, line: { color: isActual ? ORANGE : GRAY_TEXT } });
      sAg.addText(col.label,  { x: x + 0.2, y: y + 0.12, w: 4.1, h: 0.26, fontSize: 12, bold: true, color: DARK,      fontFace: "DM Sans" });
      sAg.addText("Ecommerce (VTEX)",  { x: x + 0.2, y: y + 0.46, w: 2.2, h: 0.22, fontSize: 9, color: GRAY_TEXT, fontFace: "DM Sans" });
      sAg.addText(col.ga4,          { x: x + 2.4, y: y + 0.46, w: 1.9, h: 0.22, fontSize: 10, bold: true, color: DARK, fontFace: "DM Sans", align: "right" });
      sAg.addText("Canal Agentes",  { x: x + 0.2, y: y + 0.7,  w: 2.2, h: 0.22, fontSize: 9, color: GRAY_TEXT, fontFace: "DM Sans" });
      sAg.addText(col.agentes,      { x: x + 2.4, y: y + 0.7,  w: 1.9, h: 0.22, fontSize: 10, bold: true, color: DARK, fontFace: "DM Sans", align: "right" });
      sAg.addShape(pres.shapes.RECTANGLE, { x: x + 0.2, y: y + 0.98, w: 4.1, h: 0.02, fill: { color: "E8E0D8" }, line: { color: "E8E0D8" } });
      sAg.addText("Total",   { x: x + 0.2, y: y + 1.04, w: 2.0, h: 0.3, fontSize: 10, bold: true, color: DARK, fontFace: "DM Sans" });
      sAg.addText(col.total, { x: x + 2.4, y: y + 1.04, w: 1.9, h: 0.3, fontSize: 14, bold: true, color: DARK, fontFace: "Trebuchet MS", align: "right" });
      if (isActual && DATA.CHAIDE_CONSOLIDADO_DELTA) {
        const revUp = DATA.CHAIDE_CONSOLIDADO_DELTA_UP === true;
        sAg.addShape(pres.shapes.RECTANGLE, { x: x + 0.2, y: y + 1.48, w: 1.5, h: 0.26, fill: { color: revUp ? GREEN_BG : RED_BG }, line: { color: revUp ? GREEN_BG : RED_BG } });
        sAg.addText(DATA.CHAIDE_CONSOLIDADO_DELTA, { x: x + 0.2, y: y + 1.48, w: 1.5, h: 0.26, fontSize: 11, bold: true, color: revUp ? GREEN : RED, fontFace: "DM Sans", align: "center", valign: "middle" });
        sAg.addText(`vs ${DATA.PERIODO_ANTERIOR_LABEL || "período anterior"}`, { x: x + 1.8, y: y + 1.5, w: 2.5, h: 0.22, fontSize: 9, color: GRAY_TEXT, fontFace: "DM Sans" });
      }
    });
  }



  // ── SLIDE 4 – META ADS DETALLE ────────────────────────────────────────────
  let s3 = pres.addSlide();
  s3.background = { color: WHITE };
  s3.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 1.08, fill: { color: ORANGE }, line: { color: ORANGE } });
  s3.addText("Meta Ads", { x: 0.5, y: 0.22, w: 6, h: 0.52, fontSize: 28, bold: true, color: WHITE, fontFace: "Trebuchet MS" });
  s3.addText(`${DATA.PERIODO_ACTUAL_LABEL || ""} vs ${DATA.PERIODO_ANTERIOR_LABEL || ""}`, { x: 0.5, y: 0.72, w: 6, h: 0.3, fontSize: 13, color: "FFD4B8", fontFace: "DM Sans" });
  s3.addShape(pres.shapes.RECTANGLE, { x: 7.2, y: 0.35, w: 2.3, h: 0.5, fill: { color: WHITE, transparency: 20 }, line: { color: WHITE, transparency: 50 } });
  s3.addText(`Inversión: ${DATA.META_COSTO || ""}`, { x: 7.2, y: 0.35, w: 2.3, h: 0.5, fontSize: 13, bold: true, color: WHITE, fontFace: "DM Sans", align: "center" });

  const metaKPIs = [
    { label: "Costo",       val: fmtMoneyCompact(DATA.META_COSTO), prev: DATA.META_COSTO_PREV        || "", delta: DATA.META_COSTO_DELTA        || "", up: DATA.META_COSTO_DELTA_UP        === true, warn: false },
    { label: "Clicks",      val: DATA.META_CLICKS       || "", prev: DATA.META_CLICKS_PREV       || "", delta: DATA.META_CLICKS_DELTA       || "", up: DATA.META_CLICKS_DELTA_UP       === true, warn: false },
    { label: "Impresiones", val: DATA.META_IMPRESIONES  || "", prev: DATA.META_IMPRESIONES_PREV  || "", delta: DATA.META_IMPRESIONES_DELTA  || "", up: DATA.META_IMPRESIONES_DELTA_UP  === true, warn: false },
    { label: "CTR",         val: DATA.META_CTR          || "", prev: DATA.META_CTR_PREV          || "", delta: DATA.META_CTR_DELTA          || "", up: DATA.META_CTR_DELTA_UP          === true, warn: false },
    { label: "CPC",         val: DATA.META_CPC          || "", prev: DATA.META_CPC_PREV          || "", delta: DATA.META_CPC_DELTA          || "", up: DATA.META_CPC_DELTA_UP          === true, warn: false },
    { label: "ROAS",        val: DATA.META_ROAS         || "", prev: DATA.META_ROAS_PREV         || "", delta: DATA.META_ROAS_DELTA         || "", up: DATA.META_ROAS_DELTA_UP         === true, warn: false },
  ];
  metaKPIs.forEach((k, i) => {
    const col = i % 3, row = Math.floor(i / 3);
    const x = 0.4 + col * 3.1, y = 1.3 + row * 1.6;
    s3.addShape(pres.shapes.RECTANGLE, { x, y, w: 2.8, h: 1.45, fill: { color: k.warn ? "FFF5F5" : LIGHT_BG }, line: { color: k.warn ? "F7C1C1" : "F0E8E0", width: 0.5 } });
    s3.addText(k.label, { x: x + 0.15, y: y + 0.12, w: 2.5, h: 0.28, fontSize: 11, color: GRAY_TEXT, fontFace: "DM Sans" });
    s3.addText(k.val,   { x: x + 0.15, y: y + 0.38, w: 2.5, h: 0.5,  fontSize: 24, bold: true, color: DARK, fontFace: "Trebuchet MS" });
    s3.addText(`${DATA.PERIODO_ANTERIOR_LABEL ? DATA.PERIODO_ANTERIOR_LABEL.split(" ")[0] : "Ant."}: ${k.prev}`, { x: x + 0.15, y: y + 0.88, w: 1.6, h: 0.25, fontSize: 10, color: GRAY_TEXT, fontFace: "DM Sans" });
    s3.addShape(pres.shapes.RECTANGLE, { x: x + 1.9, y: y + 0.88, w: 0.75, h: 0.25, fill: { color: k.up ? GREEN_BG : RED_BG }, line: { color: k.up ? GREEN_BG : RED_BG } });
    s3.addText(k.delta, { x: x + 1.9, y: y + 0.88, w: 0.75, h: 0.25, fontSize: 10, bold: true, color: k.up ? GREEN : RED, fontFace: "DM Sans", align: "center" });
  });

  s3.addShape(pres.shapes.RECTANGLE, { x: 0.4, y: 4.7, w: 9.2, h: 0.65, fill: { color: "FFF0EB" }, line: { color: "FA5A1E", width: 0.5 } });
  s3.addShape(pres.shapes.RECTANGLE, { x: 0.4, y: 4.7, w: 0.08, h: 0.65, fill: { color: ORANGE }, line: { color: ORANGE } });
  s3.addText(DATA.META_ALERTA || "", { x: 0.6, y: 4.72, w: 8.9, h: 0.6, fontSize: 11, color: DARK, fontFace: "DM Sans", valign: "middle" });

  // ── SLIDE 4 – GOOGLE ADS DETALLE ──────────────────────────────────────────
  let s4 = pres.addSlide();
  s4.background = { color: WHITE };
  s4.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 1.08, fill: { color: BLUE }, line: { color: BLUE } });
  s4.addText("Google Ads", { x: 0.5, y: 0.22, w: 6, h: 0.52, fontSize: 28, bold: true, color: WHITE, fontFace: "Trebuchet MS" });
  s4.addText(`${DATA.PERIODO_ACTUAL_LABEL || ""} vs ${DATA.PERIODO_ANTERIOR_LABEL || ""}`, { x: 0.5, y: 0.72, w: 6, h: 0.3, fontSize: 13, color: "B5D4F4", fontFace: "DM Sans" });
  s4.addShape(pres.shapes.RECTANGLE, { x: 7.2, y: 0.35, w: 2.3, h: 0.5, fill: { color: WHITE, transparency: 20 }, line: { color: WHITE, transparency: 50 } });
  s4.addText(`Inversión: ${DATA.GOOGLE_COSTO || ""}`, { x: 7.2, y: 0.35, w: 2.3, h: 0.5, fontSize: 13, bold: true, color: WHITE, fontFace: "DM Sans", align: "center" });

  const googleKPIs = [
    { label: "Costo",       val: fmtMoneyCompact(DATA.GOOGLE_COSTO), prev: DATA.GOOGLE_COSTO_PREV       || "", delta: DATA.GOOGLE_COSTO_DELTA       || "", good: DATA.GOOGLE_COSTO_DELTA_UP       === true },
    { label: "Clicks",      val: DATA.GOOGLE_CLICKS      || "", prev: DATA.GOOGLE_CLICKS_PREV      || "", delta: DATA.GOOGLE_CLICKS_DELTA      || "", good: DATA.GOOGLE_CLICKS_DELTA_UP      === true },
    { label: "Impresiones", val: DATA.GOOGLE_IMPRESIONES || "", prev: DATA.GOOGLE_IMPRESIONES_PREV || "", delta: DATA.GOOGLE_IMPRESIONES_DELTA || "", good: DATA.GOOGLE_IMPRESIONES_DELTA_UP === true },
    { label: "CTR",         val: DATA.GOOGLE_CTR         || "", prev: DATA.GOOGLE_CTR_PREV         || "", delta: DATA.GOOGLE_CTR_DELTA         || "", good: DATA.GOOGLE_CTR_DELTA_UP         === true },
    { label: "CPC",         val: DATA.GOOGLE_CPC         || "", prev: DATA.GOOGLE_CPC_PREV         || "", delta: DATA.GOOGLE_CPC_DELTA         || "", good: DATA.GOOGLE_CPC_DELTA_UP         === true },
    { label: "ROAS",        val: DATA.GOOGLE_ROAS        || "", prev: DATA.GOOGLE_ROAS_PREV        || "", delta: DATA.GOOGLE_ROAS_DELTA        || "", good: DATA.GOOGLE_ROAS_DELTA_UP        === true },
  ];
  googleKPIs.forEach((k, i) => {
    const col = i % 3, row = Math.floor(i / 3);
    const x = 0.4 + col * 3.1, y = 1.3 + row * 1.6;
    s4.addShape(pres.shapes.RECTANGLE, { x, y, w: 2.8, h: 1.45, fill: { color: k.good ? LIGHT_BLUE : LIGHT_BG }, line: { color: k.good ? "B5D4F4" : "F0E8E0", width: 0.5 } });
    s4.addText(k.label, { x: x + 0.15, y: y + 0.12, w: 2.5, h: 0.28, fontSize: 11, color: GRAY_TEXT, fontFace: "DM Sans" });
    s4.addText(k.val,   { x: x + 0.15, y: y + 0.38, w: 2.5, h: 0.5,  fontSize: 24, bold: true, color: DARK, fontFace: "Trebuchet MS" });
    s4.addText(`${DATA.PERIODO_ANTERIOR_LABEL ? DATA.PERIODO_ANTERIOR_LABEL.split(" ")[0] : "Ant."}: ${k.prev}`, { x: x + 0.15, y: y + 0.88, w: 1.6, h: 0.25, fontSize: 10, color: GRAY_TEXT, fontFace: "DM Sans" });
    s4.addShape(pres.shapes.RECTANGLE, { x: x + 1.9, y: y + 0.88, w: 0.75, h: 0.25, fill: { color: k.good ? GREEN_BG : RED_BG }, line: { color: k.good ? GREEN_BG : RED_BG } });
    s4.addText(k.delta, { x: x + 1.9, y: y + 0.88, w: 0.75, h: 0.25, fontSize: 10, bold: true, color: k.good ? GREEN : RED, fontFace: "DM Sans", align: "center" });
  });

  s4.addShape(pres.shapes.RECTANGLE, { x: 0.4, y: 4.7, w: 9.2, h: 0.65, fill: { color: "EAF3DE" }, line: { color: "63992250", width: 0.5 } });
  s4.addShape(pres.shapes.RECTANGLE, { x: 0.4, y: 4.7, w: 0.08, h: 0.65, fill: { color: "3B6D11" }, line: { color: "3B6D11" } });
  s4.addText(DATA.GOOGLE_ALERTA || "", { x: 0.6, y: 4.72, w: 8.9, h: 0.6, fontSize: 11, color: DARK, fontFace: "DM Sans", valign: "middle" });

  // ── SLIDE 5 – TOP CAMPAÑAS POR ROAS ──────────────────────────────────────
  let s5 = pres.addSlide();
  s5.background = { color: WHITE };
  s5.addText("Top Campañas por ROAS", { x: 0.5, y: 0.2, w: 7, h: 0.55, fontSize: 28, bold: true, color: DARK, fontFace: "Trebuchet MS" });
  s5.addText(`${DATA.PERIODO_ACTUAL_LABEL || ""}  ·  Google Ads + Meta Ads`, { x: 0.5, y: 0.76, w: 7, h: 0.3, fontSize: 13, color: GRAY_TEXT, fontFace: "DM Sans" });

  // campaigns: array of { nombre, plataforma, inversion, clicks, roas, nivel }
  // nivel: "high" | "mid" | "low"
  const campaigns = DATA.CAMPANAS || [];
  s5.addShape(pres.shapes.RECTANGLE, { x: 0.4, y: 1.2, w: 9.2, h: 0.38, fill: { color: DARK }, line: { color: DARK } });
  const headers = ["Campaña", "Plat.", "Inversión", "Clicks", "ROAS"];
  const colW    = [3.6, 0.85, 1.25, 1.15, 1.35];
  let cx = 0.55;
  headers.forEach((h, i) => {
    s5.addText(h, { x: cx, y: 1.22, w: colW[i], h: 0.34, fontSize: 10, bold: true, color: WHITE, fontFace: "DM Sans", valign: "middle" });
    cx += colW[i];
  });

  campaigns.slice(0, 8).forEach((row, i) => {
    const y  = 1.6 + i * 0.44;
    const bg = i % 2 === 0 ? WHITE : LIGHT_BG;
    s5.addShape(pres.shapes.RECTANGLE, { x: 0.4, y, w: 9.2, h: 0.43, fill: { color: bg }, line: { color: "E8E0D8", width: 0.5 } });
    let rx = 0.55;

    s5.addText(row.nombre || "", { x: rx, y: y + 0.07, w: colW[0], h: 0.3, fontSize: 10, color: DARK, fontFace: "DM Sans" });
    rx += colW[0];

    const isGoogle = (row.plataforma || "").toLowerCase() === "google";
    s5.addShape(pres.shapes.RECTANGLE, { x: rx, y: y + 0.1, w: 0.78, h: 0.24, fill: { color: isGoogle ? LIGHT_BLUE : "FFF0EB" }, line: { color: isGoogle ? "B5D4F4" : "F5C4B3" } });
    s5.addText(row.plataforma || "", { x: rx, y: y + 0.1, w: 0.78, h: 0.24, fontSize: 9, color: isGoogle ? BLUE : ORANGE, fontFace: "DM Sans", align: "center", bold: true });
    rx += colW[1];

    s5.addText(row.costo || "", { x: rx, y: y + 0.07, w: colW[2], h: 0.3, fontSize: 10, color: DARK, fontFace: "DM Sans", align: "right" });
    rx += colW[2];
    s5.addText(row.clicks || "", { x: rx, y: y + 0.07, w: colW[3], h: 0.3, fontSize: 10, color: DARK, fontFace: "DM Sans", align: "right" });
    rx += colW[3];

    const nivel    = row.nivel || "mid";
    const roasColor = nivel === "high" ? GREEN : nivel === "mid" ? AMBER : RED;
    const roasBg    = nivel === "high" ? GREEN_BG : nivel === "mid" ? AMBER_BG : RED_BG;
    s5.addShape(pres.shapes.RECTANGLE, { x: rx, y: y + 0.1, w: 1.0, h: 0.24, fill: { color: roasBg }, line: { color: roasBg } });
    s5.addText(row.roas || "", { x: rx, y: y + 0.1, w: 1.0, h: 0.24, fontSize: 10, bold: true, color: roasColor, fontFace: "DM Sans", align: "center" });
  });

  s5.addShape(pres.shapes.RECTANGLE, { x: 0.4,  y: 5.1, w: 0.55, h: 0.2, fill: { color: GREEN_BG  }, line: { color: GREEN_BG  } });
  s5.addText("ROAS alto (>30x)",    { x: 1.0,  y: 5.1, w: 1.8, h: 0.2, fontSize: 9, color: GRAY_TEXT, fontFace: "DM Sans" });
  s5.addShape(pres.shapes.RECTANGLE, { x: 2.9,  y: 5.1, w: 0.55, h: 0.2, fill: { color: AMBER_BG  }, line: { color: AMBER_BG  } });
  s5.addText("ROAS medio (5-30x)", { x: 3.5,  y: 5.1, w: 1.9, h: 0.2, fontSize: 9, color: GRAY_TEXT, fontFace: "DM Sans" });
  s5.addShape(pres.shapes.RECTANGLE, { x: 5.5,  y: 5.1, w: 0.55, h: 0.2, fill: { color: RED_BG    }, line: { color: RED_BG    } });
  s5.addText("ROAS bajo (<5x)",     { x: 6.1,  y: 5.1, w: 1.6, h: 0.2, fontSize: 9, color: GRAY_TEXT, fontFace: "DM Sans" });

  // ── SLIDE 2 – RESUMEN EJECUTIVO ───────────────────────────────────────────
  let s2 = pres.addSlide();
  s2.background = { color: WHITE };
  s2.addText("Resumen Ejecutivo", { x: 0.5, y: 0.22, w: 7, h: 0.55, fontSize: 28, bold: true, color: DARK, fontFace: "Trebuchet MS" });
  s2.addText("Inversión total · Meta Ads + Google Ads", { x: 0.5, y: 0.78, w: 7, h: 0.3, fontSize: 13, color: GRAY_TEXT, fontFace: "DM Sans" });

  const kpis = [
    { label: "Inversión total", val: fmtMoneyCompact(DATA.INVERSION_TOTAL), delta: DATA.INVERSION_DELTA || "", note: `${DATA.PERIODO_ANTERIOR_LABEL || "Año ant."}: ${DATA.INVERSION_PREV || ""}`, up: DATA.INVERSION_DELTA_UP === true },
    { label: "Clicks totales",  val: DATA.CLICKS_TOTAL || "", delta: DATA.CLICKS_DELTA || "", note: `${DATA.PERIODO_ANTERIOR_LABEL || "Año ant."}: ${DATA.CLICKS_PREV || ""}`, up: DATA.CLICKS_DELTA_UP === true },
    { label: "Impresiones",     val: DATA.IMPRESIONES_TOTAL || "", delta: DATA.IMPRESIONES_DELTA || "", note: `${DATA.PERIODO_ANTERIOR_LABEL || "Año ant."}: ${DATA.IMPRESIONES_PREV || ""}`, up: DATA.IMPRESIONES_DELTA_UP === true },
    { label: "CPC promedio",    val: DATA.CPC_TOTAL || "", delta: DATA.CPC_DELTA || "", note: `${DATA.PERIODO_ANTERIOR_LABEL || "Año ant."}: ${DATA.CPC_PREV || ""}`, up: DATA.CPC_DELTA_UP === true },
  ];
  kpis.forEach((k, i) => {
    const x = 0.4 + i * 2.32;
    s2.addShape(pres.shapes.RECTANGLE, { x, y: 1.2, w: 2.1, h: 1.55, fill: { color: LIGHT_BG }, line: { color: "F0E8E0", width: 0.5 } });
    s2.addShape(pres.shapes.RECTANGLE, { x, y: 1.2, w: 2.1, h: 0.07, fill: { color: ORANGE }, line: { color: ORANGE } });
    s2.addText(k.label, { x, y: 1.32, w: 2.1, h: 0.3, fontSize: 10, color: GRAY_TEXT, fontFace: "DM Sans", align: "center" });
    s2.addText(k.val,   { x, y: 1.62, w: 2.1, h: 0.52, fontSize: 26, bold: true, color: DARK, fontFace: "Trebuchet MS", align: "center" });
    s2.addShape(pres.shapes.RECTANGLE, { x: x + 0.6, y: 2.17, w: 0.9, h: 0.27, fill: { color: k.up ? GREEN_BG : RED_BG }, line: { color: k.up ? GREEN_BG : RED_BG } });
    s2.addText(k.delta, { x: x + 0.6, y: 2.17, w: 0.9, h: 0.27, fontSize: 11, bold: true, color: k.up ? GREEN : RED, fontFace: "DM Sans", align: "center" });
    s2.addText(k.note,  { x, y: 2.5, w: 2.1, h: 0.25, fontSize: 9, color: GRAY_TEXT, fontFace: "DM Sans", align: "center" });
  });

  s2.addText("Comparativa por plataforma", { x: 0.5, y: 2.95, w: 9, h: 0.35, fontSize: 13, bold: true, color: DARK, fontFace: "DM Sans" });

  // Meta block
  s2.addShape(pres.shapes.RECTANGLE, { x: 0.4, y: 3.35, w: 4.4, h: 1.85, fill: { color: LIGHT_BG }, line: { color: "F0E8E0", width: 0.5 } });
  s2.addShape(pres.shapes.RECTANGLE, { x: 0.4, y: 3.35, w: 4.4, h: 0.38, fill: { color: ORANGE }, line: { color: ORANGE } });
  s2.addText("Meta Ads", { x: 0.55, y: 3.38, w: 3, h: 0.32, fontSize: 13, bold: true, color: WHITE, fontFace: "DM Sans" });
  const metaStats = [
    ["Costo",  DATA.META_COSTO  || "", DATA.META_COSTO_DELTA  || "", DATA.META_COSTO_DELTA_UP  === true],
    ["Clicks", DATA.META_CLICKS || "", DATA.META_CLICKS_DELTA || "", DATA.META_CLICKS_DELTA_UP !== true],
    ["ROAS",   DATA.META_ROAS   || "", DATA.META_ROAS_DELTA   || "", DATA.META_ROAS_DELTA_UP   !== true],
    ["CPC",    DATA.META_CPC    || "", DATA.META_CPC_DELTA    || "", DATA.META_CPC_DELTA_UP    !== true],
  ];
  metaStats.forEach(([lbl, val, delta, isDown], i) => {
    const col = i % 2, row = Math.floor(i / 2);
    const bx = 0.55 + col * 2.1, by = 3.85 + row * 0.6;
    s2.addText(lbl, { x: bx, y: by, w: 1.8, h: 0.22, fontSize: 9, color: GRAY_TEXT, fontFace: "DM Sans" });
    s2.addText([
      { text: val + "  ", options: { bold: true, color: DARK } },
      { text: delta,      options: { color: isDown ? RED : GREEN, bold: true } },
    ], { x: bx, y: by + 0.2, w: 1.9, h: 0.28, fontSize: 12, fontFace: "DM Sans" });
  });

  // Google block
  s2.addShape(pres.shapes.RECTANGLE, { x: 5.2, y: 3.35, w: 4.4, h: 1.85, fill: { color: LIGHT_BLUE }, line: { color: "D0E4F5", width: 0.5 } });
  s2.addShape(pres.shapes.RECTANGLE, { x: 5.2, y: 3.35, w: 4.4, h: 0.38, fill: { color: BLUE }, line: { color: BLUE } });
  s2.addText("Google Ads", { x: 5.35, y: 3.38, w: 3, h: 0.32, fontSize: 13, bold: true, color: WHITE, fontFace: "DM Sans" });
  const googleStats = [
    ["Costo",  DATA.GOOGLE_COSTO  || "", DATA.GOOGLE_COSTO_DELTA  || "", DATA.GOOGLE_COSTO_DELTA_UP  === true],
    ["Clicks", DATA.GOOGLE_CLICKS || "", DATA.GOOGLE_CLICKS_DELTA || "", DATA.GOOGLE_CLICKS_DELTA_UP !== true],
    ["ROAS",   DATA.GOOGLE_ROAS   || "", DATA.GOOGLE_ROAS_DELTA   || "", DATA.GOOGLE_ROAS_DELTA_UP   !== true],
    ["CPC",    DATA.GOOGLE_CPC    || "", DATA.GOOGLE_CPC_DELTA    || "", DATA.GOOGLE_CPC_DELTA_UP    !== true],
  ];
  googleStats.forEach(([lbl, val, delta, isDown], i) => {
    const col = i % 2, row = Math.floor(i / 2);
    const bx = 5.35 + col * 2.1, by = 3.85 + row * 0.6;
    s2.addText(lbl, { x: bx, y: by, w: 1.8, h: 0.22, fontSize: 9, color: GRAY_TEXT, fontFace: "DM Sans" });
    s2.addText([
      { text: val + "  ", options: { bold: true, color: DARK } },
      { text: delta,      options: { color: isDown ? RED : GREEN, bold: true } },
    ], { x: bx, y: by + 0.2, w: 1.9, h: 0.28, fontSize: 12, fontFace: "DM Sans" });
  });


  // ── SLIDE 5C – TOP ANUNCIOS META POR COMPRAS ─────────────────────────────
  if (DATA.TOP_ANUNCIOS_META_TIENE_DATOS && Array.isArray(DATA.TOP_ANUNCIOS_META) && DATA.TOP_ANUNCIOS_META.length > 0) {
    const ads = DATA.TOP_ANUNCIOS_META.slice(0, 3);

    let s5b = pres.addSlide();
    s5b.background = { color: WHITE };
    s5b.addText("Top Anuncios Meta por Compras", { x: 0.5, y: 0.2, w: 7, h: 0.55, fontSize: 28, bold: true, color: DARK, fontFace: "Trebuchet MS" });
    s5b.addText(`${DATA.PERIODO_ACTUAL_LABEL || ""}  ·  Ordenados por ROAS`, { x: 0.5, y: 0.76, w: 7, h: 0.3, fontSize: 13, color: GRAY_TEXT, fontFace: "DM Sans" });

    const cardW = 2.8, cardH = 3.8, cardGap = 0.3;
    const totalW = ads.length * cardW + (ads.length - 1) * cardGap;
    const startX = (10 - totalW) / 2;

    ads.forEach((ad, i) => {
      const cx = startX + i * (cardW + cardGap);
      const cy = 1.2;

      // Card background
      s5b.addShape(pres.shapes.RECTANGLE, { x: cx, y: cy, w: cardW, h: cardH, fill: { color: LIGHT_BG }, line: { color: "F0E8E0", width: 0.5 } });
      s5b.addShape(pres.shapes.RECTANGLE, { x: cx, y: cy, w: cardW, h: 0.06, fill: { color: ORANGE }, line: { color: ORANGE } });

      // Image area (1.8 x 1.8 centered, or placeholder)
      const imgSize = 1.8;
      const imgX = cx + (cardW - imgSize) / 2;
      const imgY = cy + 0.2;

      if (ad.preview_url) {
        // Preview link button
        s5b.addShape(pres.shapes.RECTANGLE, { x: imgX, y: imgY, w: imgSize, h: imgSize, fill: { color: "FFF4EC" }, line: { color: ORANGE, width: 1 } });
        s5b.addShape(pres.shapes.OVAL, { x: imgX + 0.64, y: imgY + 0.35, w: 0.52, h: 0.52, fill: { color: ORANGE }, line: { color: ORANGE } });
        s5b.addText("▶", { x: imgX + 0.64, y: imgY + 0.35, w: 0.52, h: 0.52, fontSize: 16, color: WHITE, fontFace: "DM Sans", align: "center", valign: "middle" });
        s5b.addText("Ver anuncio", { x: imgX, y: imgY + 0.98, w: imgSize, h: 0.28, fontSize: 10, bold: true, color: DARK, fontFace: "DM Sans", align: "center" });
        s5b.addText([{ text: "Abrir preview →", options: { hyperlink: { url: ad.preview_url } } }], { x: imgX, y: imgY + 1.28, w: imgSize, h: 0.25, fontSize: 9, color: ORANGE, fontFace: "DM Sans", align: "center" });
      } else {
        // Gray placeholder sin link
        s5b.addShape(pres.shapes.RECTANGLE, { x: imgX, y: imgY, w: imgSize, h: imgSize, fill: { color: "E0E0E0" }, line: { color: "D0D0D0", width: 0.5 } });
        s5b.addText("Sin imagen", { x: imgX, y: imgY, w: imgSize, h: imgSize, fontSize: 10, color: GRAY_TEXT, fontFace: "DM Sans", align: "center", valign: "middle" });
      }

      // Ad name (truncated)
      const nombre = (ad.nombre || "").length > 45 ? (ad.nombre || "").substring(0, 42) + "..." : (ad.nombre || "");
      s5b.addText(nombre, { x: cx + 0.12, y: imgY + imgSize + 0.1, w: cardW - 0.24, h: 0.45, fontSize: 9, bold: true, color: DARK, fontFace: "DM Sans", valign: "top" });

      // Metrics grid (2x2)
      const metricsY = imgY + imgSize + 0.55;
      const metricsList = [
        { lbl: "Compras (Pixel)", val: ad.conversiones || "0" },
        { lbl: "ROAS",         val: ad.roas || "0x" },
        { lbl: "Costo",        val: ad.costo || "$0" },
        { lbl: "Clicks",       val: ad.clicks || "0" },
      ];
      metricsList.forEach((m, mi) => {
        const mcol = mi % 2, mrow = Math.floor(mi / 2);
        const mx = cx + 0.12 + mcol * 1.35;
        const my = metricsY + mrow * 0.5;
        s5b.addText(m.lbl, { x: mx, y: my, w: 1.3, h: 0.2, fontSize: 8, color: GRAY_TEXT, fontFace: "DM Sans" });
        s5b.addText(m.val, { x: mx, y: my + 0.18, w: 1.3, h: 0.25, fontSize: 12, bold: true, color: DARK, fontFace: "DM Sans" });
      });

      // Rank badge
      s5b.addShape(pres.shapes.OVAL, { x: cx + 0.1, y: cy + 0.12, w: 0.32, h: 0.32, fill: { color: ORANGE }, line: { color: ORANGE } });
      s5b.addText(`#${i + 1}`, { x: cx + 0.1, y: cy + 0.12, w: 0.32, h: 0.32, fontSize: 10, bold: true, color: WHITE, fontFace: "DM Sans", align: "center", valign: "middle" });
    });
  }

  // ── SLIDE 6 – RECOMENDACIONES ─────────────────────────────────────────────
  let s6 = pres.addSlide();
  s6.background = { color: DARK };
  s6.addShape(pres.shapes.OVAL, { x: 7.5, y: 3.5, w: 3.5, h: 3.5, fill: { color: ORANGE, transparency: 92 }, line: { color: ORANGE, transparency: 92 } });
  s6.addShape(pres.shapes.OVAL, { x: -1.0, y: -0.5, w: 2.5, h: 2.5, fill: { color: ORANGE, transparency: 88 }, line: { color: ORANGE, transparency: 88 } });
  s6.addText("Recomendaciones", { x: 0.5, y: 0.22, w: 9, h: 0.55, fontSize: 28, bold: true, color: WHITE, fontFace: "Trebuchet MS" });
  s6.addText("Acciones prioritarias para optimizar la performance", { x: 0.5, y: 0.78, w: 9, h: 0.3, fontSize: 13, color: "FF912D", fontFace: "DM Sans" });

  // RECOMENDACIONES: array of { titulo, texto }
  const recs = DATA.RECOMENDACIONES || [];
  recs.slice(0, 5).forEach((r, i) => {
    const num = String(i + 1).padStart(2, "0");
    const y   = 1.22 + i * 0.82;
    s6.addShape(pres.shapes.OVAL, { x: 0.4, y, w: 0.45, h: 0.45, fill: { color: ORANGE }, line: { color: ORANGE } });
    s6.addText(num, { x: 0.4, y, w: 0.45, h: 0.45, fontSize: 11, bold: true, color: WHITE, fontFace: "DM Sans", align: "center", valign: "middle" });
    s6.addText(r.titulo || "", { x: 1.05, y: y + 0.02, w: 8.5, h: 0.28, fontSize: 13, bold: true, color: WHITE, fontFace: "DM Sans" });
    s6.addText(r.texto  || "", { x: 1.05, y: y + 0.29, w: 8.5, h: 0.25, fontSize: 11, color: "B0B8C8", fontFace: "DM Sans" });
  });

  // ── SLIDE 8 – CIERRE ──────────────────────────────────────────────────────
  let s8 = pres.addSlide();
  s8.background = { color: ORANGE };
  s8.addShape(pres.shapes.OVAL, { x: 6.5,  y: -1.5, w: 5.5, h: 5.5, fill: { color: WHITE, transparency: 92 }, line: { color: WHITE, transparency: 92 } });
  s8.addShape(pres.shapes.OVAL, { x: -2.0, y: 3.0,  w: 4.5, h: 4.5, fill: { color: DARK,  transparency: 88 }, line: { color: DARK,  transparency: 88 } });
  s8.addShape(pres.shapes.OVAL, { x: 0.5, y: 0.5, w: 0.55, h: 0.55, fill: { color: WHITE  }, line: { color: WHITE  } });
  s8.addShape(pres.shapes.OVAL, { x: 0.65, y: 0.65, w: 0.28, h: 0.28, fill: { color: ORANGE }, line: { color: ORANGE } });
  s8.addText("Known Online", { x: 1.2, y: 0.52, w: 4, h: 0.45, fontSize: 16, bold: true, color: WHITE, fontFace: "DM Sans", margin: 0 });
  s8.addText("¡Muchas gracias!", { x: 0.5, y: 1.6, w: 9, h: 1.4, fontSize: 56, bold: true, color: WHITE, fontFace: "Trebuchet MS", align: "center" });
  s8.addText("Logramos tu transformación digital", { x: 0.5, y: 3.1, w: 9, h: 0.45, fontSize: 18, color: "FFD4B8", fontFace: "DM Sans", align: "center", italic: true });
  s8.addShape(pres.shapes.RECTANGLE, { x: 3.5, y: 3.7, w: 3.0, h: 0.04, fill: { color: WHITE, transparency: 60 }, line: { color: WHITE, transparency: 60 } });
  s8.addText(DATA.WEB || "www.knownonline.com", { x: 0.5, y: 3.9, w: 9, h: 0.35, fontSize: 14, color: WHITE, fontFace: "DM Sans", align: "center", bold: true });
  s8.addText(DATA.CONTACTO || "ariel@knownonline.com", { x: 0.5, y: 4.3, w: 9, h: 0.3, fontSize: 12, color: "FFD4B8", fontFace: "DM Sans", align: "center" });

  // ── Generate base64 ───────────────────────────────────────────────────────
  return pres.write({ outputType: "base64" });
}
