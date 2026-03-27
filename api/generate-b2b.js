const pptxgen = require("pptxgenjs");
const {
  COLORS,
  getUsdRate,
  normalizeDataForUSD,
  buildSlide_Cover,
  buildSlide_Recommendations,
  buildSlide_Close,
} = require("./lib/pptx-shared");

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
  const labelCortoActual   = DATA.PERIODO_ACTUAL_SHORT   || `${MESES[mesIdx]} '${String(añoActual).slice(2)}`;
  const labelCortoAnterior = DATA.PERIODO_ANTERIOR_SHORT || `${MESES[mesIdx]} '${String(añoActual - 1).slice(2)}`;

  // ── Brand colors ──────────────────────────────────────────────────────────
  const { ORANGE, ORANGE2, DARK, WHITE, LIGHT_BG, GRAY_TEXT, LIGHT_GRAY,
          GREEN, GREEN_BG, RED, RED_BG, AMBER, AMBER_BG, BLUE, LIGHT_BLUE } = COLORS;

  // ── Helpers ───────────────────────────────────────────────────────────────
  const parseNum = str => {
    const c = (str || "0").replace(/\./g, "").replace(",", ".").replace(/[^0-9.]/g, "");
    return parseFloat(c) || 0;
  };
  const hasGoogle = parseNum(DATA.GOOGLE_COSTO) > 0;
  const fmtMoneyCompact = val => {
    const n = parseNum(val);
    if (n >= 1_000_000) return `$${(n / 1_000_000).toFixed(2).replace(".", ",")} M`;
    if (n >= 1_000)     return `$${(n / 1_000).toFixed(1).replace(".", ",")} K`;
    return val || "";
  };

  // ── SLIDE 1 – COVER ───────────────────────────────────────────────────────
  buildSlide_Cover(pres, DATA);

  // ── SLIDE 2 – RESUMEN EJECUTIVO ───────────────────────────────────────────
  let s2 = pres.addSlide();
  s2.background = { color: WHITE };
  s2.addText("Resumen Ejecutivo", { x: 0.5, y: 0.22, w: 7, h: 0.55, fontSize: 28, bold: true, color: DARK, fontFace: "Trebuchet MS" });
  s2.addText(`Inversión · Leads · CPL  ·  Meta Ads${hasGoogle ? " + Google Ads" : ""}`, { x: 0.5, y: 0.78, w: 7, h: 0.3, fontSize: 13, color: GRAY_TEXT, fontFace: "DM Sans" });

  const kpis = [
    { label: "Inversión total", val: fmtMoneyCompact(DATA.INVERSION_TOTAL), delta: DATA.INVERSION_DELTA   || "", note: `${DATA.PERIODO_ANTERIOR_LABEL || "Año ant."}: ${DATA.INVERSION_PREV   || ""}`, up: DATA.INVERSION_DELTA_UP   === true },
    { label: "Leads totales",   val: DATA.LEADS_TOTAL  || "",               delta: DATA.LEADS_DELTA       || "", note: `${DATA.PERIODO_ANTERIOR_LABEL || "Año ant."}: ${DATA.LEADS_PREV       || ""}`, up: DATA.LEADS_DELTA_UP       === true },
    { label: "CPL promedio",    val: DATA.CPL_TOTAL    || "",               delta: DATA.CPL_DELTA         || "", note: `${DATA.PERIODO_ANTERIOR_LABEL || "Año ant."}: ${DATA.CPL_PREV         || ""}`, up: DATA.CPL_DELTA_UP         === true },
    { label: "Clicks (todos)",   val: DATA.CLICKS_TOTAL || "",               delta: DATA.CLICKS_DELTA      || "", note: `${DATA.PERIODO_ANTERIOR_LABEL || "Año ant."}: ${DATA.CLICKS_PREV      || ""}`, up: DATA.CLICKS_DELTA_UP      === true },
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
    ["CPL",    DATA.META_CPL    || "", DATA.META_CPL_DELTA    || "", DATA.META_CPL_DELTA_UP    === true],
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

  // Google block (condicional)
  if (hasGoogle) {
    s2.addShape(pres.shapes.RECTANGLE, { x: 5.2, y: 3.35, w: 4.4, h: 1.85, fill: { color: LIGHT_BLUE }, line: { color: "D0E4F5", width: 0.5 } });
    s2.addShape(pres.shapes.RECTANGLE, { x: 5.2, y: 3.35, w: 4.4, h: 0.38, fill: { color: BLUE }, line: { color: BLUE } });
    s2.addText("Google Ads", { x: 5.35, y: 3.38, w: 3, h: 0.32, fontSize: 13, bold: true, color: WHITE, fontFace: "DM Sans" });
    const googleStats = [
      ["Costo",  DATA.GOOGLE_COSTO  || "", DATA.GOOGLE_COSTO_DELTA  || "", DATA.GOOGLE_COSTO_DELTA_UP  === true],
      ["Clicks", DATA.GOOGLE_CLICKS || "", DATA.GOOGLE_CLICKS_DELTA || "", DATA.GOOGLE_CLICKS_DELTA_UP !== true],
      ["CPL",    DATA.GOOGLE_CPL    || "", DATA.GOOGLE_CPL_DELTA    || "", DATA.GOOGLE_CPL_DELTA_UP    === true],
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
  }

  // ── SLIDE – FACTURACIÓN & ROAS (CONDICIONAL VTEX) ────────────────────────
  if (DATA.ECOMMERCE_INGRESOS) {
    const _inv  = parseNum(DATA.INVERSION_TOTAL);
    const _invP = parseNum(DATA.INVERSION_PREV);
    const _rev  = parseNum(DATA.ECOMMERCE_INGRESOS);
    const _revP = parseNum(DATA.ECOMMERCE_INGRESOS_PREV);
    const _roas  = _inv  > 0 ? _rev  / _inv  : 0;
    const _roasP = _invP > 0 ? _revP / _invP : 0;
    const _roasDelta = _roasP > 0 ? ((_roas - _roasP) / _roasP * 100) : 0;
    const _roasDeltaStr = (_roasDelta >= 0 ? "+" : "") + _roasDelta.toFixed(1).replace(".", ",") + "%";
    const _roasStr  = _roas.toFixed(2).replace(".", ",")  + "x";
    const _roasPStr = _roasP.toFixed(2).replace(".", ",") + "x";

    let sVtex = pres.addSlide();
    sVtex.background = { color: WHITE };
    sVtex.addText("Facturación & ROAS", { x: 0.5, y: 0.2, w: 8, h: 0.55, fontSize: 28, bold: true, color: DARK, fontFace: "Trebuchet MS" });
    sVtex.addText(`Ecommerce  ·  ${DATA.PERIODO_ACTUAL_LABEL || ""} vs ${DATA.PERIODO_ANTERIOR_LABEL || ""}`, { x: 0.5, y: 0.76, w: 8, h: 0.3, fontSize: 13, color: GRAY_TEXT, fontFace: "DM Sans" });

    const vtexKPIs = [
      { label: "Facturación",  val: fmtMoneyCompact(DATA.ECOMMERCE_INGRESOS), prev: DATA.ECOMMERCE_INGRESOS_PREV || "", delta: DATA.ECOMMERCE_INGRESOS_DELTA || "", up: DATA.ECOMMERCE_INGRESOS_DELTA_UP === true },
      { label: "ROAS",         val: _roasStr,                                  prev: _roasPStr,                        delta: _roasDeltaStr,                      up: _roas >= _roasP                          },
      { label: "Inversión",    val: fmtMoneyCompact(DATA.INVERSION_TOTAL),     prev: DATA.INVERSION_PREV           || "", delta: DATA.INVERSION_DELTA           || "", up: DATA.INVERSION_DELTA_UP           === true },
    ];

    vtexKPIs.forEach((k, i) => {
      const x = 0.85 + i * 2.8, y = 1.3;
      sVtex.addShape(pres.shapes.RECTANGLE, { x, y, w: 2.6, h: 3.2, fill: { color: LIGHT_BG }, line: { color: "F0E8E0", width: 0.5 } });
      sVtex.addShape(pres.shapes.RECTANGLE, { x, y, w: 2.6, h: 0.06, fill: { color: ORANGE }, line: { color: ORANGE } });
      sVtex.addShape(pres.shapes.OVAL, { x: x + 0.14, y: y + 0.2, w: 0.45, h: 0.45, fill: { color: ORANGE }, line: { color: ORANGE } });
      sVtex.addText(k.label, { x: x + 0.7, y: y + 0.22, w: 1.8, h: 0.28, fontSize: 12, bold: true, color: DARK, fontFace: "DM Sans" });
      sVtex.addShape(pres.shapes.RECTANGLE, { x: x + 0.14, y: y + 0.82, w: 2.32, h: 0.02, fill: { color: "E8E0D8" }, line: { color: "E8E0D8" } });
      sVtex.addText(DATA.PERIODO_ACTUAL_LABEL || "", { x: x + 0.14, y: y + 0.92, w: 2.0, h: 0.2, fontSize: 9, color: GRAY_TEXT, fontFace: "DM Sans" });
      const fs = String(k.val).length > 12 ? 18 : String(k.val).length > 9 ? 22 : 28;
      sVtex.addText(k.val, { x: x + 0.14, y: y + 1.12, w: 2.32, h: 0.65, fontSize: fs, bold: true, color: DARK, fontFace: "Trebuchet MS" });
      sVtex.addShape(pres.shapes.RECTANGLE, { x: x + 0.14, y: y + 1.82, w: 2.32, h: 0.02, fill: { color: "E8E0D8" }, line: { color: "E8E0D8" } });
      sVtex.addText(`${DATA.PERIODO_ANTERIOR_LABEL || "Período ant."}: ${k.prev}`, { x: x + 0.14, y: y + 1.92, w: 2.32, h: 0.2, fontSize: 9, color: GRAY_TEXT, fontFace: "DM Sans" });
      sVtex.addShape(pres.shapes.RECTANGLE, { x: x + 0.5, y: y + 2.25, w: 1.6, h: 0.38, fill: { color: k.up ? GREEN_BG : RED_BG }, line: { color: k.up ? GREEN_BG : RED_BG } });
      sVtex.addText(k.delta, { x: x + 0.5, y: y + 2.25, w: 1.6, h: 0.38, fontSize: 16, bold: true, color: k.up ? GREEN : RED, fontFace: "DM Sans", align: "center", valign: "middle" });
    });
  }

  // ── SLIDE – CAMPAÑAS META ADS (tabla por campaña) ────────────────────────
  if (Array.isArray(DATA.META_CAMPANAS) && DATA.META_CAMPANAS.length > 0) {
    let smc = pres.addSlide();
    smc.background = { color: WHITE };

    // Header
    smc.addText("Campañas ", { x: 1.0, y: 0.15, w: 8.5, h: 0.6, fontSize: 28, bold: true, color: DARK, fontFace: "Trebuchet MS",
      paraSpaceAfter: 0,
      objects: [
        { text: "Campañas ", options: { bold: true, color: DARK } },
        { text: `Facebook Ads – ${DATA.PERIODO_ACTUAL_LABEL || ""}`, options: { bold: true, color: ORANGE } },
      ]
    });
    // Rewrite with two text runs
    smc.addText([
      { text: "Campañas ", options: { bold: true, color: DARK, fontSize: 26, fontFace: "Trebuchet MS" } },
      { text: `Facebook Ads – ${DATA.PERIODO_ACTUAL_LABEL || ""}`, options: { bold: true, color: ORANGE, fontSize: 26, fontFace: "Trebuchet MS" } },
    ], { x: 1.0, y: 0.15, w: 8.8, h: 0.6 });

    // Orange circle icon placeholder
    smc.addShape(pres.shapes.OVAL, { x: 0.15, y: 0.1, w: 0.72, h: 0.72, fill: { color: ORANGE }, line: { color: ORANGE } });
    smc.addText("f", { x: 0.15, y: 0.1, w: 0.72, h: 0.72, fontSize: 18, bold: true, color: WHITE, fontFace: "DM Sans", align: "center", valign: "middle" });

    // Table columns: Conjunto de anuncios | Leads | CPL | Impresiones | Clicks | CPC | Inversión | Alcance
    const mcColW  = [2.85, 0.72, 0.82, 1.08, 0.75, 0.82, 1.0, 0.88];
    const mcHdrs  = ["Conjunto de anuncios", "Leads", "CPL", "Impresiones", "Clicks", "CPC", "Inversión", "Alcance"];
    const mcY0    = 0.88;

    smc.addShape(pres.shapes.RECTANGLE, { x: 0.18, y: mcY0, w: 9.65, h: 0.34, fill: { color: "F5F5F5" }, line: { color: "E0E0E0", width: 0.5 } });
    let mcCx = 0.28;
    mcHdrs.forEach((h, i) => {
      const align = i === 0 ? "left" : "center";
      smc.addText(h, { x: mcCx, y: mcY0 + 0.02, w: mcColW[i], h: 0.3, fontSize: 8.5, bold: true, color: GRAY_TEXT, fontFace: "DM Sans", valign: "middle", align });
      mcCx += mcColW[i];
    });

    // Rows
    const maxRows = Math.min(DATA.META_CAMPANAS.length, 10);
    DATA.META_CAMPANAS.slice(0, maxRows).forEach((row, i) => {
      const ry  = mcY0 + 0.34 + i * 0.33;
      const bg  = i % 2 === 0 ? WHITE : "FAFAFA";
      smc.addShape(pres.shapes.RECTANGLE, { x: 0.18, y: ry, w: 9.65, h: 0.33, fill: { color: bg }, line: { color: "EEEEEE", width: 0.3 } });

      let rx = 0.28;
      const cells = [
        { val: row.nombre   || "", align: "left",   color: ORANGE, bold: false },
        { val: row.leads    || "", align: "center",  color: DARK,   bold: false },
        { val: row.cpl      || "", align: "center",  color: DARK,   bold: false },
        { val: row.impresiones || "", align: "center", color: DARK, bold: false },
        { val: row.clicks   || "", align: "center",  color: DARK,   bold: false },
        { val: row.cpc      || "", align: "center",  color: DARK,   bold: false },
        { val: row.costo    || "", align: "center",  color: DARK,   bold: false },
        { val: row.alcance  || "", align: "center",  color: DARK,   bold: false },
      ];
      cells.forEach((c, ci) => {
        const nombre = c.val.length > 38 ? c.val.substring(0, 36) + "…" : c.val;
        smc.addText(nombre, { x: rx, y: ry + 0.05, w: mcColW[ci], h: 0.24, fontSize: 8.5, color: c.color, fontFace: "DM Sans", align: c.align, valign: "middle" });
        rx += mcColW[ci];
      });
    });

    // Bottom KPI cards
    const kpiY    = 4.42;
    const kpiData = [
      { label: "Inversión Total",  val: DATA.META_COSTO       || "" },
      { label: "Coste por Lead",   val: DATA.META_CPL         || "" },
      { label: "Impresiones",      val: DATA.META_IMPRESIONES || "" },
      { label: "Leads Totales",    val: DATA.META_LEADS       || "" },
    ];
    kpiData.forEach((k, i) => {
      const kx = 0.18 + i * 2.44;
      smc.addShape(pres.shapes.RECTANGLE, { x: kx, y: kpiY, w: 2.3, h: 0.95, fill: { color: WHITE }, line: { color: "E8E8E8", width: 0.8 } });
      smc.addShape(pres.shapes.OVAL, { x: kx + 0.12, y: kpiY + 0.18, w: 0.48, h: 0.48, fill: { color: ORANGE }, line: { color: ORANGE } });
      smc.addText(k.label, { x: kx + 0.7, y: kpiY + 0.1,  w: 1.52, h: 0.28, fontSize: 9,  bold: true,  color: DARK,   fontFace: "DM Sans" });
      smc.addText(k.val,   { x: kx + 0.7, y: kpiY + 0.36, w: 1.52, h: 0.32, fontSize: 14, bold: true,  color: ORANGE, fontFace: "DM Sans" });
    });

    // Footer
    smc.addText(`Reporte ${DATA.CLIENTE_NOMBRE || ""} | ${DATA.AGENCIA_NOMBRE || "Known Online"}`,
      { x: 0.18, y: 5.48, w: 6, h: 0.22, fontSize: 8.5, color: GRAY_TEXT, fontFace: "DM Sans" });
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
    { label: "Leads",       val: DATA.META_LEADS        || "", prev: DATA.META_LEADS_PREV        || "", delta: DATA.META_LEADS_DELTA        || "", up: DATA.META_LEADS_DELTA_UP        === true, warn: false },
    { label: "CPL",         val: DATA.META_CPL          || "", prev: DATA.META_CPL_PREV          || "", delta: DATA.META_CPL_DELTA          || "", up: DATA.META_CPL_DELTA_UP          === true, warn: false },
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


  // ── SLIDE 4 – GOOGLE ADS DETALLE (condicional) ───────────────────────────
  if (hasGoogle) {
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
      { label: "Leads",       val: DATA.GOOGLE_LEADS       || "", prev: DATA.GOOGLE_LEADS_PREV       || "", delta: DATA.GOOGLE_LEADS_DELTA       || "", good: DATA.GOOGLE_LEADS_DELTA_UP       === true },
      { label: "CPL",         val: DATA.GOOGLE_CPL         || "", prev: DATA.GOOGLE_CPL_PREV         || "", delta: DATA.GOOGLE_CPL_DELTA         || "", good: DATA.GOOGLE_CPL_DELTA_UP         === true },
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
  }


  // ── SLIDE 5 – TOP CAMPAÑAS POR CPL ───────────────────────────────────────
  let s5 = pres.addSlide();
  s5.background = { color: WHITE };
  s5.addText("Top Campañas por CPL", { x: 0.5, y: 0.2, w: 7, h: 0.55, fontSize: 28, bold: true, color: DARK, fontFace: "Trebuchet MS" });
  s5.addText(`${DATA.PERIODO_ACTUAL_LABEL || ""}  ·  ${hasGoogle ? "Google Ads + Meta Ads" : "Meta Ads"}`, { x: 0.5, y: 0.76, w: 7, h: 0.3, fontSize: 13, color: GRAY_TEXT, fontFace: "DM Sans" });

  // campaigns: array of { nombre, plataforma, costo, leads, cpl, nivel }
  // nivel: "low" (CPL bajo = bueno) | "mid" | "high" (CPL alto = revisar)
  const campaigns = DATA.CAMPANAS || [];
  s5.addShape(pres.shapes.RECTANGLE, { x: 0.4, y: 1.2, w: 9.2, h: 0.38, fill: { color: DARK }, line: { color: DARK } });
  const headers = ["Campaña", "Plat.", "Inversión", "Leads", "CPL"];
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
    s5.addText(row.leads || "", { x: rx, y: y + 0.07, w: colW[3], h: 0.3, fontSize: 10, color: DARK, fontFace: "DM Sans", align: "right" });
    rx += colW[3];

    // nivel para CPL: "low" = bueno (verde), "mid" = amber, "high" = malo (rojo)
    const nivel   = row.nivel || "mid";
    const cplColor = nivel === "low" ? GREEN : nivel === "mid" ? AMBER : RED;
    const cplBg    = nivel === "low" ? GREEN_BG : nivel === "mid" ? AMBER_BG : RED_BG;
    s5.addShape(pres.shapes.RECTANGLE, { x: rx, y: y + 0.1, w: 1.0, h: 0.24, fill: { color: cplBg }, line: { color: cplBg } });
    s5.addText(row.cpl || "", { x: rx, y: y + 0.1, w: 1.0, h: 0.24, fontSize: 10, bold: true, color: cplColor, fontFace: "DM Sans", align: "center" });
  });


  // ── SLIDE – DISTRIBUCIÓN POR FUNNEL (condicional, parsea CAMPANAS) ────────
  {
    const nivelMap = {};
    (DATA.CAMPANAS || []).filter(c => c.plataforma === "Meta").forEach(c => {
      const parts = (c.nombre || "").split("+");
      const raw   = (parts[1] || "").trim();
      if (!raw) return;
      const nivel = raw.toLowerCase() === "lower" ? "Lower Funnel"
                  : raw.toLowerCase() === "middle" ? "Mid Funnel"
                  : raw.toLowerCase() === "upper"  ? "Upper Funnel"
                  : raw;
      if (!nivelMap[nivel]) nivelMap[nivel] = { leads: 0, costo: 0 };
      nivelMap[nivel].leads += parseNum(c.leads);
      nivelMap[nivel].costo += parseNum(c.costo);
    });
    const niveles = Object.entries(nivelMap)
      .map(([nombre, d]) => ({ nombre, leads: d.leads, costo: d.costo, cpl: d.leads > 0 ? d.costo / d.leads : 0 }))
      .sort((a, b) => b.leads - a.leads)
      .slice(0, 4);
    const totalLeadsNivel = niveles.reduce((s, n) => s + n.leads, 0);

    if (niveles.length >= 2) {
      const nivelColors = n => {
        const l = n.toLowerCase();
        if (l.includes("lower")) return { fill: GREEN_BG,   header: GREEN,  text: GREEN  };
        if (l.includes("mid"))   return { fill: AMBER_BG,   header: AMBER,  text: AMBER  };
        if (l.includes("upper")) return { fill: LIGHT_BLUE, header: BLUE,   text: BLUE   };
        return                          { fill: LIGHT_BG,   header: ORANGE, text: ORANGE };
      };
      const fmtMoney = n => {
        if (n >= 1_000_000) return `$${(n/1_000_000).toFixed(1).replace(".",",")} M`;
        if (n >= 1_000)     return `$${(n/1_000).toFixed(0).replace(".",",")} K`;
        return `$${n.toFixed(0)}`;
      };
      const fmtLeads = n => new Intl.NumberFormat("es-AR", { maximumFractionDigits: 0 }).format(n);

      let sFunnel = pres.addSlide();
      sFunnel.background = { color: WHITE };
      sFunnel.addText("Distribución por Funnel", { x: 0.5, y: 0.22, w: 7, h: 0.55, fontSize: 28, bold: true, color: DARK, fontFace: "Trebuchet MS" });
      sFunnel.addText(`Meta Ads  ·  ${DATA.PERIODO_ACTUAL_LABEL || ""}`, { x: 0.5, y: 0.78, w: 7, h: 0.3, fontSize: 13, color: GRAY_TEXT, fontFace: "DM Sans" });

      const cardW   = niveles.length === 2 ? 4.2 : 2.85;
      const cardGap = 0.2;
      const cardH   = 3.0;
      const cardY   = 1.1;
      const totalCW = niveles.length * cardW + (niveles.length - 1) * cardGap;
      const cardX0  = (10 - totalCW) / 2;

      niveles.forEach((n, i) => {
        const cx     = cardX0 + i * (cardW + cardGap);
        const colors = nivelColors(n.nombre);
        const pctStr = totalLeadsNivel > 0 ? ((n.leads / totalLeadsNivel) * 100).toFixed(1) + "%" : "0%";

        sFunnel.addShape(pres.shapes.RECTANGLE, { x: cx, y: cardY, w: cardW, h: cardH, fill: { color: colors.fill }, line: { color: colors.header, width: 0.5 } });
        sFunnel.addShape(pres.shapes.RECTANGLE, { x: cx, y: cardY, w: cardW, h: 0.07, fill: { color: colors.header }, line: { color: colors.header } });

        sFunnel.addText(n.nombre, { x: cx + 0.15, y: cardY + 0.14, w: cardW - 0.3, h: 0.3, fontSize: 12, bold: true, color: colors.text, fontFace: "DM Sans" });
        sFunnel.addText(fmtLeads(n.leads), { x: cx + 0.15, y: cardY + 0.48, w: cardW - 0.3, h: 0.65, fontSize: 32, bold: true, color: DARK, fontFace: "Trebuchet MS", align: "center" });
        sFunnel.addText("leads", { x: cx + 0.15, y: cardY + 1.1, w: cardW - 0.3, h: 0.22, fontSize: 10, color: GRAY_TEXT, fontFace: "DM Sans", align: "center" });

        const badgeX = cx + (cardW - 0.95) / 2;
        sFunnel.addShape(pres.shapes.RECTANGLE, { x: badgeX, y: cardY + 1.38, w: 0.95, h: 0.3, fill: { color: colors.header }, line: { color: colors.header } });
        sFunnel.addText(pctStr, { x: badgeX, y: cardY + 1.38, w: 0.95, h: 0.3, fontSize: 12, bold: true, color: WHITE, fontFace: "DM Sans", align: "center", valign: "middle" });

        sFunnel.addShape(pres.shapes.RECTANGLE, { x: cx + 0.15, y: cardY + 1.82, w: cardW - 0.3, h: 0.02, fill: { color: colors.header, transparency: 55 }, line: { color: colors.header, transparency: 55 } });

        sFunnel.addText("CPL Promedio", { x: cx + 0.15, y: cardY + 1.92, w: cardW - 0.3, h: 0.22, fontSize: 9, color: GRAY_TEXT, fontFace: "DM Sans" });
        sFunnel.addText(n.cpl > 0 ? fmtMoney(n.cpl) : "N/D", { x: cx + 0.15, y: cardY + 2.12, w: cardW - 0.3, h: 0.3, fontSize: 15, bold: true, color: DARK, fontFace: "DM Sans" });

        sFunnel.addText("Inversión", { x: cx + 0.15, y: cardY + 2.48, w: cardW - 0.3, h: 0.22, fontSize: 9, color: GRAY_TEXT, fontFace: "DM Sans" });
        sFunnel.addText(fmtMoney(n.costo), { x: cx + 0.15, y: cardY + 2.68, w: cardW - 0.3, h: 0.26, fontSize: 13, bold: true, color: DARK, fontFace: "DM Sans" });
      });

      // Barra proporcional
      const barY = cardY + cardH + 0.18;
      let barCx = cardX0;
      niveles.forEach(n => {
        const colors = nivelColors(n.nombre);
        const ratio  = totalLeadsNivel > 0 ? n.leads / totalLeadsNivel : 1 / niveles.length;
        const segW   = totalCW * ratio;
        sFunnel.addShape(pres.shapes.RECTANGLE, { x: barCx, y: barY, w: segW, h: 0.42, fill: { color: colors.header }, line: { color: colors.header } });
        if (segW > 0.6) {
          sFunnel.addText(`${n.nombre}  ${(ratio * 100).toFixed(0)}%`, { x: barCx, y: barY, w: segW, h: 0.42, fontSize: 8.5, bold: true, color: WHITE, fontFace: "DM Sans", align: "center", valign: "middle" });
        }
        barCx += segW;
      });
    }
  }

  // ── SLIDE – CONJUNTOS DE ANUNCIOS (condicional) ───────────────────────────
  if (Array.isArray(DATA.META_ADSETS) && DATA.META_ADSETS.length > 0) {
    let sAs = pres.addSlide();
    sAs.background = { color: WHITE };

    sAs.addText([
      { text: "Conjuntos de Anuncios ", options: { bold: true, color: DARK,   fontSize: 26, fontFace: "Trebuchet MS" } },
      { text: `Facebook Ads – ${DATA.PERIODO_ACTUAL_LABEL || ""}`, options: { bold: true, color: ORANGE, fontSize: 26, fontFace: "Trebuchet MS" } },
    ], { x: 1.0, y: 0.15, w: 8.8, h: 0.6 });
    sAs.addShape(pres.shapes.OVAL, { x: 0.15, y: 0.1, w: 0.72, h: 0.72, fill: { color: ORANGE }, line: { color: ORANGE } });
    sAs.addText("f", { x: 0.15, y: 0.1, w: 0.72, h: 0.72, fontSize: 18, bold: true, color: WHITE, fontFace: "DM Sans", align: "center", valign: "middle" });

    const asColW = [3.2, 0.85, 1.0, 1.25, 0.9, 0.88, 0.94];
    const asHdrs = ["Conjunto de anuncios", "Leads", "CPL", "Inversión", "Clicks", "CPC", "Alcance"];
    const asY0   = 0.88;
    const asW    = asColW.reduce((s, w) => s + w, 0);

    sAs.addShape(pres.shapes.RECTANGLE, { x: 0.18, y: asY0, w: asW, h: 0.34, fill: { color: "F5F5F5" }, line: { color: "E0E0E0", width: 0.5 } });
    let asCx = 0.28;
    asHdrs.forEach((h, i) => {
      sAs.addText(h, { x: asCx, y: asY0 + 0.02, w: asColW[i], h: 0.3, fontSize: 8.5, bold: true, color: GRAY_TEXT, fontFace: "DM Sans", valign: "middle", align: i === 0 ? "left" : "center" });
      asCx += asColW[i];
    });

    DATA.META_ADSETS.slice(0, 10).forEach((row, i) => {
      const ry = asY0 + 0.34 + i * 0.33;
      sAs.addShape(pres.shapes.RECTANGLE, { x: 0.18, y: ry, w: asW, h: 0.33, fill: { color: i % 2 === 0 ? WHITE : "FAFAFA" }, line: { color: "EEEEEE", width: 0.3 } });
      let rx = 0.28;
      const cells = [
        { val: row.nombre   || "", align: "left",   color: ORANGE },
        { val: row.leads    || "", align: "center",  color: DARK   },
        { val: row.cpl      || "", align: "center",  color: DARK   },
        { val: row.costo    || "", align: "center",  color: DARK   },
        { val: row.clicks   || "", align: "center",  color: DARK   },
        { val: row.cpc      || "", align: "center",  color: DARK   },
        { val: row.alcance  || "", align: "center",  color: DARK   },
      ];
      cells.forEach((c, ci) => {
        const txt = c.val.length > 42 ? c.val.substring(0, 40) + "…" : c.val;
        sAs.addText(txt, { x: rx, y: ry + 0.05, w: asColW[ci], h: 0.24, fontSize: 8.5, color: c.color, fontFace: "DM Sans", align: c.align, valign: "middle" });
        rx += asColW[ci];
      });
    });

    const asKpiY = 4.42;
    [
      { label: "Conjuntos activos", val: String(DATA.META_ADSETS.length) },
      { label: "Leads totales",     val: DATA.META_LEADS  || "" },
      { label: "CPL promedio",      val: DATA.META_CPL    || "" },
      { label: "Inversión total",   val: DATA.META_COSTO  || "" },
    ].forEach((k, i) => {
      const kx = 0.18 + i * 2.44;
      sAs.addShape(pres.shapes.RECTANGLE, { x: kx, y: asKpiY, w: 2.3, h: 0.95, fill: { color: WHITE }, line: { color: "E8E8E8", width: 0.8 } });
      sAs.addShape(pres.shapes.OVAL, { x: kx + 0.12, y: asKpiY + 0.18, w: 0.48, h: 0.48, fill: { color: ORANGE }, line: { color: ORANGE } });
      sAs.addText(k.label, { x: kx + 0.7, y: asKpiY + 0.1,  w: 1.52, h: 0.28, fontSize: 9,  bold: true, color: DARK,   fontFace: "DM Sans" });
      sAs.addText(k.val,   { x: kx + 0.7, y: asKpiY + 0.36, w: 1.52, h: 0.32, fontSize: 14, bold: true, color: ORANGE, fontFace: "DM Sans" });
    });

    sAs.addText(`Reporte ${DATA.CLIENTE_NOMBRE || ""} | ${DATA.AGENCIA_NOMBRE || "Known Online"}`,
      { x: 0.18, y: 5.48, w: 6, h: 0.22, fontSize: 8.5, color: GRAY_TEXT, fontFace: "DM Sans" });
  }

  // ── SLIDE 5C – TOP ANUNCIOS META POR LEADS ───────────────────────────────
  if (DATA.TOP_ANUNCIOS_META_TIENE_DATOS && Array.isArray(DATA.TOP_ANUNCIOS_META) && DATA.TOP_ANUNCIOS_META.length > 0) {
    const ads = DATA.TOP_ANUNCIOS_META.slice(0, 3);

    let s5b = pres.addSlide();
    s5b.background = { color: WHITE };
    s5b.addText("Top Anuncios Meta por Leads", { x: 0.5, y: 0.2, w: 7, h: 0.55, fontSize: 28, bold: true, color: DARK, fontFace: "Trebuchet MS" });
    s5b.addText(`${DATA.PERIODO_ACTUAL_LABEL || ""}  ·  Ordenados por CPL`, { x: 0.5, y: 0.76, w: 7, h: 0.3, fontSize: 13, color: GRAY_TEXT, fontFace: "DM Sans" });

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
        { lbl: "Leads",  val: ad.leads || "0" },
        { lbl: "CPL",    val: ad.cpl   || "" },
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
  buildSlide_Recommendations(pres, DATA);

  // ── SLIDE 8 – CIERRE ──────────────────────────────────────────────────────
  buildSlide_Close(pres, DATA);

  // ── Generate base64 ───────────────────────────────────────────────────────
  return pres.write({ outputType: "base64" });
}