const pptxgen = require("pptxgenjs");
const {
  COLORS,
  buildSlide_Cover,
  buildSlide_Recommendations,
  buildSlide_Close,
} = require("./lib/pptx-shared");

module.exports = async function handler(req, res) {
  if (req.method !== "POST") {
    return res.status(405).json({ error: "Method not allowed. Use POST." });
  }
  const { DATA } = req.body || {};
  if (!DATA) return res.status(400).json({ error: "Missing DATA in request body." });
  try {
    const base64 = await generatePptx(DATA);
    const filename = `Reporte_Anura_${DATA.PERIODO_ACTUAL_LABEL || "Periodo"}.pptx`.replace(/\s+/g, "_");
    return res.status(200).json({ pptx: base64, filename });
  } catch (err) {
    console.error(err);
    return res.status(500).json({ error: err.message || "Error generating PPTX." });
  }
};

async function generatePptx(DATA) {
  const pres = new pptxgen();
  pres.layout = "LAYOUT_16x9";
  pres.title  = `Reporte Anura – ${DATA.PERIODO_ACTUAL_LABEL || ""}`;

  const { ORANGE, DARK, WHITE, LIGHT_BG, GRAY_TEXT,
          GREEN, GREEN_BG, RED, RED_BG, AMBER, AMBER_BG, BLUE, LIGHT_BLUE } = COLORS;

  // ── Helpers ───────────────────────────────────────────────────────────────
  const parseNum = str => {
    if (typeof str === "number") return str;
    return parseFloat((str || "0").replace(/\./g, "").replace(",", ".").replace(/[^0-9.]/g, "")) || 0;
  };
  const fmtMoneyCompact = val => {
    const n = parseNum(val);
    if (n >= 1_000_000) return `$${(n / 1_000_000).toFixed(2).replace(".", ",")} M`;
    if (n >= 1_000)     return `$${(n / 1_000).toFixed(1).replace(".", ",")} K`;
    return String(val || "");
  };
  const fmtDelta = str => (str || "").replace(/pp$/i, "%");
  const hasGoogle = parseNum(DATA.GOOGLE_COSTO) > 0;
  const periodoAnt = DATA.PERIODO_ANTERIOR_LABEL || "Período ant.";

  // ── SLIDE 1 – COVER ───────────────────────────────────────────────────────
  buildSlide_Cover(pres, DATA);

  // ── SLIDE 2 – RESUMEN EJECUTIVO ───────────────────────────────────────────
  {
    const s = pres.addSlide();
    s.background = { color: WHITE };
    s.addText("Resumen Ejecutivo", { x: 0.5, y: 0.22, w: 7, h: 0.55, fontSize: 28, bold: true, color: DARK, fontFace: "DM Sans" });
    s.addText(`Leads Zoho · Inversión · CPL  ·  ${DATA.PERIODO_ACTUAL_LABEL || ""} vs ${periodoAnt}`, { x: 0.5, y: 0.78, w: 9, h: 0.3, fontSize: 13, color: GRAY_TEXT, fontFace: "DM Sans" });

    const kpis = [
      { label: "Leads totales",  val: String(parseNum(DATA.ZOHO_LEADS_TOTAL) || DATA.ZOHO_LEADS_TOTAL || ""), delta: fmtDelta(DATA.ZOHO_LEADS_DELTA), note: `${periodoAnt}: ${DATA.ZOHO_LEADS_PREV || ""}`, up: DATA.ZOHO_LEADS_DELTA_UP === true },
      { label: "Inversión total", val: fmtMoneyCompact(DATA.INVERSION_TOTAL), delta: fmtDelta(DATA.INVERSION_DELTA), note: `${periodoAnt}: ${fmtMoneyCompact(DATA.INVERSION_PREV)}`, up: DATA.INVERSION_DELTA_UP === true },
      { label: "CPL promedio",   val: fmtMoneyCompact(DATA.CPL_TOTAL),        delta: fmtDelta(DATA.CPL_DELTA),       note: `${periodoAnt}: ${fmtMoneyCompact(DATA.CPL_PREV)}`,        up: DATA.CPL_DELTA_UP === true },
      { label: "Clicks totales", val: DATA.CLICKS_TOTAL || "",                delta: fmtDelta(DATA.CLICKS_DELTA),    note: `${periodoAnt}: ${DATA.CLICKS_PREV || ""}`,    up: DATA.CLICKS_DELTA_UP === true },
    ];
    kpis.forEach((k, i) => {
      const x = 0.4 + i * 2.32;
      s.addShape(pres.shapes.RECTANGLE, { x, y: 1.2, w: 2.1, h: 1.55, fill: { color: LIGHT_BG }, line: { color: "F0E8E0", width: 0.5 } });
      s.addShape(pres.shapes.RECTANGLE, { x, y: 1.2, w: 2.1, h: 0.07, fill: { color: ORANGE }, line: { color: ORANGE } });
      s.addText(k.label, { x, y: 1.32, w: 2.1, h: 0.3,  fontSize: 10, color: GRAY_TEXT, fontFace: "DM Sans", align: "center" });
      s.addText(k.val,   { x, y: 1.62, w: 2.1, h: 0.52, fontSize: 26, bold: true, color: DARK, fontFace: "DM Sans", align: "center" });
      s.addShape(pres.shapes.RECTANGLE, { x: x + 0.6, y: 2.17, w: 0.9, h: 0.27, fill: { color: k.up ? GREEN_BG : RED_BG }, line: { color: k.up ? GREEN_BG : RED_BG } });
      s.addText(k.delta, { x: x + 0.6, y: 2.17, w: 0.9, h: 0.27, fontSize: 11, bold: true, color: k.up ? GREEN : RED, fontFace: "DM Sans", align: "center" });
      s.addText(k.note,  { x, y: 2.5,  w: 2.1, h: 0.25, fontSize: 9, color: GRAY_TEXT, fontFace: "DM Sans", align: "center" });
    });

    s.addText("* Leads provenientes de Zoho CRM — fuente de verdad (deduplicados por email)", { x: 0.4, y: 2.82, w: 9.2, h: 0.2, fontSize: 8, color: GRAY_TEXT, fontFace: "DM Sans", italic: true });
    s.addText("Comparativa por plataforma", { x: 0.5, y: 3.08, w: 9, h: 0.35, fontSize: 13, bold: true, color: DARK, fontFace: "DM Sans" });

    // Meta block
    const metaW = hasGoogle ? 4.4 : 9.2;
    s.addShape(pres.shapes.RECTANGLE, { x: 0.4, y: 3.48, w: metaW, h: 1.72, fill: { color: LIGHT_BG }, line: { color: "F0E8E0", width: 0.5 } });
    s.addShape(pres.shapes.RECTANGLE, { x: 0.4, y: 3.48, w: metaW, h: 0.36, fill: { color: ORANGE }, line: { color: ORANGE } });
    s.addText("Meta Ads", { x: 0.55, y: 3.5, w: 3, h: 0.3, fontSize: 13, bold: true, color: WHITE, fontFace: "DM Sans" });
    const metaStats = [
      ["Inversión", fmtMoneyCompact(DATA.META_COSTO), DATA.META_COSTO_DELTA  || "", DATA.META_COSTO_DELTA_UP  === true],
      ["Clicks",    DATA.META_CLICKS || "", DATA.META_CLICKS_DELTA || "", DATA.META_CLICKS_DELTA_UP === true],
      ["CPL",       fmtMoneyCompact(DATA.META_CPL) || "", DATA.META_CPL_DELTA || "", DATA.META_CPL_DELTA_UP === true],
      ["Leads Zoho",String(parseNum(DATA.ZOHO_LEADS_META) || DATA.ZOHO_LEADS_META || ""), "", true],
    ];
    metaStats.forEach(([lbl, val, delta, up], i) => {
      const col = i % 2, row = Math.floor(i / 2);
      const bx = 0.55 + col * (metaW / 2), by = 3.92 + row * 0.58;
      s.addText(lbl, { x: bx, y: by, w: 2.0, h: 0.22, fontSize: 9, color: GRAY_TEXT, fontFace: "DM Sans" });
      s.addText([
        { text: val + "  ", options: { bold: true, color: DARK } },
        { text: delta,      options: { color: up ? GREEN : RED, bold: true } },
      ], { x: bx, y: by + 0.2, w: 2.0, h: 0.28, fontSize: 12, fontFace: "DM Sans" });
    });

    // Google block (condicional)
    if (hasGoogle) {
      s.addShape(pres.shapes.RECTANGLE, { x: 5.2, y: 3.48, w: 4.4, h: 1.72, fill: { color: LIGHT_BLUE }, line: { color: "D0E4F5", width: 0.5 } });
      s.addShape(pres.shapes.RECTANGLE, { x: 5.2, y: 3.48, w: 4.4, h: 0.36, fill: { color: BLUE }, line: { color: BLUE } });
      s.addText("Google Ads", { x: 5.35, y: 3.5, w: 3, h: 0.3, fontSize: 13, bold: true, color: WHITE, fontFace: "DM Sans" });
      const googleStats = [
        ["Inversión", fmtMoneyCompact(DATA.GOOGLE_COSTO), DATA.GOOGLE_COSTO_DELTA  || "", DATA.GOOGLE_COSTO_DELTA_UP  === true],
        ["Clicks",    DATA.GOOGLE_CLICKS || "", DATA.GOOGLE_CLICKS_DELTA || "", DATA.GOOGLE_CLICKS_DELTA_UP === true],
        ["CPL",       fmtMoneyCompact(DATA.GOOGLE_CPL) || "", DATA.GOOGLE_CPL_DELTA || "", DATA.GOOGLE_CPL_DELTA_UP === true],
        ["Leads Zoho",String(parseNum(DATA.ZOHO_LEADS_GOOGLE) || DATA.ZOHO_LEADS_GOOGLE || ""), "", true],
      ];
      googleStats.forEach(([lbl, val, delta, up], i) => {
        const col = i % 2, row = Math.floor(i / 2);
        const bx = 5.35 + col * 2.1, by = 3.92 + row * 0.58;
        s.addText(lbl, { x: bx, y: by, w: 2.0, h: 0.22, fontSize: 9, color: GRAY_TEXT, fontFace: "DM Sans" });
        s.addText([
          { text: val + "  ", options: { bold: true, color: DARK } },
          { text: delta,      options: { color: up ? GREEN : RED, bold: true } },
        ], { x: bx, y: by + 0.2, w: 2.0, h: 0.28, fontSize: 12, fontFace: "DM Sans" });
      });
    }
  }

  // ── SLIDE 3 – GOOGLE ADS (condicional) ────────────────────────────────────
  if (hasGoogle) {
    const s = pres.addSlide();
    s.background = { color: WHITE };
    s.addShape(pres.shapes.OVAL, { x: 0.15, y: 0.1, w: 0.72, h: 0.72, fill: { color: BLUE }, line: { color: BLUE } });
    s.addText("G", { x: 0.15, y: 0.1, w: 0.72, h: 0.72, fontSize: 18, bold: true, color: WHITE, fontFace: "DM Sans", align: "center", valign: "middle" });
    s.addText([
      { text: "Google Ads ", options: { bold: true, color: BLUE, fontSize: 26, fontFace: "DM Sans" } },
      { text: `– ${DATA.PERIODO_ACTUAL_LABEL || ""}`, options: { bold: true, color: DARK, fontSize: 26, fontFace: "DM Sans" } },
    ], { x: 1.0, y: 0.15, w: 8.8, h: 0.6 });

    const gKpis = [
      { label: "Inversión",    val: fmtMoneyCompact(DATA.GOOGLE_COSTO),   prev: fmtMoneyCompact(DATA.GOOGLE_COSTO_PREV),   delta: fmtDelta(DATA.GOOGLE_COSTO_DELTA),   up: DATA.GOOGLE_COSTO_DELTA_UP   === true },
      { label: "Clicks",       val: DATA.GOOGLE_CLICKS || "",              prev: DATA.GOOGLE_CLICKS_PREV  || "",            delta: fmtDelta(DATA.GOOGLE_CLICKS_DELTA),  up: DATA.GOOGLE_CLICKS_DELTA_UP  === true },
      { label: "Impresiones",  val: DATA.GOOGLE_IMPRESIONES || "",         prev: DATA.GOOGLE_IMPRESIONES_PREV || "",         delta: fmtDelta(DATA.GOOGLE_IMPRESIONES_DELTA), up: DATA.GOOGLE_IMPRESIONES_DELTA_UP === true },
      { label: "CTR",          val: DATA.GOOGLE_CTR  || "",                prev: DATA.GOOGLE_CTR_PREV     || "",            delta: fmtDelta(DATA.GOOGLE_CTR_DELTA),     up: DATA.GOOGLE_CTR_DELTA_UP     === true },
      { label: "CPL",          val: fmtMoneyCompact(DATA.GOOGLE_CPL),     prev: fmtMoneyCompact(DATA.GOOGLE_CPL_PREV),     delta: fmtDelta(DATA.GOOGLE_CPL_DELTA),     up: DATA.GOOGLE_CPL_DELTA_UP     === true },
      { label: "Leads Zoho",   val: String(parseNum(DATA.ZOHO_LEADS_GOOGLE) || DATA.ZOHO_LEADS_GOOGLE || ""), prev: DATA.ZOHO_LEADS_GOOGLE_PREV || "", delta: "", up: true },
    ];
    const cardW = 1.48, cardH = 2.0, startX = 0.35, startY = 0.95, gap = 0.08;
    gKpis.forEach((k, i) => {
      const x = startX + i * (cardW + gap), y = startY;
      s.addShape(pres.shapes.RECTANGLE, { x, y, w: cardW, h: cardH, fill: { color: LIGHT_BLUE }, line: { color: "D0E4F5", width: 0.5 } });
      s.addShape(pres.shapes.RECTANGLE, { x, y, w: cardW, h: 0.06, fill: { color: BLUE }, line: { color: BLUE } });
      s.addText(k.label, { x, y: y + 0.1,  w: cardW, h: 0.28, fontSize: 9,  color: GRAY_TEXT, fontFace: "DM Sans", align: "center" });
      s.addText(k.val,   { x, y: y + 0.42, w: cardW, h: 0.65, fontSize: 22, bold: true, color: DARK, fontFace: "DM Sans", align: "center" });
      s.addShape(pres.shapes.RECTANGLE, { x: x + 0.2, y: y + 1.12, w: cardW - 0.4, h: 0.26, fill: { color: k.up ? GREEN_BG : RED_BG }, line: { color: k.up ? GREEN_BG : RED_BG } });
      s.addText(k.delta, { x: x + 0.2, y: y + 1.12, w: cardW - 0.4, h: 0.26, fontSize: 10, bold: true, color: k.up ? GREEN : RED, fontFace: "DM Sans", align: "center" });
      s.addText(`${periodoAnt}: ${k.prev}`, { x, y: y + 1.5, w: cardW, h: 0.22, fontSize: 8, color: GRAY_TEXT, fontFace: "DM Sans", align: "center" });
    });

    // Campañas Google (condicional)
    if (Array.isArray(DATA.GOOGLE_CAMPANAS) && DATA.GOOGLE_CAMPANAS.length > 0) {
      const tY = 3.18, hdrs = ["Campaña", "Clicks", "Impresiones", "CTR", "Inversión"], colW = [3.8, 1.2, 1.5, 1.0, 1.6];
      const tW = colW.reduce((a, b) => a + b, 0), mX = (10 - tW) / 2;
      s.addShape(pres.shapes.RECTANGLE, { x: mX, y: tY, w: tW, h: 0.32, fill: { color: BLUE }, line: { color: BLUE } });
      let cx = mX + 0.1;
      hdrs.forEach((h, i) => { s.addText(h, { x: cx, y: tY + 0.04, w: colW[i] - 0.1, h: 0.24, fontSize: 9, bold: true, color: WHITE, fontFace: "DM Sans", align: i === 0 ? "left" : "center" }); cx += colW[i]; });
      DATA.GOOGLE_CAMPANAS.slice(0, 6).forEach((row, ri) => {
        const ry = tY + 0.32 + ri * 0.32, bg = ri % 2 === 0 ? WHITE : LIGHT_BLUE;
        s.addShape(pres.shapes.RECTANGLE, { x: mX, y: ry, w: tW, h: 0.32, fill: { color: bg }, line: { color: bg } });
        const vals = [row.nombre || "", row.clicks || "", row.impresiones || "", row.ctr || "", fmtMoneyCompact(row.costo)];
        let vx = mX + 0.1;
        vals.forEach((v, vi) => { s.addText(String(v), { x: vx, y: ry + 0.05, w: colW[vi] - 0.1, h: 0.22, fontSize: 9, color: vi === 0 ? BLUE : DARK, fontFace: "DM Sans", align: vi === 0 ? "left" : "center" }); vx += colW[vi]; });
      });
    }
  }

  // ── SLIDE 4 – META ADS ────────────────────────────────────────────────────
  {
    const s = pres.addSlide();
    s.background = { color: WHITE };
    s.addShape(pres.shapes.OVAL, { x: 0.15, y: 0.1, w: 0.72, h: 0.72, fill: { color: ORANGE }, line: { color: ORANGE } });
    s.addText("f", { x: 0.15, y: 0.1, w: 0.72, h: 0.72, fontSize: 18, bold: true, color: WHITE, fontFace: "DM Sans", align: "center", valign: "middle" });
    s.addText([
      { text: "Meta Ads ", options: { bold: true, color: ORANGE, fontSize: 26, fontFace: "DM Sans" } },
      { text: `– ${DATA.PERIODO_ACTUAL_LABEL || ""}`, options: { bold: true, color: DARK, fontSize: 26, fontFace: "DM Sans" } },
    ], { x: 1.0, y: 0.15, w: 8.8, h: 0.6 });

    const mKpis = [
      { label: "Inversión",   val: fmtMoneyCompact(DATA.META_COSTO),   prev: fmtMoneyCompact(DATA.META_COSTO_PREV),   delta: fmtDelta(DATA.META_COSTO_DELTA),   up: DATA.META_COSTO_DELTA_UP   === true },
      { label: "Clicks",      val: DATA.META_CLICKS || "",              prev: DATA.META_CLICKS_PREV  || "",           delta: fmtDelta(DATA.META_CLICKS_DELTA),  up: DATA.META_CLICKS_DELTA_UP  === true },
      { label: "Impresiones", val: DATA.META_IMPRESIONES || "",         prev: DATA.META_IMPRESIONES_PREV || "",        delta: fmtDelta(DATA.META_IMPRESIONES_DELTA), up: DATA.META_IMPRESIONES_DELTA_UP === true },
      { label: "CTR",         val: DATA.META_CTR  || "",                prev: DATA.META_CTR_PREV     || "",           delta: fmtDelta(DATA.META_CTR_DELTA),     up: DATA.META_CTR_DELTA_UP     === true },
      { label: "CPL",         val: fmtMoneyCompact(DATA.META_CPL),     prev: fmtMoneyCompact(DATA.META_CPL_PREV),     delta: fmtDelta(DATA.META_CPL_DELTA),     up: DATA.META_CPL_DELTA_UP     === true },
      { label: "Leads Zoho",  val: String(parseNum(DATA.ZOHO_LEADS_META) || DATA.ZOHO_LEADS_META || ""), prev: DATA.ZOHO_LEADS_META_PREV || "", delta: "", up: true },
    ];
    const cardW = 1.48, cardH = 2.0, startX = 0.35, startY = 0.95, gap = 0.08;
    mKpis.forEach((k, i) => {
      const x = startX + i * (cardW + gap), y = startY;
      s.addShape(pres.shapes.RECTANGLE, { x, y, w: cardW, h: cardH, fill: { color: LIGHT_BG }, line: { color: "F0E8E0", width: 0.5 } });
      s.addShape(pres.shapes.RECTANGLE, { x, y, w: cardW, h: 0.06, fill: { color: ORANGE }, line: { color: ORANGE } });
      s.addText(k.label, { x, y: y + 0.1,  w: cardW, h: 0.28, fontSize: 9,  color: GRAY_TEXT, fontFace: "DM Sans", align: "center" });
      s.addText(k.val,   { x, y: y + 0.42, w: cardW, h: 0.65, fontSize: 22, bold: true, color: DARK, fontFace: "DM Sans", align: "center" });
      s.addShape(pres.shapes.RECTANGLE, { x: x + 0.2, y: y + 1.12, w: cardW - 0.4, h: 0.26, fill: { color: k.up ? GREEN_BG : RED_BG }, line: { color: k.up ? GREEN_BG : RED_BG } });
      s.addText(k.delta, { x: x + 0.2, y: y + 1.12, w: cardW - 0.4, h: 0.26, fontSize: 10, bold: true, color: k.up ? GREEN : RED, fontFace: "DM Sans", align: "center" });
      s.addText(`${periodoAnt}: ${k.prev}`, { x, y: y + 1.5, w: cardW, h: 0.22, fontSize: 8, color: GRAY_TEXT, fontFace: "DM Sans", align: "center" });
    });

    // Campañas Meta (condicional)
    if (Array.isArray(DATA.META_CAMPANAS) && DATA.META_CAMPANAS.length > 0) {
      const tY = 3.18, hdrs = ["Campaña", "Clicks", "Impresiones", "CTR", "Inversión"], colW = [3.8, 1.2, 1.5, 1.0, 1.6];
      const tW = colW.reduce((a, b) => a + b, 0), mX = (10 - tW) / 2;
      s.addShape(pres.shapes.RECTANGLE, { x: mX, y: tY, w: tW, h: 0.32, fill: { color: ORANGE }, line: { color: ORANGE } });
      let cx = mX + 0.1;
      hdrs.forEach((h, i) => { s.addText(h, { x: cx, y: tY + 0.04, w: colW[i] - 0.1, h: 0.24, fontSize: 9, bold: true, color: WHITE, fontFace: "DM Sans", align: i === 0 ? "left" : "center" }); cx += colW[i]; });
      DATA.META_CAMPANAS.slice(0, 6).forEach((row, ri) => {
        const ry = tY + 0.32 + ri * 0.32, bg = ri % 2 === 0 ? WHITE : "FFF8F5";
        s.addShape(pres.shapes.RECTANGLE, { x: mX, y: ry, w: tW, h: 0.32, fill: { color: bg }, line: { color: bg } });
        const vals = [row.nombre || "", row.clicks || "", row.impresiones || "", row.ctr || "", fmtMoneyCompact(row.costo)];
        let vx = mX + 0.1;
        vals.forEach((v, vi) => { s.addText(String(v), { x: vx, y: ry + 0.05, w: colW[vi] - 0.1, h: 0.22, fontSize: 9, color: vi === 0 ? ORANGE : DARK, fontFace: "DM Sans", align: vi === 0 ? "left" : "center" }); vx += colW[vi]; });
      });
    }
  }

  // ── SLIDE 5 – COMPOSICIÓN DE LEADS (Zoho por fuente) ─────────────────────
  {
    const s = pres.addSlide();
    s.background = { color: WHITE };
    s.addText("Composición de Leads", { x: 0.5, y: 0.22, w: 8, h: 0.55, fontSize: 28, bold: true, color: DARK, fontFace: "DM Sans" });
    s.addText(`Zoho CRM — fuente de verdad  ·  ${DATA.PERIODO_ACTUAL_LABEL || ""}`, { x: 0.5, y: 0.78, w: 9, h: 0.3, fontSize: 13, color: GRAY_TEXT, fontFace: "DM Sans" });

    const total   = parseNum(DATA.ZOHO_LEADS_TOTAL) || 1;
    const nGoogle = parseNum(DATA.ZOHO_LEADS_GOOGLE);
    const nMeta   = parseNum(DATA.ZOHO_LEADS_META);
    const nWeb    = parseNum(DATA.ZOHO_LEADS_WEB);
    const nOtros  = Math.max(total - nGoogle - nMeta - nWeb, 0);

    const fuentes = [
      { label: "Google Ads",    n: nGoogle, color: BLUE,    bg: LIGHT_BLUE },
      { label: "Meta Ads",      n: nMeta,   color: ORANGE,  bg: LIGHT_BG   },
      { label: "Formulario Web",n: nWeb,    color: "3B6D11",bg: GREEN_BG   },
      ...(nOtros > 0 ? [{ label: "Otros", n: nOtros, color: GRAY_TEXT, bg: "F1F0EC" }] : []),
    ].filter(f => f.n > 0);

    // Barra horizontal apilada
    const barX = 0.5, barY = 1.4, barW = 9.0, barH = 0.7;
    let offsetX = barX;
    fuentes.forEach(f => {
      const w = (f.n / total) * barW;
      s.addShape(pres.shapes.RECTANGLE, { x: offsetX, y: barY, w, h: barH, fill: { color: f.color }, line: { color: f.color } });
      if (w > 0.5) s.addText(`${Math.round(f.n / total * 100)}%`, { x: offsetX, y: barY, w, h: barH, fontSize: 12, bold: true, color: WHITE, fontFace: "DM Sans", align: "center", valign: "middle" });
      offsetX += w;
    });

    // Cards por fuente
    const cW = 9.0 / Math.max(fuentes.length, 1), cGap = 0.15;
    fuentes.forEach((f, i) => {
      const x = 0.5 + i * cW, y = 2.35;
      s.addShape(pres.shapes.RECTANGLE, { x: x + cGap / 2, y, w: cW - cGap, h: 2.5, fill: { color: f.bg }, line: { color: "F0E8E0", width: 0.5 } });
      s.addShape(pres.shapes.RECTANGLE, { x: x + cGap / 2, y, w: cW - cGap, h: 0.07, fill: { color: f.color }, line: { color: f.color } });
      s.addText(f.label, { x: x + cGap / 2, y: y + 0.15, w: cW - cGap, h: 0.3,  fontSize: 10, color: GRAY_TEXT, fontFace: "DM Sans", align: "center" });
      s.addText(String(f.n), { x: x + cGap / 2, y: y + 0.52, w: cW - cGap, h: 0.75, fontSize: 36, bold: true, color: f.color, fontFace: "DM Sans", align: "center" });
      s.addText(`${Math.round(f.n / total * 100)}% del total`, { x: x + cGap / 2, y: y + 1.32, w: cW - cGap, h: 0.3, fontSize: 11, color: GRAY_TEXT, fontFace: "DM Sans", align: "center" });
      const pctChange = DATA[`ZOHO_LEADS_${f.label.toUpperCase().replace(/ /g, "_")}_DELTA`] || "";
      if (pctChange) {
        const up = DATA[`ZOHO_LEADS_${f.label.toUpperCase().replace(/ /g, "_")}_DELTA_UP`] === true;
        s.addShape(pres.shapes.RECTANGLE, { x: x + cGap / 2 + 0.3, y: y + 1.72, w: cW - cGap - 0.6, h: 0.28, fill: { color: up ? GREEN_BG : RED_BG }, line: { color: up ? GREEN_BG : RED_BG } });
        s.addText(fmtDelta(pctChange), { x: x + cGap / 2 + 0.3, y: y + 1.72, w: cW - cGap - 0.6, h: 0.28, fontSize: 11, bold: true, color: up ? GREEN : RED, fontFace: "DM Sans", align: "center" });
      }
    });

    s.addText(`Total leads período: ${parseNum(DATA.ZOHO_LEADS_TOTAL)}  ·  Período anterior: ${DATA.ZOHO_LEADS_PREV || "—"}`, { x: 0.5, y: 5.08, w: 9, h: 0.25, fontSize: 9, color: GRAY_TEXT, fontFace: "DM Sans", italic: true });
  }

  // ── SLIDE 6 – CALIDAD DE LEADS (por estado Zoho) ──────────────────────────
  {
    const s = pres.addSlide();
    s.background = { color: WHITE };
    s.addText("Calidad de Leads", { x: 0.5, y: 0.22, w: 8, h: 0.55, fontSize: 28, bold: true, color: DARK, fontFace: "DM Sans" });
    s.addText(`Distribución por estado — Zoho CRM  ·  ${DATA.PERIODO_ACTUAL_LABEL || ""}`, { x: 0.5, y: 0.78, w: 9, h: 0.3, fontSize: 13, color: GRAY_TEXT, fontFace: "DM Sans" });

    const total     = parseNum(DATA.ZOHO_LEADS_TOTAL) || 1;
    const nNuevo    = parseNum(DATA.ZOHO_LEADS_NUEVO);
    const nContact  = parseNum(DATA.ZOHO_LEADS_CONTACTADO);
    const nMuerto   = parseNum(DATA.ZOHO_LEADS_MUERTO);
    const nOtros    = Math.max(total - nNuevo - nContact - nMuerto, 0);

    const estados = [
      { label: "Nuevo",       n: nNuevo,   color: "185FA5", bg: LIGHT_BLUE, desc: "Sin gestionar aún"          },
      { label: "Contactado",  n: nContact, color: AMBER,    bg: AMBER_BG,   desc: "En proceso de seguimiento"  },
      { label: "Calificado",  n: nOtros,   color: "3B6D11", bg: GREEN_BG,   desc: "Oportunidad avanzada"       },
      { label: "No califica", n: nMuerto,  color: "A32D2D", bg: RED_BG,     desc: "Descartado / sin perfil"    },
    ].filter(e => e.n > 0);

    // Barra total como referencia
    s.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 1.35, w: 9.0, h: 0.08, fill: { color: "E8E0D8" }, line: { color: "E8E0D8" } });
    let ox = 0.5;
    estados.forEach(e => { const w = (e.n / total) * 9.0; s.addShape(pres.shapes.RECTANGLE, { x: ox, y: 1.35, w, h: 0.08, fill: { color: e.color }, line: { color: e.color } }); ox += w; });

    // Cards
    const cols = Math.min(estados.length, 4);
    const cW = 9.0 / cols;
    estados.forEach((e, i) => {
      const x = 0.5 + i * cW, y = 1.6;
      s.addShape(pres.shapes.RECTANGLE, { x: x + 0.1, y, w: cW - 0.2, h: 3.2, fill: { color: e.bg }, line: { color: "F0E8E0", width: 0.5 } });
      s.addShape(pres.shapes.RECTANGLE, { x: x + 0.1, y, w: cW - 0.2, h: 0.07, fill: { color: e.color }, line: { color: e.color } });
      s.addText(e.label, { x: x + 0.1, y: y + 0.14, w: cW - 0.2, h: 0.32, fontSize: 13, bold: true, color: e.color, fontFace: "DM Sans", align: "center" });
      s.addText(String(e.n), { x: x + 0.1, y: y + 0.55, w: cW - 0.2, h: 0.9,  fontSize: 48, bold: true, color: e.color, fontFace: "DM Sans", align: "center" });
      s.addText(`${Math.round(e.n / total * 100)}%`, { x: x + 0.1, y: y + 1.55, w: cW - 0.2, h: 0.4, fontSize: 18, color: e.color, fontFace: "DM Sans", align: "center" });
      s.addShape(pres.shapes.RECTANGLE, { x: x + 0.3, y: y + 2.0, w: cW - 0.6, h: 0.02, fill: { color: "E8E0D8" }, line: { color: "E8E0D8" } });
      s.addText(e.desc, { x: x + 0.1, y: y + 2.15, w: cW - 0.2, h: 0.35, fontSize: 10, color: GRAY_TEXT, fontFace: "DM Sans", align: "center", italic: true });
    });

    s.addText(`Total: ${parseNum(DATA.ZOHO_LEADS_TOTAL)} leads procesados en Zoho CRM`, { x: 0.5, y: 5.08, w: 9, h: 0.25, fontSize: 9, color: GRAY_TEXT, fontFace: "DM Sans", italic: true });
  }

  // ── SLIDE 7 – RESULTADOS COMERCIALES (condicional) ────────────────────────
  if (Array.isArray(DATA.FUNNEL_ROWS) && DATA.FUNNEL_ROWS.length > 0) {
    const s = pres.addSlide();
    s.background = { color: WHITE };
    s.addText("Resultados Comerciales", { x: 0.5, y: 0.22, w: 8, h: 0.55, fontSize: 28, bold: true, color: DARK, fontFace: "DM Sans" });
    s.addText("Evolución mensual de leads y performance", { x: 0.5, y: 0.78, w: 9, h: 0.3, fontSize: 13, color: GRAY_TEXT, fontFace: "DM Sans" });

    const rowH = 0.42, y0 = 1.15, x0 = 0.4;
    const colW = [1.5, 1.1, 1.0, 1.0, 1.0, 1.5, 1.1, 1.1];
    const hdrs = ["Mes", "Google", "Meta", "Web", "Total", "Inversión", "CPL", "Cierres"];
    const algs = ["left", "right", "right", "right", "right", "right", "right", "right"];
    const tW   = colW.reduce((a, b) => a + b, 0);

    // Header
    s.addShape(pres.shapes.RECTANGLE, { x: x0, y: y0, w: tW, h: rowH, fill: { color: ORANGE }, line: { color: ORANGE } });
    let cx = x0 + 0.12;
    hdrs.forEach((h, i) => {
      s.addText(h, { x: cx, y: y0 + 0.1, w: colW[i] - 0.12, h: 0.24, fontSize: 9, bold: true, color: WHITE, fontFace: "DM Sans", align: algs[i] });
      cx += colW[i];
    });

    // Filas históricas
    DATA.FUNNEL_ROWS.forEach((row, ri) => {
      const ry = y0 + rowH + ri * rowH;
      const bg = ri % 2 === 0 ? WHITE : LIGHT_BG;
      s.addShape(pres.shapes.RECTANGLE, { x: x0, y: ry, w: tW, h: rowH, fill: { color: bg }, line: { color: bg } });
      const vals = [
        row.mes || "",
        String(row.leads_google || row.google || "—"),
        String(row.leads_meta   || row.meta   || "—"),
        String(row.leads_web    || row.web     || "—"),
        String(row.total        || row.leads   || "—"),
        fmtMoneyCompact(row.inversion) || "—",
        fmtMoneyCompact(row.cpl)       || "—",
        String(row.cierres || "—"),
      ];
      cx = x0 + 0.12;
      vals.forEach((v, vi) => {
        s.addText(v, { x: cx, y: ry + 0.1, w: colW[vi] - 0.12, h: 0.24, fontSize: 9.5, color: vi === 0 ? DARK : GRAY_TEXT, bold: vi === 0, fontFace: "DM Sans", align: algs[vi] });
        cx += colW[vi];
      });
    });

    // Fila mes actual (destacada)
    const lastRi = DATA.FUNNEL_ROWS.length;
    const ry = y0 + rowH + lastRi * rowH;
    s.addShape(pres.shapes.RECTANGLE, { x: x0, y: ry, w: tW, h: rowH, fill: { color: "FFF0E8" }, line: { color: ORANGE, width: 0.5 } });
    const curVals = [
      DATA.PERIODO_ACTUAL_LABEL || "Actual",
      String(parseNum(DATA.ZOHO_LEADS_GOOGLE) || DATA.ZOHO_LEADS_GOOGLE || "—"),
      String(parseNum(DATA.ZOHO_LEADS_META)   || DATA.ZOHO_LEADS_META   || "—"),
      String(parseNum(DATA.ZOHO_LEADS_WEB)    || DATA.ZOHO_LEADS_WEB    || "—"),
      String(parseNum(DATA.ZOHO_LEADS_TOTAL)  || DATA.ZOHO_LEADS_TOTAL  || "—"),
      fmtMoneyCompact(DATA.INVERSION_TOTAL) || "—",
      fmtMoneyCompact(DATA.CPL_TOTAL) || "—",
      "—",
    ];
    cx = x0 + 0.12;
    curVals.forEach((v, vi) => {
      s.addText(v, { x: cx, y: ry + 0.1, w: colW[vi] - 0.12, h: 0.24, fontSize: 9.5, bold: true, color: ORANGE, fontFace: "DM Sans", align: algs[vi] });
      cx += colW[vi];
    });
  }

  // ── SLIDE 8 – RECOMENDACIONES ─────────────────────────────────────────────
  buildSlide_Recommendations(pres, DATA);

  // ── SLIDE 9 – CIERRE ──────────────────────────────────────────────────────
  buildSlide_Close(pres, DATA);

  return await pres.write({ outputType: "base64" });
}
