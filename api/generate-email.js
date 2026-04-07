const pptxgen = require("pptxgenjs");
const {
  COLORS,
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
    const base64 = await generatePptx(DATA);
    const filename = `Reporte_Email_${DATA.CLIENTE_NOMBRE || "Cliente"}_${DATA.PERIODO_ACTUAL_LABEL || "Periodo"}.pptx`
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
  pres.title = `Reporte Email Marketing ${DATA.CLIENTE_NOMBRE || ""} - ${DATA.PERIODO_ACTUAL_LABEL || ""}`;

  // ── Brand colors ──────────────────────────────────────────────────────────
  const { ORANGE, ORANGE2, DARK, WHITE, LIGHT_BG, GRAY_TEXT,
          GREEN, GREEN_BG, RED, RED_BG, BLUE, LIGHT_BLUE } = COLORS;

  // ── Helpers ───────────────────────────────────────────────────────────────
  const parseNum  = str => { if (typeof str === "number") return str; return parseFloat((str || "0").replace(/\./g, "").replace(",", ".").replace(/[^0-9.]/g, "")) || 0; };
  const parseRate = str => parseNum((str || "0").replace("%", "").replace("pp", ""));
  const fmtDelta  = str => (str || "").replace(/pp$/i, "%");
  const fmtMoneyCompact = val => { const n = parseNum(val); if (n >= 1_000_000) return `${(n / 1_000_000).toFixed(2).replace(".", ",")} M`; if (n >= 1_000) return `${(n / 1_000).toFixed(1).replace(".", ",")} K`; return String(val || ""); };
  const hasEcommerce = !!(DATA.EMAIL_INGRESOS || DATA.EMAIL_TRANSACCIONES);
  const hasGA4       = DATA.GA4_TIENE_DATOS === true;
  const plataforma   = (DATA.PLATAFORMA_EMAIL || DATA.Plataforma || "").toUpperCase();
  // Detectar Woowup también por presencia de los arrays separados por tipo
  const isWoowup     = plataforma === "WOOWUP"
                    || Array.isArray(DATA.EMAIL_CAMPANAS_NEWSLETTER)
                    || Array.isArray(DATA.EMAIL_CAMPANAS_AUTOMATIZADA);

  const todasCampanas        = Array.isArray(DATA.EMAIL_CAMPANAS) ? DATA.EMAIL_CAMPANAS : [];
  const getTipo              = c => (c.tipo || c.Tipo || "").trim().toLowerCase();
  // Acepta arrays pre-spliteados desde n8n o splitea internamente desde EMAIL_CAMPANAS
  const campañasNewsletter   = Array.isArray(DATA.EMAIL_CAMPANAS_NEWSLETTER)
                             ? DATA.EMAIL_CAMPANAS_NEWSLETTER
                             : todasCampanas.filter(c => getTipo(c) === "newsletter");
  const campañasAutomatizada = Array.isArray(DATA.EMAIL_CAMPANAS_AUTOMATIZADA)
                             ? DATA.EMAIL_CAMPANAS_AUTOMATIZADA
                             : todasCampanas.filter(c => getTipo(c) === "automatizada" || getTipo(c) === "automtizada");
  const campañasUnicas       = todasCampanas;

  const sortByIngresos = (a, b) => parseNum((b.ingresos || "0").replace(/[^0-9,]/g, "").replace(",", "."))
                                 - parseNum((a.ingresos || "0").replace(/[^0-9,]/g, "").replace(",", "."));
  const top3Newsletter   = [...campañasNewsletter].sort(sortByIngresos).slice(0, 3);
  const top3Automatizada = [...campañasAutomatizada].sort(sortByIngresos).slice(0, 3);
  const top3Unicas       = [...campañasUnicas].sort(sortByIngresos).slice(0, 3);

  // ── Helper: tabla de campañas (paginada) ─────────────────────────────────
  function buildSlideTabla(campanas, tipo, accentColor, accentBg) {
    if (campanas.length === 0) return;

    // ── Calcular fila de totales/promedios del array completo ────────────
    const pN = str => parseNum((str || "0").replace("%", ""));
    let totalEnvios = 0, totalTrans = 0, totalIngresos = 0;
    let sumAp = 0, sumCtor = 0, sumBajas = 0;
    campanas.forEach(c => {
      const e = pN(c.envios);
      totalEnvios   += e;
      totalTrans    += pN(c.transacciones);
      totalIngresos += pN(c.ingresos);
      sumAp         += pN(c.apertura) * e;
      sumCtor       += pN(c.ctor)     * e;
      sumBajas      += pN(c.bajas)    * e;
    });
    const fN  = n => Math.round(n).toLocaleString("es-AR");
    const fP  = n => n.toFixed(1).replace(".", ",") + "%";
    const fM  = n => "$" + Math.round(n).toLocaleString("es-AR");
    const totalsRow = {
      nombre:        "TOTAL / PROMEDIO",
      envios:        fN(totalEnvios),
      apertura:      totalEnvios > 0 ? fP(sumAp   / totalEnvios) : "—",
      ctor:          totalEnvios > 0 ? fP(sumCtor  / totalEnvios) : "—",
      bajas:         totalEnvios > 0 ? fP(sumBajas / totalEnvios) : "—",
      transacciones: totalTrans    > 0 ? fN(totalTrans)    : undefined,
      ingresos:      totalIngresos > 0 ? fM(totalIngresos) : undefined,
    };

    const ROWS_PER_SLIDE = 12;
    const totalPages = Math.ceil(campanas.length / ROWS_PER_SLIDE);

    for (let page = 0; page < totalPages; page++) {
      const isLastPage = page === totalPages - 1;
      const slice      = campanas.slice(page * ROWS_PER_SLIDE, (page + 1) * ROWS_PER_SLIDE);
      const pageLabel  = totalPages > 1 ? ` (${page + 1}/${totalPages})` : "";

      const s = pres.addSlide();
      s.background = { color: WHITE };

      s.addText([
        { text: tipo + pageLabel + " ", options: { bold: true, color: accentColor, fontSize: 26, fontFace: "DM Sans" } },
        { text: `– ${DATA.PERIODO_ACTUAL_LABEL || ""}`, options: { bold: true, color: DARK, fontSize: 26, fontFace: "DM Sans" } },
      ], { x: 1.0, y: 0.15, w: 8.8, h: 0.6 });

      s.addShape(pres.shapes.OVAL, { x: 0.15, y: 0.1, w: 0.72, h: 0.72, fill: { color: accentColor }, line: { color: accentColor } });
      s.addText("@", { x: 0.15, y: 0.1, w: 0.72, h: 0.72, fontSize: 16, bold: true, color: WHITE, fontFace: "DM Sans", align: "center", valign: "middle" });

      const cols = hasEcommerce
        ? [
            { hdr: "Campaña",   key: "nombre",       w: 2.55, align: "left",   color: accentColor },
            { hdr: "Envíos",    key: "envios",        w: 0.9,  align: "center", color: DARK },
            { hdr: "Apertura",  key: "apertura",      w: 0.9,  align: "center", color: DARK },
            { hdr: "CTOR",      key: "ctor",          w: 0.75, align: "center", color: DARK },
            { hdr: "Bajas",     key: "bajas",         w: 0.75, align: "center", color: DARK },
            { hdr: "Transacc.", key: "transacciones", w: 0.85, align: "center", color: DARK },
            { hdr: "Ingresos",  key: "ingresos",      w: 1.15, align: "center", color: DARK },
          ]
        : [
            { hdr: "Campaña",  key: "nombre",   w: 3.5,  align: "left",   color: accentColor },
            { hdr: "Envíos",   key: "envios",   w: 1.2,  align: "center", color: DARK },
            { hdr: "Apertura", key: "apertura", w: 1.2,  align: "center", color: DARK },
            { hdr: "CTOR",     key: "ctor",     w: 1.0,  align: "center", color: DARK },
            { hdr: "Bajas",    key: "bajas",    w: 1.0,  align: "center", color: DARK },
          ];

      const totalW  = cols.reduce((acc, c) => acc + c.w, 0);
      const marginX = (10 - totalW) / 2;
      const tY      = 0.88;

      // Header
      s.addShape(pres.shapes.RECTANGLE, { x: marginX, y: tY, w: totalW, h: 0.34, fill: { color: accentColor }, line: { color: accentColor } });
      let cx = marginX + 0.1;
      cols.forEach(c => {
        s.addText(c.hdr, { x: cx, y: tY + 0.04, w: c.w - 0.1, h: 0.26, fontSize: 9, bold: true, color: WHITE, fontFace: "DM Sans", align: c.align, valign: "middle" });
        cx += c.w;
      });

      // Rows
      slice.forEach((row, i) => {
        const ry = tY + 0.34 + i * 0.33;
        const bg = i % 2 === 0 ? WHITE : "FAFAFA";
        s.addShape(pres.shapes.RECTANGLE, { x: marginX, y: ry, w: totalW, h: 0.33, fill: { color: bg }, line: { color: "EEEEEE", width: 0.3 } });
        let rx = marginX + 0.1;
        cols.forEach(c => {
          let val = row[c.key] || "";
          if (c.key === "nombre" && val.length > 40) val = val.substring(0, 38) + "…";
          s.addText(val, { x: rx, y: ry + 0.05, w: c.w - 0.1, h: 0.24, fontSize: 8.5, color: c.color, fontFace: "DM Sans", align: c.align, valign: "middle" });
          rx += c.w;
        });
      });

      // Fila de totales — solo en la última página
      if (isLastPage) {
        const ry = tY + 0.34 + slice.length * 0.33;
        s.addShape(pres.shapes.RECTANGLE, { x: marginX, y: ry, w: totalW, h: 0.33, fill: { color: "FFF0E8" }, line: { color: accentColor, width: 0.5 } });
        let rx = marginX + 0.1;
        cols.forEach(c => {
          const val = totalsRow[c.key] !== undefined ? (totalsRow[c.key] || "—") : "—";
          s.addText(val, { x: rx, y: ry + 0.05, w: c.w - 0.1, h: 0.24, fontSize: 8.5, bold: true, color: accentColor, fontFace: "DM Sans", align: c.align, valign: "middle" });
          rx += c.w;
        });
      }
    }
  }

  // ── Helper: top 3 por apertura ────────────────────────────────────────────
  function buildSlideTop3(top3, tipo, accentColor) {
    if (top3.length === 0) return;
    const s = pres.addSlide();
    s.background = { color: WHITE };

    s.addText(`Top 3 ${tipo}`, { x: 0.5, y: 0.22, w: 7, h: 0.55, fontSize: 28, bold: true, color: DARK, fontFace: "DM Sans" });
    s.addText("por ingresos online", { x: 0.5, y: 0.78, w: 7, h: 0.3, fontSize: 13, color: accentColor, fontFace: "DM Sans", bold: true });
    s.addText(DATA.PERIODO_ACTUAL_LABEL || "", { x: 6.5, y: 0.3, w: 3, h: 0.3, fontSize: 11, color: GRAY_TEXT, fontFace: "DM Sans", align: "right" });

    const cardW  = 2.85;
    const cardX  = [0.4, 3.55, 6.7];
    const medals = ["🥇", "🥈", "🥉"];

    top3.forEach((c, i) => {
      const x = cardX[i];
      const y = 1.25;

      s.addShape(pres.shapes.RECTANGLE, { x, y, w: cardW, h: 4.2, fill: { color: LIGHT_BG }, line: { color: "F0E8E0", width: 0.5 } });
      s.addShape(pres.shapes.RECTANGLE, { x, y, w: cardW, h: 0.08, fill: { color: accentColor }, line: { color: accentColor } });

      s.addText(`${medals[i]}  #${i + 1}`, { x, y: y + 0.15, w: cardW, h: 0.3, fontSize: 11, bold: true, color: accentColor, fontFace: "DM Sans", align: "center" });

      const nombre = (c.nombre || "").length > 50 ? (c.nombre || "").substring(0, 48) + "…" : (c.nombre || "");
      s.addText(nombre, { x: x + 0.1, y: y + 0.48, w: cardW - 0.2, h: 0.6, fontSize: 10.5, bold: true, color: DARK, fontFace: "DM Sans", align: "center", wrap: true });

      // Placeholder para captura de pantalla
      s.addShape(pres.shapes.RECTANGLE, { x: x + 0.12, y: y + 1.1, w: cardW - 0.24, h: 2.0,
        fill: { color: "F5F5F5" }, line: { color: "CCCCCC", width: 0.5, dashType: "dash" } });
      s.addText("📷  Insertar captura", { x: x + 0.12, y: y + 1.1, w: cardW - 0.24, h: 2.0,
        fontSize: 9, color: "AAAAAA", fontFace: "DM Sans", align: "center", valign: "middle" });

      [{ lbl: "Ingresos", val: c.ingresos || "—" }, { lbl: "Apertura", val: c.apertura || "—" }, { lbl: "Envíos", val: c.envios || "—" }]
        .forEach((d, di) => {
          const dy = y + 3.15 + di * 0.30;
          s.addShape(pres.shapes.RECTANGLE, { x: x + 0.15, y: dy, w: cardW - 0.3, h: 0.33, fill: { color: "F8F4F0" }, line: { color: "EEE8E0", width: 0.3 } });
          s.addText(d.lbl, { x: x + 0.22, y: dy + 0.05, w: 1.0,         h: 0.24, fontSize: 9, color: GRAY_TEXT, fontFace: "DM Sans" });
          s.addText(d.val, { x: x + 1.2,  y: dy + 0.05, w: cardW - 1.4, h: 0.24, fontSize: 9, bold: true, color: DARK, fontFace: "DM Sans", align: "right" });
        });
    });
  }

  // ── SLIDE 1 – COVER ───────────────────────────────────────────────────────
  {
    const s = pres.addSlide();
    s.background = { color: ORANGE };

    s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 0.18, h: 5.625, fill: { color: WHITE }, line: { color: WHITE } });
    s.addShape(pres.shapes.OVAL, { x: 7.8, y: -1.2, w: 4.0, h: 4.0, fill: { color: WHITE, transparency: 88 }, line: { color: WHITE, transparency: 88 } });
    s.addShape(pres.shapes.OVAL, { x: 8.5, y: -0.4, w: 2.4, h: 2.4, fill: { color: WHITE, transparency: 75 }, line: { color: WHITE, transparency: 75 } });

    s.addShape(pres.shapes.OVAL, { x: 0.5,  y: 0.45, w: 0.52, h: 0.52, fill: { color: WHITE  }, line: { color: WHITE  } });
    s.addShape(pres.shapes.OVAL, { x: 0.64, y: 0.59, w: 0.26, h: 0.26, fill: { color: ORANGE }, line: { color: ORANGE } });
    s.addText("Known Online", { x: 1.15, y: 0.48, w: 3.5, h: 0.45, fontSize: 15, color: WHITE, bold: true, fontFace: "DM Sans", margin: 0 });

    s.addShape(pres.shapes.RECTANGLE, { x: 0.45, y: 1.5, w: 1.8, h: 0.32, fill: { color: WHITE }, line: { color: WHITE } });
    s.addText(DATA.CLIENTE_NOMBRE || "CLIENTE", { x: 0.45, y: 1.5, w: 1.8, h: 0.32, fontSize: 10, color: ORANGE, bold: true, fontFace: "DM Sans", align: "center", margin: 0 });

    s.addText("Reporte\nEmail Marketing", { x: 0.45, y: 1.95, w: 7.5, h: 1.5, fontSize: 48, color: WHITE, bold: true, fontFace: "DM Sans", valign: "top" });
    s.addText(`${DATA.PERIODO_ACTUAL_LABEL || ""} vs. ${DATA.PERIODO_ANTERIOR_LABEL || ""}`, { x: 0.45, y: 3.55, w: 7, h: 0.45, fontSize: 18, color: "FFD4B8", fontFace: "DM Sans" });

    if (DATA.PLATAFORMA_EMAIL) {
      s.addText(`Plataforma: ${DATA.PLATAFORMA_EMAIL}`, { x: 0.45, y: 5.28, w: 9.1, h: 0.22, fontSize: 9, color: "FFD4B8", fontFace: "DM Sans" });
    }
  }

  // ── SLIDE 2 – RESUMEN PLATAFORMA ──────────────────────────────────────────
  {
    const s = pres.addSlide();
    s.background = { color: WHITE };

    s.addText("Resumen de Plataforma", { x: 0.5, y: 0.22, w: 8, h: 0.55, fontSize: 28, bold: true, color: DARK, fontFace: "DM Sans" });
    s.addText(`${DATA.PLATAFORMA_EMAIL || "Email Marketing"}  ·  ${DATA.PERIODO_ACTUAL_LABEL || ""} vs ${DATA.PERIODO_ANTERIOR_LABEL || ""}`, {
      x: 0.5, y: 0.78, w: 9, h: 0.3, fontSize: 13, color: GRAY_TEXT, fontFace: "DM Sans",
    });

    const kpis = [
      { label: "Envíos",           val: DATA.EMAIL_ENVIOS    || "", delta: fmtDelta(DATA.EMAIL_ENVIOS_DELTA), up: DATA.EMAIL_ENVIOS_DELTA_UP    === true, prev: DATA.EMAIL_ENVIOS_PREV    || "" },
      { label: "Tasa de apertura", val: DATA.EMAIL_APERTURA  || "", delta: fmtDelta(DATA.EMAIL_APERTURA_DELTA), up: DATA.EMAIL_APERTURA_DELTA_UP  === true, prev: DATA.EMAIL_APERTURA_PREV  || "" },
      { label: "CTOR",             val: DATA.EMAIL_CTOR      || "", delta: fmtDelta(DATA.EMAIL_CTOR_DELTA), up: DATA.EMAIL_CTOR_DELTA_UP      === true, prev: DATA.EMAIL_CTOR_PREV      || "" },
      { label: "Tasa de bajas",    val: DATA.EMAIL_BAJAS     || "", delta: fmtDelta(DATA.EMAIL_BAJAS_DELTA), up: DATA.EMAIL_BAJAS_DELTA_UP     !== true, prev: DATA.EMAIL_BAJAS_PREV     || "" },
    ];
    if (DATA.EMAIL_TRANSACCIONES) kpis.push({ label: "Transacciones",  val: DATA.EMAIL_TRANSACCIONES || "", delta: fmtDelta(DATA.EMAIL_TRANSACCIONES_DELTA), up: DATA.EMAIL_TRANSACCIONES_DELTA_UP === true, prev: DATA.EMAIL_TRANSACCIONES_PREV || "" });
    if (DATA.EMAIL_INGRESOS)      kpis.push({ label: "Ingresos online", val: DATA.EMAIL_INGRESOS      || "", delta: fmtDelta(DATA.EMAIL_INGRESOS_DELTA), up: DATA.EMAIL_INGRESOS_DELTA_UP      === true, prev: DATA.EMAIL_INGRESOS_PREV      || "" });

    const cardW      = 2.1;
    const itemsPerRow = Math.min(kpis.length, 4);
    const startX     = (10 - itemsPerRow * cardW - (itemsPerRow - 1) * 0.22) / 2;

    kpis.forEach((k, i) => {
      const row = Math.floor(i / 4);
      const col = i % 4;
      const x   = startX + col * (cardW + 0.22);
      const y   = 1.2 + row * 2.0;

      s.addShape(pres.shapes.RECTANGLE, { x, y, w: cardW, h: 1.65, fill: { color: LIGHT_BG }, line: { color: "F0E8E0", width: 0.5 } });
      s.addShape(pres.shapes.RECTANGLE, { x, y, w: cardW, h: 0.07, fill: { color: ORANGE }, line: { color: ORANGE } });
      s.addText(k.label, { x, y: y + 0.12, w: cardW, h: 0.3,  fontSize: 9.5, color: GRAY_TEXT,  fontFace: "DM Sans",       align: "center" });
      s.addText(k.val,   { x, y: y + 0.42, w: cardW, h: 0.55, fontSize: 24,  bold: true, color: DARK, fontFace: "DM Sans", align: "center" });
      s.addShape(pres.shapes.RECTANGLE, { x: x + cardW * 0.25, y: y + 1.0, w: cardW * 0.5, h: 0.27, fill: { color: k.up ? GREEN_BG : RED_BG }, line: { color: k.up ? GREEN_BG : RED_BG } });
      s.addText(k.delta, { x: x + cardW * 0.25, y: y + 1.0, w: cardW * 0.5, h: 0.27, fontSize: 11, bold: true, color: k.up ? GREEN : RED, fontFace: "DM Sans", align: "center" });
      s.addText(`${DATA.PERIODO_ANTERIOR_LABEL || "Ant."}: ${k.prev}`, { x, y: y + 1.35, w: cardW, h: 0.22, fontSize: 8.5, color: GRAY_TEXT, fontFace: "DM Sans", align: "center" });
    });
  }

  // ── SLIDES CAMPAÑAS — estructura según plataforma ────────────────────────
  if (isWoowup) {
    buildSlideTabla(campañasNewsletter,   "Newsletter",    ORANGE, LIGHT_BG);
    buildSlideTop3 (top3Newsletter,       "Newsletter",    ORANGE);
    buildSlideTabla(campañasAutomatizada, "Automatizadas", ORANGE, LIGHT_BG);
    buildSlideTop3 (top3Automatizada,     "Automatizadas", ORANGE);
  } else {
    // MAIUP / ICOMM / otras: tabla única
    const label = plataforma === "MAIUP" ? "Mailup" : plataforma === "ICOMM" ? "Icomm" : (DATA.PLATAFORMA_EMAIL || "Email");
    buildSlideTabla(campañasUnicas, label, ORANGE, LIGHT_BG);
    buildSlideTop3 (top3Unicas,    label, ORANGE);
  }

  // ── SLIDE – GA4 CANAL EMAIL (condicional) ─────────────────────────────────
  if (hasGA4) {
    const s = pres.addSlide();
    s.background = { color: WHITE };

    s.addText("Rendimiento en Sitio Web", { x: 0.5, y: 0.22, w: 7, h: 0.55, fontSize: 28, bold: true, color: DARK, fontFace: "DM Sans" });
    s.addText(`Canal Email  ·  ${DATA.PERIODO_ACTUAL_LABEL || ""} vs ${DATA.PERIODO_ANTERIOR_LABEL || ""}`, { x: 0.5, y: 0.78, w: 9, h: 0.3, fontSize: 13, color: GRAY_TEXT, fontFace: "DM Sans" });

    s.addShape(pres.shapes.RECTANGLE, { x: 8.8, y: 0.2, w: 1.0, h: 0.35, fill: { color: LIGHT_BG }, line: { color: "F0E8E0", width: 0.5 } });
    s.addText("GA4", { x: 8.8, y: 0.2, w: 1.0, h: 0.35, fontSize: 10, bold: true, color: ORANGE, fontFace: "DM Sans", align: "center", valign: "middle" });

    const fmtPrev = v => (v && v !== "0" && v !== 0) ? v : "";
    const _ga4Ses  = parseNum(DATA.GA4_SESIONES);
    const _ga4Trx  = parseNum(DATA.GA4_TRANSACCIONES);
    const _ga4TCn  = _ga4Ses > 0 ? _ga4Trx / _ga4Ses * 100 : 0;
    const _ga4TC   = _ga4TCn > 0 ? `${_ga4TCn.toFixed(2).replace(".", ",")}%` : "";
    const _ga4SesP = parseNum(DATA.GA4_SESIONES_PREV);
    const _ga4TrxP = parseNum(DATA.GA4_TRANSACCIONES_PREV);
    const _ga4TCPn = _ga4SesP > 0 ? _ga4TrxP / _ga4SesP * 100 : 0;
    const _ga4TCP  = _ga4TCPn > 0 ? `${_ga4TCPn.toFixed(2).replace(".", ",")}%` : "";
    const _ga4TCDelta = (_ga4TCn > 0 && _ga4TCPn > 0)
      ? `${(_ga4TCn >= _ga4TCPn ? "+" : "")}${(_ga4TCn - _ga4TCPn).toFixed(2).replace(".", ",")}%`
      : "";
    const _ga4TCUp = _ga4TCn >= _ga4TCPn;
    const ga4Kpis = [
      { label: "Sesiones",          val: DATA.GA4_SESIONES      || "", delta: fmtDelta(DATA.GA4_SESIONES_DELTA), up: DATA.GA4_SESIONES_DELTA_UP      === true, prev: fmtPrev(DATA.GA4_SESIONES_PREV) },
      { label: "Transacciones",     val: DATA.GA4_TRANSACCIONES || "", delta: fmtDelta(DATA.GA4_TRANSACCIONES_DELTA), up: DATA.GA4_TRANSACCIONES_DELTA_UP === true, prev: fmtPrev(DATA.GA4_TRANSACCIONES_PREV) },
      { label: "Ingresos",          val: fmtMoneyCompact(DATA.GA4_INGRESOS), delta: fmtDelta(DATA.GA4_INGRESOS_DELTA), up: DATA.GA4_INGRESOS_DELTA_UP === true, prev: fmtPrev(DATA.GA4_INGRESOS_PREV) },
      { label: "Tasa de conversión",val: _ga4TC, delta: _ga4TCDelta, up: _ga4TCUp, prev: _ga4TCP },
    ];

    const ga4CardW = 2.1;
    ga4Kpis.forEach((k, i) => {
      const x = 0.4 + i * (ga4CardW + 0.22);
      const y = 1.2;
      s.addShape(pres.shapes.RECTANGLE, { x, y, w: ga4CardW, h: 2.0, fill: { color: LIGHT_BG }, line: { color: "F0E8E0", width: 0.5 } });
      s.addShape(pres.shapes.RECTANGLE, { x, y, w: ga4CardW, h: 0.07, fill: { color: ORANGE }, line: { color: ORANGE } });
      s.addText(k.label, { x, y: y + 0.15, w: ga4CardW, h: 0.3,  fontSize: 9.5, color: GRAY_TEXT, fontFace: "DM Sans",       align: "center" });
      s.addText(k.val,   { x, y: y + 0.48, w: ga4CardW, h: 0.65, fontSize: 26,  bold: true, color: DARK, fontFace: "DM Sans", align: "center" });
      s.addShape(pres.shapes.RECTANGLE, { x: x + 0.35, y: y + 1.18, w: ga4CardW - 0.7, h: 0.27, fill: { color: k.up ? GREEN_BG : RED_BG }, line: { color: k.up ? GREEN_BG : RED_BG } });
      s.addText(k.delta, { x: x + 0.35, y: y + 1.18, w: ga4CardW - 0.7, h: 0.27, fontSize: 11, bold: true, color: k.up ? GREEN : RED, fontFace: "DM Sans", align: "center" });
      s.addText(`${DATA.PERIODO_ANTERIOR_LABEL || "Ant."}: ${k.prev}`, { x, y: y + 1.55, w: ga4CardW, h: 0.22, fontSize: 8.5, color: GRAY_TEXT, fontFace: "DM Sans", align: "center" });
    });

    if (DATA.GA4_NOTA) {
      s.addShape(pres.shapes.RECTANGLE, { x: 0.4, y: 3.55, w: 9.2, h: 0.7, fill: { color: LIGHT_BG }, line: { color: "F0E8E0", width: 0.5 } });
      s.addText(DATA.GA4_NOTA, { x: 0.55, y: 3.6, w: 9.0, h: 0.6, fontSize: 10, color: ORANGE, fontFace: "DM Sans", wrap: true });
    }
  }

  // ── SLIDE – EVOLUTIVO (condicional — si DATA.EVOLUTIVO_ROWS existe) ─────
  if (Array.isArray(DATA.EVOLUTIVO_ROWS) && DATA.EVOLUTIVO_ROWS.length > 0) {
    const rows = DATA.EVOLUTIVO_ROWS;

    // Pivot: extraer meses únicos (en orden de aparición) y KPIs únicos
    const meses = [...new Set(rows.map(r => r.Mes || r.mes || ""))].filter(Boolean);
    const kpis  = [...new Set(rows.map(r => r.KPI || r.kpi || ""))].filter(Boolean);

    // Mapa [kpi][mes] = valor
    const map = {};
    rows.forEach(r => {
      const k = r.KPI || r.kpi || "";
      const m = r.Mes || r.mes || "";
      const v = r.Valor || r.valor || "—";
      if (!map[k]) map[k] = {};
      map[k][m] = v;
    });

    const s = pres.addSlide();
    s.background = { color: WHITE };

    s.addText("Evolución de Resultados", { x: 0.5, y: 0.18, w: 9, h: 0.52, fontSize: 26, bold: true, color: DARK, fontFace: "DM Sans" });
    s.addText(`${DATA.CLIENTE_NOMBRE || ""}  ·  ${DATA.PLATAFORMA_EMAIL || "Email Marketing"}`, {
      x: 0.5, y: 0.7, w: 9, h: 0.28, fontSize: 12, color: GRAY_TEXT, fontFace: "DM Sans",
    });

    // Layout dinámico: columna KPI + una columna por mes
    const kpiColW  = 2.2;
    const mesColW  = Math.min(1.6, (9.2 - kpiColW) / meses.length);
    const totalW   = kpiColW + mesColW * meses.length;
    const marginX  = (10 - totalW) / 2;
    const tY       = 1.1;
    const rowH     = 0.38;
    const lastMes  = meses[meses.length - 1]; // mes más reciente → destacado

    // Header: KPI + meses
    s.addShape(pres.shapes.RECTANGLE, { x: marginX, y: tY, w: totalW, h: 0.34, fill: { color: ORANGE }, line: { color: ORANGE } });
    s.addText("KPI", { x: marginX + 0.1, y: tY + 0.04, w: kpiColW - 0.1, h: 0.26, fontSize: 9, bold: true, color: WHITE, fontFace: "DM Sans", valign: "middle" });
    meses.forEach((mes, mi) => {
      const cx = marginX + kpiColW + mi * mesColW;
      const isLast = mes === lastMes;
      if (isLast) s.addShape(pres.shapes.RECTANGLE, { x: cx, y: tY, w: mesColW, h: 0.34, fill: { color: "D94E10" }, line: { color: "D94E10" } });
      s.addText(mes, { x: cx, y: tY + 0.04, w: mesColW, h: 0.26, fontSize: 9, bold: true, color: WHITE, fontFace: "DM Sans", align: "center", valign: "middle" });
    });

    // Filas por KPI
    kpis.forEach((kpi, ki) => {
      const ry = tY + 0.34 + ki * rowH;
      const bg = ki % 2 === 0 ? WHITE : "FAFAFA";
      s.addShape(pres.shapes.RECTANGLE, { x: marginX, y: ry, w: totalW, h: rowH, fill: { color: bg }, line: { color: "EEEEEE", width: 0.3 } });

      s.addText(kpi, { x: marginX + 0.1, y: ry + 0.08, w: kpiColW - 0.15, h: rowH - 0.1, fontSize: 9, bold: true, color: DARK, fontFace: "DM Sans", valign: "middle" });

      meses.forEach((mes, mi) => {
        const cx  = marginX + kpiColW + mi * mesColW;
        const val = (map[kpi] && map[kpi][mes]) ? map[kpi][mes] : "—";
        const isLast = mes === lastMes;
        if (isLast) s.addShape(pres.shapes.RECTANGLE, { x: cx, y: ry, w: mesColW, h: rowH, fill: { color: "FFF0E8" }, line: { color: "EEEEEE", width: 0.3 } });
        s.addText(val, { x: cx, y: ry + 0.08, w: mesColW, h: rowH - 0.1, fontSize: 9, bold: isLast, color: isLast ? ORANGE : DARK, fontFace: "DM Sans", align: "center", valign: "middle" });
      });
    });
  }

  // ── SLIDE – RECOMENDACIONES ───────────────────────────────────────────────
  buildSlide_Recommendations(pres, DATA);

  // ── SLIDE – CIERRE ────────────────────────────────────────────────────────
  buildSlide_Close(pres, DATA);

  return await pres.write({ outputType: "base64" });
}
