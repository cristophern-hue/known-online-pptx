const XLSX = require("xlsx");
const Anthropic = require("@anthropic-ai/sdk");
const pptxgen = require("pptxgenjs");

const client = new Anthropic({ apiKey: process.env.ANTHROPIC_API_KEY });

// ── Parse multipart form-data (raw) ──────────────────────────────────────────
function parseMultipart(req) {
  return new Promise((resolve, reject) => {
    const chunks = [];
    req.on("data", c => chunks.push(c));
    req.on("end", () => {
      const body = Buffer.concat(chunks);
      const ct = req.headers["content-type"] || "";
      const boundaryMatch = ct.match(/boundary=([^\s;]+)/);
      if (!boundaryMatch) return reject(new Error("No boundary in content-type"));
      const boundary = Buffer.from("--" + boundaryMatch[1]);
      const parts = [];
      let start = 0;
      while (true) {
        const idx = body.indexOf(boundary, start);
        if (idx === -1) break;
        const partStart = idx + boundary.length + 2;
        const nextIdx = body.indexOf(boundary, partStart);
        if (nextIdx === -1) break;
        const partData = body.slice(partStart, nextIdx - 2);
        const headerEnd = partData.indexOf("\r\n\r\n");
        if (headerEnd === -1) { start = nextIdx; continue; }
        const headers = partData.slice(0, headerEnd).toString();
        const data = partData.slice(headerEnd + 4);
        const nameMatch = headers.match(/name="([^"]+)"/);
        const filenameMatch = headers.match(/filename="([^"]+)"/);
        parts.push({
          name: nameMatch ? nameMatch[1] : "",
          filename: filenameMatch ? filenameMatch[1] : null,
          data
        });
        start = nextIdx;
      }
      resolve(parts);
    });
    req.on("error", reject);
  });
}

// ── Convert Excel buffer → CSV-like text for Claude ──────────────────────────
function excelToText(buffer) {
  const wb = XLSX.read(buffer, { type: "buffer" });
  const results = [];
  for (const sheetName of wb.SheetNames) {
    const ws = wb.Sheets[sheetName];
    const csv = XLSX.utils.sheet_to_csv(ws, { blankrows: false, skipHidden: true });
    if (csv.trim().length > 0) {
      results.push(`=== HOJA: ${sheetName} ===\n${csv}`);
    }
  }
  return results.join("\n\n");
}

// ── Ask Claude to extract structured data ────────────────────────────────────
async function extractDataWithClaude(sheetText) {
  const prompt = `Sos un analista de datos experto. Te paso el contenido CSV de un Excel de Cummins con métricas ecommerce.

ESTRUCTURA DEL ARCHIVO (puede variar en el futuro, adaptate):
- Fila 1: Título "PRIMER SEMESTRE (AÑO COMERCIAL XXXX)" → ese es el año ANTERIOR
- Fila 2: Vacía
- Fila 3: Encabezados de columnas (col A vacía, datos desde col B)
- Filas 4-9: Datos mensuales (Abril a Septiembre) del año anterior
- Filas vacías de separación
- Otra fila de título con el año ACTUAL
- Otra fila de encabezados
- Filas de datos del año actual (puede que solo haya algunos meses con datos)

COLUMNAS CONOCIDAS (el orden puede cambiar, buscalas por nombre aproximado):
- Sesiones: "Visitas/Session start" o similar
- Compras: "Cantidad compras (real)" o similar
- Venta Total, Venta P&F, Venta Motores, Venta Pgen/PGen/Pgeneradoras
- Inversión SEM Total (suma de todas las inversiones SEM)
- ROAS Total, ROAS P&F, ROAS Motores, ROAS Generadores
- Tasa de Conversión
- Ticket Promedio
- Añadidos al carrito, Carros abandonados
- Cantidad leads / Solicitar Cotización
- Inicio de compras

REGLAS:
1. Ignorá celdas con #DIV/0!, #¡DIV/0!, #N/A, #ERROR, o vacías → null
2. Números en formato argentino (1.234,56) → convertí a float (1234.56)
3. Porcentajes como "0.16%" → convertí a 0.16 (como número, no como fracción)
4. Si una columna no existe en el futuro, ponés null en todos sus valores
5. Si hay más o menos meses de los esperados, adaptate a lo que hay
6. El array "meses" debe incluir TODOS los meses de ambos bloques (en orden)
7. Si una métrica no existe pero hay datos similares, mapeala al campo más cercano

Devolvé SOLO JSON válido (sin texto extra, sin markdown) con esta estructura:
{
  "ano_anterior": "2025",
  "ano_actual": "2026",
  "meses": ["Abril","Mayo","Junio","Julio","Agosto","Septiembre"],
  "anterior": {
    "sesiones": [number_o_null, ...],
    "compras": [number_o_null, ...],
    "venta_total": [number_o_null, ...],
    "venta_pf": [number_o_null, ...],
    "venta_motores": [number_o_null, ...],
    "venta_pgen": [number_o_null, ...],
    "inversion_sem": [number_o_null, ...],
    "roas_total": [number_o_null, ...],
    "roas_pf": [number_o_null, ...],
    "roas_motores": [number_o_null, ...],
    "roas_gen": [number_o_null, ...],
    "tasa_conversion": [number_o_null, ...],
    "ticket_promedio": [number_o_null, ...],
    "carrito_inicio": [number_o_null, ...],
    "carrito_abandonado": [number_o_null, ...],
    "leads": [number_o_null, ...]
  },
  "actual": {
    "sesiones": [number_o_null, ...],
    "compras": [number_o_null, ...],
    "venta_total": [number_o_null, ...],
    "venta_pf": [number_o_null, ...],
    "venta_motores": [number_o_null, ...],
    "venta_pgen": [number_o_null, ...],
    "inversion_sem": [number_o_null, ...],
    "roas_total": [number_o_null, ...],
    "roas_pf": [number_o_null, ...],
    "roas_motores": [number_o_null, ...],
    "roas_gen": [number_o_null, ...],
    "tasa_conversion": [number_o_null, ...],
    "ticket_promedio": [number_o_null, ...],
    "carrito_inicio": [number_o_null, ...],
    "carrito_abandonado": [number_o_null, ...],
    "leads": [number_o_null, ...]
  },
  "columnas_encontradas": ["lista de columnas que encontraste en el archivo"]
}

Contenido del Excel:
${sheetText.slice(0, 14000)}`;

  const msg = await client.messages.create({
    model: "claude-sonnet-4-5",
    max_tokens: 2000,
    messages: [{ role: "user", content: prompt }]
  });

  const text = msg.content[0].text.trim();
  const s = text.indexOf("{");
  const e = text.lastIndexOf("}");
  if (s === -1 || e === -1) throw new Error("Claude no devolvió JSON válido");
  return JSON.parse(text.slice(s, e + 1));
}

// ── Build PPTX with bar charts ────────────────────────────────────────────────
async function buildChartsPptx(data) {
  const pres = new pptxgen();
  pres.layout = "LAYOUT_16x9";

  const ORANGE  = "FA5A1E";
  const BLUE    = "185FA5";
  const DARK    = "1A1A2E";
  const WHITE   = "FFFFFF";
  const GRAY    = "64748B";
  const LIGHT   = "F8F9FA";

  const meses    = data.meses;
  const ant      = data.anterior;
  const act      = data.actual;
  const anoAnt   = data.ano_anterior;
  const anoAct   = data.ano_actual;

  // Filter only months with actual data in 'actual'
  const mesesConDatos = meses.filter((_, i) => act.compras[i] !== null);

  function addSlideHeader(slide, title, subtitle) {
    slide.background = { color: WHITE };
    slide.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.08, fill: { color: ORANGE }, line: { color: ORANGE } });
    slide.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0.08, w: 0.08, h: 5.54, fill: { color: ORANGE }, line: { color: ORANGE } });
    slide.addText(title,    { x: 0.3, y: 0.2,  w: 9.4, h: 0.5, fontSize: 22, bold: true, color: DARK,  fontFace: "Trebuchet MS" });
    slide.addText(subtitle, { x: 0.3, y: 0.72, w: 9.4, h: 0.28, fontSize: 11, color: GRAY, fontFace: "DM Sans" });
  }

  function nonNull(arr) {
    return arr.map(v => (v === null || v === undefined ? 0 : v));
  }

  // ── SLIDE 1 – COVER ──────────────────────────────────────────────────────
  const cover = pres.addSlide();
  cover.background = { color: DARK };
  cover.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 0.18, h: 5.625, fill: { color: ORANGE }, line: { color: ORANGE } });
  cover.addShape(pres.shapes.OVAL, { x: 7.5, y: -1.0, w: 4.0, h: 4.0, fill: { color: ORANGE, transparency: 88 }, line: { color: ORANGE, transparency: 88 } });
  cover.addText("Known Online", { x: 0.5, y: 0.48, w: 4, h: 0.4, fontSize: 14, color: ORANGE, bold: true, fontFace: "DM Sans" });
  cover.addShape(pres.shapes.RECTANGLE, { x: 0.45, y: 1.4, w: 2.0, h: 0.32, fill: { color: ORANGE }, line: { color: ORANGE } });
  cover.addText("CUMMINS", { x: 0.45, y: 1.4, w: 2.0, h: 0.32, fontSize: 10, color: WHITE, bold: true, fontFace: "DM Sans", align: "center" });
  cover.addText("Reporte\nEcommerce", { x: 0.45, y: 1.85, w: 7, h: 1.4, fontSize: 52, color: WHITE, bold: true, fontFace: "Trebuchet MS", valign: "top" });
  cover.addText(`1er Semestre ${anoAct}  vs.  ${anoAnt}`, { x: 0.45, y: 3.4, w: 7, h: 0.4, fontSize: 16, color: ORANGE, fontFace: "DM Sans" });
  cover.addText(`Gráficos de performance · Generado ${new Date().toLocaleDateString("es-AR")}`, { x: 0.45, y: 4.2, w: 8, h: 0.3, fontSize: 11, color: GRAY, fontFace: "DM Sans" });

  // ── SLIDE 2 – VENTAS TOTALES POR MES ─────────────────────────────────────
  const s2 = pres.addSlide();
  addSlideHeader(s2, "Venta Total por Mes", `Comparativo ${anoAnt} vs ${anoAct}  ·  en USD`);

  s2.addChart(pres.charts.BAR, [
    { name: anoAnt, labels: meses, values: nonNull(ant.venta_total) },
    { name: anoAct, labels: meses, values: nonNull(act.venta_total) }
  ], {
    x: 0.3, y: 1.1, w: 9.4, h: 4.2,
    barDir: "col", barGrouping: "clustered",
    chartColors: [ORANGE, BLUE],
    showLegend: true, legendPos: "t", legendFontSize: 10,
    showValue: true, dataLabelFontSize: 9,
    catAxisLabelFontSize: 10, valAxisLabelFontSize: 9,
    valAxisNumFmt: "$#,##0",
    shadow: { type: "none" }
  });

  // ── SLIDE 3 – VENTAS POR LÍNEA DE NEGOCIO ────────────────────────────────
  const s3 = pres.addSlide();
  addSlideHeader(s3, "Ventas por Línea de Negocio", `${anoAnt} vs ${anoAct}  ·  P&F / Motores / PGen`);

  // Use only months with data in 'actual'
  const idxActual = meses.map((_, i) => act.venta_pf[i] !== null ? i : -1).filter(i => i >= 0);
  const mesesSlice = idxActual.length > 0 ? idxActual.map(i => meses[i]) : meses;

  s3.addChart(pres.charts.BAR, [
    { name: `P&F ${anoAnt}`,      labels: meses, values: nonNull(ant.venta_pf) },
    { name: `Motores ${anoAnt}`,  labels: meses, values: nonNull(ant.venta_motores) },
    { name: `PGen ${anoAnt}`,     labels: meses, values: nonNull(ant.venta_pgen) },
    { name: `P&F ${anoAct}`,      labels: meses, values: nonNull(act.venta_pf) },
  ], {
    x: 0.3, y: 1.1, w: 9.4, h: 4.2,
    barDir: "col", barGrouping: "stacked",
    chartColors: ["FA5A1E", "185FA5", "3B6D11", "FF912D"],
    showLegend: true, legendPos: "t", legendFontSize: 9,
    showValue: false,
    catAxisLabelFontSize: 10, valAxisLabelFontSize: 9,
    valAxisNumFmt: "$#,##0",
    shadow: { type: "none" }
  });

  // ── SLIDE 4 – ROAS POR LÍNEA ─────────────────────────────────────────────
  const s4 = pres.addSlide();
  addSlideHeader(s4, "ROAS por Línea de Negocio", `${anoAnt} vs ${anoAct}  ·  Total / P&F / Motores / Generadores`);

  s4.addChart(pres.charts.BAR, [
    { name: `ROAS Total ${anoAnt}`,  labels: meses, values: nonNull(ant.roas_total) },
    { name: `ROAS P&F ${anoAnt}`,    labels: meses, values: nonNull(ant.roas_pf) },
    { name: `ROAS Gen ${anoAnt}`,    labels: meses, values: nonNull(ant.roas_gen) },
    { name: `ROAS Total ${anoAct}`,  labels: meses, values: nonNull(act.roas_total) },
    { name: `ROAS P&F ${anoAct}`,    labels: meses, values: nonNull(act.roas_pf) },
  ], {
    x: 0.3, y: 1.1, w: 9.4, h: 4.2,
    barDir: "col", barGrouping: "clustered",
    chartColors: ["FA5A1E", "FF912D", "3B6D11", "185FA5", "5BA4F5"],
    showLegend: true, legendPos: "t", legendFontSize: 9,
    showValue: true, dataLabelFontSize: 8,
    catAxisLabelFontSize: 10, valAxisLabelFontSize: 9,
    valAxisNumFmt: "0.00",
    shadow: { type: "none" }
  });

  // ── SLIDE 5 – FUNNEL ECOMMERCE ───────────────────────────────────────────
  const s5 = pres.addSlide();
  addSlideHeader(s5, "Funnel Ecommerce por Mes", `${anoAnt} vs ${anoAct}  ·  Inicio → Carrito → Compras`);

  s5.addChart(pres.charts.BAR, [
    { name: `Inicio compra ${anoAnt}`,   labels: meses, values: nonNull(ant.carrito_inicio) },
    { name: `Carrito aband. ${anoAnt}`,  labels: meses, values: nonNull(ant.carrito_abandonado) },
    { name: `Compras ${anoAnt}`,         labels: meses, values: nonNull(ant.compras) },
    { name: `Inicio compra ${anoAct}`,   labels: meses, values: nonNull(act.carrito_inicio) },
    { name: `Compras ${anoAct}`,         labels: meses, values: nonNull(act.compras) },
  ], {
    x: 0.3, y: 1.1, w: 9.4, h: 4.2,
    barDir: "col", barGrouping: "clustered",
    chartColors: ["FA5A1E", "FFB899", "3B6D11", "185FA5", "5BA4F5"],
    showLegend: true, legendPos: "t", legendFontSize: 9,
    showValue: true, dataLabelFontSize: 8,
    catAxisLabelFontSize: 10, valAxisLabelFontSize: 9,
    shadow: { type: "none" }
  });

  // ── SLIDE 6 – TRÁFICO Y SESIONES ─────────────────────────────────────────
  const s6 = pres.addSlide();
  addSlideHeader(s6, "Tráfico del Sitio", `${anoAnt} vs ${anoAct}  ·  Sesiones / Leads / Tasa de conversión`);

  // Bar chart for sesiones + leads
  s6.addChart(pres.charts.BAR, [
    { name: `Sesiones ${anoAnt}`,  labels: meses, values: nonNull(ant.sesiones) },
    { name: `Leads ${anoAnt}`,     labels: meses, values: nonNull(ant.leads) },
    { name: `Sesiones ${anoAct}`,  labels: meses, values: nonNull(act.sesiones) },
    { name: `Leads ${anoAct}`,     labels: meses, values: nonNull(act.leads) },
  ], {
    x: 0.3, y: 1.1, w: 6.0, h: 4.2,
    barDir: "col", barGrouping: "clustered",
    chartColors: ["FA5A1E", "FFB899", "185FA5", "5BA4F5"],
    showLegend: true, legendPos: "t", legendFontSize: 9,
    showValue: false,
    catAxisLabelFontSize: 9, valAxisLabelFontSize: 9,
    shadow: { type: "none" }
  });

  // Line chart for tasa de conversión
  s6.addChart(pres.charts.LINE, [
    { name: `Conv% ${anoAnt}`, labels: meses, values: nonNull(ant.tasa_conversion) },
    { name: `Conv% ${anoAct}`, labels: meses, values: nonNull(act.tasa_conversion) },
  ], {
    x: 6.5, y: 1.1, w: 3.2, h: 4.2,
    chartColors: [ORANGE, BLUE],
    showLegend: true, legendPos: "t", legendFontSize: 9,
    showValue: true, dataLabelFontSize: 8,
    catAxisLabelFontSize: 9, valAxisLabelFontSize: 9,
    valAxisNumFmt: "0.00%",
    shadow: { type: "none" },
    lineDataSymbol: "circle", lineDataSymbolSize: 6
  });

  // ── SLIDE 7 – INVERSIÓN SEM ──────────────────────────────────────────────
  const s7 = pres.addSlide();
  addSlideHeader(s7, "Inversión SEM vs Venta Total", `${anoAnt} vs ${anoAct}  ·  Eficiencia del gasto publicitario`);

  s7.addChart(pres.charts.BAR, [
    { name: `Inversión ${anoAnt}`,   labels: meses, values: nonNull(ant.inversion_sem) },
    { name: `Venta Total ${anoAnt}`, labels: meses, values: nonNull(ant.venta_total) },
    { name: `Inversión ${anoAct}`,   labels: meses, values: nonNull(act.inversion_sem) },
    { name: `Venta Total ${anoAct}`, labels: meses, values: nonNull(act.venta_total) },
  ], {
    x: 0.3, y: 1.1, w: 9.4, h: 4.2,
    barDir: "col", barGrouping: "clustered",
    chartColors: ["FFB899", "FA5A1E", "5BA4F5", "185FA5"],
    showLegend: true, legendPos: "t", legendFontSize: 9,
    showValue: false,
    catAxisLabelFontSize: 10, valAxisLabelFontSize: 9,
    valAxisNumFmt: "$#,##0",
    shadow: { type: "none" }
  });

  // ── SLIDE 8 – TICKET PROMEDIO ─────────────────────────────────────────────
  const s8 = pres.addSlide();
  addSlideHeader(s8, "Ticket Promedio por Mes", `${anoAnt} vs ${anoAct}  ·  Valor promedio por compra en USD`);

  s8.addChart(pres.charts.BAR, [
    { name: anoAnt, labels: meses, values: nonNull(ant.ticket_promedio) },
    { name: anoAct, labels: meses, values: nonNull(act.ticket_promedio) },
  ], {
    x: 0.3, y: 1.1, w: 9.4, h: 4.2,
    barDir: "col", barGrouping: "clustered",
    chartColors: [ORANGE, BLUE],
    showLegend: true, legendPos: "t", legendFontSize: 10,
    showValue: true, dataLabelFontSize: 10,
    catAxisLabelFontSize: 10, valAxisLabelFontSize: 9,
    valAxisNumFmt: "$#,##0",
    shadow: { type: "none" }
  });

  // ── SLIDE 9 – CIERRE ──────────────────────────────────────────────────────
  const close = pres.addSlide();
  close.background = { color: ORANGE };
  close.addShape(pres.shapes.OVAL, { x: 6.5, y: -1.5, w: 5.5, h: 5.5, fill: { color: WHITE, transparency: 92 }, line: { color: WHITE, transparency: 92 } });
  close.addShape(pres.shapes.OVAL, { x: -2.0, y: 3.0, w: 4.5, h: 4.5, fill: { color: DARK, transparency: 88 }, line: { color: DARK, transparency: 88 } });
  close.addText("Known Online", { x: 1.2, y: 0.52, w: 4, h: 0.45, fontSize: 16, bold: true, color: WHITE, fontFace: "DM Sans" });
  close.addText("¡Gracias!", { x: 0.5, y: 1.6, w: 9, h: 1.4, fontSize: 56, bold: true, color: WHITE, fontFace: "Trebuchet MS", align: "center" });
  close.addText("Logramos tu transformación digital", { x: 0.5, y: 3.1, w: 9, h: 0.45, fontSize: 18, color: "FFD4B8", fontFace: "DM Sans", align: "center", italic: true });
  close.addText("www.knownonline.com", { x: 0.5, y: 3.9, w: 9, h: 0.35, fontSize: 14, color: WHITE, fontFace: "DM Sans", align: "center", bold: true });

  return pres.write({ outputType: "base64" });
}

// ── Main handler ──────────────────────────────────────────────────────────────
module.exports = async function handler(req, res) {
  res.setHeader("Access-Control-Allow-Origin", "*");
  res.setHeader("Access-Control-Allow-Methods", "POST, OPTIONS");
  res.setHeader("Access-Control-Allow-Headers", "Content-Type");

  if (req.method === "OPTIONS") return res.status(200).end();
  if (req.method !== "POST") return res.status(405).json({ error: "Use POST" });

  if (!process.env.ANTHROPIC_API_KEY) {
    return res.status(500).json({ error: "ANTHROPIC_API_KEY no configurada en Vercel" });
  }

  try {
    const parts = await parseMultipart(req);
    const filePart = parts.find(p => p.filename);
    if (!filePart) return res.status(400).json({ error: "No se encontró archivo en el request" });

    const ext = (filePart.filename || "").split(".").pop().toLowerCase();
    if (!["xlsx", "xls", "csv"].includes(ext)) {
      return res.status(400).json({ error: "Formato no soportado. Usá .xlsx, .xls o .csv" });
    }

    const sheetText = excelToText(filePart.data);
    const data = await extractDataWithClaude(sheetText);
    const base64 = await buildChartsPptx(data);

    const filename = `Cummins_Ecommerce_${data.ano_actual}_vs_${data.ano_anterior}.pptx`;
    return res.status(200).json({ pptx: base64, filename });

  } catch (err) {
    console.error(err);
    return res.status(500).json({ error: err.message || "Error generando el reporte" });
  }
};
