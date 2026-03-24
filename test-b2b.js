const fs = require("fs");
const path = require("path");

// Simular req/res para llamar al handler
const DATA = {
  CLIENTE_NOMBRE: "Acme Industrial",
  PERIODO_ACTUAL_LABEL: "Enero 2026",
  PERIODO_ANTERIOR_LABEL: "Enero 2025",
  PERIODO_ACTUAL_SHORT: "Ene '26",
  PERIODO_ANTERIOR_SHORT: "Ene '25",
  mes_actual: "1",
  año_actual: "2026",

  // Inversión total
  INVERSION_TOTAL: "$4.250.000",
  INVERSION_PREV: "$3.800.000",
  INVERSION_DELTA: "+11,8%",
  INVERSION_DELTA_UP: true,

  // Leads
  LEADS_TOTAL: "842",
  LEADS_PREV: "710",
  LEADS_DELTA: "+18,6%",
  LEADS_DELTA_UP: true,

  // CPL
  CPL_TOTAL: "$5.047",
  CPL_PREV: "$5.352",
  CPL_DELTA: "-5,7%",
  CPL_DELTA_UP: false,

  // Clicks
  CLICKS_TOTAL: "38.450",
  CLICKS_PREV: "31.200",
  CLICKS_DELTA: "+23,2%",
  CLICKS_DELTA_UP: true,

  // GA4
  GA4_SESIONES: "24.800",
  GA4_SESIONES_PREV: "21.300",
  GA4_SESIONES_DELTA: "+16,4%",
  GA4_SESIONES_DELTA_UP: true,

  GA4_LEADS: "842",
  GA4_LEADS_PREV: "710",
  GA4_LEADS_DELTA: "+18,6%",
  GA4_LEADS_DELTA_UP: true,

  GA4_CPL: "$5.047",
  GA4_CPL_PREV: "$5.352",
  GA4_CPL_DELTA: "-5,7%",
  GA4_CPL_DELTA_UP: false,

  GA4_CONV_RATE: "3,40%",
  GA4_CONV_RATE_PREV: "3,33%",
  GA4_CONV_RATE_DELTA: "+0,1pp",
  GA4_CONV_RATE_DELTA_UP: true,

  GA4_TIEMPO_SITIO: "2:34",
  GA4_TIEMPO_SITIO_PREV: "2:18",
  GA4_TIEMPO_SITIO_DELTA: "+11,6%",
  GA4_TIEMPO_SITIO_DELTA_UP: true,

  GA4_INSIGHT: "Las sesiones orgánicas crecieron 22% impulsadas por contenido SEO publicado en Q4, reforzando la performance general del canal.",

  // META
  META_COSTO: "$2.100.000",
  META_COSTO_PREV: "$1.850.000",
  META_COSTO_DELTA: "+13,5%",
  META_COSTO_DELTA_UP: true,

  META_CLICKS: "22.400",
  META_CLICKS_PREV: "18.700",
  META_CLICKS_DELTA: "+19,8%",
  META_CLICKS_DELTA_UP: true,

  META_IMPRESIONES: "1.240.000",
  META_IMPRESIONES_PREV: "980.000",
  META_IMPRESIONES_DELTA: "+26,5%",
  META_IMPRESIONES_DELTA_UP: true,

  META_CTR: "1,81%",
  META_CTR_PREV: "1,91%",
  META_CTR_DELTA: "-0,1pp",
  META_CTR_DELTA_UP: false,

  META_CPC: "$93,75",
  META_CPC_PREV: "$98,93",
  META_CPC_DELTA: "-5,2%",
  META_CPC_DELTA_UP: false,

  META_CPL: "$5.893",
  META_CPL_PREV: "$6.412",
  META_CPL_DELTA: "-8,1%",
  META_CPL_DELTA_UP: false,

  META_ALERTA: "Meta muestra una mejora sostenida en CPL (-8,1%) con mayor volumen de impresiones. Se recomienda escalar presupuesto en los adsets con menor CPL.",

  // GOOGLE
  GOOGLE_COSTO: "$2.150.000",
  GOOGLE_COSTO_PREV: "$1.950.000",
  GOOGLE_COSTO_DELTA: "+10,3%",
  GOOGLE_COSTO_DELTA_UP: true,

  GOOGLE_CLICKS: "16.050",
  GOOGLE_CLICKS_PREV: "12.500",
  GOOGLE_CLICKS_DELTA: "+28,4%",
  GOOGLE_CLICKS_DELTA_UP: true,

  GOOGLE_IMPRESIONES: "310.000",
  GOOGLE_IMPRESIONES_PREV: "270.000",
  GOOGLE_IMPRESIONES_DELTA: "+14,8%",
  GOOGLE_IMPRESIONES_DELTA_UP: true,

  GOOGLE_CTR: "5,18%",
  GOOGLE_CTR_PREV: "4,63%",
  GOOGLE_CTR_DELTA: "+0,55pp",
  GOOGLE_CTR_DELTA_UP: true,

  GOOGLE_CPC: "$134,02",
  GOOGLE_CPC_PREV: "$156,00",
  GOOGLE_CPC_DELTA: "-14,1%",
  GOOGLE_CPC_DELTA_UP: false,

  GOOGLE_CPL: "$4.316",
  GOOGLE_CPL_PREV: "$4.875",
  GOOGLE_CPL_DELTA: "-11,5%",
  GOOGLE_CPL_DELTA_UP: false,

  GOOGLE_ALERTA: "Google Ads logró su mejor CTR histórico (5,18%) con un CPC 14% más eficiente. Performance Max y Search capturan demanda calificada en todas las etapas del funnel.",

  // Top campañas
  CAMPANAS: [
    { nombre: "Search | Maquinaria Industrial | Exact", plataforma: "Google", costo: "$820.000", leads: "194", cpl: "$4.227", nivel: "low" },
    { nombre: "Search | Repuestos | Phrase",            plataforma: "Google", costo: "$540.000", leads: "112", cpl: "$4.821", nivel: "low" },
    { nombre: "Performance Max | Catálogo 2026",        plataforma: "Google", costo: "$490.000", leads: "86",  cpl: "$5.698", nivel: "mid" },
    { nombre: "Remarketing | Leads Calientes",          plataforma: "Meta",   costo: "$380.000", leads: "68",  cpl: "$5.588", nivel: "mid" },
    { nombre: "Lead Gen | Industria | Lookalike 2%",    plataforma: "Meta",   costo: "$620.000", leads: "104", cpl: "$5.961", nivel: "mid" },
    { nombre: "Lead Gen | Ingenieros | Intereses",      plataforma: "Meta",   costo: "$540.000", leads: "72",  cpl: "$7.500", nivel: "high" },
    { nombre: "PMAX | Nuevos Mercados",                 plataforma: "Google", costo: "$300.000", leads: "36",  cpl: "$8.333", nivel: "high" },
    { nombre: "Branding | Awareness Video 30s",         plataforma: "Meta",   costo: "$560.000", leads: "41",  cpl: "$13.658", nivel: "high" },
  ],

  // Top anuncios Meta
  TOP_ANUNCIOS_META_TIENE_DATOS: true,
  TOP_ANUNCIOS_META: [
    { nombre: "Maquinaria Industrial | Video testimonial cliente minero | Ene 2026", leads: "68", cpl: "$5.588", costo: "$380.000", clicks: "4.200", preview_url: null },
    { nombre: "Lead Gen | Lookalike 2% | Carousel maquinaria pesada | Ene 2026",     leads: "54", cpl: "$6.111", costo: "$330.000", clicks: "3.800", preview_url: null },
    { nombre: "Remarketing | Video corto 15s | Visitantes 30 días | Ene 2026",       leads: "32", cpl: "$6.875", costo: "$220.000", clicks: "2.100", preview_url: null },
  ],

  // Top canales GA4
  FUENTE_MEDIO: [
    { nombre: "google / cpc",           sesiones: "14.200", txns: "497", tc: "3,50%", tc_prev: "3,10%", tc_delta: "+0,40pp", tc_delta_up: true,  revenue: "$2.150.000", revenue_prev: "$1.950.000", revenue_delta: "+10,3%", revenue_delta_up: true  },
    { nombre: "facebook / cpc",         sesiones: "6.800",  txns: "198", tc: "2,91%", tc_prev: "3,20%", tc_delta: "-0,29pp", tc_delta_up: false, revenue: "$1.100.000", revenue_prev: "$980.000",   revenue_delta: "+12,2%", revenue_delta_up: true  },
    { nombre: "google / organic",        sesiones: "1.900",  txns: "72",  tc: "3,79%", tc_prev: "3,10%", tc_delta: "+0,69pp", tc_delta_up: true,  revenue: "—",          revenue_prev: "—",          revenue_delta: "—",      revenue_delta_up: true  },
    { nombre: "instagram / cpc",         sesiones: "920",    txns: "28",  tc: "3,04%", tc_prev: "2,90%", tc_delta: "+0,14pp", tc_delta_up: true,  revenue: "—",          revenue_prev: "—",          revenue_delta: "—",      revenue_delta_up: true  },
    { nombre: "direct / none",           sesiones: "680",    txns: "18",  tc: "2,65%", tc_prev: "2,80%", tc_delta: "-0,15pp", tc_delta_up: false, revenue: "—",          revenue_prev: "—",          revenue_delta: "—",      revenue_delta_up: false },
  ],

  FUENTE_MEDIO_INSIGHT: "Google CPC mantiene la mayor tasa de conversión (3,50%) y concentra el 59% de las sesiones pagadas. Facebook baja 0,29pp en TC pero crece en revenue.",

  // Recomendaciones
  RECOMENDACIONES: [
    "Escalar presupuesto en Search Exact (CPL $4.227) que ya opera por debajo del objetivo de $5.000.",
    "Pausar campaña Branding Video 30s (CPL $13.658 — sin leads calificados en enero).",
    "Testear creative con testimonio de cliente industrial en Lookalike 2% para mejorar la TC de Meta.",
    "Incorporar retargeting de lista de clientes en Google para reducir CPL en segmentos ya conocidos.",
  ],

  // Cierre
  AGENCIA_NOMBRE: "Known Online",
  CONTACTO_EMAIL: "hola@knownonline.com",
};

async function main() {
  const handler = require("./api/generate-b2b");

  const req = {
    method: "POST",
    body: { DATA },
  };

  let statusCode, responseBody;
  const res = {
    status(code) { statusCode = code; return this; },
    json(body)   { responseBody = body; return this; },
  };

  await handler(req, res);

  if (statusCode !== 200) {
    console.error("Error:", responseBody);
    process.exit(1);
  }

  const buf = Buffer.from(responseBody.pptx, "base64");
  const outPath = path.join(__dirname, "test-b2b-output.pptx");
  fs.writeFileSync(outPath, buf);
  console.log(`✓ PPT generada: ${outPath}  (${(buf.length / 1024).toFixed(0)} KB)`);
}

main().catch(err => { console.error(err); process.exit(1); });
