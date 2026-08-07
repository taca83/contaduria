import React, { useState, useEffect, useMemo, useRef } from "react";
import * as XLSX from "xlsx";
import {
  PieChart, Pie, Cell, ResponsiveContainer, Tooltip,
  BarChart, Bar, XAxis, YAxis, CartesianGrid, Legend
} from "recharts";
import {
  Wallet, TrendingUp, TrendingDown, PiggyBank, Target, Plus, X,
  ArrowUpRight, ArrowDownRight, ArrowLeft, ArrowLeftRight, Landmark, Settings, Trash2, User, Download, Menu,
  Home, List, Upload, Pencil, Check, Mic, Square, Zap, Camera, Clock, Tag, Copy, History, LogOut
} from "lucide-react";

const DEFAULT_GASTO_CATS = ["Comida", "Tarjetas", "Ropa", "Salud", "Educación", "Transporte", "Ocio", "Servicios", "Vivienda", "Otros"];
const TAB_LABELS = {
  resumen: "Resumen",
  movimientos: "Movimientos",
  ahorros: "Ahorros e inversiones",
  presupuestos: "Presupuestos",
  rapidos: "Gastos rápidos",
  importar: "Importar",
  duplicados: "Duplicados",
  recategorizar: "Recategorizar",
  divisas: "Divisas",
  historial: "Historial de cambios",
  reset: "Reiniciar datos",
  hogar: "Mi hogar",
  categorias: "Categorías",
  recurrentes: "Gastos recurrentes",
  cuentas: "Cuentas",
  cotizaciones: "Cotizaciones",
  personas: "Por persona",
  conciliar: "Conciliar pagos",
  admin: "Panel admin (ingresos a la app)",
};
const PRIMARY_TABS = [
  ["resumen", "Resumen", Home],
  ["movimientos", "Movimientos", List],
  ["rapidos", "Gastos rápidos", Zap],
  ["importar", "Importar", Upload],
];
const SECONDARY_TABS = ["hogar", "categorias", "recurrentes", "presupuestos", "cuentas", "personas", "cotizaciones", "recategorizar", "conciliar", "duplicados", "divisas", "historial", "ahorros", "reset"];
const INGRESO_CATS = ["Sueldo", "Freelance", "Alquileres", "Otros ingresos"];
const AHORRO_INSTR = ["Plazo fijo", "Dólares (billete)", "FCI", "Acciones / CEDEARs", "Cripto", "Otro"];

const TEAL = "#0F6E6E";
const GOLD = "#C9A227";
const BRICK = "#B5473A";
const GREEN = "#2E7D4F";
const INK = "#1B2A2E";
const PAPER = "#F7F5F0";
const PAPER_DIM = "#EFEBE2";

const CAT_COLORS = ["#0F6E6E", "#C9A227", "#B5473A", "#2E7D4F", "#7A5CC7", "#3E7CB1", "#C97B3D", "#8A8F5C", "#9C6B9E"];

// Agrupa cualquier nombre de cuenta en un "medio de pago" macro (Efectivo,
// Mercado Pago, Tarjetas de crédito, Otros) — para no tener que mirar 5
// tarjetas sueltas en la pestaña Cuentas, solo el panorama general primero.
function clasificarMedioPago(cuenta) {
  const c = (cuenta || "").toLowerCase().trim();
  if (!c) return "Otros medios de pago";
  if (c.includes("efectivo")) return "Efectivo";
  if (c.includes("mercado pago") || c.includes("mercadopago")) return "Mercado Pago";
  if (c.includes("visa") || c.includes("master") || c.includes("amex") || c.includes("american express") || c.includes("tarjeta")) return "Tarjetas de crédito";
  return "Otros medios de pago";
}

// Etiqueta de tarjeta + titular, para separar "VISA Negro" de "VISA
// Nati" cuando ambos son adicionales de la misma cuenta BBVA. Usa el
// campo persona (que el parser de BBVA completa con el nombre real que
// trae el PDF); si un movimiento viejo no lo tiene, cae de respaldo al
// "(Natalia)" que se agregaba antes al final de la descripción.
function etiquetaTarjeta(entry) {
  const cuenta = (entry.account || "").trim() || "Sin cuenta especificada";
  // El ajuste de impuestos/tasas/intereses (y el "Total a pagar" de
  // respaldo cuando no se pudo itemizar) no son gasto de ningún titular
  // en particular — son un cargo de la tarjeta en sí. Sin esto, caían
  // bajo el mismo nombre que los consumos del titular principal,
  // mezclándose y sin poder distinguirse de un vistazo.
  if (/^(Impuestos y cargos de tarjeta|Total a pagar)\b/.test(entry.desc || "")) {
    return `${cuenta} · Impuestos`;
  }
  const esNatalia = entry.persona
    ? /natalia/i.test(entry.persona)
    : /\(Natalia\)\s*$/.test(entry.desc || "");
  // Si la cuenta ya se llama "Visa Signature Natalia" (porque es su
  // propia tarjeta, detectada por el destinatario del resumen), no hace
  // falta repetir el sufijo — sería "... Natalia · Natalia".
  if (esNatalia && !/natalia/i.test(cuenta)) return `${cuenta} · Natalia`;
  return cuenta;
}

// Subí este número cada vez que Claude te entregue un archivo nuevo.
// Sirve para confirmar de un vistazo que el deploy tomó la versión correcta.
// v126 · 2026-08-06 · dos pedidos: (1) nuevo modo "PDF (detecta el formato solo)" en Importar — sube uno o varios PDFs mezclados (BBVA/Mercado Pago/colegio) y cada uno se procesa con el parser que corresponde, sin elegirlo a mano; (2) nuevo modo "Captura de pantalla (ARQ, etc.)" que lee varios movimientos de una sola captura con IA (vía una Edge Function nueva de Supabase, analizar-movimientos — requiere desplegarla una vez, ver archivo aparte) — separa transferencias, ingresos y cambios de divisa automáticamente
const APP_VERSION = "v126 · 2026-08-06";

function fmtARS(n) {
  const v = Number(n) || 0;
  return v.toLocaleString("es-AR", { style: "currency", currency: "ARS", maximumFractionDigits: 0 });
}

// Unifica "negro" / "Negro" / "NEGRO" bajo un mismo nombre (mayúscula
// inicial, resto tal cual) — para que agrupar "por persona" no separe al
// mismo miembro del hogar en dos filas distintas por una diferencia de
// mayúsculas/minúsculas al cargar el nombre en algún momento.
function capitalizar(s) {
  const t = (s || "").trim();
  return t ? t.charAt(0).toUpperCase() + t.slice(1) : t;
}

// Los modales (editar movimiento, foto, voz, nuevo movimiento) se abren
// como "hoja pegada abajo" en el celu, pero en una pantalla ancha eso se
// ve raro flotando lejos del contenido — ahí los mostramos centrados,
// como un diálogo normal de escritorio.
function esPantallaAncha() {
  return typeof window !== "undefined" && window.innerWidth >= 900;
}

// Evalúa expresiones simples tipo calculadora en el campo de Monto
// ("1500+320", "8400/2"). Es un parser propio (+ - * / paréntesis) — nunca
// usa eval, y devuelve null si el texto no es una expresión numérica válida.
function evaluarExpresion(texto) {
  const limpio = (texto || "").replace(/,/g, ".").trim();
  if (!limpio) return null;
  if (!/^[\d+\-*/().\s]+$/.test(limpio)) return null;
  let i = 0;
  function parseExpr() {
    let v = parseTerm();
    while (limpio[i] === "+" || limpio[i] === "-") {
      const op = limpio[i]; i++;
      v = op === "+" ? v + parseTerm() : v - parseTerm();
    }
    return v;
  }
  function parseTerm() {
    let v = parseFactor();
    while (limpio[i] === "*" || limpio[i] === "/") {
      const op = limpio[i]; i++;
      v = op === "*" ? v * parseFactor() : v / parseFactor();
    }
    return v;
  }
  function parseFactor() {
    while (limpio[i] === " ") i++;
    if (limpio[i] === "(") {
      i++;
      const v = parseExpr();
      while (limpio[i] === " ") i++;
      if (limpio[i] === ")") i++;
      return v;
    }
    if (limpio[i] === "-") { i++; return -parseFactor(); }
    const start = i;
    while (i < limpio.length && /[\d.]/.test(limpio[i])) i++;
    if (start === i) throw new Error("expresión inválida");
    return parseFloat(limpio.slice(start, i));
  }
  try {
    while (limpio[i] === " ") i++;
    const resultado = parseExpr();
    while (limpio[i] === " ") i++;
    if (i !== limpio.length || !isFinite(resultado)) return null;
    return resultado;
  } catch {
    return null;
  }
}
// --- Categorización automática por palabra clave (mismas reglas que
// venimos aplicando a mano en los resúmenes de BBVA) ---
const CATEGORY_RULES = [
  ["Escuelas", [/\bORT\b/, "COMUNIDAD BETEL", "BETEL"]],
  ["Country/Hebraica", ["SOCIEDAD HEBRA", "SOCIEDADHEBRA", "HEBRAICA", "HEBRAIC"]],
  ["Seguros", ["ZURICH SEGUROS", "ALLIANZ", "BBVA SEGUROS", "BINA SEGUROS"]],
  ["Salud", ["OSDE", "FARMACIA", "FARMPLUS", "FARMACITY", "NUTRISUPLE", "ORTOPEDIA", "EYELIT", "IVESS"]],
  ["Ropa", ["ADIDAS", "DEXTER", "RAPANUI"]],
  ["Transporte", ["AUSOL", "CORREDORESVIALES", "AUBASA", "UBER", "PAYU*AR", "AUTOEQUIPE", "SHELL", "NAFTA"]],
  ["Servicios", ["CLARO DEB", "PERSONAL FLOW", "PERSONAL ", "EDENOR", "IPLAN", "ADOBE", "FIBER2HOME", "STRIX LOJACK", "SUSCRIPCION", "MCAFEE", "DLOCAL", "ROSMINO Y CIA"]],
  ["Ocio", ["BUQUEBUS", "GOL LINHAS", "TEATRO", "PLATEANET", "PIAF", "VELEZ", "GEXCORP"]],
  ["Comida", [
    "CARREFOUR", "COTO SUCURSAL", "SUPERMERCADOS DIA", "MC SUPERMERCADOS", "ALMAC",
    "MARKET", "LA CABRERA", "HAVANNA", "ARCOSDORADOS", "MCDONALD", "BURGER KING",
    "PEDIDOSYA", "DLO*PEDIDOSYA", " RES ", "RESTO", "CAFE", "COFFEE", "TOSTADO",
    "LUCCIANOS", "DAIU", "LA GRANJA", "GOUT GLUTEN", "SUSHI", "MEDIALUNA",
    "CONTINENTAL", "LA JUVENIL", "LOVINNE", "GUITARRITA", "BEECOFFEE", "CAFEMARTINEZ",
    "NANQUE", "ROYAL VENDING", "SUPER TESCO", "MASTER",
  ]],
];
// Clave normalizada para "aprender" categorías por descripción exacta:
// mayúsculas, sin espacios de más, y sin el sufijo de cuota (para que
// "MERPAGO*ADIDAS (cuota 3/6)" y "MERPAGO*ADIDAS (cuota 4/6)" cuenten
// como la misma descripción de fondo.
function descKey(desc) {
  return (desc || "")
    .toUpperCase()
    .replace(/\s*\(CUOTA[^)]*\)\s*/g, "")
    // El Nro. de Cupón / Nro. de operación es único por transacción — lo
    // sacamos acá para que las reglas de categorización aprendidas sigan
    // agrupando por comercio (ej. todas las de "MERPAGO*AUSOL" bajo la
    // misma regla), no una regla nueva por cada cupón distinto.
    .replace(/\s*\[(CUP[ÓO]N|OP\.)[^\]]*\]\s*/gi, "")
    .replace(/\s+/g, " ")
    .trim();
}

function inferCategory(desc, overrides = {}) {
  const key = descKey(desc);
  if (overrides[key]) return overrides[key];
  const d = (desc || "").toUpperCase();
  for (const [cat, keywords] of CATEGORY_RULES) {
    if (keywords.some((k) => (k instanceof RegExp ? k.test(d) : d.includes(k)))) return cat;
  }
  return "Otros";
}

const MESES_ES = { ene: "01", feb: "02", mar: "03", abr: "04", may: "05", jun: "06", jul: "07", ago: "08", sep: "09", set: "09", oct: "10", nov: "11", dic: "12" };
function fechaBbvaAIso(d, mmm, yy) {
  const mm = MESES_ES[mmm.toLowerCase().slice(0, 3)];
  if (!mm) return null;
  const yyyy = Number(yy) < 70 ? `20${yy}` : `19${yy}`;
  return `${yyyy}-${mm}-${String(d).padStart(2, "0")}`;
}

// pdfjs-dist se importa recién acá adentro, en el momento en que se
// necesita (al subir un PDF) — así el resto de la app puede seguir
// previsualizándose en entornos que no permiten esa librería de entrada
// (como el preview de artifacts de Claude). En el servidor real (Vite)
// esto funciona igual, solo que la carga queda diferida al primer uso.
let _pdfjsLibCache = null;
async function cargarPdfjs() {
  if (_pdfjsLibCache) return _pdfjsLibCache;
  // Build "legacy" de pdfjs-dist, fijado en la versión 3.11.174 (más
  // estable en Safari/iOS que las versiones recientes — ver
  // mozilla/pdf.js#20479). Esta versión predata el pasaje a ESM-only,
  // por eso el archivo es .js y no .mjs.
  const pdfjsLib = await import("pdfjs-dist/legacy/build/pdf.js");
  pdfjsLib.GlobalWorkerOptions.workerSrc = "/pdf.worker.min.mjs";
  _pdfjsLibCache = pdfjsLib;
  return pdfjsLib;
}

function conTimeout(promise, ms, mensaje) {
  return Promise.race([
    promise,
    new Promise((_, reject) => setTimeout(() => reject(new Error(mensaje)), ms)),
  ]);
}

// Reconstruye líneas de texto a partir de los items posicionados que
// devuelve pdf.js (que por sí solo no separa por renglones).
//
// OJO: no agrupamos por Y exacto (Math.round) — en algunos resúmenes de
// BBVA el importe de una fila se renderiza 1pt más arriba o abajo que el
// resto de la línea (fecha/descripción/cupón), así que dos ítems del
// MISMO renglón pueden caer en dos "buckets" redondeados distintos y
// separarse en dos líneas — perdiendo el importe (bug real, visto en un
// resumen de julio/26 donde esto rompía más de la mitad de las filas).
// Por eso agrupamos por CERCANÍA vertical (cluster con tolerancia) en
// vez de por igualdad exacta.
async function extraerLineasPdf(file) {
  const pdfjsLib = await cargarPdfjs();
  const buf = await file.arrayBuffer();
  const pdf = await conTimeout(pdfjsLib.getDocument({ data: buf }).promise, 25000, "El PDF tardó demasiado en procesarse (más de 25s). Puede ser un problema de conexión con el worker de lectura.");
  const lineas = [];
  const TOLERANCIA_Y = 2.5; // puntos
  for (let p = 1; p <= pdf.numPages; p++) {
    const page = await pdf.getPage(p);
    const content = await page.getTextContent();
    const items = content.items
      .filter((it) => (it.str || "").trim() !== "")
      .map((it) => ({ str: it.str, x: it.transform[4], y: it.transform[5] }))
      .sort((a, b) => b.y - a.y || a.x - b.x);

    const grupos = [];
    let grupoActual = [];
    let yPromedio = null;
    items.forEach((it) => {
      if (yPromedio === null || Math.abs(it.y - yPromedio) <= TOLERANCIA_Y) {
        grupoActual.push(it);
        // Promedio del grupo (no el último Y) para no "derivar" de a poco
        // hacia abajo en tablas con muchas filas seguidas.
        yPromedio = grupoActual.reduce((s, g) => s + g.y, 0) / grupoActual.length;
      } else {
        grupos.push(grupoActual);
        grupoActual = [it];
        yPromedio = it.y;
      }
    });
    if (grupoActual.length) grupos.push(grupoActual);

    grupos.forEach((grupo) => {
      const linea = grupo.sort((a, b) => a.x - b.x).map((it) => it.str).join(" ").replace(/\s+/g, " ").trim();
      if (linea) lineas.push(linea);
    });
  }
  return lineas;
}

// Parser específico del formato de resumen BBVA (Visa Signature / Mastercard Black)
function parsearResumenBBVA(lineasCrudas, nombreArchivo, overrides = {}) {
  const filas = [];
  const avisos = [];

  // pdf.js a veces pega texto de pie de página (el aviso legal fijo del
  // banco, la referencia "Sobre (...) N de M / Página X de Y", o
  // artefactos del código de barras) a la MISMA línea que un dato real
  // — pasa cuando una sección termina justo al final de una página. Eso
  // rompe los patrones que exigen que la línea TERMINE en el importe, y
  // se pierde el consumo en silencio (bug real: así se perdió un
  // consumo de Natalia de $15.841,50 en un resumen). Lo limpiamos acá,
  // antes de procesar nada.
  const lineas = lineasCrudas
    .map((l) => l
      .replace(/\s*Banco BBVA Argentina S\.A\.\s*-\s*IVA Responsable Inscripto CUIT Nro\.\s*[\d-]+\s*/gi, " ")
      .replace(/\s*Sobre\s*\(\d+\)\s*\d+\s*de\s*\d+\s*\/\s*P[áa]gina\s*\d+\s*de\s*\d+\s*/gi, " ")
      .replace(/Ë[^\sÌ]{0,20}Ì/g, " ")
      .replace(/\s+/g, " ")
      .trim())
    .filter(Boolean);

  const cuentaMatch = lineas.find((l) => /Visa Signature|Mastercard Black/i.test(l));
  const cuentaBase = cuentaMatch ? (cuentaMatch.match(/Visa Signature|Mastercard Black/i) || [])[0] : nombreArchivo;

  // El destinatario del resumen (el nombre que figura en el sobre, ej.
  // "WAJSMAN NATALIA" o "ISRAEL HERNAN") es el titular REAL de esa
  // cuenta — lo usamos para distinguir "tu" Visa Signature de la de
  // Natalia cuando ambos tienen una tarjeta con el mismo nombre
  // genérico. Sin esto, las dos quedaban mezcladas bajo "Visa
  // Signature" a secas. Buscamos solo en las líneas ANTES de "Tarjetas
  // de Crédito" (el destinatario del sobre), para no confundirlo con
  // los encabezados "Consumos Hernan Israel / Natalia Wajsman" que
  // aparecen más abajo y pueden estar presentes los dos en un mismo PDF.
  const idxTarjetas = lineas.findIndex((l) => /^Tarjetas de Cr[ée]dito/i.test(l));
  const textoDestinatario = (idxTarjetas > 0 ? lineas.slice(0, idxTarjetas) : lineas.slice(0, 15)).join(" ");
  let titularCuenta = null;
  if (/WAJSMAN\s+NATALIA/i.test(textoDestinatario)) titularCuenta = "Natalia";
  else if (/ISRAEL\s+HERNAN/i.test(textoDestinatario)) titularCuenta = "Hernán";
  const cuenta = titularCuenta ? `${cuentaBase} ${titularCuenta}` : cuentaBase;

  const textoCompleto = lineas.join("\n");

  // El patrón "clásico" (la etiqueta seguida directo de su fecha) sirve
  // para resúmenes donde cada dato va pegado a su título. En el formato
  // "Consolidado" más nuevo, en cambio, varios títulos van juntos en una
  // fila ("CIERRE ACTUAL VENCIMIENTO ACTUAL...") y sus valores en la fila
  // de abajo, en el mismo orden — ahí el patrón clásico no encuentra nada
  // porque entre la etiqueta y su fecha hay de por medio el resto de los
  // títulos. Por eso, si el patrón clásico falla, escaneamos esa fila de
  // encabezado y tomamos las fechas en el mismo orden en que aparecen las
  // columnas (cierre primero, vencimiento después).
  function buscarFechaEnEncabezado(cual) {
    const headerIdx = lineas.findIndex((l) => /CIERRE ACTUAL/i.test(l) && /VENCIMIENTO ACTUAL/i.test(l));
    if (headerIdx === -1) return null;
    const fechas = [];
    for (let i = headerIdx + 1; i < Math.min(headerIdx + 4, lineas.length) && fechas.length < 2; i++) {
      [...lineas[i].matchAll(/(\d{2})-([A-Za-zÁÉÍÓÚáéíóú]{3})-(\d{2})/g)].forEach((m) => {
        if (fechas.length < 2) fechas.push(m);
      });
    }
    const idx = cual === "cierre" ? 0 : 1;
    return fechas[idx] ? fechaBbvaAIso(fechas[idx][1], fechas[idx][2], fechas[idx][3]) : null;
  }

  const vencMatch = textoCompleto.match(/VENCIMIENTO ACTUAL\s*\n?\s*(\d{2})-([A-Za-zÁÉÍÓÚáéíóú]{3})-(\d{2})/i);
  const cierreMatch = textoCompleto.match(/CIERRE ACTUAL\s*\n?\s*(\d{2})-([A-Za-zÁÉÍÓÚáéíóú]{3})-(\d{2})/i);

  // Fecha de vencimiento — se usa para el ajuste de impuestos/tasas.
  let fechaVenc = null;
  let origenFecha = "";
  if (vencMatch) { fechaVenc = fechaBbvaAIso(vencMatch[1], vencMatch[2], vencMatch[3]); origenFecha = "vencimiento"; }
  if (!fechaVenc) { const f = buscarFechaEnEncabezado("vencimiento"); if (f) { fechaVenc = f; origenFecha = "vencimiento"; } }
  if (!fechaVenc && cierreMatch) { fechaVenc = fechaBbvaAIso(cierreMatch[1], cierreMatch[2], cierreMatch[3]); origenFecha = "cierre"; }
  if (!fechaVenc) { const f = buscarFechaEnEncabezado("cierre"); if (f) { fechaVenc = f; origenFecha = "cierre"; } }

  // Fecha de cierre — se usa para las compras en cuotas (ver más abajo).
  let fechaCierre = null;
  if (cierreMatch) fechaCierre = fechaBbvaAIso(cierreMatch[1], cierreMatch[2], cierreMatch[3]);
  if (!fechaCierre) fechaCierre = buscarFechaEnEncabezado("cierre");

  let seccionActual = null; // "Hernan Israel" | "Natalia Wajsman" | null
  let enSeccionPagos = false;
  // Suma de los consumos itemizados que sí pudimos leer línea por línea —
  // sirve para no volver a contarlos dentro del "Total a pagar" del
  // resumen (ver más abajo: ese total ya los incluye).
  let sumaConsumosPeriodo = 0;
  // Lo mismo pero abierto por persona — sirve para cotejar contra el
  // "TOTAL CONSUMOS DE ..." que trae el propio resumen y detectar de
  // entrada si el parseo se comió o duplicó alguna línea.
  const sumaPorPersona = {};
  // Suma de "Sus pagos y ajustes realizados" (pagos, notas de crédito,
  // etc. del período) — sirve para cotejar contra el SALDO ANTERIOR y
  // detectar si quedó algo sin cancelar del resumen previo.
  let sumaAjustesAnteriores = 0;
  const LINEA_MOV = /^(\d{2})-([A-Za-zÁÉÍÓÚáéíóú]{3})-(\d{2})\s+(.+?)\s+(\d{6})\s+(-?[\d.,]+)\s*$/;
  const LINEA_AJUSTE = /^(\d{2})-([A-Za-zÁÉÍÓÚáéíóú]{3})-(\d{2})\s+(.+?)\s+(-?[\d.,]+,\d{2})\s*$/;

  lineas.forEach((linea) => {
    if (/^Sus pagos y ajustes realizados/i.test(linea)) { enSeccionPagos = true; return; }
    if (enSeccionPagos) {
      // OJO: "SU PAGO EN PESOS" es el pago que ya hiciste del resumen
      // ANTERIOR (el que trae itemizados sus propios consumos + su
      // diferencia de impuestos/tasas, cargados cuando procesamos ESE
      // PDF). No la volvemos a cargar como gasto nuevo acá — si lo
      // hiciéramos, estaríamos sumando la misma plata una tercera vez.
      // No hace falta "avisar" nada: no es un error, es plata que ya
      // está contabilizada en el mes en que se gastó. Sí sumamos su
      // monto (junto con cualquier otro ajuste de esta sección, como
      // notas de crédito) para cotejar más abajo contra el SALDO
      // ANTERIOR y detectar si quedó algo pendiente de cancelar.
      if (/^Consumos\s+(Hernan Israel|Natalia Wajsman)/i.test(linea)) {
        enSeccionPagos = false;
        seccionActual = /Natalia/i.test(linea) ? "Natalia Wajsman" : "Hernan Israel";
        return;
      }
      const aj = linea.match(LINEA_AJUSTE);
      if (aj) sumaAjustesAnteriores += Number(aj[5].replace(/\./g, "").replace(",", "."));
      return;
    }
    if (/^Consumos\s+(Hernan Israel|Natalia Wajsman)/i.test(linea)) {
      seccionActual = /Natalia/i.test(linea) ? "Natalia Wajsman" : "Hernan Israel";
      return;
    }
    if (/^TOTAL CONSUMOS/i.test(linea)) { seccionActual = null; return; }
    if (!seccionActual) return;
    // Línea con un importe en moneda extranjera METIDO EN LA DESCRIPCIÓN
    // (Apple/Google/Spotify/Anthropic en USD, pero también vimos un cargo
    // en ILS de una tienda israelí) — el número al final de esas líneas
    // es el equivalente en dólares, NO en pesos, y NO hay que sumarlo
    // como si fuera un gasto en pesos. Antes solo se buscaba "USD"
    // textual; generalizado acá a cualquier código de 3 letras mayúsculas
    // seguido de un importe (bug real: un cargo en ILS de $4,47 dólares
    // se coló como si fueran $4,47 PESOS). No exigimos límite de palabra
    // antes del código — muchos códigos de transacción lo pegan directo
    // sin espacio (ej. "in1Tl7fDBUSD 20,00").
    const monedaExtranjera = linea.match(/[A-Z]{3}\s+\d+,\d{2}\b/);
    if (monedaExtranjera && !/,\d{2}\s*$/.test(linea.slice(0, monedaExtranjera.index).trim())) {
      return;
    }
    const m = linea.match(LINEA_MOV);
    if (!m) return;
    const [, dd, mmm, yy, descRaw, cupon, montoRaw] = m;
    const fechaCompra = fechaBbvaAIso(dd, mmm, yy);
    if (!fechaCompra) { avisos.push(`No pude leer la fecha en: "${linea}"`); return; }
    const monto = Number(montoRaw.replace(/\./g, "").replace(",", "."));
    if (!monto) return;
    sumaConsumosPeriodo += monto;
    const esCuota = /\s*C\.\d{2}\/\d{2}\s*$/.test(descRaw);
    // El Nro. de Cupón identifica cada consumo de forma única — lo
    // agregamos a la descripción para que quede como parte de la firma
    // de duplicados (fecha+monto+descripción+cuenta). Sin esto, dos
    // consumos reales y DISTINTOS con la misma fecha, mismo monto y
    // mismo comercio (ej. dos peajes de AUSOL el mismo día, ambos
    // $994,15) se detectaban como "el mismo movimiento" y se
    // descartaban al importar — un bug real, visto con varios AUSOL,
    // una recarga SUBE y hasta compras de entradas puntuales.
    const desc = `${descRaw.replace(/\s*C\.\d{2}\/\d{2}\s*$/, (s) => ` (cuota${s.trim().replace("C.", " ")})`).trim()} [Cupón ${cupon}]`;
    const persona = seccionActual === "Natalia Wajsman" ? "Natalia" : "Hernán";
    sumaPorPersona[persona] = (sumaPorPersona[persona] || 0) + monto;
    // Las compras en cuotas traen en la columna FECHA la fecha de la
    // compra ORIGINAL (a veces meses atrás), no la de esta cuota — si
    // usáramos esa fecha, el gasto de ESTE resumen quedaría esparcido
    // entre varios meses pasados y no se podría ver de un vistazo cuánto
    // costó la tarjeta este mes. Por eso, para cuotas, usamos la fecha
    // de cierre de ESTE resumen (cuando se cobra esta cuota puntual); la
    // fecha de compra original queda igual visible en la descripción.
    const fecha = esCuota && fechaCierre ? fechaCierre : fechaCompra;
    filas.push({
      date: fecha,
      type: "gasto",
      category: inferCategory(desc, overrides),
      amount: monto,
      desc: persona === "Natalia" ? `${desc} (Natalia)` : desc,
      account: cuenta,
      persona,
      origen: "pdf_bbva",
    });
  });

  // El "Total a pagar" del resumen YA INCLUYE los consumos itemizados que
  // acabamos de cargar arriba (sumaConsumosPeriodo) más impuestos, tasas
  // e intereses que no vienen desglosados línea por línea. Si cargáramos
  // el total completo de nuevo, estaríamos contando esa plata dos veces.
  // Por eso acá solo agregamos la DIFERENCIA (impuestos/tasas/intereses),
  // fechada el mismo día que el resto de los consumos de este resumen
  // (fecha de cierre) — así todo lo de ESTE resumen queda junto, en el
  // mismo mes, y entre los dos (consumos + este ajuste) dan el total
  // real de la tarjeta sin duplicar nada.
  // El total real del resumen lo buscamos primero como "SALDO ACTUAL" —
  // aparece siempre, en pesos, y es el balance real de la tarjeta. NO
  // usamos "LA SUMA DE $" como fuente principal: en los resúmenes
  // "Consolidado" más nuevos esa frase acompaña al débito automático del
  // PAGO MÍNIMO (una cuenta separada), NO al total de la tarjeta — usar
  // esa cifra hacía que el ajuste de impuestos/tasas diera negativo y se
  // perdiera en silencio (bug real, visto en un resumen de julio/26
  // donde "LA SUMA DE $" era el pago mínimo de $106.740 en vez del
  // saldo real de $791.105,75). "LA SUMA DE $" queda de respaldo nomás,
  // para algún resumen viejo donde no aparezca "SALDO ACTUAL".
  const saldoActualMatches = [...textoCompleto.matchAll(/SALDO ACTUAL\s+([\d.]+,\d{2})/gi)];
  const totalMatch = saldoActualMatches[0] || textoCompleto.match(/LA SUMA DE\s*\$\s*([\d.]+,\d{2})/i);
  const sobreMatch = textoCompleto.match(/Sobre\s*\((\d+)\)/i);

  // Validación 1: ¿quedó algo del resumen ANTERIOR sin cancelar? Si
  // SALDO ANTERIOR + los pagos/ajustes de este período no dan ~0, hay un
  // saldo pendiente (a favor o en contra) que no es evidente a simple
  // vista — avisamos y lo dejamos anotado en el ajuste de este resumen,
  // para que quede un registro permanente (no solo en el import).
  const saldoAnteriorMatch = textoCompleto.match(/SALDO ANTERIOR\s+(-?[\d.]+,\d{2})/i);
  let saldoPendienteAnteriorTexto = "";
  if (saldoAnteriorMatch) {
    const saldoAnterior = Number(saldoAnteriorMatch[1].replace(/\./g, "").replace(",", "."));
    const pendiente = Math.round((saldoAnterior + sumaAjustesAnteriores) * 100) / 100;
    if (Math.abs(pendiente) > 1) {
      avisos.push(`El saldo anterior de ${cuenta} no quedó en $0 después de los pagos/ajustes de este resumen — ${pendiente > 0 ? `parece haber quedado ${fmtARS(pendiente)} pendiente de pago` : `parece haber un saldo a favor de ${fmtARS(Math.abs(pendiente))}`}. Revisá el resumen a mano.`);
      saldoPendienteAnteriorTexto = `⚠ Saldo anterior ${pendiente > 0 ? "pendiente" : "a favor"}: ${fmtARS(Math.abs(pendiente))} · `;
    }
  }

  // Validación 2: cotejamos "TOTAL CONSUMOS DE ..." de cada titular
  // (dato que el propio resumen declara) contra lo que efectivamente
  // sumamos línea por línea — si no coincide, es señal de que el parseo
  // se comió o duplicó algo. Usamos tolerancia relativa (1% del total
  // declarado) en vez de un monto fijo — con resúmenes grandes, una
  // diferencia de unos pocos pesos por redondeo no es un error real.
  // "TOTAL CONSUMOS DE ..." aparece DOS VECES en el resumen (una en el
  // resumen de Pesos/Dólares de arriba, otra al final del detalle de esa
  // persona) — juntamos por persona para no avisar dos veces lo mismo.
  const totalesDeclaradosPorPersona = {};
  [...textoCompleto.matchAll(/TOTAL CONSUMOS DE (NATALIA WAJSMAN|HERNAN ISRAEL)\s+([\d.]+,\d{2})/gi)].forEach(([, quien, montoTxt]) => {
    const persona = /NATALIA/i.test(quien) ? "Natalia" : "Hernán";
    totalesDeclaradosPorPersona[persona] = Number(montoTxt.replace(/\./g, "").replace(",", "."));
  });
  Object.entries(totalesDeclaradosPorPersona).forEach(([persona, declarado]) => {
    const sumado = Math.round((sumaPorPersona[persona] || 0) * 100) / 100;
    const diferencia = Math.abs(declarado - sumado);
    const tolerancia = Math.max(1, declarado * 0.01); // 1% del total, con piso de $1 por redondeo
    if (diferencia > tolerancia) {
      avisos.push(`⚠ El total de consumos de ${persona} que declara el resumen (${fmtARS(declarado)}) no coincide con lo que sumé línea por línea (${fmtARS(sumado)}) en ${cuenta} — puede haber un error de parseo, revisá los movimientos de ${persona} a mano.`);
    }
  });

  if (totalMatch) {
    const totalAPagar = Number(totalMatch[1].replace(/\./g, "").replace(",", "."));
    if (saldoActualMatches.length > 1) {
      const valores = new Set(saldoActualMatches.map((m) => m[1]));
      if (valores.size > 1) {
        avisos.push(`Encontré más de un "SALDO ACTUAL" con valores distintos en ${cuenta} (${[...valores].join(" / ")}) — usé ${fmtARS(totalAPagar)}, revisá el resumen a mano.`);
      }
    }

    // Fecha del ajuste: preferimos el cierre de ESTE resumen (mismo mes
    // que los consumos de arriba); si no lo pudimos leer, caemos al
    // vencimiento, y como último recurso — nunca "hoy" — a la fecha del
    // consumo más reciente que sí pudimos leer en este mismo PDF.
    let fechaAjuste = fechaCierre;
    let origenAjuste = fechaCierre ? "cierre" : "";
    if (!fechaAjuste && fechaVenc) { fechaAjuste = fechaVenc; origenAjuste = origenFecha; }
    if (!fechaAjuste && filas.length > 0) {
      fechaAjuste = [...filas].sort((a, b) => b.date.localeCompare(a.date))[0].date;
      origenAjuste = "último consumo del resumen";
    }

    if (totalAPagar > 0 && fechaAjuste) {
      const diferencia = Math.round((totalAPagar - sumaConsumosPeriodo) * 100) / 100;
      const referenciaTotal = `Total del resumen: ${fmtARS(totalAPagar)}`;
      const vencDisplay = vencMatch ? `${vencMatch[1]}-${vencMatch[2]}-${vencMatch[3]}` : (fechaVenc ? fechaVenc.split("-").reverse().join("-") : null);
      const referenciaDoc = `${vencDisplay ? ` (vence ${vencDisplay})` : ""}${sobreMatch ? ` [Doc. ${sobreMatch[1]}]` : ""}`;

      if (sumaConsumosPeriodo <= 0) {
        // No pudimos leer ningún consumo itemizado de este resumen (por
        // ejemplo, un formato de PDF distinto al esperado) — en ese caso
        // no hay riesgo de duplicar nada, así que cargamos el total
        // completo como antes, marcado Pendiente para que se revise.
        filas.push({
          date: fechaAjuste,
          type: "gasto",
          category: "Tarjetas",
          amount: totalAPagar,
          desc: `${saldoPendienteAnteriorTexto}Total a pagar — resumen ${cuenta}${referenciaDoc}`,
          account: cuenta,
          origen: "pdf_bbva",
          pagado: false,
        });
        avisos.push(`No pude leer los consumos itemizados de ${cuenta} en este resumen, así que cargué el "Total a pagar" completo (${fmtARS(totalAPagar)}) sin desglosar — revisalo a mano.`);
      } else if (diferencia > 1) {
        filas.push({
          date: fechaAjuste,
          type: "gasto",
          category: "Tarjetas",
          amount: diferencia,
          desc: `${saldoPendienteAnteriorTexto}Impuestos y cargos de tarjeta — resumen ${cuenta}${referenciaDoc} · ${referenciaTotal}`,
          account: cuenta,
          origen: "pdf_bbva",
          pagado: true,
        });
      } else if (diferencia < -1) {
        avisos.push(`El "Total a pagar" de ${cuenta} (${fmtARS(totalAPagar)}) da MENOR que la suma de los consumos itemizados (${fmtARS(sumaConsumosPeriodo)}) — no cargué ningún ajuste, revisá el resumen a mano (puede haber un pago parcial o descuento que no estoy contemplando).`);
      }
      // Si la diferencia está entre -1 y 1 no hacemos nada (es solo redondeo).

      if (origenAjuste === "último consumo del resumen") {
        avisos.push(`La fecha del ajuste de "${cuenta}" quedó tomada por el último consumo del resumen (no encontré ni cierre ni vencimiento) — revisá que la fecha sea la correcta.`);
      }
    } else if (totalAPagar > 0) {
      avisos.push(`No pude leer ninguna fecha confiable para el "Total a pagar" de ${cuenta} — no lo agregué. Cargalo a mano si hace falta.`);
    } else {
      avisos.push("No pude leer el total a pagar del resumen (revisá el monto a mano).");
    }
  } else {
    avisos.push("No encontré la línea del total a pagar en este resumen — revisalo a mano.");
  }

  return { filas, avisos, resumenInfo: { cuenta, total: totalMatch ? Number(totalMatch[1].replace(/\./g, "").replace(",", ".")) : null } };
}

// Extrae el texto crudo de un PDF en orden de lectura (sin reconstruir
// columnas por posición) — sirve para el resumen de Mercado Pago, que es
// una sola columna de movimientos y no una tabla como BBVA.
async function extraerTextoPdf(file) {
  const pdfjsLib = await cargarPdfjs();
  const buf = await file.arrayBuffer();
  const pdf = await conTimeout(pdfjsLib.getDocument({ data: buf }).promise, 25000, "El PDF tardó demasiado en procesarse (más de 25s). Puede ser un problema de conexión con el worker de lectura.");
  let texto = "";
  for (let p = 1; p <= pdf.numPages; p++) {
    const page = await pdf.getPage(p);
    const content = await page.getTextContent();
    texto += content.items.map((it) => it.str).join(" ") + "\n";
  }
  return texto;
}

// Parser del resumen de Mercado Pago (movimientos en pesos).
// Ignora todo lo que aparezca después de "RESUMEN DE TENENCIAS EN DÓLARES".
function parsearResumenMercadoPago(fullTextCrudo, overrides = {}) {
  const filas = [];
  const avisos = [];

  // El PDF trae un pie de página fijo ("Fecha de generación: DD-MM-AAAA
  // ... Mercado Libre S.R.L. CUIT ... www.mercadopago.com.ar") que
  // pdf.js a veces mete pegado al principio de una página, sin salto de
  // línea real que lo separe de la transacción de al lado. El problema
  // real: "Fecha de generación: DD-MM-AAAA" tiene el MISMO formato que
  // una fecha de movimiento — el regex de abajo la toma como si fuera
  // una transacción real, y como no hay salto de línea que lo frene,
  // se traga TODO el texto del pie de página hasta la próxima
  // transacción real, perdiéndola en el camino y generando una fila
  // basura con una descripción gigante (bug real, visto con un PDF real
  // donde así se perdió una transacción y otra quedó con 200+
  // caracteres de descripción). Lo sacamos antes de parsear nada.
  const fullText = fullTextCrudo
    .replace(/Fecha de generaci[óo]n:\s*\d{2}-\d{2}-\d{4}/gi, " ")
    .replace(/Mercado Libre S\.R\.L\.[\s\S]*?mercadopago\.com\.ar/gi, " ")
    .replace(/Estas operaciones fueron realizadas por Industrial Valores[\s\S]*?no interviene\s*en la operaci[óo]n\.?/gi, " ");

  const seccionPesos = fullText.split(/RESUMEN DE TENENCIAS EN D[ÓO]LARES/i)[0];

  // Detecta de quién es la cuenta según el titular que figura en el propio
  // PDF, para que quede a su nombre sin importar quién lo suba a la app.
  // Miramos solo el encabezado (antes de la lista de movimientos) para no
  // confundirnos con una transferencia a la otra persona mencionada ahí.
  const encabezado = fullText.slice(0, 800);
  let titular = null;
  if (/NATALIA|WAJSMAN/i.test(encabezado)) titular = "Nati";
  else if (/HERNAN/i.test(encabezado)) titular = "negro";

  const lineRegex = /(\d{2}-\d{2}-\d{4})\s+(.+?)\s+(\d{6,})\s+\$\s*(-?[\d.]+,\d{2})\s+\$\s*(-?[\d.]+,\d{2})/g;
  const parseARS = (s) => Number(s.replace(/\./g, "").replace(",", "."));

  let m;
  let encontrados = 0;
  while ((m = lineRegex.exec(seccionPesos)) !== null) {
    encontrados++;
    const [, fechaStr, descRaw, numOp, valorStr] = m;
    const desc = descRaw.trim();

    // Salvaguarda: una descripción real de Mercado Pago es corta ("Pago
    // XYZ", "Transferencia enviada...", "Rendimientos"). Si nos quedó
    // clavado texto de portada/pie de página del PDF (le pasó a la
    // extracción de texto plana, que no reconstruye renglones por
    // posición como sí hace BBVA), la descripción capturada se vuelve
    // gigante — señal clara de que el parseo se enganchó mal. Mejor
    // saltear la fila con aviso que guardar basura en la base.
    if (desc.length > 150) {
      avisos.push(`Una fila quedó con una descripción demasiado larga (probable error de lectura del PDF) — la salteé: "${desc.slice(0, 80)}..."`);
      continue;
    }

    if (/Rendimientos/i.test(desc)) continue; // regla acordada: se excluyen
    if (/Hernan Pablo Israel/i.test(desc)) continue; // traspaso a uno mismo, excluido

    const [dd, mm, yyyy] = fechaStr.split("-");
    const fecha = `${yyyy}-${mm}-${dd}`;
    const valor = parseARS(valorStr);
    if (!valor) continue;

    const dLower = desc.toLowerCase();
    let type, category;
    if (/cocos capital/.test(dLower)) {
      // regla acordada: Cocos Capital = ahorro/inversión
      type = valor < 0 ? "ahorro" : "ingreso";
      category = valor < 0 ? "Otro" : "Otros ingresos";
    } else if (valor < 0) {
      type = "gasto";
      category = inferCategory(desc, overrides);
    } else {
      type = "ingreso";
      category = "Otros ingresos";
    }

    // El número de operación identifica cada movimiento de forma única
    // — lo agregamos a la descripción final (no a la que usamos arriba
    // para las exclusiones/categorización) para que quede como parte de
    // la firma de duplicados. Sin esto, dos recargas reales y DISTINTAS
    // con la misma fecha, mismo monto y mismo texto (ej. dos pagos de
    // "Pago AUSOL" o "Pago SUBE" del mismo día) se descartaban al
    // importar por parecer "el mismo movimiento".
    const descFinal = `${desc} [Op. ${numOp}]`;
    filas.push({ date: fecha, type, category, amount: Math.abs(valor), desc: descFinal, account: "Mercado Pago", who: titular || undefined, origen: "pdf_mercadopago" });
  }
  if (encontrados === 0) avisos.push("No encontré líneas con el formato esperado de Mercado Pago.");
  return { filas, avisos };
}

// Parser de facturas/resúmenes de aranceles de colegio. Por ahora solo
// reconoce el formato de ORT — cuando tengamos un ejemplo de Comunidad
// Betel se le agrega su propio bloque de detección acá.
const MESES_LARGOS = ["ENERO", "FEBRERO", "MARZO", "ABRIL", "MAYO", "JUNIO", "JULIO", "AGOSTO", "SETIEMBRE", "SEPTIEMBRE", "OCTUBRE", "NOVIEMBRE", "DICIEMBRE"];
function mesLargoAIso(nombreMes, anio) {
  let idx = MESES_LARGOS.indexOf(nombreMes.toUpperCase());
  if (idx === -1) return null;
  if (idx === 8) idx = 7; // "SEPTIEMBRE" y "SETIEMBRE" son el mismo mes (índice 7)
  else if (idx > 8) idx -= 1;
  return `${anio}-${String(idx + 1).padStart(2, "0")}-01`;
}

function parsearFacturaColegio(fullText, overrides = {}) {
  const filas = [];
  const avisos = [];
  const texto = fullText.replace(/\s+/g, " ");

  let escuela = null;
  if (/ASOCIACI[ÓO]N ORT ARGENTINA/i.test(texto)) escuela = "ORT";
  else if (/COMUNIDAD BETEL/i.test(texto)) escuela = "Comunidad Betel";

  if (!escuela) {
    avisos.push("No reconocí el colegio en este PDF (por ahora solo está soportado ORT). Pasámelo a mí en el chat para agregarle el formato.");
    return { filas, avisos };
  }

  const alumnoMatch = texto.match(/Alumno:\s*([A-ZÁÉÍÓÚÑ ]+?)(?:\s+ARANCEL|\s{2,}|$)/i);
  const alumno = alumnoMatch ? alumnoMatch[1].trim() : "";

  const periodoMatch = texto.match(new RegExp(`\\b(${MESES_LARGOS.join("|")})\\s+(\\d{4})\\b`, "i"));
  const cuotaMatch = texto.match(/Cuota\s+(\d+)\s+de\s+(\d+)/i);
  const importeMatch = texto.match(/ARANCEL.*?([\d.]+,\d{2})/i);

  if (!periodoMatch || !importeMatch) {
    avisos.push(`Reconocí que es ${escuela}, pero no pude leer el período o el importe del arancel. Revisá el PDF o pasámelo directamente en el chat.`);
    return { filas, avisos };
  }

  const fecha = mesLargoAIso(periodoMatch[1], periodoMatch[2]);
  const monto = Number(importeMatch[1].replace(/\./g, "").replace(",", "."));
  const cuotaTxt = cuotaMatch ? ` (cuota ${cuotaMatch[1]}/${cuotaMatch[2]})` : "";
  const desc = `${escuela} - arancel ${periodoMatch[1]} ${periodoMatch[2]}${alumno ? ` - ${alumno}` : ""}${cuotaTxt}`;

  filas.push({
    date: fecha,
    type: "gasto",
    category: overrides[descKey(desc)] || "Escuelas",
    amount: monto,
    desc,
    account: "Colegio",
    origen: "pdf_colegio",
    pagado: false, // los aranceles se suben antes de pagarlos, normalmente
  });

  return { filas, avisos };
}

// Convierte un Blob de audio a base64 (sin el prefijo "data:...;base64,")
// para poder mandarlo en un body JSON a la Edge Function de transcripción.
function blobToBase64(blob) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onloadend = () => resolve(String(reader.result).split(",")[1] || "");
    reader.onerror = reject;
    reader.readAsDataURL(blob);
  });
}

// Extrae monto / tipo / categoría / fecha de un texto dictado en español,
// para precargar el formulario de carga manual. Es heurístico a propósito:
// el usuario siempre revisa y confirma antes de guardar.
function extraerDatosDeTexto(texto, categories = []) {
  const t = (texto || "").toLowerCase();

  let type = "gasto";
  if (/cobr[ée]|me (dieron|pagaron)|ingres[eé]|deposit[eéoó]|recib[ií]/.test(t)) type = "ingreso";
  if (/cambi[ée].*d[oó]lar|cambio de d[oó]lares|d[oó]lares?\s+a\s+pesos/.test(t)) type = "cambio";

  let amount = null;
  const milMatch = t.match(/(\d+(?:[.,]\d+)?)\s*mil\b/);
  if (milMatch) {
    amount = Math.round(parseFloat(milMatch[1].replace(",", ".")) * 1000);
  } else {
    const numMatch = t.match(/\$?\s?(\d{1,3}(?:\.\d{3})+(?:,\d+)?|\d+(?:,\d+)?)/);
    if (numMatch) {
      const raw = numMatch[1];
      amount = /\.\d{3}/.test(raw)
        ? Math.round(parseFloat(raw.replace(/\./g, "").replace(",", ".")))
        : Math.round(parseFloat(raw.replace(",", ".")));
    }
  }

  // Primero probamos si el nombre de alguna categoría del hogar aparece
  // tal cual dicho en el audio; si no, caemos en las mismas reglas por
  // palabra clave que se usan para categorizar resúmenes importados.
  let category = null;
  for (const c of categories) {
    if (c && t.includes(c.toLowerCase())) { category = c; break; }
  }
  if (!category) category = inferCategory(texto, {}, categories.length ? categories : null);

  let date = todayISO();
  if (/\bayer\b/.test(t)) {
    const d = new Date();
    d.setDate(d.getDate() - 1);
    date = d.toISOString().slice(0, 10);
  }

  // Si describe algo ya hecho ("gasté", "pagué") lo marcamos Pagado. Si
  // describe algo pendiente ("tengo que pagar", "vence", "factura de...
  // que hay que pagar"), lo marcamos como Pendiente de pago.
  let pagado = true;
  if (/tengo que pagar|hay que pagar|debo pagar|falta pagar|todav[ií]a no (lo )?pagu[eé]|vence|vencimiento|a pagar\b/.test(t)) {
    pagado = false;
  }

  return { type, amount, category, date, desc: (texto || "").trim(), pagado, origen: "voz" };
}

function monthKey(dateStr) {
  return dateStr ? dateStr.slice(0, 7) : "";
}
function todayISO() {
  return new Date().toISOString().slice(0, 10);
}

const MESES_NOMBRE = ["Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"];

const ORIGEN_LABELS = {
  manual: "Carga manual",
  voz: "Por voz",
  foto: "Foto de recibo",
  pdf_bbva: "Importado (PDF BBVA)",
  pdf_mercadopago: "Importado (PDF Mercado Pago)",
  pdf_colegio: "Importado (PDF Colegio)",
  csv: "Importado (CSV/texto)",
  recurrente: "Gasto recurrente automático",
  acceso_rapido: "Acceso rápido",
  whatsapp: "Por WhatsApp",
};

// Selector de fecha propio (Día / Mes con nombre / Año) — el input nativo
// type="date" muestra el formato según el idioma del sistema operativo del
// navegador (por eso a veces se ve mm/dd/yyyy sin que podamos forzarlo por
// CSS); con selects explícitos el orden día-mes-año queda fijo siempre, y
// al escribir el mes por nombre no hay ninguna ambigüedad posible.
function FechaInput({ value, onChange }) {
  const partes = (value || todayISO()).split("-").map(Number);
  const y = partes[0] || new Date().getFullYear();
  const m = partes[1] || 1;
  const d = partes[2] || 1;
  const diasEnMes = new Date(y, m, 0).getDate();

  function actualizar(nuevoD, nuevoM, nuevoY) {
    const maxDia = new Date(nuevoY, nuevoM, 0).getDate();
    const diaFinal = Math.min(nuevoD, maxDia);
    onChange(`${nuevoY}-${String(nuevoM).padStart(2, "0")}-${String(diaFinal).padStart(2, "0")}`);
  }

  const anioActual = new Date().getFullYear();
  const anios = Array.from({ length: 8 }, (_, i) => anioActual - 6 + i);

  return (
    <div style={{ display: "flex", gap: 6 }}>
      <select value={d} onChange={(e) => actualizar(Number(e.target.value), m, y)} style={{ ...inputStyle, flex: 1 }}>
        {Array.from({ length: diasEnMes }, (_, i) => i + 1).map((dd) => <option key={dd} value={dd}>{dd}</option>)}
      </select>
      <select value={m} onChange={(e) => actualizar(d, Number(e.target.value), y)} style={{ ...inputStyle, flex: 2 }}>
        {MESES_NOMBRE.map((nombre, i) => <option key={i} value={i + 1}>{nombre}</option>)}
      </select>
      <select value={y} onChange={(e) => actualizar(d, m, Number(e.target.value))} style={{ ...inputStyle, flex: 1 }}>
        {anios.map((yy) => <option key={yy} value={yy}>{yy}</option>)}
      </select>
    </div>
  );
}
function uid() {
  return Date.now().toString(36) + Math.random().toString(36).slice(2, 7);
}

// uid() no genera UUIDs válidos (son para columnas "text" como entries.id).
// Para tablas cuyo id es "uuid" (recurring_entries, quick_entries) hace
// falta esto, o Postgres rechaza el insert.
function uidUuid() {
  if (typeof crypto !== "undefined" && crypto.randomUUID) return crypto.randomUUID();
  return "10000000-1000-4000-8000-100000000000".replace(/[018]/g, (c) =>
    (c ^ (crypto.getRandomValues(new Uint8Array(1))[0] & (15 >> (c / 4)))).toString(16)
  );
}

function safeEnv() {
  // Leemos de window en vez de import.meta.env: funciona igual en el build
  // real de Vite (ver index.html) y no rompe nada acá en el preview de
  // Claude, donde estas variables directamente no van a existir.
  try {
    return {
      url: typeof window !== "undefined" ? window.__SUPABASE_URL__ : undefined,
      key: typeof window !== "undefined" ? window.__SUPABASE_ANON_KEY__ : undefined,
    };
  } catch {
    return { url: undefined, key: undefined };
  }
}
const { url: SUPABASE_URL, key: SUPABASE_ANON_KEY } = safeEnv();
const HAS_SUPABASE = Boolean(SUPABASE_URL && SUPABASE_ANON_KEY && !SUPABASE_URL.includes("%VITE_"));

// El token de sesión del usuario logueado se guarda acá (fuera de React)
// para que sb() lo use en cada request — necesario para que las políticas
// de RLS por hogar (auth.uid()) funcionen. El refresh_token viaja con él,
// para poder renovar la sesión sola cuando el access_token vence (dura
// ~1h) SIN interrumpir a la persona con un login — antes esto solo se
// intentaba una vez al abrir la app; si la sesión seguía abierta más de
// una hora, todo empezaba a fallar hasta recargar la página.
let _accessToken = null;
let _refreshToken = null;
let _onSessionRefreshed = null; // callback que registra el componente App, para guardar en localStorage + actualizar el estado de React cuando sb() renueva la sesión sola
function setAccessToken(token) { _accessToken = token; }
function setRefreshToken(token) { _refreshToken = token; }
function setOnSessionRefreshed(fn) { _onSessionRefreshed = fn; }

let _refreshEnCurso = null; // evita disparar varios refresh en paralelo si llegan varios 401 juntos
async function refrescarTokenSupabase() {
  if (!_refreshToken) throw new Error("Sin refresh_token disponible");
  if (_refreshEnCurso) return _refreshEnCurso;
  _refreshEnCurso = (async () => {
    const renovada = await sbAuth("token?grant_type=refresh_token", {
      method: "POST",
      body: JSON.stringify({ refresh_token: _refreshToken }),
    });
    if (!renovada?.access_token) throw new Error("No se pudo renovar la sesión");
    _accessToken = renovada.access_token;
    _refreshToken = renovada.refresh_token;
    if (_onSessionRefreshed) _onSessionRefreshed(renovada);
    return renovada;
  })();
  try {
    return await _refreshEnCurso;
  } finally {
    _refreshEnCurso = null;
  }
}

async function sb(path, options = {}, _reintentado = false) {
  const res = await fetch(`${SUPABASE_URL}/rest/v1/${path}`, {
    ...options,
    headers: {
      apikey: SUPABASE_ANON_KEY,
      Authorization: `Bearer ${_accessToken || SUPABASE_ANON_KEY}`,
      "Content-Type": "application/json",
      ...(options.headers || {}),
    },
  });
  if (res.status === 401 && !_reintentado && _refreshToken) {
    // Access token vencido en medio de un uso largo — lo renovamos solos
    // y reintentamos una vez, sin que la persona vea ningún error.
    try {
      await refrescarTokenSupabase();
      return sb(path, options, true);
    } catch {
      // El refresh_token también está vencido/inválido — no hay nada
      // más para hacer del lado del cliente; sigue abajo y falla con el
      // 401 original, que es lo que fuerza el login más arriba.
    }
  }
  if (!res.ok) {
    const body = await res.text().catch(() => "");
    throw new Error(`Supabase ${path} → ${res.status}: ${body}`);
  }
  if (res.status === 204) return null;
  return res.json().catch(() => null);
}

// Pide TODAS las filas de una consulta, sin importar cuántas haya —
// PostgREST (el motor de la API de Supabase) trunca en silencio
// cualquier consulta sin límite explícito a un tope por defecto (suele
// ser 1000 filas). Con "select=*&order=date.desc" sin paginar, una vez
// que el hogar supera ese tope, las filas más VIEJAS quedaban afuera de
// la respuesta sin ningún error — bug real: así "desaparecían" enero y
// febrero de la vista después de varias importaciones, aunque los datos
// seguían enteros en la base. Acá pedimos de a páginas hasta que una
// venga incompleta (esa es la señal de que no queda nada más).
async function sbAllPages(path) {
  const TAMANO_PAGINA = 1000;
  let offset = 0;
  let todas = [];
  while (true) {
    const sep = path.includes("?") ? "&" : "?";
    const pagina = await sb(`${path}${sep}limit=${TAMANO_PAGINA}&offset=${offset}`);
    const filas = pagina || [];
    todas = todas.concat(filas);
    if (filas.length < TAMANO_PAGINA) break;
    offset += TAMANO_PAGINA;
  }
  return todas;
}

async function sbAuth(path, options = {}) {
  const res = await fetch(`${SUPABASE_URL}/auth/v1/${path}`, {
    ...options,
    headers: {
      apikey: SUPABASE_ANON_KEY,
      "Content-Type": "application/json",
      ...(options.headers || {}),
    },
  });
  const data = await res.json().catch(() => null);
  if (!res.ok) {
    throw new Error(data?.msg || data?.error_description || data?.error || `Auth ${path} → ${res.status}`);
  }
  return data;
}

// Llama a una Edge Function de Supabase (ej. la que transcribe audio).
async function sbFunction(path, body) {
  const res = await fetch(`${SUPABASE_URL}/functions/v1/${path}`, {
    method: "POST",
    headers: {
      apikey: SUPABASE_ANON_KEY,
      Authorization: `Bearer ${_accessToken || SUPABASE_ANON_KEY}`,
      "Content-Type": "application/json",
    },
    body: JSON.stringify(body),
  });
  const data = await res.json().catch(() => null);
  if (!res.ok) {
    throw new Error(data?.error || `Función ${path} → ${res.status}`);
  }
  return data;
}

// --- Desbloqueo biométrico del dispositivo (Face ID / Touch ID / huella) ---
// Usa WebAuthn con un autenticador de plataforma. Es una verificación LOCAL
// del dispositivo (no viaja a ningún servidor ni reemplaza el login real):
// sirve para no tener que mostrar los datos apenas se abre la app en el
// celu, ya con la sesión de Supabase restaurada, sin re-escribir usuario y
// contraseña cada vez.
function generarDesafioWebAuthn() {
  const arr = new Uint8Array(32);
  crypto.getRandomValues(arr);
  return arr;
}
function bufferABase64(buf) {
  return btoa(String.fromCharCode(...new Uint8Array(buf)));
}
function base64ABuffer(b64) {
  const bin = atob(b64);
  const bytes = new Uint8Array(bin.length);
  for (let i = 0; i < bin.length; i++) bytes[i] = bin.charCodeAt(i);
  return bytes.buffer;
}
async function soportaBiometria() {
  if (typeof window === "undefined" || !window.PublicKeyCredential || !PublicKeyCredential.isUserVerifyingPlatformAuthenticatorAvailable) return false;
  try {
    return await PublicKeyCredential.isUserVerifyingPlatformAuthenticatorAvailable();
  } catch {
    return false;
  }
}
async function registrarBiometria(nombre) {
  const cred = await navigator.credentials.create({
    publicKey: {
      challenge: generarDesafioWebAuthn(),
      rp: { name: "Finanzas del hogar" },
      user: { id: generarDesafioWebAuthn(), name: nombre || "usuario", displayName: nombre || "Usuario" },
      pubKeyCredParams: [{ type: "public-key", alg: -7 }, { type: "public-key", alg: -257 }],
      authenticatorSelection: { authenticatorAttachment: "platform", userVerification: "required" },
      timeout: 60000,
      attestation: "none",
    },
  });
  return bufferABase64(cred.rawId);
}
async function verificarBiometria(credentialIdB64) {
  await navigator.credentials.get({
    publicKey: {
      challenge: generarDesafioWebAuthn(),
      allowCredentials: [{ id: base64ABuffer(credentialIdB64), type: "public-key", transports: ["internal"] }],
      userVerification: "required",
      timeout: 60000,
    },
  });
}

// Mapea una fila de la tabla `entries` (Supabase) al shape que usa la app
function entryFromDb(row) {
  return {
    id: row.id,
    date: row.date,
    type: row.type,
    category: row.category,
    amount: Number(row.amount),
    desc: row.descripcion || "",
    account: row.account || "",
    who: row.who || "",
    persona: row.persona || "",
    origen: row.origen || "manual",
    createdAt: row.created_at || null,
    usdAmount: row.usd_amount != null ? Number(row.usd_amount) : undefined,
    rate: row.rate != null ? Number(row.rate) : undefined,
    moneda: row.moneda || "ARS",
    recurringId: row.recurring_id || undefined,
    generatedMonth: row.generated_month || undefined,
    pagado: row.pagado !== false, // default true si no viene (compatibilidad con filas viejas)
  };
}
// Mapea un entry de la app al shape de la tabla `entries` (Supabase)
function entryToDb(e) {
  return {
    id: e.id,
    date: e.date,
    type: e.type,
    category: e.category || "",
    amount: Number(e.amount),
    descripcion: e.desc || "",
    account: e.account || "",
    who: e.who || "",
    persona: e.persona || e.who || "",
    origen: e.origen || "manual",
    usd_amount: e.usdAmount != null ? Number(e.usdAmount) : null,
    rate: e.rate != null ? Number(e.rate) : null,
    moneda: e.moneda || "ARS",
    recurring_id: e.recurringId || null,
    generated_month: e.generatedMonth || null,
    pagado: e.pagado !== false,
  };
}

function safeGet(key) {
  try {
    const v = localStorage.getItem(key);
    return v ? JSON.parse(v) : null;
  } catch {
    return null;
  }
}
function safeSet(key, value) {
  try {
    localStorage.setItem(key, JSON.stringify(value));
  } catch (e) {
    console.error("storage error", e);
  }
}

// --- Modo local (preview sin Supabase) ---
// Se usa automáticamente cuando no hay variables de entorno de Supabase
// (por ejemplo, acá en el preview de Claude). Los datos quedan solo en
// este navegador, no se pierden los reales de producción.
function mockLoad() {
  return {
    entries: safeGet("mock_entries") || [],
    budgets: safeGet("mock_budgets") || {},
    names: safeGet("mock_names") || [],
    overrides: safeGet("mock_overrides") || {},
    categories: safeGet("mock_categories") || DEFAULT_GASTO_CATS,
    recurrentes: safeGet("mock_recurrentes") || [],
    accesosRapidos: safeGet("mock_accesos_rapidos") || [],
  };
}
function mockSaveEntries(entries) { safeSet("mock_entries", entries); }
function mockSaveCategories(categories) { safeSet("mock_categories", categories); }
function mockSaveBudgets(budgets) { safeSet("mock_budgets", budgets); }
function mockSaveNames(names) { safeSet("mock_names", names); }
function mockSaveOverrides(overrides) { safeSet("mock_overrides", overrides); }
function mockSaveRecurrentes(recurrentes) { safeSet("mock_recurrentes", recurrentes); }
function mockSaveAccesosRapidos(accesos) { safeSet("mock_accesos_rapidos", accesos); }

// Genera y descarga un Excel a partir de una lista de movimientos —
// reutilizable tanto para "exportar todo" (header) como para "exportar
// justo lo que está filtrado" (dentro de Movimientos).
function exportarEntriesAExcel(lista, sufijoNombre = "movimientos") {
  const filas = [...lista]
    .sort((a, b) => (a.date || "").localeCompare(b.date || ""))
    .map((e) => ({
      Fecha: e.date || "",
      Tipo: e.type || "",
      Categoría: e.category || "",
      Monto: Number(e.amount) || 0,
      Descripción: e.desc || "",
      Cuenta: e.account || "",
      Quién: e.who || "",
      "USD (si es cambio)": e.usdAmount || "",
      "Tipo de cambio (si es cambio)": e.rate || "",
    }));
  const ws = XLSX.utils.json_to_sheet(filas);
  ws["!cols"] = [
    { wch: 11 }, { wch: 10 }, { wch: 16 }, { wch: 13 }, { wch: 32 }, { wch: 20 }, { wch: 10 }, { wch: 14 }, { wch: 16 },
  ];
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "Movimientos");
  XLSX.writeFile(wb, `finanzas_del_hogar_${sufijoNombre}_${todayISO()}.xlsx`);
}

export default function FinanzasApp() {
  const [loading, setLoading] = useState(true);
  const [session, setSession] = useState(null); // { access_token, refresh_token, user }
  const [householdId, setHouseholdId] = useState(null);

  // Detecta sola cuando se publicó una versión nueva (después de
  // /publicar), sin que haga falta cerrar y volver a abrir la app. Compara
  // el HTML que Vercel está sirviendo AHORA contra el que quedó cargado
  // cuando arrancó esta sesión — si cambió, es porque hubo un deploy
  // nuevo (los nombres de archivo de Vite cambian con cada build). Se
  // revisa cada 3 minutos y también al volver a la pestaña/app.
  const [hayVersionNueva, setHayVersionNueva] = useState(false);
  useEffect(() => {
    if (typeof window === "undefined") return;
    let htmlBase = null;
    let cancelado = false;
    async function chequearVersionNueva() {
      try {
        const res = await fetch(`/?_v=${Date.now()}`, { cache: "no-store" });
        const html = await res.text();
        if (cancelado) return;
        if (htmlBase === null) { htmlBase = html; return; } // primera lectura: solo fija la base de comparación
        if (html !== htmlBase) setHayVersionNueva(true);
      } catch {
        // sin conexión o falla de red puntual — no pasa nada, se reintenta solo
      }
    }
    chequearVersionNueva();
    const intervalo = setInterval(chequearVersionNueva, 3 * 60 * 1000);
    const alVolver = () => { if (document.visibilityState === "visible") chequearVersionNueva(); };
    document.addEventListener("visibilitychange", alVolver);
    window.addEventListener("focus", chequearVersionNueva);
    return () => {
      cancelado = true;
      clearInterval(intervalo);
      document.removeEventListener("visibilitychange", alVolver);
      window.removeEventListener("focus", chequearVersionNueva);
    };
  }, []);

  // Cada vez que sb() renueva el access_token sola (a mitad de una
  // sesión larga), avisa acá — guardamos la sesión nueva en localStorage
  // y actualizamos el estado de React, para que quede todo sincronizado
  // sin que la persona note nada.
  setOnSessionRefreshed((s) => { safeSet("auth_session", s); setSession(s); });

  // Refresca solo los movimientos (lo que más cambia "por afuera" de esta
  // sesión, ej. algo cargado por WhatsApp mientras la app estaba abierta) —
  // liviano, sin volver a traer categorías/presupuestos/etc.
  async function refrescarEntries() {
    if (!HAS_SUPABASE || !householdId) return;
    try {
      const entRows = await sbAllPages(`entries?household_id=eq.${householdId}&select=*&order=date.desc`);
      setEntries((entRows || []).map(entryFromDb));
    } catch (e) {
      console.error("No se pudo refrescar movimientos", e);
    }
  }

  useEffect(() => {
    if (!HAS_SUPABASE || !householdId) return;
    function alVolver() {
      if (document.visibilityState === "visible") refrescarEntries();
    }
    document.addEventListener("visibilitychange", alVolver);
    window.addEventListener("focus", refrescarEntries);
    const intervalo = setInterval(refrescarEntries, 45000); // respaldo cada 45s
    return () => {
      document.removeEventListener("visibilitychange", alVolver);
      window.removeEventListener("focus", refrescarEntries);
      clearInterval(intervalo);
    };
  }, [householdId]);
  const [profileName, setProfileName] = useState(null); // display_name dentro del hogar
  const [entries, setEntries] = useState([]);
  const [budgets, setBudgets] = useState({});
  const [categoryOverrides, setCategoryOverrides] = useState({});
  const [categories, setCategories] = useState(DEFAULT_GASTO_CATS);
  const [recurrentes, setRecurrentes] = useState([]);
  const [esAdmin, setEsAdmin] = useState(false);
  const [whatsappLinks, setWhatsappLinks] = useState([]);
  const [accesosRapidos, setAccesosRapidos] = useState([]);
  const [miembrosHogar, setMiembrosHogar] = useState([]);
  const [ultimoAccesoRegistrado, setUltimoAccesoRegistrado] = useState(null);
  const [tab, setTab] = useState("resumen");
  const [menuOpen, setMenuOpen] = useState(false);
  const [showForm, setShowForm] = useState(false);
  const [formType, setFormType] = useState("gasto");
  const [showVoice, setShowVoice] = useState(false);
  const [voicePrefill, setVoicePrefill] = useState(null);
  const [showFoto, setShowFoto] = useState(false);
  const [showAddMenu, setShowAddMenu] = useState(false);
  const [saving, setSaving] = useState(false);
  const [saveError, setSaveError] = useState(null);

  // --- Layout responsive: por ahora un único breakpoint (≥900px = desktop) ---
  // que ensancha el contenido y acomoda algunas secciones en grilla. El resto
  // de la app sigue igual que en el celu.
  const [isDesktop, setIsDesktop] = useState(() => (typeof window !== "undefined" ? window.innerWidth >= 900 : false));
  useEffect(() => {
    function onResize() { setIsDesktop(window.innerWidth >= 900); }
    window.addEventListener("resize", onResize);
    return () => window.removeEventListener("resize", onResize);
  }, []);

  const [loadError, setLoadError] = useState(null);

  // --- Desbloqueo biométrico (Face ID / Touch ID / huella) del dispositivo ---
  const [biometriaSoportada, setBiometriaSoportada] = useState(false);
  const [biometriaActiva, setBiometriaActiva] = useState(() => typeof window !== "undefined" && !!safeGet("biometric_credential_id"));
  const [desbloqueado, setDesbloqueado] = useState(false);
  const [bioBusy, setBioBusy] = useState(false);
  const [bioError, setBioError] = useState(null);
  // Banner para ofrecer activar la biometría — se muestra una sola vez
  // por dispositivo si no la activaste, y desaparece para siempre en
  // cuanto la activás o la cerrás (no hace falta ir a buscarla a "Mi
  // hogar" si no sabés que existe).
  const [bannerBiometriaCerrado, setBannerBiometriaCerrado] = useState(() => typeof window !== "undefined" && !!safeGet("banner_biometria_cerrado"));

  useEffect(() => { soportaBiometria().then(setBiometriaSoportada); }, []);

  async function activarBiometria() {
    setBioError(null);
    setBioBusy(true);
    try {
      const credId = await registrarBiometria(profileName);
      safeSet("biometric_credential_id", credId);
      setBiometriaActiva(true);
    } catch (e) {
      setBioError("No pudimos activar el desbloqueo biométrico: " + e.message);
    }
    setBioBusy(false);
  }

  function desactivarBiometria() {
    safeSet("biometric_credential_id", null);
    setBiometriaActiva(false);
    setDesbloqueado(true);
  }

  async function intentarDesbloquear() {
    setBioError(null);
    setBioBusy(true);
    try {
      const credId = safeGet("biometric_credential_id");
      await verificarBiometria(credId);
      setDesbloqueado(true);
    } catch (e) {
      setBioError("No se pudo verificar. Probá de nuevo.");
    }
    setBioBusy(false);
  }

  // --- pantalla de login/registro ---
  const [authMode, setAuthMode] = useState("login"); // "login" | "signup" | "onboarding"
  const [authEmail, setAuthEmail] = useState("");
  const [authPassword, setAuthPassword] = useState("");
  const [authDisplayName, setAuthDisplayName] = useState("");
  const [authInviteCode, setAuthInviteCode] = useState(() => {
    try {
      return new URLSearchParams(window.location.search).get("invite") || "";
    } catch {
      return "";
    }
  });
  const [authHouseholdName, setAuthHouseholdName] = useState("");
  const [authError, setAuthError] = useState(null);
  const [authBusy, setAuthBusy] = useState(false);

  async function cargarHogarYDatos(sesion) {
    setAccessToken(sesion.access_token);
    setRefreshToken(sesion.refresh_token);
    const miembro = await sb("household_members?select=household_id,display_name");
    if (!miembro || miembro.length === 0) {
      setAuthMode("onboarding");
      setLoading(false);
      return;
    }
    const hh = miembro[0];
    setHouseholdId(hh.household_id);
    setProfileName(hh.display_name);
    const [entRows, budRows, ovrRows, catRows, recRows, rapRows, waRows] = await Promise.all([
      sbAllPages(`entries?household_id=eq.${hh.household_id}&select=*&order=date.desc`),
      sb(`budgets?household_id=eq.${hh.household_id}&select=*`),
      sb(`category_overrides?household_id=eq.${hh.household_id}&select=*`),
      sb(`categories?household_id=eq.${hh.household_id}&select=name&order=name.asc`),
      sb(`recurring_entries?household_id=eq.${hh.household_id}&select=*`),
      sb(`quick_entries?household_id=eq.${hh.household_id}&select=*&order=sort_order.asc`),
      sb(`whatsapp_links?household_id=eq.${hh.household_id}&select=*`),
    ]);
    const entriesCargadas = (entRows || []).map(entryFromDb);
    setEntries(entriesCargadas);
    const budObj = {};
    (budRows || []).forEach((b) => { budObj[b.category] = Number(b.limit_amount); });
    setBudgets(budObj);
    const ovrObj = {};
    (ovrRows || []).forEach((o) => { ovrObj[o.desc_key] = o.category; });
    setCategoryOverrides(ovrObj);
    setCategories((catRows || []).map((c) => c.name).length > 0 ? catRows.map((c) => c.name) : DEFAULT_GASTO_CATS);
    const recurrentesCargadas = recRows || [];
    setRecurrentes(recurrentesCargadas);
    setAccesosRapidos(rapRows || []);
    setWhatsappLinks(waRows || []);
    setTab("resumen");
    setLoading(false);
    generarRecurrentesDelMes(recurrentesCargadas, entriesCargadas, hh.household_id, hh.display_name)
      .catch((e) => console.error("No se pudieron generar los recurrentes del mes", e));
    sb("rpc/soy_admin", { method: "POST", body: "{}" })
      .then((v) => setEsAdmin(Boolean(v)))
      .catch(() => setEsAdmin(false)); // función vieja/inexistente en Supabase → no es admin, sin romper nada
    sb("login_events", { method: "POST", body: JSON.stringify([{ household_id: hh.household_id, display_name: hh.display_name }]) })
      .catch((e) => console.error("No se pudo registrar el ingreso a la app", e));
    sb("rpc/get_my_household_members", { method: "POST", body: "{}" })
      .then((rows) => setMiembrosHogar((rows || []).map((r) => r.display_name).filter(Boolean)))
      .catch((e) => console.error("No se pudo traer la lista de miembros", e));
  }

  // Si algún gasto/ingreso recurrente activo ya "llegó" (hoy >= día
  // configurado) y todavía no se generó su movimiento para este mes, lo
  // crea. Se corre cada vez que se abre la app — no hay un cron en el
  // servidor, así que el movimiento aparece la primera vez que alguien
  // entra a la app en el día correspondiente (o después).
  async function generarRecurrentesDelMes(recurrentesList, entriesList, hhId, nombre) {
    const hoy = new Date();
    const mesActual = hoy.toISOString().slice(0, 7);
    const diaHoy = hoy.getDate();
    const nuevas = [];
    for (const r of recurrentesList) {
      if (!r.activo) continue;
      if (diaHoy < r.dia_mes) continue;
      const yaExiste = entriesList.some((e) => e.recurringId === r.id && e.generatedMonth === mesActual);
      if (yaExiste) continue;
      const fecha = `${mesActual}-${String(r.dia_mes).padStart(2, "0")}`;
      nuevas.push({
        id: uid(), type: r.type, category: r.category, amount: Number(r.amount),
        desc: r.descripcion || "", date: fecha, account: r.account || "", moneda: r.moneda || "ARS",
        who: nombre, recurringId: r.id, generatedMonth: mesActual, origen: "recurrente",
      });
    }
    if (nuevas.length === 0) return;
    setEntries((prev) => [...nuevas, ...prev]);
    if (!HAS_SUPABASE) {
      mockSaveEntries([...nuevas, ...entriesList]);
      return;
    }
    await sb("entries", { method: "POST", body: JSON.stringify(nuevas.map((n) => ({ ...entryToDb(n), household_id: hhId }))) });
  }

  async function intentarRestaurarSesion() {
    const stored = safeGet("auth_session");
    if (!stored?.access_token) {
      setLoading(false);
      return;
    }
    setSession(stored);
    setLoadError(null);
    try {
      // sb() ya intenta renovar el access_token sola si vino vencido
      // (ver más arriba) — no hace falta duplicar esa lógica acá.
      await cargarHogarYDatos(stored);
    } catch (e) {
      console.error(e);
      // Solo borramos la sesión guardada si el problema es realmente de
      // autenticación (401, o el refresh_token vencido/inválido). Un
      // error de red/conexión transitorio NO debería obligar a
      // loguearse de nuevo — antes sí pasaba: cualquier hiccup al abrir
      // la app tiraba la sesión entera aunque el token siguiera siendo
      // válido, y esa era la causa real de tener que loguearse seguido.
      const esProblemaDeSesion = /→ 401\b/.test(e?.message || "") || /invalid_grant|refresh.?token/i.test(e?.message || "");
      if (esProblemaDeSesion) {
        safeSet("auth_session", null);
        setSession(null);
        setLoading(false);
      } else {
        setLoadError("No pudimos conectar para cargar tus datos. Revisá tu conexión a internet y probá de nuevo.");
        setLoading(false);
      }
    }
  }

  useEffect(() => {
    (async () => {
      if (!HAS_SUPABASE) {
        const prof = safeGet("profile");
        if (prof?.name) setProfileName(prof.name);
        const mock = mockLoad();
        setEntries(mock.entries);
        setBudgets(mock.budgets);
        setCategoryOverrides(mock.overrides);
        setCategories(mock.categories);
        setRecurrentes(mock.recurrentes);
        setAccesosRapidos(mock.accesosRapidos);
        setLoading(false);
        generarRecurrentesDelMes(mock.recurrentes, mock.entries, null, prof?.name)
          .catch((e) => console.error("No se pudieron generar los recurrentes del mes", e));
        return;
      }
      await intentarRestaurarSesion();
    })();
  }, []);

  async function handleSignup() {
    setAuthError(null);
    setAuthBusy(true);
    try {
      const data = await sbAuth("signup", {
        method: "POST",
        body: JSON.stringify({ email: authEmail.trim(), password: authPassword }),
      });
      if (!data.access_token) {
        setAuthError("Cuenta creada. Revisá tu email para confirmarla y después iniciá sesión.");
        setAuthBusy(false);
        return;
      }
      safeSet("auth_session", data);
      setSession(data);
      setAccessToken(data.access_token);
      if (authInviteCode.trim()) {
        await sb("rpc/join_household_by_code", {
          method: "POST",
          body: JSON.stringify({ p_code: authInviteCode.trim(), p_display_name: authDisplayName.trim() || "Yo" }),
        });
      } else {
        await sb("rpc/create_household", {
          method: "POST",
          body: JSON.stringify({ p_name: authHouseholdName.trim() || "Mi hogar", p_display_name: authDisplayName.trim() || "Yo" }),
        });
      }
      await cargarHogarYDatos(data);
      setDesbloqueado(true);
    } catch (e) {
      setAuthError(e.message);
    }
    setAuthBusy(false);
  }

  async function handleLogin() {
    setAuthError(null);
    setAuthBusy(true);
    try {
      const data = await sbAuth("token?grant_type=password", {
        method: "POST",
        body: JSON.stringify({ email: authEmail.trim(), password: authPassword }),
      });
      safeSet("auth_session", data);
      setSession(data);
      await cargarHogarYDatos(data);
      setDesbloqueado(true);
    } catch (e) {
      setAuthError(e.message);
    }
    setAuthBusy(false);
  }

  async function handleJoinOrCreateHousehold() {
    setAuthError(null);
    setAuthBusy(true);
    try {
      if (authInviteCode.trim()) {
        await sb("rpc/join_household_by_code", {
          method: "POST",
          body: JSON.stringify({ p_code: authInviteCode.trim(), p_display_name: authDisplayName.trim() || "Yo" }),
        });
      } else {
        await sb("rpc/create_household", {
          method: "POST",
          body: JSON.stringify({ p_name: authHouseholdName.trim() || "Mi hogar", p_display_name: authDisplayName.trim() || "Yo" }),
        });
      }
      await cargarHogarYDatos(session);
      setDesbloqueado(true);
    } catch (e) {
      setAuthError(e.message);
    }
    setAuthBusy(false);
  }

  function handleLogout() {
    safeSet("auth_session", null);
    setAccessToken(null);
    setRefreshToken(null);
    setSession(null);
    setHouseholdId(null);
    setProfileName(null);
    setEntries([]);
    setBudgets({});
    setCategoryOverrides({});
    setRecurrentes([]);
    setAccesosRapidos([]);
    setEsAdmin(false);
    setAuthMode("login");
  }

  async function addEntry(entry) {
    const full = { ...entry, id: uid(), who: profileName };
    setEntries((prev) => {
      const next = [full, ...prev];
      if (!HAS_SUPABASE) mockSaveEntries(next);
      return next;
    });
    if (!HAS_SUPABASE) return full;
    await conTimeout(
      sb("entries", { method: "POST", body: JSON.stringify([{ ...entryToDb(full), household_id: householdId }]) }),
      20000,
      "Guardar tardó demasiado (más de 20s). Puede ser un problema de conexión — probá de nuevo."
    );
    return full;
  }

  async function logAudit(accion, entrySnapshot, valorAnterior, valorNuevo) {
    const row = {
      id: uid(),
      entry_id: entrySnapshot?.id || null,
      accion,
      valor_anterior: valorAnterior != null ? String(valorAnterior) : null,
      valor_nuevo: valorNuevo != null ? String(valorNuevo) : null,
      entry_snapshot: entrySnapshot || null,
      who: profileName,
      household_id: householdId,
    };
    if (!HAS_SUPABASE) {
      const log = safeGet("mock_audit_log") || [];
      safeSet("mock_audit_log", [{ ...row, created_at: new Date().toISOString() }, ...log]);
      return;
    }
    try {
      await sb("audit_log", { method: "POST", body: JSON.stringify([row]) });
    } catch (e) {
      console.error("No se pudo guardar el historial", e);
    }
  }

  async function deleteEntry(id) {
    const anterior = entries.find((e) => e.id === id);
    setEntries((prev) => {
      const next = prev.filter((e) => e.id !== id);
      if (!HAS_SUPABASE) mockSaveEntries(next);
      return next;
    });
    await logAudit("delete", anterior, null, null);
    if (!HAS_SUPABASE) return;
    await sb(`entries?id=eq.${encodeURIComponent(id)}`, { method: "DELETE" });
  }

  async function editEntryDesc(id, nuevoDesc) {
    const anterior = entries.find((e) => e.id === id);
    setEntries((prev) => {
      const next = prev.map((e) => e.id === id ? { ...e, desc: nuevoDesc } : e);
      if (!HAS_SUPABASE) mockSaveEntries(next);
      return next;
    });
    await logAudit("edit_desc", anterior, anterior?.desc || "", nuevoDesc);
    if (!HAS_SUPABASE) return;
    await sb(`entries?id=eq.${encodeURIComponent(id)}`, { method: "PATCH", body: JSON.stringify({ descripcion: nuevoDesc }) });
  }

  async function editEntryAmount(id, nuevoMonto) {
    const anterior = entries.find((e) => e.id === id);
    setEntries((prev) => {
      const next = prev.map((e) => e.id === id ? { ...e, amount: nuevoMonto } : e);
      if (!HAS_SUPABASE) mockSaveEntries(next);
      return next;
    });
    await logAudit("edit_monto", anterior, anterior?.amount, nuevoMonto);
    if (!HAS_SUPABASE) return;
    await sb(`entries?id=eq.${encodeURIComponent(id)}`, { method: "PATCH", body: JSON.stringify({ amount: nuevoMonto }) });
  }

  // Edición completa de un movimiento (categoría, descripción, monto,
  // fecha, cuenta, pagado) en un solo viaje a la base — usada por el
  // modal de edición de Movimientos.
  async function editEntryFull(id, cambios) {
    const anterior = entries.find((e) => e.id === id);
    setEntries((prev) => {
      const next = prev.map((e) => e.id === id ? { ...e, ...cambios } : e);
      if (!HAS_SUPABASE) mockSaveEntries(next);
      return next;
    });
    await logAudit("edit_completo", anterior, anterior, cambios);
    if (!HAS_SUPABASE) return;
    const body = {};
    if (cambios.category !== undefined) body.category = cambios.category;
    if (cambios.desc !== undefined) body.descripcion = cambios.desc;
    if (cambios.amount !== undefined) body.amount = Number(cambios.amount);
    if (cambios.date !== undefined) body.date = cambios.date;
    if (cambios.account !== undefined) body.account = cambios.account;
    if (cambios.pagado !== undefined) body.pagado = cambios.pagado;
    await sb(`entries?id=eq.${encodeURIComponent(id)}`, { method: "PATCH", body: JSON.stringify(body) });
  }

  async function toggleEntryPagado(id, nuevoValor) {
    const anterior = entries.find((e) => e.id === id);
    setEntries((prev) => {
      const next = prev.map((e) => e.id === id ? { ...e, pagado: nuevoValor } : e);
      if (!HAS_SUPABASE) mockSaveEntries(next);
      return next;
    });
    await logAudit("marcar_pagado", anterior, anterior?.pagado, nuevoValor);
    if (!HAS_SUPABASE) return;
    await sb(`entries?id=eq.${encodeURIComponent(id)}`, { method: "PATCH", body: JSON.stringify({ pagado: nuevoValor }) });
  }

  // Confirma que un gasto "Pendiente" ya se pagó, apoyándose en que
  // apareció un movimiento igual (mismo monto, descripción parecida) en
  // una importación bancaria: marca el original como pagado y borra el
  // duplicado bancario, para no contar el mismo gasto dos veces.
  async function confirmarConciliacion(pendienteId, pagoId) {
    await toggleEntryPagado(pendienteId, true);
    await deleteEntry(pagoId);
  }

  async function addCategory(nombre) {
    const limpio = nombre.trim();
    if (!limpio || categories.includes(limpio)) return { ok: false, error: "Nombre vacío o ya existe." };
    const next = [...categories, limpio].sort();
    setCategories(next);
    if (!HAS_SUPABASE) { mockSaveCategories(next); return { ok: true }; }
    await sb("categories", { method: "POST", body: JSON.stringify([{ household_id: householdId, name: limpio }]) });
    return { ok: true };
  }

  function contarUsos(nombreCategoria) {
    return entries.filter((e) => e.category === nombreCategoria).length
      + Object.keys(budgets).filter((c) => c === nombreCategoria).length;
  }

  async function renameCategory(viejo, nuevo) {
    const limpio = nuevo.trim();
    if (!limpio || limpio === viejo) return { ok: false, error: "Nombre inválido." };
    if (categories.includes(limpio)) return { ok: false, error: "Ya existe una categoría con ese nombre." };
    // Renombra la categoría y todo lo que la referencia (movimientos,
    // presupuesto, reglas aprendidas) para no dejar nada huérfano.
    const nextEntries = entries.map((e) => e.category === viejo ? { ...e, category: limpio } : e);
    setEntries(nextEntries);
    const nextBudgets = { ...budgets };
    if (nextBudgets[viejo] != null) { nextBudgets[limpio] = nextBudgets[viejo]; delete nextBudgets[viejo]; }
    setBudgets(nextBudgets);
    const nextOverrides = { ...categoryOverrides };
    Object.keys(nextOverrides).forEach((k) => { if (nextOverrides[k] === viejo) nextOverrides[k] = limpio; });
    setCategoryOverrides(nextOverrides);
    const nextCats = categories.map((c) => c === viejo ? limpio : c).sort();
    setCategories(nextCats);

    if (!HAS_SUPABASE) {
      mockSaveEntries(nextEntries);
      mockSaveBudgets(nextBudgets);
      mockSaveOverrides(nextOverrides);
      mockSaveCategories(nextCats);
      return { ok: true };
    }
    await Promise.all([
      sb(`entries?household_id=eq.${householdId}&category=eq.${encodeURIComponent(viejo)}`, { method: "PATCH", body: JSON.stringify({ category: limpio }) }),
      sb(`budgets?household_id=eq.${householdId}&category=eq.${encodeURIComponent(viejo)}`, { method: "PATCH", body: JSON.stringify({ category: limpio }) }),
      sb(`category_overrides?household_id=eq.${householdId}&category=eq.${encodeURIComponent(viejo)}`, { method: "PATCH", body: JSON.stringify({ category: limpio }) }),
      sb(`categories?household_id=eq.${householdId}&name=eq.${encodeURIComponent(viejo)}`, { method: "PATCH", body: JSON.stringify({ name: limpio }) }),
    ]);
    return { ok: true };
  }

  async function deleteCategory(nombre) {
    const usos = contarUsos(nombre);
    if (usos > 0) {
      return { ok: false, error: `Hay ${usos} movimiento(s) usando "${nombre}" — recategorizalos primero.` };
    }
    const next = categories.filter((c) => c !== nombre);
    setCategories(next);
    if (!HAS_SUPABASE) { mockSaveCategories(next); return { ok: true }; }
    await sb(`categories?household_id=eq.${householdId}&name=eq.${encodeURIComponent(nombre)}`, { method: "DELETE" });
    return { ok: true };
  }

  async function resetearReglasAprendidas() {
    setCategoryOverrides({});
    if (!HAS_SUPABASE) { mockSaveOverrides({}); return; }
    await sb("category_overrides?desc_key=not.is.null", { method: "DELETE" });
  }

  async function resetearTodo() {
    setEntries([]);
    setBudgets({});
    setRecurrentes([]);
    setAccesosRapidos([]);
    if (!HAS_SUPABASE) {
      mockSaveEntries([]);
      mockSaveBudgets({});
      mockSaveRecurrentes([]);
      mockSaveAccesosRapidos([]);
      safeSet("mock_audit_log", []);
      return;
    }
    await Promise.all([
      sb("entries?id=not.is.null", { method: "DELETE" }),
      sb("budgets?category=not.is.null", { method: "DELETE" }),
      sb("recurring_entries?id=not.is.null", { method: "DELETE" }),
      sb("quick_entries?id=not.is.null", { method: "DELETE" }),
      sb("audit_log?id=not.is.null", { method: "DELETE" }),
    ]);
  }

  async function updateBudgets(next) {
    setBudgets(next);
    if (!HAS_SUPABASE) { mockSaveBudgets(next); return; }
    const rows = Object.entries(next).map(([category, limit_amount]) => ({ category, limit_amount: Number(limit_amount), household_id: householdId }));
    if (rows.length > 0) {
      await sb("budgets", { method: "POST", headers: { Prefer: "resolution=merge-duplicates" }, body: JSON.stringify(rows) });
    }
  }

  async function addRecurrente(datos) {
    const nuevo = {
      id: uidUuid(),
      type: datos.type,
      category: datos.category,
      amount: Number(datos.amount) || 0,
      descripcion: datos.desc || "",
      account: datos.account || "",
      moneda: datos.moneda || "ARS",
      dia_mes: Number(datos.diaMes) || 1,
      activo: true,
    };
    const next = [...recurrentes, nuevo];
    setRecurrentes(next);
    if (!HAS_SUPABASE) { mockSaveRecurrentes(next); return { ok: true }; }
    try {
      await sb("recurring_entries", { method: "POST", body: JSON.stringify([{ ...nuevo, household_id: householdId }]) });
      return { ok: true };
    } catch (e) {
      setRecurrentes(recurrentes); // revertir: no se guardó de verdad
      return { error: "No se pudo guardar: " + e.message };
    }
  }

  async function toggleActivoRecurrente(id, activo) {
    const next = recurrentes.map((r) => (r.id === id ? { ...r, activo } : r));
    setRecurrentes(next);
    if (!HAS_SUPABASE) { mockSaveRecurrentes(next); return; }
    await sb(`recurring_entries?id=eq.${id}`, { method: "PATCH", body: JSON.stringify({ activo }) });
  }

  async function editMontoRecurrente(id, amount) {
    const next = recurrentes.map((r) => (r.id === id ? { ...r, amount: Number(amount) || 0 } : r));
    setRecurrentes(next);
    if (!HAS_SUPABASE) { mockSaveRecurrentes(next); return; }
    await sb(`recurring_entries?id=eq.${id}`, { method: "PATCH", body: JSON.stringify({ amount: Number(amount) || 0 }) });
  }

  async function deleteRecurrente(id) {
    const next = recurrentes.filter((r) => r.id !== id);
    setRecurrentes(next);
    if (!HAS_SUPABASE) { mockSaveRecurrentes(next); return; }
    await sb(`recurring_entries?id=eq.${id}`, { method: "DELETE" });
  }

  async function addAccesoRapido(datos) {
    const nuevo = {
      id: uidUuid(),
      type: "gasto",
      category: datos.category,
      amount: Number(datos.amount) || 0,
      descripcion: datos.desc || "",
      account: datos.account || "",
      moneda: "ARS",
      sort_order: accesosRapidos.length,
      personal: datos.personal !== false,
    };
    const next = [...accesosRapidos, nuevo];
    setAccesosRapidos(next);
    if (!HAS_SUPABASE) { mockSaveAccesosRapidos(next); return { ok: true }; }
    try {
      await sb("quick_entries", { method: "POST", body: JSON.stringify([{ ...nuevo, household_id: householdId }]) });
      return { ok: true };
    } catch (e) {
      setAccesosRapidos(accesosRapidos); // revertir: no se guardó de verdad
      return { error: "No se pudo guardar: " + e.message };
    }
  }

  async function editAccesoRapido(id, cambios) {
    const next = accesosRapidos.map((a) => (a.id === id ? { ...a, ...cambios } : a));
    setAccesosRapidos(next);
    if (!HAS_SUPABASE) { mockSaveAccesosRapidos(next); return; }
    await sb(`quick_entries?id=eq.${id}`, { method: "PATCH", body: JSON.stringify(cambios) });
  }

  async function deleteAccesoRapido(id) {
    const next = accesosRapidos.filter((a) => a.id !== id);
    setAccesosRapidos(next);
    if (!HAS_SUPABASE) { mockSaveAccesosRapidos(next); return; }
    await sb(`quick_entries?id=eq.${id}`, { method: "DELETE" });
  }

  async function renombrarMiembro(nombreViejo, nombreNuevo) {
    const limpio = (nombreNuevo || "").trim();
    if (!limpio) return { error: "Poné un nombre." };
    if (limpio === nombreViejo) return { ok: true };
    if (!HAS_SUPABASE) return { error: "No aplica en vista previa local." };
    try {
      await sb("rpc/renombrar_miembro", { method: "POST", body: JSON.stringify({ p_nombre_viejo: nombreViejo, p_nombre_nuevo: limpio }) });
      await refrescarEntries();
      const rows = await sb("rpc/get_my_household_members", { method: "POST", body: "{}" }).catch(() => null);
      if (rows) setMiembrosHogar(rows.map((r) => r.display_name).filter(Boolean));
      return { ok: true, miembros: rows };
    } catch (e) {
      return { error: e.message || "No se pudo renombrar." };
    }
  }

  // Unifica/renombra una cuenta (ej. "Visa Signature" y "Visa BBVA Hernán"
  // son la misma tarjeta) — arrastra el cambio a TODOS los movimientos
  // viejos (de cualquier mes) que tenían el nombre anterior guardado.
  async function renombrarCuenta(nombreViejo, nombreNuevo) {
    const limpio = (nombreNuevo || "").trim();
    if (!limpio) return { error: "Poné un nombre." };
    if (limpio === nombreViejo) return { ok: true };
    if (!HAS_SUPABASE) return { error: "No aplica en vista previa local." };
    try {
      await sb(`entries?household_id=eq.${householdId}&account=eq.${encodeURIComponent(nombreViejo)}`, {
        method: "PATCH", body: JSON.stringify({ account: limpio }),
      });
      await refrescarEntries();
      return { ok: true };
    } catch (e) {
      return { error: e.message || "No se pudo renombrar la cuenta." };
    }
  }

  async function addWhatsappLink(phoneRaw, nombre) {
    const phone = (phoneRaw || "").replace(/[^\d]/g, ""); // solo dígitos, formato E.164 sin "+"
    if (!phone || phone.length < 10) return { error: "Poné el número completo, con código de país (ej: 5491122334455)." };
    if (!HAS_SUPABASE) return { error: "No aplica en vista previa local." };
    try {
      const nuevo = await sb("whatsapp_links", {
        method: "POST",
        headers: { Prefer: "return=representation" },
        body: JSON.stringify([{ household_id: householdId, phone_number: phone, display_name: nombre || profileName }]),
      });
      setWhatsappLinks((prev) => [...prev, ...(nuevo || [])]);
      return { ok: true };
    } catch (e) {
      return { error: e.message.includes("duplicate") ? "Ese número ya está vinculado." : "No se pudo vincular: " + e.message };
    }
  }

  async function deleteWhatsappLink(id) {
    setWhatsappLinks((prev) => prev.filter((w) => w.id !== id));
    if (!HAS_SUPABASE) return;
    await sb(`whatsapp_links?id=eq.${id}`, { method: "DELETE" });
  }

  // Un toque en un acceso rápido carga el gasto ya, con la fecha de hoy —
  // muestra un toast con "Deshacer" por unos segundos por si fue sin querer.
  async function registrarAccesoRapido(acceso) {
    const entry = await addEntry({
      type: acceso.type || "gasto",
      category: acceso.category,
      amount: Number(acceso.amount),
      desc: acceso.descripcion || acceso.category,
      date: todayISO(),
      account: acceso.account || "",
      moneda: acceso.moneda || "ARS",
      origen: "acceso_rapido",
    });
    setUltimoAccesoRegistrado({ id: entry.id, nombre: acceso.descripcion || acceso.category });
    setTimeout(() => {
      setUltimoAccesoRegistrado((cur) => (cur?.id === entry.id ? null : cur));
    }, 6000);
  }

  function deshacerUltimoAcceso() {
    if (!ultimoAccesoRegistrado) return;
    deleteEntry(ultimoAccesoRegistrado.id);
    setUltimoAccesoRegistrado(null);
  }

  // Wrapper liviano: el botón de Excel del header siempre exporta TODO.
  function exportarExcel() {
    exportarEntriesAExcel(entries, "movimientos");
  }

  const now = new Date();
  const realThisMonth = now.toISOString().slice(0, 7);
  const [selectedMonth, setSelectedMonth] = useState(realThisMonth);
  const thisMonth = selectedMonth;

  function shiftMonth(delta) {
    const [y, m] = selectedMonth.split("-").map(Number);
    const d = new Date(y, m - 1 + delta, 1);
    setSelectedMonth(d.toISOString().slice(0, 7));
  }
  function monthLabel(key) {
    const [y, m] = key.split("-").map(Number);
    return new Date(y, m - 1, 1).toLocaleDateString("es-AR", { month: "long", year: "numeric" });
  }

  const thisMonthEntries = useMemo(
    () => entries.filter((e) => monthKey(e.date) === thisMonth),
    [entries, thisMonth]
  );
  const totalIngresos = useMemo(
    () => thisMonthEntries
      .filter((e) => e.type === "ingreso" || e.type === "cambio")
      .reduce((s, e) => s + Number(e.amount), 0),
    [thisMonthEntries]
  );
  const totalGastos = useMemo(
    () => thisMonthEntries.filter((e) => e.type === "gasto" && (e.moneda || "ARS") === "ARS").reduce((s, e) => s + Number(e.amount), 0),
    [thisMonthEntries]
  );
  const totalGastosUsd = useMemo(
    () => thisMonthEntries.filter((e) => e.type === "gasto" && e.moneda === "USD").reduce((s, e) => s + Number(e.amount), 0),
    [thisMonthEntries]
  );
  const totalAhorro = useMemo(
    () => thisMonthEntries.filter((e) => e.type === "ahorro").reduce((s, e) => s + Number(e.amount), 0),
    [thisMonthEntries]
  );
  const balance = totalIngresos - totalGastos - totalAhorro;

  const totalAhorradoHistorico = useMemo(
    () => entries.filter((e) => e.type === "ahorro").reduce((s, e) => s + Number(e.amount), 0),
    [entries]
  );

  const cambios = useMemo(
    () => entries.filter((e) => e.type === "cambio").sort((a, b) => b.date.localeCompare(a.date)),
    [entries]
  );
  const cambiosStats = useMemo(() => {
    if (cambios.length === 0) return null;
    const totalUsd = cambios.reduce((s, e) => s + Number(e.usdAmount), 0);
    const totalArs = cambios.reduce((s, e) => s + Number(e.amount), 0);
    const esArq = (e) => (e.account || "").trim().toUpperCase() === "ARQ";
    const arq = cambios.filter(esArq);
    const externos = cambios.filter((e) => !esArq(e));
    const sum = (arr, key) => arr.reduce((s, e) => s + Number(e[key]), 0);
    return {
      ultimo: cambios[0], promedio: totalArs / totalUsd, totalUsd, totalArs,
      arq: { totalUsd: sum(arq, "usdAmount"), totalArs: sum(arq, "amount"), count: arq.length },
      externos: { totalUsd: sum(externos, "usdAmount"), totalArs: sum(externos, "amount"), count: externos.length },
    };
  }, [cambios]);

  const gastosPorCategoria = useMemo(() => {
    const map = {};
    thisMonthEntries.filter((e) => e.type === "gasto" && (e.moneda || "ARS") === "ARS").forEach((e) => {
      map[e.category] = (map[e.category] || 0) + Number(e.amount);
    });
    return Object.entries(map).map(([name, value]) => ({ name, value })).sort((a, b) => b.value - a.value);
  }, [thisMonthEntries]);

  // "Panel admin" solo se ofrece si la función soy_admin() confirmó que
  // este usuario está en app_admins — el chequeo real y no salteable pasa
  // igual del lado del servidor cuando se pide el contenido.
  const menuTabs = esAdmin ? [...SECONDARY_TABS, "admin"] : SECONDARY_TABS;

  if (loading) {
    return (
      <div style={{ minHeight: "100vh", background: PAPER, display: "flex", alignItems: "center", justifyContent: "center", fontFamily: "Inter, sans-serif", color: INK }}>
        Cargando...
      </div>
    );
  }

  if (loadError) {
    return (
      <div style={{ minHeight: "100vh", background: PAPER, display: "flex", flexDirection: "column", alignItems: "center", justifyContent: "center", fontFamily: "Inter, sans-serif", color: BRICK, padding: 24, textAlign: "center", whiteSpace: "pre-wrap", gap: 16 }}>
        <div>{loadError}</div>
        <button
          onClick={() => { setLoading(true); setLoadError(null); intentarRestaurarSesion(); }}
          style={{ ...btnPrimary, padding: "10px 20px" }}
        >
          Reintentar
        </button>
      </div>
    );
  }

  if (!profileName) {
    // --- Modo local (preview sin Supabase): mantenemos el selector simple ---
    if (!HAS_SUPABASE) {
      return (
        <div style={{ minHeight: "100vh", background: INK, display: "flex", alignItems: "center", justifyContent: "center", padding: 24, fontFamily: "Inter, sans-serif" }}>
          <style>{fontImports}</style>
          <div style={{ background: PAPER, borderRadius: 4, padding: "40px 32px", maxWidth: 380, width: "100%", boxShadow: "0 20px 60px rgba(0,0,0,0.4)" }}>
            <div style={{ fontFamily: "'Fraunces', serif", fontSize: 28, fontWeight: 600, color: INK, marginBottom: 6 }}>Finanzas del hogar</div>
            <div style={{ color: "#5a6b6d", fontSize: 14, marginBottom: 24 }}>Vista previa local. Decinos quién sos para empezar.</div>
            <div style={{ display: "flex", gap: 8 }}>
              <input
                value={authDisplayName}
                onChange={(e) => setAuthDisplayName(e.target.value)}
                onKeyDown={(e) => e.key === "Enter" && authDisplayName.trim() && (safeSet("profile", { name: authDisplayName.trim() }), setProfileName(authDisplayName.trim()))}
                placeholder="Ej: Juan"
                style={inputStyle}
              />
              <button onClick={() => { if (authDisplayName.trim()) { safeSet("profile", { name: authDisplayName.trim() }); setProfileName(authDisplayName.trim()); } }} style={btnPrimary}>Entrar</button>
            </div>
            <div style={{ fontSize: 11, color: GOLD, marginTop: 12, textAlign: "right" }}>⚠ Vista previa local — sin conexión a Supabase (normal en Claude)</div>
            <div style={{ fontSize: 10, color: "#c4bda8", marginTop: 18, textAlign: "right" }}>{APP_VERSION}</div>
          </div>
        </div>
      );
    }

    // --- Onboarding: ya hay sesión pero falta crear/sumarse a un hogar ---
    if (authMode === "onboarding" && session) {
      return (
        <div style={{ minHeight: "100vh", background: INK, display: "flex", alignItems: "center", justifyContent: "center", padding: 24, fontFamily: "Inter, sans-serif" }}>
          <style>{fontImports}</style>
          <div style={{ background: PAPER, borderRadius: 4, padding: "40px 32px", maxWidth: 400, width: "100%", boxShadow: "0 20px 60px rgba(0,0,0,0.4)" }}>
            <div style={{ fontFamily: "'Fraunces', serif", fontSize: 24, fontWeight: 600, color: INK, marginBottom: 6 }}>Un paso más</div>
            <div style={{ color: "#5a6b6d", fontSize: 13.5, marginBottom: 20 }}>
              Creá tu hogar, o sumate a uno existente con un código de invitación.
            </div>
            <label style={labelStyle}>Tu nombre (como te van a ver los demás)</label>
            <input value={authDisplayName} onChange={(e) => setAuthDisplayName(e.target.value)} placeholder="Ej: Juan" style={{ ...inputStyle, marginBottom: 14 }} />
            <label style={labelStyle}>Código de invitación (si te sumás a un hogar existente)</label>
            <input value={authInviteCode} onChange={(e) => setAuthInviteCode(e.target.value)} placeholder="Dejalo vacío para crear un hogar nuevo" style={{ ...inputStyle, marginBottom: 14 }} />
            {!authInviteCode.trim() && (
              <>
                <label style={labelStyle}>Nombre del hogar</label>
                <input value={authHouseholdName} onChange={(e) => setAuthHouseholdName(e.target.value)} placeholder="Ej: Familia Israel" style={{ ...inputStyle, marginBottom: 14 }} />
              </>
            )}
            {authError && <div style={{ color: BRICK, fontSize: 12.5, marginBottom: 12 }}>{authError}</div>}
            <button onClick={handleJoinOrCreateHousehold} disabled={authBusy} style={{ ...btnPrimary, width: "100%", justifyContent: "center" }}>
              {authBusy ? "Un momento..." : (authInviteCode.trim() ? "Sumarme al hogar" : "Crear mi hogar")}
            </button>
          </div>
        </div>
      );
    }

    // --- Login / registro ---
    return (
      <div style={{ minHeight: "100vh", background: INK, display: "flex", alignItems: "center", justifyContent: "center", padding: 24, fontFamily: "Inter, sans-serif" }}>
        <style>{fontImports}</style>
        <div style={{ background: PAPER, borderRadius: 4, padding: "40px 32px", maxWidth: 400, width: "100%", boxShadow: "0 20px 60px rgba(0,0,0,0.4)" }}>
          <div style={{ fontFamily: "'Fraunces', serif", fontSize: 28, fontWeight: 600, color: INK, marginBottom: 6 }}>Finanzas del hogar</div>
          <div style={{ color: "#5a6b6d", fontSize: 14, marginBottom: 20 }}>
            {authMode === "signup" ? "Creá tu cuenta para empezar." : "Iniciá sesión para continuar."}
          </div>

          <label style={labelStyle}>Email</label>
          <input type="email" value={authEmail} onChange={(e) => setAuthEmail(e.target.value)} placeholder="vos@ejemplo.com" style={{ ...inputStyle, marginBottom: 14 }} />
          <label style={labelStyle}>Contraseña</label>
          <input type="password" value={authPassword} onChange={(e) => setAuthPassword(e.target.value)} placeholder="••••••••" style={{ ...inputStyle, marginBottom: 14 }} />

          {authMode === "signup" && (
            <>
              <label style={labelStyle}>Tu nombre</label>
              <input value={authDisplayName} onChange={(e) => setAuthDisplayName(e.target.value)} placeholder="Ej: Juan" style={{ ...inputStyle, marginBottom: 14 }} />
              <label style={labelStyle}>Código de invitación (opcional, si te suma a un hogar existente)</label>
              <input value={authInviteCode} onChange={(e) => setAuthInviteCode(e.target.value)} placeholder="Dejalo vacío para crear tu propio hogar" style={{ ...inputStyle, marginBottom: 14 }} />
              {!authInviteCode.trim() && (
                <>
                  <label style={labelStyle}>Nombre del hogar</label>
                  <input value={authHouseholdName} onChange={(e) => setAuthHouseholdName(e.target.value)} placeholder="Ej: Familia Israel" style={{ ...inputStyle, marginBottom: 14 }} />
                </>
              )}
            </>
          )}

          {authError && <div style={{ color: BRICK, fontSize: 12.5, marginBottom: 12 }}>{authError}</div>}

          <button onClick={authMode === "signup" ? handleSignup : handleLogin} disabled={authBusy} style={{ ...btnPrimary, width: "100%", justifyContent: "center", marginBottom: 12 }}>
            {authBusy ? "Un momento..." : (authMode === "signup" ? "Crear cuenta" : "Iniciar sesión")}
          </button>
          <button
            onClick={() => { setAuthMode(authMode === "signup" ? "login" : "signup"); setAuthError(null); }}
            style={{ background: "none", border: "none", color: TEAL, fontSize: 13, cursor: "pointer", width: "100%" }}
          >
            {authMode === "signup" ? "¿Ya tenés cuenta? Iniciá sesión" : "¿No tenés cuenta? Creá una"}
          </button>
          <div style={{ fontSize: 10, color: "#c4bda8", marginTop: 18, textAlign: "right" }}>{APP_VERSION}</div>
        </div>
      </div>
    );
  }

  // Sesión restaurada en silencio (la persona no acaba de tipear su
  // contraseña recién) + tiene activado el desbloqueo biométrico en este
  // dispositivo: pedimos esa confirmación antes de mostrar los datos.
  if (biometriaActiva && !desbloqueado) {
    return (
      <div style={{ minHeight: "100vh", background: INK, display: "flex", alignItems: "center", justifyContent: "center", padding: 24, fontFamily: "Inter, sans-serif" }}>
        <style>{fontImports}</style>
        <div style={{ background: PAPER, borderRadius: 4, padding: "40px 32px", maxWidth: 380, width: "100%", boxShadow: "0 20px 60px rgba(0,0,0,0.4)", textAlign: "center" }}>
          <div style={{ fontFamily: "'Fraunces', serif", fontSize: 24, fontWeight: 600, color: INK, marginBottom: 6 }}>Finanzas del hogar</div>
          <div style={{ color: "#5a6b6d", fontSize: 14, marginBottom: 24 }}>Hola, {profileName}. Desbloqueá para ver tus datos.</div>
          {bioError && <div style={{ color: BRICK, fontSize: 12.5, marginBottom: 14 }}>{bioError}</div>}
          <button onClick={intentarDesbloquear} disabled={bioBusy} style={{ ...btnPrimary, width: "100%", justifyContent: "center", padding: "13px" }}>
            {bioBusy ? "Verificando..." : "Desbloquear con Face ID / huella"}
          </button>
          <button onClick={handleLogout} style={{ background: "none", border: "none", color: "#8a9698", fontSize: 12.5, marginTop: 16, cursor: "pointer" }}>
            Cerrar sesión
          </button>
        </div>
      </div>
    );
  }

  return (
    <div style={{ minHeight: "100vh", background: PAPER, fontFamily: "Inter, sans-serif", color: INK, paddingBottom: isDesktop ? 100 : 150, marginLeft: isDesktop ? 230 : 0 }}>
      <style>{fontImports}</style>
      <style>{`
        * { box-sizing: border-box; }
        button:focus-visible, input:focus-visible, select:focus-visible { outline: 2px solid ${TEAL}; outline-offset: 2px; }
        .tabbar button { transition: color 0.15s, border-color 0.15s; }
        .cat-row { transition: background 0.15s; }
        .cat-row:hover { background: ${PAPER_DIM}; }
        ::-webkit-scrollbar { height: 6px; width: 6px; }
      `}</style>

      {hayVersionNueva && (
        <div style={{ position: "sticky", top: 0, zIndex: 50, background: GOLD, color: "#fff", padding: "9px 16px", display: "flex", alignItems: "center", justifyContent: "center", gap: 10, flexWrap: "wrap", fontSize: 13, fontWeight: 600 }}>
          🔄 Hay una versión nueva de la app.
          <button
            onClick={() => window.location.reload()}
            style={{ background: "#fff", color: GOLD, border: "none", borderRadius: 6, padding: "4px 12px", fontSize: 12.5, fontWeight: 700, cursor: "pointer" }}
          >
            Actualizar ahora
          </button>
        </div>
      )}

      {/* Header */}
      <div style={{ background: INK, color: PAPER, padding: "20px 20px 26px" }}>
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start", maxWidth: isDesktop ? 1080 : 720, margin: "0 auto" }}>
          <div>
            <div style={{ display: "flex", alignItems: "center", gap: 10 }}>
              {!isDesktop && (
                <button
                  onClick={() => setMenuOpen(true)}
                  aria-label="Abrir menú"
                  style={{ background: "none", border: "1px solid rgba(255,255,255,0.3)", borderRadius: 8, width: 32, height: 32, display: "flex", alignItems: "center", justifyContent: "center", cursor: "pointer", color: PAPER, flexShrink: 0 }}
                >
                  <Menu size={18} />
                </button>
              )}
              <div style={{ fontFamily: "'Fraunces', serif", fontSize: 24, fontWeight: 600 }}>Finanzas del hogar</div>
            </div>
            <div style={{ fontSize: 12, color: "#9db3b0", marginTop: 2, marginLeft: isDesktop ? 0 : 42 }}>
              Hola, {profileName}
            </div>
            <div style={{ display: "flex", alignItems: "center", gap: 4, rowGap: 6, marginTop: 8, marginLeft: isDesktop ? 0 : 42, flexWrap: "wrap" }}>
              <button
                onClick={() => shiftMonth(-1)}
                aria-label="Mes anterior"
                style={{ background: "rgba(255,255,255,0.08)", border: "1px solid rgba(255,255,255,0.25)", borderRadius: 8, cursor: "pointer", color: PAPER, width: 36, height: 36, display: "flex", alignItems: "center", justifyContent: "center", fontSize: 16, flexShrink: 0 }}
              >◀</button>
              <span style={{ fontSize: 14, fontWeight: 700, textTransform: "capitalize", minWidth: 108, textAlign: "center", flexShrink: 0 }}>{monthLabel(selectedMonth)}</span>
              <button
                onClick={() => shiftMonth(1)}
                aria-label="Mes siguiente"
                style={{ background: "rgba(255,255,255,0.08)", border: "1px solid rgba(255,255,255,0.25)", borderRadius: 8, cursor: "pointer", color: PAPER, width: 36, height: 36, display: "flex", alignItems: "center", justifyContent: "center", fontSize: 16, flexShrink: 0 }}
              >▶</button>
              {selectedMonth !== realThisMonth && (
                <button onClick={() => setSelectedMonth(realThisMonth)} style={{ background: "none", border: `1px solid ${TEAL}`, borderRadius: 6, color: TEAL, fontSize: 11, padding: "4px 8px", marginLeft: 2, cursor: "pointer", height: 36, flexShrink: 0 }}>
                  Hoy
                </button>
              )}
            </div>
            <div style={{ fontSize: 9.5, color: "#5f7376", marginTop: 6, marginLeft: isDesktop ? 0 : 42 }}>{APP_VERSION}</div>
            {!HAS_SUPABASE && (
              <div style={{ fontSize: 10.5, color: GOLD, marginTop: 4, fontWeight: 700, marginLeft: isDesktop ? 0 : 42 }}>
                ⚠ Vista previa local — no conectado a Supabase
              </div>
            )}
          </div>
          <div style={{ display: "flex", flexDirection: "column", alignItems: "flex-end", gap: 6 }}>
            <div title={profileName} style={{
              width: 28, height: 28, borderRadius: "50%",
              background: GOLD, color: INK,
              display: "flex", alignItems: "center", justifyContent: "center",
              fontSize: 12, fontWeight: 700
            }}>{profileName?.[0]?.toUpperCase()}</div>
            <button
              onClick={exportarExcel}
              title="Exportar todos los movimientos a Excel"
              style={{
                display: "flex", alignItems: "center", gap: 5, marginTop: 40,
                background: "none", border: "1px solid rgba(255,255,255,0.3)", borderRadius: 6,
                color: PAPER, fontSize: 10.5, padding: "3px 8px", cursor: "pointer",
              }}
            >
              <Download size={12} /> Excel
            </button>
          </div>
        </div>
      </div>

      {/* Menú lateral (drawer) — solo mobile. En desktop la navegación vive en la sidebar fija. */}
      {!isDesktop && (
      <>
      {menuOpen && (
        <div onClick={() => setMenuOpen(false)} style={{ position: "fixed", inset: 0, background: "rgba(27,42,46,0.45)", zIndex: 30 }} />
      )}
      <div
        style={{
          position: "fixed", top: 0, left: 0, bottom: 0, width: 280, maxWidth: "82vw",
          background: "#fff", zIndex: 31, display: "flex", flexDirection: "column", overflowY: "auto",
          boxShadow: menuOpen ? "4px 0 24px rgba(0,0,0,0.25)" : "none",
          transform: menuOpen ? "translateX(0)" : "translateX(-100%)",
          transition: "transform 0.28s ease",
        }}
      >
        <div style={{ background: INK, color: PAPER, padding: "20px 18px", display: "flex", justifyContent: "space-between", alignItems: "center", flexShrink: 0 }}>
          <div style={{ fontFamily: "'Fraunces', serif", fontSize: 18 }}>Finanzas del hogar</div>
          <button onClick={() => setMenuOpen(false)} aria-label="Cerrar menú" style={{ background: "none", border: "none", color: PAPER, cursor: "pointer", display: "flex" }}>
            <X size={20} />
          </button>
        </div>
        <div style={{ padding: "8px 0", flex: 1 }}>
          {menuTabs.map((key) => (
            <button
              key={key}
              onClick={() => { setTab(key); setMenuOpen(false); }}
              style={{
                display: "block", width: "100%", textAlign: "left", padding: "13px 20px",
                background: tab === key ? PAPER_DIM : "transparent",
                border: "none", borderLeft: `3px solid ${tab === key ? TEAL : "transparent"}`,
                cursor: "pointer", fontSize: 14.5, fontWeight: tab === key ? 700 : 500,
                color: tab === key ? TEAL : INK,
              }}
            >
              {TAB_LABELS[key]}
            </button>
          ))}
        </div>
      </div>
      </>
      )}

      {/* Sidebar fija — solo desktop, estilo apps de escritorio (Mercado Pago, etc.) */}
      {isDesktop && (
        <div style={{
          position: "fixed", top: 0, left: 0, bottom: 0, width: 230,
          background: "#fff", borderRight: "1px solid #e6e0d0", zIndex: 15,
          display: "flex", flexDirection: "column", overflowY: "auto",
        }}>
          <div style={{ padding: "20px 18px", borderBottom: "1px solid #f0ece0" }}>
            <div style={{ fontFamily: "'Fraunces', serif", fontSize: 18 }}>Finanzas del hogar</div>
            <div style={{ fontSize: 11.5, color: "#8a9698", marginTop: 2 }}>Hola, {profileName}</div>
          </div>
          <div style={{ flex: 1, padding: "14px 10px", display: "flex", flexDirection: "column", gap: 18 }}>
            {[
              { label: null, items: PRIMARY_TABS },
              { label: "Organización", items: [
                ["categorias", TAB_LABELS.categorias, Tag],
                ["presupuestos", TAB_LABELS.presupuestos, PiggyBank],
                ["recurrentes", TAB_LABELS.recurrentes, Clock],
                ["cuentas", TAB_LABELS.cuentas, Wallet],
                ["personas", TAB_LABELS.personas, User],
              ] },
              { label: "Herramientas", items: [
                ["recategorizar", TAB_LABELS.recategorizar, ArrowLeftRight],
                ["conciliar", TAB_LABELS.conciliar, Check],
                ["duplicados", TAB_LABELS.duplicados, Copy],
                ["divisas", TAB_LABELS.divisas, Landmark],
                ["cotizaciones", TAB_LABELS.cotizaciones, TrendingUp],
              ] },
              { label: "Hogar", items: [
                ["hogar", TAB_LABELS.hogar, Settings],
                ["historial", TAB_LABELS.historial, History],
                ["ahorros", TAB_LABELS.ahorros, PiggyBank],
              ] },
              { label: "Sistema", items: [
                ["reset", TAB_LABELS.reset, Trash2],
                ...(esAdmin ? [["admin", TAB_LABELS.admin, Target]] : []),
              ] },
            ].map((grupo, gi) => (
              <div key={gi}>
                {grupo.label && (
                  <div style={{ fontSize: 10.5, fontWeight: 700, textTransform: "uppercase", letterSpacing: 0.5, color: "#8a9698", padding: "0 10px", marginBottom: 6 }}>
                    {grupo.label}
                  </div>
                )}
                <div style={{ display: "flex", flexDirection: "column", gap: 2 }}>
                  {grupo.items.map(([key, label, Icon]) => {
                    const active = tab === key;
                    return (
                      <button
                        key={key}
                        onClick={() => setTab(key)}
                        style={{
                          display: "flex", alignItems: "center", gap: 10, padding: "9px 12px", borderRadius: 8,
                          border: "none", background: active ? TEAL : "transparent", color: active ? "#fff" : INK,
                          cursor: "pointer", fontSize: 13.5, fontWeight: active ? 700 : 500, textAlign: "left",
                        }}
                      >
                        <Icon size={17} />
                        {label}
                      </button>
                    );
                  })}
                </div>
              </div>
            ))}
          </div>
          <div style={{ padding: "14px 18px", borderTop: "1px solid #f0ece0" }}>
            <button onClick={handleLogout} style={{ display: "flex", alignItems: "center", gap: 8, background: "none", border: "none", cursor: "pointer", color: BRICK, fontSize: 13, padding: 0 }}>
              <LogOut size={16} /> Cerrar sesión
            </button>
          </div>
        </div>
      )}

      <div style={{ maxWidth: isDesktop ? 1080 : 720, margin: "0 auto", padding: "0 16px" }}>
        {!isDesktop && biometriaSoportada && !biometriaActiva && !bannerBiometriaCerrado && (
          <div style={{ marginTop: 16, background: "#fff", border: `1.5px solid ${TEAL}`, borderRadius: 10, padding: "12px 14px", display: "flex", alignItems: "center", justifyContent: "space-between", gap: 10, flexWrap: "wrap" }}>
            <div style={{ fontSize: 12.5, color: INK }}>
              🔐 Este dispositivo tiene Biometria activada (Face ID / Touch ID) — activalo para abrir la app sin escribir usuario y contraseña cada vez.
            </div>
            <div style={{ display: "flex", alignItems: "center", gap: 8, flexShrink: 0 }}>
              {bioError && <div style={{ color: BRICK, fontSize: 11 }}>{bioError}</div>}
              <button
                onClick={activarBiometria}
                disabled={bioBusy}
                style={{ ...btnPrimary, padding: "6px 12px", fontSize: 12.5 }}
              >
                {bioBusy ? "Confirmando..." : "Activar ahora"}
              </button>
              <button
                onClick={() => { safeSet("banner_biometria_cerrado", true); setBannerBiometriaCerrado(true); }}
                style={{ background: "none", border: "none", color: "#8a9698", cursor: "pointer", fontSize: 18, lineHeight: 1, padding: "0 4px" }}
                aria-label="Cerrar aviso"
              >×</button>
            </div>
          </div>
        )}
        {menuTabs.includes(tab) ? (
          <div style={{ marginTop: 18, display: "flex", alignItems: "center", gap: 10 }}>
            <button
              onClick={() => setTab("resumen")}
              style={{ ...btnOutline, padding: "8px 10px", display: "flex", alignItems: "center", gap: 6 }}
              aria-label="Volver al Resumen"
            >
              <ArrowLeft size={16} /> Volver
            </button>
            <div style={{ fontFamily: "'Fraunces', serif", fontSize: 19 }}>{TAB_LABELS[tab]}</div>
          </div>
        ) : (
          <>
            {/* Barra de accesos (desktop): navegación + acciones rápidas, arriba del balance */}
            {isDesktop && (
              <div style={{ display: "flex", justifyContent: "center", gap: 8, marginTop: 18, marginBottom: 18, flexWrap: "wrap" }}>
                {[
                  ["resumen", "Resumen", Home, () => setTab("resumen")],
                  ["movimientos", "Movimientos", List, () => setTab("movimientos")],
                  ["rapidos", "Gastos rápidos", Zap, () => setTab("rapidos")],
                  ["manual", "Ingreso manual", Pencil, () => setShowForm(true)],
                  ["imagen", "Imagen", Camera, () => setShowFoto(true)],
                  ["importar", "Importar", Upload, () => setTab("importar")],
                ].map(([key, label, Icon, accion]) => {
                  const active = tab === key;
                  return (
                    <button
                      key={key}
                      onClick={accion}
                      style={{
                        display: "flex", flexDirection: "column", alignItems: "center", gap: 6,
                        width: 128, flexShrink: 0,
                        padding: "12px 6px", borderRadius: 10, cursor: "pointer",
                        border: `1.5px solid ${active ? TEAL : "#ddd6c4"}`,
                        background: active ? TEAL : "#fff",
                        color: active ? "#fff" : INK,
                        boxShadow: active ? "0 4px 14px rgba(15,110,110,0.25)" : "0 1px 4px rgba(27,42,46,0.06)",
                      }}
                    >
                      <Icon size={19} />
                      <span style={{ fontSize: 11, fontWeight: 700, textAlign: "center", lineHeight: 1.1 }}>{label}</span>
                    </button>
                  );
                })}
              </div>
            )}

            {/* Balance ticket */}
            <div style={{
              background: "#fff", marginTop: isDesktop ? 0 : -18, borderRadius: 10, padding: "20px 18px",
              boxShadow: "0 8px 24px rgba(27,42,46,0.12)", position: "relative"
            }}>
              <div style={{ fontSize: 11, letterSpacing: 1, textTransform: "uppercase", color: "#8a9698", marginBottom: 4 }}>Balance del mes</div>
              <div style={{ fontFamily: "'IBM Plex Mono', monospace", fontSize: 32, fontWeight: 600, color: balance >= 0 ? GREEN : BRICK }}>
                {fmtARS(balance)}
              </div>
              <div style={{ display: "flex", gap: 18, marginTop: 14, borderTop: `1px dashed #d8d3c6`, paddingTop: 14, flexWrap: "wrap" }}>
                <Stat icon={<ArrowUpRight size={14} color={GREEN} />} label="Ingresos" value={fmtARS(totalIngresos)} />
                <Stat icon={<ArrowDownRight size={14} color={BRICK} />} label="Gastos" value={fmtARS(totalGastos)} />
                <Stat icon={<PiggyBank size={14} color={GOLD} />} label="Ahorrado" value={fmtARS(totalAhorro)} />
              </div>
            </div>
          </>
        )}

        <div style={{ marginTop: 20 }}>
          {tab === "resumen" && (
            <ResumenTab
              gastosPorCategoria={gastosPorCategoria}
              totalAhorradoHistorico={totalAhorradoHistorico}
              entries={entries}
              thisMonthEntries={thisMonthEntries}
              cambiosStats={cambiosStats}
              totalGastosUsd={totalGastosUsd}
              isDesktop={isDesktop}
              selectedMonth={selectedMonth}
              budgets={budgets}
              onTogglePagado={toggleEntryPagado}
            />
          )}
          {tab === "movimientos" && (
            <MovimientosTab allEntries={entries} entries={thisMonthEntries} categories={categories} onDelete={deleteEntry} onEditDesc={editEntryDesc} onEditAmount={editEntryAmount} onEditFull={editEntryFull} onTogglePagado={toggleEntryPagado} profileName={profileName} monthLabel={monthLabel(selectedMonth)} isDesktop={isDesktop} />
          )}
          {tab === "ahorros" && (
            <AhorrosTab entries={entries.filter((e) => e.type === "ahorro")} onDelete={deleteEntry} totalAhorradoHistorico={totalAhorradoHistorico} />
          )}
          {tab === "presupuestos" && (
            <PresupuestosTab budgets={budgets} onUpdate={updateBudgets} gastosPorCategoria={gastosPorCategoria} categories={categories} isDesktop={isDesktop} />
          )}
          {tab === "importar" && (
            <ImportarTab
              categoryOverrides={categoryOverrides}
              onImport={async (rows, formato, archivos) => {
                const sig = (e) => `${e.date}|${Number(e.amount)}|${(e.desc || "").trim().toLowerCase()}|${(e.account || "").trim().toLowerCase()}`;
                const existentes = new Set(entries.map(sig));
                const vistosEnLote = new Set();
                const nuevos = [];
                const descartados = [];
                rows.forEach((r) => {
                  const s = sig(r);
                  if (existentes.has(s) || vistosEnLote.has(s)) {
                    descartados.push({ date: r.date, desc: r.desc, amount: r.amount, account: r.account, category: r.category, type: r.type });
                    return;
                  }
                  vistosEnLote.add(s);
                  nuevos.push({ ...r, id: uid(), who: r.who || profileName });
                });
                try {
                  if (HAS_SUPABASE && nuevos.length > 0) {
                    // Se manda en lotes de 40 filas (no todo junto en un solo
                    // pedido) — con importaciones grandes (varios PDFs a la
                    // vez) un pedido único puede tardar demasiado o fallar
                    // silenciosamente. Cada lote tiene además un límite de
                    // tiempo, para nunca quedar colgado sin avisar.
                    const TAMANO_LOTE = 40;
                    for (let i = 0; i < nuevos.length; i += TAMANO_LOTE) {
                      const lote = nuevos.slice(i, i + TAMANO_LOTE).map((n) => ({ ...entryToDb(n), household_id: householdId }));
                      await conTimeout(
                        sb("entries", { method: "POST", body: JSON.stringify(lote) }),
                        20000,
                        `La importación tardó demasiado (más de 20s) guardando el lote ${Math.floor(i / TAMANO_LOTE) + 1}. Puede ser un problema de conexión — probá de nuevo en un rato.`
                      );
                    }
                  }
                  setEntries((prev) => {
                    const next = [...nuevos, ...prev];
                    if (!HAS_SUPABASE) mockSaveEntries(next);
                    return next;
                  });
                  // Si lo importado no es del mes que se está viendo (ej.
                  // se importa un resumen de enero mientras la app está
                  // en agosto), saltamos solos al mes correspondiente —
                  // si no, el Resumen sigue mostrando el mes actual vacío
                  // y da la sensación de que la importación no hizo nada,
                  // aunque los datos sí se guardaron bien.
                  if (nuevos.length > 0) {
                    const fechaMasReciente = [...nuevos].sort((a, b) => b.date.localeCompare(a.date))[0].date;
                    const mesImportado = fechaMasReciente.slice(0, 7);
                    if (mesImportado !== selectedMonth) setSelectedMonth(mesImportado);
                  }
                  await logAudit("import", {
                    formato: formato || "?",
                    cantidad_importada: nuevos.length,
                    cantidad_duplicados: descartados.length,
                    cantidad_total: rows.length,
                    duplicados_detalle: descartados,
                    archivos: archivos && archivos.length > 0 ? archivos : undefined,
                  }, null, String(nuevos.length));
                  return { imported: nuevos.length, duplicates: descartados.length };
                } catch (err) {
                  console.error("Error importando movimientos:", err);
                  return { error: err.message || "No se pudo completar la importación." };
                }
              }}
            />
          )}
          {tab === "duplicados" && (
            <DuplicadosTab entries={entries} onDelete={deleteEntry} onForceAdd={addEntry} />
          )}
          {tab === "divisas" && (
            <DivisasTab cambiosStats={cambiosStats} cambios={cambios} onDelete={deleteEntry} />
          )}
          {tab === "historial" && (
            <HistorialTab />
          )}
          {tab === "reset" && (
            <ResetTab onReset={resetearTodo} onResetOverrides={resetearReglasAprendidas} />
          )}
          {tab === "hogar" && (
            <HogarTab
              householdId={householdId}
              onLogout={handleLogout}
              biometriaSoportada={biometriaSoportada}
              biometriaActiva={biometriaActiva}
              onActivarBiometria={activarBiometria}
              onDesactivarBiometria={desactivarBiometria}
              bioBusy={bioBusy}
              bioError={bioError}
              whatsappLinks={whatsappLinks}
              profileName={profileName}
              onAddWhatsapp={addWhatsappLink}
              onDeleteWhatsapp={deleteWhatsappLink}
              onRenombrarMiembro={renombrarMiembro}
              entries={entries}
            />
          )}
          {tab === "categorias" && (
            <CategoriasTab categories={categories} contarUsos={contarUsos} onAdd={addCategory} onRename={renameCategory} onDelete={deleteCategory} />
          )}
          {tab === "recurrentes" && (
            <RecurrentesTab
              recurrentes={recurrentes}
              categories={categories}
              onAdd={addRecurrente}
              onToggleActivo={toggleActivoRecurrente}
              onEditMonto={editMontoRecurrente}
              onDelete={deleteRecurrente}
            />
          )}
          {tab === "rapidos" && (
            <AccesosRapidosTab
              accesosRapidos={accesosRapidos}
              categories={categories}
              currentUserId={session?.user?.id}
              onRegistrar={registrarAccesoRapido}
              onAdd={addAccesoRapido}
              onEdit={editAccesoRapido}
              onDelete={deleteAccesoRapido}
            />
          )}
          {tab === "cuentas" && <CuentasTab thisMonthEntries={thisMonthEntries} monthLabel={monthLabel(selectedMonth)} onRenombrarCuenta={renombrarCuenta} />}
          {tab === "cotizaciones" && <CotizacionesTab />}
          {tab === "personas" && <PersonasTab thisMonthEntries={thisMonthEntries} monthLabel={monthLabel(selectedMonth)} />}
          {tab === "conciliar" && <ConciliarPagosTab entries={entries} onConfirmar={confirmarConciliacion} />}
          {tab === "admin" && <AdminTab />}
          {tab === "recategorizar" && (
            <RecategorizarTab
              categories={categories}
              entries={entries}
              onApply={async (ids, nuevaCategoria) => {
                const idSet = new Set(ids);
                const seleccionadas = entries.filter((e) => idSet.has(e.id));
                const keysAfectadas = new Set(seleccionadas.map((e) => descKey(e.desc)));

                // Además de lo elegido, cualquier otro movimiento ya cargado
                // con la misma descripción (aunque no estuviera visible en
                // este momento) se actualiza también.
                const idsAfectados = new Set(ids);
                entries.forEach((e) => {
                  if (keysAfectadas.has(descKey(e.desc))) idsAfectados.add(e.id);
                });

                const next = entries.map((e) => idsAfectados.has(e.id) ? { ...e, category: nuevaCategoria } : e);
                setEntries(next);

                const nuevosOverrides = { ...categoryOverrides };
                keysAfectadas.forEach((k) => { if (k) nuevosOverrides[k] = nuevaCategoria; });
                setCategoryOverrides(nuevosOverrides);

                if (!HAS_SUPABASE) {
                  mockSaveEntries(next);
                  mockSaveOverrides(nuevosOverrides);
                  return;
                }
                await Promise.all(
                  [...idsAfectados].map((id) =>
                    sb(`entries?id=eq.${encodeURIComponent(id)}`, {
                      method: "PATCH",
                      body: JSON.stringify({ category: nuevaCategoria }),
                    })
                  )
                );
                const rows = [...keysAfectadas].filter(Boolean).map((k) => ({ desc_key: k, category: nuevaCategoria, household_id: householdId }));
                if (rows.length > 0) {
                  await sb("category_overrides", { method: "POST", headers: { Prefer: "resolution=merge-duplicates" }, body: JSON.stringify(rows) });
                }
              }}
            />
          )}
        </div>
      </div>

      {ultimoAccesoRegistrado && (
        <div style={{
          position: "fixed", bottom: 154, left: "50%", transform: "translateX(-50%)",
          background: INK, color: PAPER, borderRadius: 10, padding: "10px 16px",
          display: "flex", alignItems: "center", gap: 12, boxShadow: "0 8px 24px rgba(0,0,0,0.3)",
          zIndex: 15, fontSize: 13, whiteSpace: "nowrap",
        }}>
          <span>Agregado: {ultimoAccesoRegistrado.nombre}</span>
          <button onClick={deshacerUltimoAcceso} style={{ background: "none", border: "none", color: GOLD, fontWeight: 700, cursor: "pointer", fontSize: 13 }}>
            Deshacer
          </button>
        </div>
      )}

      {/* Barra de navegación fija abajo — solo en mobile, estilo apps bancarias */}
      {!isDesktop && (
      <div style={{
        position: "fixed", bottom: 0, left: 0, right: 0, zIndex: 12,
        background: "#fff", borderTop: "1px solid #e6e0d0",
        boxShadow: "0 -4px 16px rgba(27,42,46,0.08)",
        display: "flex", justifyContent: "space-around", alignItems: "stretch",
        paddingBottom: "env(safe-area-inset-bottom, 0px)",
      }}>
        {PRIMARY_TABS.map(([key, label, Icon]) => {
          const active = tab === key;
          return (
            <button
              key={key}
              onClick={() => setTab(key)}
              style={{
                flex: 1, display: "flex", flexDirection: "column", alignItems: "center", justifyContent: "center", gap: 3,
                padding: "10px 4px 8px", background: "none", border: "none", cursor: "pointer",
                borderTop: `2px solid ${active ? TEAL : "transparent"}`,
                color: active ? TEAL : "#8a9698",
              }}
            >
              <Icon size={21} />
              <span style={{ fontSize: 10.5, fontWeight: active ? 700 : 500 }}>{label}</span>
            </button>
          );
        })}
      </div>
      )}

      {/* Menú de opciones del botón central + el botón en sí: solo mobile.
          En desktop, "Ingreso manual" e "Imagen" ya están en la barra de
          arriba, y "Por voz" no aplica tanto en escritorio. */}
      {!isDesktop && (
        <>
      {showAddMenu && (
        <div onClick={() => setShowAddMenu(false)} style={{ position: "fixed", inset: 0, zIndex: 13 }} />
      )}
      {showAddMenu && (
        <div style={{
          position: "fixed", bottom: 156, left: "50%", transform: "translateX(-50%)",
          zIndex: 14, display: "flex", flexDirection: "column", gap: 10, alignItems: "center",
        }}>
          {[
            ["manual", "Ingreso manual", Pencil, () => setShowForm(true)],
            ["voz", "Por voz", Mic, () => setShowVoice(true)],
            ["foto", "Foto de recibo", Camera, () => setShowFoto(true)],
          ].map(([key, label, Icon, accion]) => (
            <button
              key={key}
              onClick={() => { setShowAddMenu(false); accion(); }}
              style={{
                display: "flex", alignItems: "center", gap: 10, padding: "10px 16px 10px 14px",
                width: 210, background: "#fff", border: "none", borderRadius: 30, cursor: "pointer",
                boxShadow: "0 4px 14px rgba(27,42,46,0.2)", whiteSpace: "nowrap",
              }}
            >
              <span style={{
                width: 34, height: 34, borderRadius: "50%", background: PAPER_DIM, color: TEAL,
                display: "flex", alignItems: "center", justifyContent: "center", flexShrink: 0,
              }}>
                <Icon size={17} />
              </span>
              <span style={{ fontSize: 13.5, fontWeight: 600, color: INK }}>{label}</span>
            </button>
          ))}
        </div>
      )}

      {/* Botón central flotante, sobre la barra de navegación */}
      <button
        onClick={() => setShowAddMenu((v) => !v)}
        style={{
          position: "fixed", bottom: 46, left: "50%", transform: `translateX(-50%) rotate(${showAddMenu ? 45 : 0}deg)`,
          width: 62, height: 62, borderRadius: "50%",
          background: TEAL, color: "#fff", border: "4px solid #fff", boxShadow: "0 8px 20px rgba(15,110,110,0.4)",
          display: "flex", alignItems: "center", justifyContent: "center", cursor: "pointer", zIndex: 15,
          transition: "transform 0.2s ease",
        }}
        aria-label={showAddMenu ? "Cerrar opciones de carga" : "Agregar movimiento"}
      >
        <Plus size={28} />
      </button>
        </>
      )}

      {showFoto && (
        <FotoReciboModal
          categories={categories}
          onClose={() => setShowFoto(false)}
          onExtracted={(datos) => {
            setVoicePrefill(datos);
            setShowFoto(false);
            setShowForm(true);
          }}
        />
      )}

      {showVoice && (
        <VoiceEntryModal
          categories={categories}
          onClose={() => setShowVoice(false)}
          onExtracted={(datos) => {
            setVoicePrefill(datos);
            setShowVoice(false);
            setShowForm(true);
          }}
        />
      )}

      {showForm && (
        <EntryForm
          onClose={() => { setShowForm(false); setVoicePrefill(null); setSaveError(null); }}
          onSave={async (entry) => {
            setSaving(true);
            setSaveError(null);
            try {
              await addEntry(entry);
              setShowForm(false);
              setVoicePrefill(null);
            } catch (err) {
              console.error("Error guardando el movimiento:", err);
              setSaveError(err.message || "No se pudo guardar. Probá de nuevo.");
            } finally {
              setSaving(false);
            }
          }}
          saving={saving}
          saveError={saveError}
          categories={categories}
          initialData={voicePrefill}
          profileName={profileName}
          miembrosHogar={miembrosHogar}
        />
      )}
    </div>
  );
}

function Stat({ icon, label, value }) {
  return (
    <div style={{ display: "flex", flexDirection: "column", gap: 2 }}>
      <div style={{ display: "flex", alignItems: "center", gap: 5, fontSize: 11.5, color: "#8a9698" }}>{icon}{label}</div>
      <div style={{ fontFamily: "'IBM Plex Mono', monospace", fontSize: 15, fontWeight: 600 }}>{value}</div>
    </div>
  );
}

function EmptyState({ text }) {
  return (
    <div style={{ textAlign: "center", padding: "36px 16px", color: "#9a9488", fontSize: 14 }}>{text}</div>
  );
}

function DivisasTab({ cambiosStats, cambios, onDelete }) {
  if (!cambiosStats) {
    return (
      <div style={{ background: "#fff", borderRadius: 10, padding: 18, boxShadow: "0 2px 10px rgba(27,42,46,0.06)" }}>
        <div style={{ fontFamily: "'Fraunces', serif", fontSize: 17, marginBottom: 6 }}>Divisas</div>
        <EmptyState text="Todavía no cargaste ningún cambio USD→ARS." />
      </div>
    );
  }
  return (
    <div style={{ display: "flex", flexDirection: "column", gap: 14 }}>
      <div style={{ background: "#fff", borderRadius: 10, padding: 18, boxShadow: "0 2px 10px rgba(27,42,46,0.06)" }}>
        <div style={{ fontFamily: "'Fraunces', serif", fontSize: 17, marginBottom: 14 }}>Divisas</div>

        <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(150px, 1fr))", gap: 10, marginBottom: 14 }}>
          <div style={{ background: PAPER_DIM, borderRadius: 8, padding: "12px 14px" }}>
            <div style={{ fontSize: 10.5, color: "#8a9698", textTransform: "uppercase", letterSpacing: 0.5, marginBottom: 6 }}>Último cambio</div>
            <div style={{ fontFamily: "'IBM Plex Mono', monospace", fontSize: 19, fontWeight: 700, color: TEAL }}>${cambiosStats.ultimo.rate}</div>
            <div style={{ fontSize: 10.5, color: "#9a9488", marginTop: 2 }}>{cambiosStats.ultimo.date}</div>
          </div>
          <div style={{ background: PAPER_DIM, borderRadius: 8, padding: "12px 14px" }}>
            <div style={{ fontSize: 10.5, color: "#8a9698", textTransform: "uppercase", letterSpacing: 0.5, marginBottom: 6 }}>Promedio histórico</div>
            <div style={{ fontFamily: "'IBM Plex Mono', monospace", fontSize: 19, fontWeight: 700 }}>${cambiosStats.promedio.toFixed(0)}</div>
          </div>
          <div style={{ background: PAPER_DIM, borderRadius: 8, padding: "12px 14px" }}>
            <div style={{ fontSize: 10.5, color: "#8a9698", textTransform: "uppercase", letterSpacing: 0.5, marginBottom: 6 }}>USD cambiados en total</div>
            <div style={{ fontFamily: "'IBM Plex Mono', monospace", fontSize: 19, fontWeight: 700 }}>USD {cambiosStats.totalUsd.toLocaleString("es-AR")}</div>
          </div>
        </div>

        <div style={{ fontSize: 11, color: "#8a9698", textTransform: "uppercase", letterSpacing: 0.5, marginBottom: 8, paddingTop: 4, borderTop: "1px solid #f0ece0" }}>
          Por origen
        </div>
        <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(180px, 1fr))", gap: 10 }}>
          <div style={{ border: `1.5px solid ${GREEN}`, borderRadius: 8, padding: "12px 14px" }}>
            <div style={{ fontSize: 11.5, fontWeight: 700, color: GREEN, marginBottom: 6 }}>ARQ (sueldo)</div>
            <div style={{ fontFamily: "'IBM Plex Mono', monospace", fontSize: 16, fontWeight: 700 }}>USD {cambiosStats.arq.totalUsd.toLocaleString("es-AR")}</div>
            <div style={{ fontFamily: "'IBM Plex Mono', monospace", fontSize: 13, color: "#5a6b6d", marginTop: 2 }}>{fmtARS(cambiosStats.arq.totalArs)}</div>
          </div>
          <div style={{ border: `1.5px solid ${GOLD}`, borderRadius: 8, padding: "12px 14px" }}>
            <div style={{ fontSize: 11.5, fontWeight: 700, color: GOLD, marginBottom: 6 }}>Externos (cueva/change)</div>
            <div style={{ fontFamily: "'IBM Plex Mono', monospace", fontSize: 16, fontWeight: 700 }}>USD {cambiosStats.externos.totalUsd.toLocaleString("es-AR")}</div>
            <div style={{ fontFamily: "'IBM Plex Mono', monospace", fontSize: 13, color: "#5a6b6d", marginTop: 2 }}>{fmtARS(cambiosStats.externos.totalArs)}</div>
          </div>
        </div>
      </div>

      <div style={{ background: "#fff", borderRadius: 10, padding: 18, boxShadow: "0 2px 10px rgba(27,42,46,0.06)" }}>
        <div style={{ fontFamily: "'Fraunces', serif", fontSize: 16, marginBottom: 12 }}>Historial de cambios ({cambios.length})</div>
        <div style={{ display: "flex", flexDirection: "column", gap: 8 }}>
          {cambios.map((e) => {
            const esArq = (e.account || "").trim().toUpperCase() === "ARQ";
            return (
              <div key={e.id} style={{ background: PAPER_DIM, borderRadius: 8, padding: "10px 14px", display: "flex", justifyContent: "space-between", alignItems: "center", gap: 10 }}>
                <div style={{ minWidth: 0 }}>
                  <div style={{ fontSize: 13.5, fontWeight: 600 }}>
                    USD {Number(e.usdAmount).toLocaleString("es-AR")} a ${e.rate}
                    <span style={{ marginLeft: 8, fontSize: 10.5, fontWeight: 700, padding: "2px 7px", borderRadius: 10, background: esArq ? "#e4f0e8" : "#fbf1de", color: esArq ? GREEN : GOLD }}>
                      {esArq ? "ARQ (sueldo)" : (e.account || "externo")}
                    </span>
                  </div>
                  <div style={{ fontSize: 11.5, color: "#9a9488", marginTop: 2, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>
                    {e.date} · {e.who}{e.desc ? ` · ${e.desc}` : ""}
                  </div>
                </div>
                <div style={{ display: "flex", alignItems: "center", gap: 10, flexShrink: 0 }}>
                  <div style={{ fontFamily: "'IBM Plex Mono', monospace", fontWeight: 600, color: TEAL }}>{fmtARS(e.amount)}</div>
                  {onDelete && (
                    <button onClick={() => onDelete(e.id)} style={{ border: "none", background: "none", cursor: "pointer", color: "#b8b2a4" }} aria-label="Borrar">
                      <Trash2 size={15} />
                    </button>
                  )}
                </div>
              </div>
            );
          })}
        </div>
      </div>
    </div>
  );
}

function ResumenTab({ gastosPorCategoria, totalAhorradoHistorico, entries, thisMonthEntries, cambiosStats, totalGastosUsd, isDesktop, selectedMonth, budgets, onTogglePagado }) {
  const [expandedCat, setExpandedCat] = useState(null);
  const [rango, setRango] = useState(6);
  const [compareMode, setCompareMode] = useState(false);
  const [comparedMonths, setComparedMonths] = useState([]);
  // Ventana de 3 meses para "Resumen mensual" en el celu, navegable con
  // flechas — independiente del selector de rango de Evolución de arriba.
  const [resumenOffset, setResumenOffset] = useState(0);
  const [marcandoPagado, setMarcandoPagado] = useState(null);

  // Pendientes de pago de TODO el historial (no solo el mes que estás
  // mirando) — para que nada quede olvidado aunque cambies de mes.
  const pendientesDePago = useMemo(() => {
    return entries
      .filter((e) => e.type === "gasto" && e.pagado === false)
      .sort((a, b) => (a.date || "").localeCompare(b.date || "")); // más viejo primero
  }, [entries]);
  const totalPendiente = pendientesDePago.reduce((s, e) => s + Number(e.amount), 0);

  async function handleMarcarPagado(id) {
    setMarcandoPagado(id);
    await onTogglePagado(id, true);
    setMarcandoPagado(null);
  }

  // Umbral para las alertas de "gasto que subió mucho" — se guarda en este
  // dispositivo para no tener que reconfigurarlo cada vez.
  const [umbralAlerta, setUmbralAlerta] = useState(() => Number(safeGet("umbral_alerta_gastos")) || 20);
  useEffect(() => { safeSet("umbral_alerta_gastos", umbralAlerta); }, [umbralAlerta]);
  // Plegada por defecto — ocupaba mucho lugar arriba de todo lo demás.
  const [alertasAbiertas, setAlertasAbiertas] = useState(false);

  const alertasGasto = useMemo(() => {
    const [y, m] = selectedMonth.split("-").map(Number);
    const mesAnteriorKey = new Date(y, m - 2, 1).toISOString().slice(0, 7);

    const porCategoriaActual = {};
    const porCategoriaAnterior = {};
    entries.forEach((e) => {
      if (e.type !== "gasto" || (e.moneda || "ARS") !== "ARS") return;
      const mk = monthKey(e.date);
      if (mk === selectedMonth) porCategoriaActual[e.category] = (porCategoriaActual[e.category] || 0) + Number(e.amount);
      else if (mk === mesAnteriorKey) porCategoriaAnterior[e.category] = (porCategoriaAnterior[e.category] || 0) + Number(e.amount);
    });

    const resultado = [];
    Object.keys(porCategoriaActual).forEach((cat) => {
      const actual = porCategoriaActual[cat];
      const anterior = porCategoriaAnterior[cat] || 0;
      if (anterior > 0) {
        const variacion = ((actual - anterior) / anterior) * 100;
        if (variacion >= umbralAlerta) resultado.push({ categoria: cat, actual, anterior, variacion, nueva: false });
      } else if (actual > 0) {
        resultado.push({ categoria: cat, actual, anterior: 0, variacion: null, nueva: true });
      }
    });
    return resultado.sort((a, b) => (b.variacion ?? 999999) - (a.variacion ?? 999999));
  }, [entries, selectedMonth, umbralAlerta]);

  const alertasPresupuesto = useMemo(() => {
    const gastoMap = Object.fromEntries(gastosPorCategoria.map((g) => [g.name, g.value]));
    const resultado = [];
    Object.entries(budgets || {}).forEach(([cat, limite]) => {
      const limiteNum = Number(limite) || 0;
      if (limiteNum <= 0) return;
      const gastado = gastoMap[cat] || 0;
      const pct = (gastado / limiteNum) * 100;
      if (pct >= 80) resultado.push({ categoria: cat, gastado, limite: limiteNum, pct, superado: gastado > limiteNum });
    });
    return resultado.sort((a, b) => b.pct - a.pct);
  }, [budgets, gastosPorCategoria]);

  function calcularMes(key) {
    const [y, m] = key.split("-").map(Number);
    const d = new Date(y, m - 1, 1);
    const label = d.toLocaleDateString("es-AR", { month: "short", year: "2-digit" });
    const ing = entries.filter((e) => (e.type === "ingreso" || e.type === "cambio") && monthKey(e.date) === key).reduce((s, e) => s + Number(e.amount), 0);
    const gas = entries.filter((e) => e.type === "gasto" && monthKey(e.date) === key).reduce((s, e) => s + Number(e.amount), 0);
    // Solo tarjetas de crédito (BBVA Visa, Mastercard Black, etc.) — no
    // incluye Efectivo, Mercado Pago ni ARQ, para responder puntualmente
    // "cuánto gasté de tarjeta este mes".
    const tarj = entries
      .filter((e) => e.type === "gasto" && monthKey(e.date) === key && clasificarMedioPago(e.account) === "Tarjetas de crédito")
      .reduce((s, e) => s + Number(e.amount), 0);
    return { mes: label, Ingresos: ing, Gastos: gas, Tarjetas: tarj };
  }

  // Siempre el último año hasta el mes actual (nunca meses futuros, aunque
  // ya haya movimientos cargados a futuro como la casita de Hebraica).
  const availableMonths = useMemo(() => {
    const now = new Date();
    const months = [];
    for (let i = 0; i <= 12; i++) {
      const d = new Date(now.getFullYear(), now.getMonth() - i, 1);
      months.push(d.toISOString().slice(0, 7));
    }
    return months;
  }, []);

  const resumenMensualMobile = useMemo(() => {
    const now = new Date();
    const months = [];
    for (let i = 2; i >= 0; i--) {
      const d = new Date(now.getFullYear(), now.getMonth() - resumenOffset - i, 1);
      months.push(calcularMes(d.toISOString().slice(0, 7)));
    }
    return months;
  }, [entries, resumenOffset]);

  const chartData = useMemo(() => {
    if (compareMode) {
      return [...comparedMonths].sort().map(calcularMes);
    }
    const now = new Date();
    let mesesAMostrar = rango;
    if (rango === "todo") {
      const fechas = entries.map((e) => e.date).filter(Boolean).sort();
      if (fechas.length === 0) return [];
      const primera = new Date(fechas[0] + "-01");
      mesesAMostrar = (now.getFullYear() - primera.getFullYear()) * 12 + (now.getMonth() - primera.getMonth()) + 1;
      mesesAMostrar = Math.min(mesesAMostrar, 60);
    }
    const months = [];
    for (let i = mesesAMostrar - 1; i >= 0; i--) {
      const d = new Date(now.getFullYear(), now.getMonth() - i, 1);
      const key = d.toISOString().slice(0, 7);
      const label = d.toLocaleDateString("es-AR", { month: "short", year: mesesAMostrar > 12 ? "2-digit" : undefined });
      const ing = entries.filter((e) => (e.type === "ingreso" || e.type === "cambio") && monthKey(e.date) === key).reduce((s, e) => s + Number(e.amount), 0);
      const gas = entries.filter((e) => e.type === "gasto" && monthKey(e.date) === key).reduce((s, e) => s + Number(e.amount), 0);
      const tarj = entries
        .filter((e) => e.type === "gasto" && monthKey(e.date) === key && clasificarMedioPago(e.account) === "Tarjetas de crédito")
        .reduce((s, e) => s + Number(e.amount), 0);
      months.push({ mes: label, Ingresos: ing, Gastos: gas, Tarjetas: tarj });
    }
    return months;
  }, [entries, rango, compareMode, comparedMonths]);

  function toggleMesComparado(key) {
    setComparedMonths((prev) => prev.includes(key) ? prev.filter((k) => k !== key) : [...prev, key]);
  }

  // Cuántas etiquetas de mes entran cómodas en el eje X del gráfico de
  // "Evolución" antes de amontonarse — bastante menos en el celu (pantalla
  // angosta) que en escritorio. Con más meses que ese máximo, salteamos
  // etiquetas parejo para no superponer texto.
  const xAxisInterval = useMemo(() => {
    const maxEtiquetas = isDesktop ? 12 : 6;
    return chartData.length > maxEtiquetas ? Math.ceil(chartData.length / maxEtiquetas) - 1 : 0;
  }, [chartData.length, isDesktop]);

  // Lista de tarjetas/titulares distintos que aparecen en el historial
  // (ej. "Visa BBVA Hernán", "Visa BBVA Hernán · Natalia", "Mastercard
  // Black"), ordenadas de mayor a menor gasto histórico — define las
  // columnas de la tabla de desglose de abajo.
  const cuentasTarjetaLabels = useMemo(() => {
    const map = {};
    entries.forEach((e) => {
      if (e.type !== "gasto" || (e.moneda || "ARS") !== "ARS" || clasificarMedioPago(e.account) !== "Tarjetas de crédito") return;
      const label = etiquetaTarjeta(e);
      map[label] = (map[label] || 0) + Number(e.amount);
    });
    return Object.entries(map).sort((a, b) => b[1] - a[1]).map(([label]) => label);
  }, [entries]);

  // Rango propio para "Gasto por tarjeta, mes a mes" — independiente del
  // selector de "Evolución" de arriba. Antes compartían el mismo rango y
  // no quedaba claro cómo ver meses más viejos en la tabla de tarjetas
  // sin ir a cambiar el selector de otra sección.
  const [rangoTarjetas, setRangoTarjetas] = useState(6);
  const chartDataTarjetasPorCuenta = useMemo(() => {
    function calcularMesPorCuenta(key) {
      const [y, m] = key.split("-").map(Number);
      const d = new Date(y, m - 1, 1);
      const label = d.toLocaleDateString("es-AR", { month: "short", year: "2-digit" });
      const fila = { mes: label };
      cuentasTarjetaLabels.forEach((c) => { fila[c] = 0; });
      entries.forEach((e) => {
        if (e.type !== "gasto" || (e.moneda || "ARS") !== "ARS" || monthKey(e.date) !== key || clasificarMedioPago(e.account) !== "Tarjetas de crédito") return;
        const label2 = etiquetaTarjeta(e);
        fila[label2] = (fila[label2] || 0) + Number(e.amount);
      });
      return fila;
    }
    const now = new Date();
    let mesesAMostrar = rangoTarjetas;
    if (rangoTarjetas === "todo") {
      const fechas = entries.map((e) => e.date).filter(Boolean).sort();
      if (fechas.length === 0) return [];
      const primera = new Date(fechas[0] + "-01");
      mesesAMostrar = (now.getFullYear() - primera.getFullYear()) * 12 + (now.getMonth() - primera.getMonth()) + 1;
      mesesAMostrar = Math.min(mesesAMostrar, 60);
    }
    const months = [];
    for (let i = mesesAMostrar - 1; i >= 0; i--) {
      const d = new Date(now.getFullYear(), now.getMonth() - i, 1);
      const key = d.toISOString().slice(0, 7);
      months.push(calcularMesPorCuenta(key));
    }
    return months;
  }, [entries, rangoTarjetas, cuentasTarjetaLabels]);

  return (
    <div style={{ display: "flex", flexDirection: "column", gap: 24 }}>
      {pendientesDePago.length > 0 && (
        <div style={{ background: "#fff", borderRadius: 10, padding: 18, boxShadow: "0 2px 10px rgba(27,42,46,0.06)" }}>
          <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", flexWrap: "wrap", gap: 8, marginBottom: 12 }}>
            <div style={{ fontFamily: "'Fraunces', serif", fontSize: 17 }}>💳 Pendientes de pago</div>
            <div style={{ fontFamily: "'IBM Plex Mono', monospace", fontWeight: 700, color: GOLD }}>{fmtARS(totalPendiente)} en total</div>
          </div>
          <div style={{ display: "flex", flexDirection: "column", gap: 8 }}>
            {pendientesDePago.map((e) => (
              <div key={e.id} style={{ display: "flex", justifyContent: "space-between", alignItems: "center", padding: "8px 10px", background: "#fdf1de", borderRadius: 8, gap: 10 }}>
                <div style={{ minWidth: 0 }}>
                  <div style={{ fontSize: 13, fontWeight: 700, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>{e.desc || e.category}</div>
                  <div style={{ fontSize: 11, color: "#8a9698" }}>{e.date} · {e.category}{e.account ? ` · ${e.account}` : ""}</div>
                </div>
                <div style={{ display: "flex", alignItems: "center", gap: 10, flexShrink: 0 }}>
                  <div style={{ fontFamily: "'IBM Plex Mono', monospace", fontWeight: 700 }}>{fmtARS(e.amount)}</div>
                  <button
                    onClick={() => handleMarcarPagado(e.id)}
                    disabled={marcandoPagado === e.id}
                    style={{ ...btnPrimary, background: GREEN, padding: "6px 10px", fontSize: 11.5 }}
                  >
                    {marcandoPagado === e.id ? "..." : "Marcar pagado"}
                  </button>
                </div>
              </div>
            ))}
          </div>
        </div>
      )}

      {alertasPresupuesto.length > 0 && (
        <div style={{ background: "#fff", borderRadius: 10, padding: 18, boxShadow: "0 2px 10px rgba(27,42,46,0.06)" }}>
          <div style={{ fontFamily: "'Fraunces', serif", fontSize: 17, marginBottom: 12 }}>🎯 Presupuestos cerca del límite</div>
          <div style={{ display: "flex", flexDirection: "column", gap: 10 }}>
            {alertasPresupuesto.map((a) => (
              <div key={a.categoria}>
                <div style={{ display: "flex", justifyContent: "space-between", fontSize: 13, marginBottom: 4 }}>
                  <span style={{ fontWeight: 700 }}>{a.categoria}</span>
                  <span style={{ fontFamily: "'IBM Plex Mono', monospace", fontWeight: 700, color: a.superado ? BRICK : GOLD }}>
                    {fmtARS(a.gastado)} / {fmtARS(a.limite)}
                  </span>
                </div>
                <div style={{ height: 6, borderRadius: 4, background: PAPER_DIM, overflow: "hidden" }}>
                  <div style={{ width: `${Math.min(100, a.pct)}%`, height: "100%", background: a.superado ? BRICK : GOLD }} />
                </div>
                <div style={{ fontSize: 11, color: a.superado ? BRICK : GOLD, marginTop: 3, fontWeight: 600 }}>
                  {a.superado ? `¡Superado! (${Math.round(a.pct)}%)` : `${Math.round(a.pct)}% del presupuesto`}
                </div>
              </div>
            ))}
          </div>
        </div>
      )}

      <div style={{ background: "#fff", borderRadius: 10, padding: 18, boxShadow: "0 2px 10px rgba(27,42,46,0.06)" }}>
        <div
          onClick={() => setAlertasAbiertas((v) => !v)}
          style={{ display: "flex", justifyContent: "space-between", alignItems: "center", flexWrap: "wrap", gap: 8, cursor: "pointer" }}
        >
          <div style={{ fontFamily: "'Fraunces', serif", fontSize: 17 }}>⚠️ Aumentos de gasto</div>
          <div style={{ display: "flex", alignItems: "center", gap: 10 }}>
            {!alertasAbiertas && (
              <span style={{ fontSize: 12, color: alertasGasto.length > 0 ? BRICK : "#8a9698", fontWeight: alertasGasto.length > 0 ? 700 : 400 }}>
                {alertasGasto.length === 0 ? "Nada relevante" : `${alertasGasto.length} categoría${alertasGasto.length > 1 ? "s" : ""} subió más de ${umbralAlerta}%`}
              </span>
            )}
            <span style={{ fontSize: 12, color: TEAL, fontWeight: 700 }}>{alertasAbiertas ? "Ocultar ▲" : "Ver detalle ▼"}</span>
          </div>
        </div>
        {alertasAbiertas && (
          <>
            <div style={{ display: "flex", justifyContent: "flex-end", alignItems: "center", gap: 6, fontSize: 12, color: "#8a9698", marginTop: 10 }}>
              Avisame si sube más de
              <input
                type="number"
                value={umbralAlerta}
                onChange={(e) => setUmbralAlerta(Math.max(1, Number(e.target.value) || 1))}
                style={{ width: 52, padding: "3px 6px", borderRadius: 6, border: "1px solid #ddd6c4", fontSize: 12.5, textAlign: "center" }}
              />
              %
            </div>
            <div style={{ fontSize: 12, color: "#8a9698", margin: "10px 0 12px" }}>
              Comparando el mes que estás mirando con el anterior.
            </div>
            {alertasGasto.length === 0 ? (
              <EmptyState text={`Ninguna categoría subió más de ${umbralAlerta}% respecto al mes anterior.`} />
            ) : (
              <div style={{ display: "flex", flexDirection: "column", gap: 8 }}>
                {alertasGasto.map((a) => (
                  <div key={a.categoria} style={{ display: "flex", justifyContent: "space-between", alignItems: "center", padding: "8px 10px", background: "#fbeee6", borderRadius: 8 }}>
                    <div>
                      <div style={{ fontSize: 13.5, fontWeight: 700 }}>{a.categoria}</div>
                      <div style={{ fontSize: 11.5, color: "#8a9698" }}>
                        {a.nueva ? "Sin gasto el mes anterior" : `${fmtARS(a.anterior)} → ${fmtARS(a.actual)}`}
                      </div>
                    </div>
                    <div style={{ textAlign: "right" }}>
                      <div style={{ fontFamily: "'IBM Plex Mono', monospace", fontWeight: 700, color: BRICK, fontSize: 14 }}>{fmtARS(a.actual)}</div>
                      <div style={{ fontSize: 11.5, fontWeight: 700, color: BRICK }}>{a.nueva ? "Nuevo" : `+${Math.round(a.variacion)}%`}</div>
                    </div>
                  </div>
                ))}
              </div>
            )}
          </>
        )}
      </div>

      <div style={isDesktop ? { display: "grid", gridTemplateColumns: "1fr 1fr", gap: 24, alignItems: "start" } : { display: "contents" }}>
      <div style={{ background: "#fff", borderRadius: 10, padding: 18, boxShadow: "0 2px 10px rgba(27,42,46,0.06)" }}>
        <div style={{ fontFamily: "'Fraunces', serif", fontSize: 17, marginBottom: 6 }}>Gastos por categoría (este mes)</div>
        {totalGastosUsd > 0 && (
          <div style={{ fontSize: 12.5, color: "#8a9698", marginBottom: 12 }}>
            + <b style={{ color: BRICK, fontFamily: "'IBM Plex Mono', monospace" }}>USD {totalGastosUsd.toLocaleString("es-AR", { maximumFractionDigits: 2 })}</b> en gastos cargados directo en dólares (no incluidos en el total de pesos de arriba).
          </div>
        )}
        {gastosPorCategoria.length === 0 ? <EmptyState text="Todavía no cargaste gastos este mes." /> : (
          <div style={{ display: "flex", alignItems: "center", gap: 12, flexWrap: "wrap" }}>
            <div style={{ width: 180, height: 180 }}>
              <ResponsiveContainer>
                <PieChart>
                  <Pie data={gastosPorCategoria} dataKey="value" nameKey="name" innerRadius={45} outerRadius={80} paddingAngle={2}>
                    {gastosPorCategoria.map((_, i) => <Cell key={i} fill={CAT_COLORS[i % CAT_COLORS.length]} />)}
                  </Pie>
                  <Tooltip formatter={(v) => fmtARS(v)} />
                </PieChart>
              </ResponsiveContainer>
            </div>
            <div style={{ flex: 1, minWidth: 160 }}>
              {gastosPorCategoria.map((c, i) => {
                const isOpen = expandedCat === c.name;
                const movs = isOpen
                  ? thisMonthEntries
                      .filter((e) => e.type === "gasto" && (e.moneda || "ARS") === "ARS" && e.category === c.name)
                      .sort((a, b) => (b.date || "").localeCompare(a.date || ""))
                  : [];
                return (
                  <div key={c.name}>
                    <div
                      className="cat-row"
                      onClick={() => setExpandedCat(isOpen ? null : c.name)}
                      style={{ display: "flex", justifyContent: "space-between", padding: "5px 6px", borderRadius: 6, fontSize: 13, cursor: "pointer", background: isOpen ? "#f2eee2" : "transparent" }}
                    >
                      <span style={{ display: "flex", alignItems: "center", gap: 6 }}>
                        <span style={{ width: 8, height: 8, borderRadius: "50%", background: CAT_COLORS[i % CAT_COLORS.length] }} />
                        {c.name}
                      </span>
                      <span style={{ display: "flex", alignItems: "center", gap: 6 }}>
                        <span style={{ fontFamily: "'IBM Plex Mono', monospace" }}>{fmtARS(c.value)}</span>
                        <span style={{ fontSize: 10, color: "#8a9698" }}>{isOpen ? "▲" : "▼"}</span>
                      </span>
                    </div>
                    {isOpen && (
                      <div style={{ padding: "4px 6px 10px 22px", display: "flex", flexDirection: "column", gap: 4 }}>
                        {movs.length === 0 ? (
                          <div style={{ fontSize: 12, color: "#8a9698" }}>Sin movimientos.</div>
                        ) : movs.map((e) => (
                          <div key={e.id} style={{ display: "flex", justifyContent: "space-between", fontSize: 12, color: "#5a6b6d" }}>
                            <span>{e.date} · {e.desc || "(sin descripción)"}{e.account ? ` · ${e.account}` : ""}</span>
                            <span style={{ fontFamily: "'IBM Plex Mono', monospace", flexShrink: 0, marginLeft: 8 }}>{fmtARS(e.amount)}</span>
                          </div>
                        ))}
                      </div>
                    )}
                  </div>
                );
              })}
            </div>
          </div>
        )}
      </div>

      <div style={{ background: "#fff", borderRadius: 10, padding: 18, boxShadow: "0 2px 10px rgba(27,42,46,0.06)" }}>
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", flexWrap: "wrap", gap: 8, marginBottom: 12 }}>
          <div style={{ fontFamily: "'Fraunces', serif", fontSize: 17 }}>Evolución</div>
          <div style={{ display: "flex", gap: 4, flexWrap: "wrap" }}>
            {!compareMode && [[3, "3M"], [6, "6M"], [12, "12M"], [24, "24M"], ["todo", "Todo"]].map(([v, l]) => (
              <button key={l} onClick={() => setRango(v)} style={{
                padding: "5px 10px", borderRadius: 16, fontSize: 11.5, fontWeight: 700, cursor: "pointer",
                border: `1px solid ${rango === v ? TEAL : "#ddd6c4"}`,
                background: rango === v ? TEAL : "#fff", color: rango === v ? "#fff" : INK
              }}>{l}</button>
            ))}
            <button
              onClick={() => setCompareMode((v) => !v)}
              style={{
                padding: "5px 10px", borderRadius: 16, fontSize: 11.5, fontWeight: 700, cursor: "pointer",
                border: `1px solid ${compareMode ? GOLD : "#ddd6c4"}`,
                background: compareMode ? GOLD : "#fff", color: compareMode ? "#fff" : INK,
              }}
            >
              {compareMode ? "✕ Comparar meses" : "Comparar meses"}
            </button>
          </div>
        </div>

        {compareMode && (
          <div style={{ marginBottom: 14 }}>
            <div style={{ fontSize: 12, color: "#8a9698", marginBottom: 6 }}>Elegí los meses que querés comparar (no hace falta que sean seguidos):</div>
            {availableMonths.length === 0 ? (
              <div style={{ fontSize: 12.5, color: "#8a9698" }}>Todavía no hay movimientos cargados.</div>
            ) : (
              <div style={{ display: "flex", flexWrap: "wrap", gap: 6 }}>
                {availableMonths.map((key) => {
                  const [y, m] = key.split("-").map(Number);
                  const label = new Date(y, m - 1, 1).toLocaleDateString("es-AR", { month: "short", year: "2-digit" });
                  const active = comparedMonths.includes(key);
                  return (
                    <button
                      key={key}
                      onClick={() => toggleMesComparado(key)}
                      style={{
                        padding: "5px 10px", borderRadius: 16, fontSize: 11.5, fontWeight: 700, cursor: "pointer", textTransform: "capitalize",
                        border: `1px solid ${active ? TEAL : "#ddd6c4"}`,
                        background: active ? TEAL : "#fff", color: active ? "#fff" : INK,
                      }}
                    >
                      {label}
                    </button>
                  );
                })}
              </div>
            )}
          </div>
        )}

        {compareMode && comparedMonths.length === 0 ? (
          <EmptyState text="Elegí al menos un mes para comparar." />
        ) : (
          <div style={{ width: "100%", height: 220 }}>
            <ResponsiveContainer>
              <BarChart data={chartData}>
                <CartesianGrid strokeDasharray="3 3" stroke="#eee6d5" vertical={false} />
                <XAxis dataKey="mes" tick={{ fontSize: 11, fill: "#8a9698" }} axisLine={false} tickLine={false} interval={xAxisInterval} />
                <YAxis tick={{ fontSize: 11, fill: "#8a9698" }} axisLine={false} tickLine={false} width={40} tickFormatter={(v) => `${Math.round(v / 1000)}k`} />
                <Tooltip formatter={(v) => fmtARS(v)} />
                <Legend wrapperStyle={{ fontSize: 12 }} />
                <Bar dataKey="Ingresos" fill={GREEN} radius={[3, 3, 0, 0]} />
                <Bar dataKey="Gastos" fill={BRICK} radius={[3, 3, 0, 0]} />
              </BarChart>
            </ResponsiveContainer>
          </div>
        )}
      </div>
      </div>

      <div style={{ background: "#fff", borderRadius: 10, padding: 18, boxShadow: "0 2px 10px rgba(27,42,46,0.06)" }}>
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", flexWrap: "wrap", gap: 8, marginBottom: 12 }}>
          <div style={{ fontFamily: "'Fraunces', serif", fontSize: 17 }}>Resumen mensual</div>
          {!isDesktop && (
            <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
              <button onClick={() => setResumenOffset((v) => v + 3)} aria-label="Meses anteriores" style={{ background: "none", border: "1px solid #ddd6c4", borderRadius: 6, width: 26, height: 26, cursor: "pointer", color: INK }}>◀</button>
              {resumenOffset > 0 && (
                <button onClick={() => setResumenOffset(0)} style={{ background: "none", border: `1px solid ${TEAL}`, borderRadius: 6, color: TEAL, fontSize: 11, padding: "3px 8px", cursor: "pointer" }}>Mes actual</button>
              )}
              <button
                onClick={() => setResumenOffset((v) => Math.max(0, v - 3))}
                disabled={resumenOffset === 0}
                aria-label="Meses siguientes"
                style={{ background: "none", border: "1px solid #ddd6c4", borderRadius: 6, width: 26, height: 26, cursor: resumenOffset === 0 ? "default" : "pointer", color: resumenOffset === 0 ? "#ddd6c4" : INK }}
              >▶</button>
            </div>
          )}
        </div>
        {isDesktop ? (
          chartData.length === 0 ? <EmptyState text={compareMode ? "Elegí al menos un mes para comparar." : "Sin datos para este rango."} /> : (
          <div style={{ display: "flex", flexDirection: "column" }}>
            <div style={{ display: "flex", fontSize: 11, textTransform: "uppercase", letterSpacing: 0.5, color: "#8a9698", paddingBottom: 8, borderBottom: "1px solid #eee6d5" }}>
              <div style={{ flex: 1.2 }}>Mes</div>
              <div style={{ flex: 1, textAlign: "right" }}>Ingresos</div>
              <div style={{ flex: 1, textAlign: "right" }}>Gastos</div>
              <div style={{ flex: 1, textAlign: "right" }}>Tarjetas</div>
              <div style={{ flex: 1, textAlign: "right" }}>Balance</div>
            </div>
            {[...chartData].reverse().map((m, i) => {
              const bal = m.Ingresos - m.Gastos;
              return (
                <div key={i} style={{ display: "flex", padding: "8px 0", borderBottom: "1px solid #f2eee2", fontSize: 12.5, alignItems: "center" }}>
                  <div style={{ flex: 1.2, fontWeight: 600 }}>{m.mes}</div>
                  <div style={{ flex: 1, textAlign: "right", color: GREEN, fontFamily: "'IBM Plex Mono', monospace", whiteSpace: "nowrap" }}>{fmtARS(m.Ingresos)}</div>
                  <div style={{ flex: 1, textAlign: "right", color: BRICK, fontFamily: "'IBM Plex Mono', monospace", whiteSpace: "nowrap" }}>{fmtARS(m.Gastos)}</div>
                  <div style={{ flex: 1, textAlign: "right", color: GOLD, fontFamily: "'IBM Plex Mono', monospace", whiteSpace: "nowrap" }}>{fmtARS(m.Tarjetas)}</div>
                  <div style={{ flex: 1, textAlign: "right", fontWeight: 700, color: bal >= 0 ? GREEN : BRICK, fontFamily: "'IBM Plex Mono', monospace", whiteSpace: "nowrap" }}>{fmtARS(bal)}</div>
                </div>
              );
            })}
          </div>
          )
        ) : (
          // En el celu: siempre 3 meses fijos, navegables con las flechas
          // de arriba — independiente del rango elegido en Evolución.
          <div style={{ display: "flex", flexDirection: "column", gap: 8 }}>
            {[...resumenMensualMobile].reverse().map((m, i) => {
              const bal = m.Ingresos - m.Gastos;
              return (
                <div key={i} style={{ border: "1px solid #f2eee2", borderRadius: 8, padding: "10px 12px" }}>
                  <div style={{ fontWeight: 700, fontSize: 13, marginBottom: 6, textTransform: "capitalize" }}>{m.mes}</div>
                  <div style={{ display: "flex", justifyContent: "space-between", fontSize: 12.5, marginBottom: 3 }}>
                    <span style={{ color: "#8a9698" }}>Ingresos</span>
                    <span style={{ color: GREEN, fontFamily: "'IBM Plex Mono', monospace", whiteSpace: "nowrap" }}>{fmtARS(m.Ingresos)}</span>
                  </div>
                  <div style={{ display: "flex", justifyContent: "space-between", fontSize: 12.5, marginBottom: 3 }}>
                    <span style={{ color: "#8a9698" }}>Gastos</span>
                    <span style={{ color: BRICK, fontFamily: "'IBM Plex Mono', monospace", whiteSpace: "nowrap" }}>{fmtARS(m.Gastos)}</span>
                  </div>
                  <div style={{ display: "flex", justifyContent: "space-between", fontSize: 12.5, marginBottom: 3 }}>
                    <span style={{ color: "#8a9698" }}>Tarjetas</span>
                    <span style={{ color: GOLD, fontFamily: "'IBM Plex Mono', monospace", whiteSpace: "nowrap" }}>{fmtARS(m.Tarjetas)}</span>
                  </div>
                  <div style={{ display: "flex", justifyContent: "space-between", fontSize: 12.5, paddingTop: 5, marginTop: 3, borderTop: "1px dashed #eee6d5" }}>
                    <span style={{ fontWeight: 700 }}>Balance</span>
                    <span style={{ fontWeight: 700, color: bal >= 0 ? GREEN : BRICK, fontFamily: "'IBM Plex Mono', monospace", whiteSpace: "nowrap" }}>{fmtARS(bal)}</span>
                  </div>
                </div>
              );
            })}
          </div>
        )}
      </div>

      <div style={{ background: "#fff", borderRadius: 10, padding: 18, boxShadow: "0 2px 10px rgba(27,42,46,0.06)" }}>
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", flexWrap: "wrap", gap: 8, marginBottom: 4 }}>
          <div style={{ fontFamily: "'Fraunces', serif", fontSize: 17 }}>Gasto por tarjeta, mes a mes</div>
          <div style={{ display: "flex", gap: 4, flexWrap: "wrap" }}>
            {[[3, "3M"], [6, "6M"], [12, "12M"], [24, "24M"], ["todo", "Todo"]].map(([v, l]) => (
              <button key={l} onClick={() => setRangoTarjetas(v)} style={{
                padding: "5px 10px", borderRadius: 16, fontSize: 11.5, fontWeight: 700, cursor: "pointer",
                border: `1px solid ${rangoTarjetas === v ? TEAL : "#ddd6c4"}`,
                background: rangoTarjetas === v ? TEAL : "#fff", color: rangoTarjetas === v ? "#fff" : INK
              }}>{l}</button>
            ))}
          </div>
        </div>
        <div style={{ fontSize: 12, color: "#8a9698", marginBottom: 12 }}>
          Cada tarjeta por separado (no sumadas) — y, dentro de una misma tarjeta compartida, separado también por titular cuando el resumen lo distingue. Elegí arriba cuántos meses ver — tiene su propio rango, independiente de "Evolución".
        </div>
        {cuentasTarjetaLabels.length === 0 ? (
          <EmptyState text="Todavía no hay gastos de tarjeta cargados." />
        ) : (
          <div style={{ display: "flex", flexDirection: "column", gap: 8, maxHeight: 520, overflowY: "auto", paddingRight: 4 }}>
            {[...chartDataTarjetasPorCuenta].reverse().map((m, i) => {
              const total = cuentasTarjetaLabels.reduce((s, c) => s + (m[c] || 0), 0);
              return (
                <div key={i} style={{ border: "1px solid #f2eee2", borderRadius: 8, padding: "10px 12px" }}>
                  <div style={{ fontWeight: 700, fontSize: 13, marginBottom: 6, textTransform: "capitalize" }}>{m.mes}</div>
                  {cuentasTarjetaLabels.filter((c) => m[c] > 0).length === 0 ? (
                    <div style={{ fontSize: 12, color: "#8a9698" }}>Sin gastos de tarjeta este mes.</div>
                  ) : (
                    <div style={isDesktop ? { display: "grid", gridTemplateColumns: "repeat(auto-fill, minmax(220px, 1fr))", columnGap: 20, rowGap: 3 } : undefined}>
                      {cuentasTarjetaLabels.filter((c) => m[c] > 0).map((c) => (
                        <div key={c} style={{ display: "flex", justifyContent: "space-between", fontSize: 12.5, marginBottom: 3 }}>
                          <span style={{ color: "#8a9698" }}>{c}</span>
                          <span style={{ color: GOLD, fontFamily: "'IBM Plex Mono', monospace", whiteSpace: "nowrap" }}>{fmtARS(m[c])}</span>
                        </div>
                      ))}
                    </div>
                  )}
                  {cuentasTarjetaLabels.filter((c) => m[c] > 0).length > 0 && (
                    <div style={{ display: "flex", justifyContent: "space-between", fontSize: 12.5, paddingTop: 5, marginTop: 3, borderTop: "1px dashed #eee6d5" }}>
                      <span style={{ fontWeight: 700 }}>Total tarjeta</span>
                      <span style={{ fontWeight: 700, fontFamily: "'IBM Plex Mono', monospace", whiteSpace: "nowrap" }}>{fmtARS(total)}</span>
                    </div>
                  )}
                </div>
              );
            })}
          </div>
        )}
      </div>

      <div style={{ background: INK, color: PAPER, borderRadius: 10, padding: 18, display: "flex", justifyContent: "space-between", alignItems: "center" }}>
        <div>
          <div style={{ fontSize: 12, color: "#9db3b0" }}>Ahorrado / invertido histórico</div>
          <div style={{ fontFamily: "'IBM Plex Mono', monospace", fontSize: 22, fontWeight: 600, marginTop: 2 }}>{fmtARS(totalAhorradoHistorico)}</div>
        </div>
        <PiggyBank size={30} color={GOLD} />
      </div>
    </div>
  );
}

function MovimientosTab({ entries, allEntries, categories, onDelete, onEditDesc, onEditAmount, onEditFull, onTogglePagado, profileName, monthLabel }) {
  const [filter, setFilter] = useState("todos");
  const [busq, setBusq] = useState("");
  const [alcance, setAlcance] = useState("mes"); // "mes" | "historial"
  const [fechaDesde, setFechaDesde] = useState("");
  const [fechaHasta, setFechaHasta] = useState("");
  const [editingId, setEditingId] = useState(null);
  const buscando = busq.trim().length > 0;
  const usaRangoFecha = Boolean(fechaDesde || fechaHasta);
  const usaTodo = alcance === "historial" || usaRangoFecha;

  const base = usaTodo ? allEntries : entries;
  const filtered = base.filter((e) => {
    const pasaTipo = filter === "todos" ? (e.type === "gasto" || e.type === "ingreso" || e.type === "cambio")
      : filter === "pendientes" ? (e.type === "gasto" && e.pagado === false)
      : e.type === filter;
    if (!pasaTipo) return false;
    if (usaRangoFecha) {
      if (fechaDesde && (e.date || "") < fechaDesde) return false;
      if (fechaHasta && (e.date || "") > fechaHasta) return false;
    }
    if (!buscando) return true;
    const q = busq.trim().toLowerCase();
    const qNum = Number(busq.trim().replace(/\./g, "").replace(",", "."));
    const matcheaMonto = !isNaN(qNum) && qNum > 0 && (
      Math.abs(Number(e.amount) - qNum) < 0.01 || // monto exacto
      String(Math.round(Number(e.amount))).includes(String(Math.round(qNum))) // "150" encuentra 150000, 21500, etc.
    );
    return (
      e.desc?.toLowerCase().includes(q) ||
      e.category?.toLowerCase().includes(q) ||
      e.account?.toLowerCase().includes(q) ||
      e.who?.toLowerCase().includes(q) ||
      matcheaMonto
    );
  }).sort((a, b) => (usaTodo ? (b.date || "").localeCompare(a.date || "") : 0));

  return (
    <div>
      {monthLabel && (
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", flexWrap: "wrap", gap: 8, marginBottom: 10 }}>
          <div style={{ fontSize: 12, color: "#8a9698", textTransform: "capitalize" }}>
            {usaTodo ? <>Mostrando <b>todo el historial</b></> : <>Mostrando: <b>{monthLabel}</b></>}
          </div>
          <button
            onClick={() => exportarEntriesAExcel(filtered, "filtrados")}
            disabled={filtered.length === 0}
            title="Exporta solo lo que ves ahora (con la búsqueda/filtro aplicado)"
            style={{ ...btnOutline, padding: "5px 10px", fontSize: 11.5, opacity: filtered.length === 0 ? 0.5 : 1 }}
          >
            <Download size={13} /> Exportar esta vista ({filtered.length})
          </button>
        </div>
      )}
      <input
        value={busq}
        onChange={(e) => setBusq(e.target.value)}
        placeholder="Buscar por descripción, categoría, cuenta, monto o quién lo cargó..."
        style={{ ...inputStyle, marginBottom: 10 }}
      />
      <div style={{ display: "flex", flexWrap: "wrap", gap: 6, marginBottom: 10 }}>
        {[["mes", "Buscar en el mes"], ["historial", "Buscar en todo el historial"]].map(([k, l]) => (
          <button key={k} onClick={() => setAlcance(k)} disabled={usaRangoFecha} style={{
            padding: "6px 12px", borderRadius: 20, fontSize: 12, cursor: usaRangoFecha ? "default" : "pointer",
            border: `1px solid ${alcance === k && !usaRangoFecha ? GOLD : "#ddd6c4"}`,
            background: alcance === k && !usaRangoFecha ? GOLD : "#fff", color: alcance === k && !usaRangoFecha ? "#fff" : INK, fontWeight: 600,
            opacity: usaRangoFecha ? 0.5 : 1,
          }}>{l}</button>
        ))}
      </div>
      <div style={{ display: "flex", flexWrap: "wrap", alignItems: "center", gap: 8, marginBottom: 14, padding: "8px 10px", background: usaRangoFecha ? "#fbf1de" : PAPER_DIM, borderRadius: 8 }}>
        <span style={{ fontSize: 12, color: "#8a9698", fontWeight: 600 }}>O elegí una fecha puntual / rango:</span>
        <input
          type="date"
          value={fechaDesde}
          onChange={(e) => setFechaDesde(e.target.value)}
          style={{ ...inputStyle, padding: "5px 8px", fontSize: 12.5, width: 140 }}
          aria-label="Desde"
        />
        <span style={{ fontSize: 12, color: "#8a9698" }}>a</span>
        <input
          type="date"
          value={fechaHasta}
          onChange={(e) => setFechaHasta(e.target.value)}
          style={{ ...inputStyle, padding: "5px 8px", fontSize: 12.5, width: 140 }}
          aria-label="Hasta"
        />
        <span style={{ fontSize: 11, color: "#8a9698" }}>(dejá "Hasta" vacío para desde esa fecha en adelante; poné la misma en las dos para un solo día)</span>
        {usaRangoFecha && (
          <button onClick={() => { setFechaDesde(""); setFechaHasta(""); }} style={{ background: "none", border: `1px solid ${BRICK}`, color: BRICK, borderRadius: 6, fontSize: 11.5, padding: "4px 10px", cursor: "pointer", fontWeight: 700 }}>
            ✕ Quitar filtro de fecha
          </button>
        )}
      </div>
      <div
        style={{
          display: "flex", gap: 8, marginBottom: 14,
          overflowX: "auto", overflowY: "hidden", flexWrap: "nowrap",
          paddingBottom: 4, marginLeft: -2, paddingLeft: 2,
          scrollbarWidth: "none", msOverflowStyle: "none", WebkitOverflowScrolling: "touch",
        }}
        className="chips-scroll"
      >
        <style>{`.chips-scroll::-webkit-scrollbar { display: none; }`}</style>
        {[
          ["todos", "Todos", List],
          ["gasto", "Gastos", ArrowDownRight],
          ["ingreso", "Ingresos", ArrowUpRight],
          ["cambio", "Cambios USD", ArrowLeftRight],
          ["pendientes", "Pendientes", Clock],
        ].map(([k, l, Icon]) => {
          const active = filter === k;
          const cantidad = base.filter((e) =>
            k === "todos" ? (e.type === "gasto" || e.type === "ingreso" || e.type === "cambio")
            : k === "pendientes" ? (e.type === "gasto" && e.pagado === false)
            : e.type === k
          ).length;
          return (
            <button
              key={k}
              onClick={() => setFilter(k)}
              style={{
                display: "flex", alignItems: "center", gap: 6, flexShrink: 0,
                padding: "8px 14px", borderRadius: 22, fontSize: 12.5, fontWeight: 700, cursor: "pointer",
                border: "none",
                background: active ? TEAL : "#fff",
                color: active ? "#fff" : INK,
                boxShadow: active ? "0 3px 10px rgba(15,110,110,0.3)" : "0 1px 3px rgba(27,42,46,0.08)",
                transition: "background 0.15s ease, box-shadow 0.15s ease",
              }}
            >
              <Icon size={14} />
              {l}
              {cantidad > 0 && (
                <span style={{
                  fontSize: 10.5, fontWeight: 700, padding: "1px 6px", borderRadius: 10,
                  background: active ? "rgba(255,255,255,0.25)" : PAPER_DIM,
                  color: active ? "#fff" : "#8a9698",
                }}>{cantidad}</span>
              )}
            </button>
          );
        })}
      </div>
      {filtered.length === 0 ? <EmptyState text="No hay movimientos para mostrar." /> : (
        <div style={{ display: "flex", flexDirection: "column", gap: 8 }}>
          {filtered.map((e) => {
            return (
              <div key={e.id} style={{ background: "#fff", borderRadius: 8, padding: "12px 14px", display: "flex", justifyContent: "space-between", alignItems: "center", boxShadow: "0 1px 4px rgba(27,42,46,0.06)", gap: 10 }}>
                <div style={{ display: "flex", gap: 10, alignItems: "center", minWidth: 0, flex: 1 }}>
                  <div style={{
                    width: 34, height: 34, borderRadius: "50%", flexShrink: 0,
                    background: e.type === "ingreso" ? "#e4f0e8" : e.type === "cambio" ? "#e3eeee" : "#f7e9e6",
                    display: "flex", alignItems: "center", justifyContent: "center"
                  }}>
                    {e.type === "ingreso" ? <TrendingUp size={16} color={GREEN} /> : e.type === "cambio" ? <Landmark size={16} color={TEAL} /> : <TrendingDown size={16} color={BRICK} />}
                  </div>
                  <div style={{ minWidth: 0, flex: 1 }}>
                    <div style={{ display: "flex", alignItems: "center", gap: 6, minWidth: 0, flexWrap: "wrap" }}>
                      <div style={{ fontSize: 14, fontWeight: 600, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>{e.category}{e.desc ? ` · ${e.desc}` : ""}</div>
                      {e.type !== "cambio" && (
                        <button onClick={() => setEditingId(e.id)} style={{ border: "none", background: "none", cursor: "pointer", color: "#c4bda8", flexShrink: 0, padding: 2 }} aria-label="Editar movimiento">
                          <Pencil size={13} />
                        </button>
                      )}
                      {e.type === "gasto" && e.pagado === false && (
                        <button
                          onClick={() => onTogglePagado(e.id, true)}
                          title="Tocá para marcar como pagado"
                          style={{
                            fontSize: 10, fontWeight: 700, padding: "2px 8px", borderRadius: 10,
                            background: "#fdf1de", color: GOLD, border: `1px solid ${GOLD}`, cursor: "pointer", flexShrink: 0,
                          }}
                        >
                          Pendiente
                        </button>
                      )}
                    </div>
                    <div style={{ fontSize: 11.5, color: "#9a9488" }}>
                      {e.date} · {e.who}{e.account ? ` · ${e.account}` : ""}{e.type === "cambio" ? ` · USD ${e.usdAmount} a $${e.rate}` : ""}
                    </div>
                  </div>
                </div>
                <div style={{ display: "flex", alignItems: "center", gap: 10, flexShrink: 0 }}>
                  <div style={{ fontFamily: "'IBM Plex Mono', monospace", fontWeight: 600, color: e.type === "ingreso" ? GREEN : e.type === "cambio" ? TEAL : BRICK }}>
                    {e.type === "ingreso" ? "+" : e.type === "cambio" ? "" : "-"}{fmtARS(e.amount)}
                  </div>
                  <button onClick={() => onDelete(e.id)} style={{ border: "none", background: "none", cursor: "pointer", color: "#b8b2a4" }} aria-label="Borrar">
                    <Trash2 size={15} />
                  </button>
                </div>
              </div>
            );
          })}
        </div>
      )}

      {editingId && (() => {
        const entryEditando = filtered.find((en) => en.id === editingId) || allEntries.find((en) => en.id === editingId);
        if (!entryEditando) return null;
        return (
          <EditarMovimientoModal
            entry={entryEditando}
            categories={categories}
            onClose={() => setEditingId(null)}
            onSave={async (cambios) => { await onEditFull(entryEditando.id, cambios); setEditingId(null); }}
            onDelete={() => { onDelete(entryEditando.id); setEditingId(null); }}
          />
        );
      })()}
    </div>
  );
}

function EditarMovimientoModal({ entry, categories, onClose, onSave, onDelete }) {
  const [category, setCategory] = useState(entry.category || (categories?.[0] || ""));
  const [desc, setDesc] = useState(entry.desc || "");
  const [amount, setAmount] = useState(String(entry.amount ?? ""));
  const [date, setDate] = useState(entry.date || todayISO());
  const [account, setAccount] = useState(entry.account || "");
  const [pagado, setPagado] = useState(entry.pagado !== false);
  const [guardando, setGuardando] = useState(false);
  const [confirmarBorrado, setConfirmarBorrado] = useState(false);

  async function handleGuardar() {
    const montoNum = Number(amount);
    if (!montoNum || montoNum <= 0) return;
    setGuardando(true);
    await onSave({
      category, desc: desc.trim(), amount: montoNum, date, account: account.trim(),
      ...(entry.type === "gasto" ? { pagado } : {}),
    });
    setGuardando(false);
  }

  return (
    <div style={{ position: "fixed", inset: 0, background: "rgba(27,42,46,0.5)", display: "flex", alignItems: esPantallaAncha() ? "center" : "flex-end", justifyContent: "center", zIndex: 25 }} onClick={onClose}>
      <div onClick={(e) => e.stopPropagation()} style={{ background: PAPER, width: esPantallaAncha() ? 480 : "100%", maxWidth: 480, borderRadius: esPantallaAncha() ? 16 : "16px 16px 0 0", padding: "20px 20px 28px", maxHeight: "88vh", overflowY: "auto", boxShadow: esPantallaAncha() ? "0 20px 60px rgba(27,42,46,0.3)" : "none" }}>
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 4 }}>
          <div style={{ fontFamily: "'Fraunces', serif", fontSize: 19 }}>Editar movimiento</div>
          <button onClick={onClose} style={{ border: "none", background: "none", cursor: "pointer" }} aria-label="Cerrar"><X size={22} /></button>
        </div>
        <div style={{ fontSize: 11, color: "#8a9698", marginBottom: 16 }}>
          Cargado por <b>{entry.who || "—"}</b> · {ORIGEN_LABELS[entry.origen] || "Manual"}
          {entry.createdAt && <> · {new Date(entry.createdAt).toLocaleString("es-AR", { day: "2-digit", month: "2-digit", year: "numeric", hour: "2-digit", minute: "2-digit" })}</>}
        </div>

        <label style={labelStyle}>Categoría</label>
        <select value={category} onChange={(e) => setCategory(e.target.value)} style={{ ...inputStyle, marginBottom: 14 }}>
          {(categories || []).includes(category) ? null : <option value={category}>{category}</option>}
          {(categories || []).map((c) => <option key={c} value={c}>{c}</option>)}
        </select>

        <label style={labelStyle}>Descripción</label>
        <input value={desc} onChange={(e) => setDesc(e.target.value)} placeholder="Descripción (opcional)" style={{ ...inputStyle, marginBottom: 14 }} />

        <label style={labelStyle}>Monto</label>
        <input
          type="number"
          value={amount}
          onChange={(e) => setAmount(e.target.value)}
          style={{ ...inputStyle, marginBottom: 14, fontFamily: "'IBM Plex Mono', monospace", fontSize: 18 }}
        />

        <label style={labelStyle}>Fecha</label>
        <div style={{ marginBottom: 14 }}><FechaInput value={date} onChange={setDate} /></div>

        <label style={labelStyle}>Cuenta (opcional)</label>
        <input value={account} onChange={(e) => setAccount(e.target.value)} placeholder="Ej: Visa BBVA, Efectivo..." style={{ ...inputStyle, marginBottom: 14 }} />

        {entry.type === "gasto" && (
          <>
            <label style={labelStyle}>Estado</label>
            <div style={{ display: "flex", gap: 6, marginBottom: 20 }}>
              {[[true, "Pagado", GREEN], [false, "Pendiente de pago", GOLD]].map(([val, label, color]) => (
                <button key={String(val)} onClick={() => setPagado(val)} style={{
                  flex: 1, padding: "8px 6px", borderRadius: 8, fontSize: 12.5, fontWeight: 700, cursor: "pointer",
                  border: `1.5px solid ${pagado === val ? color : "#ddd6c4"}`,
                  background: pagado === val ? color : "#fff", color: pagado === val ? "#fff" : INK,
                }}>{label}</button>
              ))}
            </div>
          </>
        )}

        <button onClick={handleGuardar} disabled={guardando} style={{ ...btnPrimary, width: "100%", justifyContent: "center", padding: "13px", marginBottom: 10 }}>
          {guardando ? "Guardando..." : "Guardar cambios"}
        </button>

        {!confirmarBorrado ? (
          <button onClick={() => setConfirmarBorrado(true)} style={{ ...btnOutline, width: "100%", justifyContent: "center", color: BRICK, borderColor: BRICK }}>
            <Trash2 size={15} /> Borrar este movimiento
          </button>
        ) : (
          <div style={{ display: "flex", gap: 8 }}>
            <button onClick={onDelete} style={{ ...btnPrimary, background: BRICK, flex: 1, justifyContent: "center" }}>Sí, borrar</button>
            <button onClick={() => setConfirmarBorrado(false)} style={{ ...btnOutline, flex: 1, justifyContent: "center" }}>Cancelar</button>
          </div>
        )}
      </div>
    </div>
  );
}

function AhorrosTab({ entries, onDelete, totalAhorradoHistorico }) {
  const porInstrumento = useMemo(() => {
    const map = {};
    entries.forEach((e) => { map[e.category] = (map[e.category] || 0) + Number(e.amount); });
    return Object.entries(map).map(([name, value]) => ({ name, value })).sort((a, b) => b.value - a.value);
  }, [entries]);

  return (
    <div style={{ display: "flex", flexDirection: "column", gap: 16 }}>
      <div style={{ background: "#fff", borderRadius: 10, padding: 18, boxShadow: "0 2px 10px rgba(27,42,46,0.06)" }}>
        <div style={{ fontSize: 12, color: "#8a9698", textTransform: "uppercase", letterSpacing: 0.5 }}>Total ahorrado / invertido</div>
        <div style={{ fontFamily: "'IBM Plex Mono', monospace", fontSize: 26, fontWeight: 600, color: GOLD, marginTop: 4 }}>{fmtARS(totalAhorradoHistorico)}</div>
        {porInstrumento.length > 0 && (
          <div style={{ marginTop: 14, display: "flex", flexDirection: "column", gap: 6 }}>
            {porInstrumento.map((p, i) => (
              <div key={p.name} style={{ display: "flex", justifyContent: "space-between", fontSize: 13 }}>
                <span style={{ display: "flex", alignItems: "center", gap: 6 }}>
                  <span style={{ width: 8, height: 8, borderRadius: "50%", background: CAT_COLORS[i % CAT_COLORS.length] }} />
                  {p.name}
                </span>
                <span style={{ fontFamily: "'IBM Plex Mono', monospace" }}>{fmtARS(p.value)}</span>
              </div>
            ))}
          </div>
        )}
      </div>
      {entries.length === 0 ? <EmptyState text="Todavía no registraste ahorros ni inversiones." /> : (
        <div style={{ display: "flex", flexDirection: "column", gap: 8 }}>
          {entries.map((e) => (
            <div key={e.id} style={{ background: "#fff", borderRadius: 8, padding: "12px 14px", display: "flex", justifyContent: "space-between", alignItems: "center", boxShadow: "0 1px 4px rgba(27,42,46,0.06)" }}>
              <div>
                <div style={{ fontSize: 14, fontWeight: 600 }}>{e.category}{e.desc ? ` · ${e.desc}` : ""}</div>
                <div style={{ fontSize: 11.5, color: "#9a9488" }}>{e.date} · {e.who}</div>
              </div>
              <div style={{ display: "flex", alignItems: "center", gap: 10 }}>
                <div style={{ fontFamily: "'IBM Plex Mono', monospace", fontWeight: 600, color: GOLD }}>{fmtARS(e.amount)}</div>
                <button onClick={() => onDelete(e.id)} style={{ border: "none", background: "none", cursor: "pointer", color: "#b8b2a4" }} aria-label="Borrar">
                  <Trash2 size={15} />
                </button>
              </div>
            </div>
          ))}
        </div>
      )}
    </div>
  );
}

function PresupuestosTab({ budgets, onUpdate, gastosPorCategoria, categories, isDesktop }) {
  const [local, setLocal] = useState(budgets);
  useEffect(() => setLocal(budgets), [budgets]);

  function handleChange(cat, val) {
    setLocal((prev) => ({ ...prev, [cat]: val }));
  }
  function handleBlur(cat, val) {
    onUpdate({ ...budgets, [cat]: Number(val) || 0 });
  }
  const gastoMap = Object.fromEntries(gastosPorCategoria.map((g) => [g.name, g.value]));

  return (
    <div style={{ display: "flex", flexDirection: "column", gap: 10 }}>
      <div style={{ fontSize: 13, color: "#8a9698", marginBottom: 4 }}>Definí un tope mensual por categoría. Se compara contra lo gastado este mes.</div>
      <div style={isDesktop ? { display: "grid", gridTemplateColumns: "1fr 1fr", gap: 10 } : { display: "contents" }}>
      {categories.map((cat) => {
        const limit = Number(local[cat]) || 0;
        const spent = gastoMap[cat] || 0;
        const pct = limit > 0 ? Math.min(100, (spent / limit) * 100) : 0;
        const over = limit > 0 && spent > limit;
        return (
          <div key={cat} style={{ background: "#fff", borderRadius: 8, padding: "12px 14px", boxShadow: "0 1px 4px rgba(27,42,46,0.06)" }}>
            <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 6 }}>
              <div style={{ fontSize: 14, fontWeight: 600 }}>{cat}</div>
              <input
                type="number"
                value={local[cat] ?? ""}
                onChange={(e) => handleChange(cat, e.target.value)}
                onBlur={(e) => handleBlur(cat, e.target.value)}
                placeholder="Sin tope"
                style={{ width: 100, padding: "4px 8px", borderRadius: 6, border: "1px solid #ddd6c4", fontSize: 13, textAlign: "right", fontFamily: "'IBM Plex Mono', monospace" }}
              />
            </div>
            {limit > 0 && (
              <>
                <div style={{ height: 6, borderRadius: 4, background: PAPER_DIM, overflow: "hidden" }}>
                  <div style={{ width: `${pct}%`, height: "100%", background: over ? BRICK : TEAL, transition: "width 0.3s" }} />
                </div>
                <div style={{ fontSize: 11.5, color: over ? BRICK : "#8a9698", marginTop: 4 }}>
                  {fmtARS(spent)} de {fmtARS(limit)} {over ? "· ¡superado!" : ""}
                </div>
              </>
            )}
          </div>
        );
      })}
      </div>
    </div>
  );
}

const CSV_PLACEHOLDER = `fecha,tipo,categoria,monto,descripcion,cuenta,usd,tc
2026-07-03,gasto,Comida,15400,Supermercado Coto,Visa BBVA Hernán,,
2026-07-10,ingreso,Sueldo,850000,Sueldo julio,,,
2026-07-15,cambio,,1561456.98,Cambio en ARQ,ARQ,1000,1561.46`;

// build-fix-negativos-v2
function parseCSV(text) {
  const lines = text.trim().split("\n").filter((l) => l.trim());
  if (lines.length === 0) return { rows: [], errors: [] };
  const header = lines[0].toLowerCase().includes("fecha") ? lines.slice(1) : lines;
  const rows = [];
  const errors = [];
  header.forEach((line, i) => {
    const parts = line.split(",").map((p) => p.trim());
    if (parts.length < 4) { errors.push(`Línea ${i + 1}: faltan columnas`); return; }
    const [fecha, tipo, categoria, monto, desc, cuenta, usd, tc] = parts;
    const tipoNorm = tipo.toLowerCase();
    if (!["gasto", "ingreso", "ahorro", "cambio"].includes(tipoNorm)) { errors.push(`Línea ${i + 1}: tipo "${tipo}" inválido (usar gasto/ingreso/ahorro/cambio)`); return; }
    const montoNum = Number(monto.replace(/[^0-9.-]/g, ""));
    if (montoNum === 0 || Number.isNaN(montoNum)) { errors.push(`Línea ${i + 1}: monto inválido`); return; }
    if (!/^\d{4}-\d{2}-\d{2}$/.test(fecha)) { errors.push(`Línea ${i + 1}: fecha debe ser AAAA-MM-DD`); return; }
    if (tipoNorm === "cambio") {
      const usdNum = Number((usd || "").replace(/[^0-9.-]/g, ""));
      const tcNum = Number((tc || "").replace(/[^0-9.-]/g, ""));
      if (!usdNum || usdNum <= 0) { errors.push(`Línea ${i + 1}: cambio necesita columna usd válida`); return; }
      rows.push({ date: fecha, type: "cambio", category: "Cambio USD→ARS", amount: montoNum, desc: desc || "", account: cuenta || "", usdAmount: usdNum, rate: tcNum || Math.round(montoNum / usdNum), origen: "csv" });
      return;
    }
    rows.push({ date: fecha, type: tipoNorm, category: categoria || "Otros", amount: montoNum, desc: desc || "", account: cuenta || "", origen: "csv" });
  });
  return { rows, errors };
}

function ImportarTab({ onImport, categoryOverrides }) {
  const [modo, setModo] = useState("auto"); // "auto" | "bbva" | "mercadopago" | "colegio" | "csv"
  const [text, setText] = useState("");
  const [result, setResult] = useState(null);
  const [importing, setImporting] = useState(false);
  const [fileName, setFileName] = useState("");
  const [lastImportMsg, setLastImportMsg] = useState(null);
  const [pdfProcesando, setPdfProcesando] = useState(false);
  const [pdfAvisos, setPdfAvisos] = useState([]);
  const [pdfNombres, setPdfNombres] = useState([]);
  const [pdfResumenes, setPdfResumenes] = useState([]);

  // Detecta el formato de cada PDF por su propio contenido (no hace
  // falta elegir el banco/billetera de antemano) y lo manda al parser
  // que corresponda. Si no reconoce ninguno, avisa en vez de forzar un
  // parser equivocado.
  async function handlePdfFilesAuto(e) {
    const files = Array.from(e.target.files || []);
    if (files.length === 0) return;
    setPdfProcesando(true);
    setPdfAvisos([]);
    setPdfNombres(files.map((f) => f.name));
    setPdfResumenes([]);
    let todasLasFilas = [];
    let avisos = [];
    let resumenes = [];
    for (const file of files) {
      try {
        const texto = await conTimeout(extraerTextoPdf(file), 30000, `${file.name}: tardó demasiado en procesarse (más de 30s), lo salteo.`);
        let formatoDetectado = null;
        if (/ASOCIACI[ÓO]N ORT ARGENTINA/i.test(texto) || /COMUNIDAD BETEL/i.test(texto)) {
          formatoDetectado = "colegio";
        } else if (/Mercado\s*Pago/i.test(texto) && /(RESUMEN DE CUENTA|ID de la operaci[óo]n)/i.test(texto)) {
          formatoDetectado = "mercadopago";
        } else if (/BBVA/i.test(texto) && /(Visa Signature|Mastercard Black)/i.test(texto)) {
          formatoDetectado = "bbva";
        }

        if (formatoDetectado === "bbva") {
          const lineas = await conTimeout(extraerLineasPdf(file), 30000, `${file.name}: tardó demasiado en procesarse (más de 30s), lo salteo.`);
          const { filas, avisos: av, resumenInfo } = parsearResumenBBVA(lineas, file.name, categoryOverrides);
          if (resumenInfo) {
            const sumaFilas = filas.reduce((s, f) => s + Number(f.amount), 0);
            resumenes.push({ archivo: file.name, cuenta: resumenInfo.cuenta, total: resumenInfo.total, sumaFilas });
          }
          todasLasFilas = todasLasFilas.concat(filas);
          avisos = avisos.concat(av.map((a) => `${file.name} (BBVA): ${a}`));
        } else if (formatoDetectado === "mercadopago") {
          const { filas, avisos: av } = parsearResumenMercadoPago(texto, categoryOverrides);
          todasLasFilas = todasLasFilas.concat(filas);
          avisos = avisos.concat(av.map((a) => `${file.name} (Mercado Pago): ${a}`));
        } else if (formatoDetectado === "colegio") {
          const { filas, avisos: av } = parsearFacturaColegio(texto, categoryOverrides);
          todasLasFilas = todasLasFilas.concat(filas);
          avisos = avisos.concat(av.map((a) => `${file.name} (Colegio): ${a}`));
        } else {
          avisos.push(`${file.name}: no reconocí el formato (no parece ser BBVA, Mercado Pago, ni un colegio soportado por ahora). Pasámelo a mí en el chat para agregar ese formato.`);
        }
      } catch (err) {
        avisos.push(`${file.name}: no pude leer el PDF (${err.message}).`);
      }
    }
    setPdfAvisos(avisos);
    setPdfResumenes(resumenes);
    setResult({ rows: todasLasFilas, errors: [] });
    setPdfProcesando(false);
  }

  // Achica una imagen antes de mandarla a analizar (mismo criterio que la
  // foto de recibo individual: para leer texto alcanza con bastante menos
  // que lo que saca la cámara, así sale más rápido y barato).
  function comprimirImagen(file) {
    return new Promise((resolve, reject) => {
      const reader = new FileReader();
      reader.onload = (e) => {
        const img = new Image();
        img.onload = () => {
          const maxAncho = 1280;
          const escala = Math.min(1, maxAncho / img.width);
          const canvas = document.createElement("canvas");
          canvas.width = Math.round(img.width * escala);
          canvas.height = Math.round(img.height * escala);
          const ctx = canvas.getContext("2d");
          ctx.drawImage(img, 0, 0, canvas.width, canvas.height);
          canvas.toBlob((blob) => resolve(blob), "image/jpeg", 0.8);
        };
        img.onerror = reject;
        img.src = e.target.result;
      };
      reader.onerror = reject;
      reader.readAsDataURL(file);
    });
  }

  // Lee una o varias capturas de pantalla que muestran VARIOS movimientos
  // juntos (ej. el listado de "Transacciones" de ARQ: transferencias
  // enviadas, recibidas, y cambios de divisa) y arma las filas para
  // importar, igual que con un PDF.
  async function handleCapturas(e) {
    const files = Array.from(e.target.files || []);
    if (files.length === 0) return;
    setPdfProcesando(true);
    setPdfAvisos([]);
    setPdfNombres(files.map((f) => f.name));
    setPdfResumenes([]);
    let todasLasFilas = [];
    let avisos = [];
    for (const file of files) {
      try {
        const blob = await comprimirImagen(file);
        const image_base64 = await blobToBase64(blob);
        const resultado = await conTimeout(
          sbFunction("analizar-movimientos", { image_base64, mime_type: "image/jpeg" }),
          45000,
          `${file.name}: tardó demasiado en analizarse (más de 45s), la salteo.`
        );
        if (resultado?.error) throw new Error(resultado.error);
        const movimientos = Array.isArray(resultado?.movimientos) ? resultado.movimientos : [];
        if (movimientos.length === 0) {
          avisos.push(`${file.name}: no encontré movimientos legibles en esta captura.`);
        }
        movimientos.forEach((m) => {
          if (!m.fecha || !m.monto || !m.tipo) return;
          if (m.tipo === "cambio") {
            const montoArs = Number(m.monto_destino ?? m.monto);
            const montoUsd = Number(m.monto);
            todasLasFilas.push({
              date: m.fecha,
              type: "cambio",
              category: "Cambio USD→ARS",
              amount: montoArs,
              usdAmount: montoUsd,
              rate: montoUsd > 0 ? Math.round((montoArs / montoUsd) * 100) / 100 : null,
              desc: m.desc || "Cambio de divisa",
              account: "ARQ",
              origen: "foto",
            });
          } else {
            const tipo = m.tipo === "ingreso" ? "ingreso" : "gasto";
            todasLasFilas.push({
              date: m.fecha,
              type: tipo,
              category: tipo === "gasto" ? inferCategory(m.desc || "", categoryOverrides) : "Otros ingresos",
              amount: Number(m.monto),
              desc: m.desc || "",
              account: "ARQ",
              origen: "foto",
            });
          }
        });
      } catch (err) {
        avisos.push(`${file.name}: no pude leer la captura (${err.message}).`);
      }
    }
    setPdfAvisos(avisos);
    setResult({ rows: todasLasFilas, errors: [] });
    setPdfProcesando(false);
  }

  async function handlePdfFiles(e) {
    const files = Array.from(e.target.files || []);
    if (files.length === 0) return;
    setPdfProcesando(true);
    setPdfAvisos([]);
    setPdfNombres(files.map((f) => f.name));
    setPdfResumenes([]);
    let todasLasFilas = [];
    let avisos = [];
    let resumenes = [];
    for (const file of files) {
      try {
        const lineas = await conTimeout(extraerLineasPdf(file), 30000, `${file.name}: tardó demasiado en procesarse (más de 30s), lo salteo.`);
        const { filas, avisos: av, resumenInfo } = parsearResumenBBVA(lineas, file.name, categoryOverrides);
        if (filas.length === 0) {
          avisos.push(`${file.name}: no reconocí movimientos con el formato BBVA. Puede ser otro banco/billetera — pasámelo a mí directamente en el chat para procesarlo.`);
        }
        if (resumenInfo) {
          const sumaFilas = filas.reduce((s, f) => s + Number(f.amount), 0);
          resumenes.push({ archivo: file.name, cuenta: resumenInfo.cuenta, total: resumenInfo.total, sumaFilas });
        }
        todasLasFilas = todasLasFilas.concat(filas);
        avisos = avisos.concat(av.map((a) => `${file.name}: ${a}`));
      } catch (err) {
        avisos.push(`${file.name}: no pude leer el PDF (${err.message}).`);
      }
    }
    setPdfAvisos(avisos);
    setPdfResumenes(resumenes);
    setResult({ rows: todasLasFilas, errors: [] });
    setPdfProcesando(false);
  }

  async function handlePdfFilesMP(e) {
    const files = Array.from(e.target.files || []);
    if (files.length === 0) return;
    setPdfProcesando(true);
    setPdfAvisos([]);
    setPdfNombres(files.map((f) => f.name));
    let todasLasFilas = [];
    let avisos = [];
    for (const file of files) {
      try {
        const texto = await conTimeout(extraerTextoPdf(file), 30000, `${file.name}: tardó demasiado en procesarse (más de 30s), lo salteo.`);
        const { filas, avisos: av } = parsearResumenMercadoPago(texto, categoryOverrides);
        if (filas.length === 0) {
          avisos.push(`${file.name}: no reconocí movimientos con el formato de Mercado Pago. Pasámelo a mí en el chat para revisarlo.`);
        }
        todasLasFilas = todasLasFilas.concat(filas);
        avisos = avisos.concat(av.map((a) => `${file.name}: ${a}`));
      } catch (err) {
        avisos.push(`${file.name}: no pude leer el PDF (${err.message}).`);
      }
    }
    setPdfAvisos(avisos);
    setResult({ rows: todasLasFilas, errors: [] });
    setPdfProcesando(false);
  }

  async function handlePdfFilesColegio(e) {
    const files = Array.from(e.target.files || []);
    if (files.length === 0) return;
    setPdfProcesando(true);
    setPdfAvisos([]);
    setPdfNombres(files.map((f) => f.name));
    let todasLasFilas = [];
    let avisos = [];
    for (const file of files) {
      try {
        const texto = await conTimeout(extraerTextoPdf(file), 30000, `${file.name}: tardó demasiado en procesarse (más de 30s), lo salteo.`);
        const { filas, avisos: av } = parsearFacturaColegio(texto, categoryOverrides);
        todasLasFilas = todasLasFilas.concat(filas);
        avisos = avisos.concat(av.map((a) => `${file.name}: ${a}`));
      } catch (err) {
        avisos.push(`${file.name}: no pude leer el PDF (${err.message}).`);
      }
    }
    setPdfAvisos(avisos);
    setResult({ rows: todasLasFilas, errors: [] });
    setPdfProcesando(false);
  }

  function handleFile(e) {
    const file = e.target.files?.[0];
    if (!file) return;
    setFileName(file.name);
    const reader = new FileReader();
    reader.onload = (ev) => {
      setText(ev.target.result);
      setResult(null);
    };
    reader.onerror = () => alert("No pude leer el archivo. Probá copiando y pegando el texto.");
    reader.readAsText(file);
  }

  function handlePreview() {
    setResult(parseCSV(text));
  }
  async function handleImport() {
    if (!result || result.rows.length === 0) return;
    setImporting(true);
    const formatoLabel = { bbva: "PDF BBVA", mercadopago: "PDF Mercado Pago", colegio: "PDF Colegio", csv: "CSV/texto" }[modo] || modo;
    try {
      const res = await onImport(result.rows, formatoLabel, pdfNombres);
      if (res?.error) {
        setLastImportMsg({ tone: "warn", text: `No se completó la importación: ${res.error}` });
        return;
      }
      setText("");
      setResult(null);
      setPdfNombres([]);
      if (res && typeof res === "object") {
        setLastImportMsg({
          tone: res.duplicates > 0 ? "warn" : "ok",
          text: res.duplicates > 0
            ? `Se importaron ${res.imported} movimientos. Se descartaron ${res.duplicates} por ser duplicados (misma fecha, monto, descripción y cuenta que uno ya cargado).`
            : `Se importaron ${res.imported} movimientos. Sin duplicados detectados.`,
        });
      } else {
        setLastImportMsg({ tone: "ok", text: `Se importaron ${res} movimientos.` });
      }
    } catch (err) {
      setLastImportMsg({ tone: "warn", text: `No se pudo importar: ${err.message || "error inesperado"}. Probá de nuevo.` });
    } finally {
      setImporting(false);
    }
  }

  function cambiarModo(nuevoModo) {
    setModo(nuevoModo);
    setResult(null);
    setPdfAvisos([]);
    setPdfNombres([]);
    setText("");
    setFileName("");
  }

  return (
    <div style={{ display: "flex", flexDirection: "column", gap: 14 }}>
      {lastImportMsg && (
        <div style={{
          background: lastImportMsg.tone === "warn" ? "#fbf1de" : "#e8f3ec",
          border: `1px solid ${lastImportMsg.tone === "warn" ? GOLD : GREEN}`,
          borderRadius: 8, padding: "12px 14px", fontSize: 13, display: "flex", justifyContent: "space-between", gap: 10, alignItems: "center",
        }}>
          <span>{lastImportMsg.text}</span>
          <button onClick={() => setLastImportMsg(null)} style={{ background: "none", border: "none", cursor: "pointer", color: "#8a9698" }}>
            <X size={15} />
          </button>
        </div>
      )}

      <div style={{ background: "#fff", borderRadius: 10, padding: 18, boxShadow: "0 2px 10px rgba(27,42,46,0.06)" }}>
        <div style={{ fontFamily: "'Fraunces', serif", fontSize: 17, marginBottom: 12 }}>¿Qué querés importar?</div>
        <div style={{ display: "flex", flexWrap: "wrap", gap: 8 }}>
          {[
            ["auto", "PDF (detecta el formato solo)"],
            ["bbva", "PDF tarjeta BBVA"],
            ["mercadopago", "PDF Mercado Pago"],
            ["colegio", "PDF colegio (arancel)"],
            ["captura", "Captura de pantalla (ARQ, etc.)"],
            ["csv", "Pegar texto / CSV"],
          ].map(([k, l]) => (
            <button key={k} onClick={() => cambiarModo(k)} style={{
              padding: "9px 14px", borderRadius: 8, fontSize: 13, fontWeight: 700, cursor: "pointer",
              border: `1.5px solid ${modo === k ? TEAL : "#ddd6c4"}`,
              background: modo === k ? TEAL : "#fff", color: modo === k ? "#fff" : INK,
            }}>{l}</button>
          ))}
        </div>
      </div>

      {modo === "auto" && (
        <div style={{ background: "#fff", borderRadius: 10, padding: 18, boxShadow: "0 2px 10px rgba(27,42,46,0.06)" }}>
          <div style={{ fontFamily: "'Fraunces', serif", fontSize: 17, marginBottom: 6 }}>Subir PDF (detecta el banco/billetera solo)</div>
          <div style={{ fontSize: 13, color: "#8a9698", marginBottom: 12 }}>
            Subí uno o varios PDFs mezclados — de tarjeta BBVA, resumen de Mercado Pago, o arancel de colegio — y cada uno se procesa con el parser que corresponda, sin que tengas que elegirlo vos. Si alguno no se reconoce, avisa cuál.
          </div>
          <label style={{ ...btnOutline, cursor: "pointer", marginBottom: 10, display: "inline-flex" }}>
            <input type="file" accept="application/pdf" multiple onChange={handlePdfFilesAuto} style={{ display: "none" }} />
            {pdfProcesando ? "Leyendo PDF(s)..." : "Elegir PDF(s)"}
          </label>
          {pdfNombres.length > 0 && <div style={{ fontSize: 12, color: TEAL, marginBottom: 8 }}>{pdfNombres.join(", ")}</div>}
          {pdfResumenes.length > 0 && (
            <div style={{ display: "flex", flexDirection: "column", gap: 6, marginBottom: pdfAvisos.length > 0 ? 8 : 0 }}>
              {pdfResumenes.map((r, i) => {
                const diff = r.total != null ? Math.round((r.total - r.sumaFilas) * 100) / 100 : null;
                const coincide = diff !== null && Math.abs(diff) <= 1;
                return (
                  <div key={i} style={{ background: coincide ? "#e8f3ec" : "#fdf1de", border: `1px solid ${coincide ? GREEN : GOLD}`, borderRadius: 8, padding: "10px 12px", fontSize: 12.5 }}>
                    <div style={{ fontWeight: 700, marginBottom: 2 }}>{r.cuenta} <span style={{ fontWeight: 400, color: "#8a9698" }}>({r.archivo})</span></div>
                    <div>Total del resumen: <b style={{ fontFamily: "'IBM Plex Mono', monospace" }}>{r.total != null ? fmtARS(r.total) : "no encontrado"}</b></div>
                    <div>Suma de lo que se va a importar: <b style={{ fontFamily: "'IBM Plex Mono', monospace" }}>{fmtARS(r.sumaFilas)}</b></div>
                    {!coincide && r.total != null && (
                      <div style={{ color: BRICK, marginTop: 2 }}>⚠ No coincide (diferencia de {fmtARS(Math.abs(diff))}) — revisá los avisos de abajo antes de importar.</div>
                    )}
                  </div>
                );
              })}
            </div>
          )}
          {pdfAvisos.length > 0 && (
            <div style={{ background: "#fbf1de", border: `1px solid ${GOLD}`, borderRadius: 8, padding: 10, fontSize: 12, marginTop: 4 }}>
              {pdfAvisos.map((a, i) => <div key={i}>{a}</div>)}
            </div>
          )}
        </div>
      )}

      {modo === "captura" && (
        <div style={{ background: "#fff", borderRadius: 10, padding: 18, boxShadow: "0 2px 10px rgba(27,42,46,0.06)" }}>
          <div style={{ fontFamily: "'Fraunces', serif", fontSize: 17, marginBottom: 6 }}>Captura de pantalla con varios movimientos (ARQ, etc.)</div>
          <div style={{ fontSize: 13, color: "#8a9698", marginBottom: 12 }}>
            Para apps como ARQ que no dan un PDF descargable — subí la captura de la lista de "Transacciones" tal cual, con todos los movimientos que se vean (transferencias, cambios de divisa). Se leen con IA, quedan a nombre de la cuenta "ARQ", y podés revisar todo antes de confirmar. Podés subir varias capturas juntas si necesitás cubrir más de una pantalla.
          </div>
          <label style={{ ...btnOutline, cursor: "pointer", marginBottom: 10, display: "inline-flex" }}>
            <input type="file" accept="image/*" multiple onChange={handleCapturas} style={{ display: "none" }} />
            {pdfProcesando ? "Leyendo captura(s)..." : "Elegir captura(s)"}
          </label>
          {pdfNombres.length > 0 && <div style={{ fontSize: 12, color: TEAL, marginBottom: 8 }}>{pdfNombres.join(", ")}</div>}
          {pdfAvisos.length > 0 && (
            <div style={{ background: "#fbf1de", border: `1px solid ${GOLD}`, borderRadius: 8, padding: 10, fontSize: 12, marginTop: 4 }}>
              {pdfAvisos.map((a, i) => <div key={i}>{a}</div>)}
            </div>
          )}
        </div>
      )}

      {modo === "mercadopago" && (
        <div style={{ background: "#fff", borderRadius: 10, padding: 18, boxShadow: "0 2px 10px rgba(27,42,46,0.06)" }}>
          <div style={{ fontFamily: "'Fraunces', serif", fontSize: 17, marginBottom: 6 }}>Subir resumen en PDF (Mercado Pago)</div>
          <div style={{ fontSize: 13, color: "#8a9698", marginBottom: 12 }}>
            Solo la sección en pesos (ignora tenencias en dólares). Excluye rendimientos y traspasos a vos mismo automáticamente.
          </div>
          <label style={{ ...btnOutline, cursor: "pointer", marginBottom: 10, display: "inline-flex" }}>
            <input type="file" accept="application/pdf" multiple onChange={handlePdfFilesMP} style={{ display: "none" }} />
            {pdfProcesando ? "Leyendo PDF..." : "Elegir PDF(s)"}
          </label>
          {pdfNombres.length > 0 && <div style={{ fontSize: 12, color: TEAL, marginBottom: 8 }}>{pdfNombres.join(", ")}</div>}
          {pdfAvisos.length > 0 && (
            <div style={{ background: "#fbf1de", border: `1px solid ${GOLD}`, borderRadius: 8, padding: 10, fontSize: 12, marginTop: 4 }}>
              {pdfAvisos.map((a, i) => <div key={i}>{a}</div>)}
            </div>
          )}
        </div>
      )}

      {modo === "bbva" && (
        <div style={{ background: "#fff", borderRadius: 10, padding: 18, boxShadow: "0 2px 10px rgba(27,42,46,0.06)" }}>
          <div style={{ fontFamily: "'Fraunces', serif", fontSize: 17, marginBottom: 6 }}>Subir resumen en PDF (BBVA)</div>
          <div style={{ fontSize: 13, color: "#8a9698", marginBottom: 12 }}>
            Subí uno o varios PDFs de resumen de Visa/Mastercard BBVA — se leen y categorizan solos, sin pasar por el chat. Después revisá abajo antes de importar.
          </div>
          <label style={{ ...btnOutline, cursor: "pointer", marginBottom: 10, display: "inline-flex" }}>
            <input type="file" accept="application/pdf" multiple onChange={handlePdfFiles} style={{ display: "none" }} />
            {pdfProcesando ? "Leyendo PDF..." : "Elegir PDF(s)"}
          </label>
          {pdfNombres.length > 0 && <div style={{ fontSize: 12, color: TEAL, marginBottom: 8 }}>{pdfNombres.join(", ")}</div>}
          {pdfResumenes.length > 0 && (
            <div style={{ display: "flex", flexDirection: "column", gap: 6, marginBottom: pdfAvisos.length > 0 ? 8 : 0 }}>
              {pdfResumenes.map((r, i) => {
                const diff = r.total != null ? Math.round((r.total - r.sumaFilas) * 100) / 100 : null;
                const coincide = diff !== null && Math.abs(diff) <= 1;
                return (
                  <div key={i} style={{ background: coincide ? "#e8f3ec" : "#fdf1de", border: `1px solid ${coincide ? GREEN : GOLD}`, borderRadius: 8, padding: "10px 12px", fontSize: 12.5 }}>
                    <div style={{ fontWeight: 700, marginBottom: 2 }}>{r.cuenta} <span style={{ fontWeight: 400, color: "#8a9698" }}>({r.archivo})</span></div>
                    <div>Total del resumen: <b style={{ fontFamily: "'IBM Plex Mono', monospace" }}>{r.total != null ? fmtARS(r.total) : "no encontrado"}</b></div>
                    <div>Suma de lo que se va a importar: <b style={{ fontFamily: "'IBM Plex Mono', monospace" }}>{fmtARS(r.sumaFilas)}</b></div>
                    {!coincide && r.total != null && (
                      <div style={{ color: BRICK, marginTop: 2 }}>⚠ No coincide (diferencia de {fmtARS(Math.abs(diff))}) — revisá los avisos de abajo antes de importar.</div>
                    )}
                  </div>
                );
              })}
            </div>
          )}
          {pdfAvisos.length > 0 && (
            <div style={{ background: "#fbf1de", border: `1px solid ${GOLD}`, borderRadius: 8, padding: 10, fontSize: 12, marginTop: 4 }}>
              {pdfAvisos.map((a, i) => <div key={i}>{a}</div>)}
            </div>
          )}
        </div>
      )}

      {modo === "colegio" && (
        <div style={{ background: "#fff", borderRadius: 10, padding: 18, boxShadow: "0 2px 10px rgba(27,42,46,0.06)" }}>
          <div style={{ fontFamily: "'Fraunces', serif", fontSize: 17, marginBottom: 6 }}>Subir factura de arancel (colegio)</div>
          <div style={{ fontSize: 13, color: "#8a9698", marginBottom: 12 }}>
            Por ahora reconoce el formato de <b>ORT</b>. Lee el alumno, el período y el importe del arancel, y lo categoriza como "Escuelas". Comunidad Betel todavía no está soportado — mandame un PDF de ejemplo en el chat para agregarlo.
          </div>
          <label style={{ ...btnOutline, cursor: "pointer", marginBottom: 10, display: "inline-flex" }}>
            <input type="file" accept="application/pdf" multiple onChange={handlePdfFilesColegio} style={{ display: "none" }} />
            {pdfProcesando ? "Leyendo PDF..." : "Elegir PDF(s)"}
          </label>
          {pdfNombres.length > 0 && <div style={{ fontSize: 12, color: TEAL, marginBottom: 8 }}>{pdfNombres.join(", ")}</div>}
          {pdfAvisos.length > 0 && (
            <div style={{ background: "#fbf1de", border: `1px solid ${GOLD}`, borderRadius: 8, padding: 10, fontSize: 12, marginTop: 4 }}>
              {pdfAvisos.map((a, i) => <div key={i}>{a}</div>)}
            </div>
          )}
        </div>
      )}

      {modo === "csv" && (
        <div style={{ background: "#fff", borderRadius: 10, padding: 18, boxShadow: "0 2px 10px rgba(27,42,46,0.06)" }}>
          <div style={{ fontFamily: "'Fraunces', serif", fontSize: 17, marginBottom: 6 }}>Importar movimientos (texto/CSV)</div>
          <div style={{ fontSize: 13, color: "#8a9698", marginBottom: 12 }}>
            Pegá las líneas que te pasó Claude (o armá las tuyas) con este formato: <br />
            <span style={{ fontFamily: "'IBM Plex Mono', monospace", fontSize: 11.5 }}>fecha,tipo,categoria,monto,descripcion,cuenta,usd,tc</span>
            <br />Cuenta es opcional. Las columnas usd/tc solo se usan si tipo es "cambio" (monto = ARS resultante, usd = dólares, tc = tipo de cambio).
          </div>
          <label style={{ ...btnOutline, cursor: "pointer", marginBottom: 12, display: "inline-flex" }}>
            <input type="file" accept=".csv,text/csv,text/plain" onChange={handleFile} style={{ display: "none" }} />
            Elegir archivo CSV
          </label>
          {fileName && <div style={{ fontSize: 12, color: TEAL, marginBottom: 10 }}>Archivo cargado: {fileName}</div>}
          <div style={{ fontSize: 11.5, color: "#9a9488", marginBottom: 8 }}>— o pegá el texto manualmente —</div>
          <textarea
            value={text}
            onChange={(e) => { setText(e.target.value); setResult(null); }}
            placeholder={CSV_PLACEHOLDER}
            rows={8}
            style={{ ...inputStyle, fontFamily: "'IBM Plex Mono', monospace", fontSize: 12.5, resize: "vertical" }}
          />
          <div style={{ display: "flex", gap: 8, marginTop: 12 }}>
            <button onClick={handlePreview} style={btnOutline} disabled={!text.trim()}>Previsualizar</button>
            {result && result.rows.length > 0 && (
              <button onClick={handleImport} disabled={importing} style={btnPrimary}>
                {importing ? "Importando..." : `Importar ${result.rows.length} movimientos`}
              </button>
            )}
          </div>
        </div>
      )}

      {result && modo !== "csv" && (
        <div style={{ display: "flex", justifyContent: "flex-end" }}>
          {result.rows.length > 0 && (
            <button onClick={handleImport} disabled={importing} style={btnPrimary}>
              {importing ? "Importando..." : `Importar ${result.rows.length} movimientos`}
            </button>
          )}
        </div>
      )}

      {result && (
        <div style={{ background: "#fff", borderRadius: 10, padding: 18, boxShadow: "0 2px 10px rgba(27,42,46,0.06)" }}>
          <div style={{ fontSize: 13, fontWeight: 700, marginBottom: 8 }}>
            {result.rows.length} filas listas · {result.errors.length} con error
          </div>
          {result.errors.length > 0 && (
            <div style={{ background: "#f7e9e6", borderRadius: 6, padding: 10, fontSize: 12, color: BRICK, marginBottom: 10 }}>
              {result.errors.map((e, i) => <div key={i}>{e}</div>)}
            </div>
          )}
          <div style={{ maxHeight: 220, overflowY: "auto", display: "flex", flexDirection: "column", gap: 6 }}>
            {result.rows.map((r, i) => (
              <div key={i} style={{ display: "flex", justifyContent: "space-between", fontSize: 12.5, padding: "4px 0", borderBottom: "1px solid #f0ece0" }}>
                <span>{r.date} · {r.category}{r.desc ? ` · ${r.desc}` : ""}{r.account ? ` · ${r.account}` : ""}</span>
                <span style={{ fontFamily: "'IBM Plex Mono', monospace", color: r.type === "ingreso" ? GREEN : r.type === "ahorro" ? GOLD : r.type === "cambio" ? TEAL : BRICK }}>
                  {fmtARS(r.amount)}
                </span>
              </div>
            ))}
          </div>
        </div>
      )}
    </div>
  );
}

// Nuevo apartado en "Mi hogar": de dónde viene cada movimiento cargado
// (manual, voz, WhatsApp, foto, PDF, recurrente, acceso rápido) y el
// detalle de cada carga de archivo puntual (PDF/CSV), con cuántos
// movimientos entraron en cada una.
function ImportacionesCard({ entries }) {
  const [cargas, setCargas] = useState(null); // null = cargando

  useEffect(() => {
    (async () => {
      let rows;
      if (!HAS_SUPABASE) {
        rows = safeGet("mock_audit_log") || [];
      } else {
        try {
          rows = await sb("audit_log?select=*&accion=eq.import&order=created_at.desc&limit=100");
        } catch (e) {
          console.error(e);
          rows = [];
        }
      }
      setCargas((rows || []).filter((r) => r.accion === "import" || !HAS_SUPABASE));
    })();
  }, []);

  const porOrigen = useMemo(() => {
    const map = {};
    (entries || []).forEach((e) => {
      const o = e.origen || "manual";
      map[o] = (map[o] || 0) + 1;
    });
    return Object.entries(map)
      .map(([origen, cantidad]) => ({ origen, cantidad, label: ORIGEN_LABELS[origen] || origen }))
      .sort((a, b) => b.cantidad - a.cantidad);
  }, [entries]);

  return (
    <div style={{ background: "#fff", borderRadius: 10, padding: 18, boxShadow: "0 2px 10px rgba(27,42,46,0.06)" }}>
      <div style={{ fontFamily: "'Fraunces', serif", fontSize: 17, marginBottom: 4 }}>Importaciones de datos</div>
      <div style={{ fontSize: 12.5, color: "#8a9698", marginBottom: 16 }}>
        De dónde vienen tus movimientos cargados, y el detalle de cada carga de archivo (PDF/CSV).
      </div>

      <div style={{ fontSize: 13, fontWeight: 700, marginBottom: 8 }}>Movimientos por origen (todo el historial)</div>
      <div style={{ display: "flex", flexDirection: "column", gap: 2, marginBottom: 20 }}>
        {porOrigen.length === 0 ? (
          <div style={{ fontSize: 12.5, color: "#8a9698" }}>Todavía no hay movimientos cargados.</div>
        ) : porOrigen.map((o) => (
          <div key={o.origen} style={{ display: "flex", justifyContent: "space-between", fontSize: 12.5, padding: "5px 0", borderBottom: "1px solid #f2eee2" }}>
            <span>{o.label}</span>
            <span style={{ fontFamily: "'IBM Plex Mono', monospace", fontWeight: 700 }}>{o.cantidad}</span>
          </div>
        ))}
      </div>

      <div style={{ fontSize: 13, fontWeight: 700, marginBottom: 8 }}>Cargas de archivo (PDF/CSV), una por una</div>
      {cargas === null ? (
        <div style={{ fontSize: 12.5, color: "#8a9698" }}>Cargando...</div>
      ) : cargas.length === 0 ? (
        <div style={{ fontSize: 12.5, color: "#8a9698" }}>Todavía no importaste ningún archivo.</div>
      ) : (
        <div style={{ display: "flex", flexDirection: "column", gap: 8 }}>
          {cargas.map((c) => {
            const snap = c.entry_snapshot || {};
            let fecha = "";
            try { fecha = new Date(c.created_at).toLocaleString("es-AR"); } catch {}
            const archivos = Array.isArray(snap.archivos) ? snap.archivos : [];
            return (
              <div key={c.id} style={{ border: "1px solid #f2eee2", borderRadius: 8, padding: "10px 12px" }}>
                <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 4, flexWrap: "wrap", gap: 4 }}>
                  <span style={{ fontSize: 12.5, fontWeight: 700 }}>{snap.formato || "Importación"}{archivos.length > 1 ? ` (${archivos.length} archivos)` : ""}</span>
                  <span style={{ fontSize: 11, color: "#8a9698" }}>{fecha}{c.who ? ` · ${c.who}` : ""}</span>
                </div>
                {archivos.length > 0 && (
                  <div style={{ fontSize: 11.5, color: TEAL, marginBottom: 4 }}>{archivos.join(", ")}</div>
                )}
                <div style={{ fontSize: 12.5 }}>
                  <b style={{ color: GREEN }}>{snap.cantidad_importada ?? 0}</b> movimientos cargados
                  {snap.cantidad_total != null && Number(snap.cantidad_total) !== Number(snap.cantidad_importada) && (
                    <span style={{ color: "#8a9698" }}> de {snap.cantidad_total} en el archivo</span>
                  )}
                  {Number(snap.cantidad_duplicados) > 0 && (
                    <span style={{ color: GOLD }}> · {snap.cantidad_duplicados} descartados por duplicado</span>
                  )}
                </div>
              </div>
            );
          })}
        </div>
      )}
    </div>
  );
}

function HogarTab({ householdId, onLogout, biometriaSoportada, biometriaActiva, onActivarBiometria, onDesactivarBiometria, bioBusy, bioError, whatsappLinks, profileName, onAddWhatsapp, onDeleteWhatsapp, onRenombrarMiembro, entries }) {
  const [editandoNombreDe, setEditandoNombreDe] = useState(null);
  const [nombreNuevo, setNombreNuevo] = useState("");
  const [renombrando, setRenombrando] = useState(false);
  const [renombrarError, setRenombrarError] = useState(null);
  const [hh, setHh] = useState(null);
  const [miembros, setMiembros] = useState(null);
  const [copiado, setCopiado] = useState(false);
  const [qrDataUrl, setQrDataUrl] = useState(null);
  const [waNumero, setWaNumero] = useState("");
  const [waNombre, setWaNombre] = useState(profileName || "");
  const [waBusy, setWaBusy] = useState(false);
  const [waError, setWaError] = useState(null);

  useEffect(() => {
    (async () => {
      if (!HAS_SUPABASE) return;
      try {
        const [hhRows, memRows] = await Promise.all([
          sb(`households?id=eq.${householdId}&select=*`),
          sb("rpc/get_my_household_members", { method: "POST", body: "{}" }),
        ]);
        setHh(hhRows?.[0] || null);
        setMiembros(memRows || []);
      } catch (e) {
        console.error(e);
      }
    })();
  }, [householdId]);

  const inviteLink = hh?.invite_code && typeof window !== "undefined"
    ? `${window.location.origin}/?invite=${hh.invite_code}`
    : null;

  useEffect(() => {
    if (!inviteLink) return;
    (async () => {
      try {
        // qrcode se carga recién acá (no como import fijo arriba del
        // archivo), para no romper el preview de Claude — mismo patrón
        // que usamos con pdfjs-dist.
        const QRCode = await import("qrcode");
        const url = await QRCode.toDataURL(inviteLink, { width: 240, margin: 1 });
        setQrDataUrl(url);
      } catch (e) {
        console.error("No se pudo generar el QR", e);
      }
    })();
  }, [inviteLink]);

  function copiarCodigo() {
    if (!hh?.invite_code) return;
    navigator.clipboard?.writeText(hh.invite_code);
    setCopiado(true);
    setTimeout(() => setCopiado(false), 2000);
  }

  function empezarRenombrar(nombreActual) {
    setEditandoNombreDe(nombreActual);
    setNombreNuevo(nombreActual);
    setRenombrarError(null);
  }

  async function handleRenombrar(nombreViejo) {
    setRenombrarError(null);
    setRenombrando(true);
    const res = await onRenombrarMiembro(nombreViejo, nombreNuevo);
    setRenombrando(false);
    if (res?.error) { setRenombrarError(res.error); return; }
    if (res?.miembros) setMiembros(res.miembros);
    else setMiembros((prev) => (prev || []).map((m) => (m.display_name === nombreViejo ? { ...m, display_name: nombreNuevo.trim() } : m)));
    setEditandoNombreDe(null);
  }

  async function handleAddWhatsapp() {
    setWaError(null);
    setWaBusy(true);
    const res = await onAddWhatsapp(waNumero, waNombre);
    setWaBusy(false);
    if (res?.error) { setWaError(res.error); return; }
    setWaNumero("");
  }

  return (
    <div style={{ display: "flex", flexDirection: "column", gap: 14 }}>
      <div style={{ background: "#fff", borderRadius: 10, padding: 18, boxShadow: "0 2px 10px rgba(27,42,46,0.06)" }}>
        <div style={{ fontFamily: "'Fraunces', serif", fontSize: 17, marginBottom: 6 }}>Mi hogar</div>
        {!HAS_SUPABASE ? (
          <EmptyState text="No aplica en vista previa local." />
        ) : !hh ? (
          <div style={{ fontSize: 13, color: "#8a9698" }}>Cargando...</div>
        ) : (
          <>
            <div style={{ fontSize: 13, color: "#8a9698", marginBottom: 4 }}>{hh.name}</div>
            <div style={{ fontSize: 12.5, color: "#8a9698", marginBottom: 6 }}>
              Compartí este código con quien quieras sumar a tu hogar (van a ver y editar los mismos datos que vos):
            </div>
            <div style={{ display: "flex", gap: 8, alignItems: "center", marginBottom: 14 }}>
              <div style={{ fontFamily: "'IBM Plex Mono', monospace", fontSize: 20, fontWeight: 700, background: PAPER_DIM, padding: "8px 14px", borderRadius: 8, letterSpacing: 2 }}>
                {hh.invite_code}
              </div>
              <button onClick={copiarCodigo} style={btnOutline}>{copiado ? "¡Copiado!" : "Copiar"}</button>
            </div>

            {inviteLink && (
              <div style={{ textAlign: "center", marginBottom: 16, paddingTop: 4 }}>
                <a
                  href={`https://wa.me/?text=${encodeURIComponent(`Te invito a sumarte a mi hogar en Finanzas del hogar 🏠\nEntrá acá: ${inviteLink}`)}`}
                  target="_blank"
                  rel="noopener noreferrer"
                  style={{ display: "inline-flex", alignItems: "center", gap: 8, padding: "11px 18px", background: "#25D366", color: "#fff", borderRadius: 8, textDecoration: "none", fontWeight: 700, fontSize: 13.5, marginBottom: 16 }}
                >
                  Enviar invitación por WhatsApp
                </a>
                {qrDataUrl && (
                  <div>
                    <img src={qrDataUrl} alt="Código QR de invitación" style={{ width: 180, height: 180 }} />
                    <div style={{ fontSize: 12, color: "#8a9698", marginTop: 4 }}>O escaneá este QR</div>
                  </div>
                )}
              </div>
            )}

            {miembros && miembros.length > 0 && (
              <div>
                <div style={{ fontSize: 11.5, color: "#8a9698", textTransform: "uppercase", letterSpacing: 0.5, marginBottom: 6 }}>Miembros</div>
                <div style={{ fontSize: 11.5, color: "#8a9698", marginBottom: 10 }}>
                  Si algún movimiento viejo quedó guardado con un nombre distinto (ej. "Nati" en vez de "Natalia"), corregilo acá — el cambio se aplica también a todos los movimientos ya cargados con el nombre anterior.
                </div>
                <div style={{ display: "flex", flexDirection: "column", gap: 8 }}>
                  {miembros.map((m, i) => {
                    const enEdicion = editandoNombreDe === m.display_name;
                    return (
                      <div key={i}>
                        {enEdicion ? (
                          <div style={{ display: "flex", alignItems: "center", gap: 6 }}>
                            <input
                              value={nombreNuevo}
                              onChange={(e) => setNombreNuevo(e.target.value)}
                              autoFocus
                              style={{ ...inputStyle, flex: 1, padding: "6px 8px", fontSize: 13 }}
                              onKeyDown={(e) => { if (e.key === "Enter") handleRenombrar(m.display_name); if (e.key === "Escape") setEditandoNombreDe(null); }}
                            />
                            <button onClick={() => handleRenombrar(m.display_name)} disabled={renombrando} aria-label="Confirmar" style={{ background: "none", border: "none", cursor: "pointer", color: GREEN }}>
                              <Check size={17} />
                            </button>
                            <button onClick={() => setEditandoNombreDe(null)} aria-label="Cancelar" style={{ background: "none", border: "none", cursor: "pointer", color: "#8a9698" }}>
                              <X size={17} />
                            </button>
                          </div>
                        ) : (
                          <div style={{ display: "flex", alignItems: "center", gap: 6, fontSize: 13 }}>
                            <span style={{ flex: 1 }}>{m.display_name} {m.role === "owner" ? "· dueño/a" : ""}</span>
                            <button onClick={() => empezarRenombrar(m.display_name)} aria-label="Editar nombre" style={{ background: "none", border: "none", cursor: "pointer", color: TEAL }}>
                              <Pencil size={14} />
                            </button>
                          </div>
                        )}
                        {enEdicion && renombrarError && <div style={{ color: BRICK, fontSize: 11.5, marginTop: 4 }}>{renombrarError}</div>}
                      </div>
                    );
                  })}
                </div>
              </div>
            )}
          </>
        )}
      </div>

      {HAS_SUPABASE && (
        <div style={{ background: "#fff", borderRadius: 10, padding: 18, boxShadow: "0 2px 10px rgba(27,42,46,0.06)" }}>
          <div style={{ fontFamily: "'Fraunces', serif", fontSize: 17, marginBottom: 6 }}>Seguridad</div>
          {!biometriaSoportada ? (
            <div style={{ fontSize: 12.5, color: "#8a9698" }}>
              Este dispositivo/navegador no tiene Face ID, Touch ID o huella disponible para la app.
            </div>
          ) : biometriaActiva ? (
            <>
              <div style={{ fontSize: 12.5, color: "#8a9698", marginBottom: 12 }}>
                Activado: cada vez que abras la app en este dispositivo (con la sesión ya guardada) te vamos a pedir Face ID / Touch ID / huella antes de mostrar tus datos.
              </div>
              <button onClick={onDesactivarBiometria} style={{ ...btnOutline, justifyContent: "center", width: "100%" }}>
                Desactivar desbloqueo biométrico
              </button>
            </>
          ) : (
            <>
              <div style={{ fontSize: 12.5, color: "#8a9698", marginBottom: 12 }}>
                Podés activar que la app te pida Face ID / Touch ID / huella cada vez que la abras en este dispositivo, en vez de mostrar los datos directo. Te va a pedir confirmarlo ahora una vez con tu biometría para guardarlo.
              </div>
              {bioError && <div style={{ color: BRICK, fontSize: 12, marginBottom: 10 }}>{bioError}</div>}
              <button onClick={onActivarBiometria} disabled={bioBusy} style={{ ...btnPrimary, justifyContent: "center", width: "100%" }}>
                {bioBusy ? "Confirmando..." : "Activar desbloqueo biométrico"}
              </button>
            </>
          )}
        </div>
      )}

      {HAS_SUPABASE && (
        <div style={{ background: "#fff", borderRadius: 10, padding: 18, boxShadow: "0 2px 10px rgba(27,42,46,0.06)" }}>
          <div style={{ fontFamily: "'Fraunces', serif", fontSize: 17, marginBottom: 6 }}>WhatsApp</div>
          <div style={{ fontSize: 12.5, color: "#8a9698", marginBottom: 14 }}>
            Vinculá tu número para cargar gastos/ingresos escribiéndole o mandándole un audio al bot, y para preguntarle cosas como "cuánto gasté en comida este mes?".
          </div>

          {whatsappLinks && whatsappLinks.length > 0 && (
            <div style={{ display: "flex", flexDirection: "column", gap: 6, marginBottom: 14 }}>
              {whatsappLinks.map((w) => (
                <div key={w.id} style={{ display: "flex", justifyContent: "space-between", alignItems: "center", background: PAPER_DIM, borderRadius: 8, padding: "8px 12px" }}>
                  <span style={{ fontSize: 13 }}>
                    <b>{w.display_name}</b> · +{w.phone_number}
                  </span>
                  <button onClick={() => onDeleteWhatsapp(w.id)} style={{ background: "none", border: "none", cursor: "pointer", color: BRICK }}>
                    <Trash2 size={15} />
                  </button>
                </div>
              ))}
            </div>
          )}

          <label style={labelStyle}>Tu número (con código de país, sin +)</label>
          <input value={waNumero} onChange={(e) => setWaNumero(e.target.value)} placeholder="Ej: 5491122334455" style={{ ...inputStyle, marginBottom: 10 }} />
          <label style={labelStyle}>Nombre (para saber que sos vos)</label>
          <input value={waNombre} onChange={(e) => setWaNombre(e.target.value)} placeholder="Ej: Hernán" style={{ ...inputStyle, marginBottom: 10 }} />
          {waError && <div style={{ color: BRICK, fontSize: 12, marginBottom: 10 }}>{waError}</div>}
          <button onClick={handleAddWhatsapp} disabled={waBusy || !waNumero.trim()} style={{ ...btnPrimary, justifyContent: "center", width: "100%" }}>
            {waBusy ? "Vinculando..." : "Vincular número"}
          </button>
        </div>
      )}

      <ImportacionesCard entries={entries} />

      <button onClick={onLogout} style={{ ...btnOutline, color: BRICK, borderColor: BRICK, justifyContent: "center" }}>
        Cerrar sesión
      </button>
    </div>
  );
}

function CategoriasTab({ categories, contarUsos, onAdd, onRename, onDelete }) {
  const [nuevoNombre, setNuevoNombre] = useState("");
  const [addError, setAddError] = useState(null);
  const [editKey, setEditKey] = useState(null);
  const [editValue, setEditValue] = useState("");
  const [busy, setBusy] = useState(null);
  const [deleteError, setDeleteError] = useState({});

  async function handleAdd() {
    setAddError(null);
    const res = await onAdd(nuevoNombre);
    if (!res.ok) { setAddError(res.error); return; }
    setNuevoNombre("");
  }

  function startEdit(cat) {
    setEditKey(cat);
    setEditValue(cat);
  }
  function cancelEdit() {
    setEditKey(null);
    setEditValue("");
  }
  async function saveRename(viejo) {
    setBusy(viejo);
    const res = await onRename(viejo, editValue);
    setBusy(null);
    if (res.ok) { setEditKey(null); setEditValue(""); }
    else setAddError(res.error);
  }
  async function handleDelete(cat) {
    setBusy(cat);
    setDeleteError((prev) => ({ ...prev, [cat]: null }));
    const res = await onDelete(cat);
    setBusy(null);
    if (!res.ok) setDeleteError((prev) => ({ ...prev, [cat]: res.error }));
  }

  return (
    <div style={{ background: "#fff", borderRadius: 10, padding: 18, boxShadow: "0 2px 10px rgba(27,42,46,0.06)" }}>
      <div style={{ fontFamily: "'Fraunces', serif", fontSize: 17, marginBottom: 6 }}>Categorías</div>
      <div style={{ fontSize: 13, color: "#8a9698", marginBottom: 14 }}>
        Agregá, renombrá o borrá categorías. Para borrar una que ya tiene gastos cargados, primero hay que recategorizar todo lo que la usa (pestaña Recategorizar).
      </div>

      <div style={{ display: "flex", gap: 8, marginBottom: 8 }}>
        <input
          value={nuevoNombre}
          onChange={(e) => setNuevoNombre(e.target.value)}
          onKeyDown={(e) => e.key === "Enter" && handleAdd()}
          placeholder="Nueva categoría, ej: Mascotas"
          style={{ ...inputStyle, flex: 1 }}
        />
        <button onClick={handleAdd} style={btnPrimary}><Plus size={16} /> Agregar</button>
      </div>
      {addError && <div style={{ color: BRICK, fontSize: 12, marginBottom: 14 }}>{addError}</div>}

      <div style={{ display: "flex", flexDirection: "column", gap: 6, marginTop: 10 }}>
        {categories.map((cat) => {
          const usos = contarUsos(cat);
          const editing = editKey === cat;
          return (
            <div key={cat} style={{ background: PAPER_DIM, borderRadius: 8, padding: "8px 10px" }}>
              {editing ? (
                <div style={{ display: "flex", gap: 6, alignItems: "center" }}>
                  <input
                    autoFocus
                    value={editValue}
                    onChange={(e) => setEditValue(e.target.value)}
                    onKeyDown={(e) => { if (e.key === "Enter") saveRename(cat); if (e.key === "Escape") cancelEdit(); }}
                    style={{ ...inputStyle, padding: "5px 8px", fontSize: 13, flex: 1 }}
                  />
                  <button onClick={() => saveRename(cat)} disabled={busy === cat} style={{ border: "none", background: "none", cursor: "pointer", color: GREEN }}><Check size={16} /></button>
                  <button onClick={cancelEdit} style={{ border: "none", background: "none", cursor: "pointer", color: "#b8b2a4" }}><X size={16} /></button>
                </div>
              ) : (
                <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", gap: 8 }}>
                  <div style={{ fontSize: 13.5, fontWeight: 600 }}>
                    {cat} <span style={{ fontSize: 11, color: "#8a9698", fontWeight: 400 }}>({usos} movimiento{usos !== 1 ? "s" : ""})</span>
                  </div>
                  <div style={{ display: "flex", gap: 10, flexShrink: 0 }}>
                    <button onClick={() => startEdit(cat)} style={{ border: "none", background: "none", cursor: "pointer", color: TEAL }} aria-label="Renombrar"><Pencil size={14} /></button>
                    <button onClick={() => handleDelete(cat)} disabled={busy === cat} style={{ border: "none", background: "none", cursor: "pointer", color: BRICK }} aria-label="Borrar"><Trash2 size={14} /></button>
                  </div>
                </div>
              )}
              {deleteError[cat] && <div style={{ color: BRICK, fontSize: 11.5, marginTop: 6 }}>{deleteError[cat]}</div>}
            </div>
          );
        })}
      </div>
    </div>
  );
}

function ConciliarPagosTab({ entries, onConfirmar }) {
  const [descartadas, setDescartadas] = useState([]); // ids de pares "pendiente-pago" descartados en esta sesión
  const [confirmando, setConfirmando] = useState(null);

  const sugerencias = useMemo(() => {
    const pendientes = entries.filter((e) => e.type === "gasto" && e.pagado === false);
    const pagados = entries.filter((e) => e.type === "gasto" && e.pagado !== false);
    const resultado = [];
    pendientes.forEach((p) => {
      pagados.forEach((pg) => {
        if (pg.id === p.id) return;
        const descP = (p.desc || p.category || "").toLowerCase().trim();
        const descPg = (pg.desc || pg.category || "").toLowerCase().trim();
        if (!descP || !descPg) return;

        // Señal fuerte: comparten un número largo (factura/comprobante/
        // "Doc.") — casi siempre significa que son el mismo movimiento,
        // aunque el monto no coincida perfecto (redondeos, intereses).
        const numerosP = descP.match(/\d{5,}/g) || [];
        const numerosPg = descPg.match(/\d{5,}/g) || [];
        const numeroComun = numerosP.find((n) => numerosPg.includes(n));

        const montoCoincide = Math.abs(Number(p.amount) - Number(pg.amount)) <= 0.5;
        const textoCoincide =
          descP.includes(descPg) || descPg.includes(descP) ||
          descP.split(/\s+/).some((w) => w.length > 3 && descPg.includes(w));

        if (!numeroComun && !(montoCoincide && textoCoincide)) return;

        const clave = `${p.id}-${pg.id}`;
        if (descartadas.includes(clave)) return;
        resultado.push({ clave, pendiente: p, pago: pg, numeroComun: numeroComun || null, montoCoincide });
      });
    });
    // Las coincidencias por número de comprobante van primero — son las más confiables.
    return resultado.sort((a, b) => (b.numeroComun ? 1 : 0) - (a.numeroComun ? 1 : 0));
  }, [entries, descartadas]);

  async function handleConfirmar(clave, pendienteId, pagoId) {
    setConfirmando(clave);
    await onConfirmar(pendienteId, pagoId);
    setConfirmando(null);
  }

  return (
    <div style={{ display: "flex", flexDirection: "column", gap: 14 }}>
      <div style={{ fontSize: 13, color: "#8a9698" }}>
        Compara automáticamente tus gastos marcados como "Pendiente" contra el resto de tus movimientos, buscando el mismo monto y una descripción parecida — o, más confiable todavía, el mismo número de factura/comprobante compartido. Al confirmar, marca la factura como pagada y borra el duplicado del resumen, para no contar el gasto dos veces.
      </div>

      {sugerencias.length === 0 ? (
        <EmptyState text="No hay coincidencias para revisar por ahora." />
      ) : (
        <div style={{ display: "flex", flexDirection: "column", gap: 10 }}>
          {sugerencias.map(({ clave, pendiente, pago, numeroComun, montoCoincide }) => (
            <div key={clave} style={{ background: "#fff", borderRadius: 10, padding: 14, boxShadow: "0 2px 10px rgba(27,42,46,0.06)" }}>
              {numeroComun ? (
                <div style={{ display: "inline-flex", alignItems: "center", gap: 5, fontSize: 10.5, fontWeight: 700, color: GREEN, background: "#e8f3ec", padding: "3px 8px", borderRadius: 10, marginBottom: 8 }}>
                  🔗 Mismo N° de comprobante ({numeroComun}){!montoCoincide ? " — monto distinto, revisá antes de confirmar" : ""}
                </div>
              ) : (
                <div style={{ fontSize: 10.5, color: "#8a9698", marginBottom: 8 }}>Coincidencia por monto y descripción parecida</div>
              )}
              <div style={{ display: "flex", flexDirection: "column", gap: 8, marginBottom: 12 }}>
                <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", padding: "8px 10px", background: "#fdf1de", borderRadius: 8 }}>
                  <div>
                    <div style={{ fontSize: 10, fontWeight: 700, color: GOLD, marginBottom: 2 }}>PENDIENTE</div>
                    <div style={{ fontSize: 13, fontWeight: 600 }}>{pendiente.desc || pendiente.category}</div>
                    <div style={{ fontSize: 11, color: "#8a9698" }}>{pendiente.date} · {pendiente.category}</div>
                  </div>
                  <div style={{ fontFamily: "'IBM Plex Mono', monospace", fontWeight: 700 }}>{fmtARS(pendiente.amount)}</div>
                </div>
                <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", padding: "8px 10px", background: "#e8f0f0", borderRadius: 8 }}>
                  <div>
                    <div style={{ fontSize: 10, fontWeight: 700, color: TEAL, marginBottom: 2 }}>MOVIMIENTO ENCONTRADO</div>
                    <div style={{ fontSize: 13, fontWeight: 600 }}>{pago.desc || pago.category}</div>
                    <div style={{ fontSize: 11, color: "#8a9698" }}>{pago.date} · {pago.account || "sin cuenta"}</div>
                  </div>
                  <div style={{ fontFamily: "'IBM Plex Mono', monospace", fontWeight: 700 }}>{fmtARS(pago.amount)}</div>
                </div>
              </div>
              <div style={{ display: "flex", gap: 8 }}>
                <button
                  onClick={() => setDescartadas((prev) => [...prev, clave])}
                  style={{ ...btnOutline, flex: 1, justifyContent: "center", fontSize: 12.5 }}
                >
                  No es la misma
                </button>
                <button
                  onClick={() => handleConfirmar(clave, pendiente.id, pago.id)}
                  disabled={confirmando === clave}
                  style={{ ...btnPrimary, flex: 1, justifyContent: "center", fontSize: 12.5 }}
                >
                  {confirmando === clave ? "Confirmando..." : "Sí, es el mismo pago"}
                </button>
              </div>
            </div>
          ))}
        </div>
      )}
    </div>
  );
}

const CRIPTOS_DISPONIBLES = [
  { id: "bitcoin", symbol: "BTC", nombre: "Bitcoin" },
  { id: "ethereum", symbol: "ETH", nombre: "Ethereum" },
  { id: "tether", symbol: "USDT", nombre: "Tether" },
  { id: "usd-coin", symbol: "USDC", nombre: "USD Coin" },
  { id: "binancecoin", symbol: "BNB", nombre: "BNB" },
  { id: "solana", symbol: "SOL", nombre: "Solana" },
  { id: "ripple", symbol: "XRP", nombre: "XRP" },
  { id: "cardano", symbol: "ADA", nombre: "Cardano" },
  { id: "dogecoin", symbol: "DOGE", nombre: "Dogecoin" },
  { id: "polkadot", symbol: "DOT", nombre: "Polkadot" },
  { id: "litecoin", symbol: "LTC", nombre: "Litecoin" },
  { id: "avalanche-2", symbol: "AVAX", nombre: "Avalanche" },
  { id: "chainlink", symbol: "LINK", nombre: "Chainlink" },
  { id: "matic-network", symbol: "MATIC", nombre: "Polygon" },
  { id: "tron", symbol: "TRX", nombre: "TRON" },
];

const DOLAR_LABELS = {
  oficial: "Oficial",
  blue: "Blue",
  bolsa: "MEP",
  contadoconliqui: "CCL",
  mayorista: "Mayorista",
  tarjeta: "Tarjeta",
  cripto: "Cripto (USDT)",
};

function CotizacionesTab() {
  const [dolares, setDolares] = useState(null);
  const [cripto, setCripto] = useState(null);
  const [cargando, setCargando] = useState(true);
  const [error, setError] = useState(null);
  const [ultimaActualizacion, setUltimaActualizacion] = useState(null);
  const [trackeadas, setTrackeadas] = useState(() => safeGet("cripto_trackeadas") || ["bitcoin", "ethereum", "tether"]);
  const [mostrarAgregar, setMostrarAgregar] = useState(false);

  useEffect(() => { safeSet("cripto_trackeadas", trackeadas); }, [trackeadas]);

  async function cargarTodo() {
    setCargando(true);
    setError(null);
    try {
      const [dolaresRes, criptoRes] = await Promise.all([
        fetch("https://dolarapi.com/v1/dolares").then((r) => r.json()),
        trackeadas.length > 0
          ? fetch(`https://api.coingecko.com/api/v3/simple/price?ids=${trackeadas.join(",")}&vs_currencies=usd&include_24hr_change=true`).then((r) => r.json())
          : Promise.resolve({}),
      ]);
      setDolares(Array.isArray(dolaresRes) ? dolaresRes : []);
      setCripto(criptoRes || {});
      setUltimaActualizacion(new Date());
    } catch (e) {
      setError("No se pudieron traer las cotizaciones. Probá de nuevo en un rato.");
    }
    setCargando(false);
  }

  useEffect(() => { cargarTodo(); }, [trackeadas]); // eslint-disable-line react-hooks/exhaustive-deps

  function toggleCripto(id) {
    setTrackeadas((prev) => (prev.includes(id) ? prev.filter((c) => c !== id) : [...prev, id]));
  }

  return (
    <div style={{ display: "flex", flexDirection: "column", gap: 14 }}>
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center" }}>
        <div style={{ fontSize: 12.5, color: "#8a9698" }}>
          {ultimaActualizacion ? `Actualizado ${ultimaActualizacion.toLocaleTimeString("es-AR", { hour: "2-digit", minute: "2-digit" })}` : "Cotizaciones en vivo"}
        </div>
        <button onClick={cargarTodo} disabled={cargando} style={{ ...btnOutline, padding: "6px 10px", fontSize: 12 }}>
          {cargando ? "Actualizando..." : "Actualizar"}
        </button>
      </div>

      {error && <div style={{ color: BRICK, fontSize: 12.5 }}>{error}</div>}

      <div style={{ background: "#fff", borderRadius: 10, padding: 18, boxShadow: "0 2px 10px rgba(27,42,46,0.06)" }}>
        <div style={{ fontFamily: "'Fraunces', serif", fontSize: 17, marginBottom: 12 }}>Dólar</div>
        {!dolares ? (
          <div style={{ fontSize: 13, color: "#8a9698" }}>Cargando...</div>
        ) : dolares.length === 0 ? (
          <EmptyState text="No se pudo traer la cotización del dólar." />
        ) : (
          <div style={{ display: "flex", flexDirection: "column", gap: 2 }}>
            {dolares.map((d) => (
              <div key={d.casa} style={{ display: "flex", justifyContent: "space-between", alignItems: "center", padding: "8px 0", borderBottom: "1px solid #f2eee2" }}>
                <span style={{ fontSize: 13.5, fontWeight: 600 }}>{DOLAR_LABELS[d.casa] || d.nombre}</span>
                <div style={{ textAlign: "right" }}>
                  <div style={{ fontFamily: "'IBM Plex Mono', monospace", fontSize: 13.5, fontWeight: 700 }}>{fmtARS(d.venta)}</div>
                  <div style={{ fontSize: 10.5, color: "#8a9698" }}>compra {fmtARS(d.compra)}</div>
                </div>
              </div>
            ))}
          </div>
        )}
      </div>

      <div style={{ background: "#fff", borderRadius: 10, padding: 18, boxShadow: "0 2px 10px rgba(27,42,46,0.06)" }}>
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 12 }}>
          <div style={{ fontFamily: "'Fraunces', serif", fontSize: 17 }}>Crypto</div>
          <button onClick={() => setMostrarAgregar((v) => !v)} style={{ ...btnOutline, padding: "5px 10px", fontSize: 11.5 }}>
            {mostrarAgregar ? "Listo" : "+ Agregar"}
          </button>
        </div>

        {mostrarAgregar && (
          <div style={{ display: "flex", flexWrap: "wrap", gap: 6, marginBottom: 14 }}>
            {CRIPTOS_DISPONIBLES.map((c) => {
              const active = trackeadas.includes(c.id);
              return (
                <button key={c.id} onClick={() => toggleCripto(c.id)} style={{
                  padding: "5px 10px", borderRadius: 16, fontSize: 11.5, fontWeight: 700, cursor: "pointer",
                  border: `1px solid ${active ? TEAL : "#ddd6c4"}`,
                  background: active ? TEAL : "#fff", color: active ? "#fff" : INK,
                }}>{c.symbol}</button>
              );
            })}
          </div>
        )}

        {trackeadas.length === 0 ? (
          <EmptyState text="No estás siguiendo ninguna cripto. Tocá + Agregar para elegir." />
        ) : !cripto ? (
          <div style={{ fontSize: 13, color: "#8a9698" }}>Cargando...</div>
        ) : (
          <div style={{ display: "flex", flexDirection: "column", gap: 2 }}>
            {trackeadas.map((id) => {
              const info = CRIPTOS_DISPONIBLES.find((c) => c.id === id);
              const precio = cripto[id]?.usd;
              const variacion = cripto[id]?.usd_24h_change;
              return (
                <div key={id} style={{ display: "flex", justifyContent: "space-between", alignItems: "center", padding: "8px 0", borderBottom: "1px solid #f2eee2" }}>
                  <div>
                    <div style={{ fontSize: 13.5, fontWeight: 600 }}>{info?.nombre || id}</div>
                    <div style={{ fontSize: 10.5, color: "#8a9698" }}>{info?.symbol}</div>
                  </div>
                  <div style={{ textAlign: "right" }}>
                    <div style={{ fontFamily: "'IBM Plex Mono', monospace", fontSize: 13.5, fontWeight: 700 }}>
                      {precio != null ? `US$ ${precio.toLocaleString("es-AR", { maximumFractionDigits: precio < 1 ? 4 : 2 })}` : "—"}
                    </div>
                    {variacion != null && (
                      <div style={{ fontSize: 10.5, fontWeight: 700, color: variacion >= 0 ? GREEN : BRICK }}>
                        {variacion >= 0 ? "▲" : "▼"} {Math.abs(variacion).toFixed(1)}% (24h)
                      </div>
                    )}
                  </div>
                </div>
              );
            })}
          </div>
        )}
      </div>

      <div style={{ fontSize: 10.5, color: "#c4bda8", textAlign: "center" }}>
        Fuente: DolarAPI.com y CoinGecko. Puede haber unos minutos de demora respecto al mercado real.
      </div>
    </div>
  );
}

function CuentasTab({ thisMonthEntries, monthLabel, onRenombrarCuenta }) {
  const [expandedMedio, setExpandedMedio] = useState(null);
  const [expandedCuenta, setExpandedCuenta] = useState(null);
  const [editandoCuentaDe, setEditandoCuentaDe] = useState(null);
  const [nombreCuentaNuevo, setNombreCuentaNuevo] = useState("");
  const [renombrandoCuenta, setRenombrandoCuenta] = useState(false);
  const [errorRenombrarCuenta, setErrorRenombrarCuenta] = useState(null);

  async function handleRenombrarCuenta(nombreViejo) {
    setErrorRenombrarCuenta(null);
    setRenombrandoCuenta(true);
    const res = await onRenombrarCuenta(nombreViejo, nombreCuentaNuevo);
    setRenombrandoCuenta(false);
    if (res?.error) { setErrorRenombrarCuenta(res.error); return; }
    setEditandoCuentaDe(null);
  }

  const gastosArs = useMemo(
    () => thisMonthEntries.filter((e) => e.type === "gasto" && (e.moneda || "ARS") === "ARS"),
    [thisMonthEntries]
  );

  // Nivel 1: por medio de pago macro (Efectivo / Mercado Pago / Tarjetas de crédito / Otros)
  const gastosPorMedio = useMemo(() => {
    const map = {};
    gastosArs.forEach((e) => {
      const medio = clasificarMedioPago(e.account);
      map[medio] = (map[medio] || 0) + Number(e.amount);
    });
    return Object.entries(map).map(([name, value]) => ({ name, value })).sort((a, b) => b.value - a.value);
  }, [gastosArs]);

  // Nivel 2: dentro de un medio, por cuenta/tarjeta real
  function cuentasDentroDe(medio) {
    const map = {};
    gastosArs.forEach((e) => {
      if (clasificarMedioPago(e.account) !== medio) return;
      const cuenta = (e.account || "").trim() || "Sin cuenta especificada";
      map[cuenta] = (map[cuenta] || 0) + Number(e.amount);
    });
    return Object.entries(map).map(([name, value]) => ({ name, value })).sort((a, b) => b.value - a.value);
  }

  const ingresosPorCuenta = useMemo(() => {
    const map = {};
    thisMonthEntries
      .filter((e) => e.type === "ingreso" || e.type === "cambio")
      .forEach((e) => {
        const cuenta = (e.account || "").trim() || "Sin cuenta especificada";
        map[cuenta] = (map[cuenta] || 0) + Number(e.amount);
      });
    return Object.entries(map).map(([name, value]) => ({ name, value })).sort((a, b) => b.value - a.value);
  }, [thisMonthEntries]);

  return (
    <div style={{ display: "flex", flexDirection: "column", gap: 14 }}>
      <div style={{ fontSize: 12.5, color: "#8a9698" }}>
        Gastos de {monthLabel}, agrupados por medio de pago — Efectivo, Mercado Pago, Tarjetas de crédito (todas juntas) y Otros. Tocá cualquiera para ver el detalle por tarjeta/cuenta real, y de ahí los movimientos. Si dos nombres son en realidad la misma tarjeta, tocá el lápiz para unificarlos.
      </div>

      <div style={{ background: "#fff", borderRadius: 10, padding: 18, boxShadow: "0 2px 10px rgba(27,42,46,0.06)" }}>
        <div style={{ fontFamily: "'Fraunces', serif", fontSize: 17, marginBottom: 12 }}>Gastos por medio de pago</div>
        {gastosPorMedio.length === 0 ? <EmptyState text="Todavía no cargaste gastos este mes." /> : (
          <div style={{ display: "flex", alignItems: "center", gap: 12, flexWrap: "wrap" }}>
            <div style={{ width: 180, height: 180 }}>
              <ResponsiveContainer>
                <PieChart>
                  <Pie data={gastosPorMedio} dataKey="value" nameKey="name" innerRadius={45} outerRadius={80} paddingAngle={2}>
                    {gastosPorMedio.map((_, i) => <Cell key={i} fill={CAT_COLORS[i % CAT_COLORS.length]} />)}
                  </Pie>
                  <Tooltip formatter={(v) => fmtARS(v)} />
                </PieChart>
              </ResponsiveContainer>
            </div>
            <div style={{ flex: 1, minWidth: 200 }}>
              {gastosPorMedio.map((medio, i) => {
                const medioAbierto = expandedMedio === medio.name;
                const cuentas = medioAbierto ? cuentasDentroDe(medio.name) : [];
                return (
                  <div key={medio.name}>
                    <div
                      onClick={() => { setExpandedMedio(medioAbierto ? null : medio.name); setExpandedCuenta(null); }}
                      style={{ display: "flex", alignItems: "center", gap: 8, padding: "8px 0", cursor: "pointer" }}
                    >
                      <div style={{ width: 10, height: 10, borderRadius: "50%", background: CAT_COLORS[i % CAT_COLORS.length], flexShrink: 0 }} />
                      <div style={{ flex: 1, fontSize: 14, fontWeight: 700 }}>{medio.name}</div>
                      <div style={{ fontFamily: "'IBM Plex Mono', monospace", fontSize: 14, fontWeight: 700 }}>{fmtARS(medio.value)}</div>
                    </div>
                    {medioAbierto && (
                      <div style={{ marginLeft: 18, marginBottom: 8, display: "flex", flexDirection: "column", gap: 2 }}>
                        {cuentas.map((c) => {
                          const cuentaAbierta = expandedCuenta === c.name;
                          const movs = cuentaAbierta
                            ? gastosArs.filter((e) => ((e.account || "").trim() || "Sin cuenta especificada") === c.name)
                                .sort((a, b) => (b.date || "").localeCompare(a.date || ""))
                            : [];
                          return (
                            <div key={c.name}>
                              {editandoCuentaDe === c.name ? (
                                <div style={{ padding: "6px 0" }}>
                                  <div style={{ display: "flex", alignItems: "center", gap: 6 }}>
                                    <input
                                      value={nombreCuentaNuevo}
                                      onChange={(e) => setNombreCuentaNuevo(e.target.value)}
                                      autoFocus
                                      style={{ ...inputStyle, flex: 1, padding: "6px 8px", fontSize: 13 }}
                                      onKeyDown={(e) => { if (e.key === "Enter") handleRenombrarCuenta(c.name); if (e.key === "Escape") setEditandoCuentaDe(null); }}
                                    />
                                    <button onClick={() => handleRenombrarCuenta(c.name)} disabled={renombrandoCuenta} aria-label="Confirmar" style={{ background: "none", border: "none", cursor: "pointer", color: GREEN }}>
                                      <Check size={17} />
                                    </button>
                                    <button onClick={() => setEditandoCuentaDe(null)} aria-label="Cancelar" style={{ background: "none", border: "none", cursor: "pointer", color: "#8a9698" }}>
                                      <X size={17} />
                                    </button>
                                  </div>
                                  <div style={{ fontSize: 11, color: "#8a9698", marginTop: 4 }}>
                                    Va a renombrar TODOS los movimientos viejos de "{c.name}" (de cualquier mes) — útil si en realidad es la misma tarjeta que otra cuenta.
                                  </div>
                                  {errorRenombrarCuenta && <div style={{ color: BRICK, fontSize: 11.5, marginTop: 4 }}>{errorRenombrarCuenta}</div>}
                                </div>
                              ) : (
                                <div style={{ display: "flex", alignItems: "center", gap: 8, padding: "6px 0" }}>
                                  <div onClick={() => setExpandedCuenta(cuentaAbierta ? null : c.name)} style={{ display: "flex", alignItems: "center", gap: 8, flex: 1, cursor: "pointer", minWidth: 0 }}>
                                    <div style={{ flex: 1, fontSize: 13, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>{c.name}</div>
                                    <div style={{ fontFamily: "'IBM Plex Mono', monospace", fontSize: 13, fontWeight: 600 }}>{fmtARS(c.value)}</div>
                                  </div>
                                  <button
                                    onClick={() => { setEditandoCuentaDe(c.name); setNombreCuentaNuevo(c.name); setErrorRenombrarCuenta(null); }}
                                    aria-label="Renombrar/unificar cuenta"
                                    style={{ background: "none", border: "none", cursor: "pointer", color: "#c4bda8", flexShrink: 0, padding: 2 }}
                                  >
                                    <Pencil size={13} />
                                  </button>
                                </div>
                              )}
                              {cuentaAbierta && (
                                <div style={{ marginLeft: 14, marginBottom: 6, display: "flex", flexDirection: "column", gap: 3 }}>
                                  {movs.map((m) => (
                                    <div key={m.id} style={{ display: "flex", justifyContent: "space-between", fontSize: 11.5, color: "#5a6b6d" }}>
                                      <span style={{ overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap", maxWidth: 160 }}>{m.desc || m.category}</span>
                                      <span style={{ fontFamily: "'IBM Plex Mono', monospace" }}>{fmtARS(m.amount)}</span>
                                    </div>
                                  ))}
                                </div>
                              )}
                            </div>
                          );
                        })}
                      </div>
                    )}
                  </div>
                );
              })}
            </div>
          </div>
        )}
      </div>

      {ingresosPorCuenta.length > 0 && (
        <div style={{ background: "#fff", borderRadius: 10, padding: 18, boxShadow: "0 2px 10px rgba(27,42,46,0.06)" }}>
          <div style={{ fontFamily: "'Fraunces', serif", fontSize: 17, marginBottom: 12 }}>Ingresos por cuenta</div>
          <div style={{ display: "flex", flexDirection: "column", gap: 4 }}>
            {ingresosPorCuenta.map((c) => (
              <div key={c.name} style={{ display: "flex", justifyContent: "space-between", fontSize: 13, padding: "4px 0" }}>
                <span>{c.name}</span>
                <span style={{ fontFamily: "'IBM Plex Mono', monospace", color: GREEN, fontWeight: 600 }}>{fmtARS(c.value)}</span>
              </div>
            ))}
          </div>
        </div>
      )}
    </div>
  );
}

function PersonasTab({ thisMonthEntries, monthLabel }) {
  const [expandedPersona, setExpandedPersona] = useState(null);

  const gastosPorPersona = useMemo(() => {
    const map = {};
    thisMonthEntries
      .filter((e) => e.type === "gasto" && (e.moneda || "ARS") === "ARS")
      .forEach((e) => {
        const persona = capitalizar(e.persona || e.who) || "Sin nombre";
        map[persona] = (map[persona] || 0) + Number(e.amount);
      });
    return Object.entries(map).map(([name, value]) => ({ name, value })).sort((a, b) => b.value - a.value);
  }, [thisMonthEntries]);

  const ingresosPorPersona = useMemo(() => {
    const map = {};
    thisMonthEntries
      .filter((e) => e.type === "ingreso" || e.type === "cambio")
      .forEach((e) => {
        const persona = capitalizar(e.persona || e.who) || "Sin nombre";
        map[persona] = (map[persona] || 0) + Number(e.amount);
      });
    return Object.entries(map).map(([name, value]) => ({ name, value })).sort((a, b) => b.value - a.value);
  }, [thisMonthEntries]);

  return (
    <div style={{ display: "flex", flexDirection: "column", gap: 14 }}>
      <div style={{ fontSize: 12.5, color: "#8a9698" }}>
        Gastos e ingresos de {monthLabel}, agrupados por quién lo gastó realmente (no por quién lo cargó en la app) — se elige al cargar cada movimiento.
      </div>

      <div style={{ background: "#fff", borderRadius: 10, padding: 18, boxShadow: "0 2px 10px rgba(27,42,46,0.06)" }}>
        <div style={{ fontFamily: "'Fraunces', serif", fontSize: 17, marginBottom: 12 }}>Gastos por persona</div>
        {gastosPorPersona.length === 0 ? <EmptyState text="Todavía no cargaste gastos este mes." /> : (
          <div style={{ display: "flex", alignItems: "center", gap: 12, flexWrap: "wrap" }}>
            <div style={{ width: 180, height: 180 }}>
              <ResponsiveContainer>
                <PieChart>
                  <Pie data={gastosPorPersona} dataKey="value" nameKey="name" innerRadius={45} outerRadius={80} paddingAngle={2}>
                    {gastosPorPersona.map((_, i) => <Cell key={i} fill={CAT_COLORS[i % CAT_COLORS.length]} />)}
                  </Pie>
                  <Tooltip formatter={(v) => fmtARS(v)} />
                </PieChart>
              </ResponsiveContainer>
            </div>
            <div style={{ flex: 1, minWidth: 160 }}>
              {gastosPorPersona.map((p, i) => {
                const isOpen = expandedPersona === `gasto-${p.name}`;
                const movs = isOpen
                  ? thisMonthEntries
                      .filter((e) => e.type === "gasto" && (e.moneda || "ARS") === "ARS" && (capitalizar(e.persona || e.who) || "Sin nombre") === p.name)
                      .sort((a, b) => (b.date || "").localeCompare(a.date || ""))
                  : [];
                return (
                  <div key={p.name}>
                    <div
                      onClick={() => setExpandedPersona(isOpen ? null : `gasto-${p.name}`)}
                      style={{ display: "flex", alignItems: "center", gap: 8, padding: "6px 0", cursor: "pointer" }}
                    >
                      <div style={{ width: 10, height: 10, borderRadius: "50%", background: CAT_COLORS[i % CAT_COLORS.length], flexShrink: 0 }} />
                      <div style={{ flex: 1, fontSize: 13 }}>{p.name}</div>
                      <div style={{ fontFamily: "'IBM Plex Mono', monospace", fontSize: 13, fontWeight: 600 }}>{fmtARS(p.value)}</div>
                    </div>
                    {isOpen && (
                      <div style={{ marginLeft: 18, marginBottom: 6, display: "flex", flexDirection: "column", gap: 3 }}>
                        {movs.map((m) => (
                          <div key={m.id} style={{ display: "flex", justifyContent: "space-between", fontSize: 11.5, color: "#5a6b6d" }}>
                            <span style={{ overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap", maxWidth: 160 }}>{m.desc || m.category}</span>
                            <span style={{ fontFamily: "'IBM Plex Mono', monospace" }}>{fmtARS(m.amount)}</span>
                          </div>
                        ))}
                      </div>
                    )}
                  </div>
                );
              })}
            </div>
          </div>
        )}
      </div>

      {ingresosPorPersona.length > 0 && (
        <div style={{ background: "#fff", borderRadius: 10, padding: 18, boxShadow: "0 2px 10px rgba(27,42,46,0.06)" }}>
          <div style={{ fontFamily: "'Fraunces', serif", fontSize: 17, marginBottom: 12 }}>Ingresos por persona</div>
          <div style={{ display: "flex", flexDirection: "column", gap: 4 }}>
            {ingresosPorPersona.map((p) => (
              <div key={p.name} style={{ display: "flex", justifyContent: "space-between", fontSize: 13, padding: "4px 0" }}>
                <span>{p.name}</span>
                <span style={{ fontFamily: "'IBM Plex Mono', monospace", color: GREEN, fontWeight: 600 }}>{fmtARS(p.value)}</span>
              </div>
            ))}
          </div>
        </div>
      )}
    </div>
  );
}

function AdminTab() {
  const [filas, setFilas] = useState(null);
  const [cargando, setCargando] = useState(true);
  const [error, setError] = useState(null);
  const [expandido, setExpandido] = useState(null);

  useEffect(() => {
    (async () => {
      try {
        const rows = await sb("rpc/admin_resumen_logins", { method: "POST", body: "{}" });
        setFilas(rows || []);
      } catch (e) {
        setError(e.message || "No se pudo cargar.");
      }
      setCargando(false);
    })();
  }, []);

  function fmtFecha(iso) {
    return new Date(iso).toLocaleString("es-AR", { day: "2-digit", month: "short", year: "numeric", hour: "2-digit", minute: "2-digit" });
  }

  if (cargando) return <div style={{ fontSize: 13, color: "#8a9698" }}>Cargando...</div>;
  if (error) return <div style={{ color: BRICK, fontSize: 13 }}>No se pudo cargar el panel admin: {error}</div>;
  if (!filas || filas.length === 0) return <EmptyState text="Todavía no hay ingresos a la app registrados en ningún hogar." />;

  // Agrupamos hogar -> persona -> [fechas de ingreso]
  const hogares = {};
  filas.forEach((f) => {
    if (!hogares[f.household_id]) hogares[f.household_id] = { nombre: f.household_name, personas: {} };
    const persona = f.member_name || "(sin nombre)";
    if (!hogares[f.household_id].personas[persona]) hogares[f.household_id].personas[persona] = [];
    hogares[f.household_id].personas[persona].push(f.login_at);
  });

  return (
    <div style={{ display: "flex", flexDirection: "column", gap: 14 }}>
      <div style={{ fontSize: 12.5, color: "#8a9698" }}>
        Ingresos a la app (logins) por hogar y persona — el más reciente de cada uno queda destacado, el resto se puede desplegar. Acceso restringido a administradores.
      </div>
      {Object.entries(hogares).map(([hhId, hh]) => (
        <div key={hhId} style={{ background: "#fff", borderRadius: 10, padding: 16, boxShadow: "0 2px 10px rgba(27,42,46,0.06)" }}>
          <div style={{ fontFamily: "'Fraunces', serif", fontSize: 16, marginBottom: 6 }}>{hh.nombre}</div>
          {Object.entries(hh.personas).map(([nombre, logins]) => {
            const ordenados = [...logins].sort((a, b) => new Date(b) - new Date(a));
            const ultimo = ordenados[0];
            const resto = ordenados.slice(1);
            const key = `${hhId}-${nombre}`;
            const abierto = expandido === key;
            return (
              <div key={key} style={{ borderTop: "1px solid #f0ece0", paddingTop: 8, marginTop: 8 }}>
                <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center" }}>
                  <span style={{ fontSize: 13.5, fontWeight: 600 }}>{nombre}</span>
                  <span style={{ fontSize: 11, color: "#8a9698" }}>{ordenados.length} ingreso{ordenados.length === 1 ? "" : "s"}</span>
                </div>
                <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginTop: 4 }}>
                  <span style={{ fontSize: 12, color: "#8a9698" }}>Último ingreso</span>
                  <span style={{ fontFamily: "'IBM Plex Mono', monospace", color: TEAL, fontWeight: 700, fontSize: 13 }}>{fmtFecha(ultimo)}</span>
                </div>
                {resto.length > 0 && (
                  <>
                    <button
                      onClick={() => setExpandido(abierto ? null : key)}
                      style={{ background: "none", border: "none", color: "#8a9698", fontSize: 11.5, cursor: "pointer", padding: "6px 0", textDecoration: "underline" }}
                    >
                      {abierto ? "Ocultar anteriores" : `Ver ${resto.length} ingreso${resto.length === 1 ? "" : "s"} anterior${resto.length === 1 ? "" : "es"}`}
                    </button>
                    {abierto && (
                      <div style={{ display: "flex", flexDirection: "column", gap: 3, marginTop: 2 }}>
                        {resto.map((l, i) => (
                          <div key={i} style={{ fontSize: 11.5, color: "#5a6b6d" }}>{fmtFecha(l)}</div>
                        ))}
                      </div>
                    )}
                  </>
                )}
              </div>
            );
          })}
        </div>
      ))}
    </div>
  );
}

function AccesosRapidosTab({ accesosRapidos, categories, currentUserId, onRegistrar, onAdd, onEdit, onDelete }) {
  const [editando, setEditando] = useState(false);
  const [mostrarForm, setMostrarForm] = useState(false);
  const [nombre, setNombre] = useState("");
  const [categoria, setCategoria] = useState(categories[0] || "Otros");
  const [monto, setMonto] = useState("");
  const [cuenta, setCuenta] = useState("");
  const [personal, setPersonal] = useState(true);
  const [agregando, setAgregando] = useState(false);
  const [addError, setAddError] = useState(null);

  const [editandoId, setEditandoId] = useState(null);
  const [editNombre, setEditNombre] = useState("");
  const [editMonto, setEditMonto] = useState("");

  async function handleAdd() {
    if (!nombre.trim() || !monto || Number(monto) <= 0) return;
    setAddError(null);
    setAgregando(true);
    const res = await onAdd({ desc: nombre.trim(), category: categoria, amount: monto, account: cuenta, personal });
    setAgregando(false);
    if (res?.error) { setAddError(res.error); return; }
    setNombre("");
    setMonto("");
    setCuenta("");
    setPersonal(true);
    setMostrarForm(false);
  }

  function empezarEdicion(a) {
    setEditandoId(a.id);
    setEditNombre(a.descripcion || "");
    setEditMonto(String(a.amount));
  }
  async function confirmarEdicion(id) {
    await onEdit(id, { descripcion: editNombre, amount: Number(editMonto) || 0 });
    setEditandoId(null);
  }

  return (
    <div style={{ display: "flex", flexDirection: "column", gap: 14 }}>
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", gap: 10 }}>
        <div style={{ fontSize: 13, color: "#8a9698" }}>
          Un toque y listo: gastos que se repiten con el mismo monto y categoría (psicóloga, nafta, cochera...).
        </div>
        <div style={{ display: "flex", gap: 6, flexShrink: 0 }}>
          <button onClick={() => setEditando((v) => !v)} style={{ ...btnOutline, padding: "6px 10px", fontSize: 12 }}>
            {editando ? "Listo" : "Editar"}
          </button>
          <button onClick={() => setMostrarForm((v) => !v)} style={mostrarForm ? { ...btnOutline, padding: "6px 10px", fontSize: 12 } : { ...btnPrimary, padding: "6px 10px", fontSize: 12 }}>
            {mostrarForm ? "Cancelar" : (<><Plus size={14} /> Agregar</>)}
          </button>
        </div>
      </div>

      {mostrarForm && (
        <div style={{ background: "#fff", borderRadius: 10, padding: 18, boxShadow: "0 2px 10px rgba(27,42,46,0.06)" }}>
          <div style={{ fontFamily: "'Fraunces', serif", fontSize: 17, marginBottom: 12 }}>Nuevo acceso rápido</div>
          <label style={labelStyle}>Nombre</label>
          <input value={nombre} onChange={(e) => setNombre(e.target.value)} placeholder="Ej: Psicóloga" style={{ ...inputStyle, marginBottom: 14 }} />
          <label style={labelStyle}>Categoría</label>
          <select value={categoria} onChange={(e) => setCategoria(e.target.value)} style={{ ...inputStyle, marginBottom: 14 }}>
            {categories.map((c) => <option key={c} value={c}>{c}</option>)}
          </select>
          <label style={labelStyle}>Monto (ARS)</label>
          <input type="number" value={monto} onChange={(e) => setMonto(e.target.value)} placeholder="0" style={{ ...inputStyle, marginBottom: 14, fontSize: 18, fontFamily: "'IBM Plex Mono', monospace" }} />
          <label style={labelStyle}>Cuenta (opcional)</label>
          <input value={cuenta} onChange={(e) => setCuenta(e.target.value)} placeholder="Ej: Efectivo" style={{ ...inputStyle, marginBottom: 14 }} />
          <label style={labelStyle}>Visibilidad</label>
          <div style={{ display: "flex", gap: 6, marginBottom: 14 }}>
            {[[true, "Personal"], [false, "Del hogar"]].map(([val, label]) => (
              <button key={String(val)} onClick={() => setPersonal(val)} style={{
                flex: 1, padding: "8px 6px", borderRadius: 8, fontSize: 12.5, fontWeight: 700, cursor: "pointer",
                border: `1.5px solid ${personal === val ? TEAL : "#ddd6c4"}`,
                background: personal === val ? TEAL : "#fff", color: personal === val ? "#fff" : INK,
              }}>{label}</button>
            ))}
          </div>
          <div style={{ fontSize: 11, color: "#8a9698", marginTop: -8, marginBottom: 14 }}>
            {personal ? "Solo vos lo vas a ver y usar." : "Todos los miembros del hogar lo van a ver y podrán usarlo (pero solo vos podés editarlo o borrarlo)."}
          </div>
          {addError && <div style={{ color: BRICK, fontSize: 12.5, marginBottom: 10 }}>{addError}</div>}
          <button onClick={handleAdd} disabled={agregando} style={{ ...btnPrimary, width: "100%", justifyContent: "center" }}>
            {agregando ? "Agregando..." : "Agregar acceso rápido"}
          </button>
        </div>
      )}

      <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 10 }}>
        {accesosRapidos.length === 0 ? (
          <div style={{ gridColumn: "1 / -1" }}>
            <EmptyState text="Todavía no tenés accesos rápidos. Tocá Editar para crear el primero." />
          </div>
        ) : accesosRapidos.map((a) => {
          const enEdicion = editandoId === a.id;
          const esMio = !currentUserId || a.user_id === currentUserId;
          if (editando) {
            return (
              <div key={a.id} style={{ background: "#fff", borderRadius: 10, padding: 12, boxShadow: "0 1px 4px rgba(27,42,46,0.06)" }}>
                {enEdicion ? (
                  <>
                    <input value={editNombre} onChange={(e) => setEditNombre(e.target.value)} style={{ ...inputStyle, marginBottom: 6, padding: "6px 8px", fontSize: 13 }} />
                    <input type="number" value={editMonto} onChange={(e) => setEditMonto(e.target.value)} style={{ ...inputStyle, marginBottom: 8, padding: "6px 8px", fontSize: 13, fontFamily: "'IBM Plex Mono', monospace" }} />
                    <div style={{ display: "flex", gap: 6 }}>
                      <button onClick={() => confirmarEdicion(a.id)} style={{ ...btnPrimary, flex: 1, justifyContent: "center", padding: "6px", fontSize: 12 }}>Guardar</button>
                      <button onClick={() => setEditandoId(null)} style={{ ...btnOutline, flex: 1, justifyContent: "center", padding: "6px", fontSize: 12 }}>Cancelar</button>
                    </div>
                  </>
                ) : (
                  <>
                    <div style={{ display: "flex", alignItems: "center", gap: 5, marginBottom: 2 }}>
                      <div style={{ fontSize: 13, fontWeight: 600, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap", flex: 1 }}>{a.descripcion || a.category}</div>
                    </div>
                    <div style={{ display: "flex", alignItems: "center", gap: 4, marginBottom: 8 }}>
                      <span style={{ fontSize: 10, fontWeight: 700, padding: "1px 6px", borderRadius: 8, background: a.personal ? "#e8f0f0" : "#fdf1de", color: a.personal ? TEAL : GOLD }}>
                        {a.personal ? "Personal" : "Del hogar"}
                      </span>
                      {!esMio && <span style={{ fontSize: 10, color: "#8a9698" }}>· de {a.who || "otro miembro"}</span>}
                    </div>
                    <div style={{ fontSize: 11.5, color: "#8a9698", marginBottom: 8 }}>{a.category}</div>
                    <div style={{ fontFamily: "'IBM Plex Mono', monospace", fontWeight: 700, fontSize: 15, marginBottom: 8 }}>{fmtARS(a.amount)}</div>
                    {esMio && (
                      <div style={{ display: "flex", gap: 10 }}>
                        <button onClick={() => empezarEdicion(a)} style={{ background: "none", border: "none", cursor: "pointer", color: TEAL }}><Pencil size={15} /></button>
                        <button onClick={() => onDelete(a.id)} style={{ background: "none", border: "none", cursor: "pointer", color: BRICK }}><Trash2 size={15} /></button>
                      </div>
                    )}
                  </>
                )}
              </div>
            );
          }
          return (
            <button
              key={a.id}
              onClick={() => onRegistrar(a)}
              style={{
                background: "#fff", border: "1.5px solid #ddd6c4", borderRadius: 10, padding: "16px 12px",
                cursor: "pointer", textAlign: "left", boxShadow: "0 1px 4px rgba(27,42,46,0.06)",
              }}
            >
              <div style={{ fontSize: 13.5, fontWeight: 700, marginBottom: 4, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>{a.descripcion || a.category}</div>
              <span style={{ fontSize: 9.5, fontWeight: 700, padding: "1px 6px", borderRadius: 8, background: a.personal ? "#e8f0f0" : "#fdf1de", color: a.personal ? TEAL : GOLD }}>
                {a.personal ? "Personal" : "Del hogar"}
              </span>
              <div style={{ fontSize: 11, color: "#8a9698", marginTop: 6, marginBottom: 8 }}>{a.category}</div>
              <div style={{ fontFamily: "'IBM Plex Mono', monospace", fontWeight: 700, fontSize: 16, color: BRICK }}>{fmtARS(a.amount)}</div>
            </button>
          );
        })}
      </div>
    </div>
  );
}

function RecurrentesTab({ recurrentes, categories, onAdd, onToggleActivo, onEditMonto, onDelete }) {
  const [tipo, setTipo] = useState("gasto");
  const [categoria, setCategoria] = useState(categories[0] || "Otros");
  const [monto, setMonto] = useState("");
  const [desc, setDesc] = useState("");
  const [cuenta, setCuenta] = useState("");
  const [diaMes, setDiaMes] = useState("1");
  const [agregando, setAgregando] = useState(false);

  const [editandoId, setEditandoId] = useState(null);
  const [editValue, setEditValue] = useState("");
  const [addError, setAddError] = useState(null);

  const cats = tipo === "gasto" ? categories : INGRESO_CATS;

  useEffect(() => { setCategoria(cats[0] || ""); }, [tipo]); // eslint-disable-line

  async function handleAdd() {
    if (!monto || Number(monto) <= 0) return;
    setAddError(null);
    setAgregando(true);
    const res = await onAdd({ type: tipo, category: categoria, amount: monto, desc, account: cuenta, diaMes });
    setAgregando(false);
    if (res?.error) { setAddError(res.error); return; }
    setMonto("");
    setDesc("");
  }

  function empezarEdicion(r) {
    setEditandoId(r.id);
    setEditValue(String(r.amount));
  }
  async function confirmarEdicion(id) {
    await onEditMonto(id, editValue);
    setEditandoId(null);
  }

  return (
    <div style={{ display: "flex", flexDirection: "column", gap: 14 }}>
      <div style={{ fontSize: 13, color: "#8a9698", marginBottom: -4 }}>
        Cargá acá lo que se repite todos los meses (alquiler, seguros, suscripciones, sueldo). La app crea el movimiento solo, el día del mes que elijas — se genera la primera vez que alguien abre la app en esa fecha o después.
      </div>

      <div style={{ background: "#fff", borderRadius: 10, padding: 18, boxShadow: "0 2px 10px rgba(27,42,46,0.06)" }}>
        <div style={{ fontFamily: "'Fraunces', serif", fontSize: 17, marginBottom: 12 }}>Nuevo recurrente</div>

        <div style={{ display: "flex", gap: 6, marginBottom: 14 }}>
          {[["gasto", "Gasto", BRICK], ["ingreso", "Ingreso", GREEN]].map(([k, l, c]) => (
            <button key={k} onClick={() => setTipo(k)} style={{
              flex: 1, padding: "8px 6px", borderRadius: 8, fontSize: 12.5, fontWeight: 700, cursor: "pointer",
              border: `1.5px solid ${tipo === k ? c : "#ddd6c4"}`,
              background: tipo === k ? c : "#fff", color: tipo === k ? "#fff" : INK,
            }}>{l}</button>
          ))}
        </div>

        <label style={labelStyle}>Monto (ARS)</label>
        <input type="number" value={monto} onChange={(e) => setMonto(e.target.value)} placeholder="0" style={{ ...inputStyle, marginBottom: 14, fontSize: 18, fontFamily: "'IBM Plex Mono', monospace" }} />

        <label style={labelStyle}>Categoría</label>
        <select value={categoria} onChange={(e) => setCategoria(e.target.value)} style={{ ...inputStyle, marginBottom: 14 }}>
          {cats.map((c) => <option key={c} value={c}>{c}</option>)}
        </select>

        <label style={labelStyle}>Descripción (opcional)</label>
        <input value={desc} onChange={(e) => setDesc(e.target.value)} placeholder="Ej: Alquiler casita Hebraica" style={{ ...inputStyle, marginBottom: 14 }} />

        <div style={{ display: "flex", gap: 10, marginBottom: 14 }}>
          <div style={{ flex: 1 }}>
            <label style={labelStyle}>Cuenta (opcional)</label>
            <input value={cuenta} onChange={(e) => setCuenta(e.target.value)} placeholder="Ej: ARQ" style={inputStyle} />
          </div>
          <div style={{ width: 100 }}>
            <label style={labelStyle}>Día del mes</label>
            <input type="number" min={1} max={28} value={diaMes} onChange={(e) => setDiaMes(e.target.value)} style={inputStyle} />
          </div>
        </div>
        <div style={{ fontSize: 11, color: "#8a9698", marginTop: -8, marginBottom: 14 }}>
          Máximo día 28, para que funcione igual todos los meses (incluyendo febrero).
        </div>

        {addError && <div style={{ color: BRICK, fontSize: 12.5, marginBottom: 10 }}>{addError}</div>}
        <button onClick={handleAdd} disabled={agregando || !monto} style={{ ...btnPrimary, width: "100%", justifyContent: "center" }}>
          {agregando ? "Agregando..." : "Agregar recurrente"}
        </button>
      </div>

      <div style={{ display: "flex", flexDirection: "column", gap: 8 }}>
        {recurrentes.length === 0 ? (
          <EmptyState text="Todavía no cargaste ningún gasto o ingreso recurrente." />
        ) : recurrentes.map((r) => {
          const enEdicion = editandoId === r.id;
          return (
            <div key={r.id} style={{ background: "#fff", borderRadius: 8, padding: "12px 14px", boxShadow: "0 1px 4px rgba(27,42,46,0.06)", opacity: r.activo ? 1 : 0.55 }}>
              <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", gap: 8 }}>
                <div style={{ flex: 1, minWidth: 0 }}>
                  <div style={{ display: "flex", alignItems: "center", gap: 6, marginBottom: 2 }}>
                    <span style={{
                      fontSize: 10, fontWeight: 700, padding: "1px 7px", borderRadius: 10,
                      background: r.type === "gasto" ? "#fbe9e6" : "#e8f3ec",
                      color: r.type === "gasto" ? BRICK : GREEN,
                    }}>{r.type === "gasto" ? "Gasto" : "Ingreso"}</span>
                    <span style={{ fontSize: 13, fontWeight: 600 }}>{r.category}</span>
                  </div>
                  <div style={{ fontSize: 11.5, color: "#8a9698", overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>
                    {r.descripcion || "(sin descripción)"} · día {r.dia_mes}{r.account ? ` · ${r.account}` : ""}
                  </div>
                </div>
                {enEdicion ? (
                  <div style={{ display: "flex", alignItems: "center", gap: 4 }}>
                    <input
                      type="number"
                      value={editValue}
                      onChange={(e) => setEditValue(e.target.value)}
                      autoFocus
                      style={{ width: 90, padding: "4px 6px", borderRadius: 6, border: "1px solid #ddd6c4", fontSize: 13, textAlign: "right", fontFamily: "'IBM Plex Mono', monospace" }}
                      onKeyDown={(e) => { if (e.key === "Enter") confirmarEdicion(r.id); if (e.key === "Escape") setEditandoId(null); }}
                    />
                    <button onClick={() => confirmarEdicion(r.id)} style={{ background: "none", border: "none", cursor: "pointer", color: GREEN }}><Check size={16} /></button>
                  </div>
                ) : (
                  <button onClick={() => empezarEdicion(r)} style={{ background: "none", border: "none", cursor: "pointer", fontFamily: "'IBM Plex Mono', monospace", fontWeight: 700, fontSize: 14, color: r.type === "gasto" ? BRICK : GREEN, display: "flex", alignItems: "center", gap: 4 }}>
                    {fmtARS(r.amount)} <Pencil size={12} color="#8a9698" />
                  </button>
                )}
              </div>
              <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginTop: 10, paddingTop: 10, borderTop: "1px dashed #eee6d5" }}>
                <label style={{ display: "flex", alignItems: "center", gap: 6, fontSize: 12, cursor: "pointer", color: "#8a9698" }}>
                  <input type="checkbox" checked={r.activo} onChange={(e) => onToggleActivo(r.id, e.target.checked)} />
                  Activo
                </label>
                <button onClick={() => onDelete(r.id)} aria-label="Eliminar recurrente" style={{ background: "none", border: "none", cursor: "pointer", color: BRICK }}>
                  <Trash2 size={16} />
                </button>
              </div>
            </div>
          );
        })}
      </div>
    </div>
  );
}

function ResetTab({ onReset, onResetOverrides }) {
  const [confirmText, setConfirmText] = useState("");
  const [resetting, setResetting] = useState(false);
  const [done, setDone] = useState(false);
  const FRASE = "BORRAR TODO";

  const [confirmandoReglas, setConfirmandoReglas] = useState(false);
  const [resetingReglas, setResetingReglas] = useState(false);
  const [reglasReseteadas, setReglasReseteadas] = useState(false);

  async function handleReset() {
    setResetting(true);
    await onReset();
    setResetting(false);
    setDone(true);
    setConfirmText("");
  }

  async function handleResetReglas() {
    setResetingReglas(true);
    await onResetOverrides();
    setResetingReglas(false);
    setConfirmandoReglas(false);
    setReglasReseteadas(true);
  }

  return (
    <div style={{ display: "flex", flexDirection: "column", gap: 14 }}>
      <div style={{ background: "#fff", borderRadius: 10, padding: 18, boxShadow: "0 2px 10px rgba(27,42,46,0.06)" }}>
        <div style={{ fontFamily: "'Fraunces', serif", fontSize: 17, marginBottom: 6 }}>Reglas de categorización aprendidas</div>
        <div style={{ fontSize: 13, color: "#8a9698", marginBottom: 14 }}>
          Son las reglas que la app aprendió cada vez que recategorizaste algo a mano (para que las próximas importaciones de PDF/CSV categoricen solas). Borrarlas no toca ningún movimiento ni categoría — solo hace que tengas que volver a enseñarle esas reglas si las necesitás de nuevo.
        </div>
        {reglasReseteadas ? (
          <div style={{ background: "#e8f3ec", border: `1px solid ${GREEN}`, borderRadius: 8, padding: 12, fontSize: 13 }}>
            Listo, se borraron las reglas aprendidas.
          </div>
        ) : confirmandoReglas ? (
          <div style={{ display: "flex", gap: 8 }}>
            <button onClick={handleResetReglas} disabled={resetingReglas} style={{ ...btnPrimary, background: BRICK, flex: 1, justifyContent: "center" }}>
              {resetingReglas ? "Borrando..." : "Sí, borrar las reglas"}
            </button>
            <button onClick={() => setConfirmandoReglas(false)} style={{ ...btnOutline, flex: 1, justifyContent: "center" }}>
              Cancelar
            </button>
          </div>
        ) : (
          <button onClick={() => setConfirmandoReglas(true)} style={{ ...btnOutline, color: BRICK, borderColor: BRICK, justifyContent: "center", width: "100%" }}>
            Borrar reglas de categorización aprendidas
          </button>
        )}
      </div>

      <div style={{ background: "#fff", borderRadius: 10, padding: 18, boxShadow: "0 2px 10px rgba(27,42,46,0.06)" }}>
        <div style={{ fontFamily: "'Fraunces', serif", fontSize: 17, marginBottom: 6, color: BRICK }}>Reiniciar datos</div>
        <div style={{ fontSize: 13, color: "#8a9698", marginBottom: 14 }}>
          Pensado para esta etapa de pruebas. Esto borra <b>TODOS</b> los movimientos, presupuestos, recurrentes, accesos rápidos y el historial de cambios — de forma permanente, sin poder deshacerlo. No borra tu perfil, tus categorías, ni las reglas de categorización aprendidas.
        </div>

        {done ? (
          <div style={{ background: "#e8f3ec", border: `1px solid ${GREEN}`, borderRadius: 8, padding: 12, fontSize: 13 }}>
            Listo, se borró todo. La app queda como recién instalada.
          </div>
        ) : (
          <>
            <label style={labelStyle}>Para confirmar, escribí exactamente: <b>{FRASE}</b></label>
            <input
              value={confirmText}
              onChange={(e) => setConfirmText(e.target.value)}
              placeholder={FRASE}
              style={{ ...inputStyle, marginBottom: 14 }}
            />
            <button
              onClick={handleReset}
              disabled={confirmText.trim().toUpperCase() !== FRASE || resetting}
              style={{ ...btnPrimary, background: BRICK, width: "100%", justifyContent: "center" }}
            >
              {resetting ? "Borrando todo..." : "Borrar todos los datos"}
            </button>
          </>
        )}
      </div>
    </div>
  );
}

function HistorialTab() {
  const [log, setLog] = useState(null);

  useEffect(() => {
    (async () => {
      if (!HAS_SUPABASE) {
        setLog(safeGet("mock_audit_log") || []);
        return;
      }
      try {
        const rows = await sb("audit_log?select=*&order=created_at.desc&limit=300");
        setLog(rows || []);
      } catch (e) {
        console.error(e);
        setLog([]);
      }
    })();
  }, []);

  return (
    <div style={{ background: "#fff", borderRadius: 10, padding: 18, boxShadow: "0 2px 10px rgba(27,42,46,0.06)" }}>
      <div style={{ fontFamily: "'Fraunces', serif", fontSize: 17, marginBottom: 6 }}>Historial de cambios</div>
      <div style={{ fontSize: 13, color: "#8a9698", marginBottom: 14 }}>
        Ediciones de monto o descripción, y movimientos borrados — con su valor anterior para poder consultarlo.
      </div>

      {log === null ? (
        <div style={{ fontSize: 13, color: "#8a9698" }}>Cargando...</div>
      ) : log.length === 0 ? (
        <EmptyState text="Todavía no hay cambios registrados." />
      ) : (
        <div style={{ display: "flex", flexDirection: "column", gap: 8, maxHeight: 480, overflowY: "auto" }}>
          {log.map((r) => {
            const snap = r.entry_snapshot || {};
            let fechaCambio = "";
            try { fechaCambio = new Date(r.created_at).toLocaleString("es-AR"); } catch {}
            let texto;
            if (r.accion === "delete") {
              texto = `Borrado: ${snap.category || "(sin categoría)"} · ${snap.desc || "(sin descripción)"} · ${fmtARS(snap.amount || 0)} (${snap.date || "?"})`;
            } else if (r.accion === "edit_monto") {
              texto = `Monto editado: ${fmtARS(Number(r.valor_anterior) || 0)} → ${fmtARS(Number(r.valor_nuevo) || 0)} · ${snap.desc || snap.category || ""}`;
            } else if (r.accion === "edit_desc") {
              texto = `Descripción editada: "${r.valor_anterior || "(vacío)"}" → "${r.valor_nuevo || "(vacío)"}"`;
            } else if (r.accion === "import") {
              texto = `Importación (${snap.formato || "?"}): ${snap.cantidad_importada ?? 0} cargados de ${snap.cantidad_total ?? "?"}${snap.cantidad_duplicados ? `, ${snap.cantidad_duplicados} duplicados descartados` : ""}`;
            } else {
              texto = r.accion;
            }
            return (
              <div key={r.id} style={{ padding: "9px 12px", borderRadius: 6, background: r.accion === "delete" ? "#f7e9e6" : r.accion === "import" ? "#e3eeee" : "#f2eee2", fontSize: 12.5 }}>
                <div>{texto}</div>
                <div style={{ fontSize: 11, color: "#9a9488", marginTop: 3 }}>{fechaCambio} · {r.who || "?"}</div>
              </div>
            );
          })}
        </div>
      )}
    </div>
  );
}

function DuplicadosTab({ entries, onDelete, onForceAdd }) {
  const [borrando, setBorrando] = useState(false);
  const [rechazados, setRechazados] = useState(null);
  const [editKey, setEditKey] = useState(null);
  const [editDesc, setEditDesc] = useState("");
  const [editMonto, setEditMonto] = useState("");
  const [procesando, setProcesando] = useState(null);
  const [procesados, setProcesados] = useState(new Set());

  useEffect(() => {
    (async () => {
      let rows;
      if (!HAS_SUPABASE) {
        rows = safeGet("mock_audit_log") || [];
      } else {
        try {
          rows = await sb("audit_log?select=*&accion=eq.import&order=created_at.desc&limit=50");
        } catch (e) {
          console.error(e);
          rows = [];
        }
      }
      const conDuplicados = (rows || []).filter((r) => (r.entry_snapshot?.cantidad_duplicados || 0) > 0);
      setRechazados(conDuplicados);
    })();
  }, []);

  function keyDe(rowId, i) {
    return `${rowId}-${i}`;
  }
  function startEdit(rowId, i, d) {
    setEditKey(keyDe(rowId, i));
    setEditDesc(d.desc || "");
    setEditMonto(String(d.amount ?? ""));
  }
  function cancelEdit() {
    setEditKey(null);
    setEditDesc("");
    setEditMonto("");
  }
  async function procesarDeTodosModos(rowId, i, d) {
    const k = keyDe(rowId, i);
    setProcesando(k);
    await onForceAdd({
      date: d.date,
      type: d.type || "gasto",
      category: d.category || "Otros",
      amount: Number(editKey === k ? editMonto : d.amount) || 0,
      desc: editKey === k ? editDesc.trim() : (d.desc || ""),
      account: d.account || "",
    });
    setProcesados((prev) => new Set(prev).add(k));
    setProcesando(null);
    setEditKey(null);
  }

  const grupos = useMemo(() => {
    const sig = (e) => `${e.date}|${Number(e.amount)}|${(e.desc || "").trim().toLowerCase()}|${(e.account || "").trim().toLowerCase()}`;
    const map = {};
    entries.forEach((e) => {
      const s = sig(e);
      if (!map[s]) map[s] = [];
      map[s].push(e);
    });
    return Object.values(map).filter((g) => g.length > 1);
  }, [entries]);

  const totalSobrantes = grupos.reduce((s, g) => s + (g.length - 1), 0);

  async function borrarSobrantes() {
    if (totalSobrantes === 0) return;
    setBorrando(true);
    // Nos quedamos con el primero de cada grupo (el más viejo cargado) y borramos el resto
    const idsABorrar = grupos.flatMap((g) => g.slice(1).map((e) => e.id));
    for (const id of idsABorrar) {
      await onDelete(id);
    }
    setBorrando(false);
  }

  return (
    <div style={{ display: "flex", flexDirection: "column", gap: 14 }}>
      {rechazados && rechazados.length > 0 && (
        <div style={{ background: "#fff", borderRadius: 10, padding: 18, boxShadow: "0 2px 10px rgba(27,42,46,0.06)" }}>
          <div style={{ fontFamily: "'Fraunces', serif", fontSize: 17, marginBottom: 6 }}>Descartados al importar</div>
          <div style={{ fontSize: 13, color: "#8a9698", marginBottom: 14 }}>
            Movimientos que no se cargaron porque ya existía uno igual (misma fecha, monto, descripción y cuenta) al momento de importar.
          </div>
          <div style={{ display: "flex", flexDirection: "column", gap: 10, maxHeight: 360, overflowY: "auto" }}>
            {rechazados.map((r) => {
              let fecha = "";
              try { fecha = new Date(r.created_at).toLocaleString("es-AR"); } catch {}
              const detalle = r.entry_snapshot?.duplicados_detalle || [];
              return (
                <div key={r.id} style={{ background: "#fbf1de", borderRadius: 8, padding: 10 }}>
                  <div style={{ fontSize: 12.5, fontWeight: 700, marginBottom: 6 }}>
                    Importación {r.entry_snapshot?.formato || ""} · {fecha} · {detalle.length} descartado{detalle.length !== 1 ? "s" : ""}
                  </div>
                  <div style={{ display: "flex", flexDirection: "column", gap: 6 }}>
                    {detalle.map((d, i) => {
                      const k = keyDe(r.id, i);
                      const editing = editKey === k;
                      const yaCargado = procesados.has(k);
                      return (
                        <div key={i} style={{ background: "#fff", borderRadius: 6, padding: "6px 8px" }}>
                          {editing ? (
                            <div style={{ display: "flex", gap: 6, alignItems: "center", flexWrap: "wrap" }}>
                              <input
                                autoFocus
                                value={editDesc}
                                onChange={(ev) => setEditDesc(ev.target.value)}
                                placeholder="Descripción"
                                style={{ ...inputStyle, padding: "5px 8px", fontSize: 12, flex: 1, minWidth: 120 }}
                              />
                              <input
                                type="number"
                                value={editMonto}
                                onChange={(ev) => setEditMonto(ev.target.value)}
                                placeholder="Monto"
                                style={{ ...inputStyle, padding: "5px 8px", fontSize: 12, width: 100, fontFamily: "'IBM Plex Mono', monospace" }}
                              />
                              <button onClick={() => procesarDeTodosModos(r.id, i, d)} disabled={procesando === k} style={{ ...btnPrimary, padding: "5px 10px", fontSize: 11.5 }}>
                                {procesando === k ? "Cargando..." : "Cargar de todos modos"}
                              </button>
                              <button onClick={cancelEdit} style={{ border: "none", background: "none", cursor: "pointer", color: "#b8b2a4" }} aria-label="Cancelar">
                                <X size={15} />
                              </button>
                            </div>
                          ) : (
                            <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", gap: 8, fontSize: 12, color: yaCargado ? "#b8b2a4" : "#5a6b6d", textDecoration: yaCargado ? "line-through" : "none" }}>
                              <span style={{ overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap", flex: 1 }}>
                                {d.date} · {d.desc || "(sin descripción)"}{d.account ? ` · ${d.account}` : ""}
                              </span>
                              <span style={{ fontFamily: "'IBM Plex Mono', monospace", flexShrink: 0 }}>{fmtARS(d.amount || 0)}</span>
                              {!yaCargado && (
                                <button onClick={() => startEdit(r.id, i, d)} style={{ border: "none", background: "none", cursor: "pointer", color: TEAL, flexShrink: 0 }} aria-label="Editar y procesar">
                                  <Pencil size={13} />
                                </button>
                              )}
                              {yaCargado && <span style={{ fontSize: 10.5, color: GREEN, flexShrink: 0, fontWeight: 700 }}>Cargado ✓</span>}
                            </div>
                          )}
                        </div>
                      );
                    })}
                  </div>
                </div>
              );
            })}
          </div>
        </div>
      )}

      <div style={{ background: "#fff", borderRadius: 10, padding: 18, boxShadow: "0 2px 10px rgba(27,42,46,0.06)" }}>
      <div style={{ fontFamily: "'Fraunces', serif", fontSize: 17, marginBottom: 6 }}>Duplicados</div>
      <div style={{ fontSize: 13, color: "#8a9698", marginBottom: 14 }}>
        Considera duplicado a todo movimiento con misma fecha, monto, descripción y cuenta. Útil para limpiar lo que se haya importado más de una vez antes de este chequeo.
      </div>

      {grupos.length === 0 ? (
        <EmptyState text="No se encontraron duplicados." />
      ) : (
        <>
          <div style={{ fontSize: 13, marginBottom: 12 }}>
            Se encontraron <b>{grupos.length}</b> grupos con duplicados ({totalSobrantes} movimientos de más).
          </div>
          <div style={{ maxHeight: 320, overflowY: "auto", display: "flex", flexDirection: "column", gap: 10, marginBottom: 14 }}>
            {grupos.map((g, i) => (
              <div key={i} style={{ background: "#f7e9e6", borderRadius: 8, padding: 10 }}>
                <div style={{ fontSize: 12.5, fontWeight: 700, marginBottom: 4 }}>
                  {g.length} copias · {g[0].date} · {fmtARS(g[0].amount)} · {g[0].desc || "(sin descripción)"} {g[0].account ? `· ${g[0].account}` : ""}
                </div>
                <div style={{ fontSize: 11.5, color: "#8a9698" }}>Se conserva la primera cargada, se borran las otras {g.length - 1}.</div>
              </div>
            ))}
          </div>
          <button onClick={borrarSobrantes} disabled={borrando} style={{ ...btnPrimary, background: BRICK, width: "100%", justifyContent: "center" }}>
            {borrando ? "Borrando..." : `Borrar ${totalSobrantes} duplicados`}
          </button>
        </>
      )}
      </div>
    </div>
  );
}

function CategoryChips({ selected, onSelect, cats }) {
  return (
    <div style={{ display: "flex", flexWrap: "wrap", gap: 8, marginBottom: 14 }}>
      {cats.map((c) => {
        const i = cats.indexOf(c);
        const color = CAT_COLORS[(i >= 0 ? i : 0) % CAT_COLORS.length];
        const active = selected === c;
        return (
          <button
            key={c}
            onClick={() => onSelect(c)}
            style={{
              display: "flex", alignItems: "center", gap: 6, padding: "7px 12px", borderRadius: 20,
              border: `1.5px solid ${active ? color : "#ddd6c4"}`, background: active ? color : "#fff",
              color: active ? "#fff" : INK, fontSize: 12.5, fontWeight: 700, cursor: "pointer",
            }}
          >
            <span style={{ width: 8, height: 8, borderRadius: "50%", background: active ? "#fff" : color }} />
            {c}
          </button>
        );
      })}
    </div>
  );
}

function RecategorizarTab({ entries, onApply, categories }) {
  const [desde, setDesde] = useState(categories[0]);
  const [hacia, setHacia] = useState(categories[0]);
  const [busq, setBusq] = useState("");
  const [applying, setApplying] = useState(false);
  const [incluidos, setIncluidos] = useState(new Set());

  const coincidencias = useMemo(() => {
    return entries.filter((e) => {
      if (busq.trim()) return e.desc?.toLowerCase().includes(busq.trim().toLowerCase());
      return e.category === desde;
    });
  }, [entries, desde, busq]);

  useEffect(() => {
    setIncluidos(new Set());
  }, [desde, busq]);

  const seleccionados = coincidencias.filter((e) => incluidos.has(e.id));

  function toggleIncluido(id) {
    setIncluidos((prev) => {
      const next = new Set(prev);
      next.has(id) ? next.delete(id) : next.add(id);
      return next;
    });
  }
  function seleccionarTodos() {
    setIncluidos(new Set(coincidencias.map((e) => e.id)));
  }
  function deseleccionarTodos() {
    setIncluidos(new Set());
  }

  async function handleApply() {
    if (seleccionados.length === 0) return;
    setApplying(true);
    await onApply(seleccionados.map((e) => e.id), hacia);
    setApplying(false);
    setIncluidos(new Set());
  }

  return (
    <div style={{ background: "#fff", borderRadius: 10, padding: 18, boxShadow: "0 2px 10px rgba(27,42,46,0.06)" }}>
      <div style={{ fontFamily: "'Fraunces', serif", fontSize: 17, marginBottom: 6 }}>Recategorizar en bloque</div>
      <div style={{ fontSize: 13, color: "#8a9698", marginBottom: 14 }}>
        Elegí una categoría actual (o buscá por texto), revisá qué movimientos coinciden, destildá los que no correspondan, y elegí a qué categoría pasarlos.
      </div>

      <label style={labelStyle}>Buscar por texto en descripción (opcional)</label>
      <input value={busq} onChange={(e) => setBusq(e.target.value)} placeholder="Ej: Hebraica" style={{ ...inputStyle, marginBottom: 14 }} />

      {!busq.trim() && (
        <>
          <label style={labelStyle}>Categoría actual</label>
          <CategoryChips selected={desde} onSelect={setDesde} cats={categories} />
        </>
      )}

      <label style={labelStyle}>Cambiar a</label>
      <CategoryChips selected={hacia} onSelect={setHacia} cats={categories} />

      <div style={{ fontSize: 12.5, color: "#8a9698", marginBottom: 8, display: "flex", justifyContent: "space-between", alignItems: "center", flexWrap: "wrap", gap: 6 }}>
        <span>Coinciden <b>{coincidencias.length}</b> movimientos · <b>{seleccionados.length}</b> seleccionados para mover.</span>
        {coincidencias.length > 0 && (
          <span style={{ display: "flex", gap: 10 }}>
            <button onClick={seleccionarTodos} style={{ background: "none", border: "none", color: TEAL, fontSize: 12, fontWeight: 700, cursor: "pointer", padding: 0 }}>Seleccionar todos</button>
            <button onClick={deseleccionarTodos} style={{ background: "none", border: "none", color: "#8a9698", fontSize: 12, fontWeight: 700, cursor: "pointer", padding: 0 }}>Ninguno</button>
          </span>
        )}
      </div>

      {coincidencias.length > 0 && (
        <div style={{ maxHeight: 260, overflowY: "auto", display: "flex", flexDirection: "column", gap: 6, marginBottom: 14, border: "1px solid #f0ece0", borderRadius: 8, padding: 8 }}>
          {coincidencias.map((e) => {
            const incluido = incluidos.has(e.id);
            return (
              <label key={e.id} style={{ display: "flex", alignItems: "center", gap: 8, padding: "6px 4px", fontSize: 12.5, cursor: "pointer", background: incluido ? "#e8f3ec" : "transparent", borderRadius: 6 }}>
                <input type="checkbox" checked={incluido} onChange={() => toggleIncluido(e.id)} style={{ flexShrink: 0 }} />
                <span style={{ flex: 1, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>
                  {e.date} · {e.desc || "(sin descripción)"}{e.account ? ` · ${e.account}` : ""}
                </span>
                <span style={{ fontFamily: "'IBM Plex Mono', monospace", flexShrink: 0 }}>{fmtARS(e.amount)}</span>
              </label>
            );
          })}
        </div>
      )}

      <button onClick={handleApply} disabled={seleccionados.length === 0 || applying} style={{ ...btnPrimary, width: "100%", justifyContent: "center" }}>
        {applying ? "Aplicando..." : `Recategorizar ${seleccionados.length} movimientos a "${hacia}"`}
      </button>
    </div>
  );
}

function FotoReciboModal({ onClose, onExtracted, categories }) {
  const [imagenUrl, setImagenUrl] = useState(null);
  const [analizando, setAnalizando] = useState(false);
  const [error, setError] = useState(null);
  const fileRef = useRef(null);
  const blobRef = useRef(null);

  // Achica la foto antes de mandarla (las fotos de celu pesan varios MB;
  // para leer un ticket alcanza con mucho menos, así sale más rápido y
  // barato el análisis).
  function comprimirYMostrar(file) {
    const reader = new FileReader();
    reader.onload = (e) => {
      const img = new Image();
      img.onload = () => {
        const maxAncho = 1280;
        const escala = Math.min(1, maxAncho / img.width);
        const canvas = document.createElement("canvas");
        canvas.width = Math.round(img.width * escala);
        canvas.height = Math.round(img.height * escala);
        const ctx = canvas.getContext("2d");
        ctx.drawImage(img, 0, 0, canvas.width, canvas.height);
        canvas.toBlob((blob) => {
          blobRef.current = blob;
          setImagenUrl(URL.createObjectURL(blob));
        }, "image/jpeg", 0.75);
      };
      img.src = e.target.result;
    };
    reader.readAsDataURL(file);
  }

  function handleFile(e) {
    setError(null);
    const file = e.target.files?.[0];
    if (file) comprimirYMostrar(file);
    e.target.value = ""; // permite volver a elegir el mismo archivo si hace falta
  }

  function sacarOtra() {
    setImagenUrl(null);
    blobRef.current = null;
    setError(null);
  }

  async function analizar() {
    if (!blobRef.current) return;
    setAnalizando(true);
    setError(null);
    try {
      const image_base64 = await blobToBase64(blobRef.current);
      const resultado = await sbFunction("analizar-recibo", { image_base64, mime_type: "image/jpeg", categorias: categories });
      if (resultado?.error) throw new Error(resultado.error);
      const categoriaFinal = categories.includes(resultado.categoria) ? resultado.categoria : "Otros";
      const detalle = [resultado.comercio, resultado.desc].filter(Boolean).join(" — ");
      onExtracted({
        type: "gasto",
        amount: resultado.monto != null ? Number(resultado.monto) : "",
        category: categoriaFinal,
        date: resultado.fecha || todayISO(),
        desc: detalle,
        pagado: resultado.pagado !== false,
        origen: "foto",
      });
    } catch (e) {
      setError("No pude leer el recibo: " + e.message);
    }
    setAnalizando(false);
  }

  return (
    <div style={{ position: "fixed", inset: 0, background: "rgba(27,42,46,0.5)", display: "flex", alignItems: esPantallaAncha() ? "center" : "flex-end", justifyContent: "center", zIndex: 25 }} onClick={onClose}>
      <div onClick={(e) => e.stopPropagation()} style={{ background: PAPER, width: esPantallaAncha() ? 480 : "100%", maxWidth: 480, borderRadius: esPantallaAncha() ? 16 : "16px 16px 0 0", padding: "20px 20px 28px", maxHeight: "88vh", overflowY: "auto", boxShadow: esPantallaAncha() ? "0 20px 60px rgba(27,42,46,0.3)" : "none" }}>
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 6 }}>
          <div style={{ fontFamily: "'Fraunces', serif", fontSize: 19 }}>Foto de recibo</div>
          <button onClick={onClose} style={{ border: "none", background: "none", cursor: "pointer" }} aria-label="Cerrar"><X size={22} /></button>
        </div>
        <div style={{ fontSize: 12.5, color: "#8a9698", marginBottom: 16 }}>
          Sacale una foto al ticket o comprobante — la IA lee el monto, la fecha y elige la categoría. Después revisás todo antes de guardar.
        </div>

        {error && <div style={{ color: BRICK, fontSize: 12.5, marginBottom: 12 }}>{error}</div>}

        {!imagenUrl ? (
          <button
            onClick={() => fileRef.current?.click()}
            style={{ ...btnPrimary, width: "100%", justifyContent: "center", padding: "14px" }}
          >
            <Camera size={18} /> Sacar foto / elegir imagen
          </button>
        ) : (
          <>
            <img src={imagenUrl} alt="Recibo" style={{ width: "100%", borderRadius: 10, marginBottom: 12, maxHeight: 320, objectFit: "contain", background: "#111" }} />
            <div style={{ display: "flex", gap: 8, marginBottom: 14 }}>
              <button onClick={sacarOtra} style={{ ...btnOutline, flex: 1, justifyContent: "center" }}>Sacar otra</button>
              <button onClick={analizar} disabled={analizando} style={{ ...btnPrimary, flex: 1, justifyContent: "center" }}>
                {analizando ? "Analizando..." : "Analizar recibo"}
              </button>
            </div>
          </>
        )}

        <input ref={fileRef} type="file" accept="image/*" onChange={handleFile} style={{ display: "none" }} />
      </div>
    </div>
  );
}

function VoiceEntryModal({ onClose, onExtracted, categories }) {
  const [recording, setRecording] = useState(false);
  const [audioUrl, setAudioUrl] = useState(null);
  const [transcribing, setTranscribing] = useState(false);
  const [transcript, setTranscript] = useState("");
  const [error, setError] = useState(null);
  const mediaRecorderRef = useRef(null);
  const chunksRef = useRef([]);
  const blobRef = useRef(null);
  const streamRef = useRef(null);

  useEffect(() => {
    return () => { streamRef.current?.getTracks().forEach((t) => t.stop()); };
  }, []);

  async function startRecording() {
    setError(null);
    setTranscript("");
    try {
      const stream = await navigator.mediaDevices.getUserMedia({ audio: true });
      streamRef.current = stream;
      const mimeType = MediaRecorder.isTypeSupported("audio/webm")
        ? "audio/webm"
        : (MediaRecorder.isTypeSupported("audio/mp4") ? "audio/mp4" : "");
      const mr = new MediaRecorder(stream, mimeType ? { mimeType } : undefined);
      chunksRef.current = [];
      mr.ondataavailable = (e) => { if (e.data.size > 0) chunksRef.current.push(e.data); };
      mr.onstop = () => {
        const blob = new Blob(chunksRef.current, { type: mr.mimeType || "audio/webm" });
        blobRef.current = blob;
        setAudioUrl(URL.createObjectURL(blob));
        stream.getTracks().forEach((t) => t.stop());
      };
      mediaRecorderRef.current = mr;
      mr.start();
      setRecording(true);
    } catch (e) {
      setError("No pude acceder al micrófono: " + e.message);
    }
  }

  function stopRecording() {
    mediaRecorderRef.current?.stop();
    setRecording(false);
  }

  function grabarDeNuevo() {
    setAudioUrl(null);
    blobRef.current = null;
    setTranscript("");
    setError(null);
  }

  async function transcribir() {
    if (!blobRef.current) return;
    setTranscribing(true);
    setError(null);
    try {
      const audio_base64 = await blobToBase64(blobRef.current);
      const data = await sbFunction("transcribir-audio", { audio_base64, mime_type: blobRef.current.type });
      setTranscript(data.text || "");
    } catch (e) {
      setError("No pude transcribir el audio: " + e.message);
    }
    setTranscribing(false);
  }

  function usarTranscripcion() {
    onExtracted(extraerDatosDeTexto(transcript, categories));
  }

  return (
    <div style={{ position: "fixed", inset: 0, background: "rgba(27,42,46,0.5)", display: "flex", alignItems: esPantallaAncha() ? "center" : "flex-end", justifyContent: "center", zIndex: 25 }} onClick={onClose}>
      <div onClick={(e) => e.stopPropagation()} style={{ background: PAPER, width: esPantallaAncha() ? 480 : "100%", maxWidth: 480, borderRadius: esPantallaAncha() ? 16 : "16px 16px 0 0", padding: "20px 20px 28px", maxHeight: "88vh", overflowY: "auto", boxShadow: esPantallaAncha() ? "0 20px 60px rgba(27,42,46,0.3)" : "none" }}>
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 6 }}>
          <div style={{ fontFamily: "'Fraunces', serif", fontSize: 19 }}>Cargar por voz</div>
          <button onClick={onClose} style={{ border: "none", background: "none", cursor: "pointer" }} aria-label="Cerrar"><X size={22} /></button>
        </div>
        <div style={{ fontSize: 12.5, color: "#8a9698", marginBottom: 16 }}>
          Grabá diciendo el gasto o ingreso ("gasté 15 mil en el supermercado con la visa"), transcribilo y revisá los datos antes de guardar.
        </div>

        {error && <div style={{ color: BRICK, fontSize: 12.5, marginBottom: 12 }}>{error}</div>}

        {!audioUrl ? (
          <button
            onClick={recording ? stopRecording : startRecording}
            style={{
              ...btnPrimary, width: "100%", justifyContent: "center", padding: "14px",
              background: recording ? BRICK : TEAL,
            }}
          >
            {recording ? <Square size={18} /> : <Mic size={18} />}
            {recording ? "Detener grabación" : "Grabar"}
          </button>
        ) : (
          <>
            <audio controls src={audioUrl} style={{ width: "100%", marginBottom: 12 }} />
            <div style={{ display: "flex", gap: 8, marginBottom: 14 }}>
              <button onClick={grabarDeNuevo} style={{ ...btnOutline, flex: 1, justifyContent: "center" }}>Grabar de nuevo</button>
              <button onClick={transcribir} disabled={transcribing} style={{ ...btnPrimary, flex: 1, justifyContent: "center" }}>
                {transcribing ? "Transcribiendo..." : "Transcribir"}
              </button>
            </div>
          </>
        )}

        {transcript && (
          <>
            <label style={labelStyle}>Texto transcripto (editalo si hace falta)</label>
            <textarea
              value={transcript}
              onChange={(e) => setTranscript(e.target.value)}
              rows={3}
              style={{ ...inputStyle, marginBottom: 14, resize: "vertical" }}
            />
            <button onClick={usarTranscripcion} style={{ ...btnPrimary, width: "100%", justifyContent: "center" }}>
              Usar este texto para cargar el movimiento
            </button>
          </>
        )}
      </div>
    </div>
  );
}

function EntryForm({ onClose, onSave, saving, saveError, categories, initialData, profileName, miembrosHogar }) {
  const [type, setType] = useState(initialData?.type || "gasto");
  const [amount, setAmount] = useState(initialData?.amount != null ? String(initialData.amount) : "");
  const [category, setCategory] = useState(initialData?.category || categories[0]);
  const [desc, setDesc] = useState(initialData?.desc || "");
  const [date, setDate] = useState(initialData?.date || todayISO());
  const [usdAmount, setUsdAmount] = useState("");
  const [rate, setRate] = useState("");
  const [account, setAccount] = useState("");
  const [moneda, setMoneda] = useState("ARS");
  const [pagado, setPagado] = useState(initialData?.pagado !== undefined ? initialData.pagado : true);
  const opcionesPersona = miembrosHogar && miembrosHogar.length > 0 ? miembrosHogar : [profileName].filter(Boolean);
  const [persona, setPersona] = useState(initialData?.persona || profileName || opcionesPersona[0] || "");

  const cats = type === "gasto" ? categories : type === "ingreso" ? INGRESO_CATS : AHORRO_INSTR;
  const arsFromCambio = (Number(usdAmount) || 0) * (Number(rate) || 0);
  const tieneOperador = /[+\-*/()]/.test(amount);
  const montoCalculado = evaluarExpresion(amount);

  // Al montar con datos precargados (dictado por voz) no queremos que este
  // efecto pise la categoría ya elegida — solo debe correr cuando el usuario
  // cambia el tipo a mano, de ahí el guard con el ref.
  const yaMontado = useRef(false);
  useEffect(() => {
    if (!yaMontado.current) { yaMontado.current = true; return; }
    setCategory((type === "gasto" ? categories : type === "ingreso" ? INGRESO_CATS : AHORRO_INSTR)[0]);
  }, [type]);

  function agregarOperador(op) {
    setAmount((prev) => prev + op);
  }

  function handleSubmit() {
    const origen = initialData?.origen || "manual";
    if (type === "cambio") {
      if (!usdAmount || Number(usdAmount) <= 0 || !rate || Number(rate) <= 0) return;
      onSave({
        type: "cambio", category: "Cambio USD→ARS", amount: arsFromCambio,
        usdAmount: Number(usdAmount), rate: Number(rate), desc, date, account, persona, origen
      });
      return;
    }
    if (!montoCalculado || montoCalculado <= 0) return;
    onSave({ type, amount: montoCalculado, category, desc, date, account, moneda: type === "gasto" || type === "ingreso" ? moneda : "ARS", pagado: type === "gasto" ? pagado : true, persona, origen });
  }

  return (
    <div style={{ position: "fixed", inset: 0, background: "rgba(27,42,46,0.5)", display: "flex", alignItems: esPantallaAncha() ? "center" : "flex-end", justifyContent: "center", zIndex: 20 }} onClick={onClose}>
      <div onClick={(e) => e.stopPropagation()} style={{ background: PAPER, width: esPantallaAncha() ? 480 : "100%", maxWidth: 480, borderRadius: esPantallaAncha() ? 16 : "16px 16px 0 0", padding: "20px 20px 28px", maxHeight: "88vh", overflowY: "auto", boxShadow: esPantallaAncha() ? "0 20px 60px rgba(27,42,46,0.3)" : "none" }}>
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 16 }}>
          <div style={{ fontFamily: "'Fraunces', serif", fontSize: 19 }}>Nuevo movimiento</div>
          <button onClick={onClose} style={{ border: "none", background: "none", cursor: "pointer" }} aria-label="Cerrar"><X size={22} /></button>
        </div>

        <div style={{ display: "flex", gap: 6, marginBottom: 16 }}>
          {[["gasto", "Gasto", BRICK], ["ingreso", "Ingreso", GREEN], ["ahorro", "Ahorro/Inversión", GOLD], ["cambio", "Cambio USD→ARS", TEAL]].map(([k, l, c]) => (
            <button key={k} onClick={() => setType(k)} style={{
              flex: "1 1 40%", padding: "9px 6px", borderRadius: 8, fontSize: 12, fontWeight: 700, cursor: "pointer",
              border: `1.5px solid ${type === k ? c : "#ddd6c4"}`,
              background: type === k ? c : "#fff", color: type === k ? "#fff" : INK
            }}>{l}</button>
          ))}
        </div>

        {type === "cambio" ? (
          <>
            <label style={labelStyle}>Monto en USD que cambiaste</label>
            <input type="number" value={usdAmount} onChange={(e) => setUsdAmount(e.target.value)} placeholder="0" style={{ ...inputStyle, marginBottom: 14, fontSize: 18, fontFamily: "'IBM Plex Mono', monospace" }} autoFocus />
            <label style={labelStyle}>Tipo de cambio (ARS por USD)</label>
            <input type="number" value={rate} onChange={(e) => setRate(e.target.value)} placeholder="Ej: 1250" style={{ ...inputStyle, marginBottom: 10, fontSize: 18, fontFamily: "'IBM Plex Mono', monospace" }} />
            {arsFromCambio > 0 && (
              <div style={{ fontSize: 13, color: TEAL, fontWeight: 700, marginBottom: 14 }}>= {fmtARS(arsFromCambio)}</div>
            )}
            <label style={labelStyle}>Cuenta / origen</label>
            <input value={account} onChange={(e) => setAccount(e.target.value)} placeholder="Ej: ARQ, Cueva, Change" style={{ ...inputStyle, marginBottom: 14 }} />
            <div style={{ fontSize: 11.5, color: "#8a9698", marginBottom: 14, marginTop: -8 }}>
              "ARQ" se muestra como sueldo; cualquier otro nombre queda marcado como cambio externo (igual suma a Ingresos).
            </div>
          </>
        ) : (
          <>
            {(type === "gasto" || type === "ingreso") && (
              <>
                <label style={labelStyle}>Moneda</label>
                <div style={{ display: "flex", gap: 6, marginBottom: 14 }}>
                  {["ARS", "USD"].map((m) => (
                    <button key={m} onClick={() => setMoneda(m)} style={{
                      flex: 1, padding: "8px 6px", borderRadius: 8, fontSize: 12.5, fontWeight: 700, cursor: "pointer",
                      border: `1.5px solid ${moneda === m ? TEAL : "#ddd6c4"}`,
                      background: moneda === m ? TEAL : "#fff", color: moneda === m ? "#fff" : INK,
                    }}>{m}</button>
                  ))}
                </div>
              </>
            )}
            <label style={labelStyle}>Monto ({moneda === "USD" ? "USD" : "ARS"})</label>
            <input
              type="text"
              inputMode="decimal"
              value={amount}
              onChange={(e) => setAmount(e.target.value)}
              placeholder="0"
              style={{ ...inputStyle, marginBottom: 8, fontSize: 20, fontFamily: "'IBM Plex Mono', monospace" }}
              autoFocus
            />
            <div style={{ display: "flex", gap: 6, marginBottom: 6 }}>
              {["+", "-", "*", "/"].map((op) => (
                <button
                  key={op}
                  type="button"
                  onClick={() => agregarOperador(op)}
                  style={{ flex: 1, padding: "7px 0", borderRadius: 6, border: "1px solid #ddd6c4", background: "#fff", color: TEAL, fontWeight: 700, fontSize: 15, cursor: "pointer" }}
                >
                  {op === "*" ? "×" : op === "/" ? "÷" : op}
                </button>
              ))}
            </div>
            {tieneOperador && (
              <div style={{ fontSize: 13, fontWeight: 700, marginBottom: 14, color: montoCalculado != null ? TEAL : BRICK }}>
                {montoCalculado != null ? `= ${fmtARS(montoCalculado)}` : "Expresión inválida"}
              </div>
            )}
            {!tieneOperador && <div style={{ marginBottom: 8 }} />}

            <label style={labelStyle}>{type === "ahorro" ? "Instrumento" : "Categoría"}</label>
            <select value={category} onChange={(e) => setCategory(e.target.value)} style={{ ...inputStyle, marginBottom: 14 }}>
              {cats.map((c) => <option key={c} value={c}>{c}</option>)}
            </select>

            <label style={labelStyle}>Tarjeta / cuenta (opcional)</label>
            <input value={account} onChange={(e) => setAccount(e.target.value)} placeholder="Ej: Visa BBVA Hernán" style={{ ...inputStyle, marginBottom: 14 }} />
            {moneda === "USD" && (type === "gasto" || type === "ingreso") && (
              <div style={{ fontSize: 11.5, color: "#8a9698", marginTop: -8, marginBottom: 14 }}>
                Este movimiento queda en dólares y no se mezcla con los totales en pesos.
              </div>
            )}
          </>
        )}

        <label style={labelStyle}>Descripción (opcional)</label>
        <input value={desc} onChange={(e) => setDesc(e.target.value)} placeholder="Ej: supermercado del sábado" style={{ ...inputStyle, marginBottom: 14 }} />

        {opcionesPersona.length > 1 && (
          <>
            <label style={labelStyle}>¿Quién lo gastó?</label>
            <div style={{ display: "flex", gap: 6, marginBottom: 14, flexWrap: "wrap" }}>
              {opcionesPersona.map((nombre) => (
                <button key={nombre} onClick={() => setPersona(nombre)} style={{
                  padding: "8px 12px", borderRadius: 8, fontSize: 12.5, fontWeight: 700, cursor: "pointer",
                  border: `1.5px solid ${persona === nombre ? TEAL : "#ddd6c4"}`,
                  background: persona === nombre ? TEAL : "#fff", color: persona === nombre ? "#fff" : INK,
                }}>{nombre}</button>
              ))}
            </div>
          </>
        )}

        {type === "gasto" && (
          <>
            <label style={labelStyle}>Estado</label>
            <div style={{ display: "flex", gap: 6, marginBottom: 14 }}>
              {[[true, "Pagado", GREEN], [false, "Pendiente de pago", GOLD]].map(([val, label, color]) => (
                <button key={String(val)} onClick={() => setPagado(val)} style={{
                  flex: 1, padding: "8px 6px", borderRadius: 8, fontSize: 12.5, fontWeight: 700, cursor: "pointer",
                  border: `1.5px solid ${pagado === val ? color : "#ddd6c4"}`,
                  background: pagado === val ? color : "#fff", color: pagado === val ? "#fff" : INK,
                }}>{label}</button>
              ))}
            </div>
            {!pagado && (
              <div style={{ fontSize: 11.5, color: "#8a9698", marginTop: -8, marginBottom: 14 }}>
                Útil para facturas/aranceles que cargás antes de pagarlos — lo marcás como "Pagado" después, desde Movimientos.
              </div>
            )}
          </>
        )}

        <label style={labelStyle}>Fecha</label>
        <div style={{ marginBottom: 20 }}><FechaInput value={date} onChange={setDate} /></div>

        {saveError && (
          <div style={{ color: BRICK, fontSize: 12.5, marginBottom: 10, background: "#fbeee6", padding: "8px 10px", borderRadius: 6 }}>
            {saveError}
          </div>
        )}

        <button onClick={handleSubmit} disabled={saving} style={{ ...btnPrimary, width: "100%", justifyContent: "center", padding: "13px", fontSize: 15 }}>
          {saving ? "Guardando..." : "Guardar movimiento"}
        </button>
      </div>
    </div>
  );
}

const labelStyle = { fontSize: 11.5, textTransform: "uppercase", letterSpacing: 0.5, color: "#8a9698", display: "block", marginBottom: 5, fontWeight: 600 };
const inputStyle = { width: "100%", padding: "10px 12px", borderRadius: 8, border: "1px solid #ddd6c4", fontSize: 14, fontFamily: "Inter, sans-serif", background: "#fff", color: INK };
const btnPrimary = { background: TEAL, color: "#fff", border: "none", borderRadius: 8, padding: "10px 16px", fontSize: 14, fontWeight: 700, cursor: "pointer", display: "flex", alignItems: "center", gap: 6 };
const btnOutline = { background: "#fff", color: INK, border: `1.5px solid #ddd6c4`, borderRadius: 8, padding: "11px 14px", fontSize: 14, fontWeight: 600, cursor: "pointer", display: "flex", alignItems: "center", gap: 8 };

const fontImports = `
@import url('https://fonts.googleapis.com/css2?family=Fraunces:wght@500;600;700&family=Inter:wght@400;500;600;700&family=IBM+Plex+Mono:wght@500;600&display=swap');
`;
