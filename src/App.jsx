import React, { useState, useEffect, useMemo } from "react";
import * as XLSX from "xlsx";
import {
  PieChart, Pie, Cell, ResponsiveContainer, Tooltip,
  BarChart, Bar, XAxis, YAxis, CartesianGrid, Legend
} from "recharts";
import {
  Wallet, TrendingUp, TrendingDown, PiggyBank, Target, Plus, X,
  ArrowUpRight, ArrowDownRight, Landmark, Settings, Trash2, User, Download, Menu,
  Home, List, Upload, Pencil, Check
} from "lucide-react";

const DEFAULT_GASTO_CATS = ["Comida", "Tarjetas", "Ropa", "Salud", "Educación", "Transporte", "Ocio", "Servicios", "Vivienda", "Otros"];
const TAB_LABELS = {
  resumen: "Resumen",
  movimientos: "Movimientos",
  ahorros: "Ahorros e inversiones",
  presupuestos: "Presupuestos",
  importar: "Importar",
  duplicados: "Duplicados",
  recategorizar: "Recategorizar",
  divisas: "Divisas",
  historial: "Historial de cambios",
  reset: "Reiniciar datos",
  hogar: "Mi hogar",
  categorias: "Categorías",
};
const PRIMARY_TABS = [
  ["resumen", "Resumen", Home],
  ["movimientos", "Movimientos", List],
  ["presupuestos", "Presupuestos", Target],
  ["importar", "Importar", Upload],
];
const SECONDARY_TABS = ["hogar", "categorias", "recategorizar", "duplicados", "divisas", "historial", "ahorros", "reset"];
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

// Subí este número cada vez que Claude te entregue un archivo nuevo.
// Sirve para confirmar de un vistazo que el deploy tomó la versión correcta.
// v47 · 2026-08-02 · layout responsive: contenido más ancho y gráficos/presupuestos en grilla en desktop
const APP_VERSION = "v47 · 2026-08-02";

function fmtARS(n) {
  const v = Number(n) || 0;
  return v.toLocaleString("es-AR", { style: "currency", currency: "ARS", maximumFractionDigits: 0 });
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
async function extraerLineasPdf(file) {
  const pdfjsLib = await cargarPdfjs();
  const buf = await file.arrayBuffer();
  const pdf = await conTimeout(pdfjsLib.getDocument({ data: buf }).promise, 25000, "El PDF tardó demasiado en procesarse (más de 25s). Puede ser un problema de conexión con el worker de lectura.");
  const lineas = [];
  for (let p = 1; p <= pdf.numPages; p++) {
    const page = await pdf.getPage(p);
    const content = await page.getTextContent();
    const porY = {};
    content.items.forEach((it) => {
      const y = Math.round(it.transform[5]);
      if (!porY[y]) porY[y] = [];
      porY[y].push(it);
    });
    const ys = Object.keys(porY).map(Number).sort((a, b) => b - a);
    ys.forEach((y) => {
      const linea = porY[y].sort((a, b) => a.transform[4] - b.transform[4]).map((it) => it.str).join(" ").replace(/\s+/g, " ").trim();
      if (linea) lineas.push(linea);
    });
  }
  return lineas;
}

// Parser específico del formato de resumen BBVA (Visa Signature / Mastercard Black)
function parsearResumenBBVA(lineas, nombreArchivo, overrides = {}) {
  const filas = [];
  const avisos = [];

  const cuentaMatch = lineas.find((l) => /Visa Signature|Mastercard Black/i.test(l));
  const cuenta = cuentaMatch ? (cuentaMatch.match(/Visa Signature|Mastercard Black/i) || [])[0] : nombreArchivo;

  let seccionActual = null; // "Hernan Israel" | "Natalia Wajsman" | null
  const LINEA_MOV = /^(\d{2})-([A-Za-zÁÉÍÓÚáéíóú]{3})-(\d{2})\s+(.+?)\s+(\d{6})\s+(-?[\d.,]+)\s*$/;

  lineas.forEach((linea) => {
    if (/^Consumos\s+(Hernan Israel|Natalia Wajsman)/i.test(linea)) {
      seccionActual = /Natalia/i.test(linea) ? "Natalia Wajsman" : "Hernan Israel";
      return;
    }
    if (/^TOTAL CONSUMOS/i.test(linea)) { seccionActual = null; return; }
    if (!seccionActual) return;
    if (/\bUSD\b\s*\d/.test(linea) && !/,\d{2}\s*$/.test(linea.replace(/USD.*$/, ""))) {
      // Línea con importe solo en USD (Apple, Google, Spotify, etc.) — la salteamos,
      // igual que venimos haciendo a mano, porque no es deuda en pesos.
      return;
    }
    const m = linea.match(LINEA_MOV);
    if (!m) return;
    const [, dd, mmm, yy, descRaw, , montoRaw] = m;
    const fecha = fechaBbvaAIso(dd, mmm, yy);
    if (!fecha) { avisos.push(`No pude leer la fecha en: "${linea}"`); return; }
    const monto = Number(montoRaw.replace(/\./g, "").replace(",", "."));
    if (!monto) return;
    const desc = descRaw.replace(/\s*C\.\d{2}\/\d{2}\s*$/, (s) => ` (cuota${s.trim().replace("C.", " ")})`).trim();
    filas.push({
      date: fecha,
      type: "gasto",
      category: inferCategory(desc, overrides),
      amount: monto,
      desc: seccionActual === "Natalia Wajsman" ? `${desc} (Natalia)` : desc,
      account: cuenta,
    });
  });

  return { filas, avisos };
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
function parsearResumenMercadoPago(fullText, overrides = {}) {
  const filas = [];
  const avisos = [];
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
    const [, fechaStr, descRaw, , valorStr] = m;
    const desc = descRaw.trim();

    if (/^Rendimientos$/i.test(desc)) continue; // regla acordada: se excluyen
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

    filas.push({ date: fecha, type, category, amount: Math.abs(valor), desc, account: "Mercado Pago", who: titular || undefined });
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
  });

  return { filas, avisos };
}

function monthKey(dateStr) {
  return dateStr ? dateStr.slice(0, 7) : "";
}
function todayISO() {
  return new Date().toISOString().slice(0, 10);
}
function uid() {
  return Date.now().toString(36) + Math.random().toString(36).slice(2, 7);
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
// de RLS por hogar (auth.uid()) funcionen.
let _accessToken = null;
function setAccessToken(token) { _accessToken = token; }

async function sb(path, options = {}) {
  const res = await fetch(`${SUPABASE_URL}/rest/v1/${path}`, {
    ...options,
    headers: {
      apikey: SUPABASE_ANON_KEY,
      Authorization: `Bearer ${_accessToken || SUPABASE_ANON_KEY}`,
      "Content-Type": "application/json",
      ...(options.headers || {}),
    },
  });
  if (!res.ok) {
    const body = await res.text().catch(() => "");
    throw new Error(`Supabase ${path} → ${res.status}: ${body}`);
  }
  if (res.status === 204) return null;
  return res.json().catch(() => null);
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
    usdAmount: row.usd_amount != null ? Number(row.usd_amount) : undefined,
    rate: row.rate != null ? Number(row.rate) : undefined,
    moneda: row.moneda || "ARS",
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
    usd_amount: e.usdAmount != null ? Number(e.usdAmount) : null,
    rate: e.rate != null ? Number(e.rate) : null,
    moneda: e.moneda || "ARS",
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
  };
}
function mockSaveEntries(entries) { safeSet("mock_entries", entries); }
function mockSaveCategories(categories) { safeSet("mock_categories", categories); }
function mockSaveBudgets(budgets) { safeSet("mock_budgets", budgets); }
function mockSaveNames(names) { safeSet("mock_names", names); }
function mockSaveOverrides(overrides) { safeSet("mock_overrides", overrides); }

export default function FinanzasApp() {
  const [loading, setLoading] = useState(true);
  const [session, setSession] = useState(null); // { access_token, refresh_token, user }
  const [householdId, setHouseholdId] = useState(null);
  const [profileName, setProfileName] = useState(null); // display_name dentro del hogar
  const [entries, setEntries] = useState([]);
  const [budgets, setBudgets] = useState({});
  const [categoryOverrides, setCategoryOverrides] = useState({});
  const [categories, setCategories] = useState(DEFAULT_GASTO_CATS);
  const [tab, setTab] = useState("resumen");
  const [menuOpen, setMenuOpen] = useState(false);
  const [showForm, setShowForm] = useState(false);
  const [formType, setFormType] = useState("gasto");
  const [saving, setSaving] = useState(false);

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

  async function cargarHogarYDatos(accessToken) {
    setAccessToken(accessToken);
    const miembro = await sb("household_members?select=household_id,display_name");
    if (!miembro || miembro.length === 0) {
      setAuthMode("onboarding");
      setLoading(false);
      return;
    }
    const hh = miembro[0];
    setHouseholdId(hh.household_id);
    setProfileName(hh.display_name);
    const [entRows, budRows, ovrRows, catRows] = await Promise.all([
      sb(`entries?household_id=eq.${hh.household_id}&select=*&order=date.desc`),
      sb(`budgets?household_id=eq.${hh.household_id}&select=*`),
      sb(`category_overrides?household_id=eq.${hh.household_id}&select=*`),
      sb(`categories?household_id=eq.${hh.household_id}&select=name&order=name.asc`),
    ]);
    setEntries((entRows || []).map(entryFromDb));
    const budObj = {};
    (budRows || []).forEach((b) => { budObj[b.category] = Number(b.limit_amount); });
    setBudgets(budObj);
    const ovrObj = {};
    (ovrRows || []).forEach((o) => { ovrObj[o.desc_key] = o.category; });
    setCategoryOverrides(ovrObj);
    setCategories((catRows || []).map((c) => c.name).length > 0 ? catRows.map((c) => c.name) : DEFAULT_GASTO_CATS);
    setTab("resumen");
    setLoading(false);
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
        setLoading(false);
        return;
      }
      try {
        const stored = safeGet("auth_session");
        if (!stored?.access_token) {
          setLoading(false);
          return;
        }
        setSession(stored);
        await cargarHogarYDatos(stored.access_token);
      } catch (e) {
        console.error(e);
        // token vencido u otro problema — pedimos login de nuevo
        safeSet("auth_session", null);
        setSession(null);
        setLoading(false);
      }
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
      await cargarHogarYDatos(data.access_token);
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
      await cargarHogarYDatos(data.access_token);
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
      await cargarHogarYDatos(session.access_token);
    } catch (e) {
      setAuthError(e.message);
    }
    setAuthBusy(false);
  }

  function handleLogout() {
    safeSet("auth_session", null);
    setAccessToken(null);
    setSession(null);
    setHouseholdId(null);
    setProfileName(null);
    setEntries([]);
    setBudgets({});
    setCategoryOverrides({});
    setAuthMode("login");
  }

  async function addEntry(entry) {
    const full = { ...entry, id: uid(), who: profileName };
    setEntries((prev) => {
      const next = [full, ...prev];
      if (!HAS_SUPABASE) mockSaveEntries(next);
      return next;
    });
    if (!HAS_SUPABASE) return;
    await sb("entries", { method: "POST", body: JSON.stringify([{ ...entryToDb(full), household_id: householdId }]) });
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

  async function resetearTodo() {
    setEntries([]);
    setBudgets({});
    setCategoryOverrides({});
    if (!HAS_SUPABASE) {
      mockSaveEntries([]);
      mockSaveBudgets({});
      mockSaveOverrides({});
      safeSet("mock_audit_log", []);
      return;
    }
    await Promise.all([
      sb("entries?id=not.is.null", { method: "DELETE" }),
      sb("budgets?category=not.is.null", { method: "DELETE" }),
      sb("category_overrides?desc_key=not.is.null", { method: "DELETE" }),
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

  function exportarExcel() {
    const filas = [...entries]
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
    XLSX.writeFile(wb, `contaduria_movimientos_${todayISO()}.xlsx`);
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

  if (loading) {
    return (
      <div style={{ minHeight: "100vh", background: PAPER, display: "flex", alignItems: "center", justifyContent: "center", fontFamily: "Inter, sans-serif", color: INK }}>
        Cargando...
      </div>
    );
  }

  if (loadError) {
    return (
      <div style={{ minHeight: "100vh", background: PAPER, display: "flex", alignItems: "center", justifyContent: "center", fontFamily: "Inter, sans-serif", color: BRICK, padding: 24, textAlign: "center", whiteSpace: "pre-wrap" }}>
        {loadError}
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
            <div style={{ fontFamily: "'Fraunces', serif", fontSize: 28, fontWeight: 600, color: INK, marginBottom: 6 }}>Contaduría</div>
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
          <div style={{ fontFamily: "'Fraunces', serif", fontSize: 28, fontWeight: 600, color: INK, marginBottom: 6 }}>Contaduría</div>
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

  return (
    <div style={{ minHeight: "100vh", background: PAPER, fontFamily: "Inter, sans-serif", color: INK, paddingBottom: 90 }}>
      <style>{fontImports}</style>
      <style>{`
        * { box-sizing: border-box; }
        button:focus-visible, input:focus-visible, select:focus-visible { outline: 2px solid ${TEAL}; outline-offset: 2px; }
        .tabbar button { transition: color 0.15s, border-color 0.15s; }
        .cat-row { transition: background 0.15s; }
        .cat-row:hover { background: ${PAPER_DIM}; }
        ::-webkit-scrollbar { height: 6px; width: 6px; }
      `}</style>

      {/* Header */}
      <div style={{ background: INK, color: PAPER, padding: "20px 20px 26px" }}>
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start", maxWidth: isDesktop ? 1080 : 720, margin: "0 auto" }}>
          <div>
            <div style={{ fontFamily: "'Fraunces', serif", fontSize: 24, fontWeight: 600 }}>Contaduría</div>
            <div style={{ fontSize: 12, color: "#9db3b0", marginTop: 2 }}>
              Hola, {profileName}
            </div>
            <div style={{ display: "flex", alignItems: "center", gap: 10, marginTop: 6 }}>
              <button onClick={() => shiftMonth(-1)} aria-label="Mes anterior" style={{ background: "none", border: "none", cursor: "pointer", color: PAPER, opacity: 0.8, padding: 2 }}>◀</button>
              <span style={{ fontSize: 13, fontWeight: 700, textTransform: "capitalize", minWidth: 130, textAlign: "center" }}>{monthLabel(selectedMonth)}</span>
              <button onClick={() => shiftMonth(1)} aria-label="Mes siguiente" style={{ background: "none", border: "none", cursor: "pointer", color: PAPER, opacity: 0.8, padding: 2 }}>▶</button>
              {selectedMonth !== realThisMonth && (
                <button onClick={() => setSelectedMonth(realThisMonth)} style={{ background: "none", border: `1px solid ${TEAL}`, borderRadius: 6, color: TEAL, fontSize: 11, padding: "2px 8px", cursor: "pointer" }}>
                  Hoy
                </button>
              )}
            </div>
            <div style={{ fontSize: 9.5, color: "#5f7376", marginTop: 6 }}>{APP_VERSION}</div>
            {!HAS_SUPABASE && (
              <div style={{ fontSize: 10.5, color: GOLD, marginTop: 4, fontWeight: 700 }}>
                ⚠ Vista previa local — no conectado a Supabase
              </div>
            )}
          </div>
          <div style={{ display: "flex", gap: 6, alignItems: "center", position: "relative" }}>
            <div title={profileName} style={{
              width: 28, height: 28, borderRadius: "50%",
              background: GOLD, color: INK,
              display: "flex", alignItems: "center", justifyContent: "center",
              fontSize: 12, fontWeight: 700
            }}>{profileName?.[0]?.toUpperCase()}</div>
            <button
              onClick={() => setMenuOpen((v) => !v)}
              aria-label="Más opciones"
              style={{ background: "none", border: "1px solid rgba(255,255,255,0.3)", borderRadius: 8, width: 32, height: 28, display: "flex", alignItems: "center", justifyContent: "center", cursor: "pointer", color: PAPER }}
            >
              <Menu size={16} />
            </button>
            {menuOpen && (
              <>
                <div onClick={() => setMenuOpen(false)} style={{ position: "fixed", inset: 0, zIndex: 9 }} />
                <div style={{
                  position: "absolute", top: "calc(100% + 6px)", right: 0, zIndex: 10, minWidth: 190,
                  background: "#fff", borderRadius: 8, boxShadow: "0 8px 28px rgba(27,42,46,0.25)",
                  border: "1px solid #ddd6c4", overflow: "hidden",
                }}>
                  {SECONDARY_TABS.map((key) => (
                    <button
                      key={key}
                      onClick={() => { setTab(key); setMenuOpen(false); }}
                      style={{
                        display: "block", width: "100%", textAlign: "left", padding: "11px 16px",
                        background: tab === key ? PAPER_DIM : "#fff", border: "none", borderBottom: "1px solid #f0ece0",
                        cursor: "pointer", fontSize: 13.5, fontWeight: tab === key ? 700 : 500,
                        color: tab === key ? TEAL : INK,
                      }}
                    >
                      {TAB_LABELS[key]}
                    </button>
                  ))}
                </div>
              </>
            )}
          </div>
        </div>
      </div>

      <div style={{ maxWidth: isDesktop ? 1080 : 720, margin: "0 auto", padding: "0 16px" }}>
        {SECONDARY_TABS.includes(tab) ? (
          <div style={{ marginTop: 18, display: "flex", alignItems: "center", gap: 10 }}>
            <button
              onClick={() => setTab("resumen")}
              style={{ ...btnOutline, padding: "8px 10px" }}
              aria-label="Volver al Resumen"
            >
              ← Volver
            </button>
            <div style={{ fontFamily: "'Fraunces', serif", fontSize: 19 }}>{TAB_LABELS[tab]}</div>
          </div>
        ) : (
          <>
            {/* Balance ticket */}
            <div style={{
              background: "#fff", marginTop: -18, borderRadius: 10, padding: "20px 18px",
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

            {/* Secciones principales */}
            <div style={{ marginTop: 22 }}>
              <div style={{ display: "grid", gridTemplateColumns: "repeat(4, 1fr)", gap: 8 }}>
                {PRIMARY_TABS.map(([key, label, Icon]) => (
                  <button
                    key={key}
                    onClick={() => setTab(key)}
                    style={{
                      display: "flex", flexDirection: "column", alignItems: "center", gap: 6,
                      padding: "14px 6px", borderRadius: 10, cursor: "pointer",
                      border: `1.5px solid ${tab === key ? TEAL : "#ddd6c4"}`,
                      background: tab === key ? TEAL : "#fff",
                      color: tab === key ? "#fff" : INK,
                      boxShadow: tab === key ? "0 4px 14px rgba(15,110,110,0.25)" : "0 1px 4px rgba(27,42,46,0.06)",
                    }}
                  >
                    <Icon size={20} />
                    <span style={{ fontSize: 11.5, fontWeight: 700, textAlign: "center", lineHeight: 1.1 }}>{label}</span>
                  </button>
                ))}
              </div>
              <div style={{ display: "flex", justifyContent: "flex-end", marginTop: 10 }}>
                <button
                  onClick={exportarExcel}
                  title="Exportar todos los movimientos a Excel"
                  style={{ ...btnOutline, padding: "8px 12px", fontSize: 12.5 }}
                >
                  <Download size={15} /> Excel
                </button>
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
            />
          )}
          {tab === "movimientos" && (
            <MovimientosTab allEntries={entries} entries={thisMonthEntries} onDelete={deleteEntry} onEditDesc={editEntryDesc} onEditAmount={editEntryAmount} profileName={profileName} monthLabel={monthLabel(selectedMonth)} />
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
              onImport={async (rows, formato) => {
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
                if (HAS_SUPABASE && nuevos.length > 0) {
                  await sb("entries", { method: "POST", body: JSON.stringify(nuevos.map((n) => ({ ...entryToDb(n), household_id: householdId }))) });
                }
                setEntries((prev) => {
                  const next = [...nuevos, ...prev];
                  if (!HAS_SUPABASE) mockSaveEntries(next);
                  return next;
                });
                await logAudit("import", {
                  formato: formato || "?",
                  cantidad_importada: nuevos.length,
                  cantidad_duplicados: descartados.length,
                  cantidad_total: rows.length,
                  duplicados_detalle: descartados,
                }, null, String(nuevos.length));
                return { imported: nuevos.length, duplicates: descartados.length };
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
            <ResetTab onReset={resetearTodo} />
          )}
          {tab === "hogar" && (
            <HogarTab householdId={householdId} onLogout={handleLogout} />
          )}
          {tab === "categorias" && (
            <CategoriasTab categories={categories} contarUsos={contarUsos} onAdd={addCategory} onRename={renameCategory} onDelete={deleteCategory} />
          )}
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

      {/* FAB */}
      <button
        onClick={() => setShowForm(true)}
        style={{
          position: "fixed", bottom: 22, right: 22, width: 56, height: 56, borderRadius: "50%",
          background: TEAL, color: "#fff", border: "none", boxShadow: "0 8px 20px rgba(15,110,110,0.4)",
          display: "flex", alignItems: "center", justifyContent: "center", cursor: "pointer", zIndex: 10
        }}
        aria-label="Agregar movimiento"
      >
        <Plus size={26} />
      </button>

      {showForm && (
        <EntryForm
          onClose={() => setShowForm(false)}
          onSave={async (entry) => { setSaving(true); await addEntry(entry); setSaving(false); setShowForm(false); }}
          saving={saving}
          categories={categories}
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
      <div style={{ background: "#fff", borderRadius: 10, padding: 18, boxShadow: "0 2px 10px rgba(27,42,46,0.06)", display: "flex", gap: 20, flexWrap: "wrap" }}>
        <div style={{ fontFamily: "'Fraunces', serif", fontSize: 17, width: "100%", marginBottom: 4 }}>Divisas</div>
        <div>
          <div style={{ fontSize: 11.5, color: "#8a9698", textTransform: "uppercase", letterSpacing: 0.5 }}>Último cambio</div>
          <div style={{ fontFamily: "'IBM Plex Mono', monospace", fontSize: 18, fontWeight: 600, color: TEAL, marginTop: 2 }}>
            ${cambiosStats.ultimo.rate} <span style={{ fontSize: 11, color: "#9a9488", fontWeight: 400 }}>({cambiosStats.ultimo.date})</span>
          </div>
        </div>
        <div>
          <div style={{ fontSize: 11.5, color: "#8a9698", textTransform: "uppercase", letterSpacing: 0.5 }}>Promedio histórico</div>
          <div style={{ fontFamily: "'IBM Plex Mono', monospace", fontSize: 18, fontWeight: 600, marginTop: 2 }}>${cambiosStats.promedio.toFixed(0)}</div>
        </div>
        <div>
          <div style={{ fontSize: 11.5, color: "#8a9698", textTransform: "uppercase", letterSpacing: 0.5 }}>USD cambiados en total</div>
          <div style={{ fontFamily: "'IBM Plex Mono', monospace", fontSize: 18, fontWeight: 600, marginTop: 2 }}>USD {cambiosStats.totalUsd.toLocaleString("es-AR")}</div>
        </div>
        <div style={{ width: "100%", display: "flex", gap: 20, flexWrap: "wrap", paddingTop: 12, marginTop: 4, borderTop: "1px solid #f0ece0" }}>
          <div>
            <div style={{ fontSize: 11.5, color: "#8a9698", textTransform: "uppercase", letterSpacing: 0.5 }}>Cambios ARQ (sueldo)</div>
            <div style={{ fontFamily: "'IBM Plex Mono', monospace", fontSize: 15, fontWeight: 600, marginTop: 2, color: GREEN }}>
              USD {cambiosStats.arq.totalUsd.toLocaleString("es-AR")} · {fmtARS(cambiosStats.arq.totalArs)}
            </div>
          </div>
          <div>
            <div style={{ fontSize: 11.5, color: "#8a9698", textTransform: "uppercase", letterSpacing: 0.5 }}>Cambios externos (cueva/change)</div>
            <div style={{ fontFamily: "'IBM Plex Mono', monospace", fontSize: 15, fontWeight: 600, marginTop: 2, color: GOLD }}>
              USD {cambiosStats.externos.totalUsd.toLocaleString("es-AR")} · {fmtARS(cambiosStats.externos.totalArs)}
            </div>
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

function ResumenTab({ gastosPorCategoria, totalAhorradoHistorico, entries, thisMonthEntries, cambiosStats, totalGastosUsd, isDesktop }) {
  const [expandedCat, setExpandedCat] = useState(null);
  const [rango, setRango] = useState(6);

  const chartData = useMemo(() => {
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
      months.push({ mes: label, Ingresos: ing, Gastos: gas });
    }
    return months;
  }, [entries, rango]);

  return (
    <div style={{ display: "flex", flexDirection: "column", gap: 24 }}>
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
          <div style={{ display: "flex", gap: 4 }}>
            {[[6, "6M"], [12, "12M"], [24, "24M"], ["todo", "Todo"]].map(([v, l]) => (
              <button key={l} onClick={() => setRango(v)} style={{
                padding: "5px 10px", borderRadius: 16, fontSize: 11.5, fontWeight: 700, cursor: "pointer",
                border: `1px solid ${rango === v ? TEAL : "#ddd6c4"}`,
                background: rango === v ? TEAL : "#fff", color: rango === v ? "#fff" : INK
              }}>{l}</button>
            ))}
          </div>
        </div>
        <div style={{ width: "100%", height: 220 }}>
          <ResponsiveContainer>
            <BarChart data={chartData}>
              <CartesianGrid strokeDasharray="3 3" stroke="#eee6d5" vertical={false} />
              <XAxis dataKey="mes" tick={{ fontSize: 11, fill: "#8a9698" }} axisLine={false} tickLine={false} interval={chartData.length > 12 ? 1 : 0} />
              <YAxis tick={{ fontSize: 11, fill: "#8a9698" }} axisLine={false} tickLine={false} width={40} tickFormatter={(v) => `${Math.round(v / 1000)}k`} />
              <Tooltip formatter={(v) => fmtARS(v)} />
              <Legend wrapperStyle={{ fontSize: 12 }} />
              <Bar dataKey="Ingresos" fill={GREEN} radius={[3, 3, 0, 0]} />
              <Bar dataKey="Gastos" fill={BRICK} radius={[3, 3, 0, 0]} />
            </BarChart>
          </ResponsiveContainer>
        </div>
      </div>
      </div>

      <div style={{ background: "#fff", borderRadius: 10, padding: 18, boxShadow: "0 2px 10px rgba(27,42,46,0.06)" }}>
        <div style={{ fontFamily: "'Fraunces', serif", fontSize: 17, marginBottom: 12 }}>Resumen mensual</div>
        {chartData.length === 0 ? <EmptyState text="Sin datos para este rango." /> : (
          <div style={{ display: "flex", flexDirection: "column" }}>
            <div style={{ display: "flex", fontSize: 11, textTransform: "uppercase", letterSpacing: 0.5, color: "#8a9698", paddingBottom: 8, borderBottom: "1px solid #eee6d5" }}>
              <div style={{ flex: 1.2 }}>Mes</div>
              <div style={{ flex: 1, textAlign: "right" }}>Ingresos</div>
              <div style={{ flex: 1, textAlign: "right" }}>Gastos</div>
              <div style={{ flex: 1, textAlign: "right" }}>Balance</div>
            </div>
            {[...chartData].reverse().map((m, i) => {
              const bal = m.Ingresos - m.Gastos;
              return (
                <div key={i} style={{ display: "flex", padding: "8px 0", borderBottom: "1px solid #f2eee2", fontSize: 12.5, alignItems: "center" }}>
                  <div style={{ flex: 1.2, fontWeight: 600 }}>{m.mes}</div>
                  <div style={{ flex: 1, textAlign: "right", color: GREEN, fontFamily: "'IBM Plex Mono', monospace" }}>{fmtARS(m.Ingresos)}</div>
                  <div style={{ flex: 1, textAlign: "right", color: BRICK, fontFamily: "'IBM Plex Mono', monospace" }}>{fmtARS(m.Gastos)}</div>
                  <div style={{ flex: 1, textAlign: "right", fontWeight: 700, color: bal >= 0 ? GREEN : BRICK, fontFamily: "'IBM Plex Mono', monospace" }}>{fmtARS(bal)}</div>
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

function MovimientosTab({ entries, allEntries, onDelete, onEditDesc, onEditAmount, profileName, monthLabel }) {
  const [filter, setFilter] = useState("todos");
  const [busq, setBusq] = useState("");
  const [alcance, setAlcance] = useState("mes"); // "mes" | "historial"
  const [editingId, setEditingId] = useState(null);
  const [editText, setEditText] = useState("");
  const [editAmount, setEditAmount] = useState("");

  function startEdit(e) {
    setEditingId(e.id);
    setEditText(e.desc || "");
    setEditAmount(String(e.amount ?? ""));
  }
  function cancelEdit() {
    setEditingId(null);
    setEditText("");
    setEditAmount("");
  }
  async function saveEdit(e) {
    const nuevoTexto = editText.trim();
    const nuevoMonto = Number(editAmount);
    if (nuevoTexto !== (e.desc || "")) await onEditDesc(e.id, nuevoTexto);
    if (nuevoMonto > 0 && nuevoMonto !== Number(e.amount)) await onEditAmount(e.id, nuevoMonto);
    setEditingId(null);
    setEditText("");
    setEditAmount("");
  }
  const buscando = busq.trim().length > 0;
  const usaTodo = alcance === "historial";

  const base = usaTodo ? allEntries : entries;
  const filtered = base.filter((e) => {
    const pasaTipo = filter === "todos" ? (e.type === "gasto" || e.type === "ingreso") : e.type === filter;
    if (!pasaTipo) return false;
    if (!buscando) return true;
    const q = busq.trim().toLowerCase();
    return (
      e.desc?.toLowerCase().includes(q) ||
      e.category?.toLowerCase().includes(q) ||
      e.account?.toLowerCase().includes(q) ||
      e.who?.toLowerCase().includes(q)
    );
  }).sort((a, b) => (usaTodo ? (b.date || "").localeCompare(a.date || "") : 0));

  return (
    <div>
      {monthLabel && (
        <div style={{ fontSize: 12, color: "#8a9698", marginBottom: 10, textTransform: "capitalize" }}>
          {usaTodo ? <>Mostrando <b>todo el historial</b></> : <>Mostrando: <b>{monthLabel}</b></>}
        </div>
      )}
      <input
        value={busq}
        onChange={(e) => setBusq(e.target.value)}
        placeholder="Buscar por descripción, categoría, cuenta o quién lo cargó..."
        style={{ ...inputStyle, marginBottom: 10 }}
      />
      <div style={{ display: "flex", gap: 6, marginBottom: 10 }}>
        {[["mes", "Buscar en el mes"], ["historial", "Buscar en todo el historial"]].map(([k, l]) => (
          <button key={k} onClick={() => setAlcance(k)} style={{
            padding: "6px 12px", borderRadius: 20, fontSize: 12, cursor: "pointer",
            border: `1px solid ${alcance === k ? GOLD : "#ddd6c4"}`,
            background: alcance === k ? GOLD : "#fff", color: alcance === k ? "#fff" : INK, fontWeight: 600
          }}>{l}</button>
        ))}
      </div>
      <div style={{ display: "flex", gap: 6, marginBottom: 14 }}>
        {[["todos", "Todos"], ["gasto", "Gastos"], ["ingreso", "Ingresos"], ["cambio", "Cambios USD"]].map(([k, l]) => (
          <button key={k} onClick={() => setFilter(k)} style={{
            padding: "6px 12px", borderRadius: 20, fontSize: 12.5, cursor: "pointer",
            border: `1px solid ${filter === k ? TEAL : "#ddd6c4"}`,
            background: filter === k ? TEAL : "#fff", color: filter === k ? "#fff" : INK, fontWeight: 600
          }}>{l}</button>
        ))}
      </div>
      {filtered.length === 0 ? <EmptyState text="No hay movimientos para mostrar." /> : (
        <div style={{ display: "flex", flexDirection: "column", gap: 8 }}>
          {filtered.map((e) => {
            const editing = editingId === e.id;
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
                    {editing ? (
                      <div style={{ display: "flex", flexDirection: "column", gap: 6 }}>
                        <div style={{ display: "flex", gap: 6, alignItems: "center" }}>
                          <input
                            autoFocus
                            value={editText}
                            onChange={(ev) => setEditText(ev.target.value)}
                            onKeyDown={(ev) => { if (ev.key === "Enter") saveEdit(e); if (ev.key === "Escape") cancelEdit(); }}
                            placeholder="Descripción"
                            style={{ ...inputStyle, padding: "5px 8px", fontSize: 13, flex: 1 }}
                          />
                          <input
                            type="number"
                            value={editAmount}
                            onChange={(ev) => setEditAmount(ev.target.value)}
                            onKeyDown={(ev) => { if (ev.key === "Enter") saveEdit(e); if (ev.key === "Escape") cancelEdit(); }}
                            placeholder="Monto"
                            style={{ ...inputStyle, padding: "5px 8px", fontSize: 13, width: 110, fontFamily: "'IBM Plex Mono', monospace" }}
                          />
                          <button onClick={() => saveEdit(e)} style={{ border: "none", background: "none", cursor: "pointer", color: GREEN, flexShrink: 0 }} aria-label="Guardar">
                            <Check size={16} />
                          </button>
                          <button onClick={cancelEdit} style={{ border: "none", background: "none", cursor: "pointer", color: "#b8b2a4", flexShrink: 0 }} aria-label="Cancelar">
                            <X size={16} />
                          </button>
                        </div>
                      </div>
                    ) : (
                      <div style={{ display: "flex", alignItems: "center", gap: 6, minWidth: 0 }}>
                        <div style={{ fontSize: 14, fontWeight: 600, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>{e.category}{e.desc ? ` · ${e.desc}` : ""}</div>
                        {e.type !== "cambio" && (
                          <button onClick={() => startEdit(e)} style={{ border: "none", background: "none", cursor: "pointer", color: "#c4bda8", flexShrink: 0, padding: 2 }} aria-label="Editar descripción y monto">
                            <Pencil size={13} />
                          </button>
                        )}
                      </div>
                    )}
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
      rows.push({ date: fecha, type: "cambio", category: "Cambio USD→ARS", amount: montoNum, desc: desc || "", account: cuenta || "", usdAmount: usdNum, rate: tcNum || Math.round(montoNum / usdNum) });
      return;
    }
    rows.push({ date: fecha, type: tipoNorm, category: categoria || "Otros", amount: montoNum, desc: desc || "", account: cuenta || "" });
  });
  return { rows, errors };
}

function ImportarTab({ onImport, categoryOverrides }) {
  const [modo, setModo] = useState("bbva"); // "bbva" | "mercadopago" | "colegio" | "csv"
  const [text, setText] = useState("");
  const [result, setResult] = useState(null);
  const [importing, setImporting] = useState(false);
  const [fileName, setFileName] = useState("");
  const [lastImportMsg, setLastImportMsg] = useState(null);
  const [pdfProcesando, setPdfProcesando] = useState(false);
  const [pdfAvisos, setPdfAvisos] = useState([]);
  const [pdfNombres, setPdfNombres] = useState([]);

  async function handlePdfFiles(e) {
    const files = Array.from(e.target.files || []);
    if (files.length === 0) return;
    setPdfProcesando(true);
    setPdfAvisos([]);
    setPdfNombres(files.map((f) => f.name));
    let todasLasFilas = [];
    let avisos = [];
    for (const file of files) {
      try {
        const lineas = await extraerLineasPdf(file);
        const { filas, avisos: av } = parsearResumenBBVA(lineas, file.name, categoryOverrides);
        if (filas.length === 0) {
          avisos.push(`${file.name}: no reconocí movimientos con el formato BBVA. Puede ser otro banco/billetera — pasámelo a mí directamente en el chat para procesarlo.`);
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
        const texto = await extraerTextoPdf(file);
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
        const texto = await extraerTextoPdf(file);
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
    const res = await onImport(result.rows, formatoLabel);
    setImporting(false);
    setText("");
    setResult(null);
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
            ["bbva", "PDF tarjeta BBVA"],
            ["mercadopago", "PDF Mercado Pago"],
            ["colegio", "PDF colegio (arancel)"],
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

function HogarTab({ householdId, onLogout }) {
  const [hh, setHh] = useState(null);
  const [miembros, setMiembros] = useState(null);
  const [copiado, setCopiado] = useState(false);
  const [qrDataUrl, setQrDataUrl] = useState(null);

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
                  href={`https://wa.me/?text=${encodeURIComponent(`Te invito a sumarte a nuestro hogar en Contaduría 🏠\nEntrá acá: ${inviteLink}`)}`}
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
                <div style={{ display: "flex", flexDirection: "column", gap: 4 }}>
                  {miembros.map((m, i) => (
                    <div key={i} style={{ fontSize: 13 }}>{m.display_name} {m.role === "owner" ? "· dueño/a" : ""}</div>
                  ))}
                </div>
              </div>
            )}
          </>
        )}
      </div>

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

function ResetTab({ onReset }) {
  const [confirmText, setConfirmText] = useState("");
  const [resetting, setResetting] = useState(false);
  const [done, setDone] = useState(false);
  const FRASE = "BORRAR TODO";

  async function handleReset() {
    setResetting(true);
    await onReset();
    setResetting(false);
    setDone(true);
    setConfirmText("");
  }

  return (
    <div style={{ background: "#fff", borderRadius: 10, padding: 18, boxShadow: "0 2px 10px rgba(27,42,46,0.06)" }}>
      <div style={{ fontFamily: "'Fraunces', serif", fontSize: 17, marginBottom: 6, color: BRICK }}>Reiniciar datos</div>
      <div style={{ fontSize: 13, color: "#8a9698", marginBottom: 14 }}>
        Pensado para esta etapa de pruebas. Esto borra <b>TODOS</b> los movimientos, presupuestos, reglas de categorización aprendidas, y el historial de cambios — de forma permanente, sin poder deshacerlo. No borra tu perfil ni los nombres guardados.
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

function EntryForm({ onClose, onSave, saving, categories }) {
  const [type, setType] = useState("gasto");
  const [amount, setAmount] = useState("");
  const [category, setCategory] = useState(categories[0]);
  const [desc, setDesc] = useState("");
  const [date, setDate] = useState(todayISO());
  const [usdAmount, setUsdAmount] = useState("");
  const [rate, setRate] = useState("");
  const [account, setAccount] = useState("");
  const [moneda, setMoneda] = useState("ARS");

  const cats = type === "gasto" ? categories : type === "ingreso" ? INGRESO_CATS : AHORRO_INSTR;
  const arsFromCambio = (Number(usdAmount) || 0) * (Number(rate) || 0);

  useEffect(() => {
    setCategory((type === "gasto" ? categories : type === "ingreso" ? INGRESO_CATS : AHORRO_INSTR)[0]);
  }, [type]);

  function handleSubmit() {
    if (type === "cambio") {
      if (!usdAmount || Number(usdAmount) <= 0 || !rate || Number(rate) <= 0) return;
      onSave({
        type: "cambio", category: "Cambio USD→ARS", amount: arsFromCambio,
        usdAmount: Number(usdAmount), rate: Number(rate), desc, date, account
      });
      return;
    }
    if (!amount || Number(amount) <= 0) return;
    onSave({ type, amount: Number(amount), category, desc, date, account, moneda: type === "gasto" || type === "ingreso" ? moneda : "ARS" });
  }

  return (
    <div style={{ position: "fixed", inset: 0, background: "rgba(27,42,46,0.5)", display: "flex", alignItems: "flex-end", zIndex: 20 }} onClick={onClose}>
      <div onClick={(e) => e.stopPropagation()} style={{ background: PAPER, width: "100%", borderRadius: "16px 16px 0 0", padding: "20px 20px 28px", maxHeight: "88vh", overflowY: "auto" }}>
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
            <input type="number" value={amount} onChange={(e) => setAmount(e.target.value)} placeholder="0" style={{ ...inputStyle, marginBottom: 14, fontSize: 20, fontFamily: "'IBM Plex Mono', monospace" }} autoFocus />

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

        <label style={labelStyle}>Fecha</label>
        <input type="date" value={date} onChange={(e) => setDate(e.target.value)} style={{ ...inputStyle, marginBottom: 20 }} />

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
