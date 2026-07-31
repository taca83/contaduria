import React, { useState, useEffect, useMemo } from "react";
import * as XLSX from "xlsx";
import {
  PieChart, Pie, Cell, ResponsiveContainer, Tooltip,
  BarChart, Bar, XAxis, YAxis, CartesianGrid, Legend
} from "recharts";
import {
  Wallet, TrendingUp, TrendingDown, PiggyBank, Target, Plus, X,
  ArrowUpRight, ArrowDownRight, Landmark, Settings, Trash2, User, Download
} from "lucide-react";

const GASTO_CATS = ["Comida", "Transporte", "Vivienda", "Servicios", "Salud", "Seguros", "Country/Hebraica", "Ocio", "Educación", "Ropa", "Otros"];
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

function fmtARS(n) {
  const v = Number(n) || 0;
  return v.toLocaleString("es-AR", { style: "currency", currency: "ARS", maximumFractionDigits: 0 });
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

const SUPABASE_URL = import.meta.env.VITE_SUPABASE_URL;
const SUPABASE_ANON_KEY = import.meta.env.VITE_SUPABASE_ANON_KEY;

async function sb(path, options = {}) {
  const res = await fetch(`${SUPABASE_URL}/rest/v1/${path}`, {
    ...options,
    headers: {
      apikey: SUPABASE_ANON_KEY,
      Authorization: `Bearer ${SUPABASE_ANON_KEY}`,
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

export default function FinanzasApp() {
  const [loading, setLoading] = useState(true);
  const [profileName, setProfileName] = useState(null);
  const [nameInput, setNameInput] = useState("");
  const [config, setConfig] = useState({ names: [] });
  const [entries, setEntries] = useState([]);
  const [budgets, setBudgets] = useState({});
  const [tab, setTab] = useState("resumen");
  const [showForm, setShowForm] = useState(false);
  const [formType, setFormType] = useState("gasto");
  const [saving, setSaving] = useState(false);

  const [loadError, setLoadError] = useState(null);

  useEffect(() => {
    (async () => {
      const prof = safeGet("profile");
      if (prof?.name) setProfileName(prof.name);
      try {
        const [cfgRows, entRows, budRows] = await Promise.all([
          sb("config?id=eq.1&select=names"),
          sb("entries?select=*&order=date.desc"),
          sb("budgets?select=*"),
        ]);
        setConfig({ names: cfgRows?.[0]?.names || [] });
        setEntries((entRows || []).map(entryFromDb));
        const budObj = {};
        (budRows || []).forEach((b) => { budObj[b.category] = Number(b.limit_amount); });
        setBudgets(budObj);
      } catch (e) {
        console.error(e);
        setLoadError(`No se pudo conectar con la base de datos.\n\nDetalle técnico: ${e.message}`);
      }
      setLoading(false);
    })();
  }, []);

  async function chooseName(name) {
    const clean = name.trim();
    if (!clean) return;
    setProfileName(clean);
    safeSet("profile", { name: clean });
    if (!config.names.includes(clean)) {
      const next = { names: [...config.names, clean] };
      setConfig(next);
      await sb("config?id=eq.1", { method: "PATCH", body: JSON.stringify({ names: next.names }) });
    }
  }

  async function addEntry(entry) {
    const full = { ...entry, id: uid(), who: profileName };
    setEntries((prev) => [full, ...prev]);
    await sb("entries", { method: "POST", body: JSON.stringify([entryToDb(full)]) });
  }

  async function deleteEntry(id) {
    setEntries((prev) => prev.filter((e) => e.id !== id));
    await sb(`entries?id=eq.${encodeURIComponent(id)}`, { method: "DELETE" });
  }

  async function updateBudgets(next) {
    setBudgets(next);
    const rows = Object.entries(next).map(([category, limit_amount]) => ({ category, limit_amount: Number(limit_amount) }));
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
  const thisMonth = now.toISOString().slice(0, 7);

  const thisMonthEntries = useMemo(
    () => entries.filter((e) => monthKey(e.date) === thisMonth),
    [entries, thisMonth]
  );
  const totalIngresos = useMemo(
    () => thisMonthEntries
      .filter((e) => e.type === "ingreso" || (e.type === "cambio" && e.account === "ARQ"))
      .reduce((s, e) => s + Number(e.amount), 0),
    [thisMonthEntries]
  );
  const totalGastos = useMemo(
    () => thisMonthEntries.filter((e) => e.type === "gasto").reduce((s, e) => s + Number(e.amount), 0),
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
    return { ultimo: cambios[0], promedio: totalArs / totalUsd, totalUsd, totalArs };
  }, [cambios]);

  const gastosPorCategoria = useMemo(() => {
    const map = {};
    thisMonthEntries.filter((e) => e.type === "gasto").forEach((e) => {
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
    return (
      <div style={{ minHeight: "100vh", background: INK, display: "flex", alignItems: "center", justifyContent: "center", padding: 24, fontFamily: "Inter, sans-serif" }}>
        <style>{fontImports}</style>
        <div style={{ background: PAPER, borderRadius: 4, padding: "40px 32px", maxWidth: 380, width: "100%", boxShadow: "0 20px 60px rgba(0,0,0,0.4)" }}>
          <div style={{ fontFamily: "'Fraunces', serif", fontSize: 28, fontWeight: 600, color: INK, marginBottom: 6 }}>Contaduría</div>
          <div style={{ color: "#5a6b6d", fontSize: 14, marginBottom: 24 }}>Finanzas compartidas. Decinos quién sos para empezar.</div>
          {config.names.length > 0 && (
            <div style={{ display: "flex", flexDirection: "column", gap: 8, marginBottom: 20 }}>
              {config.names.map((n) => (
                <button key={n} onClick={() => chooseName(n)} style={btnOutline}>
                  <User size={16} /> Soy {n}
                </button>
              ))}
            </div>
          )}
          <div style={{ fontSize: 12, color: "#8a8f5c", marginBottom: 8, textTransform: "uppercase", letterSpacing: 0.5 }}>
            {config.names.length > 0 ? "O ingresá otro nombre" : "Tu nombre"}
          </div>
          <div style={{ display: "flex", gap: 8 }}>
            <input
              value={nameInput}
              onChange={(e) => setNameInput(e.target.value)}
              onKeyDown={(e) => e.key === "Enter" && chooseName(nameInput)}
              placeholder="Ej: Juan"
              style={inputStyle}
            />
            <button onClick={() => chooseName(nameInput)} style={btnPrimary}>Entrar</button>
          </div>
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
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start", maxWidth: 720, margin: "0 auto" }}>
          <div>
            <div style={{ fontFamily: "'Fraunces', serif", fontSize: 24, fontWeight: 600 }}>Contaduría</div>
            <div style={{ fontSize: 12, color: "#9db3b0", marginTop: 2 }}>
              Hola, {profileName} · {now.toLocaleDateString("es-AR", { month: "long", year: "numeric" })}
            </div>
          </div>
          <div style={{ display: "flex", gap: 6 }}>
            {config.names.map((n) => (
              <div key={n} title={n} style={{
                width: 28, height: 28, borderRadius: "50%",
                background: n === profileName ? GOLD : "#33474a",
                color: n === profileName ? INK : PAPER,
                display: "flex", alignItems: "center", justifyContent: "center",
                fontSize: 12, fontWeight: 700
              }}>{n[0]?.toUpperCase()}</div>
            ))}
          </div>
        </div>
      </div>

      <div style={{ maxWidth: 720, margin: "0 auto", padding: "0 16px" }}>
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

        {/* Tabs */}
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginTop: 22, gap: 8 }}>
          <div className="tabbar" style={{ display: "flex", gap: 4, borderBottom: `1px solid #ddd6c4`, overflowX: "auto", flex: 1 }}>
            {[
              ["resumen", "Resumen"],
              ["movimientos", "Movimientos"],
              ["ahorros", "Ahorros e inversiones"],
              ["presupuestos", "Presupuestos"],
              ["importar", "Importar"],
              ["duplicados", "Duplicados"],
              ["recategorizar", "Recategorizar"],
            ].map(([key, label]) => (
              <button key={key} onClick={() => setTab(key)} style={{
                background: "none", border: "none", padding: "10px 12px", cursor: "pointer",
                fontSize: 13.5, fontWeight: 600, whiteSpace: "nowrap",
                color: tab === key ? INK : "#9a9488",
                borderBottom: tab === key ? `2px solid ${TEAL}` : "2px solid transparent",
              }}>{label}</button>
            ))}
          </div>
          <button
            onClick={exportarExcel}
            title="Exportar todos los movimientos a Excel"
            style={{ ...btnOutline, padding: "8px 12px", fontSize: 12.5, flexShrink: 0, marginBottom: 4 }}
          >
            <Download size={15} /> Excel
          </button>
        </div>

        <div style={{ marginTop: 20 }}>
          {tab === "resumen" && (
            <ResumenTab
              gastosPorCategoria={gastosPorCategoria}
              totalAhorradoHistorico={totalAhorradoHistorico}
              entries={entries}
              cambiosStats={cambiosStats}
            />
          )}
          {tab === "movimientos" && (
            <MovimientosTab entries={entries} onDelete={deleteEntry} profileName={profileName} />
          )}
          {tab === "ahorros" && (
            <AhorrosTab entries={entries.filter((e) => e.type === "ahorro")} onDelete={deleteEntry} totalAhorradoHistorico={totalAhorradoHistorico} />
          )}
          {tab === "presupuestos" && (
            <PresupuestosTab budgets={budgets} onUpdate={updateBudgets} gastosPorCategoria={gastosPorCategoria} />
          )}
          {tab === "importar" && (
            <ImportarTab
              onImport={async (rows) => {
                const sig = (e) => `${e.date}|${Number(e.amount)}|${(e.desc || "").trim().toLowerCase()}|${(e.account || "").trim().toLowerCase()}`;
                const existentes = new Set(entries.map(sig));
                const vistosEnLote = new Set();
                const nuevos = [];
                let duplicados = 0;
                rows.forEach((r) => {
                  const s = sig(r);
                  if (existentes.has(s) || vistosEnLote.has(s)) { duplicados++; return; }
                  vistosEnLote.add(s);
                  nuevos.push({ ...r, id: uid(), who: profileName });
                });
                if (nuevos.length > 0) {
                  await sb("entries", { method: "POST", body: JSON.stringify(nuevos.map(entryToDb)) });
                }
                setEntries((prev) => [...nuevos, ...prev]);
                return { imported: nuevos.length, duplicates: duplicados };
              }}
            />
          )}
          {tab === "duplicados" && (
            <DuplicadosTab entries={entries} onDelete={deleteEntry} />
          )}
          {tab === "recategorizar" && (
            <RecategorizarTab
              entries={entries}
              onApply={async (ids, nuevaCategoria) => {
                const idSet = new Set(ids);
                const next = entries.map((e) => idSet.has(e.id) ? { ...e, category: nuevaCategoria } : e);
                setEntries(next);
                await Promise.all(
                  ids.map((id) =>
                    sb(`entries?id=eq.${encodeURIComponent(id)}`, {
                      method: "PATCH",
                      body: JSON.stringify({ category: nuevaCategoria }),
                    })
                  )
                );
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

function ResumenTab({ gastosPorCategoria, totalAhorradoHistorico, entries, cambiosStats }) {
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
      const ing = entries.filter((e) => (e.type === "ingreso" || (e.type === "cambio" && e.account === "ARQ")) && monthKey(e.date) === key).reduce((s, e) => s + Number(e.amount), 0);
      const gas = entries.filter((e) => e.type === "gasto" && monthKey(e.date) === key).reduce((s, e) => s + Number(e.amount), 0);
      months.push({ mes: label, Ingresos: ing, Gastos: gas });
    }
    return months;
  }, [entries, rango]);

  return (
    <div style={{ display: "flex", flexDirection: "column", gap: 24 }}>
      {cambiosStats && (
        <div style={{ background: "#fff", borderRadius: 10, padding: 18, boxShadow: "0 2px 10px rgba(27,42,46,0.06)", display: "flex", gap: 20, flexWrap: "wrap" }}>
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
        </div>
      )}
      <div style={{ background: "#fff", borderRadius: 10, padding: 18, boxShadow: "0 2px 10px rgba(27,42,46,0.06)" }}>
        <div style={{ fontFamily: "'Fraunces', serif", fontSize: 17, marginBottom: 12 }}>Gastos por categoría (este mes)</div>
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
              {gastosPorCategoria.map((c, i) => (
                <div key={c.name} className="cat-row" style={{ display: "flex", justifyContent: "space-between", padding: "5px 6px", borderRadius: 6, fontSize: 13 }}>
                  <span style={{ display: "flex", alignItems: "center", gap: 6 }}>
                    <span style={{ width: 8, height: 8, borderRadius: "50%", background: CAT_COLORS[i % CAT_COLORS.length] }} />
                    {c.name}
                  </span>
                  <span style={{ fontFamily: "'IBM Plex Mono', monospace" }}>{fmtARS(c.value)}</span>
                </div>
              ))}
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

function MovimientosTab({ entries, onDelete, profileName }) {
  const [filter, setFilter] = useState("todos");
  const filtered = entries.filter((e) => filter === "todos" ? (e.type === "gasto" || e.type === "ingreso") : e.type === filter);

  return (
    <div>
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
          {filtered.map((e) => (
            <div key={e.id} style={{ background: "#fff", borderRadius: 8, padding: "12px 14px", display: "flex", justifyContent: "space-between", alignItems: "center", boxShadow: "0 1px 4px rgba(27,42,46,0.06)" }}>
              <div style={{ display: "flex", gap: 10, alignItems: "center", minWidth: 0 }}>
                <div style={{
                  width: 34, height: 34, borderRadius: "50%", flexShrink: 0,
                  background: e.type === "ingreso" ? "#e4f0e8" : e.type === "cambio" ? "#e3eeee" : "#f7e9e6",
                  display: "flex", alignItems: "center", justifyContent: "center"
                }}>
                  {e.type === "ingreso" ? <TrendingUp size={16} color={GREEN} /> : e.type === "cambio" ? <Landmark size={16} color={TEAL} /> : <TrendingDown size={16} color={BRICK} />}
                </div>
                <div style={{ minWidth: 0 }}>
                  <div style={{ fontSize: 14, fontWeight: 600, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>{e.category}{e.desc ? ` · ${e.desc}` : ""}</div>
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
          ))}
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

function PresupuestosTab({ budgets, onUpdate, gastosPorCategoria }) {
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
      {GASTO_CATS.map((cat) => {
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

function ImportarTab({ onImport }) {
  const [text, setText] = useState("");
  const [result, setResult] = useState(null);
  const [importing, setImporting] = useState(false);
  const [fileName, setFileName] = useState("");

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
    const res = await onImport(result.rows);
    setImporting(false);
    setText("");
    setResult(null);
    if (res && typeof res === "object") {
      const msg = res.duplicates > 0
        ? `Se importaron ${res.imported} movimientos. Se descartaron ${res.duplicates} por ser duplicados (misma fecha, monto, descripción y cuenta que uno ya cargado).`
        : `Se importaron ${res.imported} movimientos. Sin duplicados detectados.`;
      alert(msg);
    } else {
      alert(`Se importaron ${res} movimientos.`);
    }
  }

  return (
    <div style={{ display: "flex", flexDirection: "column", gap: 14 }}>
      <div style={{ background: "#fff", borderRadius: 10, padding: 18, boxShadow: "0 2px 10px rgba(27,42,46,0.06)" }}>
        <div style={{ fontFamily: "'Fraunces', serif", fontSize: 17, marginBottom: 6 }}>Importar movimientos</div>
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

function DuplicadosTab({ entries, onDelete }) {
  const [borrando, setBorrando] = useState(false);

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
  );
}

function RecategorizarTab({ entries, onApply }) {
  const [desde, setDesde] = useState(GASTO_CATS[0]);
  const [hacia, setHacia] = useState(GASTO_CATS[0]);
  const [busq, setBusq] = useState("");
  const [applying, setApplying] = useState(false);

  const coincidencias = useMemo(() => {
    return entries.filter((e) => {
      if (busq.trim()) return e.desc?.toLowerCase().includes(busq.trim().toLowerCase());
      return e.category === desde;
    });
  }, [entries, desde, busq]);

  async function handleApply() {
    if (coincidencias.length === 0) return;
    setApplying(true);
    await onApply(coincidencias.map((e) => e.id), hacia);
    setApplying(false);
    alert(`${coincidencias.length} movimientos actualizados a "${hacia}".`);
  }

  return (
    <div style={{ background: "#fff", borderRadius: 10, padding: 18, boxShadow: "0 2px 10px rgba(27,42,46,0.06)" }}>
      <div style={{ fontFamily: "'Fraunces', serif", fontSize: 17, marginBottom: 6 }}>Recategorizar en bloque</div>
      <div style={{ fontSize: 13, color: "#8a9698", marginBottom: 14 }}>
        Cambiá la categoría de muchos movimientos a la vez. Filtrá por categoría actual o por texto de la descripción.
      </div>

      <label style={labelStyle}>Buscar por texto en descripción (opcional, ej: "Hebraica")</label>
      <input value={busq} onChange={(e) => setBusq(e.target.value)} placeholder="Ej: Hebraica" style={{ ...inputStyle, marginBottom: 14 }} />

      {!busq.trim() && (
        <>
          <label style={labelStyle}>Categoría actual</label>
          <select value={desde} onChange={(e) => setDesde(e.target.value)} style={{ ...inputStyle, marginBottom: 14 }}>
            {GASTO_CATS.map((c) => <option key={c} value={c}>{c}</option>)}
          </select>
        </>
      )}

      <label style={labelStyle}>Cambiar a</label>
      <select value={hacia} onChange={(e) => setHacia(e.target.value)} style={{ ...inputStyle, marginBottom: 14 }}>
        {GASTO_CATS.map((c) => <option key={c} value={c}>{c}</option>)}
      </select>

      <div style={{ fontSize: 12.5, color: "#8a9698", marginBottom: 12 }}>
        Coinciden <b>{coincidencias.length}</b> movimientos.
      </div>

      <button onClick={handleApply} disabled={coincidencias.length === 0 || applying} style={{ ...btnPrimary, width: "100%", justifyContent: "center" }}>
        {applying ? "Aplicando..." : `Recategorizar ${coincidencias.length} movimientos`}
      </button>
    </div>
  );
}

function EntryForm({ onClose, onSave, saving }) {
  const [type, setType] = useState("gasto");
  const [amount, setAmount] = useState("");
  const [category, setCategory] = useState(GASTO_CATS[0]);
  const [desc, setDesc] = useState("");
  const [date, setDate] = useState(todayISO());
  const [usdAmount, setUsdAmount] = useState("");
  const [rate, setRate] = useState("");
  const [account, setAccount] = useState("");

  const cats = type === "gasto" ? GASTO_CATS : type === "ingreso" ? INGRESO_CATS : AHORRO_INSTR;
  const arsFromCambio = (Number(usdAmount) || 0) * (Number(rate) || 0);

  useEffect(() => {
    setCategory((type === "gasto" ? GASTO_CATS : type === "ingreso" ? INGRESO_CATS : AHORRO_INSTR)[0]);
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
    onSave({ type, amount: Number(amount), category, desc, date, account });
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
          </>
        ) : (
          <>
            <label style={labelStyle}>Monto (ARS)</label>
            <input type="number" value={amount} onChange={(e) => setAmount(e.target.value)} placeholder="0" style={{ ...inputStyle, marginBottom: 14, fontSize: 20, fontFamily: "'IBM Plex Mono', monospace" }} autoFocus />

            <label style={labelStyle}>{type === "ahorro" ? "Instrumento" : "Categoría"}</label>
            <select value={category} onChange={(e) => setCategory(e.target.value)} style={{ ...inputStyle, marginBottom: 14 }}>
              {cats.map((c) => <option key={c} value={c}>{c}</option>)}
            </select>

            <label style={labelStyle}>Tarjeta / cuenta (opcional)</label>
            <input value={account} onChange={(e) => setAccount(e.target.value)} placeholder="Ej: Visa BBVA Hernán" style={{ ...inputStyle, marginBottom: 14 }} />
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
