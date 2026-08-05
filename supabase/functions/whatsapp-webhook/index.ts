// supabase/functions/whatsapp-webhook/index.ts
//
// Webhook de WhatsApp (Meta Cloud API). Recibe mensajes de texto o audio,
// identifica de qué hogar es ese número, y:
//  - si es una carga ("gasté 5000 en nafta") → crea el movimiento
//  - si es una pregunta ("cuánto gasté en comida este mes?") → responde
//    en base a los datos reales de ese hogar (nunca inventa números)
//
// Usa la service_role key para leer/escribir sin depender de un JWT de
// sesión (un mensaje de WhatsApp no trae uno) — por eso whatsapp_links y
// entries se tocan acá con privilegios elevados, algo que nunca debe
// hacerse desde el cliente.

import { serve } from "https://deno.land/std@0.224.0/http/server.ts";

const SUPABASE_URL = Deno.env.get("SUPABASE_URL");
const SERVICE_ROLE_KEY = Deno.env.get("SUPABASE_SERVICE_ROLE_KEY");
const OPENAI_API_KEY = Deno.env.get("OPENAI_API_KEY");
const WHATSAPP_TOKEN = Deno.env.get("WHATSAPP_ACCESS_TOKEN");
const WHATSAPP_PHONE_ID = Deno.env.get("WHATSAPP_PHONE_NUMBER_ID");
const VERIFY_TOKEN = Deno.env.get("WHATSAPP_VERIFY_TOKEN");

const CORS_HEADERS = { "Access-Control-Allow-Origin": "*" };

// Chequeo al arrancar: si falta algún secreto, lo avisamos en los Logs de
// entrada — así se puede descartar esto de un vistazo sin adivinar.
for (const [nombre, valor] of [
  ["OPENAI_API_KEY", OPENAI_API_KEY],
  ["WHATSAPP_ACCESS_TOKEN", WHATSAPP_TOKEN],
  ["WHATSAPP_PHONE_NUMBER_ID", WHATSAPP_PHONE_ID],
  ["WHATSAPP_VERIFY_TOKEN", VERIFY_TOKEN],
]) {
  if (!valor) console.error(`⚠️ Falta el secret ${nombre} (Deno.env.get devolvió vacío/undefined).`);
}

// --- Helper: consultas a Postgres vía REST, con la service_role key ---
async function db(path: string, options: RequestInit = {}) {
  const res = await fetch(`${SUPABASE_URL}/rest/v1/${path}`, {
    ...options,
    headers: {
      apikey: SERVICE_ROLE_KEY!,
      Authorization: `Bearer ${SERVICE_ROLE_KEY}`,
      "Content-Type": "application/json",
      ...(options.headers || {}),
    },
  });
  const raw = await res.text();
  if (!res.ok) throw new Error(`DB ${path} → ${res.status}: ${raw}`);
  if (!raw) return null;
  try {
    return JSON.parse(raw);
  } catch {
    console.error(`DB ${path}: la respuesta no fue JSON válido:`, raw);
    throw new Error(`DB ${path}: respuesta no-JSON`);
  }
}

// Meta a veces guarda mal los destinatarios de prueba argentinos: en vez del
// formato real que usa WhatsApp (54 9 11 NNNNNNNN), les queda guardado el
// formato viejo (54 11 15 NNNNNNNN). Si el envío es rechazado por eso
// (código 131030) y el número es de Buenos Aires, reintentamos una vez con
// el formato alternativo antes de darnos por vencidos.
function formatoAlternativoBA(numero: string): string | null {
  const m = numero.match(/^549(11)(\d{8})$/); // 54 9 11 + 8 dígitos
  if (!m) return null;
  return `54${m[1]}15${m[2]}`; // 54 11 15 + los mismos 8 dígitos
}

// --- Helper: mandar un mensaje de texto por WhatsApp ---
async function enviarWhatsapp(to: string, texto: string) {
  async function intentar(destino: string) {
    const res = await fetch(`https://graph.facebook.com/v20.0/${WHATSAPP_PHONE_ID}/messages`, {
      method: "POST",
      headers: { Authorization: `Bearer ${WHATSAPP_TOKEN}`, "Content-Type": "application/json" },
      body: JSON.stringify({ messaging_product: "whatsapp", to: destino, type: "text", text: { body: texto } }),
    });
    const raw = await res.text();
    return { ok: res.ok, status: res.status, raw };
  }

  let resultado = await intentar(to);
  if (!resultado.ok) {
    console.error(`enviarWhatsapp → Meta rechazó el envío a "${to}" (status ${resultado.status}):`, resultado.raw);
    const esRechazoDeLista = resultado.raw.includes("131030");
    const alternativo = esRechazoDeLista ? formatoAlternativoBA(to) : null;
    if (alternativo) {
      console.log(`Reintentando con formato alternativo de Buenos Aires: "${alternativo}"`);
      resultado = await intentar(alternativo);
      if (!resultado.ok) {
        console.error(`enviarWhatsapp → también falló con "${alternativo}" (status ${resultado.status}):`, resultado.raw);
      } else {
        console.log(`enviarWhatsapp → OK con formato alternativo "${alternativo}":`, resultado.raw);
      }
    }
  } else {
    console.log("enviarWhatsapp → OK:", resultado.raw);
  }
}

// --- Helper: descargar y transcribir una nota de voz de WhatsApp ---
async function transcribirAudioWhatsapp(mediaId: string): Promise<string> {
  const metaRes = await fetch(`https://graph.facebook.com/v20.0/${mediaId}`, {
    headers: { Authorization: `Bearer ${WHATSAPP_TOKEN}` },
  });
  const metaRaw = await metaRes.text();
  if (!metaRes.ok) throw new Error(`Meta media info → ${metaRes.status}: ${metaRaw}`);
  const meta = JSON.parse(metaRaw);

  const audioRes = await fetch(meta.url, { headers: { Authorization: `Bearer ${WHATSAPP_TOKEN}` } });
  const blob = await audioRes.blob();

  const form = new FormData();
  form.append("file", blob, "audio.ogg");
  form.append("model", "whisper-1");
  form.append("language", "es");
  const whisperRes = await fetch("https://api.openai.com/v1/audio/transcriptions", {
    method: "POST",
    headers: { Authorization: `Bearer ${OPENAI_API_KEY}` },
    body: form,
  });
  const whisperRaw = await whisperRes.text();
  if (!whisperRes.ok) throw new Error(`Whisper → ${whisperRes.status}: ${whisperRaw}`);
  const data = JSON.parse(whisperRaw);
  return data.text || "";
}

// --- Helper: arma un snapshot resumido (para que la IA entienda qué datos
// hay disponibles) a partir de los entries ya traídos ---
function armarSnapshot(entries: any[]) {
  const hoy = new Date();
  const mesKey = hoy.toISOString().slice(0, 7);
  const gastos = entries.filter((e: any) => e.type === "gasto" && (e.moneda || "ARS") === "ARS");
  const ingresos = entries.filter((e: any) => e.type === "ingreso" || e.type === "cambio");
  const totalGastos = gastos.reduce((s: number, e: any) => s + Number(e.amount), 0);
  const totalIngresos = ingresos.reduce((s: number, e: any) => s + Number(e.amount), 0);
  const porCategoria: Record<string, number> = {};
  gastos.forEach((e: any) => { porCategoria[e.category] = (porCategoria[e.category] || 0) + Number(e.amount); });

  return {
    mes: mesKey,
    total_gastos: totalGastos,
    total_ingresos: totalIngresos,
    balance: totalIngresos - totalGastos,
    gastos_por_categoria: porCategoria,
  };
}

function fmtARS(n: number) {
  return `$${Math.round(n).toLocaleString("es-AR")}`;
}

// Arma el desglose "Nombre: $monto, Nombre2: $monto2" solo si hay más de
// una persona involucrada — si no, deja el texto limpio sin desglosar.
function desgloseTexto(lista: any[]) {
  const porPersona: Record<string, number> = {};
  lista.forEach((e) => {
    const quien = e.who || "(sin nombre)";
    porPersona[quien] = (porPersona[quien] || 0) + Number(e.amount);
  });
  const personas = Object.keys(porPersona);
  if (personas.length <= 1) return "";
  return " (" + personas.map((p) => `${p}: ${fmtARS(porPersona[p])}`).join(", ") + ")";
}

// --- La suma y el filtrado los hace el código, con matemática exacta —
// nunca le pedimos a la IA que sume una lista, porque no es 100% confiable
// haciendo aritmética sobre texto. La IA solo dice QUÉ se preguntó
// (consulta.tipo/valor); esta función calcula la respuesta real. ---
function responderConsulta(consulta: any, entries: any[]): string {
  const gastos = entries.filter((e: any) => e.type === "gasto" && (e.moneda || "ARS") === "ARS");
  const ingresos = entries.filter((e: any) => e.type === "ingreso" || e.type === "cambio");

  if (consulta?.tipo === "balance") {
    const tg = gastos.reduce((s, e) => s + Number(e.amount), 0);
    const ti = ingresos.reduce((s, e) => s + Number(e.amount), 0);
    return `Balance del mes: ${fmtARS(ti - tg)} (ingresos ${fmtARS(ti)}, gastos ${fmtARS(tg)}).`;
  }
  if (consulta?.tipo === "total_gastos") {
    const total = gastos.reduce((s, e) => s + Number(e.amount), 0);
    return `Gastaron ${fmtARS(total)} este mes${desgloseTexto(gastos)}.`;
  }
  if (consulta?.tipo === "total_ingresos") {
    const total = ingresos.reduce((s, e) => s + Number(e.amount), 0);
    return `Ingresaron ${fmtARS(total)} este mes${desgloseTexto(ingresos)}.`;
  }
  if (consulta?.tipo === "busqueda" && consulta.valor) {
    const valor = String(consulta.valor).toLowerCase();
    const filtrados = gastos.filter((e) =>
      (e.category || "").toLowerCase() === valor || (e.descripcion || "").toLowerCase().includes(valor)
    );
    if (filtrados.length === 0) return `No encontré gastos de "${consulta.valor}" este mes.`;
    const total = filtrados.reduce((s, e) => s + Number(e.amount), 0);
    return `Gastaron ${fmtARS(total)} en ${consulta.valor} este mes${desgloseTexto(filtrados)}.`;
  }
  return ""; // no reconocida — el llamador usa el fallback de la IA
}

// --- Helper: le pide a OpenAI que clasifique el mensaje (nunca que sume) ---
async function interpretarMensaje(texto: string, categorias: string[], snapshot: any) {
  const systemPrompt = `Sos el clasificador de una app de finanzas familiares en Argentina (pesos ARS). Te llega un mensaje de WhatsApp que puede ser:
1) Una carga de gasto/ingreso (ej: "gasté 5000 en nafta", "cobré 200000 de sueldo").
2) Una pregunta sobre los datos del mes (ej: "cuánto gasté en comida?", "cuál es mi balance?", "cuánto gasté en nafta?").
3) Otra cosa (saludo, algo que no entendés).

Categorías válidas para gastos: ${categorias.join(", ")}.
Elegí la categoría que mejor represente el SIGNIFICADO del gasto, no solo si la palabra exacta aparece en el nombre de la categoría — por ejemplo, nafta/combustible/estacionamiento/peaje/Uber van en "Transporte" si esa categoría existe; supermercado/restaurante/delivery van en "Comida"; etc. Usá "Otros" únicamente cuando el gasto de verdad no encaje en ninguna categoría de la lista, no como opción por defecto.

Resumen del mes (solo para que entiendas qué se puede preguntar, NO hagas cuentas vos): ${JSON.stringify(snapshot)}.

IMPORTANTE: si intent es "pregunta", vos NO calculás el número — solo identificás qué tipo de pregunta es, en "consulta":
- {"tipo": "balance"} → preguntan el balance general.
- {"tipo": "total_gastos"} → preguntan cuánto gastaron en total.
- {"tipo": "total_ingresos"} → preguntan cuánto ingresó en total.
- {"tipo": "busqueda", "valor": "lo que preguntan"} → preguntan por algo específico: una categoría (ej. "Transporte"), un concepto (ej. "nafta", "psicóloga", "supermercado"), lo que sea — poné en "valor" la palabra tal cual la usaron, sin traducirla a categoría ni modificarla.
- {"tipo": "otro"} → no encaja en nada de lo anterior.

Respondé SOLO con JSON, sin texto adicional, con este formato exacto:
{"intent": "carga" | "pregunta" | "otro", "entry": {"type": "gasto"|"ingreso", "amount": number, "category": string, "desc": string} | null, "consulta": {"tipo": string, "valor": string | null} | null, "respuesta_otro": string | null}

Si intent es "carga", completá "entry" y dejá "consulta" y "respuesta_otro" en null.
Si intent es "pregunta", completá "consulta" y dejá "entry" y "respuesta_otro" en null.
Si intent es "otro", dejá "entry" y "consulta" en null, y poné en "respuesta_otro" una aclaración breve de cómo usar el bot.`;

  const res = await fetch("https://api.openai.com/v1/chat/completions", {
    method: "POST",
    headers: { Authorization: `Bearer ${OPENAI_API_KEY}`, "Content-Type": "application/json" },
    body: JSON.stringify({
      model: "gpt-4o-mini",
      response_format: { type: "json_object" },
      messages: [
        { role: "system", content: systemPrompt },
        { role: "user", content: texto },
      ],
    }),
  });
  const raw = await res.text();
  if (!res.ok) {
    console.error("OpenAI devolvió un error:", res.status, raw);
    throw new Error(`OpenAI → ${res.status}: ${raw}`);
  }
  let data;
  try {
    data = JSON.parse(raw);
  } catch {
    console.error("La respuesta de OpenAI no fue JSON válido:", raw);
    throw new Error("OpenAI: respuesta no-JSON");
  }

  const contenido = data?.choices?.[0]?.message?.content;
  if (!contenido) {
    console.error("OpenAI no devolvió contenido esperado. Respuesta completa:", JSON.stringify(data));
    throw new Error("OpenAI: sin contenido en la respuesta");
  }
  try {
    return JSON.parse(contenido);
  } catch {
    console.error("El contenido de OpenAI no fue JSON válido:", contenido);
    throw new Error("OpenAI: contenido no-JSON");
  }
}

serve(async (req) => {
  const url = new URL(req.url);

  // --- Verificación del webhook (Meta hace un GET la primera vez) ---
  if (req.method === "GET") {
    const mode = url.searchParams.get("hub.mode");
    const token = url.searchParams.get("hub.verify_token");
    const challenge = url.searchParams.get("hub.challenge");
    if (mode === "subscribe" && token === VERIFY_TOKEN) {
      return new Response(challenge, { status: 200 });
    }
    return new Response("Forbidden", { status: 403 });
  }

  // --- Mensaje entrante ---
  try {
    const rawBody = await req.text();
    if (!rawBody) {
      console.log("POST sin body (probablemente un ping de verificación de Meta) — se ignora.");
      return new Response("ok", { headers: CORS_HEADERS });
    }

    let body;
    try {
      body = JSON.parse(rawBody);
    } catch {
      console.error("Body recibido no es JSON válido:", rawBody);
      return new Response("ok", { headers: CORS_HEADERS });
    }

    console.log("Body recibido:", JSON.stringify(body));

    const value = body?.entry?.[0]?.changes?.[0]?.value;
    const mensaje = value?.messages?.[0];
    if (!mensaje) {
      console.log("No hay 'messages' en el payload (probablemente una notificación de estado, no un mensaje) — se ignora.");
      return new Response("ok", { headers: CORS_HEADERS });
    }

    const from = mensaje.from; // ej: "5491122334455"

    // 1) ¿De qué hogar es este número?
    const links = await db(`whatsapp_links?phone_number=eq.${from}&select=*`);
    if (!links || links.length === 0) {
      console.log(`No hay whatsapp_links para el número "${from}" — pidiendo que lo vinculen.`);
      await enviarWhatsapp(from, "Todavía no vinculé tu número a ningún hogar. Entrá a la app → Mi hogar → WhatsApp, y agregá este número primero.");
      return new Response("ok", { headers: CORS_HEADERS });
    }
    const link = links[0];
    console.log(`Número vinculado a household_id=${link.household_id}, display_name=${link.display_name}`);

    // 2) Texto del mensaje (directo, o transcripto si es audio)
    let texto = "";
    if (mensaje.type === "text") {
      texto = mensaje.text.body;
    } else if (mensaje.type === "audio") {
      texto = await transcribirAudioWhatsapp(mensaje.audio.id);
    } else {
      await enviarWhatsapp(from, "Por ahora solo entiendo texto o notas de voz.");
      return new Response("ok", { headers: CORS_HEADERS });
    }
    console.log("Texto a interpretar:", texto);

    // 3) Categorías del hogar + datos del mes (una sola consulta, se reutiliza)
    const catRows = await db(`categories?household_id=eq.${link.household_id}&select=name`);
    const categorias = (catRows || []).map((c: any) => c.name);
    const mesKey = new Date().toISOString().slice(0, 7);
    const entries = await db(
      `entries?household_id=eq.${link.household_id}&date=gte.${mesKey}-01&select=type,category,amount,descripcion,date,moneda,who&order=date.desc&limit=300`
    ) || [];
    const snapshot = armarSnapshot(entries);

    // 4) Interpretar con OpenAI (clasifica, no suma)
    const resultado = await interpretarMensaje(texto, categorias, snapshot);
    console.log("Resultado de OpenAI:", JSON.stringify(resultado));

    if (resultado.intent === "carga" && resultado.entry) {
      const categoriaFinal = categorias.includes(resultado.entry.category) ? resultado.entry.category : "Otros";
      await db("entries", {
        method: "POST",
        body: JSON.stringify([{
          id: crypto.randomUUID(),
          household_id: link.household_id,
          type: resultado.entry.type,
          category: categoriaFinal,
          amount: Number(resultado.entry.amount),
          descripcion: resultado.entry.desc || texto,
          date: new Date().toISOString().slice(0, 10),
          who: link.display_name,
          moneda: "ARS",
          origen: "whatsapp",
        }]),
      });
      const signo = resultado.entry.type === "gasto" ? "Gasto" : "Ingreso";
      await enviarWhatsapp(from, `✅ ${signo} cargado: $${Number(resultado.entry.amount).toLocaleString("es-AR")} en ${categoriaFinal}.`);
    } else if (resultado.intent === "pregunta" && resultado.consulta) {
      const respuestaExacta = responderConsulta(resultado.consulta, entries);
      await enviarWhatsapp(from, respuestaExacta || "No entendí bien qué querés saber. Probá algo como \"cuánto gasté en comida?\" o \"cuál es mi balance?\".");
    } else {
      await enviarWhatsapp(from, resultado.respuesta_otro || "No entendí bien. Probá algo como \"gasté 5000 en nafta\" o \"cuánto gasté este mes en comida\".");
    }

    return new Response("ok", { headers: CORS_HEADERS });
  } catch (e) {
    console.error(e);
    return new Response("ok", { headers: CORS_HEADERS }); // igual 200, para que Meta no reintente en loop
  }
});
