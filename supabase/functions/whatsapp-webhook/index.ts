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
  if (!res.ok) throw new Error(`DB ${path} → ${res.status}: ${await res.text()}`);
  if (res.status === 204) return null;
  return res.json();
}

// --- Helper: mandar un mensaje de texto por WhatsApp ---
async function enviarWhatsapp(to: string, texto: string) {
  await fetch(`https://graph.facebook.com/v20.0/${WHATSAPP_PHONE_ID}/messages`, {
    method: "POST",
    headers: { Authorization: `Bearer ${WHATSAPP_TOKEN}`, "Content-Type": "application/json" },
    body: JSON.stringify({ messaging_product: "whatsapp", to, type: "text", text: { body: texto } }),
  });
}

// --- Helper: descargar y transcribir una nota de voz de WhatsApp ---
async function transcribirAudioWhatsapp(mediaId: string): Promise<string> {
  const metaRes = await fetch(`https://graph.facebook.com/v20.0/${mediaId}`, {
    headers: { Authorization: `Bearer ${WHATSAPP_TOKEN}` },
  });
  const meta = await metaRes.json();
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
  const data = await whisperRes.json();
  return data.text || "";
}

// --- Helper: armar un snapshot de los datos del mes para que el modelo
// pueda responder preguntas sin inventar números ---
async function snapshotDelMes(householdId: string) {
  const hoy = new Date();
  const mesKey = hoy.toISOString().slice(0, 7);
  const entries = await db(
    `entries?household_id=eq.${householdId}&date=gte.${mesKey}-01&select=type,category,amount,descripcion,date,moneda&order=date.desc&limit=300`
  );
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
    ultimos_movimientos: entries.slice(0, 15).map((e: any) => ({
      fecha: e.date, tipo: e.type, categoria: e.category, monto: Number(e.amount), desc: e.descripcion,
    })),
  };
}

// --- Helper: le pide a OpenAI que clasifique el mensaje y responda ---
async function interpretarMensaje(texto: string, categorias: string[], snapshot: any) {
  const systemPrompt = `Sos el asistente de una app de finanzas familiares en Argentina (pesos ARS). Te llega un mensaje de WhatsApp que puede ser:
1) Una carga de gasto/ingreso (ej: "gasté 5000 en nafta", "cobré 200000 de sueldo").
2) Una pregunta sobre los datos del mes (ej: "cuánto gasté en comida?", "cuál es mi balance?").
3) Otra cosa (saludo, algo que no entendés).

Categorías válidas para gastos: ${categorias.join(", ")}.
Datos reales del mes actual (usalos para responder preguntas, NUNCA inventes números): ${JSON.stringify(snapshot)}.

Respondé SOLO con JSON, sin texto adicional, con este formato exacto:
{"intent": "carga" | "pregunta" | "otro", "entry": {"type": "gasto"|"ingreso", "amount": number, "category": string, "desc": string} | null, "respuesta": string | null}

Si intent es "carga", completá "entry" y dejá "respuesta" en null.
Si intent es "pregunta", completá "respuesta" (en español, corta, con los montos reales del snapshot) y dejá "entry" en null.
Si intent es "otro", dejá "entry" en null y poné en "respuesta" una aclaración breve de cómo usar el bot.`;

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
  const data = await res.json();
  return JSON.parse(data.choices[0].message.content);
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
    const body = await req.json();
    const value = body?.entry?.[0]?.changes?.[0]?.value;
    const mensaje = value?.messages?.[0];
    if (!mensaje) return new Response("ok", { headers: CORS_HEADERS }); // notificación de estado, no un mensaje

    const from = mensaje.from; // ej: "5491122334455"

    // 1) ¿De qué hogar es este número?
    const links = await db(`whatsapp_links?phone_number=eq.${from}&select=*`);
    if (!links || links.length === 0) {
      await enviarWhatsapp(from, "Todavía no vinculé tu número a ningún hogar. Entrá a la app → Mi hogar → WhatsApp, y agregá este número primero.");
      return new Response("ok", { headers: CORS_HEADERS });
    }
    const link = links[0];

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

    // 3) Categorías del hogar + snapshot del mes
    const catRows = await db(`categories?household_id=eq.${link.household_id}&select=name`);
    const categorias = (catRows || []).map((c: any) => c.name);
    const snapshot = await snapshotDelMes(link.household_id);

    // 4) Interpretar con OpenAI
    const resultado = await interpretarMensaje(texto, categorias, snapshot);

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
        }]),
      });
      const signo = resultado.entry.type === "gasto" ? "Gasto" : "Ingreso";
      await enviarWhatsapp(from, `✅ ${signo} cargado: $${Number(resultado.entry.amount).toLocaleString("es-AR")} en ${categoriaFinal}.`);
    } else {
      await enviarWhatsapp(from, resultado.respuesta || "No entendí bien. Probá algo como \"gasté 5000 en nafta\" o \"cuánto gasté este mes en comida\".");
    }

    return new Response("ok", { headers: CORS_HEADERS });
  } catch (e) {
    console.error(e);
    return new Response("ok", { headers: CORS_HEADERS }); // igual 200, para que Meta no reintente en loop
  }
});
