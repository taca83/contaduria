// supabase/functions/analizar-recibo/index.ts
//
// Recibe una foto de un ticket/comprobante (base64) y le pide a un modelo
// de IA con visión (OpenAI) que extraiga monto, fecha, categoría y detalle.
// Reutiliza la misma OPENAI_API_KEY que ya está configurada para la
// transcripción de audio — no hace falta agregar ningún secret nuevo.
//
// La IA nunca carga nada sola: el cliente siempre muestra los datos
// extraídos en el formulario para que la persona los confirme o corrija
// antes de guardar.

import { serve } from "https://deno.land/std@0.224.0/http/server.ts";

const OPENAI_API_KEY = Deno.env.get("OPENAI_API_KEY");
const CORS_HEADERS = {
  "Access-Control-Allow-Origin": "*",
  "Access-Control-Allow-Headers": "authorization, x-client-info, apikey, content-type",
};

serve(async (req) => {
  if (req.method === "OPTIONS") {
    return new Response("ok", { headers: CORS_HEADERS });
  }

  if (!OPENAI_API_KEY) {
    console.error("⚠️ Falta el secret OPENAI_API_KEY.");
  }

  try {
    const rawBody = await req.text();
    if (!rawBody) {
      return new Response(JSON.stringify({ error: "Body vacío." }), {
        status: 400,
        headers: { ...CORS_HEADERS, "Content-Type": "application/json" },
      });
    }

    let body;
    try {
      body = JSON.parse(rawBody);
    } catch {
      console.error("Body no es JSON válido:", rawBody.slice(0, 200));
      return new Response(JSON.stringify({ error: "Body no es JSON válido." }), {
        status: 400,
        headers: { ...CORS_HEADERS, "Content-Type": "application/json" },
      });
    }

    const { image_base64, mime_type, categorias } = body;
    if (!image_base64) {
      return new Response(JSON.stringify({ error: "Falta la imagen." }), {
        status: 400,
        headers: { ...CORS_HEADERS, "Content-Type": "application/json" },
      });
    }

    const dataUrl = `data:${mime_type || "image/jpeg"};base64,${image_base64}`;
    const hoy = new Date().toISOString().slice(0, 10);
    const listaCategorias = Array.isArray(categorias) ? categorias : [];

    const systemPrompt = `Sos un asistente que lee fotos de tickets/recibos de compra en Argentina (montos en pesos ARS) y extrae los datos del gasto para cargarlo en una app de finanzas familiares.

Categorías válidas: ${listaCategorias.join(", ")}.
Elegí la categoría que mejor represente el gasto por su significado (ej. un ticket de supermercado va en "Comida", una factura de nafta en "Transporte", una farmacia en "Salud"). Usá "Otros" solo si de verdad no encaja en ninguna categoría de la lista.

Si el ticket muestra una fecha, usala en formato YYYY-MM-DD. Si no se ve ninguna fecha, usá "${hoy}" (hoy).
Si no podés leer el monto TOTAL con claridad, poné "monto": null en vez de inventar un número.
"comercio" es el nombre del negocio/local si se ve (ej. "Carrefour", "YPF"). "desc" es cualquier detalle extra corto y útil (ej. "2 productos", vacío si no hay nada relevante).

Respondé SOLO con este JSON, sin texto adicional ni backticks:
{"monto": number | null, "fecha": "YYYY-MM-DD", "categoria": string, "comercio": string, "desc": string}`;

    const res = await fetch("https://api.openai.com/v1/chat/completions", {
      method: "POST",
      headers: { Authorization: `Bearer ${OPENAI_API_KEY}`, "Content-Type": "application/json" },
      body: JSON.stringify({
        model: "gpt-4o-mini",
        response_format: { type: "json_object" },
        messages: [
          { role: "system", content: systemPrompt },
          {
            role: "user",
            content: [
              { type: "text", text: "Leé este recibo y extraé los datos del gasto." },
              { type: "image_url", image_url: { url: dataUrl } },
            ],
          },
        ],
      }),
    });

    const raw = await res.text();
    if (!res.ok) {
      console.error("OpenAI devolvió un error:", res.status, raw);
      return new Response(JSON.stringify({ error: `OpenAI → ${res.status}` }), {
        status: 500,
        headers: { ...CORS_HEADERS, "Content-Type": "application/json" },
      });
    }

    let data;
    try {
      data = JSON.parse(raw);
    } catch {
      console.error("Respuesta de OpenAI no fue JSON válido:", raw);
      return new Response(JSON.stringify({ error: "OpenAI: respuesta no-JSON" }), {
        status: 500,
        headers: { ...CORS_HEADERS, "Content-Type": "application/json" },
      });
    }

    const contenido = data?.choices?.[0]?.message?.content;
    if (!contenido) {
      console.error("OpenAI no devolvió contenido esperado:", JSON.stringify(data));
      return new Response(JSON.stringify({ error: "OpenAI: sin contenido en la respuesta" }), {
        status: 500,
        headers: { ...CORS_HEADERS, "Content-Type": "application/json" },
      });
    }

    let resultado;
    try {
      resultado = JSON.parse(contenido);
    } catch {
      console.error("Contenido de OpenAI no fue JSON válido:", contenido);
      return new Response(JSON.stringify({ error: "OpenAI: contenido no-JSON" }), {
        status: 500,
        headers: { ...CORS_HEADERS, "Content-Type": "application/json" },
      });
    }

    return new Response(JSON.stringify(resultado), {
      headers: { ...CORS_HEADERS, "Content-Type": "application/json" },
    });
  } catch (e) {
    console.error(e);
    return new Response(JSON.stringify({ error: e.message || "Error inesperado" }), {
      status: 500,
      headers: { ...CORS_HEADERS, "Content-Type": "application/json" },
    });
  }
});
