// supabase/functions/transcribir-audio/index.ts
import { serve } from "https://deno.land/std@0.224.0/http/server.ts";

const CORS_HEADERS = {
  "Access-Control-Allow-Origin": "*",
  "Access-Control-Allow-Headers": "authorization, x-client-info, apikey, content-type",
};

serve(async (req) => {
  if (req.method === "OPTIONS") {
    return new Response("ok", { headers: CORS_HEADERS });
  }

  try {
    const { audio_base64, mime_type } = await req.json();
    if (!audio_base64) {
      return new Response(JSON.stringify({ error: "Falta el audio." }), {
        status: 400,
        headers: { ...CORS_HEADERS, "Content-Type": "application/json" },
      });
    }

    const bytes = Uint8Array.from(atob(audio_base64), (c) => c.charCodeAt(0));
    const tipo = mime_type || "audio/webm";
    const extension = tipo.includes("mp4") ? "mp4" : "webm";
    const blob = new Blob([bytes], { type: tipo });

    const form = new FormData();
    form.append("file", blob, `audio.${extension}`);
    form.append("model", "whisper-1");
    form.append("language", "es");

    const resp = await fetch("https://api.openai.com/v1/audio/transcriptions", {
      method: "POST",
      headers: { Authorization: `Bearer ${Deno.env.get("OPENAI_API_KEY")}` },
      body: form,
    });
    const data = await resp.json();

    if (!resp.ok) {
      return new Response(JSON.stringify({ error: data?.error?.message || "Error al transcribir el audio." }), {
        status: 500,
        headers: { ...CORS_HEADERS, "Content-Type": "application/json" },
      });
    }

    return new Response(JSON.stringify({ text: data.text }), {
      headers: { ...CORS_HEADERS, "Content-Type": "application/json" },
    });
  } catch (e) {
    return new Response(JSON.stringify({ error: e.message || "Error inesperado." }), {
      status: 500,
      headers: { ...CORS_HEADERS, "Content-Type": "application/json" },
    });
  }
});
