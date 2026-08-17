// Función intermedia (Vercel Serverless Function): recibe pedidos del sitio y los
// reenvía a Google Apps Script desde el servidor, no desde el navegador.
//
// Por qué existe: llamar a Apps Script directo desde el navegador (fetch del lado
// del cliente) falla de forma intermitente por un problema de redirección/CORS —
// ya lo encontramos una vez migrando Control DJ. Reenviar desde acá (servidor a
// servidor) lo evita por completo.
//
// La URL real de Apps Script NUNCA queda expuesta en el código del sitio — vive
// solo acá, como variable de entorno del lado del servidor (APPS_SCRIPT_URL,
// sin el prefijo VITE_, así Vercel no la incluye en el bundle público).

export default async function handler(req, res) {
  const APPS_SCRIPT_URL = process.env.APPS_SCRIPT_URL;

  if (!APPS_SCRIPT_URL) {
    return res.status(500).json({ ok: false, msg: "Falta configurar APPS_SCRIPT_URL en Vercel." });
  }

  try {
    if (req.method === "GET") {
      const params = new URLSearchParams(req.query).toString();
      const upstream = await fetch(`${APPS_SCRIPT_URL}?${params}`, { redirect: "follow" });
      const data = await upstream.json();
      return res.status(200).json(data);
    }

    if (req.method === "POST") {
      const upstream = await fetch(APPS_SCRIPT_URL, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify(req.body),
        redirect: "follow",
      });
      const data = await upstream.json();
      return res.status(200).json(data);
    }

    return res.status(405).json({ ok: false, msg: "Método no soportado" });
  } catch (e) {
    return res.status(500).json({ ok: false, msg: String(e) });
  }
}
