const APPS_SCRIPT_URL = "https://script.google.com/macros/s/AKfycbyTqHC1as7ahGbG14_r5WJq3Dkg10SHGYT65Kr58adf-3wnvCokYPelSy4YejoVANUT/exec";

export default async function handler(req, res) {
  if (req.method === "OPTIONS") {
    res.status(204).end();
    return;
  }

  if (req.method !== "POST" && req.method !== "GET") {
    res.status(405).json({ ok: false, error: "Method not allowed" });
    return;
  }

  try {
    if (req.method === "GET") {
      const r = await fetch(APPS_SCRIPT_URL, { method: "GET" });
      const text = await r.text();
      res.status(r.status).send(text);
      return;
    }

    const body = typeof req.body === "string" ? req.body : JSON.stringify(req.body || {});
    const r = await fetch(APPS_SCRIPT_URL, {
      method: "POST",
      headers: { "Content-Type": "text/plain;charset=utf-8" },
      body
    });

    const text = await r.text();
    res.status(r.status).setHeader("Content-Type", "application/json; charset=utf-8").send(text);
  } catch (error) {
    res.status(500).json({ ok: false, error: `Proxy error: ${error?.message || error}` });
  }
}
