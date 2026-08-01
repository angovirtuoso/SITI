/**
 * Proxy serverless para Vercel.
 * Configura APPS_SCRIPT_URL en Vercel con la URL /exec del Web App de Apps Script.
 */
export default async function handler(req, res) {
  const scriptUrl = process.env.APPS_SCRIPT_URL;
  if (!scriptUrl) {
    return res.status(500).json({ ok: false, message: 'Falta configurar APPS_SCRIPT_URL en Vercel.' });
  }

  const url = new URL(scriptUrl);
  Object.entries(req.query || {}).forEach(([key, value]) => {
    if (key !== 'path') url.searchParams.set(key, Array.isArray(value) ? value[0] : value);
  });

  try {
    const options = { method: req.method, redirect: 'follow' };
    if (req.method !== 'GET' && req.method !== 'HEAD') {
      options.headers = { 'Content-Type': 'text/plain;charset=utf-8' };
      options.body = typeof req.body === 'string' ? req.body : JSON.stringify(req.body || {});
    }
    const response = await fetch(url, options);
    const body = await response.text();
    res.status(response.ok ? 200 : response.status);
    res.setHeader('Content-Type', 'application/json; charset=utf-8');
    return res.send(body);
  } catch (error) {
    return res.status(502).json({ ok: false, message: `No se pudo conectar con Apps Script: ${error.message}` });
  }
}

