/**
 * Vercel Serverless Proxy
 *
 * POST /api/hero          → HERO GraphQL
 * POST /api/hero?upload=1 → HERO File-Upload REST endpoint (gibt UUID zurück)
 */
export default async function handler(req, res) {
  res.setHeader("Access-Control-Allow-Origin", "*");
  res.setHeader("Access-Control-Allow-Methods", "POST, OPTIONS");
  res.setHeader("Access-Control-Allow-Headers", "Content-Type, Authorization");

  if (req.method === "OPTIONS") return res.status(200).end();
  if (req.method !== "POST") return res.status(405).json({ error: "Method not allowed" });

  const authHeader = req.headers["authorization"];
  if (!authHeader) return res.status(401).json({ error: "Authorization header missing" });

  // ── File-Upload (Base64 → HERO REST → UUID) ──────────────────────────────
  if (req.query.upload === "1") {
    const { filename, content_base64, content_type } = req.body;

    // Base64 → Buffer
    const buffer = Buffer.from(content_base64, "base64");

    // Multipart-Body manuell aufbauen
    const boundary = "----HEROBoundary" + Date.now();
    const CRLF = "\r\n";
    const header =
      `--${boundary}${CRLF}` +
      `Content-Disposition: form-data; name="file"; filename="${filename}"${CRLF}` +
      `Content-Type: ${content_type}${CRLF}${CRLF}`;
    const footer = `${CRLF}--${boundary}--${CRLF}`;

    const body = Buffer.concat([
      Buffer.from(header, "utf8"),
      buffer,
      Buffer.from(footer, "utf8"),
    ]);

    try {
      const uploadRes = await fetch(
        "https://login.hero-software.de/api/external/v7/file_uploads",
        {
          method: "POST",
          headers: {
            Authorization: authHeader,
            "Content-Type": `multipart/form-data; boundary=${boundary}`,
          },
          body,
        }
      );

      const text = await uploadRes.text();
      let data;
      try { data = JSON.parse(text); } catch { data = { raw: text }; }

      return res.status(uploadRes.status).json(data);
    } catch (err) {
      return res.status(502).json({ error: "Upload fehlgeschlagen", detail: err.message });
    }
  }

  // ── GraphQL Proxy ─────────────────────────────────────────────────────────
  try {
    const heroRes = await fetch(
      "https://login.hero-software.de/api/external/v7/graphql",
      {
        method: "POST",
        headers: {
          Authorization: authHeader,
          "Content-Type": "application/json",
        },
        body: JSON.stringify(req.body),
      }
    );

    const data = await heroRes.json();
    return res.status(heroRes.status).json(data);
  } catch (err) {
    return res.status(502).json({ error: "GraphQL-Anfrage fehlgeschlagen", detail: err.message });
  }
}
