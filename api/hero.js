/**
 * Vercel Serverless Proxy
 *
 * POST /api/hero            → HERO GraphQL v7 (Suche, Logbuch)
 * POST /api/hero?upload=1   → HERO REST v8 FileUploads (Datei → UUID)
 * POST /api/hero?v8=1       → HERO GraphQL v8 (upload_document)
 */
export default async function handler(req, res) {
  res.setHeader("Access-Control-Allow-Origin", "*");
  res.setHeader("Access-Control-Allow-Methods", "POST, OPTIONS");
  res.setHeader("Access-Control-Allow-Headers", "Content-Type, Authorization");

  if (req.method === "OPTIONS") return res.status(200).end();
  if (req.method !== "POST") return res.status(405).json({ error: "Method not allowed" });

  // API-Key aus Bearer-Header extrahieren
  const authHeader = req.headers["authorization"] || "";
  const apiKey = authHeader.replace(/^Bearer\s+/i, "").trim();
  if (!apiKey) return res.status(401).json({ error: "Authorization header missing" });

  // ── Datei-Upload: REST v8 → gibt UUID zurück ─────────────────────────────
  if (req.query.upload === "1") {
    const { filename, content_base64, content_type } = req.body;
    const buffer = Buffer.from(content_base64, "base64");

    // UUID selbst generieren (wie in der Python-Referenzimplementierung)
    const fileUuid = crypto.randomUUID();

    const boundary = "----HEROBoundary" + Date.now();
    const CRLF = "\r\n";

    // Multipart: field "uuid"
    const uuidPart =
      `--${boundary}${CRLF}` +
      `Content-Disposition: form-data; name="uuid"${CRLF}${CRLF}` +
      `${fileUuid}${CRLF}`;

    // Multipart: field "file"
    const filePart =
      `--${boundary}${CRLF}` +
      `Content-Disposition: form-data; name="file"; filename="${filename}"${CRLF}` +
      `Content-Type: ${content_type}${CRLF}${CRLF}`;

    const footer = `${CRLF}--${boundary}--${CRLF}`;

    const body = Buffer.concat([
      Buffer.from(uuidPart, "utf8"),
      Buffer.from(filePart, "utf8"),
      buffer,
      Buffer.from(footer, "utf8"),
    ]);

    try {
      const uploadRes = await fetch(
        "https://login.hero-software.de/app/v8/FileUploads/upload",
        {
          method: "POST",
          headers: {
            "x-auth-token": apiKey,
            "Content-Type": `multipart/form-data; boundary=${boundary}`,
          },
          body,
        }
      );

      const text = await uploadRes.text();
      console.log("File upload raw response:", text);
      // UUID zurückgeben – egal was HERO antwortet, wir kennen die UUID
      return res.status(uploadRes.status).json({ uuid: fileUuid, raw: text });
    } catch (err) {
      return res.status(502).json({ error: "Upload fehlgeschlagen", detail: err.message });
    }
  }

  // ── GraphQL v7 (Suche, Logbuch, Introspection) ────────────────────────────
  try {
    const heroRes = await fetch(
      "https://login.hero-software.de/api/external/v7/graphql",
      {
        method: "POST",
        headers: {
          "Authorization": `Bearer ${apiKey}`,
          "Content-Type": "application/json",
        },
        body: JSON.stringify(req.body),
      }
    );
    const data = await heroRes.json();
    return res.status(heroRes.status).json(data);
  } catch (err) {
    return res.status(502).json({ error: "GraphQL v7 fehlgeschlagen", detail: err.message });
  }
}
