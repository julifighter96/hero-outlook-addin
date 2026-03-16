/**
 * Vercel Serverless Proxy – leitet GraphQL-Requests an HERO API weiter.
 * Umgeht das CORS-Problem, da der Request vom Server ausgeht.
 */
export default async function handler(req, res) {
  // CORS für das Add-In erlauben
  res.setHeader("Access-Control-Allow-Origin", "*");
  res.setHeader("Access-Control-Allow-Methods", "POST, OPTIONS");
  res.setHeader("Access-Control-Allow-Headers", "Content-Type, Authorization");

  if (req.method === "OPTIONS") {
    return res.status(200).end();
  }

  if (req.method !== "POST") {
    return res.status(405).json({ error: "Method not allowed" });
  }

  const authHeader = req.headers["authorization"];
  if (!authHeader) {
    return res.status(401).json({ error: "Authorization header missing" });
  }

  try {
    const heroRes = await fetch("https://login.hero-software.de/api/external/v7/graphql", {
      method: "POST",
      headers: {
        "Authorization": authHeader,
        "Content-Type": "application/json",
      },
      body: JSON.stringify(req.body),
    });

    const data = await heroRes.json();
    return res.status(heroRes.status).json(data);
  } catch (err) {
    return res.status(502).json({ error: "Upstream request failed", detail: err.message });
  }
}
