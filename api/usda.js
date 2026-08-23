export default async function handler(request, response) {
  if (request.method !== "GET") {
    response.status(405).json({ error: "Method not allowed" });
    return;
  }

  const query = String(request.query.query || "").trim();
  if (!query) {
    response.status(400).json({ error: "query is required" });
    return;
  }

  const apiKey = process.env.USDA_API_KEY || process.env.VITE_USDA_API_KEY;
  if (!apiKey) {
    response.status(503).json({ error: "USDA API key is not configured." });
    return;
  }

  const upstreamUrl = new URL(
    "https://api.nal.usda.gov/fdc/v1/foods/search/"
  );
  upstreamUrl.searchParams.set("query", query);
  upstreamUrl.searchParams.set("pageSize", "10");
  upstreamUrl.searchParams.set("api_key", apiKey);

  try {
    const upstreamResponse = await fetch(upstreamUrl);
    const body = await upstreamResponse.text();

    response.status(upstreamResponse.status);
    response.setHeader(
      "Content-Type",
      upstreamResponse.headers.get("content-type") || "application/json"
    );
    response.send(body);
  } catch (error) {
    console.error("USDA API failed:", error);
    response.status(502).json({ error: "USDA search is unavailable." });
  }
}
