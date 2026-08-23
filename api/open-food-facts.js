export default async function handler(request, response) {
  if (request.method !== "GET") {
    response.status(405).json({ error: "Method not allowed" });
    return;
  }

  const searchTerms = String(request.query.search_terms || "").trim();
  if (!searchTerms) {
    response.status(400).json({ error: "search_terms is required" });
    return;
  }

  const upstreamUrl = new URL("https://api.openfoodfacts.org/api/v2/search");
  upstreamUrl.searchParams.set("search_terms", searchTerms);
  upstreamUrl.searchParams.set("page_size", "10");
  upstreamUrl.searchParams.set(
    "fields",
    "code,product_name,brands,quantity,nutriments"
  );

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
    console.error("Open Food Facts API failed:", error);
    response.status(502).json({
      error: "Open Food Facts search is unavailable.",
    });
  }
}
