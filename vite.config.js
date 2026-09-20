import { defineConfig } from "vite";
import { loadEnv } from "vite";
import react from "@vitejs/plugin-react";
import tailwindcss from "@tailwindcss/vite";
import { VitePWA } from "vite-plugin-pwa";

function usdaProxy(apiKey) {
  return {
    name: "usda-proxy",
    configureServer(server) {
      server.middlewares.use(async function (request, response, next) {
        if (!request.url || !request.url.startsWith("/api/usda")) {
          next();
          return;
        }

        if (!apiKey) {
          response.statusCode = 503;
          response.end(JSON.stringify({ error: "USDA API key is not configured." }));
          return;
        }

        try {
          const incomingUrl = new URL(request.url, "http://localhost");
          const upstreamUrl = new URL(
            "https://api.nal.usda.gov/fdc/v1/foods/search/"
          );

          incomingUrl.searchParams.forEach(function (value, key) {
            upstreamUrl.searchParams.set(key, value);
          });
          upstreamUrl.searchParams.set("api_key", apiKey);

          const upstreamResponse = await fetch(upstreamUrl);
          const body = await upstreamResponse.text();

          response.statusCode = upstreamResponse.status;
          response.setHeader(
            "Content-Type",
            upstreamResponse.headers.get("content-type") || "application/json"
          );
          response.end(body);
        } catch (error) {
          console.error("USDA proxy failed:", error);
          response.statusCode = 502;
          response.end(JSON.stringify({ error: "USDA search is unavailable." }));
        }
      });
    },
  };
}

function openFoodFactsProxy() {
  return {
    name: "open-food-facts-proxy",
    configureServer(server) {
      server.middlewares.use(async function (request, response, next) {
        if (!request.url || !request.url.startsWith("/api/open-food-facts")) {
          next();
          return;
        }

        try {
          const incomingUrl = new URL(request.url, "http://localhost");
          const upstreamUrl = new URL(
            "https://api.openfoodfacts.org/api/v2/search"
          );

          incomingUrl.searchParams.forEach(function (value, key) {
            upstreamUrl.searchParams.set(key, value);
          });

          const upstreamResponse = await fetch(upstreamUrl);
          const body = await upstreamResponse.text();

          response.statusCode = upstreamResponse.status;
          response.setHeader(
            "Content-Type",
            upstreamResponse.headers.get("content-type") || "application/json"
          );
          response.end(body);
        } catch (error) {
          console.error("Open Food Facts proxy failed:", error);
          response.statusCode = 502;
          response.end(
            JSON.stringify({ error: "Open Food Facts search is unavailable." })
          );
        }
      });
    },
  };
}

export default defineConfig(({ mode }) => {
  const env = loadEnv(mode, "src", "");
  const usdaApiKey = env.VITE_USDA_API_KEY;

  return {
    envDir: "src",
    plugins: [
      usdaProxy(usdaApiKey),
      openFoodFactsProxy(),
      react(),
    tailwindcss(),

    VitePWA({
      registerType: "autoUpdate",

      manifest: {
        name: "Calories + Fitness Tracker",
        short_name: "Calories",
        description:
          "Track calories, fitness, nutrition, and weight progress.",

        theme_color: "#10b981",
        background_color: "#ffffff",
        display: "standalone",
        start_url: "/",

        icons: [
          {
            src: "/icon-192.png",
            sizes: "192x192",
            type: "image/png",
          },
          {
            src: "/icon-512.png",
            sizes: "512x512",
            type: "image/png",
          },
        ],
      },
    }),
    ],
  };
});