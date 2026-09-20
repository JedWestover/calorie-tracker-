import React from "react";
import ReactDOM from "react-dom/client";
import { PublicClientApplication } from "@azure/msal-browser";
import { MsalContext, MsalProvider } from "@azure/msal-react";

import { registerSW } from "virtual:pwa-register";

registerSW({
  immediate: true,
});

import App from "./App.jsx";
import "./index.css";

import { msalConfig } from "./authConfig";

const app = window.isSecureContext ? (
  <MsalProvider instance={new PublicClientApplication(msalConfig)}>
    <App />
  </MsalProvider>
) : (
  <MsalContext.Provider
    value={{ instance: null, accounts: [], inProgress: "none" }}
  >
    <App />
  </MsalContext.Provider>
);

ReactDOM.createRoot(document.getElementById("root")).render(
  <React.StrictMode>{app}</React.StrictMode>
);