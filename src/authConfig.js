export const msalConfig = {
  auth: {
    clientId: "fec6c40d-7add-4d2e-b9a9-490ec46b902d",
    authority: "https://login.microsoftonline.com/consumers",
    redirectUri: window.location.origin,
  },
};

export const loginRequest = {
  scopes: ["User.Read", "Files.ReadWrite"],
};