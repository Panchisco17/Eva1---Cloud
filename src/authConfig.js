export const msalConfig = {
    auth: {
        clientId: "291fadd9-b4ff-41f2-9e0e-19f92e1cda52",
        authority: "https://login.microsoftonline.com/fe70fa16-a45c-4810-b613-80f2842db999",
        redirectUri: "http://localhost:3000",
    }
};

export const loginRequest = {
    scopes: ["api://97fd6d32-e248-449e-9862-b31ee3e937ec/Solicitudes.Read"]
};