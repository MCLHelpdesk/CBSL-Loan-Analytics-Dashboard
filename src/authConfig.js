import { LogLevel } from "@azure/msal-browser";

export const msalConfig = {
  auth: {
    clientId: import.meta.env.VITE_CLIENT_ID,
    authority: `https://login.microsoftonline.com/${import.meta.env.VITE_TENANT_ID}`,
    redirectUri: window.location.origin,
  },

  cache: {
    cacheLocation: "localStorage",
    storeAuthStateInCookie: false,
  },

  system: {
    loggerOptions: {
      loggerCallback: (level, message, containsPii) => {
        if (containsPii) {
          return;
        }

        switch (level) {
          case LogLevel.Error:
            console.error(message);
            break;

          case LogLevel.Warning:
            console.warn(message);
            break;

          case LogLevel.Info:
            console.info(message);
            break;

          case LogLevel.Verbose:
            console.debug(message);
            break;

          default:
            break;
        }
      },
    },
  },
};

export const loginRequest = {
  scopes: [
    "User.Read",
    "Sites.Read.All",
    "Sites.ReadWrite.All",
  ],
};