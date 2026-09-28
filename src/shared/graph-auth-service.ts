/*
 * Autenticazione Microsoft Graph tramite
 * Nested App Authentication.
 */

/* global Office */

import {
  createNestablePublicClientApplication,
  IPublicClientApplication,
  InteractionRequiredAuthError,
} from "@azure/msal-browser";

const CLIENT_ID =
  "84170119-cf48-41a1-b14b-f28116320cb4";

const TENANT_ID =
  "35683b98-5a14-4911-a637-34d25c24a531";

const GRAPH_SCOPES = [
  "Mail.Read",
  "User.Read",
];

let msalInstance:
  | IPublicClientApplication
  | undefined;

async function getMsalInstance():
  Promise<IPublicClientApplication> {
  if (msalInstance) {
    return msalInstance;
  }

  const naaSupported =
    Office.context.requirements.isSetSupported(
      "NestedAppAuth",
      "1.1"
    );

  if (!naaSupported) {
    throw new Error(
      "Questo client Outlook non supporta Nested App Authentication."
    );
  }

  msalInstance =
    await createNestablePublicClientApplication({
      auth: {
        clientId: CLIENT_ID,
        authority:
          `https://login.microsoftonline.com/${TENANT_ID}`,
      },
      cache: {
        cacheLocation: "localStorage",
      },
    });

  return msalInstance;
}

/*
 * Usata dal runtime automatico.
 *
 * Non deve mai aprire finestre, popup o dialoghi,
 * perché gli eventi automatici Outlook non
 * possono interagire con l'interfaccia utente.
 */
export async function acquireGraphTokenSilent():
  Promise<string> {
  const instance = await getMsalInstance();

  try {
    const response =
      await instance.acquireTokenSilent({
        scopes: GRAPH_SCOPES,
        loginHint:
          Office.context.mailbox.userProfile
            .emailAddress,
      });

    return response.accessToken;
  } catch (error) {

  throw new Error(
    "MSAL: " +
    JSON.stringify(error)
  );

}
}

/*
 * Usata esclusivamente dal task pane.
 *
 * Può mostrare l'interfaccia Microsoft per
 * l'accesso o il consenso dell'utente.
 */
export async function acquireGraphTokenInteractive():
  Promise<string> {
  const instance = await getMsalInstance();

  try {
    const silentResponse =
      await instance.acquireTokenSilent({
        scopes: GRAPH_SCOPES,
        loginHint:
          Office.context.mailbox.userProfile
            .emailAddress,
      });

    return silentResponse.accessToken;
  } catch (silentError) {
    if (
      !(
        silentError instanceof
        InteractionRequiredAuthError
      )
    ) {
      throw silentError;
    }

    const interactiveResponse =
      await instance.acquireTokenPopup({
        scopes: GRAPH_SCOPES,
        loginHint:
          Office.context.mailbox.userProfile
            .emailAddress,
      });

    return interactiveResponse.accessToken;
  }
}