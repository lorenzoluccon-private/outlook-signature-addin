/*
 * Autenticazione Microsoft Graph tramite
 * Nested App Authentication.
 */

/* global Office, window */

import {
  createNestablePublicClientApplication,
  IPublicClientApplication,
  InteractionRequiredAuthError,
} from "@azure/msal-browser";

const CLIENT_ID = "84170119-cf48-41a1-b14b-f28116320cb4";

const TENANT_ID = "35683b98-5a14-4911-a637-34d25c24a531";

const GRAPH_SCOPES = ["Mail.Read", "User.Read"];

/*
 * Timeout applicato soltanto all'autenticazione silenziosa
 * utilizzata dal runtime automatico.
 *
 * Un evento automatico Outlook non deve restare bloccato
 * indefinitamente in attesa di MSAL.
 */
const SILENT_TOKEN_TIMEOUT_MILLISECONDS = 5000;

let msalInstance: IPublicClientApplication | undefined;

let msalInstancePromise: Promise<IPublicClientApplication> | undefined;

function getErrorMessage(error: unknown): string {
  if (error instanceof Error) {
    return error.message;
  }

  return String(error);
}

function withTimeout<T>(
  operation: Promise<T>,
  milliseconds: number,
  timeoutMessage: string
): Promise<T> {
  return new Promise<T>((resolve, reject) => {
    const timeoutId = window.setTimeout(() => {
      reject(new Error(timeoutMessage));
    }, milliseconds);

    operation.then(
      (value) => {
        window.clearTimeout(timeoutId);
        resolve(value);
      },
      (error) => {
        window.clearTimeout(timeoutId);
        reject(error);
      }
    );
  });
}

async function createMsalInstance(): Promise<IPublicClientApplication> {
  const naaSupported = Office.context.requirements.isSetSupported("NestedAppAuth", "1.1");

  if (!naaSupported) {
    throw new Error("Questo client Outlook non supporta Nested App Authentication.");
  }

  return createNestablePublicClientApplication({
    auth: {
      clientId: CLIENT_ID,
      authority: `https://login.microsoftonline.com/${TENANT_ID}`,
    },
    cache: {
      cacheLocation: "localStorage",
    },
  });
}

async function getMsalInstance(): Promise<IPublicClientApplication> {
  if (msalInstance) {
    return msalInstance;
  }

  /*
   * Se più chiamate arrivano contemporaneamente,
   * riutilizziamo la stessa inizializzazione MSAL.
   */
  if (!msalInstancePromise) {
    msalInstancePromise = createMsalInstance()
      .then((instance) => {
        msalInstance = instance;
        return instance;
      })
      .catch((error) => {
        /*
         * Se l'inizializzazione fallisce, permettiamo
         * un nuovo tentativo alla chiamata successiva.
         */
        msalInstancePromise = undefined;
        throw error;
      });
  }

  return msalInstancePromise;
}

/*
 * Usata dal runtime automatico.
 *
 * Non deve mai aprire finestre, popup o dialoghi,
 * perché gli eventi automatici Outlook non
 * possono interagire con l'interfaccia utente.
 */
export async function acquireGraphTokenSilent(): Promise<string> {
  const silentTokenOperation = (async (): Promise<string> => {
    const instance = await getMsalInstance();

    const response = await instance.acquireTokenSilent({
      scopes: GRAPH_SCOPES,
      loginHint: Office.context.mailbox.userProfile.emailAddress,
    });

    return response.accessToken;
  })();

  try {
    return await withTimeout(
      silentTokenOperation,
      SILENT_TOKEN_TIMEOUT_MILLISECONDS,
      `L'autenticazione silenziosa Microsoft Graph ` +
        `non ha risposto entro ${SILENT_TOKEN_TIMEOUT_MILLISECONDS / 1000} secondi.`
    );
  } catch (error) {
    throw new Error("Autenticazione Microsoft Graph non riuscita: " + getErrorMessage(error));
  }
}

/*
 * Usata esclusivamente dal task pane.
 *
 * Può mostrare l'interfaccia Microsoft per
 * l'accesso o il consenso dell'utente.
 *
 * Il flusso interattivo non usa il timeout breve
 * del runtime automatico, perché l'utente può
 * impiegare del tempo per completare l'accesso.
 */
export async function acquireGraphTokenInteractive(): Promise<string> {
  const instance = await getMsalInstance();

  try {
    const silentResponse = await instance.acquireTokenSilent({
      scopes: GRAPH_SCOPES,
      loginHint: Office.context.mailbox.userProfile.emailAddress,
    });

    return silentResponse.accessToken;
  } catch (silentError) {
    if (!(silentError instanceof InteractionRequiredAuthError)) {
      throw silentError;
    }

    const interactiveResponse = await instance.acquireTokenPopup({
      scopes: GRAPH_SCOPES,
      loginHint: Office.context.mailbox.userProfile.emailAddress,
    });

    return interactiveResponse.accessToken;
  }
}
