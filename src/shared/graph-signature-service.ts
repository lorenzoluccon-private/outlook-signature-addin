/* global fetch, window, AbortController, AbortSignal, DOMException, Response, URLSearchParams */

import { acquireGraphTokenInteractive, acquireGraphTokenSilent } from "./graph-auth-service";

const GRAPH_BASE_URL = "https://graph.microsoft.com/v1.0";

const SIGNATURE_FOLDER_NAME = "FGSignatureCache";

const SIGNATURE_MESSAGE_SUBJECT = "FG_SIGNATURE_V1";

/*
 * Tempo massimo complessivo per le chiamate HTTP Graph.
 *
 * Lo stesso AbortSignal viene utilizzato per:
 * - ricerca della cartella;
 * - ricerca del messaggio.
 *
 * Non sono quindi 8 secondi per richiesta:
 * sono 8 secondi complessivi per entrambe.
 */
const GRAPH_FETCH_TIMEOUT_MILLISECONDS = 8000;

interface GraphCollection<T> {
  value: T[];
}

interface GraphMailFolder {
  id: string;
  displayName: string;
  isHidden?: boolean;
}

interface GraphMessage {
  id: string;
  subject: string;
  body?: {
    contentType?: string;
    content?: string;
  };
  lastModifiedDateTime?: string;
}

export interface GraphSignatureResult {
  html: string;
  lastModifiedDateTime: string | null;
}

function getErrorMessage(error: unknown): string {
  if (error instanceof Error) {
    return error.message;
  }

  return String(error);
}

function isAbortError(error: unknown): boolean {
  return error instanceof DOMException && error.name === "AbortError";
}

async function graphFetch(
  url: string,
  accessToken: string,
  signal: AbortSignal
): Promise<Response> {
  let response: Response;

  try {
    response = await fetch(url, {
      method: "GET",
      headers: {
        Authorization: `Bearer ${accessToken}`,
        Prefer: 'outlook.body-content-type="html"',
      },
      cache: "no-store",
      signal,
    });
  } catch (error) {
    if (isAbortError(error)) {
      throw new Error(
        `Microsoft Graph non ha completato il recupero ` +
          `della firma entro ${GRAPH_FETCH_TIMEOUT_MILLISECONDS / 1000} secondi.`
      );
    }

    throw new Error(
      "Errore di rete durante la chiamata a Microsoft Graph: " + getErrorMessage(error)
    );
  }

  if (!response.ok) {
    let errorDetail = "";

    try {
      errorDetail = await response.text();
    } catch {
      errorDetail = "";
    }

    const detailSuffix = errorDetail.trim() !== "" ? ` ${errorDetail}` : "";

    throw new Error(`Microsoft Graph ha restituito HTTP ` + `${response.status}.${detailSuffix}`);
  }

  return response;
}

async function findSignatureFolder(
  accessToken: string,
  signal: AbortSignal
): Promise<GraphMailFolder> {
  const params = new URLSearchParams({
    includeHiddenFolders: "true",
    $select: "id,displayName,isHidden",
    $top: "100",
  });

  const url = `${GRAPH_BASE_URL}/me/mailFolders?` + params.toString();

  const response = await graphFetch(url, accessToken, signal);

  const data = (await response.json()) as GraphCollection<GraphMailFolder>;

  if (!Array.isArray(data.value)) {
    throw new Error("Microsoft Graph ha restituito un elenco cartelle non valido.");
  }

  const folder = data.value.find((item) => item.displayName === SIGNATURE_FOLDER_NAME);

  if (!folder) {
    throw new Error(
      `La cartella nascosta ${SIGNATURE_FOLDER_NAME} ` + "non è presente nella mailbox."
    );
  }

  return folder;
}

async function findSignatureMessage(
  accessToken: string,
  folderId: string,
  signal: AbortSignal
): Promise<GraphMessage> {
  const params = new URLSearchParams({
    $filter: `subject eq '${SIGNATURE_MESSAGE_SUBJECT}'`,
    $select: "id,subject,body,lastModifiedDateTime",
    $top: "1",
  });

  const encodedFolderId = encodeURIComponent(folderId);

  const url =
    `${GRAPH_BASE_URL}/me/mailFolders/` + `${encodedFolderId}/messages?` + params.toString();

  const response = await graphFetch(url, accessToken, signal);

  const data = (await response.json()) as GraphCollection<GraphMessage>;

  if (!Array.isArray(data.value)) {
    throw new Error("Microsoft Graph ha restituito un elenco messaggi non valido.");
  }

  const message = data.value[0];

  if (!message) {
    throw new Error(`Il messaggio ${SIGNATURE_MESSAGE_SUBJECT} ` + "non è presente nella mailbox.");
  }

  return message;
}

async function fetchSignatureUsingToken(accessToken: string): Promise<GraphSignatureResult> {
  const controller = new AbortController();

  const timeoutId = window.setTimeout(() => {
    controller.abort();
  }, GRAPH_FETCH_TIMEOUT_MILLISECONDS);

  try {
    const folder = await findSignatureFolder(accessToken, controller.signal);

    const message = await findSignatureMessage(accessToken, folder.id, controller.signal);

    const html = message.body?.content;

    if (!html || html.trim() === "") {
      throw new Error("La firma salvata in Microsoft 365 è vuota.");
    }

    return {
      html,
      lastModifiedDateTime: message.lastModifiedDateTime || null,
    };
  } finally {
    window.clearTimeout(timeoutId);
  }
}

export async function fetchSignatureFromGraph(): Promise<GraphSignatureResult> {
  const accessToken = await acquireGraphTokenSilent();

  return fetchSignatureUsingToken(accessToken);
}

export async function initializeGraphSignatureAccess(): Promise<GraphSignatureResult> {
  const accessToken = await acquireGraphTokenInteractive();

  return fetchSignatureUsingToken(accessToken);
}
