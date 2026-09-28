import {
  acquireGraphTokenInteractive,
  acquireGraphTokenSilent,
} from "./graph-auth-service";

const GRAPH_BASE_URL =
  "https://graph.microsoft.com/v1.0";

const SIGNATURE_FOLDER_NAME =
  "FGSignatureCache";

const SIGNATURE_MESSAGE_SUBJECT =
  "FG_SIGNATURE_V1";

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

async function graphFetch(
  url: string,
  accessToken: string
): Promise<Response> {
  const response = await fetch(url, {
    method: "GET",
    headers: {
      Authorization: `Bearer ${accessToken}`,
      Prefer: 'outlook.body-content-type="html"',
    },
    cache: "no-store",
  });

  if (!response.ok) {
    let errorDetail = "";

    try {
      errorDetail = await response.text();
    } catch {
      errorDetail = "";
    }

    throw new Error(
      `Microsoft Graph ha restituito HTTP ${response.status}. ` +
        errorDetail
    );
  }

  return response;
}

async function findSignatureFolder(
  accessToken: string
): Promise<GraphMailFolder> {
  const params = new URLSearchParams({
    includeHiddenFolders: "true",
    "$select": "id,displayName,isHidden",
    "$top": "100",
  });

  const url =
    `${GRAPH_BASE_URL}/me/mailFolders?` +
    params.toString();

  const response = await graphFetch(
    url,
    accessToken
  );

  const data =
    (await response.json()) as
      GraphCollection<GraphMailFolder>;

  const folder = data.value.find(
    (item) =>
      item.displayName ===
      SIGNATURE_FOLDER_NAME
  );

  if (!folder) {
    throw new Error(
      `La cartella nascosta ${SIGNATURE_FOLDER_NAME} ` +
        "non è presente nella mailbox."
    );
  }

  return folder;
}

async function findSignatureMessage(
  accessToken: string,
  folderId: string
): Promise<GraphMessage> {
  const params = new URLSearchParams({
    "$filter":
      `subject eq '${SIGNATURE_MESSAGE_SUBJECT}'`,
    "$select":
      "id,subject,body,lastModifiedDateTime",
    "$top": "1",
  });

  const encodedFolderId =
    encodeURIComponent(folderId);

  const url =
    `${GRAPH_BASE_URL}/me/mailFolders/` +
    `${encodedFolderId}/messages?` +
    params.toString();

  const response = await graphFetch(
    url,
    accessToken
  );

  const data =
    (await response.json()) as
      GraphCollection<GraphMessage>;

  const message = data.value[0];

  if (!message) {
    throw new Error(
      `Il messaggio ${SIGNATURE_MESSAGE_SUBJECT} ` +
        "non è presente nella mailbox."
    );
  }

  return message;
}

async function fetchSignatureUsingToken(
  accessToken: string
): Promise<GraphSignatureResult> {
  const folder =
    await findSignatureFolder(accessToken);

  const message =
    await findSignatureMessage(
      accessToken,
      folder.id
    );

  const html = message.body?.content;

  if (!html || html.trim() === "") {
    throw new Error(
      "La firma salvata in Microsoft 365 è vuota."
    );
  }

  return {
    html,
    lastModifiedDateTime:
      message.lastModifiedDateTime || null,
  };
}

export async function fetchSignatureFromGraph():
Promise<GraphSignatureResult> {

  const accessToken =
    await acquireGraphTokenSilent();

  return fetchSignatureUsingToken(
    accessToken
  );
}

export async function initializeGraphSignatureAccess():
Promise<GraphSignatureResult> {
  const accessToken =
    await acquireGraphTokenInteractive();

  return fetchSignatureUsingToken(
    accessToken
  );
}