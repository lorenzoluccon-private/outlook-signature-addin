/*
 * Pannello di gestione della firma aziendale.
 */

/* global document, Office */

import {
  formatUpdateDate,
  getUserEmail,
  loadAndInsertSignature,
  SignatureResult,
} from "../shared/signature-service";

import {
  initializeGraphSignatureAccess,
} from "../shared/graph-signature-service";

const ELEMENT_SIDELOAD_MESSAGE = "sideload-msg";
const ELEMENT_APP_BODY = "app-body";
const ELEMENT_RUN_BUTTON = "run";
const ELEMENT_CONNECT_BUTTON = "connect-microsoft-365";
const ELEMENT_STATUS = "item-subject";

Office.onReady((info) => {
  if (info.host !== Office.HostType.Outlook) {
    return;
  }

  initializeTaskPane();
});

function initializeTaskPane(): void {
  const sideloadMessage = document.getElementById(
    ELEMENT_SIDELOAD_MESSAGE
  );

  const appBody = document.getElementById(
    ELEMENT_APP_BODY
  );

  const runButton = document.getElementById(
    ELEMENT_RUN_BUTTON
  );

  const connectButton = document.getElementById(
    ELEMENT_CONNECT_BUTTON
  );

  if (sideloadMessage) {
    sideloadMessage.style.display = "none";
  }

  if (appBody) {
    appBody.style.display = "flex";
  }

  if (runButton) {
    runButton.onclick = runOutlook;

    runButton.addEventListener(
      "keydown",
      (event: KeyboardEvent) => {
        if (
          event.key === "Enter" ||
          event.key === " "
        ) {
          event.preventDefault();
          void runOutlook();
        }
      }
    );
  }

  if (connectButton) {
    connectButton.onclick =
      connectToMicrosoft365;

    connectButton.addEventListener(
      "keydown",
      (event: KeyboardEvent) => {
        if (
          event.key === "Enter" ||
          event.key === " "
        ) {
          event.preventDefault();
          void connectToMicrosoft365();
        }
      }
    );
  }

  showInitialStatus();
}

function showInitialStatus(): void {
  try {
    const email = getUserEmail();

    setStatus(
      `Account rilevato: ${email}`,
      "info"
    );
  } catch (error) {
    setStatus(
      getErrorMessage(error),
      "error"
    );
  }
}

function setElementEnabled(
  elementId: string,
  enabled: boolean
): void {
  const element =
    document.getElementById(elementId);

  if (!element) {
    return;
  }

  element.setAttribute(
    "aria-disabled",
    enabled ? "false" : "true"
  );

  element.style.pointerEvents =
    enabled ? "auto" : "none";

  element.style.opacity =
    enabled ? "1" : "0.6";
}

function setButtonEnabled(
  enabled: boolean
): void {
  setElementEnabled(
    ELEMENT_RUN_BUTTON,
    enabled
  );
}

function setConnectButtonEnabled(
  enabled: boolean
): void {
  setElementEnabled(
    ELEMENT_CONNECT_BUTTON,
    enabled
  );
}

function setStatus(
  message: string,
  type: "info" | "success" | "error"
): void {
  const statusElement = document.getElementById(
    ELEMENT_STATUS
  );

  if (!statusElement) {
    return;
  }

  statusElement.textContent = message;
  statusElement.style.display = "block";
  statusElement.style.fontWeight = "600";
  statusElement.style.lineHeight = "1.4";

  switch (type) {
    case "success":
      statusElement.style.color = "#107c10";
      break;

    case "error":
      statusElement.style.color = "#a4262c";
      break;

    default:
      statusElement.style.color = "#323130";
      break;
  }
}

function getSuccessMessage(
  result: SignatureResult
): string {
  const sourceDescription =
    result.source === "graph"
      ? "Microsoft 365"
      : "cache locale";

  const updatedAt =
    formatUpdateDate(result.updatedAt);

  return [
    "Firma inserita correttamente.",
    `Origine: ${sourceDescription}.`,
    `Ultimo aggiornamento: ${updatedAt}.`,
  ].join(" ");
}

function getErrorMessage(
  error: unknown
): string {
  if (error instanceof Error) {
    return error.message;
  }

  return "Si è verificato un errore sconosciuto.";
}

export async function connectToMicrosoft365():
Promise<void> {
  setConnectButtonEnabled(false);
  setButtonEnabled(false);

  setStatus(
    "Connessione a Microsoft 365 in corso...",
    "info"
  );

  try {
    const result =
      await initializeGraphSignatureAccess();

    const byteLength =
      new TextEncoder().encode(
        result.html
      ).length;

    const updatedAt =
      formatUpdateDate(
        result.lastModifiedDateTime
      );

    setStatus(
      [
        "Connessione a Microsoft 365 completata.",
        "Firma Graph trovata correttamente.",
        `Dimensione: ${byteLength} byte.`,
        `Ultimo aggiornamento: ${updatedAt}.`,
        "L'inserimento automatico può ora utilizzare Microsoft Graph.",
      ].join(" "),
      "success"
    );
  } catch (error) {
    setStatus(
      "Connessione a Microsoft 365 non riuscita: " +
        getErrorMessage(error),
      "error"
    );
  } finally {
    setConnectButtonEnabled(true);
    setButtonEnabled(true);
  }
}

export async function runOutlook():
Promise<void> {
  setButtonEnabled(false);
  setConnectButtonEnabled(false);

  setStatus(
    "Recupero della firma aziendale in corso...",
    "info"
  );

  try {
    const result =
      await loadAndInsertSignature(true);

    setStatus(
      getSuccessMessage(result),
      "success"
    );
  } catch (error) {
    setStatus(
      "Impossibile inserire la firma: " +
        getErrorMessage(error),
      "error"
    );
  } finally {
    setButtonEnabled(true);
    setConnectButtonEnabled(true);
  }
}