/*
 * Runtime automatico leggero per la firma aziendale FG.
 */

/* global Office */

import {
  loadAndInsertSignature,
} from "../shared/signature-service";

function getErrorMessage(error: unknown): string {
  if (error instanceof Error) {
    return error.message;
  }

  return "Errore sconosciuto.";
}

async function onNewMessageCompose(
  event: Office.AddinCommands.Event
): Promise<void> {
  try {
    await loadAndInsertSignature();
  } catch (error) {
    console.error(
      "[Firma aziendale FG] Inserimento automatico non riuscito:",
      getErrorMessage(error)
    );
  } finally {
    event.completed();
  }
}

/*
 * L'associazione deve essere eseguita direttamente.
 * Outlook Classic per Windows non esegue Office.onReady()
 * all'avvio del runtime JavaScript per gli eventi automatici.
 */
Office.actions.associate(
  "onNewMessageCompose",
  onNewMessageCompose
);