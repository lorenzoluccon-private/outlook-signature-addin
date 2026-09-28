/*
 * Comandi manuali Outlook per la firma aziendale FG.
 */

/* global Office */

import {
  loadAndInsertSignature,
  SignatureResult,
} from "../shared/signature-service";

function notifyInfo(messageText: string): void {
  const item = Office.context.mailbox.item;

  if (!item?.notificationMessages) {
    return;
  }

  const message: Office.NotificationMessageDetails = {
    type:
      Office.MailboxEnums.ItemNotificationMessageType
        .InformationalMessage,
    message: messageText,
    icon: "Icon.80x80",
    persistent: false,
  };

  item.notificationMessages.replaceAsync(
    "CompanySignatureNotification",
    message
  );
}

function notifyError(messageText: string): void {
  const item = Office.context.mailbox.item;

  if (!item?.notificationMessages) {
    return;
  }

  const message: Office.NotificationMessageDetails = {
    type:
      Office.MailboxEnums.ItemNotificationMessageType
        .ErrorMessage,
    message: messageText,
  };

  item.notificationMessages.replaceAsync(
    "CompanySignatureError",
    message
  );
}

function getSuccessMessage(
  result: SignatureResult
): string {
  if (result.source === "graph") {
    return "Firma aziendale aggiornata e inserita correttamente.";
  }

  return "Firma aziendale inserita dalla cache.";
}

function getErrorMessage(error: unknown): string {

  if (error instanceof Error) {
    return error.message;
  }

  try {
    return JSON.stringify(error);
  } catch {
    return String(error);
  }
}

export async function insertCompanySignature(
  event: Office.AddinCommands.Event
): Promise<void> {
  try {
    const result = await loadAndInsertSignature();

    notifyInfo(
      getSuccessMessage(result)
    );
} catch (error) {

  notifyError(
    "TIPO: " +
    typeof error +
    " | VALORE: " +
    String(error)
  );

} finally {
    event.completed();
  }
}

Office.onReady(() => {
  Office.actions.associate(
    "action",
    insertCompanySignature
  );
});