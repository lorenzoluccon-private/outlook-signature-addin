/*
 * Company Signature - firma da API con cache 6 ore
 */

/* global Office */

const SIGNATURE_API_TEMPLATE =
  "https://localhost:3000/signature-api?cmd=getEmailSign&email={{email}}&applicationId=FGIT&appId=FGIT";

const CACHE_TTL_HOURS = 6;

const CACHE_KEY_HTML = "companySignature.html";
const CACHE_KEY_EMAIL = "companySignature.email";
const CACHE_KEY_UPDATED_AT = "companySignature.updatedAt";

function getUserEmail(): string {
  const profileEmail = Office.context.mailbox.userProfile.emailAddress;

  if (!profileEmail || profileEmail.trim() === "") {
    throw new Error(
      "Impossibile determinare l'indirizzo email dell'utente Outlook."
    );
  }

  return profileEmail.trim().toLowerCase();
}

function buildSignatureApiUrl(email: string): string {
  return SIGNATURE_API_TEMPLATE.replace(
    "{{email}}",
    encodeURIComponent(email)
  );
}

function getCachedSignature(email: string): string | null {
  const cachedHtml =
    Office.context.roamingSettings.get(CACHE_KEY_HTML) as string | undefined;

  const cachedEmail =
    Office.context.roamingSettings.get(CACHE_KEY_EMAIL) as string | undefined;

  const cachedUpdatedAt =
    Office.context.roamingSettings.get(CACHE_KEY_UPDATED_AT) as string | undefined;

  if (!cachedHtml || !cachedEmail || !cachedUpdatedAt) {
    return null;
  }

  if (cachedEmail.toLowerCase() !== email.toLowerCase()) {
    return null;
  }

  return cachedHtml;
}

function isCacheFresh(email: string): boolean {
  const cachedHtml =
    Office.context.roamingSettings.get(CACHE_KEY_HTML) as string | undefined;

  const cachedEmail =
    Office.context.roamingSettings.get(CACHE_KEY_EMAIL) as string | undefined;

  const cachedUpdatedAt =
    Office.context.roamingSettings.get(CACHE_KEY_UPDATED_AT) as string | undefined;

  if (!cachedHtml || !cachedEmail || !cachedUpdatedAt) {
    return false;
  }

  if (cachedEmail.toLowerCase() !== email.toLowerCase()) {
    return false;
  }

  const updatedAtTime = new Date(cachedUpdatedAt).getTime();

  if (Number.isNaN(updatedAtTime)) {
    return false;
  }

  const ageMs = Date.now() - updatedAtTime;
  const ttlMs = CACHE_TTL_HOURS * 60 * 60 * 1000;

  return ageMs >= 0 && ageMs < ttlMs;
}

function saveCachedSignature(
  email: string,
  html: string
): Promise<void> {
  return new Promise((resolve, reject) => {
    Office.context.roamingSettings.set(
      CACHE_KEY_HTML,
      html
    );

    Office.context.roamingSettings.set(
      CACHE_KEY_EMAIL,
      email
    );

    Office.context.roamingSettings.set(
      CACHE_KEY_UPDATED_AT,
      new Date().toISOString()
    );

    Office.context.roamingSettings.saveAsync(
      (result) => {
        if (
          result.status ===
          Office.AsyncResultStatus.Succeeded
        ) {
          resolve();
        } else {
          reject(result.error);
        }
      }
    );
  });
}

async function fetchSignatureFromApi(
  email: string
): Promise<string> {

  const url = buildSignatureApiUrl(email);

  try {
    const response = await fetch(url, {
      method: "GET",
      cache: "no-store",
      mode: "cors",
    });

    const contentType =
      response.headers.get("content-type") || "";

    if (!response.ok) {
      const errorBody =
        await response.text().catch(() => "");

      throw new Error(
        `API firma non disponibile. HTTP ${response.status}. Risposta: ${errorBody.substring(0, 300)}`
      );
    }

    if (
      contentType
        .toLowerCase()
        .includes("application/json")
    ) {
      const data = await response.json();

      if (typeof data === "string") {
        return data;
      }

      if (data && typeof data.html === "string") {
        return data.html;
      }

      if (
        data &&
        typeof data.signature === "string"
      ) {
        return data.signature;
      }

      throw new Error(
        "La risposta JSON non contiene html o signature."
      );
    }

    const html = await response.text();

    if (!html || html.trim() === "") {
      throw new Error(
        "L'API ha restituito una firma vuota."
      );
    }

    return html;

  } catch (error) {

    if (error instanceof TypeError) {
      throw new Error(
        `Fetch fallito verso API firma. Possibili cause: CORS, TLS, redirect, rete o blocco WebView. Dettaglio: ${error.message}`
      );
    }

    if (error instanceof Error) {
      throw new Error(
        `Errore chiamata API firma: ${error.message}`
      );
    }

    throw new Error(
      "Errore sconosciuto durante la chiamata API firma."
    );
  }
}

async function getSignatureHtml(): Promise<{
  html: string;
  source: "api" | "cache";
}> {

  const email = getUserEmail();

  if (isCacheFresh(email)) {

    const cached =
      getCachedSignature(email);

    if (cached) {
      return {
        html: cached,
        source: "cache",
      };
    }
  }

  try {

    const freshHtml =
      await fetchSignatureFromApi(email);

    await saveCachedSignature(
      email,
      freshHtml
    );

    return {
      html: freshHtml,
      source: "api",
    };

  } catch (error) {

    const cached =
      getCachedSignature(email);

    if (cached) {
      return {
        html: cached,
        source: "cache",
      };
    }

    throw error;
  }
}

function notifyInfo(messageText: string) {

  const message: Office.NotificationMessageDetails =
    {
      type:
        Office.MailboxEnums
          .ItemNotificationMessageType
          .InformationalMessage,

      message: messageText,

      icon: "Icon.80x80",

      persistent: false,
    };

  Office.context.mailbox.item
    .notificationMessages
    .replaceAsync(
      "CompanySignatureNotification",
      message
    );
}

function notifyError(messageText: string) {

  const message: Office.NotificationMessageDetails =
    {
      type:
        Office.MailboxEnums
          .ItemNotificationMessageType
          .ErrorMessage,

      message: messageText,
    };

  Office.context.mailbox.item
    .notificationMessages
    .replaceAsync(
      "CompanySignatureError",
      message
    );
}

function setSignature(
  html: string
): Promise<void> {

  return new Promise((resolve, reject) => {

    Office.context.mailbox.item.body
      .setSignatureAsync(
        html,
        {
          coercionType:
            Office.CoercionType.Html,
        },
        (result) => {

          if (
            result.status ===
            Office.AsyncResultStatus.Succeeded
          ) {
            resolve();
          } else {
            reject(result.error);
          }
        }
      );
  });
}

export async function setNotificationInOutlook(
  event: Office.AddinCommands.Event
) {

  try {

    const result =
      await getSignatureHtml();

    await setSignature(result.html);

    if (result.source === "api") {

      notifyInfo(
        "Firma aggiornata dall'API e inserita correttamente."
      );

    } else {

      notifyInfo(
        "Firma inserita dalla cache locale."
      );
    }

  } catch (error) {

    const message =
      error instanceof Error
        ? error.message
        : "Errore sconosciuto.";

    notifyError(
      `Errore firma: ${message}`
    );

  } finally {

    event.completed();
  }
}

Office.onReady(async () => {
  Office.actions.associate(
    "action",
    setNotificationInOutlook
  );
});