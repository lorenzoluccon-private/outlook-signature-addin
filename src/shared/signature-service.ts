/*
 * Servizio condiviso per il recupero, la cache
 * e l'inserimento della firma aziendale.
 */

/* global Office, window, console */

import { fetchSignatureFromGraph } from "./graph-signature-service";

import { fetchSignatureFromLegacyApi } from "./legacy-signature-service";

const CACHE_TTL_HOURS = 6;

/*
 * Il salvataggio nei roamingSettings non deve
 * trattenere indefinitamente l'inserimento.
 */
const CACHE_SAVE_TIMEOUT_MILLISECONDS = 2500;

const CACHE_KEY_HTML = "companySignature.html";

const CACHE_KEY_EMAIL = "companySignature.email";

const CACHE_KEY_UPDATED_AT = "companySignature.updatedAt";

export type SignatureSource = "graph" | "legacy" | "cache";

export interface SignatureResult {
  html: string;
  source: SignatureSource;
  email: string;
  updatedAt: string | null;
}

interface CachedSignature {
  html: string;
  email: string;
  updatedAt: string;
}

/*
 * Condivide il recupero in corso all'interno dello
 * stesso runtime JavaScript.
 *
 * In questo modo task pane, comando o altre chiamate
 * nello stesso runtime non avviano più recuperi
 * contemporanei della stessa firma.
 */
let pendingSignatureRequest: Promise<SignatureResult> | null = null;

export function getUserEmail(): string {
  const profileEmail = Office.context.mailbox.userProfile.emailAddress;

  if (!profileEmail || profileEmail.trim() === "") {
    throw new Error("Impossibile determinare l'indirizzo email dell'utente Outlook.");
  }

  return profileEmail.trim().toLowerCase();
}

function getErrorMessage(error: unknown): string {
  if (error instanceof Error) {
    return error.message;
  }

  return String(error);
}

function createTimeoutPromise(milliseconds: number, message: string): Promise<never> {
  return new Promise((_, reject) => {
    window.setTimeout(() => {
      reject(new Error(message));
    }, milliseconds);
  });
}

function getCachedSignature(email: string): CachedSignature | null {
  try {
    const settings = Office.context.roamingSettings;

    const cachedHtml = settings.get(CACHE_KEY_HTML) as string | undefined;

    const cachedEmail = settings.get(CACHE_KEY_EMAIL) as string | undefined;

    const cachedUpdatedAt = settings.get(CACHE_KEY_UPDATED_AT) as string | undefined;

    if (!cachedHtml || !cachedEmail || !cachedUpdatedAt) {
      return null;
    }

    if (cachedEmail.trim().toLowerCase() !== email.trim().toLowerCase()) {
      return null;
    }

    const timestamp = new Date(cachedUpdatedAt).getTime();

    if (Number.isNaN(timestamp)) {
      return null;
    }

    return {
      html: cachedHtml,
      email: cachedEmail,
      updatedAt: cachedUpdatedAt,
    };
  } catch (error) {
    console.warn("[Firma aziendale FG] Lettura cache non riuscita:", getErrorMessage(error));

    return null;
  }
}

function isCacheFresh(cachedSignature: CachedSignature): boolean {
  const updatedAtTime = new Date(cachedSignature.updatedAt).getTime();

  if (Number.isNaN(updatedAtTime)) {
    return false;
  }

  const cacheAgeMilliseconds = Date.now() - updatedAtTime;

  const cacheTtlMilliseconds = CACHE_TTL_HOURS * 60 * 60 * 1000;

  return cacheAgeMilliseconds >= 0 && cacheAgeMilliseconds < cacheTtlMilliseconds;
}

function saveCachedSignatureInternal(
  email: string,
  html: string,
  updatedAt: string
): Promise<void> {
  return new Promise((resolve, reject) => {
    try {
      const settings = Office.context.roamingSettings;

      settings.set(CACHE_KEY_HTML, html);

      settings.set(CACHE_KEY_EMAIL, email);

      settings.set(CACHE_KEY_UPDATED_AT, updatedAt);

      settings.saveAsync((result) => {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
          resolve();
          return;
        }

        reject(new Error(result.error?.message || "Impossibile salvare la firma nella cache."));
      });
    } catch (error) {
      reject(
        error instanceof Error
          ? error
          : new Error("Errore imprevisto durante il salvataggio della cache.")
      );
    }
  });
}

/*
 * Il salvataggio cache è best effort.
 *
 * Se fallisce o supera 2,5 secondi, la firma
 * appena recuperata viene comunque restituita.
 */
async function trySaveCachedSignature(
  email: string,
  html: string,
  updatedAt: string
): Promise<boolean> {
  try {
    await Promise.race([
      saveCachedSignatureInternal(email, html, updatedAt),
      createTimeoutPromise(
        CACHE_SAVE_TIMEOUT_MILLISECONDS,
        "Timeout durante il salvataggio della firma nella cache."
      ),
    ]);

    return true;
  } catch (error) {
    console.warn(
      "[Firma aziendale FG] Firma recuperata, ma cache non salvata:",
      getErrorMessage(error)
    );

    return false;
  }
}

function createCachedResult(cachedSignature: CachedSignature, email: string): SignatureResult {
  return {
    html: cachedSignature.html,
    source: "cache",
    email,
    updatedAt: cachedSignature.updatedAt,
  };
}

async function retrieveSignature(
  email: string,
  cachedSignature: CachedSignature | null
): Promise<SignatureResult> {
  try {
    const graphSignature = await fetchSignatureFromGraph();

    /*
     * updatedAt indica il momento in cui questa copia
     * è stata memorizzata nei roamingSettings.
     *
     * Non utilizziamo lastModifiedDateTime per il TTL,
     * perché rappresenta la modifica del messaggio
     * nella mailbox, non l'aggiornamento della cache.
     */
    const updatedAt = new Date().toISOString();

    await trySaveCachedSignature(email, graphSignature.html, updatedAt);

    return {
      html: graphSignature.html,
      source: "graph",
      email,
      updatedAt,
    };
  } catch (graphError) {
    console.warn(
      "[Firma aziendale FG] Graph non disponibile. Tentativo Legacy API:",
      getErrorMessage(graphError)
    );

    try {
      const legacyHtml = await fetchSignatureFromLegacyApi(email);

      const updatedAt = new Date().toISOString();

      await trySaveCachedSignature(email, legacyHtml, updatedAt);

      return {
        html: legacyHtml,
        source: "legacy",
        email,
        updatedAt,
      };
    } catch (legacyError) {
      /*
       * La cache scaduta è il fallback conclusivo.
       *
       * È preferibile inserire una firma precedente
       * piuttosto che lasciare il messaggio senza firma.
       */
      if (cachedSignature) {
        console.warn(
          "[Firma aziendale FG] Graph e Legacy non disponibili. Utilizzo cache scaduta.",
          {
            graph: getErrorMessage(graphError),
            legacy: getErrorMessage(legacyError),
            cachedAt: cachedSignature.updatedAt,
          }
        );

        return createCachedResult(cachedSignature, email);
      }

      throw new Error(
        "Impossibile recuperare la firma. " +
          "GRAPH=" +
          getErrorMessage(graphError) +
          " | LEGACY=" +
          getErrorMessage(legacyError)
      );
    }
  }
}

export async function getSignature(forceRefresh = false): Promise<SignatureResult> {
  const email = getUserEmail();

  const cachedSignature = getCachedSignature(email);

  /*
   * Se la cache è fresca, non viene effettuata
   * alcuna chiamata di rete.
   */
  if (cachedSignature && !forceRefresh && isCacheFresh(cachedSignature)) {
    return createCachedResult(cachedSignature, email);
  }

  /*
   * Se nello stesso runtime è già in corso un
   * recupero, ne riutilizziamo la Promise.
   */
  if (pendingSignatureRequest) {
    return pendingSignatureRequest;
  }

  pendingSignatureRequest = retrieveSignature(email, cachedSignature).finally(() => {
    /*
     * È fondamentale eliminare sempre la Promise,
     * anche dopo timeout o errore.
     *
     * In caso contrario tutte le mail successive
     * continuerebbero a ricevere lo stesso errore
     * o resterebbero collegate alla vecchia chiamata.
     */
    pendingSignatureRequest = null;
  });

  return pendingSignatureRequest;
}

export function insertSignature(html: string): Promise<void> {
  return new Promise((resolve, reject) => {
    const item = Office.context.mailbox.item;

    if (!item?.body) {
      reject(new Error("Non è disponibile un messaggio in modalità composizione."));
      return;
    }

    if (!html || html.trim() === "") {
      reject(new Error("La firma da inserire è vuota."));
      return;
    }

    item.body.setSignatureAsync(
      html,
      {
        coercionType: Office.CoercionType.Html,
      },
      (result) => {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
          resolve();
          return;
        }

        reject(
          new Error(result.error?.message || "Outlook non ha inserito la firma nel messaggio.")
        );
      }
    );
  });
}

export async function loadAndInsertSignature(forceRefresh = false): Promise<SignatureResult> {
  const signature = await getSignature(forceRefresh);

  await insertSignature(signature.html);

  return signature;
}

export function formatUpdateDate(updatedAt: string | null): string {
  if (!updatedAt) {
    return "Non disponibile";
  }

  const date = new Date(updatedAt);

  if (Number.isNaN(date.getTime())) {
    return "Non disponibile";
  }

  return new Intl.DateTimeFormat("it-IT", {
    dateStyle: "short",
    timeStyle: "short",
  }).format(date);
}
