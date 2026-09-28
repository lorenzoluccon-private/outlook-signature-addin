/*
 * Servizio condiviso per il recupero, la cache
 * e l'inserimento della firma aziendale.
 */

/* global Office */

import {
  fetchSignatureFromGraph,
} from "./graph-signature-service";
import {
  fetchSignatureFromLegacyApi,
} from "./legacy-signature-service";

const CACHE_TTL_HOURS = 6;

const CACHE_KEY_HTML = "companySignature.html";
const CACHE_KEY_EMAIL = "companySignature.email";
const CACHE_KEY_UPDATED_AT = "companySignature.updatedAt";

export type SignatureSource = "graph" | "cache";

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

export function getUserEmail(): string {
  const profileEmail =
    Office.context.mailbox.userProfile.emailAddress;

  if (!profileEmail || profileEmail.trim() === "") {
    throw new Error(
      "Impossibile determinare l'indirizzo email dell'utente Outlook."
    );
  }

  return profileEmail.trim().toLowerCase();
}

function getCachedSignature(
  email: string
): CachedSignature | null {
  const settings = Office.context.roamingSettings;

  const cachedHtml =
    settings.get(CACHE_KEY_HTML) as string | undefined;

  const cachedEmail =
    settings.get(CACHE_KEY_EMAIL) as string | undefined;

  const cachedUpdatedAt =
    settings.get(CACHE_KEY_UPDATED_AT) as string | undefined;

  if (!cachedHtml || !cachedEmail || !cachedUpdatedAt) {
    return null;
  }

  if (
    cachedEmail.toLowerCase() !==
    email.toLowerCase()
  ) {
    return null;
  }

  const timestamp =
    new Date(cachedUpdatedAt).getTime();

  if (Number.isNaN(timestamp)) {
    return null;
  }

  return {
    html: cachedHtml,
    email: cachedEmail,
    updatedAt: cachedUpdatedAt,
  };
}

function isCacheFresh(
  cachedSignature: CachedSignature
): boolean {
  const updatedAtTime =
    new Date(cachedSignature.updatedAt).getTime();

  if (Number.isNaN(updatedAtTime)) {
    return false;
  }

  const cacheAgeMilliseconds =
    Date.now() - updatedAtTime;

  const cacheTtlMilliseconds =
    CACHE_TTL_HOURS * 60 * 60 * 1000;

  return (
    cacheAgeMilliseconds >= 0 &&
    cacheAgeMilliseconds < cacheTtlMilliseconds
  );
}

function saveCachedSignature(
  email: string,
  html: string
): Promise<string> {
  return new Promise((resolve, reject) => {
    const updatedAt =
      new Date().toISOString();

    const settings =
      Office.context.roamingSettings;

    settings.set(CACHE_KEY_HTML, html);
    settings.set(CACHE_KEY_EMAIL, email);
    settings.set(CACHE_KEY_UPDATED_AT, updatedAt);

    settings.saveAsync((result) => {
      if (
        result.status ===
        Office.AsyncResultStatus.Succeeded
      ) {
        resolve(updatedAt);
        return;
      }

      reject(
        new Error(
          result.error?.message ||
          "Impossibile salvare la firma nella cache."
        )
      );
    });
  });
}

export async function getSignature(
  forceRefresh = false
): Promise<SignatureResult> {

  const email = getUserEmail();

  const cachedSignature =
    getCachedSignature(email);

  if (
    cachedSignature &&
    !forceRefresh &&
    isCacheFresh(cachedSignature)
  ) {
    return {
      html: cachedSignature.html,
      source: "cache",
      email,
      updatedAt: cachedSignature.updatedAt,
    };
  }

  try {

    const graphSignature =
      await fetchSignatureFromGraph();

    const updatedAt =
      graphSignature.lastModifiedDateTime ||
      new Date().toISOString();

    try {
      await saveCachedSignature(
        email,
        graphSignature.html
      );

    } catch (cacheError) {
      console.error(
        "[Firma Graph] Firma recuperata, ma salvataggio cache non riuscito:",
        cacheError
      );
    }

    return {
      html: graphSignature.html,
      source: "graph",
      email,
      updatedAt,
    };

  } catch (graphError) {

    console.warn(
      "[Firma Graph] Graph non disponibile, fallback Legacy API",
      graphError
    );

    try {
      const legacyHtml =
        await fetchSignatureFromLegacyApi(email);

console.log(
  "LEGACY HTML LENGTH:",
  legacyHtml.length
);        
      try {
        await saveCachedSignature(
          email,
          legacyHtml
        );
      } catch (cacheError) {
        console.error(
          "[Firma Legacy] Salvataggio cache fallito",
          cacheError
        );
      }

      return {
        html: legacyHtml,
        source: "cache",
        email,
        updatedAt: new Date().toISOString(),
      };

    } catch (legacyError) {

      throw new Error(
        "GRAPH=" +
        String(graphError) +
        " | LEGACY=" +
        String(legacyError)
      );

    }
  }
}

export function insertSignature(
  html: string
): Promise<void> {
  return new Promise((resolve, reject) => {
    const item = Office.context.mailbox.item;

    if (!item?.body) {
      reject(
        new Error(
          "Non è disponibile un messaggio in modalità composizione."
        )
      );
      return;
    }

    item.body.setSignatureAsync(
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
          return;
        }

        reject(
          new Error(
            result.error?.message ||
            "Outlook non ha inserito la firma nel messaggio."
          )
        );
      }
    );
  });
}

export async function loadAndInsertSignature(
  forceRefresh = false
): Promise<SignatureResult> {

  const signature =
    await getSignature(forceRefresh);

  await insertSignature(signature.html);

  return signature;
}

export function formatUpdateDate(
  updatedAt: string | null
): string {

  if (!updatedAt) {
    return "Non disponibile";
  }

  const date = new Date(updatedAt);

  if (Number.isNaN(date.getTime())) {
    return "Non disponibile";
  }

  return new Intl.DateTimeFormat(
    "it-IT",
    {
      dateStyle: "short",
      timeStyle: "short",
    }
  ).format(date);
}