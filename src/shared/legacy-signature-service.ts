/* global fetch, window, AbortController, DOMException, URLSearchParams */

const LEGACY_API_BASE_URL = "https://signature.fgautomazioni.it/signature-api";

const LEGACY_API_TIMEOUT_MILLISECONDS = 6000;

function getErrorMessage(error: unknown): string {
  if (error instanceof Error) {
    return error.message;
  }

  return String(error);
}

function isAbortError(error: unknown): boolean {
  return error instanceof DOMException && error.name === "AbortError";
}

export async function fetchSignatureFromLegacyApi(email: string): Promise<string> {
  const params = new URLSearchParams({
    cmd: "getEmailSign",
    email,
    applicationId: "FGIT",
    appId: "FGIT",
  });

  const url = `${LEGACY_API_BASE_URL}?${params.toString()}`;

  const controller = new AbortController();

  const timeoutId = window.setTimeout(() => {
    controller.abort();
  }, LEGACY_API_TIMEOUT_MILLISECONDS);

  try {
    const response = await fetch(url, {
      method: "GET",
      cache: "no-store",
      signal: controller.signal,
    });

    if (!response.ok) {
      throw new Error(`La Legacy API ha restituito HTTP ` + `${response.status}.`);
    }

    const html = await response.text();

    if (!html || html.trim() === "") {
      throw new Error("La Legacy API ha restituito una firma vuota.");
    }

    return html;
  } catch (error) {
    if (isAbortError(error)) {
      throw new Error(
        `La Legacy API non ha risposto entro ${LEGACY_API_TIMEOUT_MILLISECONDS / 1000} secondi.`
      );
    }

    throw new Error("Recupero firma dalla Legacy API non riuscito: " + getErrorMessage(error));
  } finally {
    window.clearTimeout(timeoutId);
  }
}
