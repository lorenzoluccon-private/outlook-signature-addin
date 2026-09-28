export async function fetchSignatureFromLegacyApi(
  email: string
): Promise<string> {

  const url =
    "https://signature.fgautomazioni.it/signature-api" +
    "?cmd=getEmailSign" +
    "&email=" + encodeURIComponent(email) +
    "&applicationId=FGIT" +
    "&appId=FGIT";

  try {

    const response = await fetch(url);

    const html = await response.text();

    return `
      <div style="border:2px solid green;padding:10px">
        FETCH OK<br>
        STATUS=${response.status}<br>
        LENGTH=${html.length}
      </div>
    `;

  } catch (e) {

    return "";
    /*
    return `
      <div style="border:2px solid red;padding:10px">
        FETCH KO<br>
        ${String(e)}
      </div>
    `;
    */
  }
}