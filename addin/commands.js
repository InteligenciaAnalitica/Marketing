/* global Office, OfficeRuntime */

const AAD_APP_CLIENT_ID = "8100d731-7de2-416e-a112-a6c62a361a9a";

const LOGO_URL =
  "https://raw.githubusercontent.com/InteligenciaAnalitica/Marketing/main/firma/logo-horizontal-transparente.png";
const SITE_URL = "https://www.inteligenciaanalitica.com";

Office.onReady(() => {
  // El add-in ya queda listo para que Office llame a insertSignature.
});

Office.actions.associate("insertSignature", insertSignature);

async function insertSignature(event) {
  let displayName = "";
  let jobTitle = "";

  try {
    displayName = Office.context.mailbox.userProfile.displayName || "";
  } catch (e) {
    // sin perfil disponible, seguimos con lo que haya
  }

  try {
    const perfil = await obtenerPerfilDesdeGraph();
    if (perfil) {
      displayName = perfil.displayName || displayName;
      jobTitle = perfil.jobTitle || "";
    }
  } catch (e) {
    console.log("No se pudo leer el perfil desde Microsoft Graph:", e);
    // seguimos igual, la firma se arma solo con lo que Office.js ya sabe (nombre)
  }

  const html = construirFirmaHtml(displayName, jobTitle);

  Office.context.mailbox.item.body.setSignatureAsync(
    html,
    { coercionType: Office.CoercionType.Html },
    (result) => {
      if (result.status === Office.AsyncResultStatus.Failed) {
        console.log("Error al insertar la firma:", result.error.message);
      }
      event.completed();
    }
  );
}

async function obtenerPerfilDesdeGraph() {
  if (typeof msal === "undefined") {
    throw new Error("msal-browser no está cargado.");
  }

  const pca = await msal.createNestablePublicClientApplication({
    auth: {
      clientId: AAD_APP_CLIENT_ID,
      authority: "https://login.microsoftonline.com/common",
    },
  });

  const tokenRequest = { scopes: ["User.Read"] };
  let tokenResponse;

  const cuentas = pca.getAllAccounts();
  try {
    if (cuentas.length > 0) {
      tokenResponse = await pca.acquireTokenSilent({
        ...tokenRequest,
        account: cuentas[0],
      });
    } else {
      tokenResponse = await pca.ssoSilent(tokenRequest);
    }
  } catch (e) {
    tokenResponse = await pca.acquireTokenPopup(tokenRequest);
  }

  const respuesta = await fetch(
    "https://graph.microsoft.com/v1.0/me?$select=displayName,jobTitle,mail",
    {
      headers: { Authorization: `Bearer ${tokenResponse.accessToken}` },
    }
  );

  if (!respuesta.ok) {
    throw new Error(`Graph respondió ${respuesta.status}`);
  }

  return await respuesta.json();
}

function construirFirmaHtml(displayName, jobTitle) {
  const nombre = escaparHtml(displayName || "");
  const puesto = escaparHtml(jobTitle || "");

  return `
    <table role="presentation" cellpadding="0" cellspacing="0" style="border-collapse:collapse;width:auto;margin-top:14px;font-family:Calibri,Arial,sans-serif;">
      <tr>
        <td style="padding-top:12px;padding-bottom:8px;">
          <a href="${SITE_URL}" target="_blank" style="text-decoration:none;">
            <img src="${LOGO_URL}" width="110" style="display:block;height:auto;border:0;outline:none;" alt="Inteligencia Analitica">
          </a>
        </td>
      </tr>
      <tr>
        <td style="font-size:12.5px;line-height:18px;white-space:nowrap;font-family:Calibri,Arial,sans-serif;">
          <span style="font-weight:700;color:#003580;">${nombre}</span>
          <span style="color:#9aa2ad;">&nbsp;|&nbsp;</span>
          <span style="color:#5d6470;">${puesto}</span>
        </td>
      </tr>
      <tr>
        <td style="color:#5d6470;font-size:12px;line-height:20px;padding-top:2px;font-family:Calibri,Arial,sans-serif;">
          <a href="${SITE_URL}" target="_blank" style="text-decoration:none;color:#5d6470;">
            <span style="background:#e3f2fc;">www.inteligenciaanalitica.com</span>
          </a>
        </td>
      </tr>
    </table>
  `;
}

function escaparHtml(texto) {
  return String(texto)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;");
}
