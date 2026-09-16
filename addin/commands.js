/* global Office, OfficeRuntime */

const AAD_APP_CLIENT_ID = "8100d731-7de2-416e-a112-a6c62a361a9a";

const LOGO_URL =
  "https://raw.githubusercontent.com/InteligenciaAnalitica/Marketing/main/firma/logo-horizontal-transparente.png";
const WHATSAPP_ICON_URL =
  "https://raw.githubusercontent.com/InteligenciaAnalitica/Marketing/main/firma/whatsapp-icon.png";
const SITE_URL = "https://www.inteligenciaanalitica.com";

Office.onReady(() => {
  // El add-in ya queda listo para que Office llame a insertSignature.
});

Office.actions.associate("insertSignature", insertSignature);
Office.actions.associate("autoInsertSignature", autoInsertSignature);

async function obtenerDatosPerfil() {
  let displayName = "";
  let jobTitle = "";
  let mobilePhone = "";

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
      mobilePhone = perfil.mobilePhone || "";
    }
  } catch (e) {
    console.log("No se pudo leer el perfil desde Microsoft Graph:", e);
    // seguimos igual, la firma se arma solo con lo que Office.js ya sabe (nombre)
  }

  return { displayName, jobTitle, mobilePhone };
}

// Botón "Insertar firma": inserta en la posición del cursor. Sirve para
// volver a insertarla a mano, o para clientes donde el evento automático
// todavía no esté disponible.
async function insertSignature(event) {
  const { displayName, jobTitle, mobilePhone } = await obtenerDatosPerfil();
  const html = construirFirmaHtml(displayName, jobTitle, mobilePhone);

  Office.context.mailbox.item.body.setSelectedDataAsync(
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

// Se dispara solo al abrir un correo nuevo, una respuesta o un reenvío
// (Responder y Responder a todos usan el mismo tipo de evento "reply").
async function autoInsertSignature(event) {
  const { displayName, jobTitle, mobilePhone } = await obtenerDatosPerfil();
  const html = construirFirmaHtml(displayName, jobTitle, mobilePhone);

  Office.context.mailbox.item.body.setSignatureAsync(
    html,
    { coercionType: Office.CoercionType.Html },
    (result) => {
      if (result.status === Office.AsyncResultStatus.Failed) {
        console.log("Error al insertar la firma automática:", result.error.message);
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
    "https://graph.microsoft.com/v1.0/me?$select=displayName,jobTitle,mail,mobilePhone",
    {
      headers: { Authorization: `Bearer ${tokenResponse.accessToken}` },
    }
  );

  if (!respuesta.ok) {
    throw new Error(`Graph respondió ${respuesta.status}`);
  }

  return await respuesta.json();
}

function construirFirmaHtml(displayName, jobTitle, mobilePhone) {
  const nombre = escaparHtml(displayName || "");
  const puesto = escaparHtml(jobTitle || "");
  const filaWhatsapp = construirFilaWhatsapp(mobilePhone);

  return `
    <table role="presentation" cellpadding="0" cellspacing="0" style="border-collapse:collapse;width:auto;margin-top:0;font-family:Calibri,Arial,sans-serif;">
      <tr>
        <td style="padding-top:2px;padding-bottom:8px;">
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
      ${filaWhatsapp}
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

// Solo se agrega si "Telefono movil" esta cargado en Microsoft 365. Si esta
// vacio, esta funcion devuelve un string vacio y la fila directamente no existe.
function construirFilaWhatsapp(mobilePhone) {
  if (!mobilePhone || !mobilePhone.trim()) {
    return "";
  }

  const telefonoVisible = escaparHtml(mobilePhone.trim());
  const telefonoWa = mobilePhone.replace(/[^\d]/g, "");

  return `
      <tr>
        <td style="padding-top:1px;font-family:Calibri,Arial,sans-serif;">
          <a href="https://wa.me/${telefonoWa}" target="_blank" style="text-decoration:none;">
            <table role="presentation" cellpadding="0" cellspacing="0" style="border-collapse:collapse;">
              <tr>
                <td valign="middle" style="padding-right:5px;">
                  <img src="${WHATSAPP_ICON_URL}" width="13" height="13" alt="WhatsApp" style="display:block;border:0;">
                </td>
                <td valign="middle" style="color:#5d6470;font-size:12px;line-height:14px;white-space:nowrap;">
                  ${telefonoVisible}
                </td>
              </tr>
            </table>
          </a>
        </td>
      </tr>`;
}

function escaparHtml(texto) {
  return String(texto)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;");
}
