// ============================================================================
// PUENTE OFICIAL DE CORREO — CENTRAL 1410 PEÑA SEGURO (central1410@penaflor.cl)
// ============================================================================
// INSTRUCCIONES DE INSTALACIÓN (1 sola vez, en la cuenta central1410@penaflor.cl):
// 1. Inicia sesión en https://script.google.com con central1410@penaflor.cl
// 2. Crea un "Nuevo proyecto" llamado: Puente_Correo_Central1410
// 3. Pega este código completo y guarda (Ctrl+S)
// 4. Haz clic en Implementar -> Nueva implementación -> Aplicación web
//    - Ejecutar como: Yo (central1410@penaflor.cl)
//    - Quién tiene acceso: Cualquier usuario
// 5. Autoriza los permisos de correo y copia la URL /exec resultante
// 6. Pega esa URL en el Panel Administrador -> Pestaña "Correos" -> Puente central1410@penaflor.cl
// ============================================================================

const DEFAULT_BRIDGE_SECRET = "PENA1410-MAIL-BRIDGE";

function doGet() {
  return ContentService
    .createTextOutput(JSON.stringify({
      ok: true,
      service: "Puente Correo Central 1410 Peña Seguro",
      sender: Session.getActiveUser().getEmail() || "central1410@penaflor.cl",
      remainingDailyQuota: MailApp.getRemainingDailyQuota()
    }))
    .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  try {
    const raw = (e && e.postData && e.postData.contents) ? e.postData.contents : "{}";
    const data = JSON.parse(raw);

    const expectedSecret = PropertiesService.getScriptProperties().getProperty("MAIL_BRIDGE_SECRET") || DEFAULT_BRIDGE_SECRET;
    if (String(data.secret || "").trim() !== String(expectedSecret).trim()) {
      return jsonOut({ ok: false, error: "Clave secreta inválida en puente de correo." });
    }

    const to = String(data.to || "").trim();
    const subject = String(data.subject || "Copia Oficial de Procedimiento — Peña Seguro 1410").trim();
    const htmlBody = String(data.htmlBody || "").trim();
    const textBody = String(data.textBody || "Copia de procedimiento Peña Seguro 1410.").trim();

    if (!to) {
      return jsonOut({ ok: false, error: "No se especificaron destinatarios (to)." });
    }

    const options = {
      to: to,
      subject: subject,
      body: textBody,
      htmlBody: htmlBody || undefined,
      name: String(data.senderName || "Central 1410 — Peña Seguro").trim(),
      replyTo: String(data.replyTo || "central1410@penaflor.cl").trim()
    };
    if (data.cc) options.cc = String(data.cc).trim();

    MailApp.sendEmail(options);

    return jsonOut({
      ok: true,
      sentFrom: Session.getActiveUser().getEmail() || "central1410@penaflor.cl",
      to: to,
      remainingDailyQuota: MailApp.getRemainingDailyQuota()
    });
  } catch (err) {
    return jsonOut({ ok: false, error: err.toString() });
  }
}

function jsonOut(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
