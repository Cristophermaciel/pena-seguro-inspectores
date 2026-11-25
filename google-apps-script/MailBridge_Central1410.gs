// ============================================================================
// PUENTE OFICIAL DE CORREO — CENTRAL 1410 PEÑA SEGURO (central1410@penaflor.cl)
// Soporta HTML oficial + Imágenes incrustadas (inline) y adjuntas reales (v43)
// ============================================================================

const DEFAULT_BRIDGE_SECRET = "PENA1410-MAIL-BRIDGE";

function doGet() {
  return ContentService
    .createTextOutput(JSON.stringify({
      ok: true,
      service: "Puente Correo Central 1410 Peña Seguro (v43 con imágenes)",
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
    let htmlBody = String(data.htmlBody || "").trim();
    const textBody = String(data.textBody || "Copia de procedimiento Peña Seguro 1410.").trim();

    if (!to) {
      return jsonOut({ ok: false, error: "No se especificaron destinatarios (to)." });
    }

    const attachments = [];
    const inlineImages = {};
    const imgs = Array.isArray(data.imagesBase64) ? data.imagesBase64 : [];
    imgs.slice(0, 4).forEach((item, idx) => {
      try {
        const b64 = typeof item === "string" ? item : (item && item.dataUrl ? item.dataUrl : "");
        if (!b64 || b64.indexOf("data:image/") !== 0) return;
        const parts = b64.split(",");
        if (parts.length < 2) return;
        const mimeMatch = parts[0].match(/data:(.*?);/);
        const mime = mimeMatch ? mimeMatch[1] : "image/jpeg";
        const ext = mime.indexOf("png") >= 0 ? "png" : "jpg";
        const fname = (item && item.name) ? item.name : ("Evidencia_" + (idx + 1) + "." + ext);
        const blob = Utilities.newBlob(Utilities.base64Decode(parts[1]), mime, fname);
        attachments.push(blob);
        inlineImages["evidencia_" + idx] = blob;
      } catch (ignoreImg) {}
    });

    const options = {
      to: to,
      subject: subject,
      body: textBody,
      htmlBody: htmlBody || undefined,
      name: String(data.senderName || "Central 1410 — Peña Seguro").trim(),
      replyTo: String(data.replyTo || "central1410@penaflor.cl").trim()
    };
    if (data.cc) options.cc = String(data.cc).trim();
    if (attachments.length > 0) {
      options.attachments = attachments;
      options.inlineImages = inlineImages;
    }

    MailApp.sendEmail(options);

    return jsonOut({
      ok: true,
      sentFrom: Session.getActiveUser().getEmail() || "central1410@penaflor.cl",
      to: to,
      imagesSent: attachments.length,
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
