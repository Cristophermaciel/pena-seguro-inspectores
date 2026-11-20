// ============================================================================
// PEÑA SEGURO 1410 - BACKEND Y APP WEB EN GOOGLE APPS SCRIPT (USUARIO + ADMIN)
// Planilla Oficial: BD_Peña_Seguro_1410_Inspectores
// ID: 1E8xewBE2aUIzy0yUfydUc1ljiVWh09WVGwsAooD-_FQ
// ============================================================================

const CONFIG = {
  SPREADSHEET_ID: "1E8xewBE2aUIzy0yUfydUc1ljiVWh09WVGwsAooD-_FQ",
  FOLDER_ID_FOTOS: "",
  FOLDER_NAME_AUTO: "Peña_Seguro_1410_Evidencias",
  EMAIL_CENTRAL: "central1410@penaflor.cl",
  EMAIL_JEFATURA: "inspeccion@penaflor.cl",
  ENVIAR_CORREOS_REALES: false
};

const SHEETS = {
  TURNOS: "Turnos_y_Checklist",
  PROCEDIMIENTOS: "Procedimientos",
  TAREAS: "Tareas_Cuadrante",
  CIERRES: "Cierres_Turno",
  USUARIOS: "Usuarios",
  CATALOGO_TAREAS: "Catalogo_Tareas",
  CONFIG_FLOTA: "Config_Flota",
  MANTENCIONES: "Mantenciones"
};

/**
 * Función principal por defecto al presionar "Ejecutar" en la barra de Apps Script.
 * Crea y verifica las 7 pestañas del Excel y la carpeta de fotos en Google Drive.
 */
function inicializarHojas() {
  const ss = obtenerSpreadsheet();

  const definiciones = [
    {
      name: SHEETS.TURNOS,
      headers: [
        "ID Turno", "Fecha/Hora Inicio", "RUT Funcionario", "Rango Horario",
        "Móvil / Unidad", "Cuadrante", "Km Inicial", "Estado Checklist",
        "Ítems Malos", "Observaciones Checklist", "Foto General / Odómetro (Drive)",
        "Foto Desperfecto (Drive)", "Detalle Completo Checklist (JSON)",
        "Estado Turno", "Fecha/Hora Término", "Km Término", "Km Recorridos"
      ]
    },
    {
      name: SHEETS.PROCEDIMIENTOS,
      headers: [
        "ID Procedimiento", "Fecha/Hora Registro", "ID Turno", "RUT Funcionario",
        "Rango Horario", "Móvil / Unidad", "Cuadrante", "Tipo de Procedimiento",
        "Dirección / Coordenadas GPS", "Observaciones / Detalle",
        "Fotografía Evidencia (Drive)", "Notificar Central 1410", "Notificar Jefatura"
      ]
    },
    {
      name: SHEETS.TAREAS,
      headers: [
        "ID Registro", "Fecha/Hora Visto Bueno", "Fecha Operativa (YYYY-MM-DD)",
        "ID Tarea", "Descripción Tarea", "ID Turno", "RUT Funcionario",
        "Rango Horario", "Móvil / Unidad", "Cuadrante",
        "Ubicación Verificación GPS", "Observaciones", "Evidencia Fotográfica (Drive)"
      ]
    },
    {
      name: SHEETS.CIERRES,
      headers: [
        "ID Cierre", "Fecha/Hora Término", "ID Turno", "RUT Funcionario",
        "Rango Horario", "Móvil / Unidad", "Cuadrante", "Km Inicial",
        "Km Término", "Total Km Recorridos"
      ]
    },
    {
      name: SHEETS.USUARIOS,
      headers: [
        "RUT", "Nombre Completo", "Cargo", "Contraseña", "Rol", "Estado", "Fecha Creación"
      ]
    },
    {
      name: SHEETS.CATALOGO_TAREAS,
      headers: [
        "ID Tarea", "Rango Horario", "Cuadrante", "Descripción de la Tarea", "Estado"
      ]
    },
    {
      name: SHEETS.CONFIG_FLOTA,
      headers: [
        "ID Unidad", "Nombre Unidad", "Tipo", "Patente / Código", "Estado Operativo",
        "Cuadrante Base", "Km Actual", "Km Última Mantención", "Fecha Última Mantención"
      ]
    },
    {
      name: SHEETS.MANTENCIONES,
      headers: [
        "ID Mantención", "Fecha/Hora", "ID Unidad", "Nombre Unidad", "Tipo",
        "Patente / Código", "Km Mantención", "Próxima Mantención (Km)",
        "Observaciones", "Registrado Por"
      ]
    }
  ];

  definiciones.forEach(def => {
    let sheet = ss.getSheetByName(def.name);
    if (!sheet) {
      sheet = ss.insertSheet(def.name);
    }
    if (sheet.getLastRow() === 0) {
      sheet.appendRow(def.headers);
      const headerRange = sheet.getRange(1, 1, 1, def.headers.length);
      headerRange.setBackground("#1e3a8a");
      headerRange.setFontColor("#ffffff");
      headerRange.setFontWeight("bold");
      sheet.setFrozenRows(1);
    } else if (sheet.getLastColumn() < def.headers.length) {
      const startCol = sheet.getLastColumn() + 1;
      const missingHeaders = def.headers.slice(startCol - 1);
      const addRange = sheet.getRange(1, startCol, 1, missingHeaders.length);
      addRange.setValues([missingHeaders]);
      addRange.setBackground("#1e3a8a");
      addRange.setFontColor("#ffffff");
      addRange.setFontWeight("bold");
    }
  });

  obtenerCarpetaEvidencias();
  Logger.log("¡Las 8 hojas y la carpeta de Drive están listas y verificadas!");
}

/**
 * Alias de seguridad por si Apps Script ejecuta myFunction() por defecto.
 */
function myFunction() {
  inicializarHojas();
}

function obtenerSpreadsheet() {
  if (CONFIG.SPREADSHEET_ID && CONFIG.SPREADSHEET_ID.trim() !== "") {
    return SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID.trim());
  }
  return SpreadsheetApp.getActiveSpreadsheet();
}

/**
 * Sirve la aplicación web completa cuando se abre el enlace /exec en el navegador,
 * o responde datos JSON si se consulta con ?action=bootstrap.
 */
function doGet(e) {
  try {
    const action = (e && e.parameter && e.parameter.action) ? e.parameter.action : "";

    if (!action) {
      const candidateFiles = ["Index", "index", "pe_a_seguro_inspectores"];
      for (let i = 0; i < candidateFiles.length; i++) {
        try {
          return HtmlService.createHtmlOutputFromFile(candidateFiles[i])
            .setTitle("Peña Seguro 1410 - Inspectores y Administración")
            .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
            .addMetaTag("viewport", "width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no");
        } catch (ignore) {}
      }

      // Si no se creó el archivo Index.html dentro de Apps Script, devolver HtmlService (evita el bloqueo multi-cuenta de ContentService)
      const statusHtml = `
        <!DOCTYPE html>
        <html lang="es">
        <head>
          <meta charset="UTF-8">
          <meta name="viewport" content="width=device-width, initial-scale=1.0">
          <title>Peña Seguro 1410 - API Conectada</title>
          <style>
            body { font-family: Arial, sans-serif; background: #0f172a; color: #fff; display: flex; align-items: center; justify-content: center; min-height: 95vh; margin: 0; padding: 16px; }
            .card { background: #1e3a8a; border: 1px solid #3b82f6; border-radius: 16px; max-width: 480px; padding: 28px; text-align: center; box-shadow: 0 15px 30px rgba(0,0,0,0.4); }
            .badge { background: #10b981; color: #fff; padding: 6px 14px; border-radius: 999px; font-size: 12px; font-weight: bold; display: inline-block; margin-bottom: 12px; }
            h1 { margin: 0 0 8px 0; font-size: 22px; }
            p { color: #bfdbfe; font-size: 14px; line-height: 1.5; }
            a { display: inline-block; margin-top: 14px; background: #facc15; color: #1e3a8a; font-weight: bold; text-decoration: none; padding: 12px 20px; border-radius: 10px; }
          </style>
        </head>
        <body>
          <div class="card">
            <div class="badge">✓ BACKEND GOOGLE APPS SCRIPT ACTIVO</div>
            <h1>PEÑA SEGURO 1410</h1>
            <p>Tu conexión con la planilla <strong>BD_Peña_Seguro_1410_Inspectores</strong> está operativa y recibiendo datos desde tu aplicación web.</p>
            <a href="https://docs.google.com/spreadsheets/d/${CONFIG.SPREADSHEET_ID}/edit" target="_blank">Abrir Planilla Excel en Drive</a>
          </div>
        </body>
        </html>`;
      return HtmlService.createHtmlOutput(statusHtml)
        .setTitle("Peña Seguro 1410 - API Activa")
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
    }

    if (action === "bootstrap" || action === "getCompletedTasks") {
      return respuestaJSON(apiGetBootstrapData());
    }

    return respuestaJSON({ ok: true, service: "Peña Seguro 1410 API" });
  } catch (err) {
    return respuestaJSON({ ok: false, error: err.toString() });
  }
}

/**
 * Recibe envíos POST desde GitHub Pages o dispositivos móviles.
 */
function doPost(e) {
  try {
    let payload = {};
    if (e.postData && e.postData.contents) {
      payload = JSON.parse(e.postData.contents);
    } else if (e.parameter && e.parameter.payload) {
      payload = JSON.parse(e.parameter.payload);
    }
    const result = apiHandleAction(payload);
    return respuestaJSON(result);
  } catch (err) {
    return respuestaJSON({ ok: false, error: err.toString() });
  }
}

/**
 * Carga todos los datos maestros desde el Excel en Drive.
 */
function apiGetBootstrapData() {
  inicializarHojas();
  const ss = obtenerSpreadsheet();
  const todayStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd");

  // 1. Usuarios
  const shUsers = ss.getSheetByName(SHEETS.USUARIOS);
  const users = [];
  if (shUsers && shUsers.getLastRow() > 1) {
    const rows = shUsers.getRange(2, 1, shUsers.getLastRow() - 1, 7).getValues();
    rows.forEach(r => {
      if (String(r[0]).trim()) {
        users.push({
          rut: String(r[0]).trim(),
          nombre: String(r[1]).trim(),
          cargo: String(r[2]).trim(),
          password: String(r[3]).trim(),
          rol: String(r[4]).trim().toUpperCase() || "INSPECTOR",
          estado: String(r[5]).trim().toUpperCase() || "ACTIVO",
          fechaCreacion: String(r[6]).trim()
        });
      }
    });
  }

  // 2. Catálogo de Tareas
  const shCat = ss.getSheetByName(SHEETS.CATALOGO_TAREAS);
  const tasksCatalog = [];
  if (shCat && shCat.getLastRow() > 1) {
    const rows = shCat.getRange(2, 1, shCat.getLastRow() - 1, 5).getValues();
    rows.forEach(r => {
      if (String(r[0]).trim()) {
        tasksCatalog.push({
          id: String(r[0]).trim(),
          rango: String(r[1]).trim().toUpperCase(),
          cuadrante: String(r[2]).trim().toUpperCase(),
          desc: String(r[3]).trim(),
          estado: String(r[4]).trim().toUpperCase() || "ACTIVA"
        });
      }
    });
  }

  // 3. Flota (incluyendo kilometraje actual y control de mantenciones)
  const shFlota = ss.getSheetByName(SHEETS.CONFIG_FLOTA);
  const fleet = [];
  const maxKmByVehicle = {};

  // Escanear últimos kilometrajes conocidos en Turnos y Cierres para mantener kmActual sincronizado
  const shTurnosKm = ss.getSheetByName(SHEETS.TURNOS);
  if (shTurnosKm && shTurnosKm.getLastRow() > 1) {
    const tRows = shTurnosKm.getRange(2, 1, shTurnosKm.getLastRow() - 1, 16).getValues();
    tRows.forEach(r => {
      const key = String(r[4] || "").trim().toUpperCase();
      if (!key) return;
      const kmIni = Number(r[6] || 0);
      const kmFin = Number(r[15] || 0);
      const best = Math.max(isNaN(kmIni) ? 0 : kmIni, isNaN(kmFin) ? 0 : kmFin);
      if (best > (maxKmByVehicle[key] || 0)) {
        maxKmByVehicle[key] = best;
      }
    });
  }
  const shCierresKm = ss.getSheetByName(SHEETS.CIERRES);
  if (shCierresKm && shCierresKm.getLastRow() > 1) {
    const cRows = shCierresKm.getRange(2, 1, shCierresKm.getLastRow() - 1, 9).getValues();
    cRows.forEach(r => {
      const key = String(r[5] || "").trim().toUpperCase();
      if (!key) return;
      const kmFin = Number(r[8] || 0);
      if (!isNaN(kmFin) && kmFin > (maxKmByVehicle[key] || 0)) {
        maxKmByVehicle[key] = kmFin;
      }
    });
  }

  if (shFlota && shFlota.getLastRow() > 1) {
    const numCols = Math.min(9, Math.max(6, shFlota.getLastColumn()));
    const rows = shFlota.getRange(2, 1, shFlota.getLastRow() - 1, numCols).getValues();
    rows.forEach(r => {
      if (String(r[0]).trim()) {
        const nombre = String(r[1]).trim();
        const storedKm = Number(r[6] || 0);
        const inferredKm = maxKmByVehicle[nombre.toUpperCase()] || 0;
        const kmActual = Math.max(isNaN(storedKm) ? 0 : storedKm, inferredKm);
        const rawUltima = r.length > 7 ? r[7] : "";
        const kmUltimaMantencion = (rawUltima !== "" && rawUltima !== null && rawUltima !== undefined && !isNaN(Number(rawUltima)))
          ? Number(rawUltima)
          : "";
        fleet.push({
          id: String(r[0]).trim(),
          nombre: nombre,
          tipo: String(r[2]).trim(),
          patente: String(r[3]).trim(),
          estado: String(r[4]).trim().toUpperCase() || "OPERATIVO",
          cuadranteBase: String(r[5]).trim(),
          kmActual: kmActual,
          kmUltimaMantencion: kmUltimaMantencion,
          fechaUltimaMantencion: r.length > 8 ? String(r[8] || "").trim() : ""
        });
      }
    });
  }

  // 4. Tareas completadas hoy
  const shTareas = ss.getSheetByName(SHEETS.TAREAS);
  const completedTaskIds = [];
  if (shTareas && shTareas.getLastRow() > 1) {
    const rows = shTareas.getRange(2, 1, shTareas.getLastRow() - 1, 4).getValues();
    rows.forEach(r => {
      const rowDate = String(r[2]).trim();
      const taskId = String(r[3]).trim();
      if (rowDate === todayStr && taskId && completedTaskIds.indexOf(taskId) === -1) {
        completedTaskIds.push(taskId);
      }
    });
  }

  // 5. Últimos turnos y procedimientos para el Monitor Admin
  const shTurnos = ss.getSheetByName(SHEETS.TURNOS);
  const recentShifts = [];
  if (shTurnos && shTurnos.getLastRow() > 1) {
    const count = Math.min(25, shTurnos.getLastRow() - 1);
    const startRow = shTurnos.getLastRow() - count + 1;
    const rows = shTurnos.getRange(startRow, 1, count, 17).getValues();
    rows.reverse().forEach(r => {
      recentShifts.push({
        turnoId: String(r[0]),
        fechaInicio: String(r[1]),
        rut: String(r[2]),
        rangoHorario: String(r[3]),
        movil: String(r[4]),
        cuadrante: String(r[5]),
        kmInicial: r[6],
        estadoChecklist: String(r[7]),
        itemsMalos: String(r[8]),
        observaciones: String(r[9]),
        fotoGeneral: String(r[10]),
        fotoDesperfecto: String(r[11]),
        estadoTurno: String(r[13]),
        fechaTermino: String(r[14]),
        kmTermino: r[15],
        kmRecorridos: r[16]
      });
    });
  }

  const shProc = ss.getSheetByName(SHEETS.PROCEDIMIENTOS);
  const recentProcedures = [];
  if (shProc && shProc.getLastRow() > 1) {
    const count = Math.min(25, shProc.getLastRow() - 1);
    const startRow = shProc.getLastRow() - count + 1;
    const rows = shProc.getRange(startRow, 1, count, 13).getValues();
    rows.reverse().forEach(r => {
      recentProcedures.push({
        procId: String(r[0]),
        fecha: String(r[1]),
        rut: String(r[3]),
        rango: String(r[4]),
        movil: String(r[5]),
        cuadrante: String(r[6]),
        tipo: String(r[7]),
        direccion: String(r[8]),
        detalle: String(r[9]),
        foto: String(r[10])
      });
    });
  }

  return {
    ok: true,
    date: todayStr,
    spreadsheetUrl: "https://docs.google.com/spreadsheets/d/" + CONFIG.SPREADSHEET_ID + "/edit",
    users: users,
    tasksCatalog: tasksCatalog,
    fleet: fleet,
    completedTaskIds: completedTaskIds,
    recentShifts: recentShifts,
    recentProcedures: recentProcedures
  };
}

/**
 * Procesa todas las acciones tanto de Inspectores como de Administradores.
 */
function apiHandleAction(payload) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);

  try {
    inicializarHojas();
    const action = payload.action;
    const ss = obtenerSpreadsheet();
    const nowFormatted = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd-MM-yyyy HH:mm:ss");
    const todayISO = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd");

    if (action === "ADMIN_SAVE_USER") {
      const sheet = ss.getSheetByName(SHEETS.USUARIOS);
      const rut = String(payload.rut || "").trim();
      if (!rut) return { ok: false, error: "RUT requerido" };

      const lastRow = sheet.getLastRow();
      let foundRow = -1;
      if (lastRow > 1) {
        const ruts = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
        for (let i = 0; i < ruts.length; i++) {
          if (String(ruts[i][0]).trim().toLowerCase() === rut.toLowerCase()) {
            foundRow = i + 2;
            break;
          }
        }
      }

      const rowValues = [
        rut,
        payload.nombre || "",
        payload.cargo || "Inspector Municipal",
        payload.password || "1234",
        (payload.rol || "INSPECTOR").toUpperCase(),
        (payload.estado || "ACTIVO").toUpperCase(),
        payload.fechaCreacion || nowFormatted.slice(0, 10)
      ];

      if (foundRow > 0) {
        sheet.getRange(foundRow, 1, 1, 7).setValues([rowValues]);
      } else {
        sheet.appendRow(rowValues);
      }
      return { ok: true, action: action, rut: rut };
    }

    if (action === "ADMIN_DELETE_USER") {
      const sheet = ss.getSheetByName(SHEETS.USUARIOS);
      const rut = String(payload.rut || "").trim().toLowerCase();
      const lastRow = sheet.getLastRow();
      if (lastRow > 1) {
        const ruts = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
        for (let i = ruts.length - 1; i >= 0; i--) {
          if (String(ruts[i][0]).trim().toLowerCase() === rut) {
            sheet.deleteRow(i + 2);
            return { ok: true, action: action };
          }
        }
      }
      return { ok: false, error: "Usuario no encontrado" };
    }

    if (action === "ADMIN_SAVE_TASK") {
      const sheet = ss.getSheetByName(SHEETS.CATALOGO_TAREAS);
      const id = String(payload.id || `t_${Date.now()}`).trim();
      const lastRow = sheet.getLastRow();
      let foundRow = -1;
      if (lastRow > 1) {
        const ids = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
        for (let i = 0; i < ids.length; i++) {
          if (String(ids[i][0]).trim() === id) {
            foundRow = i + 2;
            break;
          }
        }
      }

      const rowValues = [
        id,
        (payload.rango || "MANANA").toUpperCase(),
        (payload.cuadrante || "SRA1").toUpperCase(),
        payload.desc || "",
        (payload.estado || "ACTIVA").toUpperCase()
      ];

      if (foundRow > 0) {
        sheet.getRange(foundRow, 1, 1, 5).setValues([rowValues]);
      } else {
        sheet.appendRow(rowValues);
      }
      return { ok: true, action: action, id: id };
    }

    if (action === "ADMIN_DELETE_TASK") {
      const sheet = ss.getSheetByName(SHEETS.CATALOGO_TAREAS);
      const id = String(payload.id || "").trim();
      const lastRow = sheet.getLastRow();
      if (lastRow > 1) {
        const ids = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
        for (let i = ids.length - 1; i >= 0; i--) {
          if (String(ids[i][0]).trim() === id) {
            sheet.deleteRow(i + 2);
            return { ok: true, action: action };
          }
        }
      }
      return { ok: false, error: "Tarea no encontrada" };
    }

    if (action === "ADMIN_SAVE_VEHICLE") {
      const sheet = ss.getSheetByName(SHEETS.CONFIG_FLOTA);
      const id = String(payload.id || `UND-${Date.now()}`).trim();
      const lastRow = sheet.getLastRow();
      let foundRow = -1;
      if (lastRow > 1) {
        const ids = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
        for (let i = 0; i < ids.length; i++) {
          if (String(ids[i][0]).trim() === id) {
            foundRow = i + 2;
            break;
          }
        }
      }

      const kmActualVal = (payload.kmActual !== undefined && payload.kmActual !== "" && !isNaN(Number(payload.kmActual)))
        ? Number(payload.kmActual)
        : 0;
      const kmUltimaVal = (payload.kmUltimaMantencion !== undefined && payload.kmUltimaMantencion !== "" && !isNaN(Number(payload.kmUltimaMantencion)))
        ? Number(payload.kmUltimaMantencion)
        : "";
      const fechaUltimaVal = String(payload.fechaUltimaMantencion || "").trim();

      const rowValues = [
        id,
        payload.nombre || "",
        payload.tipo || "Móvil",
        payload.patente || "",
        (payload.estado || "OPERATIVO").toUpperCase(),
        payload.cuadranteBase || "SRA1",
        kmActualVal,
        kmUltimaVal,
        fechaUltimaVal
      ];

      if (foundRow > 0) {
        sheet.getRange(foundRow, 1, 1, 9).setValues([rowValues]);
      } else {
        sheet.appendRow(rowValues);
      }
      return { ok: true, action: action, id: id };
    }

    if (action === "ADMIN_LOG_MAINTENANCE") {
      const sheetFlota = ss.getSheetByName(SHEETS.CONFIG_FLOTA);
      const sheetMant = ss.getSheetByName(SHEETS.MANTENCIONES);
      const id = String(payload.id || "").trim();
      const nombre = String(payload.nombre || "").trim();
      const tipo = String(payload.tipo || "Móvil").trim();
      const patente = String(payload.patente || "").trim();
      const kmMant = Number(payload.kmMantencion || payload.kmActual || 0);
      const isMoto = tipo.toLowerCase().indexOf("moto") >= 0 || nombre.toLowerCase().indexOf("tango") >= 0;
      const intervalo = isMoto ? 3000 : 10000;
      const kmProxima = kmMant + intervalo;
      const fechaMant = String(payload.fechaMantencion || nowFormatted).trim();
      const obs = String(payload.observaciones || (isMoto ? "Cambio de aceite realizado (3.000 km)" : "Mantención preventiva realizada (10.000 km)")).trim();
      const registradoPor = String(payload.registradoPor || "Administrador").trim();

      const lastRow = sheetFlota.getLastRow();
      let foundRow = -1;
      if (lastRow > 1) {
        const rows = sheetFlota.getRange(2, 1, lastRow - 1, 2).getValues();
        for (let i = 0; i < rows.length; i++) {
          if ((id && String(rows[i][0]).trim() === id) || (nombre && String(rows[i][1]).trim().toLowerCase() === nombre.toLowerCase())) {
            foundRow = i + 2;
            break;
          }
        }
      }

      if (foundRow > 0) {
        sheetFlota.getRange(foundRow, 5).setValue((payload.estado || "OPERATIVO").toUpperCase());
        sheetFlota.getRange(foundRow, 7, 1, 3).setValues([[kmMant, kmMant, fechaMant]]);
      } else if (nombre) {
        sheetFlota.appendRow([
          id || `UND-${Date.now()}`,
          nombre,
          tipo,
          patente,
          "OPERATIVO",
          payload.cuadranteBase || "252A",
          kmMant,
          kmMant,
          fechaMant
        ]);
      }

      if (sheetMant) {
        sheetMant.appendRow([
          `MNT-${Date.now()}`,
          fechaMant,
          id,
          nombre,
          tipo,
          patente,
          kmMant,
          kmProxima,
          obs,
          registradoPor
        ]);
      }

      return { ok: true, action: action, id: id, kmMantencion: kmMant, kmProxima: kmProxima };
    }

    if (action === "ADMIN_DELETE_VEHICLE") {
      const sheet = ss.getSheetByName(SHEETS.CONFIG_FLOTA);
      const id = String(payload.id || "").trim();
      const lastRow = sheet.getLastRow();
      if (lastRow > 1) {
        const ids = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
        for (let i = ids.length - 1; i >= 0; i--) {
          if (String(ids[i][0]).trim() === id) {
            sheet.deleteRow(i + 2);
            return { ok: true, action: action };
          }
        }
      }
      return { ok: false, error: "Vehículo no encontrado" };
    }

    if (action === "INICIO_TURNO_CHECKLIST") {
      const sheet = ss.getSheetByName(SHEETS.TURNOS);

      const urlFotoGeneral = payload.fotoGeneralBase64
        ? guardarFotoEnDrive(payload.fotoGeneralBase64, `Turno_${payload.turnoId}_General.jpg`)
        : "Sin foto";

      const urlFotoDesperfecto = payload.fotoDesperfectoBase64
        ? guardarFotoEnDrive(payload.fotoDesperfectoBase64, `Turno_${payload.turnoId}_Desperfecto.jpg`)
        : "Sin desperfectos";

      sheet.appendRow([
        payload.turnoId || `T-${Date.now()}`,
        payload.fechaInicio || nowFormatted,
        payload.rut || "",
        payload.rangoHorario || "",
        payload.movil || "",
        payload.cuadrante || "",
        Number(payload.kmInicial || 0),
        payload.estadoChecklist || "OK",
        payload.itemsMalos || "Ninguno",
        payload.observaciones || "",
        urlFotoGeneral,
        urlFotoDesperfecto,
        JSON.stringify(payload.respuestasChecklist || {}),
        "ACTIVO",
        "",
        "",
        ""
      ]);

      actualizarKmFlota(ss, payload.movil, Number(payload.kmInicial || 0));

      return {
        ok: true,
        action: action,
        turnoId: payload.turnoId,
        urlFotoGeneral: urlFotoGeneral,
        urlFotoDesperfecto: urlFotoDesperfecto
      };
    }

    if (action === "REGISTRAR_PROCEDIMIENTO") {
      const sheet = ss.getSheetByName(SHEETS.PROCEDIMIENTOS);
      const procId = payload.procId || `PROC-${Date.now()}`;

      const urlFoto = payload.fotoBase64
        ? guardarFotoEnDrive(payload.fotoBase64, `${procId}_${payload.tipo || "Proc"}.jpg`)
        : "Sin foto";

      sheet.appendRow([
        procId,
        payload.fechaRegistro || nowFormatted,
        payload.turnoId || "",
        payload.rut || "",
        payload.rangoHorario || "",
        payload.movil || "",
        payload.cuadrante || "",
        payload.tipo || "",
        payload.direccion || "",
        payload.detalle || "",
        urlFoto,
        payload.notificarCentral ? "SÍ" : "NO",
        payload.notificarJefatura ? "SÍ" : "NO"
      ]);

      if (CONFIG.ENVIAR_CORREOS_REALES && (payload.notificarCentral || payload.notificarJefatura)) {
        enviarNotificacionCorreo(payload, urlFoto, nowFormatted);
      }

      return { ok: true, action: action, procId: procId, urlFoto: urlFoto };
    }

    if (action === "RESOLVER_TAREA") {
      const sheet = ss.getSheetByName(SHEETS.TAREAS);
      const regId = `TAR-${Date.now()}`;

      const urlFoto = payload.fotoBase64
        ? guardarFotoEnDrive(payload.fotoBase64, `${regId}_${payload.taskId}.jpg`)
        : "Sin foto";

      sheet.appendRow([
        regId,
        payload.fechaRegistro || nowFormatted,
        todayISO,
        payload.taskId || "",
        payload.taskDesc || "",
        payload.turnoId || "",
        payload.rut || "",
        payload.rangoHorario || "",
        payload.movil || "",
        payload.cuadrante || "",
        payload.direccion || "",
        payload.detalle || "",
        urlFoto
      ]);

      return { ok: true, action: action, regId: regId, urlFoto: urlFoto };
    }

    if (action === "CERRAR_TURNO") {
      const sheetCierres = ss.getSheetByName(SHEETS.CIERRES);
      const sheetTurnos = ss.getSheetByName(SHEETS.TURNOS);
      const kmIni = Number(payload.kmInicial || 0);
      const kmFin = Number(payload.kmTermino || 0);
      const kmRecorridos = Math.max(0, kmFin - kmIni);

      sheetCierres.appendRow([
        `CIE-${Date.now()}`,
        payload.fechaTermino || nowFormatted,
        payload.turnoId || "",
        payload.rut || "",
        payload.rangoHorario || "",
        payload.movil || "",
        payload.cuadrante || "",
        kmIni,
        kmFin,
        kmRecorridos
      ]);

      if (payload.turnoId && sheetTurnos.getLastRow() > 1) {
        const ids = sheetTurnos.getRange(2, 1, sheetTurnos.getLastRow() - 1, 1).getValues();
        for (let i = ids.length - 1; i >= 0; i--) {
          if (String(ids[i][0]) === String(payload.turnoId)) {
            const rowIndex = i + 2;
            sheetTurnos.getRange(rowIndex, 14, 1, 4).setValues([[
              "FINALIZADO",
              payload.fechaTermino || nowFormatted,
              kmFin,
              kmRecorridos
            ]]);
            break;
          }
        }
      }

      actualizarKmFlota(ss, payload.movil, kmFin);

      return { ok: true, action: action, kmRecorridos: kmRecorridos };
    }

    return { ok: false, error: "Acción no reconocida: " + action };
  } catch (err) {
    return { ok: false, error: err.toString() };
  } finally {
    lock.releaseLock();
  }
}

/**
 * Actualiza automáticamente el Km Actual del vehículo en la hoja Config_Flota
 * cada vez que un inspector abre o cierra turno con un kilometraje mayor.
 */
function actualizarKmFlota(ss, movilNombre, kmNuevo) {
  try {
    const nombre = String(movilNombre || "").trim();
    const km = Number(kmNuevo || 0);
    if (!nombre || isNaN(km) || km <= 0) return;

    const shFlota = ss.getSheetByName(SHEETS.CONFIG_FLOTA);
    if (!shFlota || shFlota.getLastRow() <= 1) return;

    const numCols = Math.min(9, Math.max(6, shFlota.getLastColumn()));
    const rows = shFlota.getRange(2, 1, shFlota.getLastRow() - 1, numCols).getValues();
    for (let i = 0; i < rows.length; i++) {
      if (String(rows[i][1] || "").trim().toLowerCase() === nombre.toLowerCase()) {
        const rowIndex = i + 2;
        const tipo = String(rows[i][2] || "Móvil").trim();
        const kmPrevio = Number(rows[i][6] || 0);
        const rawUltima = rows[i].length > 7 ? rows[i][7] : "";
        const isMoto = tipo.toLowerCase().indexOf("moto") >= 0 || nombre.toLowerCase().indexOf("tango") >= 0;
        const intervalo = isMoto ? 3000 : 10000;

        if (km > kmPrevio) {
          shFlota.getRange(rowIndex, 7).setValue(km);
          if (rawUltima === "" || rawUltima === null || rawUltima === undefined) {
            const refKm = kmPrevio > 0 ? kmPrevio : km;
            const anchored = Math.floor(Math.max(0, refKm - 1) / intervalo) * intervalo;
            shFlota.getRange(rowIndex, 8).setValue(anchored);
          }
        }
        break;
      }
    }
  } catch (ignore) {}
}

function guardarFotoEnDrive(dataUrl, fileName) {
  try {
    const folder = obtenerCarpetaEvidencias();
    const parts = dataUrl.split(",");
    const base64Data = parts.length > 1 ? parts[1] : parts[0];
    const mimeMatch = parts[0].match(/:(.*?);/);
    const contentType = mimeMatch ? mimeMatch[1] : "image/jpeg";

    const decoded = Utilities.base64Decode(base64Data);
    const blob = Utilities.newBlob(decoded, contentType, fileName);
    const file = folder.createFile(blob);

    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    return file.getUrl();
  } catch (e) {
    return "Error al guardar foto: " + e.toString();
  }
}

function obtenerCarpetaEvidencias() {
  if (CONFIG.FOLDER_ID_FOTOS && CONFIG.FOLDER_ID_FOTOS.trim() !== "") {
    return DriveApp.getFolderById(CONFIG.FOLDER_ID_FOTOS.trim());
  }
  const folders = DriveApp.getFoldersByName(CONFIG.FOLDER_NAME_AUTO);
  if (folders.hasNext()) {
    return folders.next();
  }
  return DriveApp.createFolder(CONFIG.FOLDER_NAME_AUTO);
}

function enviarNotificacionCorreo(payload, urlFoto, fechaHora) {
  const destinatarios = [];
  if (payload.notificarCentral && CONFIG.EMAIL_CENTRAL) destinatarios.push(CONFIG.EMAIL_CENTRAL);
  if (payload.notificarJefatura && CONFIG.EMAIL_JEFATURA) destinatarios.push(CONFIG.EMAIL_JEFATURA);
  if (destinatarios.length === 0) return;

  const subject = `[PEÑA SEGURO 1410] ${payload.tipo} - Cuadrante ${payload.cuadrante} (${payload.movil})`;
  const body = [
    "REPORTE DE PROCEDIMIENTO OPERATIVO - PEÑA SEGURO 1410",
    "-----------------------------------------------------",
    "Fecha y Hora: " + fechaHora,
    "RUT Inspector: " + payload.rut,
    "Móvil / Unidad: " + payload.movil,
    "Cuadrante: " + payload.cuadrante,
    "Rango Horario: " + payload.rangoHorario,
    "Tipo Procedimiento: " + payload.tipo,
    "Dirección / GPS: " + payload.direccion,
    "Observaciones: " + (payload.detalle || "Sin observaciones"),
    "Evidencia Fotográfica: " + urlFoto
  ].join("\n");

  MailApp.sendEmail(destinatarios.join(","), subject, body);
}

function respuestaJSON(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
