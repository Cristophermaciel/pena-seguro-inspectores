// ============================================================================
// PEÑA SEGURO 1410 - BACKEND Y APP WEB EN GOOGLE APPS SCRIPT (USUARIO + ADMIN)
// Planilla Oficial: BD_Peña_Seguro_1410_Inspectores
// ID: 1E8xewBE2aUIzy0yUfydUc1ljiVWh09WVGwsAooD-_FQ
// ============================================================================

const CONFIG = {
  SPREADSHEET_ID: "1E8xewBE2aUIzy0yUfydUc1ljiVWh09WVGwsAooD-_FQ",
  FOLDER_ID_FOTOS: "",
  FOLDER_NAME_AUTO: "Peña_Seguro_1410_Evidencias",
  FOLDER_NAME_PDF_CENCO: "Peña_Seguro_1410_Actas_CENCO_PDF",
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
  MANTENCIONES: "Mantenciones",
  ALERTAS_FLOTA: "Alertas_Flota",
  // --- Hojas del Sistema CAD CENCO 1410 ---
  CENCO_BITACORA: "CENCO_Bitacora_CAD",
  CENCO_TURNOS: "CENCO_Turnos",
  CENCO_CCTV_ALERTAS: "CENCO_Alertas_CCTV",
  CENCO_CAMARAS: "CENCO_Catalogo_Camaras",
  CENCO_DESPACHOS: "CENCO_Despachos"
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
    },
    {
      name: SHEETS.ALERTAS_FLOTA,
      headers: [
        "ID Alerta", "Fecha/Hora", "Móvil / Unidad", "Cuadrante", "Inspector",
        "Categoría", "Prioridad", "Ítem Checklist", "Estado Reportado",
        "Acción Recomendada", "Observaciones", "Foto Evidencia",
        "Estado Gestión", "Fecha Resolución", "Resuelto Por", "Nota Resolución"
      ]
    },
    {
      name: SHEETS.CENCO_BITACORA,
      headers: [
        "ID Folio CAD", "Fecha/Hora Ingreso", "Origen / Canal", "Estado CAD",
        "Prioridad", "Categoría", "Subcategoría", "Cuadrante",
        "Dirección / Sector", "Referencia / Cámara CCTV", "Nombre Denunciante",
        "Teléfono Denunciante", "Detalle / Relato", "Ingresado Por",
        "Despachador Asignado", "Móvil(es) Despachado(s)", "Hora Despacho",
        "Hora Arribo", "Hora Cierre", "Resultado / Resolución",
        "Imágenes Evidencia (Drive)", "ID Turno CENCO", "ID Proc Inspector",
        "Historial Cadena (JSON)"
      ]
    },
    {
      name: SHEETS.CENCO_TURNOS,
      headers: [
        "ID Turno CENCO", "Fecha/Hora Inicio", "RUT / Usuario", "Nombre Operador",
        "Rol CENCO", "Puesto / Consola", "Rango Horario", "Novedades Inicio Turno",
        "Estado Turno", "Fecha/Hora Cierre", "Total Ingresados", "Total Despachados",
        "Total Finalizados", "Resumen Cierre Turno", "URL PDF Acta Turno (Drive)"
      ]
    },
    {
      name: SHEETS.CENCO_CCTV_ALERTAS,
      headers: [
        "ID Alerta CCTV", "Fecha/Hora", "ID Cámara", "Nombre / Ubicación Cámara",
        "Cuadrante", "Categoría Evento", "Prioridad", "Descripción Visual",
        "Operador CCTV", "Imágenes Captura CCTV (Drive)", "Estado",
        "Móvil Cercano Sugerido", "Folio CAD Vinculado", "Atendido Por"
      ]
    },
    {
      name: SHEETS.CENCO_CAMARAS,
      headers: [
        "ID Cámara", "Nombre / Intersección", "Cuadrante", "Sector / Villa",
        "Tipo Cámara", "Estado", "Coordenadas GPS", "Observaciones"
      ]
    },
    {
      name: SHEETS.CENCO_DESPACHOS,
      headers: [
        "ID Despacho", "Fecha/Hora Despacho", "ID Folio CAD", "Móvil / Unidad",
        "Cuadrante", "Despachador", "Estado Unidad", "Hora Arribo",
        "Hora Liberación", "Observación Despacho"
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
      headerRange.setBackground(def.name.indexOf("CENCO_") === 0 ? "#0f172a" : "#1e3a8a");
      headerRange.setFontColor("#ffffff");
      headerRange.setFontWeight("bold");
      sheet.setFrozenRows(1);
    } else if (sheet.getLastColumn() < def.headers.length) {
      const startCol = sheet.getLastColumn() + 1;
      const missingHeaders = def.headers.slice(startCol - 1);
      const addRange = sheet.getRange(1, startCol, 1, missingHeaders.length);
      addRange.setValues([missingHeaders]);
      addRange.setBackground(def.name.indexOf("CENCO_") === 0 ? "#0f172a" : "#1e3a8a");
      addRange.setFontColor("#ffffff");
      addRange.setFontWeight("bold");
    }
  });

  // Sembrar catálogo base de cámaras CCTV si está vacío
  const shCam = ss.getSheetByName(SHEETS.CENCO_CAMARAS);
  if (shCam && shCam.getLastRow() <= 1) {
    const defaultCams = [
      ["CAM-01", "Plaza de Armas / Av. Vicuña Mackenna con Alcalde Luis Araya", "252A", "Peñaflor Centro", "PTZ 360°", "OPERATIVA", "-33.6063, -70.8768", "Cobertura Plaza y Bancos"],
      ["CAM-02", "Manuel Rodríguez con Av. Caupolicán", "252A", "Peñaflor Centro", "PTZ 360°", "OPERATIVA", "-33.6085, -70.8792", "Control de flujo vehicular y colegios"],
      ["CAM-03", "CESFAM Peñaflor / SAR Alcalde Luis Araya Cereceda", "252A", "Villa San Fernando", "Fija HD", "OPERATIVA", "-33.6039, -70.8745", "Acceso recinto salud"],
      ["CAM-04", "Av. Malloquito con 14 de Julio", "252B", "Población 14 de Julio", "PTZ 360°", "OPERATIVA", "-33.6152, -70.8721", "Sector Malloquito y Villa Duralitte"],
      ["CAM-05", "Cancha Los Diamelos Sur / Av. del Parque", "252B", "Villa Los Diamelos", "PTZ 360°", "OPERATIVA", "-33.6190, -70.8695", "Foco preventivo nocturno"],
      ["CAM-06", "Acceso Parque Municipal El Trapiche / Camino El Trapiche", "252C", "El Trapiche", "LPR Patentes", "OPERATIVA", "-33.6124, -70.8920", "Lectura patentes acceso poniente"],
      ["CAM-07", "Av. Las Palmeras con Villa Alameda", "252C", "Villa Alameda", "PTZ 360°", "OPERATIVA", "-33.6098, -70.8884", "Vigilancia especial Las Palmeras"],
      ["CAM-08", "Plaza de Malloco / Av. Balmaceda con Toribio Larraín", "253A", "Malloco Norte", "PTZ 360°", "OPERATIVA", "-33.6130, -70.8542", "Cobertura Plaza Malloco y comercio"],
      ["CAM-09", "CESFAM Dr. Fernando Monckeberg / Los Bosques", "253A", "Villa Los Bosques", "Fija HD", "OPERATIVA", "-33.6105, -70.8510", "Acceso CESFAM Malloco"],
      ["CAM-10", "Centro Seguridad Comunitario Las Praderas (Sector Amarillo)", "253B", "Población Las Praderas", "PTZ 360°", "OPERATIVA", "-33.6215, -70.8568", "Monitoreo permanente Las Praderas"],
      ["CAM-11", "Av. La Paz con Punta Arenas / Villa Berlín", "253B", "Malloco Sur", "PTZ 360°", "OPERATIVA", "-33.6240, -70.8595", "Cruce comercial y florerías"],
      ["CAM-12", "Camino a Pelvín / Cruce Acceso Rural", "SRA", "Sector Pelvín", "LPR Patentes", "OPERATIVA", "-33.5910, -70.8950", "Control ingreso rural norte"],
      ["CAM-13", "Camino El Guanaco / Peñaflor Alto", "SRB", "Peñaflor Alto", "PTZ 360°", "OPERATIVA", "-33.5880, -70.8710", "Vigilancia sector alto"],
      ["CAM-14", "Ruta G-78 Camino a Valparaíso con Calle Bilbao", "SRC", "Camino Valparaíso", "LPR Patentes", "OPERATIVA", "-33.6045, -70.8615", "Pórtico lector de patentes Ruta G-78"],
      ["CAM-15", "Cuesta de Mallarauco / Cruce El Sotillo", "SRA1", "Mallarauco", "LPR Patentes", "OPERATIVA", "-33.6010, -70.9180", "Control acceso poniente Mallarauco"]
    ];
    shCam.getRange(2, 1, defaultCams.length, 8).setValues(defaultCams);
  }

  obtenerCarpetaEvidencias();
  Logger.log("¡Todas las hojas de Inspectores y CENCO CAD (14 pestañas) y Drive están listas!");
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

  // 5. Últimos turnos y procedimientos para el Monitor Admin y CENCO CAD
  // Cruzar con Cierres_Turno para garantizar que ningún turno cerrado aparezca como ACTIVO
  const closedTurnoIds = {};
  const shCierresCheck = ss.getSheetByName(SHEETS.CIERRES);
  if (shCierresCheck && shCierresCheck.getLastRow() > 1) {
    const cCount = Math.min(200, shCierresCheck.getLastRow() - 1);
    const cStart = shCierresCheck.getLastRow() - cCount + 1;
    const cIds = shCierresCheck.getRange(cStart, 3, cCount, 1).getValues();
    cIds.forEach(cr => {
      const tid = String(cr[0] || "").trim();
      if (tid) closedTurnoIds[tid] = true;
    });
  }

  const shTurnos = ss.getSheetByName(SHEETS.TURNOS);
  const recentShifts = [];
  if (shTurnos && shTurnos.getLastRow() > 1) {
    const count = Math.min(60, shTurnos.getLastRow() - 1);
    const startRow = shTurnos.getLastRow() - count + 1;
    const rows = shTurnos.getRange(startRow, 1, count, 17).getValues();
    rows.reverse().forEach(r => {
      const tid = String(r[0] || "").trim();
      const rawEstado = String(r[13] || "").trim().toUpperCase();
      const fechaFin = String(r[14] || "").trim();
      const kmFinVal = Number(r[15] || 0);
      const isClosed = rawEstado === "FINALIZADO" || rawEstado === "CERRADO" || Boolean(fechaFin) || kmFinVal > 0 || Boolean(tid && closedTurnoIds[tid]);
      const estadoChecklistStr = String(r[7] || "").trim();
      const detailChecklistStr = String(r[12] || "").trim();
      recentShifts.push({
        turnoId: tid,
        fechaInicio: String(r[1]),
        rut: String(r[2]),
        rangoHorario: String(r[3]),
        movil: String(r[4]),
        cuadrante: String(r[5]),
        kmInicial: r[6],
        estadoChecklist: estadoChecklistStr,
        hasChecklist: Boolean(estadoChecklistStr || detailChecklistStr),
        itemsMalos: String(r[8]),
        observaciones: String(r[9]),
        fotoGeneral: String(r[10]),
        fotoDesperfecto: String(r[11]),
        estadoTurno: isClosed ? "FINALIZADO" : (rawEstado || "ACTIVO"),
        fechaTermino: fechaFin,
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

  // 6. Alertas de Flota desde Checklist (Alertas_Flota)
  const shAlertas = ss.getSheetByName(SHEETS.ALERTAS_FLOTA);
  const fleetAlerts = [];
  if (shAlertas && shAlertas.getLastRow() > 1) {
    const count = Math.min(120, shAlertas.getLastRow() - 1);
    const startRow = shAlertas.getLastRow() - count + 1;
    const rows = shAlertas.getRange(startRow, 1, count, 16).getValues();
    rows.reverse().forEach(r => {
      if (String(r[0]).trim()) {
        fleetAlerts.push({
          id: String(r[0]).trim(),
          fecha: String(r[1]).trim(),
          movil: String(r[2]).trim(),
          cuadrante: String(r[3]).trim(),
          inspector: String(r[4]).trim(),
          categoria: String(r[5]).trim(),
          prioridad: String(r[6]).trim().toUpperCase() || "MEDIA",
          itemLabel: String(r[7]).trim(),
          estadoReportado: String(r[8]).trim(),
          accionRecomendada: String(r[9]).trim(),
          observaciones: String(r[10]).trim(),
          foto: String(r[11]).trim(),
          estadoGestion: String(r[12]).trim().toUpperCase() || "PENDIENTE",
          fechaResolucion: String(r[13]).trim(),
          resueltoPor: String(r[14]).trim(),
          notaResolucion: String(r[15]).trim()
        });
      }
    });
  }

  // 7. Datos CENCO CAD (Bitácora, Turnos, Alertas CCTV, Cámaras, Despachos)
  const shCencoBit = ss.getSheetByName(SHEETS.CENCO_BITACORA);
  const cencoIncidents = [];
  if (shCencoBit && shCencoBit.getLastRow() > 1) {
    const count = Math.min(200, shCencoBit.getLastRow() - 1);
    const startRow = shCencoBit.getLastRow() - count + 1;
    const numCols = Math.min(24, Math.max(20, shCencoBit.getLastColumn()));
    const rows = shCencoBit.getRange(startRow, 1, count, numCols).getValues();
    rows.reverse().forEach(r => {
      if (String(r[0]).trim()) {
        let historial = [];
        try {
          historial = r[23] ? JSON.parse(String(r[23])) : [];
        } catch (ignore) { historial = []; }
        const rawImgs = String(r[20] || "").trim();
        const imagenes = rawImgs && rawImgs !== "Sin foto"
          ? rawImgs.split(/\s*\|\s*/).filter(Boolean)
          : [];
        cencoIncidents.push({
          folioId: String(r[0]).trim(),
          fechaIngreso: String(r[1]).trim(),
          origen: String(r[2]).trim() || "LLAMADO 1410",
          estado: String(r[3]).trim().toUpperCase() || "PENDIENTE_DESPACHO",
          prioridad: String(r[4]).trim().toUpperCase() || "MEDIA",
          categoria: String(r[5]).trim(),
          subcategoria: String(r[6]).trim(),
          cuadrante: String(r[7]).trim(),
          direccion: String(r[8]).trim(),
          referenciaCamara: String(r[9]).trim(),
          denunciante: String(r[10]).trim(),
          telefono: String(r[11]).trim(),
          detalle: String(r[12]).trim(),
          ingresadoPor: String(r[13]).trim(),
          despachador: String(r[14]).trim(),
          movilesAsignados: String(r[15]).trim(),
          horaDespacho: String(r[16]).trim(),
          horaArribo: String(r[17]).trim(),
          horaCierre: String(r[18]).trim(),
          resolucion: String(r[19]).trim(),
          imagenes: imagenes,
          turnoCencoId: String(r[21] || "").trim(),
          procInspectorId: String(r[22] || "").trim(),
          historial: Array.isArray(historial) ? historial : []
        });
      }
    });
  }

  const shCencoTurnos = ss.getSheetByName(SHEETS.CENCO_TURNOS);
  const cencoShifts = [];
  if (shCencoTurnos && shCencoTurnos.getLastRow() > 1) {
    const count = Math.min(60, shCencoTurnos.getLastRow() - 1);
    const startRow = shCencoTurnos.getLastRow() - count + 1;
    const rows = shCencoTurnos.getRange(startRow, 1, count, 15).getValues();
    rows.reverse().forEach(r => {
      if (String(r[0]).trim()) {
        cencoShifts.push({
          turnoCencoId: String(r[0]).trim(),
          fechaInicio: String(r[1]).trim(),
          rut: String(r[2]).trim(),
          operador: String(r[3]).trim(),
          rolCenco: String(r[4]).trim(),
          puesto: String(r[5]).trim(),
          rangoHorario: String(r[6]).trim(),
          novedadesInicio: String(r[7]).trim(),
          estado: String(r[8]).trim().toUpperCase() || "ACTIVO",
          fechaCierre: String(r[9]).trim(),
          totalIngresados: Number(r[10] || 0),
          totalDespachados: Number(r[11] || 0),
          totalFinalizados: Number(r[12] || 0),
          resumenCierre: String(r[13]).trim(),
          pdfUrl: String(r[14]).trim()
        });
      }
    });
  }

  const shCctv = ss.getSheetByName(SHEETS.CENCO_CCTV_ALERTAS);
  const cencoCctvAlerts = [];
  if (shCctv && shCctv.getLastRow() > 1) {
    const count = Math.min(100, shCctv.getLastRow() - 1);
    const startRow = shCctv.getLastRow() - count + 1;
    const rows = shCctv.getRange(startRow, 1, count, 14).getValues();
    rows.reverse().forEach(r => {
      if (String(r[0]).trim()) {
        const rawImgs = String(r[9] || "").trim();
        const imagenes = rawImgs && rawImgs !== "Sin foto"
          ? rawImgs.split(/\s*\|\s*/).filter(Boolean)
          : [];
        cencoCctvAlerts.push({
          alertaId: String(r[0]).trim(),
          fecha: String(r[1]).trim(),
          camaraId: String(r[2]).trim(),
          camaraNombre: String(r[3]).trim(),
          cuadrante: String(r[4]).trim(),
          categoria: String(r[5]).trim(),
          prioridad: String(r[6]).trim().toUpperCase() || "ALTA",
          descripcion: String(r[7]).trim(),
          operadorCctv: String(r[8]).trim(),
          imagenes: imagenes,
          estado: String(r[10]).trim().toUpperCase() || "ENVIADA_A_DESPACHADOR",
          movilSugerido: String(r[11]).trim(),
          folioVinculado: String(r[12]).trim(),
          atendidoPor: String(r[13]).trim()
        });
      }
    });
  }

  const shCam = ss.getSheetByName(SHEETS.CENCO_CAMARAS);
  const cencoCameras = [];
  if (shCam && shCam.getLastRow() > 1) {
    const rows = shCam.getRange(2, 1, shCam.getLastRow() - 1, 8).getValues();
    rows.forEach(r => {
      if (String(r[0]).trim()) {
        cencoCameras.push({
          id: String(r[0]).trim(),
          nombre: String(r[1]).trim(),
          cuadrante: String(r[2]).trim(),
          sector: String(r[3]).trim(),
          tipo: String(r[4]).trim(),
          estado: String(r[5]).trim().toUpperCase() || "OPERATIVA",
          gps: String(r[6]).trim(),
          observaciones: String(r[7]).trim()
        });
      }
    });
  }

  const shDesp = ss.getSheetByName(SHEETS.CENCO_DESPACHOS);
  const cencoDispatches = [];
  if (shDesp && shDesp.getLastRow() > 1) {
    const count = Math.min(120, shDesp.getLastRow() - 1);
    const startRow = shDesp.getLastRow() - count + 1;
    const rows = shDesp.getRange(startRow, 1, count, 10).getValues();
    rows.reverse().forEach(r => {
      if (String(r[0]).trim()) {
        cencoDispatches.push({
          despachoId: String(r[0]).trim(),
          fecha: String(r[1]).trim(),
          folioId: String(r[2]).trim(),
          movil: String(r[3]).trim(),
          cuadrante: String(r[4]).trim(),
          despachador: String(r[5]).trim(),
          estadoUnidad: String(r[6]).trim().toUpperCase() || "DESPACHADO",
          horaArribo: String(r[7]).trim(),
          horaLiberacion: String(r[8]).trim(),
          observacion: String(r[9]).trim()
        });
      }
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
    recentProcedures: recentProcedures,
    fleetAlerts: fleetAlerts,
    cencoIncidents: cencoIncidents,
    cencoShifts: cencoShifts,
    cencoCctvAlerts: cencoCctvAlerts,
    cencoCameras: cencoCameras,
    cencoDispatches: cencoDispatches
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
      const kmProxima = (Math.floor((kmMant + 500) / intervalo) + 1) * intervalo;
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

    if (action === "ADMIN_RESOLVE_ALERT") {
      const shAlertas = ss.getSheetByName(SHEETS.ALERTAS_FLOTA);
      const alertId = String(payload.id || "").trim();
      const fechaRes = String(payload.fechaResolucion || nowFormatted).trim();
      const resueltoPor = String(payload.resueltoPor || "Administrador").trim();
      const nota = String(payload.notaResolucion || "Atendido / Resuelto").trim();

      if (shAlertas && alertId && shAlertas.getLastRow() > 1) {
        const ids = shAlertas.getRange(2, 1, shAlertas.getLastRow() - 1, 1).getValues();
        for (let i = ids.length - 1; i >= 0; i--) {
          if (String(ids[i][0]).trim() === alertId) {
            shAlertas.getRange(i + 2, 13, 1, 4).setValues([["RESUELTO", fechaRes, resueltoPor, nota]]);
            return { ok: true, action: action, id: alertId };
          }
        }
      }
      return { ok: true, action: action, id: alertId, note: "Resuelto localmente" };
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

      // Guardar alertas clasificadas del checklist en Alertas_Flota
      const shAlertas = ss.getSheetByName(SHEETS.ALERTAS_FLOTA);
      if (shAlertas && Array.isArray(payload.alertasChecklist) && payload.alertasChecklist.length > 0) {
        payload.alertasChecklist.forEach(al => {
          shAlertas.appendRow([
            al.id || `ALR-${Date.now()}-${Math.floor(Math.random() * 1000)}`,
            payload.fechaInicio || nowFormatted,
            payload.movil || "",
            payload.cuadrante || "",
            payload.rut || "",
            al.categoria || "OTROS",
            (al.prioridad || "MEDIA").toUpperCase(),
            al.itemLabel || al.itemId || "",
            al.estadoReportado || "",
            al.accionRecomendada || "Revisar / Gestionar",
            payload.observaciones || "",
            urlFotoDesperfecto !== "Sin desperfectos" ? urlFotoDesperfecto : urlFotoGeneral,
            "PENDIENTE",
            "",
            "",
            ""
          ]);
        });
      }

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

      // --- INTEGRACIÓN AUTOMÁTICA INSPECTORES -> BITÁCORA CENCO CAD ---
      const shCencoBit = ss.getSheetByName(SHEETS.CENCO_BITACORA);
      let cencoFolioId = "";
      if (shCencoBit) {
        cencoFolioId = "CAD-INSP-" + String(Date.now()).slice(-6);
        const tipoUpper = String(payload.tipo || "PROCEDIMIENTO TERRENO").toUpperCase();
        const prioridadAuto = (tipoUpper.indexOf("INCIVILIDAD") >= 0 || tipoUpper.indexOf("FISCALIZ") >= 0) ? "ALTA" : "MEDIA";
        const historialInit = [{
          fecha: payload.fechaRegistro || nowFormatted,
          rol: "INSPECTOR TERRENO",
          usuario: payload.rut || "Inspector",
          accion: "Procedimiento ingresado directo desde terreno (" + (payload.movil || "Unidad") + ")"
        }];
        shCencoBit.appendRow([
          cencoFolioId,
          payload.fechaRegistro || nowFormatted,
          "INSPECTOR TERRENO",
          "FINALIZADO",
          prioridadAuto,
          tipoUpper,
          payload.subtipo || "Ingreso Directo Terreno",
          payload.cuadrante || "",
          payload.direccion || "",
          "Móvil en terreno: " + (payload.movil || "-"),
          "Inspector Municipal",
          "",
          payload.detalle || "",
          "INSPECTOR: " + (payload.rut || ""),
          "INGRESO DIRECTO",
          payload.movil || "",
          payload.fechaRegistro || nowFormatted,
          payload.fechaRegistro || nowFormatted,
          payload.fechaRegistro || nowFormatted,
          "Registrado y cerrado en terreno por Inspector",
          urlFoto,
          payload.turnoId || "",
          procId,
          JSON.stringify(historialInit)
        ]);
      }

      if (CONFIG.ENVIAR_CORREOS_REALES && (payload.notificarCentral || payload.notificarJefatura)) {
        enviarNotificacionCorreo(payload, urlFoto, nowFormatted);
      }

      return { ok: true, action: action, procId: procId, cencoFolioId: cencoFolioId, urlFoto: urlFoto };
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

      if (sheetTurnos.getLastRow() > 1) {
        const lastRowT = sheetTurnos.getLastRow();
        const tData = sheetTurnos.getRange(2, 1, lastRowT - 1, 14).getValues();
        const targetTid = String(payload.turnoId || "").trim();
        const targetMovil = String(payload.movil || "").trim().toUpperCase();
        for (let i = tData.length - 1; i >= 0; i--) {
          const rowTid = String(tData[i][0] || "").trim();
          const rowMovil = String(tData[i][4] || "").trim().toUpperCase();
          const rowEstado = String(tData[i][13] || "").trim().toUpperCase();
          const matchById = targetTid && rowTid === targetTid;
          const matchOpenMovil = targetMovil && rowMovil === targetMovil && rowEstado === "ACTIVO";
          if (matchById || matchOpenMovil) {
            const rowIndex = i + 2;
            sheetTurnos.getRange(rowIndex, 14, 1, 4).setValues([[
              "FINALIZADO",
              payload.fechaTermino || nowFormatted,
              kmFin,
              kmRecorridos
            ]]);
          }
        }
      }

      actualizarKmFlota(ss, payload.movil, kmFin);

      return { ok: true, action: action, kmRecorridos: kmRecorridos };
    }

    // ========================================================================
    // ACCIONES SISTEMA CAD CENCO 1410 (TELEFONISTA -> DESPACHADOR <-> CCTV)
    // ========================================================================
    if (action === "CENCO_INICIO_TURNO") {
      const shTurnos = ss.getSheetByName(SHEETS.CENCO_TURNOS);
      const turnoCencoId = payload.turnoCencoId || ("CENCO-T-" + Date.now());
      shTurnos.appendRow([
        turnoCencoId,
        payload.fechaInicio || nowFormatted,
        payload.rut || "",
        payload.operador || "",
        (payload.rolCenco || "CENCO_TELEFONISTA").toUpperCase(),
        payload.puesto || "Consola 1",
        (payload.rangoHorario || "MANANA").toUpperCase(),
        payload.novedadesInicio || "Sin novedades al recibir turno",
        "ACTIVO",
        "",
        0,
        0,
        0,
        "",
        ""
      ]);
      return { ok: true, action: action, turnoCencoId: turnoCencoId };
    }

    if (action === "CENCO_CERRAR_TURNO") {
      const shTurnos = ss.getSheetByName(SHEETS.CENCO_TURNOS);
      const turnoCencoId = String(payload.turnoCencoId || "").trim();
      const fechaCierre = payload.fechaCierre || nowFormatted;
      let pdfUrl = "";
      if (payload.pdfBase64) {
        const safeName = ("Acta_Turno_CENCO_" + (turnoCencoId || Date.now()) + ".pdf").replace(/[^a-zA-Z0-9._-]/g, "_");
        pdfUrl = guardarPdfCencoEnDrive(payload.pdfBase64, safeName);
      }

      let updated = false;
      if (shTurnos && turnoCencoId && shTurnos.getLastRow() > 1) {
        const ids = shTurnos.getRange(2, 1, shTurnos.getLastRow() - 1, 1).getValues();
        for (let i = ids.length - 1; i >= 0; i--) {
          if (String(ids[i][0]).trim() === turnoCencoId) {
            shTurnos.getRange(i + 2, 9, 1, 7).setValues([[
              "FINALIZADO",
              fechaCierre,
              Number(payload.totalIngresados || 0),
              Number(payload.totalDespachados || 0),
              Number(payload.totalFinalizados || 0),
              payload.resumenCierre || "",
              pdfUrl || "Generado localmente"
            ]]);
            updated = true;
            break;
          }
        }
      }
      if (!updated && shTurnos) {
        shTurnos.appendRow([
          turnoCencoId || ("CENCO-T-" + Date.now()),
          payload.fechaInicio || fechaCierre,
          payload.rut || "",
          payload.operador || "",
          (payload.rolCenco || "CENCO_SUPERVISOR").toUpperCase(),
          payload.puesto || "Consola CENCO",
          (payload.rangoHorario || "MANANA").toUpperCase(),
          payload.novedadesInicio || "",
          "FINALIZADO",
          fechaCierre,
          Number(payload.totalIngresados || 0),
          Number(payload.totalDespachados || 0),
          Number(payload.totalFinalizados || 0),
          payload.resumenCierre || "",
          pdfUrl || "Generado localmente"
        ]);
      }
      return { ok: true, action: action, turnoCencoId: turnoCencoId, pdfUrl: pdfUrl };
    }

    if (action === "CENCO_CREAR_INCIDENTE") {
      const shBit = ss.getSheetByName(SHEETS.CENCO_BITACORA);
      const folioId = payload.folioId || ("CAD-" + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyyMMdd") + "-" + String(Date.now()).slice(-4));
      const urlsFotos = [];
      if (Array.isArray(payload.imagenesBase64)) {
        payload.imagenesBase64.forEach((b64, idx) => {
          if (b64 && String(b64).indexOf("data:") === 0) {
            urlsFotos.push(guardarFotoEnDrive(b64, `${folioId}_Img${idx + 1}.jpg`));
          } else if (b64 && String(b64).indexOf("http") === 0) {
            urlsFotos.push(String(b64));
          }
        });
      }
      const imgsStr = urlsFotos.length > 0 ? urlsFotos.join(" | ") : "Sin foto";
      const estadoInit = (payload.estado || "PENDIENTE_DESPACHO").toUpperCase();
      const movilesStr = String(payload.movilesAsignados || "").trim();
      const horaDesp = movilesStr ? (payload.horaDespacho || nowFormatted) : "";
      const historial = Array.isArray(payload.historial) && payload.historial.length > 0
        ? payload.historial
        : [{
            fecha: payload.fechaIngreso || nowFormatted,
            rol: payload.origen || "TELEFONISTA",
            usuario: payload.ingresadoPor || "Operador CENCO",
            accion: "Folio CAD ingresado y enviado a cola de despacho"
          }];

      shBit.appendRow([
        folioId,
        payload.fechaIngreso || nowFormatted,
        payload.origen || "LLAMADO 1410",
        estadoInit,
        (payload.prioridad || "MEDIA").toUpperCase(),
        payload.categoria || "GENERAL",
        payload.subcategoria || "",
        payload.cuadrante || "252A",
        payload.direccion || "",
        payload.referenciaCamara || "",
        payload.denunciante || "",
        payload.telefono || "",
        payload.detalle || "",
        payload.ingresadoPor || "",
        payload.despachador || "",
        movilesStr,
        horaDesp,
        "",
        estadoInit === "FINALIZADO" ? nowFormatted : "",
        payload.resolucion || "",
        imgsStr,
        payload.turnoCencoId || "",
        payload.procInspectorId || "",
        JSON.stringify(historial)
      ]);

      // Si ya viene con móvil despachado, registrar también en CENCO_Despachos
      if (movilesStr) {
        const shDesp = ss.getSheetByName(SHEETS.CENCO_DESPACHOS);
        if (shDesp) {
          movilesStr.split(",").map(m => m.trim()).filter(Boolean).forEach(mov => {
            shDesp.appendRow([
              "DSP-" + Date.now() + "-" + Math.floor(Math.random() * 100),
              horaDesp,
              folioId,
              mov,
              payload.cuadrante || "",
              payload.despachador || payload.ingresadoPor || "",
              "DESPACHADO",
              "",
              "",
              payload.detalle || "Despacho directo"
            ]);
          });
        }
      }

      return { ok: true, action: action, folioId: folioId, imagenes: urlsFotos };
    }

    if (action === "CENCO_ACTUALIZAR_INCIDENTE") {
      const shBit = ss.getSheetByName(SHEETS.CENCO_BITACORA);
      const folioId = String(payload.folioId || "").trim();
      if (!folioId || !shBit || shBit.getLastRow() <= 1) {
        return { ok: false, error: "Folio CAD no encontrado" };
      }

      const ids = shBit.getRange(2, 1, shBit.getLastRow() - 1, 1).getValues();
      let rowIndex = -1;
      for (let i = ids.length - 1; i >= 0; i--) {
        if (String(ids[i][0]).trim() === folioId) {
          rowIndex = i + 2;
          break;
        }
      }
      if (rowIndex < 0) return { ok: false, error: "Folio CAD no encontrado: " + folioId };

      const currentRow = shBit.getRange(rowIndex, 1, 1, 24).getValues()[0];
      const newEstado = (payload.estado || currentRow[3] || "DESPACHADO").toUpperCase();
      const newDespachador = payload.despachador !== undefined && payload.despachador !== "" ? payload.despachador : currentRow[14];
      const newMoviles = payload.movilesAsignados !== undefined ? String(payload.movilesAsignados).trim() : String(currentRow[15] || "").trim();
      const newHoraDespacho = payload.horaDespacho || currentRow[16] || (newMoviles ? nowFormatted : "");
      const newHoraArribo = payload.horaArribo || currentRow[17] || ((newEstado === "EN_LUGAR" || newEstado === "EN_PROCEDIMIENTO") ? nowFormatted : "");
      const newHoraCierre = payload.horaCierre || currentRow[18] || ((newEstado === "FINALIZADO" || newEstado === "CANCELADO") ? nowFormatted : "");
      const newResolucion = payload.resolucion !== undefined && payload.resolucion !== "" ? payload.resolucion : currentRow[19];

      // Imágenes existentes + nuevas imágenes agregadas a la bitácora
      const existingImgsStr = String(currentRow[20] || "").trim();
      const allImgs = existingImgsStr && existingImgsStr !== "Sin foto"
        ? existingImgsStr.split(/\s*\|\s*/).filter(Boolean)
        : [];
      const newlySavedUrls = [];
      if (Array.isArray(payload.nuevasImagenesBase64)) {
        payload.nuevasImagenesBase64.forEach((b64, idx) => {
          if (b64 && String(b64).indexOf("data:") === 0) {
            const url = guardarFotoEnDrive(b64, `${folioId}_Bitacora_${Date.now()}_${idx + 1}.jpg`);
            allImgs.push(url);
            newlySavedUrls.push(url);
          }
        });
      }

      let historial = [];
      try { historial = currentRow[23] ? JSON.parse(String(currentRow[23])) : []; } catch (ignore) { historial = []; }
      if (!Array.isArray(historial)) historial = [];
      if (payload.eventoHistorial) {
        historial.push({
          fecha: nowFormatted,
          rol: payload.eventoHistorial.rol || "CENCO",
          usuario: payload.eventoHistorial.usuario || newDespachador || "Operador",
          accion: payload.eventoHistorial.accion || ("Actualización estado: " + newEstado)
        });
      }

      shBit.getRange(rowIndex, 4).setValue(newEstado);
      shBit.getRange(rowIndex, 15, 1, 7).setValues([[
        newDespachador,
        newMoviles,
        newHoraDespacho,
        newHoraArribo,
        newHoraCierre,
        newResolucion,
        allImgs.length > 0 ? allImgs.join(" | ") : "Sin foto"
      ]]);
      shBit.getRange(rowIndex, 24).setValue(JSON.stringify(historial));

      // Registrar despacho si se asignaron móviles en esta acción
      if (payload.registrarDespacho && newMoviles) {
        const shDesp = ss.getSheetByName(SHEETS.CENCO_DESPACHOS);
        if (shDesp) {
          newMoviles.split(",").map(m => m.trim()).filter(Boolean).forEach(mov => {
            shDesp.appendRow([
              "DSP-" + Date.now() + "-" + Math.floor(Math.random() * 100),
              nowFormatted,
              folioId,
              mov,
              currentRow[7] || "",
              newDespachador || "",
              newEstado,
              newHoraArribo,
              newHoraCierre,
              payload.notaDespacho || newResolucion || "Despacho asignado por CENCO"
            ]);
          });
        }
      }

      return { ok: true, action: action, folioId: folioId, estado: newEstado, imagenes: allImgs, nuevasImagenes: newlySavedUrls };
    }

    if (action === "CENCO_CREAR_ALERTA_CCTV") {
      const shCctv = ss.getSheetByName(SHEETS.CENCO_CCTV_ALERTAS);
      const alertaId = payload.alertaId || ("CCTV-" + String(Date.now()).slice(-6));
      const urlsFotos = [];
      if (Array.isArray(payload.imagenesBase64)) {
        payload.imagenesBase64.forEach((b64, idx) => {
          if (b64 && String(b64).indexOf("data:") === 0) {
            urlsFotos.push(guardarFotoEnDrive(b64, `${alertaId}_Cam_${idx + 1}.jpg`));
          }
        });
      }
      const imgsStr = urlsFotos.length > 0 ? urlsFotos.join(" | ") : "Sin foto";

      shCctv.appendRow([
        alertaId,
        payload.fecha || nowFormatted,
        payload.camaraId || "",
        payload.camaraNombre || "",
        payload.cuadrante || "252A",
        payload.categoria || "AVISTAMIENTO CCTV",
        (payload.prioridad || "ALTA").toUpperCase(),
        payload.descripcion || "",
        payload.operadorCctv || "",
        imgsStr,
        "ENVIADA_A_DESPACHADOR",
        payload.movilSugerido || "",
        payload.folioVinculado || "",
        ""
      ]);

      // Si CCTV pidió crear también folio CAD inmediato para el Despachador
      let folioCreado = payload.folioVinculado || "";
      if (payload.crearFolioInmediato) {
        const shBit = ss.getSheetByName(SHEETS.CENCO_BITACORA);
        folioCreado = "CAD-CCTV-" + String(Date.now()).slice(-5);
        const hist = [{
          fecha: payload.fecha || nowFormatted,
          rol: "CCTV",
          usuario: payload.operadorCctv || "Operador CCTV",
          accion: "Alerta visual desde cámara " + (payload.camaraId || "") + " enviada a Despachador"
        }];
        shBit.appendRow([
          folioCreado,
          payload.fecha || nowFormatted,
          "CCTV",
          "PENDIENTE_DESPACHO",
          (payload.prioridad || "ALTA").toUpperCase(),
          payload.categoria || "AVISTAMIENTO CCTV",
          "Cámara " + (payload.camaraId || ""),
          payload.cuadrante || "252A",
          payload.camaraNombre || "",
          (payload.camaraId || "") + " - " + (payload.camaraNombre || ""),
          "Operador CCTV",
          "",
          payload.descripcion || "",
          "CCTV: " + (payload.operadorCctv || ""),
          "",
          "",
          "",
          "",
          "",
          "",
          imgsStr,
          payload.turnoCencoId || "",
          alertaId,
          JSON.stringify(hist)
        ]);
        // Actualizar folio vinculado en última fila de CCTV
        shCctv.getRange(shCctv.getLastRow(), 13).setValue(folioCreado);
      }

      return { ok: true, action: action, alertaId: alertaId, folioVinculado: folioCreado, imagenes: urlsFotos };
    }

    if (action === "CENCO_ATENDER_ALERTA_CCTV") {
      const shCctv = ss.getSheetByName(SHEETS.CENCO_CCTV_ALERTAS);
      const alertaId = String(payload.alertaId || "").trim();
      if (shCctv && alertaId && shCctv.getLastRow() > 1) {
        const ids = shCctv.getRange(2, 1, shCctv.getLastRow() - 1, 1).getValues();
        for (let i = ids.length - 1; i >= 0; i--) {
          if (String(ids[i][0]).trim() === alertaId) {
            shCctv.getRange(i + 2, 11, 1, 4).setValues([[
              (payload.estado || "DESPACHADO").toUpperCase(),
              payload.movilSugerido || "",
              payload.folioVinculado || "",
              payload.atendidoPor || "Despachador"
            ]]);
            return { ok: true, action: action, alertaId: alertaId };
          }
        }
      }
      return { ok: true, action: action, alertaId: alertaId };
    }

    if (action === "CENCO_GUARDAR_CAMARA") {
      const shCam = ss.getSheetByName(SHEETS.CENCO_CAMARAS);
      const id = String(payload.id || ("CAM-" + Date.now())).trim();
      let foundRow = -1;
      if (shCam.getLastRow() > 1) {
        const ids = shCam.getRange(2, 1, shCam.getLastRow() - 1, 1).getValues();
        for (let i = 0; i < ids.length; i++) {
          if (String(ids[i][0]).trim().toUpperCase() === id.toUpperCase()) {
            foundRow = i + 2;
            break;
          }
        }
      }
      const rowVals = [
        id,
        payload.nombre || "",
        (payload.cuadrante || "252A").toUpperCase(),
        payload.sector || "",
        payload.tipo || "PTZ 360°",
        (payload.estado || "OPERATIVA").toUpperCase(),
        payload.gps || "",
        payload.observaciones || ""
      ];
      if (foundRow > 0) {
        shCam.getRange(foundRow, 1, 1, 8).setValues([rowVals]);
      } else {
        shCam.appendRow(rowVals);
      }
      return { ok: true, action: action, id: id };
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

function obtenerCarpetaPdfCenco() {
  const folderName = CONFIG.FOLDER_NAME_PDF_CENCO || "Peña_Seguro_1410_Actas_CENCO_PDF";
  const folders = DriveApp.getFoldersByName(folderName);
  if (folders.hasNext()) {
    return folders.next();
  }
  return DriveApp.createFolder(folderName);
}

function guardarPdfCencoEnDrive(dataUrlOrBase64, fileName) {
  try {
    const folder = obtenerCarpetaPdfCenco();
    const parts = String(dataUrlOrBase64).split(",");
    const base64Data = parts.length > 1 ? parts[1] : parts[0];
    const decoded = Utilities.base64Decode(base64Data);
    const blob = Utilities.newBlob(decoded, "application/pdf", fileName);
    const file = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    return file.getUrl();
  } catch (e) {
    return "Error al guardar PDF: " + e.toString();
  }
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
