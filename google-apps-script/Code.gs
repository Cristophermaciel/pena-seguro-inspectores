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
  ENVIAR_CORREOS_REALES: false,
  WIALON_HOSTS: ["https://hst-api.wialon.us", "https://hst-api.wialon.com", "https://cliente.gps-vision.cl"]
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
  CENCO_DESPACHOS: "CENCO_Despachos",
  // --- Hojas de Inteligencia Operativa y Mapa de Calor (v24) ---
  INTEL_MECHEROS: "Intel_Mecheros",
  INTEL_ROBADOS: "Intel_Vehiculos_Robados",
  INTEL_CALLE: "Intel_Situacion_Calle",
  INTEL_CALOR: "Intel_Puntos_Calor",
  // --- Directorio Oficial de Funcionarios (v34) ---
  STAFF_FUNCIONARIOS: "Staff_Funcionarios",
  // --- Directorio Oficial de Correos Destinatarios (v40) ---
  CORREOS_DESTINATARIOS: "Correos_Destinatarios"
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
        "Total Finalizados", "Resumen Cierre Turno", "URL PDF Acta Turno (Drive)",
        "Conformación Completa (JSON)"
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
    },
    {
      name: SHEETS.INTEL_MECHEROS,
      headers: [
        "ID", "Fecha Registro", "Nombre / Alias", "Cuadrante", "Sector / Comercio Habitual",
        "Vestimenta / Señas", "Nivel Riesgo", "Estado", "Observaciones", "Foto (Drive)", "Registrado Por"
      ]
    },
    {
      name: SHEETS.INTEL_ROBADOS,
      headers: [
        "ID", "Fecha Registro", "Patente", "Marca / Modelo", "Color",
        "Tipo Encargo", "Cuadrante / Lugar", "Estado", "Observaciones", "Foto (Drive)", "Registrado Por"
      ]
    },
    {
      name: SHEETS.INTEL_CALLE,
      headers: [
        "ID", "Fecha Catastro", "Nombre / Apodo", "RUT / Edad", "Cuadrante",
        "Ubicación Habitual", "GPS", "Perfil Social / Salud", "Estado / Derivación", "Observaciones", "Foto (Drive)", "Registrado Por"
      ]
    },
    {
      name: SHEETS.INTEL_CALOR,
      headers: [
        "ID", "Fecha Registro", "Tipo Punto", "Cuadrante", "Dirección",
        "GPS", "Estado", "Descripción", "Foto (Drive)", "Registrado Por"
      ]
    },
    {
      name: SHEETS.STAFF_FUNCIONARIOS,
      headers: [
        "ID", "N° Funcionario", "Nombre Completo", "RUT", "Cargo Actual", "Estado"
      ]
    },
    {
      name: SHEETS.CORREOS_DESTINATARIOS,
      headers: [
        "ID", "Nombre / Área", "Cargo / Institución", "Correo Electrónico", "Estado"
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

  // Sembrar o migrar catálogo oficial de 22 cámaras reales de Peñaflor (KMZ)
  const shCam = ss.getSheetByName(SHEETS.CENCO_CAMARAS);
  if (shCam) {
    const defaultCams = [
      ["CAM-101", "Av. Troncal - Ulkantún", "253A", "Malloco", "PTZ 360°", "OPERATIVA", "-33.604367, -70.864410", ""],
      ["CAM-102", "21 de mayo - María Ester Garay", "252A", "Peñaflor", "PTZ 360°", "OPERATIVA", "-33.608441, -70.899745", ""],
      ["CAM-103", "Bilbao - Vicuña Mackena", "252C", "Peñaflor", "PTZ 360°", "OPERATIVA", "-33.611011, -70.880928", ""],
      ["CAM-104", "AV. Peñaflor - Calle 2", "252C", "Peñaflor", "PTZ 360°", "OPERATIVA", "-33.612676, -70.905891", ""],
      ["CAM-105", "Vicuña Mackena - C. Eduardo Yañez", "253A", "Malloco", "PTZ 360°", "OPERATIVA", "-33.610024, -70.860956", ""],
      ["CAM-106", "Malloquito - Progreso", "252B", "Peñaflor", "PTZ 360°", "CON OBSERVACIONES", "-33.603285, -70.884017", "Con problemas eléctricos no disponible, funciona a veces"],
      ["CAM-107", "Catorce de Julio - 12 de Septiembre", "252B", "Peñaflor", "PTZ 360°", "OPERATIVA", "-33.600836, -70.880500", ""],
      ["CAM-108", "Av. Bernardo O. - Av. La Paz", "253A", "Malloco", "PTZ 360°", "OPERATIVA", "-33.604786, -70.852579", ""],
      ["CAM-109", "Balmaceda - Eliana Del Pin", "253A", "Malloco", "PTZ 360°", "CON OBSERVACIONES", "-33.607362, -70.855136", "Cámara necesita cambio"],
      ["CAM-110", "Av. Peñaflor - Camino Mallarauco - Vicuña Mackena - San Martín", "252C", "Peñaflor", "PTZ 360°", "OPERATIVA", "-33.610697, -70.904997", ""],
      ["CAM-111", "Vicuña Mackena - Concepción - Plaza 22", "253B", "Malloco", "PTZ 360°", "CON OBSERVACIONES", "-33.610482, -70.871418", "Problema eléctrico, cámara necesita cambio de ubicación"],
      ["CAM-112", "Vicuña Mackena - Los Guindos", "252C", "Peñaflor", "PTZ 360°", "OPERATIVA", "-33.610687, -70.875501", ""],
      ["CAM-113", "21 de Mayo - Francisco Bilbao", "252B", "Peñaflor", "PTZ 360°", "OPERATIVA", "-33.605944, -70.878208", ""],
      ["CAM-114", "21 de Mayo - Concepción", "253A", "Malloco", "PTZ 360°", "OPERATIVA", "-33.608604, -70.870417", ""],
      ["CAM-115", "Av. Miraflores - Av. Balmaceda", "253A", "Malloco", "PTZ 360°", "OPERATIVA", "-33.602943, -70.850953", ""],
      ["CAM-116", "Plaza de Peñaflor - Alcalde Luis Araya Cereceda - Dr. Prado Tagle", "252A", "Peñaflor", "PTZ 360°", "OPERATIVA", "-33.608006, -70.900423", ""],
      ["CAM-117", "Vicuña Mackena - 21 de Mayo", "252A", "Peñaflor", "PTZ 360°", "CON OBSERVACIONES", "-33.611106, -70.901979", "Problemas de focalización, cámara necesita reubicación"],
      ["CAM-118", "Basterrica - 21 de Mayo", "252A", "Peñaflor", "PTZ 360°", "OPERATIVA", "-33.609379, -70.900961", ""],
      ["CAM-119", "Torre Interior Municipalidad", "252A", "Peñaflor", "PTZ 360°", "OPERATIVA", "-33.607930, -70.900869", ""],
      ["CAM-120", "Plaza de Malloco - Manuel González", "253A", "Malloco", "PTZ 360°", "OPERATIVA", "-33.604736, -70.852585", ""],
      ["CAM-121", "Ulkantún Fija", "253A", "Malloco", "Fija HD", "OPERATIVA", "-33.604376, -70.864491", ""],
      ["CAM-122", "21 de Mayo Fija", "252A", "Peñaflor", "Fija HD", "OPERATIVA", "-33.608376, -70.899820", ""]
    ];

    let shouldSeedRealCams = shCam.getLastRow() <= 1;
    if (!shouldSeedRealCams && shCam.getLastRow() > 1) {
      const existingIds = shCam.getRange(2, 1, shCam.getLastRow() - 1, 1).getValues().map(r => String(r[0] || "").trim().toUpperCase());
      const hasLegacyDummy = existingIds.indexOf("CAM-01") >= 0 && existingIds.indexOf("CAM-101") === -1;
      if (hasLegacyDummy) {
        shCam.getRange(2, 1, shCam.getLastRow() - 1, 8).clearContent();
        shouldSeedRealCams = true;
      }
    }
    if (shouldSeedRealCams) {
      shCam.getRange(2, 1, defaultCams.length, 8).setValues(defaultCams);
    }
  }

  // Sembrar nómina oficial de 78 funcionarios si la hoja Staff_Funcionarios está vacía
  const shStaff = ss.getSheetByName(SHEETS.STAFF_FUNCIONARIOS);
  if (shStaff && shStaff.getLastRow() <= 1) {
    const defaultStaff = [
      ["STF-001", "725", "ERWIN ORLANDO RODRIGUEZ ERICES", "14244841-6", "INSPECTOR", "ACTIVO"],
      ["STF-002", "726", "PLUTARCO ABURTO CERNA", "09274315-2", "PORTERO", "ACTIVO"],
      ["STF-003", "727", "HECTOR ANDRES ALEGRIA CERECEDA", "16561498-4", "INSPECTOR", "ACTIVO"],
      ["STF-004", "755", "FELIPE GABRIEL GALLARDO CHACON", "17852292-2", "INSPECTOR", "ACTIVO"],
      ["STF-005", "762", "LUIS ALEJANDRO YAÑEZ MAULEN", "16609509-3", "INSPECTOR", "ACTIVO"],
      ["STF-006", "788", "RODRIGO ALEXIS DEL RIO URRA", "11708551-1", "INSPECTOR", "ACTIVO"],
      ["STF-007", "827", "LUIS RICARDO TORO CONTRERAS", "19071479-9", "INSPECTOR", "ACTIVO"],
      ["STF-008", "899", "HECTOR FELIPE GOMEZ ROMERO", "14383154-K", "INSPECTOR", "ACTIVO"],
      ["STF-009", "900", "HECTOR ANTONIO VALENZUELA ZUÑIGA", "14343442-7", "INSPECTOR", "ACTIVO"],
      ["STF-010", "902", "MICHAEL ANDRES COPELLI LECAROS", "18157270-1", "INSPECTOR", "ACTIVO"],
      ["STF-011", "911", "VANESSA VIRGINIA MERINO RAMIREZ", "16644249-4", "INSPECTOR", "ACTIVO"],
      ["STF-012", "913", "MANUEL ESTEBAN AGUSTO CABRERA", "10025056-K", "INSPECTOR", "ACTIVO"],
      ["STF-013", "914", "LUIS ALBERTO GONZALEZ ESPINOSA", "09757919-9", "INSPECTOR", "ACTIVO"],
      ["STF-014", "915", "MANUEL ALEJANDRO CARREÑO RODRIGUEZ", "18623363-8", "INSPECTOR", "ACTIVO"],
      ["STF-015", "916", "JUAN DEL TRANSITO HERNANDEZ CISTERNA", "10564788-3", "INSPECTOR", "ACTIVO"],
      ["STF-016", "919", "JUAN MANUEL URBINA RIVAS", "11332496-1", "INSPECTOR", "ACTIVO"],
      ["STF-017", "922", "PATRICIO EDUARDO URRA TORRES", "11645831-4", "INSPECTOR", "ACTIVO"],
      ["STF-018", "925", "JOSE ARSENIO CASTILLO HIDALGO", "17661284-3", "INSPECTOR", "ACTIVO"],
      ["STF-019", "926", "CARLOS MAURICIO ATENAS NAVARRO", "15649947-1", "INSPECTOR", "ACTIVO"],
      ["STF-020", "927", "ALEJANDRO JESUS MARAMBIO HERNANDEZ", "21726114-7", "INSPECTOR", "ACTIVO"],
      ["STF-021", "929", "PATRICIO RAUL FUENTEALBA BARRIA", "12995904-5", "INSPECTOR", "ACTIVO"],
      ["STF-022", "930", "JERSON ANDRES ARANEDA ZAPATA", "13756174-3", "INSPECTOR", "ACTIVO"],
      ["STF-023", "932", "ADRIAN ANTONIO CIFUENTES MUÑOZ", "14568998-8", "INSPECTOR", "ACTIVO"],
      ["STF-024", "933", "IVAN JONATHAN MICHEL VILLARROEL MADRID", "18959871-8", "INSPECTOR", "ACTIVO"],
      ["STF-025", "944", "LUIS ALEJANDRO BENITEZ BENITEZ", "16352853-3", "INSPECTOR", "ACTIVO"],
      ["STF-026", "958", "CESAR ANTONIO CASTILLO MIRANDA", "17802628-3", "INSPECTOR", "ACTIVO"],
      ["STF-027", "959", "EDUARDO ANDRES VILLAGAS SALAZAR", "16713389-4", "INSPECTOR", "ACTIVO"],
      ["STF-028", "961", "FRANCISCO JAVIER SOLIS GALAZ", "20580191-K", "INSPECTOR", "ACTIVO"],
      ["STF-029", "962", "HECTOR EDUARDO VIDAL ORELLANA", "10062276-9", "INSPECTOR", "ACTIVO"],
      ["STF-030", "963", "IGNACIO ALEJANDRO ARRIAGADA LOBOS", "20136357-8", "INSPECTOR", "ACTIVO"],
      ["STF-031", "964", "JAYSON ELIAS ALISTE URTUBIA", "19063906-1", "INSPECTOR", "ACTIVO"],
      ["STF-032", "965", "JOSE LUIS ROJAS CALDERON", "12223257-3", "INSPECTOR", "ACTIVO"],
      ["STF-033", "966", "LUIS ALBERTO CONTRERAS RICHARD", "10029751-5", "INSPECTOR", "ACTIVO"],
      ["STF-034", "967", "LUIS FELIPE LEIVA PAREJA", "18555957-2", "INSPECTOR", "ACTIVO"],
      ["STF-035", "968", "MAURICIO ANTONIO AYALA TORO", "14500745-3", "INSPECTOR", "ACTIVO"],
      ["STF-036", "969", "MAURICIO ALVARO PAILLAQUEO CARILAO", "11488504-5", "INSPECTOR", "ACTIVO"],
      ["STF-037", "970", "RODRIGO HERNAN GONZALEZ REBOLLEDO", "19650331-5", "INSPECTOR", "ACTIVO"],
      ["STF-038", "971", "VICTOR ALONSO GARCIA DIAZ", "18186814-7", "INSPECTOR", "ACTIVO"],
      ["STF-039", "972", "YENNIFER CAROLINA RIOS LARA", "16872698-8", "INSPECTOR", "ACTIVO"],
      ["STF-040", "973", "CLAUDIO HERNAN SALINAS ESPINOZA", "13557190-3", "INSPECTOR", "ACTIVO"],
      ["STF-041", "742", "YONATHAN SEBASTIAN CATRICHEO ABURTO", "18589636-6", "INSPECTOR", "ACTIVO"],
      ["STF-042", "H", "MANUEL VERA DUARTE", "20.537.044-7", "INSPECTOR CCTV", "ACTIVO"],
      ["STF-043", "352", "PAUL BERENGUELA ARAVENA", "6605297-4", "INSPECTOR CCTV", "ACTIVO"],
      ["STF-044", "670", "CARLA JONES MENA", "13026661-4", "INSPECTOR CCTV", "ACTIVO"],
      ["STF-045", "531", "RAUL SOTO QUIROGA", "18326298-K", "INSPECTOR CCTV", "ACTIVO"],
      ["STF-046", "864", "YESSICA ZUÑIGA ACEVEDO", "9752391-6", "INSPECTOR DESPACHADOR", "ACTIVO"],
      ["STF-047", "862", "FELIPE MILLAN FLORES", "15339144-0", "INSPECTOR DESPACHADOR", "ACTIVO"],
      ["STF-048", "865", "LETICIA CABRERA CARREÑO", "9106120-1", "INSPECTOR CCTV", "ACTIVO"],
      ["STF-049", "H", "CLAUDIA GONZALEZ JEREZ", "11.811.919-3", "INSPECTOR CALLCENTER", "ACTIVO"],
      ["STF-050", "H", "MILVA CARCAMO DIAZ", "13.610.243-5", "INSPECTOR CALLCENTER", "ACTIVO"],
      ["STF-051", "H", "DANIEL MONTECINO EDWARD", "17.933.483-6", "INSPECTOR DESPACHADOR", "ACTIVO"],
      ["STF-052", "H", "BALDRA GALAZ PÉREZ", "12.314.014-1", "INSPECTOR CCTV", "ACTIVO"],
      ["STF-053", "H", "JORGE SOTO BRANTE", "11.656.428-9", "INSPECTOR CCTV", "ACTIVO"],
      ["STF-054", "H", "JOSHUA IBARRA DIAZ", "21.729.845-8", "INSPECTOR CALLCENTER", "ACTIVO"],
      ["STF-055", "H", "YASNA AGUILA ROBLES", "13673508-K", "INSPECTOR DESPACHADOR", "ACTIVO"],
      ["STF-056", "H", "MARCELA MASSONE GABLER", "14.525.223-7", "INSPECTOR DESPACHADOR", "ACTIVO"],
      ["STF-057", "H", "ROSA MADRID GONZALEZ", "14.307.743-8", "INSPECTOR CCTV", "ACTIVO"],
      ["STF-058", "H", "SOFIA CABRERA BERMÚDEZ", "22.266.795-K", "INSPECTOR CCTV", "ACTIVO"],
      ["STF-059", "416", "NICOLAS ANDRES VALENZUELA BERRIOS", "", "INSPECTOR", "ACTIVO"],
      ["STF-060", "533", "DIEGO ANDRES GONZALEZ ATENAS", "", "INSPECTOR", "ACTIVO"],
      ["STF-061", "417", "JORGE DEL CARMEN GUZMAN OYARZUN", "", "INSPECTOR", "ACTIVO"],
      ["STF-062", "363", "LUIS FRANCISCO RIOS PALMA", "", "INSPECTOR", "ACTIVO"],
      ["STF-063", "540", "LUIS FELIPE PUEBLA CARCAMO", "", "INSPECTOR", "ACTIVO"],
      ["STF-064", "662", "ALDO GIORGIO SOTO LLORENTE", "", "INSPECTOR", "ACTIVO"],
      ["STF-065", "270", "HERIBERTO ANTONIO CANCINO VASQUEZ", "", "INSPECTOR", "ACTIVO"],
      ["STF-066", "632", "LUIS MALLEA CONTRERAS", "", "INSPECTOR", "ACTIVO"],
      ["STF-067", "671", "ORLANDO POBLETE VERGARA", "", "INSPECTOR", "ACTIVO"],
      ["STF-068", "431", "VICTOR FABIAN FERNANDEZ ARAVENA", "", "INSPECTOR", "ACTIVO"],
      ["STF-069", "600", "LUIS EDUARDO MORALES OSORIO", "", "INSPECTOR", "ACTIVO"],
      ["STF-070", "379", "JUAN MARCELO HERNANDEZ MELLADO", "", "INSPECTOR", "ACTIVO"],
      ["STF-071", "660", "AARON NATANAEL CACERES MORALES", "", "INSPECTOR", "ACTIVO"],
      ["STF-072", "484", "VICTOR RETAMALES MADARIAGA", "", "PORTERO", "ACTIVO"],
      ["STF-073", "659", "SERGIO FARIAS FLORES", "", "PORTERO", "ACTIVO"],
      ["STF-074", "467", "MARIO MUÑOZ ESCOBAR", "", "PORTERO", "ACTIVO"],
      ["STF-075", "628", "FERNANDO AMIGO ALLENDE", "", "PORTERO", "ACTIVO"],
      ["STF-076", "204", "LUIS OMAR FUENTES DIAZ", "", "PORTERO", "ACTIVO"],
      ["STF-077", "464", "JOSE VIELMA DIAZ", "", "PORTERO", "ACTIVO"],
      ["STF-078", "510", "JUAN IGNACIO CESPEDES AZOCAR", "", "PORTERO", "ACTIVO"]
    ];
    shStaff.getRange(2, 1, defaultStaff.length, 6).setValues(defaultStaff);
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

    if (action === "gpsLive") {
      return respuestaJSON(apiGetGpsLive());
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
 * Optimización para alta concurrencia (hasta 100 usuarios simultáneos):
 * Evita re-verificar las 19 pestañas y Drive en cada lectura/escritura.
 */
function ensureSheetsInitializedFast() {
  try {
    const cache = CacheService.getScriptCache();
    if (cache.get("SHEETS_INIT_OK_V46") === "1") return;
    inicializarHojas();
    cache.put("SHEETS_INIT_OK_V46", "1", 21600); // 6 horas
  } catch (e) {
    inicializarHojas();
  }
}

function getCachedBootstrapData() {
  try {
    const cache = CacheService.getScriptCache();
    const metaRaw = cache.get("BOOT_CACHE_META_V46");
    if (!metaRaw) return null;
    const meta = JSON.parse(metaRaw);
    if (!meta || !meta.chunks) return null;
    const keys = [];
    for (let i = 0; i < meta.chunks; i++) keys.push("BOOT_CACHE_C_" + i);
    const map = cache.getAll(keys);
    let fullStr = "";
    for (let i = 0; i < meta.chunks; i++) {
      const part = map["BOOT_CACHE_C_" + i];
      if (part === undefined || part === null) return null;
      fullStr += part;
    }
    return JSON.parse(fullStr);
  } catch (e) {
    return null;
  }
}

function setCachedBootstrapData(obj) {
  try {
    const cache = CacheService.getScriptCache();
    const fullStr = JSON.stringify(obj);
    const CHUNK_SIZE = 90000;
    const count = Math.ceil(fullStr.length / CHUNK_SIZE);
    if (count > 12) return; // Seguridad
    const batch = { "BOOT_CACHE_META_V46": JSON.stringify({ chunks: count, ts: Date.now() }) };
    for (let i = 0; i < count; i++) {
      batch["BOOT_CACHE_C_" + i] = fullStr.slice(i * CHUNK_SIZE, (i + 1) * CHUNK_SIZE);
    }
    cache.putAll(batch, 20); // 20s TTL compartido entre los 100 usuarios
  } catch (e) {}
}

function invalidateBootstrapCache() {
  try {
    CacheService.getScriptCache().remove("BOOT_CACHE_META_V46");
  } catch (e) {}
}

/**
 * Carga todos los datos maestros desde el Excel en Drive (con caché RAM compartida 20s para 100 usuarios).
 */
function apiGetBootstrapData(forceRefresh) {
  if (!forceRefresh) {
    const cached = getCachedBootstrapData();
    if (cached && cached.ok) return cached;
  }
  ensureSheetsInitializedFast();
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
  const nowMs = Date.now();
  const MAX_SHIFT_AGE_MS = 14 * 60 * 60 * 1000; // 14 horas máximo de vigencia de turno activo
  function parseFechaTurnoMs(val) {
    if (!val) return NaN;
    if (val instanceof Date && !isNaN(val.getTime())) return val.getTime();
    const str = String(val).trim();
    const m = str.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{4})(?:[,\sT]+(\d{1,2}):(\d{2})(?::(\d{2}))?)?/);
    if (m) {
      const d = new Date(Number(m[3]), Number(m[2]) - 1, Number(m[1]), Number(m[4] || 0), Number(m[5] || 0), Number(m[6] || 0));
      if (!isNaN(d.getTime())) return d.getTime();
    }
    const parsed = Date.parse(str);
    return isNaN(parsed) ? NaN : parsed;
  }
  if (shTurnos && shTurnos.getLastRow() > 1) {
    const count = Math.min(60, shTurnos.getLastRow() - 1);
    const startRow = shTurnos.getLastRow() - count + 1;
    const rows = shTurnos.getRange(startRow, 1, count, 17).getValues();
    rows.reverse().forEach(r => {
      const tid = String(r[0] || "").trim();
      const rawEstado = String(r[13] || "").trim().toUpperCase();
      const fechaFin = String(r[14] || "").trim();
      const kmFinVal = Number(r[15] || 0);
      const startMs = parseFechaTurnoMs(r[1]);
      const isExpired = !isNaN(startMs) && (nowMs - startMs > MAX_SHIFT_AGE_MS);
      const isClosed = rawEstado === "FINALIZADO" || rawEstado === "CERRADO" || Boolean(fechaFin) || kmFinVal > 0 || Boolean(tid && closedTurnoIds[tid]) || isExpired;
      const estadoChecklistStr = String(r[7] || "").trim();
      const detailChecklistStr = String(r[12] || "").trim();
      recentShifts.push({
        turnoId: tid,
        fechaInicio: String(r[1]),
        fechaInicioIso: !isNaN(startMs) ? new Date(startMs).toISOString() : "",
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
    const numColsT = Math.min(16, Math.max(15, shCencoTurnos.getLastColumn()));
    const rows = shCencoTurnos.getRange(startRow, 1, count, numColsT).getValues();
    rows.reverse().forEach(r => {
      if (String(r[0]).trim()) {
        let conformacionObj = null;
        try { conformacionObj = r[15] ? JSON.parse(String(r[15])) : null; } catch (ignore) { conformacionObj = null; }
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
          pdfUrl: String(r[14]).trim(),
          conformacion: conformacionObj
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

  // 8. Datos de Inteligencia Operativa y Mapa de Calor (v24)
  const shMech = ss.getSheetByName(SHEETS.INTEL_MECHEROS);
  const intelMecheros = [];
  if (shMech && shMech.getLastRow() > 1) {
    const rows = shMech.getRange(2, 1, shMech.getLastRow() - 1, 11).getValues();
    rows.reverse().forEach(r => {
      if (String(r[0]).trim()) {
        intelMecheros.push({
          id: String(r[0]).trim(),
          fecha: String(r[1]).trim(),
          alias: String(r[2]).trim(),
          cuadrante: String(r[3]).trim(),
          sector: String(r[4]).trim(),
          senas: String(r[5]).trim(),
          riesgo: String(r[6]).trim().toUpperCase() || "ALTO",
          estado: String(r[7]).trim().toUpperCase() || "ACTIVO",
          observaciones: String(r[8]).trim(),
          foto: String(r[9]).trim(),
          registradoPor: String(r[10]).trim()
        });
      }
    });
  }

  const shRob = ss.getSheetByName(SHEETS.INTEL_ROBADOS);
  const intelRobados = [];
  if (shRob && shRob.getLastRow() > 1) {
    const rows = shRob.getRange(2, 1, shRob.getLastRow() - 1, 11).getValues();
    rows.reverse().forEach(r => {
      if (String(r[0]).trim()) {
        intelRobados.push({
          id: String(r[0]).trim(),
          fecha: String(r[1]).trim(),
          patente: String(r[2]).trim().toUpperCase().replace(/[^A-Z0-9]/g, ""),
          marcaModelo: String(r[3]).trim(),
          color: String(r[4]).trim(),
          tipoEncargo: String(r[5]).trim() || "ROBO VEHÍCULO",
          cuadrante: String(r[6]).trim(),
          estado: String(r[7]).trim().toUpperCase() || "ENCARGO VIGENTE",
          observaciones: String(r[8]).trim(),
          foto: String(r[9]).trim(),
          registradoPor: String(r[10]).trim()
        });
      }
    });
  }

  const shCalle = ss.getSheetByName(SHEETS.INTEL_CALLE);
  const intelCalle = [];
  if (shCalle && shCalle.getLastRow() > 1) {
    const rows = shCalle.getRange(2, 1, shCalle.getLastRow() - 1, 12).getValues();
    rows.reverse().forEach(r => {
      if (String(r[0]).trim()) {
        intelCalle.push({
          id: String(r[0]).trim(),
          fecha: String(r[1]).trim(),
          nombre: String(r[2]).trim(),
          rutEdad: String(r[3]).trim(),
          cuadrante: String(r[4]).trim(),
          ubicacion: String(r[5]).trim(),
          gps: String(r[6]).trim(),
          perfil: String(r[7]).trim(),
          estado: String(r[8]).trim() || "EN SEGUIMIENTO",
          observaciones: String(r[9]).trim(),
          foto: String(r[10]).trim(),
          registradoPor: String(r[11]).trim()
        });
      }
    });
  }

  const shCalor = ss.getSheetByName(SHEETS.INTEL_CALOR);
  const intelPuntosCalor = [];
  if (shCalor && shCalor.getLastRow() > 1) {
    const rows = shCalor.getRange(2, 1, shCalor.getLastRow() - 1, 10).getValues();
    rows.reverse().forEach(r => {
      if (String(r[0]).trim()) {
        intelPuntosCalor.push({
          id: String(r[0]).trim(),
          fecha: String(r[1]).trim(),
          tipo: String(r[2]).trim().toUpperCase() || "INCIVILIDAD",
          cuadrante: String(r[3]).trim(),
          direccion: String(r[4]).trim(),
          gps: String(r[5]).trim(),
          estado: String(r[6]).trim().toUpperCase() || "ACTIVO",
          descripcion: String(r[7]).trim(),
          foto: String(r[8]).trim(),
          registradoPor: String(r[9]).trim()
        });
      }
    });
  }

  const shStaffDir = ss.getSheetByName(SHEETS.STAFF_FUNCIONARIOS);
  const staffDirectory = [];
  if (shStaffDir && shStaffDir.getLastRow() > 1) {
    const rows = shStaffDir.getRange(2, 1, shStaffDir.getLastRow() - 1, 6).getValues();
    rows.forEach(r => {
      if (String(r[0]).trim() && String(r[2]).trim()) {
        staffDirectory.push({
          id: String(r[0]).trim(),
          numero: String(r[1]).trim(),
          nombre: String(r[2]).trim(),
          rut: String(r[3]).trim(),
          cargo: String(r[4]).trim().toUpperCase() || "INSPECTOR",
          estado: String(r[5]).trim().toUpperCase() || "ACTIVO"
        });
      }
    });
  }

  const shEmailsDir = ss.getSheetByName(SHEETS.CORREOS_DESTINATARIOS);
  const emailDirectory = [];
  if (shEmailsDir && shEmailsDir.getLastRow() > 1) {
    const rows = shEmailsDir.getRange(2, 1, shEmailsDir.getLastRow() - 1, 5).getValues();
    rows.forEach(r => {
      if (String(r[0]).trim() && String(r[3]).trim()) {
        emailDirectory.push({
          id: String(r[0]).trim(),
          nombre: String(r[1]).trim(),
          cargo: String(r[2]).trim(),
          email: String(r[3]).trim().toLowerCase(),
          estado: String(r[4]).trim().toUpperCase() || "ACTIVO"
        });
      }
    });
  }

  const bootstrapPayload = {
    ok: true,
    date: todayStr,
    spreadsheetUrl: "https://docs.google.com/spreadsheets/d/" + CONFIG.SPREADSHEET_ID + "/edit",
    gpsConfigured: Boolean(PropertiesService.getScriptProperties().getProperty("WIALON_TOKEN")),
    telegramConfigured: Boolean(PropertiesService.getScriptProperties().getProperty("TELEGRAM_BOT_TOKEN") && PropertiesService.getScriptProperties().getProperty("TELEGRAM_CHAT_ID")),
    mailBridgeConfigured: Boolean(PropertiesService.getScriptProperties().getProperty("MAIL_BRIDGE_URL")),
    users: users,
    staffDirectory: staffDirectory,
    emailDirectory: emailDirectory,
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
    cencoDispatches: cencoDispatches,
    cencoGeofenceAuths: getCloudGeofenceAuths(),
    intelMecheros: intelMecheros,
    intelRobados: intelRobados,
    intelCalle: intelCalle,
    intelPuntosCalor: intelPuntosCalor
  };
  setCachedBootstrapData(bootstrapPayload);
  return bootstrapPayload;
}

/**
 * Procesa todas las acciones tanto de Inspectores como de Administradores.
 */
function apiHandleAction(payload) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);

  try {
    ensureSheetsInitializedFast();
    invalidateBootstrapCache();
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

    if (action === "ADMIN_SAVE_STAFF") {
      const sheet = ss.getSheetByName(SHEETS.STAFF_FUNCIONARIOS);
      const id = String(payload.id || ("STF-" + Date.now())).trim();
      const nombre = String(payload.nombre || "").trim().toUpperCase();
      if (!nombre) return { ok: false, error: "Nombre completo requerido" };
      const rowValues = [
        id,
        String(payload.numero || "").trim(),
        nombre,
        String(payload.rut || "").trim(),
        String(payload.cargo || "INSPECTOR").trim().toUpperCase(),
        String(payload.estado || "ACTIVO").trim().toUpperCase()
      ];
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
      if (foundRow > 0) {
        sheet.getRange(foundRow, 1, 1, 6).setValues([rowValues]);
      } else {
        sheet.appendRow(rowValues);
      }
      return { ok: true, action: action, id: id };
    }

    if (action === "ADMIN_DELETE_STAFF") {
      const sheet = ss.getSheetByName(SHEETS.STAFF_FUNCIONARIOS);
      const id = String(payload.id || "").trim();
      if (!id) return { ok: false, error: "ID de funcionario requerido" };
      const lastRow = sheet.getLastRow();
      if (lastRow > 1) {
        const ids = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
        for (let i = ids.length - 1; i >= 0; i--) {
          if (String(ids[i][0]).trim() === id) {
            sheet.deleteRow(i + 2);
            return { ok: true, action: action, id: id };
          }
        }
      }
      return { ok: false, error: "Funcionario no encontrado" };
    }

    if (action === "ADMIN_SAVE_EMAIL") {
      const sheet = ss.getSheetByName(SHEETS.CORREOS_DESTINATARIOS);
      const id = String(payload.id || ("MAIL-" + Date.now())).trim();
      const email = String(payload.email || "").trim().toLowerCase();
      if (!email || email.indexOf("@") < 1) return { ok: false, error: "Correo electrónico inválido" };
      const rowValues = [
        id,
        String(payload.nombre || email).trim(),
        String(payload.cargo || "Destinatario Institucional").trim(),
        email,
        String(payload.estado || "ACTIVO").trim().toUpperCase()
      ];
      const lastRow = sheet.getLastRow();
      let foundRow = -1;
      if (lastRow > 1) {
        const ids = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
        for (let i = 0; i < ids.length; i++) {
          if (String(ids[i][0]).trim() === id) { foundRow = i + 2; break; }
        }
      }
      if (foundRow > 0) sheet.getRange(foundRow, 1, 1, 5).setValues([rowValues]);
      else sheet.appendRow(rowValues);
      return { ok: true, action: action, id: id };
    }

    if (action === "ADMIN_DELETE_EMAIL") {
      const sheet = ss.getSheetByName(SHEETS.CORREOS_DESTINATARIOS);
      const id = String(payload.id || "").trim();
      const lastRow = sheet.getLastRow();
      if (lastRow > 1) {
        const ids = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
        for (let i = ids.length - 1; i >= 0; i--) {
          if (String(ids[i][0]).trim() === id) {
            sheet.deleteRow(i + 2);
            return { ok: true, action: action, id: id };
          }
        }
      }
      return { ok: false, error: "Correo no encontrado" };
    }

    if (action === "ADMIN_SAVE_MAIL_BRIDGE") {
      const props = PropertiesService.getScriptProperties();
      const url = String(payload.bridgeUrl || "").trim();
      const secret = String(payload.bridgeSecret || "PENA1410-MAIL-BRIDGE").trim();
      if (url) props.setProperty("MAIL_BRIDGE_URL", url);
      else props.deleteProperty("MAIL_BRIDGE_URL");
      props.setProperty("MAIL_BRIDGE_SECRET", secret);
      return { ok: true, action: action, configured: Boolean(url) };
    }

    if (action === "CENCO_ENVIAR_CORREO_PROCEDIMIENTO") {
      const recipients = Array.isArray(payload.recipients)
        ? payload.recipients.map(e => String(e || "").trim().toLowerCase()).filter(e => e.indexOf("@") > 0)
        : [];
      if (recipients.length === 0) return { ok: false, error: "Debe seleccionar al menos un correo destinatario." };

      const inc = payload.incident || {};
      const folioId = String(inc.folioId || payload.folioId || "SIN-FOLIO").trim();
      const notaExtra = String(payload.notaAdicional || "").trim();
      const enviadoPor = String(payload.enviadoPor || "Central CENCO 1410").trim();
      const esc = (s) => String(s || "-").replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");

      const rawImgs = [].concat(
        Array.isArray(payload.imagenesBase64) ? payload.imagenesBase64 : [],
        Array.isArray(inc.imagenesBase64) ? inc.imagenesBase64 : [],
        Array.isArray(inc.imagenes) ? inc.imagenes : []
      );
      const driveLinks = [];
      const imagesBase64ToSend = [];
      const seenImgKeys = {};

      rawImgs.forEach((entry, idx) => {
        const s = String(entry || "").trim();
        if (!s || s === "Sin foto") return;
        if (s.indexOf("data:image/") === 0) {
          const sig = s.slice(0, 64) + "_" + s.length;
          if (!seenImgKeys[sig] && imagesBase64ToSend.length < 4) {
            seenImgKeys[sig] = true;
            imagesBase64ToSend.push({ name: folioId + "_Evidencia_" + (imagesBase64ToSend.length + 1) + ".jpg", dataUrl: s });
          }
        } else if (s.indexOf("http") === 0) {
          if (driveLinks.indexOf(s) === -1) driveLinks.push(s);
          // Intentar leer el archivo real desde Drive para incrustarlo y adjuntarlo en el correo
          if (imagesBase64ToSend.length < 4) {
            try {
              const idMatch = s.match(/\/d\/([a-zA-Z0-9_-]{15,})/) || s.match(/[?&]id=([a-zA-Z0-9_-]{15,})/);
              if (idMatch && idMatch[1] && !seenImgKeys[idMatch[1]]) {
                seenImgKeys[idMatch[1]] = true;
                const fBlob = DriveApp.getFileById(idMatch[1]).getBlob();
                const mime = fBlob.getContentType() || "image/jpeg";
                if (mime.indexOf("image/") === 0 && fBlob.getBytes().length <= 4 * 1024 * 1024) {
                  const dataUrl = "data:" + mime + ";base64," + Utilities.base64Encode(fBlob.getBytes());
                  imagesBase64ToSend.push({ name: fBlob.getName() || (folioId + "_Evidencia_" + (imagesBase64ToSend.length + 1) + ".jpg"), dataUrl: dataUrl });
                }
              }
            } catch (ignoreDriveBlob) {}
          }
        }
      });

      let imgsHtml = "";
      if (imagesBase64ToSend.length > 0 || driveLinks.length > 0) {
        imgsHtml = "<div style=\"margin-top:14px;\"><strong style=\"color:#0a2463;\">Evidencia Fotográfica (" + Math.max(imagesBase64ToSend.length, driveLinks.length) + "):</strong>";
        if (imagesBase64ToSend.length > 0) {
          imgsHtml += "<div style=\"margin-top:8px;\">" +
            imagesBase64ToSend.map((im, i) => "<div style=\"margin-bottom:10px;\"><img src=\"cid:evidencia_" + i + "\" alt=\"Evidencia " + (i + 1) + "\" style=\"max-width:100%;max-height:360px;border-radius:8px;border:1px solid #cbd5e1;display:block;\"></div>").join("") +
            "</div>";
        }
        if (driveLinks.length > 0) {
          imgsHtml += "<p style=\"font-size:12px;margin-top:6px;\">" + driveLinks.map((u, idx) => "<a href=\"" + u + "\" target=\"_blank\">Abrir Fotografía #" + (idx + 1) + " en Google Drive</a>").join(" &nbsp;|&nbsp; ") + "</p>";
        }
        imgsHtml += "</div>";
      }
      const hist = Array.isArray(inc.historial) ? inc.historial : [];
      const histHtml = hist.length > 0
        ? "<h3 style=\"color:#0a2463;margin:16px 0 6px;font-size:13px;\">Bitácora / Historial del Procedimiento</h3><ul style=\"padding-left:18px;font-size:12px;color:#334155;\">" +
          hist.map(h => "<li style=\"margin-bottom:4px;\"><strong>[" + esc(h.fecha) + "] " + esc(h.usuario) + ":</strong> " + esc(h.accion) + "</li>").join("") +
          "</ul>"
        : "";

      const subject = "[Peña Seguro 1410] Copia Procedimiento " + folioId + " — " + (inc.categoria || "General") + " (" + (inc.cuadrante || "Comuna") + ")";
      const htmlBody =
        "<div style=\"font-family:Arial,sans-serif;max-width:680px;margin:0 auto;border:1px solid #cbd5e1;border-radius:10px;overflow:hidden;\">" +
        "<div style=\"background:#0a2463;color:#ffffff;padding:16px 20px;\">" +
        "<div style=\"font-size:11px;letter-spacing:1px;text-transform:uppercase;color:#93c5fd;\">Dirección de Seguridad Pública • I. Municipalidad de Peñaflor</div>" +
        "<div style=\"font-size:18px;font-weight:bold;margin-top:4px;\">COPIA OFICIAL DE PROCEDIMIENTO — " + esc(folioId) + "</div>" +
        "</div>" +
        "<div style=\"padding:18px 20px;color:#0f172a;font-size:13px;line-height:1.5;\">" +
        (notaExtra ? "<div style=\"background:#eff6ff;border-left:4px solid #2563eb;padding:10px 12px;margin-bottom:14px;border-radius:4px;\"><strong>Nota de Central 1410:</strong> " + esc(notaExtra) + "</div>" : "") +
        "<table style=\"width:100%;border-collapse:collapse;font-size:12px;margin-bottom:14px;\">" +
        "<tr><td style=\"padding:6px;border-bottom:1px solid #e2e8f0;color:#475569;width:34%;\"><strong>Fecha / Hora Ingreso:</strong></td><td style=\"padding:6px;border-bottom:1px solid #e2e8f0;\">" + esc(inc.fechaIngreso) + "</td></tr>" +
        "<tr><td style=\"padding:6px;border-bottom:1px solid #e2e8f0;color:#475569;\"><strong>Estado / Prioridad:</strong></td><td style=\"padding:6px;border-bottom:1px solid #e2e8f0;\"><strong>" + esc(inc.estado) + "</strong> • Prioridad: " + esc(inc.prioridad) + "</td></tr>" +
        "<tr><td style=\"padding:6px;border-bottom:1px solid #e2e8f0;color:#475569;\"><strong>Categoría / Tipo:</strong></td><td style=\"padding:6px;border-bottom:1px solid #e2e8f0;\">" + esc(inc.categoria) + (inc.subcategoria ? " — " + esc(inc.subcategoria) : "") + "</td></tr>" +
        "<tr><td style=\"padding:6px;border-bottom:1px solid #e2e8f0;color:#475569;\"><strong>Cuadrante / Dirección:</strong></td><td style=\"padding:6px;border-bottom:1px solid #e2e8f0;\">[" + esc(inc.cuadrante) + "] " + esc(inc.direccion) + "</td></tr>" +
        "<tr><td style=\"padding:6px;border-bottom:1px solid #e2e8f0;color:#475569;\"><strong>Origen / Canal:</strong></td><td style=\"padding:6px;border-bottom:1px solid #e2e8f0;\">" + esc(inc.origen) + "</td></tr>" +
        "<tr><td style=\"padding:6px;border-bottom:1px solid #e2e8f0;color:#475569;\"><strong>Móvil(es) Asignado(s):</strong></td><td style=\"padding:6px;border-bottom:1px solid #e2e8f0;\">" + esc(inc.movilesAsignados || "Sin móvil asignado") + "</td></tr>" +
        "<tr><td style=\"padding:6px;border-bottom:1px solid #e2e8f0;color:#475569;\"><strong>Despachador / Ingresado por:</strong></td><td style=\"padding:6px;border-bottom:1px solid #e2e8f0;\">" + esc(inc.despachador || inc.ingresadoPor) + "</td></tr>" +
        "</table>" +
        "<div style=\"background:#f8fafc;border:1px solid #e2e8f0;border-radius:6px;padding:12px;margin-bottom:12px;\"><strong>Relato / Detalle:</strong><br>" + esc(inc.detalle) + "</div>" +
        (inc.resolucion ? "<div style=\"background:#ecfdf5;border:1px solid #a7f3d0;border-radius:6px;padding:12px;margin-bottom:12px;\"><strong>Resolución / Resultado:</strong><br>" + esc(inc.resolucion) + "</div>" : "") +
        imgsHtml + histHtml +
        "<hr style=\"border:none;border-top:1px solid #e2e8f0;margin:16px 0;\">" +
        "<div style=\"font-size:11px;color:#64748b;\">Enviado por <strong>" + esc(enviadoPor) + "</strong> el " + esc(nowFormatted) + " desde Central 1410 Peña Seguro (central1410@penaflor.cl).</div>" +
        "</div></div>";

      const textBody = "PEÑA SEGURO 1410 — COPIA PROCEDIMIENTO " + folioId + "\nFecha: " + (inc.fechaIngreso || "-") +
        "\nCategoría: " + (inc.categoria || "-") + "\nCuadrante/Dirección: [" + (inc.cuadrante || "-") + "] " + (inc.direccion || "-") +
        "\nMóvil: " + (inc.movilesAsignados || "-") + "\nDetalle: " + (inc.detalle || "-") +
        (inc.resolucion ? ("\nResolución: " + inc.resolucion) : "") +
        (notaExtra ? ("\nNota Central 1410: " + notaExtra) : "");

      const props = PropertiesService.getScriptProperties();
      const bridgeUrl = String(props.getProperty("MAIL_BRIDGE_URL") || "").trim();
      const bridgeSecret = String(props.getProperty("MAIL_BRIDGE_SECRET") || "PENA1410-MAIL-BRIDGE").trim();
      let sentFrom = "central1410@penaflor.cl";
      let usedBridge = false;

      if (bridgeUrl) {
        const resp = UrlFetchApp.fetch(bridgeUrl, {
          method: "post",
          contentType: "application/json",
          payload: JSON.stringify({
            secret: bridgeSecret,
            to: recipients.join(","),
            subject: subject,
            htmlBody: htmlBody,
            textBody: textBody,
            senderName: "Central 1410 — Peña Seguro",
            replyTo: "central1410@penaflor.cl",
            imagesBase64: imagesBase64ToSend
          }),
          muteHttpExceptions: true
        });
        const json = JSON.parse(resp.getContentText() || "{}");
        if (!json.ok) {
          return { ok: false, error: "Error en puente central1410@penaflor.cl: " + (json.error || resp.getResponseCode()) };
        }
        sentFrom = json.sentFrom || "central1410@penaflor.cl";
        usedBridge = true;
      } else {
        // Respaldo automático si aún no se configura la URL del puente central1410@penaflor.cl
        const fallbackAttachments = [];
        const fallbackInline = {};
        imagesBase64ToSend.forEach((im, idx) => {
          try {
            const parts = im.dataUrl.split(",");
            const mimeMatch = parts[0].match(/data:(.*?);/);
            const mime = mimeMatch ? mimeMatch[1] : "image/jpeg";
            const blob = Utilities.newBlob(Utilities.base64Decode(parts[1]), mime, im.name || ("Evidencia_" + (idx + 1) + ".jpg"));
            fallbackAttachments.push(blob);
            fallbackInline["evidencia_" + idx] = blob;
          } catch (ignoreFbImg) {}
        });
        const mailOpts = {
          to: recipients.join(","),
          subject: subject,
          body: textBody,
          htmlBody: htmlBody,
          name: "Central 1410 — Peña Seguro",
          replyTo: "central1410@penaflor.cl"
        };
        if (fallbackAttachments.length > 0) {
          mailOpts.attachments = fallbackAttachments;
          mailOpts.inlineImages = fallbackInline;
        }
        MailApp.sendEmail(mailOpts);
        sentFrom = Session.getActiveUser().getEmail() || "Respaldo Apps Script (reply-to central1410@penaflor.cl)";
      }

      // Registrar en el historial del folio en CENCO_Bitacora_CAD
      const shBit = ss.getSheetByName(SHEETS.CENCO_BITACORA);
      if (shBit && folioId && shBit.getLastRow() > 1) {
        const ids = shBit.getRange(2, 1, shBit.getLastRow() - 1, 1).getValues();
        for (let i = ids.length - 1; i >= 0; i--) {
          if (String(ids[i][0]).trim() === folioId) {
            let hArr = [];
            try { hArr = JSON.parse(String(shBit.getRange(i + 2, 24).getValue() || "[]")); if (!Array.isArray(hArr)) hArr = []; } catch (ignore) { hArr = []; }
            hArr.push({
              fecha: nowFormatted,
              rol: "CORREO_OFICIAL",
              usuario: enviadoPor,
              accion: "📧 Copia enviada por correo a: " + recipients.join(", ") + (notaExtra ? (" | Nota: " + notaExtra) : "")
            });
            shBit.getRange(i + 2, 24).setValue(JSON.stringify(hArr));
            break;
          }
        }
      }

      return { ok: true, action: action, folioId: folioId, recipients: recipients, sentFrom: sentFrom, usedBridge: usedBridge };
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
          obtenerTurnoCencoActivoId(ss) || payload.turnoId || "",
          procId,
          JSON.stringify(historialInit)
        ]);
      }

      if (CONFIG.ENVIAR_CORREOS_REALES && (payload.notificarCentral || payload.notificarJefatura)) {
        enviarNotificacionCorreo(payload, urlFoto, nowFormatted);
      }

      // Respaldo automático en Telegram (fecha, hora, lugar e imagen)
      enviarRespaldoTelegram({
        titulo: "PROCEDIMIENTO INSPECTOR",
        folio: cencoFolioId || procId,
        fecha: payload.fechaRegistro || nowFormatted,
        tipo: payload.tipo || "Procedimiento Terreno",
        cuadrante: payload.cuadrante || "-",
        lugar: payload.direccion || "-",
        movil: payload.movil || "-",
        responsable: payload.rut || "Inspector Municipal",
        detalle: payload.detalle || "Sin observaciones",
        urlFotoDrive: urlFoto,
        imagenesBase64: payload.fotoBase64 ? [payload.fotoBase64] : []
      });

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

      // --- INTEGRACIÓN AUTOMÁTICA TAREAS DIARIAS -> BITÁCORA CENCO CAD ---
      const shCencoBitTar = ss.getSheetByName(SHEETS.CENCO_BITACORA);
      let cencoFolioId = "";
      if (shCencoBitTar) {
        cencoFolioId = "CAD-TAR-" + String(Date.now()).slice(-6);
        const descTarea = String(payload.taskDesc || "Tarea Diaria de Cuadrante").trim();
        const obsTarea = String(payload.detalle || "Visto bueno otorgado en terreno.").trim();
        const detalleCompleto = "TAREA CUMPLIDA: " + descTarea + (obsTarea ? " | Obs: " + obsTarea : "");
        const historialInit = [{
          fecha: payload.fechaRegistro || nowFormatted,
          rol: "INSPECTOR TERRENO",
          usuario: payload.rut || "Inspector",
          accion: "Tarea diaria cumplida y reportada desde terreno (" + (payload.movil || "Unidad") + ")"
        }];
        shCencoBitTar.appendRow([
          cencoFolioId,
          payload.fechaRegistro || nowFormatted,
          "TAREA DIARIA",
          "FINALIZADO",
          "MEDIA",
          "TAREA DIARIA CUMPLIDA",
          descTarea.slice(0, 120),
          payload.cuadrante || "",
          payload.direccion || "",
          "Móvil en terreno: " + (payload.movil || "-"),
          "Inspector Municipal",
          "",
          detalleCompleto,
          "INSPECTOR: " + (payload.rut || ""),
          "INGRESO DIRECTO",
          payload.movil || "",
          payload.fechaRegistro || nowFormatted,
          payload.fechaRegistro || nowFormatted,
          payload.fechaRegistro || nowFormatted,
          "Visto bueno de tarea diaria registrado en terreno",
          urlFoto,
          obtenerTurnoCencoActivoId(ss) || payload.turnoId || "",
          regId,
          JSON.stringify(historialInit)
        ]);
      }

      // Respaldo automático en Telegram (Tarea Diaria con fecha, hora, lugar e imagen)
      enviarRespaldoTelegram({
        titulo: "TAREA DIARIA CUMPLIDA",
        folio: cencoFolioId || regId,
        fecha: payload.fechaRegistro || nowFormatted,
        tipo: payload.taskDesc || "Tarea Diaria de Cuadrante",
        cuadrante: payload.cuadrante || "-",
        lugar: payload.direccion || "-",
        movil: payload.movil || "-",
        responsable: payload.rut || "Inspector Municipal",
        detalle: payload.detalle || "Visto bueno otorgado en terreno.",
        urlFotoDrive: urlFoto,
        imagenesBase64: payload.fotoBase64 ? [payload.fotoBase64] : []
      });

      return { ok: true, action: action, regId: regId, cencoFolioId: cencoFolioId, urlFoto: urlFoto };
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
    if (action === "CENCO_INICIO_TURNO" || action === "CENCO_GUARDAR_CONFORMACION") {
      const shTurnos = ss.getSheetByName(SHEETS.CENCO_TURNOS);
      const turnoCencoId = payload.turnoCencoId || ("CENCO-T-" + Date.now());
      const confJson = payload.conformacion ? JSON.stringify(payload.conformacion) : "";
      let updatedRow = -1;
      if (shTurnos.getLastRow() > 1) {
        const ids = shTurnos.getRange(2, 1, shTurnos.getLastRow() - 1, 1).getValues();
        for (let i = ids.length - 1; i >= 0; i--) {
          if (String(ids[i][0] || "").trim() === String(turnoCencoId)) {
            updatedRow = i + 2;
            break;
          }
        }
      }
      if (updatedRow > 0) {
        shTurnos.getRange(updatedRow, 7, 1, 2).setValues([[
          (payload.rangoHorario || "MANANA").toUpperCase(),
          payload.novedadesInicio || "Conformación de turno actualizada"
        ]]);
        if (confJson) shTurnos.getRange(updatedRow, 16).setValue(confJson);
      } else {
        shTurnos.appendRow([
          turnoCencoId,
          payload.fechaInicio || nowFormatted,
          payload.rut || "",
          payload.operador || "",
          (payload.rolCenco || "CENCO_DESPACHADOR").toUpperCase(),
          payload.puesto || "Central Peña Seguro — Despacho Radial",
          (payload.rangoHorario || "MANANA").toUpperCase(),
          payload.novedadesInicio || "Conformación de turno CENCO",
          "ACTIVO",
          "",
          0,
          0,
          0,
          "",
          "",
          confJson
        ]);
      }

      // Registrar o actualizar en Turnos_y_Checklist los móviles y motoristas agregados manualmente por el despachador
      if (payload.conformacion) {
        const shInspTurnos = ss.getSheetByName(SHEETS.TURNOS);
        const manualUnits = [].concat(
          Array.isArray(payload.conformacion.moviles) ? payload.conformacion.moviles : [],
          Array.isArray(payload.conformacion.motoristas) ? payload.conformacion.motoristas : []
        );
        if (shInspTurnos && manualUnits.length > 0) {
          const existingIdsMap = {};
          const lastInspRow = shInspTurnos.getLastRow();
          if (lastInspRow > 1) {
            const idsCol = shInspTurnos.getRange(2, 1, lastInspRow - 1, 1).getValues();
            for (let i = 0; i < idsCol.length; i++) {
              const kId = String(idsCol[i][0] || "").trim();
              if (kId) existingIdsMap[kId] = i + 2;
            }
          }
          manualUnits.forEach((u, idx) => {
            if (!u || !u.movil || u.origen === "AUTO_CHECKLIST") return;
            const mTid = String(u.turnoId || ("MAN-" + Date.now() + "-" + idx)).trim();
            const obsStr = "Radio: " + (u.radio || "-") + " | Pila: " + (u.pila || "-") + " | Cám. Corporal: " + (u.camaraCorporal || "NO");
            const rowData = [
              mTid,
              nowFormatted,
              u.nombre || "Inspector",
              (payload.rangoHorario || "MANANA").toUpperCase(),
              u.movil,
              u.cuadrante || "252A",
              Number(u.kmInicio || 0),
              "Conformado Manual CENCO",
              "Ninguno",
              obsStr,
              "Conformación CENCO",
              "Sin desperfectos",
              JSON.stringify({ manualCenco: true, radio: u.radio || "", pila: u.pila || "", camaraCorporal: u.camaraCorporal || "NO" }),
              "ACTIVO",
              "",
              "",
              ""
            ];
            if (existingIdsMap[mTid]) {
              shInspTurnos.getRange(existingIdsMap[mTid], 1, 1, 17).setValues([rowData]);
            } else {
              shInspTurnos.appendRow(rowData);
              existingIdsMap[mTid] = shInspTurnos.getLastRow();
            }
            if (Number(u.kmInicio || 0) > 0) actualizarKmFlota(ss, u.movil, Number(u.kmInicio));
          });
        }

        // Enviar Conformación Oficial del Turno completa como respaldo automático a Telegram
        if (payload.conformacionResumen) {
          enviarRespaldoTelegram({
            titulo: "CONFORMACIÓN DE TURNO " + (payload.rangoHorario || ""),
            folio: turnoCencoId,
            fecha: payload.fechaInicio || nowFormatted,
            tipo: "Inicio y Conformación de Turno CENCO 1410",
            cuadrante: "TODA LA COMUNA",
            lugar: "Central CENCO 1410 Peñaflor",
            movil: manualUnits.length + " unidad(es) en terreno",
            responsable: payload.operador || "Despachador CENCO",
            detalle: payload.conformacionResumen,
            imagenesBase64: []
          });
        }
      }

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

      // 1. Cerrar TODOS los vehículos activos en Turnos_y_Checklist con sus Km de Término y registrar en Cierres_Turno
      const vehiculosCierre = Array.isArray(payload.vehiculosCierre) ? payload.vehiculosCierre : [];
      const cierreByTid = {};
      const cierreByMov = {};
      const normVeh = (v) => String(v || "").replace(/[^A-Z0-9]/gi, "").toUpperCase();
      vehiculosCierre.forEach(v => {
        if (!v) return;
        if (v.turnoId) cierreByTid[String(v.turnoId).trim()] = v;
        const nk = normVeh(v.movil);
        if (nk) cierreByMov[nk] = v;
      });

      const shInspTurnos = ss.getSheetByName(SHEETS.TURNOS);
      const shCierres = ss.getSheetByName(SHEETS.CIERRES);
      const closedMovKeys = {};
      const closedSummaryLines = [];
      let closedVehiclesCount = 0;

      if (shInspTurnos && shInspTurnos.getLastRow() > 1) {
        const lastRowT = shInspTurnos.getLastRow();
        const scanT = Math.min(80, lastRowT - 1);
        const startRowT = lastRowT - scanT + 1;
        const rowsT = shInspTurnos.getRange(startRowT, 1, scanT, 17).getValues();
        for (let i = 0; i < rowsT.length; i++) {
          const r = rowsT[i];
          const tid = String(r[0] || "").trim();
          const est = String(r[13] || "").trim().toUpperCase();
          const fFin = String(r[14] || "").trim();
          if (est === "FINALIZADO" || est === "CERRADO" || fFin) continue;
          const mov = String(r[4] || "").trim();
          const nk = normVeh(mov);
          const match = (tid && cierreByTid[tid]) || (nk && cierreByMov[nk]) || null;
          const kmIni = Number(r[6] || (match ? match.kmInicial : 0) || 0);
          const kmTerm = match && Number(match.kmTermino || 0) > 0 ? Number(match.kmTermino) : kmIni;
          const kmRec = kmTerm > kmIni ? (kmTerm - kmIni) : 0;

          shInspTurnos.getRange(startRowT + i, 14, 1, 4).setValues([["FINALIZADO", fechaCierre, kmTerm || "", kmRec]]);
          if (shCierres) {
            shCierres.appendRow([
              "CIE-CENCO-" + Date.now() + "-" + i,
              fechaCierre,
              tid,
              String(r[2] || (match ? match.rut : "") || ""),
              String(r[3] || payload.rangoHorario || "MANANA"),
              mov,
              String(r[5] || (match ? match.cuadrante : "") || ""),
              kmIni,
              kmTerm,
              kmRec
            ]);
          }
          if (kmTerm > 0 && mov) actualizarKmFlota(ss, mov, kmTerm);
          if (nk) closedMovKeys[nk] = true;
          closedVehiclesCount++;
          closedSummaryLines.push("• " + mov + " (" + (r[2] || "-") + "): Km " + kmIni + " → " + kmTerm + " (+" + kmRec + " km)");
        }
      }

      // Registrar también cualquier vehículo en vehiculosCierre que no figurara aún como ACTIVO en Turnos_y_Checklist
      vehiculosCierre.forEach((v, idx) => {
        const nk = normVeh(v && v.movil);
        if (!nk || closedMovKeys[nk]) return;
        const kmIni = Number(v.kmInicial || 0);
        const kmTerm = Number(v.kmTermino || 0) || kmIni;
        const kmRec = kmTerm > kmIni ? (kmTerm - kmIni) : 0;
        if (shCierres) {
          shCierres.appendRow([
            "CIE-CENCO-M-" + Date.now() + "-" + idx,
            fechaCierre,
            v.turnoId || ("MAN-" + Date.now()),
            v.rut || "",
            v.rangoHorario || payload.rangoHorario || "MANANA",
            v.movil,
            v.cuadrante || "",
            kmIni,
            kmTerm,
            kmRec
          ]);
        }
        if (kmTerm > 0) actualizarKmFlota(ss, v.movil, kmTerm);
        closedMovKeys[nk] = true;
        closedVehiclesCount++;
        closedSummaryLines.push("• " + v.movil + " (" + (v.rut || "-") + "): Km " + kmIni + " → " + kmTerm + " (+" + kmRec + " km)");
      });

      // 2. Finalizar TODOS los procedimientos de Geocerca GPS abiertos en CENCO_Bitacora_CAD
      const shBit = ss.getSheetByName(SHEETS.CENCO_BITACORA);
      let closedGeofencesCount = 0;
      if (shBit && shBit.getLastRow() > 1) {
        const lastRowB = shBit.getLastRow();
        const scanB = Math.min(150, lastRowB - 1);
        const startRowB = lastRowB - scanB + 1;
        const rowsB = shBit.getRange(startRowB, 1, scanB, 24).getValues();
        for (let i = 0; i < rowsB.length; i++) {
          const r = rowsB[i];
          const fId = String(r[0] || "").trim();
          const orig = String(r[2] || "").trim().toUpperCase();
          const est = String(r[3] || "").trim().toUpperCase();
          const cat = String(r[5] || "").trim().toUpperCase();
          const isGeo = fId.indexOf("CAD-GPS-") === 0 || orig === "ALERTA GPS GEOCERCA" || cat.indexOf("GEOCERCA") >= 0;
          if (!isGeo) continue;
          if (est === "FINALIZADO" || est === "CANCELADO" || est === "DESCARTADA") continue;

          const prevRes = String(r[19] || "").trim();
          const newRes = prevRes
            ? (prevRes + " | Finalizado al Cierre de Turno CENCO (" + turnoCencoId + ")")
            : ("Finalizado automáticamente al Cierre de Turno CENCO (" + turnoCencoId + ")");
          shBit.getRange(startRowB + i, 4).setValue("FINALIZADO");
          shBit.getRange(startRowB + i, 19).setValue(fechaCierre);
          shBit.getRange(startRowB + i, 20).setValue(newRes);
          closedGeofencesCount++;
        }
      }

      // 3. Expirar autorizaciones TODO_TURNO en nube para que el turno entrante inicie limpio
      const updatedAuths = getCloudGeofenceAuths().map(a => {
        if (a && !a.revocada && (String(a.modo || "").toUpperCase() === "TODO_TURNO" || (turnoCencoId && a.turnoCencoId === turnoCencoId))) {
          a.revocada = true;
        }
        return a;
      });
      saveCloudGeofenceAuths(updatedAuths);

      // 4. Enviar resumen de Cierre de Turno y kilometrajes finales a Telegram
      const tgLines = [
        "🔒 CIERRE OFICIAL DE TURNO CENCO (" + (payload.rangoHorario || "-") + ")",
        "👤 Cierra: " + (payload.operador || "-") + (payload.recibeTurno ? ("  |  Recibe: " + payload.recibeTurno) : ""),
        "📊 Totales: " + Number(payload.totalIngresados || 0) + " registrados • " + Number(payload.totalFinalizados || 0) + " finalizados • " + closedGeofencesCount + " geocercas cerradas",
        "📝 Novedades: " + (payload.resumenCierre || "Sin novedades")
      ];
      if (closedSummaryLines.length > 0) {
        tgLines.push("🚓 VEHÍCULOS CERRADOS (" + closedVehiclesCount + "):");
        closedSummaryLines.forEach(l => tgLines.push(l));
      }
      enviarRespaldoTelegram({
        titulo: "CIERRE DE TURNO CENCO " + (payload.rangoHorario || ""),
        folio: turnoCencoId || ("CENCO-T-" + Date.now()),
        fecha: fechaCierre,
        tipo: "Cierre Oficial de Turno y Kilometrajes",
        cuadrante: "TODA LA COMUNA",
        lugar: "Central CENCO 1410 Peñaflor",
        movil: closedVehiclesCount + " vehículo(s) cerrado(s)",
        responsable: payload.operador || "Operador CENCO",
        detalle: tgLines.join("\n"),
        urlFotoDrive: pdfUrl || "",
        imagenesBase64: []
      });

      return {
        ok: true,
        action: action,
        turnoCencoId: turnoCencoId,
        pdfUrl: pdfUrl,
        closedVehiclesCount: closedVehiclesCount,
        closedGeofencesCount: closedGeofencesCount
      };
    }

    if (action === "CENCO_CREAR_INCIDENTE") {
      const shBit = ss.getSheetByName(SHEETS.CENCO_BITACORA);
      const folioId = payload.folioId || ("CAD-" + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyyMMdd") + "-" + String(Date.now()).slice(-4));
      const movilesStrPre = String(payload.movilesAsignados || "").trim();
      const origenUp = String(payload.origen || "").trim().toUpperCase();
      const isGpsGeoAlert = origenUp === "ALERTA GPS GEOCERCA" || String(folioId).indexOf("CAD-GPS-") === 0;

      // Bloqueo en servidor para alertas GPS: respetar autorizaciones vigentes y evitar duplicados entre múltiples pantallas CENCO
      if (isGpsGeoAlert && movilesStrPre) {
        const nowCheckMs = Date.now();
        if (hasActiveCloudGeofenceAuth(movilesStrPre, nowCheckMs)) {
          return { ok: true, action: action, folioId: folioId, skipped: true, reason: "MOVIL_AUTORIZADO" };
        }
        const normMovKey = movilesStrPre.replace(/[^A-Z0-9]/gi, "").toUpperCase();
        const gpsCache = CacheService.getScriptCache();
        const lockKey = "GPS_GEO_CD_" + normMovKey;
        if (gpsCache.get(lockKey)) {
          return { ok: true, action: action, folioId: folioId, skipped: true, reason: "COOLDOWN_ACTIVO" };
        }
        gpsCache.put(lockKey, folioId, 2700); // 45 minutos anti-saturación en servidor
      }

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

      // Respaldo automático en Telegram (Procedimiento CAD / Alerta GPS / Alerta LPR con fecha, hora, lugar e imágenes)
      if (estadoInit !== "DESCARTADA") {
        enviarRespaldoTelegram({
          titulo: (payload.origen || "PROCEDIMIENTO CENCO CAD").toUpperCase(),
          folio: folioId,
          fecha: payload.fechaIngreso || nowFormatted,
          tipo: (payload.categoria || "Procedimiento") + (payload.subcategoria ? " — " + payload.subcategoria : ""),
          cuadrante: payload.cuadrante || "-",
          lugar: payload.direccion || "-",
          movil: movilesStr || "Pendiente despacho",
          responsable: payload.ingresadoPor || "Operador CENCO",
          detalle: payload.detalle || "-",
          urlFotoDrive: urlsFotos[0] || "",
          imagenesBase64: Array.isArray(payload.imagenesBase64) ? payload.imagenesBase64 : []
        });
      }

      return { ok: true, action: action, folioId: folioId, imagenes: urlsFotos };
    }

    if (action === "CENCO_AUTORIZAR_CUADRANTE") {
      const shBit = ss.getSheetByName(SHEETS.CENCO_BITACORA);
      const movil = String(payload.movil || "").trim();
      const conductor = String(payload.conductor || "Inspector en Turno").trim();
      const cuadAsig = String(payload.cuadranteAsignado || "-").trim();
      const cuadDet = String(payload.cuadranteDetectado || "-").trim();
      const tipoAut = String(payload.tipoAutorizacion || "").trim();
      const despachador = String(payload.despachador || "Despachador CENCO").trim();
      const jefeTurno = String(payload.jefeTurno || "Jefe de Turno").trim();
      const vigenciaTexto = String(payload.vigenciaTexto || "TODO EL TURNO").trim();
      const obs = String(payload.observaciones || "").trim();
      const turnoId = obtenerTurnoCencoActivoId(ss, payload.turnoCencoId);
      const modo = String(payload.modo || (vigenciaTexto.toUpperCase().indexOf("TODO") >= 0 ? "TODO_TURNO" : "RANGO")).toUpperCase();
      const nowMsAuth = Date.now();
      const startMsAuth = Number(payload.startMs || nowMsAuth);
      const endMsAuth = modo === "TODO_TURNO"
        ? Math.max(Number(payload.endMs || 0), nowMsAuth + 12 * 60 * 60 * 1000)
        : Number(payload.endMs || (nowMsAuth + 60 * 60 * 1000));
      const authId = String(payload.authId || ("AUT-" + nowMsAuth)).trim();

      // Persistir autorización en nube (ScriptProperties) para que TODOS los equipos CENCO la respeten y no generen alertas repetidas
      if (movil) {
        const normKey = movil.replace(/[^A-Z0-9]/gi, "").toUpperCase();
        const existingAuths = getCloudGeofenceAuths().filter(a =>
          !(String(a.movil || "").replace(/[^A-Z0-9]/gi, "").toUpperCase() === normKey && !a.revocada)
        );
        existingAuths.unshift({
          id: authId,
          movil: movil,
          conductor: conductor,
          cuadranteAsignado: cuadAsig,
          cuadranteDetectado: cuadDet,
          tipoAutorizacion: tipoAut,
          despachador: despachador,
          jefeTurno: jefeTurno,
          modo: modo,
          startMs: startMsAuth,
          endMs: endMsAuth,
          vigenciaTexto: vigenciaTexto,
          observaciones: obs,
          turnoCencoId: turnoId,
          createdMs: nowMsAuth,
          revocada: false
        });
        saveCloudGeofenceAuths(existingAuths);
      }

      const detalleLines = [
        "✅ AUTORIZACIÓN / JUSTIFICACIÓN DE SALIDA DE CUADRANTE",
        "🚓 Móvil: " + movil + " (Conductor: " + conductor + ")",
        "📍 Cuadrante Asignado: " + cuadAsig + " → Ubicación Actual: " + cuadDet,
        "📌 Tipo de Autorización: " + tipoAut,
        "🎧 Despachador CENCO: " + despachador,
        "👮 Jefe de Turno: " + jefeTurno,
        "⏱️ Vigencia / Rango Horario: " + vigenciaTexto
      ];
      if (obs) detalleLines.push("📝 Observaciones: " + obs);
      const resumenTexto = detalleLines.join("\n");

      let folioId = String(payload.folioId || "").trim();
      let updatedExisting = false;

      // Buscar la última alerta GPS abierta de este móvil para justificarla y cerrarla como AUTORIZADO
      if (shBit && shBit.getLastRow() > 1 && movil) {
        const lastRow = shBit.getLastRow();
        const scanCount = Math.min(80, lastRow - 1);
        const startRow = lastRow - scanCount + 1;
        const rows = shBit.getRange(startRow, 1, scanCount, 24).getValues();
        const movUp = movil.toUpperCase();
        for (let i = rows.length - 1; i >= 0; i--) {
          const r = rows[i];
          const rFolio = String(r[0] || "").trim();
          const rEstado = String(r[3] || "").trim().toUpperCase();
          const rMovil = String(r[15] || "").trim().toUpperCase();
          const isMatch = (folioId && rFolio === folioId) ||
            (rFolio.indexOf("CAD-GPS-") === 0 && rMovil.indexOf(movUp) >= 0 && rEstado !== "FINALIZADO" && rEstado !== "AUTORIZADO");
          if (isMatch) {
            folioId = rFolio;
            let hist = [];
            try { hist = JSON.parse(String(r[23] || "[]")); if (!Array.isArray(hist)) hist = []; } catch (ignore) { hist = []; }
            hist.push({
              fecha: nowFormatted,
              rol: "CENCO_DESPACHADOR",
              usuario: despachador,
              accion: resumenTexto
            });
            const targetRow = startRow + i;
            shBit.getRange(targetRow, 4).setValue("AUTORIZADO");
            shBit.getRange(targetRow, 15).setValue(despachador);
            shBit.getRange(targetRow, 19).setValue(nowFormatted);
            shBit.getRange(targetRow, 20).setValue(resumenTexto);
            shBit.getRange(targetRow, 24).setValue(JSON.stringify(hist));
            updatedExisting = true;
            break;
          }
        }
      }

      if (!updatedExisting && shBit) {
        folioId = folioId || ("CAD-AUT-" + String(Date.now()).slice(-5));
        const hist = [{
          fecha: nowFormatted,
          rol: "CENCO_DESPACHADOR",
          usuario: despachador,
          accion: resumenTexto
        }];
        shBit.appendRow([
          folioId,
          nowFormatted,
          "AUTORIZACIÓN CUADRANTE CENCO",
          "AUTORIZADO",
          "MEDIA",
          "Autorización Fuera de Cuadrante",
          tipoAut,
          cuadAsig,
          "Cuadrante " + cuadAsig + " → " + cuadDet,
          "",
          "Jefe de Turno: " + jefeTurno,
          "",
          resumenTexto,
          despachador,
          despachador,
          movil,
          nowFormatted,
          nowFormatted,
          nowFormatted,
          resumenTexto,
          "",
          turnoId,
          "",
          JSON.stringify(hist)
        ]);
      }

      // Enviar respuesta / autorización automáticamente a Telegram
      enviarRespaldoTelegram({
        titulo: "AUTORIZACIÓN SALIDA DE CUADRANTE",
        folio: folioId || ("CAD-AUT-" + String(Date.now()).slice(-5)),
        fecha: nowFormatted,
        tipo: tipoAut,
        cuadrante: cuadAsig + " → " + cuadDet,
        lugar: "Central CENCO 1410 Peñaflor",
        movil: movil,
        responsable: despachador + " / Jefe Turno: " + jefeTurno,
        detalle: resumenTexto,
        imagenesBase64: []
      });

      return { ok: true, action: action, folioId: folioId, authId: authId };
    }

    if (action === "CENCO_REVOCAR_AUTORIZACION") {
      const authId = String(payload.authId || "").trim();
      const movil = String(payload.movil || "").trim();
      const normKey = movil.replace(/[^A-Z0-9]/gi, "").toUpperCase();
      const updated = getCloudGeofenceAuths().map(a => {
        if ((authId && a.id === authId) || (normKey && String(a.movil || "").replace(/[^A-Z0-9]/gi, "").toUpperCase() === normKey)) {
          a.revocada = true;
        }
        return a;
      });
      saveCloudGeofenceAuths(updated);
      return { ok: true, action: action, authId: authId };
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
      const prevEstado = String(currentRow[3] || "").trim().toUpperCase();
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

      // Notificación automática a Telegram cuando un procedimiento pasa a FINALIZADO o CANCELADO
      if ((newEstado === "FINALIZADO" || newEstado === "CANCELADO") && prevEstado !== newEstado) {
        const catStr = String(currentRow[5] || "Procedimiento").trim();
        const subStr = String(currentRow[6] || "").trim();
        const detalleCierre = [
          newEstado === "FINALIZADO" ? "✅ PROCEDIMIENTO CERRADO / FINALIZADO" : "🚫 PROCEDIMIENTO CANCELADO",
          "📋 Relato inicial: " + String(currentRow[12] || "-"),
          "🏁 Resolución / Resultado: " + String(newResolucion || (payload.eventoHistorial && payload.eventoHistorial.accion) || "Finalizado sin observaciones adicionales")
        ].join("\n");
        enviarRespaldoTelegram({
          titulo: newEstado === "FINALIZADO" ? "PROCEDIMIENTO FINALIZADO" : "PROCEDIMIENTO CANCELADO",
          folio: folioId,
          fecha: newHoraCierre || nowFormatted,
          tipo: catStr + (subStr ? " — " + subStr : ""),
          cuadrante: String(currentRow[7] || "-"),
          lugar: String(currentRow[8] || "-"),
          movil: newMoviles || "Sin móvil asignado",
          responsable: (payload.eventoHistorial && payload.eventoHistorial.usuario) || newDespachador || String(currentRow[13] || "Operador CENCO"),
          detalle: detalleCierre,
          urlFotoDrive: allImgs[0] || "",
          imagenesBase64: Array.isArray(payload.nuevasImagenesBase64) ? payload.nuevasImagenesBase64 : []
        });
      }

      return { ok: true, action: action, folioId: folioId, estado: newEstado, imagenes: allImgs, nuevasImagenes: newlySavedUrls };
    }

    if (action === "ADMIN_SAVE_WIALON_TOKEN") {
      const cleanToken = String(payload.token || "").trim();
      const props = PropertiesService.getScriptProperties();
      if (!cleanToken) {
        props.deleteProperty("WIALON_TOKEN");
        CacheService.getScriptCache().remove("GPS_WIALON_LIVE_V1");
        return { ok: true, action: action, configured: false, message: "Token GPS eliminado" };
      }
      props.setProperty("WIALON_TOKEN", cleanToken);
      CacheService.getScriptCache().remove("GPS_WIALON_LIVE_V1");
      const testRes = apiGetGpsLive(true);
      return {
        ok: Boolean(testRes.ok),
        action: action,
        configured: Boolean(testRes.ok),
        unitsCount: (testRes.units || []).length,
        error: testRes.error || ""
      };
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

    if (action === "CENCO_ELIMINAR_CAMARA") {
      const shCam = ss.getSheetByName(SHEETS.CENCO_CAMARAS);
      const id = String(payload.id || "").trim().toUpperCase();
      if (id && shCam && shCam.getLastRow() > 1) {
        const ids = shCam.getRange(2, 1, shCam.getLastRow() - 1, 1).getValues();
        for (let i = ids.length - 1; i >= 0; i--) {
          if (String(ids[i][0] || "").trim().toUpperCase() === id) {
            shCam.deleteRow(i + 2);
            break;
          }
        }
      }
      return { ok: true, action: action, id: id };
    }

    // ========================================================================
    // INTELIGENCIA OPERATIVA (MECHEROS, VEHÍCULOS ROBADOS, SITUACIÓN CALLE, MAPA CALOR)
    // ========================================================================
    if (action === "INTEL_SAVE_MECHERO") {
      const sh = ss.getSheetByName(SHEETS.INTEL_MECHEROS);
      const id = String(payload.id || ("MECH-" + String(Date.now()).slice(-6))).trim();
      const fotoUrl = (payload.fotoBase64 && String(payload.fotoBase64).indexOf("data:") === 0)
        ? guardarFotoEnDrive(payload.fotoBase64, id + "_Mechero.jpg")
        : (payload.foto || "Sin foto");
      const rowVals = [
        id, payload.fecha || nowFormatted, payload.alias || "", (payload.cuadrante || "252A").toUpperCase(),
        payload.sector || "", payload.senas || "", (payload.riesgo || "ALTO").toUpperCase(),
        (payload.estado || "ACTIVO").toUpperCase(), payload.observaciones || "", fotoUrl, payload.registradoPor || ""
      ];
      upsertRowById(sh, id, rowVals);
      return { ok: true, action: action, id: id, foto: fotoUrl };
    }

    if (action === "INTEL_DELETE_MECHERO") {
      deleteRowById(ss.getSheetByName(SHEETS.INTEL_MECHEROS), payload.id);
      return { ok: true, action: action, id: payload.id };
    }

    if (action === "INTEL_SAVE_ROBADO") {
      const sh = ss.getSheetByName(SHEETS.INTEL_ROBADOS);
      const id = String(payload.id || ("ROB-" + String(Date.now()).slice(-6))).trim();
      const patenteClean = String(payload.patente || "").trim().toUpperCase().replace(/[^A-Z0-9]/g, "");
      const fotoUrl = (payload.fotoBase64 && String(payload.fotoBase64).indexOf("data:") === 0)
        ? guardarFotoEnDrive(payload.fotoBase64, id + "_" + (patenteClean || "Veh") + ".jpg")
        : (payload.foto || "Sin foto");
      const rowVals = [
        id, payload.fecha || nowFormatted, patenteClean, payload.marcaModelo || "", payload.color || "",
        payload.tipoEncargo || "ROBO VEHÍCULO", (payload.cuadrante || "252A").toUpperCase(),
        (payload.estado || "ENCARGO VIGENTE").toUpperCase(), payload.observaciones || "", fotoUrl, payload.registradoPor || ""
      ];
      upsertRowById(sh, id, rowVals);
      return { ok: true, action: action, id: id, patente: patenteClean, foto: fotoUrl };
    }

    if (action === "INTEL_DELETE_ROBADO") {
      deleteRowById(ss.getSheetByName(SHEETS.INTEL_ROBADOS), payload.id);
      return { ok: true, action: action, id: payload.id };
    }

    if (action === "INTEL_SAVE_CALLE") {
      const sh = ss.getSheetByName(SHEETS.INTEL_CALLE);
      const id = String(payload.id || ("CALLE-" + String(Date.now()).slice(-6))).trim();
      const fotoUrl = (payload.fotoBase64 && String(payload.fotoBase64).indexOf("data:") === 0)
        ? guardarFotoEnDrive(payload.fotoBase64, id + "_Calle.jpg")
        : (payload.foto || "Sin foto");
      const rowVals = [
        id, payload.fecha || nowFormatted, payload.nombre || "", payload.rutEdad || "",
        (payload.cuadrante || "252A").toUpperCase(), payload.ubicacion || "", payload.gps || "",
        payload.perfil || "", payload.estado || "EN SEGUIMIENTO", payload.observaciones || "", fotoUrl, payload.registradoPor || ""
      ];
      upsertRowById(sh, id, rowVals);
      return { ok: true, action: action, id: id, foto: fotoUrl };
    }

    if (action === "INTEL_DELETE_CALLE") {
      deleteRowById(ss.getSheetByName(SHEETS.INTEL_CALLE), payload.id);
      return { ok: true, action: action, id: payload.id };
    }

    if (action === "INTEL_SAVE_PUNTO_CALOR") {
      const sh = ss.getSheetByName(SHEETS.INTEL_CALOR);
      const id = String(payload.id || ("HEAT-" + String(Date.now()).slice(-6))).trim();
      const fotoUrl = (payload.fotoBase64 && String(payload.fotoBase64).indexOf("data:") === 0)
        ? guardarFotoEnDrive(payload.fotoBase64, id + "_PuntoCalor.jpg")
        : (payload.foto || "Sin foto");
      const rowVals = [
        id, payload.fecha || nowFormatted, (payload.tipo || "INCIVILIDAD").toUpperCase(),
        (payload.cuadrante || "252A").toUpperCase(), payload.direccion || "", payload.gps || "",
        (payload.estado || "ACTIVO").toUpperCase(), payload.descripcion || "", fotoUrl, payload.registradoPor || ""
      ];
      upsertRowById(sh, id, rowVals);
      return { ok: true, action: action, id: id, foto: fotoUrl };
    }

    if (action === "INTEL_DELETE_PUNTO_CALOR") {
      deleteRowById(ss.getSheetByName(SHEETS.INTEL_CALOR), payload.id);
      return { ok: true, action: action, id: payload.id };
    }

    // ========================================================================
    // CONFIGURACIÓN Y DETECCIÓN AUTOMÁTICA DE BOT TELEGRAM (RESPALDO REPORTES)
    // ========================================================================
    if (action === "ADMIN_DETECT_TELEGRAM_CHAT") {
      const token = String(payload.token || PropertiesService.getScriptProperties().getProperty("TELEGRAM_BOT_TOKEN") || "").trim();
      if (!token) return { ok: false, error: "Ingrese primero el Token del Bot de @BotFather" };
      const resp = UrlFetchApp.fetch("https://api.telegram.org/bot" + encodeURIComponent(token) + "/getUpdates", { method: "get", muteHttpExceptions: true });
      const data = JSON.parse(resp.getContentText() || "{}");
      if (!data.ok) return { ok: false, error: data.description || "Token Telegram inválido" };
      const chatsMap = {};
      (data.result || []).forEach(u => {
        const msg = u.message || u.channel_post || u.my_chat_member || {};
        const ch = msg.chat;
        if (ch && ch.id) {
          chatsMap[String(ch.id)] = {
            chatId: String(ch.id),
            title: ch.title || [ch.first_name, ch.last_name].filter(Boolean).join(" ") || ch.username || String(ch.id),
            type: ch.type || "chat"
          };
        }
      });
      const chats = Object.keys(chatsMap).map(k => chatsMap[k]);
      return { ok: true, action: action, chats: chats };
    }

    if (action === "ADMIN_SAVE_TELEGRAM_CONFIG") {
      const token = String(payload.token || "").trim();
      const chatId = String(payload.chatId || "").trim();
      const props = PropertiesService.getScriptProperties();
      if (!token || !chatId) {
        return { ok: false, error: "Debe ingresar tanto el Token del Bot como el Chat ID" };
      }
      props.setProperty("TELEGRAM_BOT_TOKEN", token);
      props.setProperty("TELEGRAM_CHAT_ID", chatId);
      const testRes = enviarRespaldoTelegram({
        titulo: "PRUEBA DE CONEXIÓN EXITOSA",
        folio: "TEST-TG-1410",
        fecha: nowFormatted,
        tipo: "Vinculación Bot de Respaldo Operativo",
        cuadrante: "TODA LA COMUNA",
        lugar: "Central CENCO CAD — Peña Seguro 1410",
        movil: "CENTRAL 1410",
        responsable: payload.usuario || "Administrador",
        detalle: "✅ El respaldo automático en Telegram quedó activo. Cada procedimiento, tarea diaria y alerta enviará fecha, hora, lugar e imágenes a este chat.",
        imagenesBase64: []
      }, true);
      return { ok: Boolean(testRes && testRes.ok), action: action, configured: Boolean(testRes && testRes.ok), error: (testRes && testRes.error) || "" };
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

function upsertRowById(sh, id, rowVals) {
  if (!sh) return;
  const target = String(id || "").trim().toUpperCase();
  if (target && sh.getLastRow() > 1) {
    const ids = sh.getRange(2, 1, sh.getLastRow() - 1, 1).getValues();
    for (let i = 0; i < ids.length; i++) {
      if (String(ids[i][0] || "").trim().toUpperCase() === target) {
        sh.getRange(i + 2, 1, 1, rowVals.length).setValues([rowVals]);
        return;
      }
    }
  }
  sh.appendRow(rowVals);
}

function deleteRowById(sh, id) {
  if (!sh || sh.getLastRow() <= 1) return;
  const target = String(id || "").trim().toUpperCase();
  if (!target) return;
  const ids = sh.getRange(2, 1, sh.getLastRow() - 1, 1).getValues();
  for (let i = ids.length - 1; i >= 0; i--) {
    if (String(ids[i][0] || "").trim().toUpperCase() === target) {
      sh.deleteRow(i + 2);
      return;
    }
  }
}

function obtenerTurnoCencoActivoId(ss) {
  try {
    const sh = ss.getSheetByName(SHEETS.CENCO_TURNOS);
    if (!sh || sh.getLastRow() <= 1) return "";
    const count = Math.min(20, sh.getLastRow() - 1);
    const startRow = sh.getLastRow() - count + 1;
    const rows = sh.getRange(startRow, 1, count, 9).getValues();
    for (let i = rows.length - 1; i >= 0; i--) {
      if (String(rows[i][8] || "").trim().toUpperCase() === "ACTIVO") {
        return String(rows[i][0] || "").trim();
      }
    }
  } catch (ignore) {}
  return "";
}

function respuestaJSON(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

/**
 * Consulta en tiempo real la API oficial Wialon Remote API (cliente.gps-vision.cl)
 * usando el token de solo lectura almacenado en ScriptProperties ("WIALON_TOKEN").
 * Incluye caché de 15 segundos para máxima velocidad sin sobrecargar Wialon.
 */
function apiGetGpsLive(bypassCache) {
  try {
    const token = String(PropertiesService.getScriptProperties().getProperty("WIALON_TOKEN") || "").trim();
    if (!token) {
      return { ok: false, configured: false, units: [], error: "Token Wialon GPS no configurado" };
    }

    const cache = CacheService.getScriptCache();
    const cacheKey = "GPS_WIALON_LIVE_V1";
    if (!bypassCache) {
      const cached = cache.get(cacheKey);
      if (cached) {
        try { return JSON.parse(cached); } catch (ignore) {}
      }
    }

    const props = PropertiesService.getScriptProperties();
    const savedHost = String(props.getProperty("WIALON_API_HOST") || "").trim();
    const candidateHosts = [];
    if (savedHost) candidateHosts.push(savedHost);
    (CONFIG.WIALON_HOSTS || ["https://hst-api.wialon.us", "https://hst-api.wialon.com"]).forEach(h => {
      if (candidateHosts.indexOf(h) === -1) candidateHosts.push(h);
    });

    // 1. Iniciar sesión con token oficial Wialon (probando automáticamente el host Wialon Hosting correcto)
    let sid = "";
    let baseUrl = "";
    let lastLoginErr = "No se pudo conectar con Wialon";
    for (let i = 0; i < candidateHosts.length; i++) {
      const h = candidateHosts[i];
      const tryBase = h + "/wialon/ajax.html";
      try {
        const loginUrl = tryBase + "?svc=token/login&params=" + encodeURIComponent(JSON.stringify({ token: token }));
        const loginResp = UrlFetchApp.fetch(loginUrl, { method: "post", muteHttpExceptions: true });
        const rawText = loginResp.getContentText() || "";
        if (rawText.trim().charAt(0) !== "{") {
          lastLoginErr = h + " no respondió JSON";
          continue;
        }
        const loginJson = JSON.parse(rawText);
        if (loginJson && loginJson.eid) {
          sid = loginJson.eid;
          baseUrl = tryBase;
          if (savedHost !== h) props.setProperty("WIALON_API_HOST", h);
          break;
        } else {
          lastLoginErr = "Wialon (" + h + ") error " + (loginJson.error || "?") + ": " + (loginJson.reason || "Token inválido");
        }
      } catch (eHost) {
        lastLoginErr = h + ": " + eHost.toString();
      }
    }

    if (!sid || !baseUrl) {
      return { ok: false, configured: true, units: [], error: lastLoginErr };
    }

    // 2. Consultar todas las unidades (flags 1025 = base 1 + last position 1024)
    const searchParams = {
      spec: {
        itemsType: "avl_unit",
        propName: "sys_name",
        propValueMask: "*",
        sortType: "sys_name"
      },
      force: 1,
      flags: 1025,
      from: 0,
      to: 0
    };
    const searchUrl = baseUrl + "?svc=core/search_items&params=" + encodeURIComponent(JSON.stringify(searchParams)) + "&sid=" + encodeURIComponent(sid);
    const searchResp = UrlFetchApp.fetch(searchUrl, { method: "post", muteHttpExceptions: true });
    const searchJson = JSON.parse(searchResp.getContentText() || "{}");

    // 3. Cerrar sesión Wialon limpiamente para no acumular sesiones
    try {
      UrlFetchApp.fetch(baseUrl + "?svc=core/logout&params={}&sid=" + encodeURIComponent(sid), { method: "post", muteHttpExceptions: true });
    } catch (ignore) {}

    if (searchJson.error) {
      return { ok: false, configured: true, units: [], error: "Wialon search error: " + searchJson.error };
    }

    const nowSec = Math.floor(Date.now() / 1000);
    const items = Array.isArray(searchJson.items) ? searchJson.items : [];
    const units = items.map(u => {
      const pos = u.pos || {};
      const lat = typeof pos.y === "number" ? pos.y : null;
      const lon = typeof pos.x === "number" ? pos.x : null;
      const ts = Number(pos.t || 0);
      const ageSec = ts > 0 ? Math.max(0, nowSec - ts) : 999999;
      return {
        id: u.id,
        nombre: String(u.nm || "").trim(),
        lat: lat,
        lon: lon,
        velocidad: Number(pos.s || 0),
        rumbo: Number(pos.c || 0),
        satelites: Number(pos.sc || 0),
        ultimoReporteEpoch: ts,
        ageSec: ageSec,
        online: ts > 0 && ageSec <= 600
      };
    }).filter(u => u.lat !== null && u.lon !== null);

    const payload = {
      ok: true,
      configured: true,
      fetchedAt: new Date().toISOString(),
      units: units
    };

    try {
      cache.put(cacheKey, JSON.stringify(payload), 15);
    } catch (ignore) {}

    return payload;
  } catch (err) {
    return { ok: false, configured: true, units: [], error: err.toString() };
  }
}

/**
 * Envía respaldo automático a Telegram (Bot API) con fecha, hora, lugar, móvil, detalle e imágenes reales.
 * Usa TELEGRAM_BOT_TOKEN y TELEGRAM_CHAT_ID almacenados en ScriptProperties.
 */
function enviarRespaldoTelegram(datos, isTest) {
  try {
    const props = PropertiesService.getScriptProperties();
    const token = String(props.getProperty("TELEGRAM_BOT_TOKEN") || "").trim();
    const chatId = String(props.getProperty("TELEGRAM_CHAT_ID") || "").trim();
    if (!token || !chatId) return { ok: false, error: "Telegram no configurado" };

    const esc = (s) => String(s || "").replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
    const fullDetalle = String(datos.detalle || "-");
    const lines = [
      "🛡️ <b>PEÑA SEGURO 1410 — " + esc(datos.titulo || "RESPALDO OPERATIVO") + "</b>",
      "🆔 <b>Folio:</b> <code>" + esc(datos.folio || "-") + "</code>",
      "📅 <b>Fecha y Hora:</b> " + esc(datos.fecha || "-"),
      "📌 <b>Categoría / Tipo:</b> " + esc(datos.tipo || "-"),
      "📍 <b>Cuadrante / Lugar:</b> [" + esc(datos.cuadrante || "-") + "] " + esc(datos.lugar || "-"),
      "🚓 <b>Móvil:</b> " + esc(datos.movil || "-") + "  |  👤 <b>Resp.:</b> " + esc(datos.responsable || "-"),
      "📝 <b>Detalle:</b>\n" + esc(fullDetalle)
    ];
    if (datos.urlFotoDrive && datos.urlFotoDrive !== "Sin foto" && String(datos.urlFotoDrive).indexOf("http") === 0) {
      lines.push("🔗 <a href=\"" + datos.urlFotoDrive + "\">Ver evidencia en Google Drive</a>");
    }
    const textMsg = lines.join("\n");
    const apiBase = "https://api.telegram.org/bot" + encodeURIComponent(token);

    // Helper para dividir mensajes largos (>3900 chars) sin cortar líneas ni etiquetas
    const sendTextChunks = (fullText) => {
      const MAX_LEN = 3900;
      const chunks = [];
      let remaining = String(fullText || "");
      while (remaining.length > MAX_LEN) {
        let cut = remaining.lastIndexOf("\n", MAX_LEN);
        if (cut < MAX_LEN * 0.5) cut = MAX_LEN;
        chunks.push(remaining.slice(0, cut));
        remaining = remaining.slice(cut).replace(/^\n+/, "");
      }
      if (remaining) chunks.push(remaining);
      let lastJson = { ok: false };
      chunks.forEach((chunk, idx) => {
        const prefix = chunks.length > 1 ? ("📄 <i>Parte " + (idx + 1) + "/" + chunks.length + "</i>\n") : "";
        const resp = UrlFetchApp.fetch(apiBase + "/sendMessage", {
          method: "post",
          contentType: "application/json",
          payload: JSON.stringify({
            chat_id: chatId,
            text: prefix + chunk,
            parse_mode: "HTML",
            disable_web_page_preview: true
          }),
          muteHttpExceptions: true
        });
        lastJson = JSON.parse(resp.getContentText() || "{}");
      });
      return lastJson;
    };

    // Extraer imágenes base64 válidas
    const validImgs = (Array.isArray(datos.imagenesBase64) ? datos.imagenesBase64 : [])
      .filter(b => b && String(b).indexOf("data:image/") === 0);

    if (validImgs.length > 0) {
      // Enviar primera imagen con caption seguro (máx 1024 chars en Telegram)
      const firstParts = validImgs[0].split(",");
      const firstMime = (firstParts[0].match(/data:(.*?);/) || [null, "image/jpeg"])[1];
      const firstBlob = Utilities.newBlob(Utilities.base64Decode(firstParts[1]), firstMime, (datos.folio || "evidencia") + "_1.jpg");
      const shortCaption = textMsg.length <= 1020
        ? textMsg
        : (lines.slice(0, 6).join("\n") + "\n📝 <i>Ver detalle completo en el mensaje siguiente...</i>").slice(0, 1020);
      const resp1 = UrlFetchApp.fetch(apiBase + "/sendPhoto", {
        method: "post",
        payload: {
          chat_id: chatId,
          caption: shortCaption,
          parse_mode: "HTML",
          photo: firstBlob
        },
        muteHttpExceptions: true
      });
      const json1 = JSON.parse(resp1.getContentText() || "{}");
      if (textMsg.length > 1020) {
        sendTextChunks(textMsg);
      }

      // Si hay más imágenes (2da, 3ra, 4ta), enviarlas inmediatamente debajo
      for (let i = 1; i < Math.min(validImgs.length, 4); i++) {
        try {
          const p = validImgs[i].split(",");
          const m = (p[0].match(/data:(.*?);/) || [null, "image/jpeg"])[1];
          const b = Utilities.newBlob(Utilities.base64Decode(p[1]), m, (datos.folio || "evidencia") + "_" + (i + 1) + ".jpg");
          UrlFetchApp.fetch(apiBase + "/sendPhoto", {
            method: "post",
            payload: {
              chat_id: chatId,
              caption: "📸 Evidencia #" + (i + 1) + " — Folio " + (datos.folio || "-"),
              photo: b
            },
            muteHttpExceptions: true
          });
        } catch (ignoreExtra) {}
      }
      return { ok: Boolean(json1.ok), error: json1.description || "" };
    } else {
      const json = sendTextChunks(textMsg);
      return { ok: Boolean(json.ok), error: json.description || "" };
    }
  } catch (err) {
    if (isTest) return { ok: false, error: err.toString() };
    return { ok: false };
  }
}

/**
 * Persistencia en nube de Autorizaciones de Salida de Cuadrante (sincronizada a todos los puestos CENCO).
 */
function getCloudGeofenceAuths() {
  try {
    const raw = PropertiesService.getScriptProperties().getProperty("CENCO_GEOFENCE_AUTHS_JSON");
    if (!raw) return [];
    const list = JSON.parse(raw);
    if (!Array.isArray(list)) return [];
    const cutoff = Date.now() - 36 * 60 * 60 * 1000;
    return list.filter(a => a && Number(a.createdMs || 0) >= cutoff);
  } catch (e) {
    return [];
  }
}

function saveCloudGeofenceAuths(list) {
  try {
    const cutoff = Date.now() - 36 * 60 * 60 * 1000;
    const pruned = (Array.isArray(list) ? list : [])
      .filter(a => a && Number(a.createdMs || 0) >= cutoff)
      .slice(0, 80);
    PropertiesService.getScriptProperties().setProperty("CENCO_GEOFENCE_AUTHS_JSON", JSON.stringify(pruned));
  } catch (e) {}
}

function hasActiveCloudGeofenceAuth(movil, nowMs) {
  const key = String(movil || "").replace(/[^A-Z0-9]/gi, "").toUpperCase();
  if (!key) return false;
  const now = nowMs || Date.now();
  const list = getCloudGeofenceAuths();
  for (let i = 0; i < list.length; i++) {
    const a = list[i];
    if (!a || a.revocada) continue;
    const aKey = String(a.movil || "").replace(/[^A-Z0-9]/gi, "").toUpperCase();
    if (aKey !== key) continue;
    if (String(a.modo || "").toUpperCase() === "TODO_TURNO") {
      const endT = Number(a.endMs || (Number(a.createdMs || 0) + 12 * 60 * 60 * 1000));
      if (now <= endT) return true;
    } else {
      const sMs = Number(a.startMs || 0);
      const eMs = Number(a.endMs || 0);
      if (sMs && eMs && now >= (sMs - 120000) && now <= eMs) return true;
    }
  }
  return false;
}
