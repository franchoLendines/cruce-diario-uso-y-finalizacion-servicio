/**
 * =================================================================================
 * CRUCE DIARIO DE USO Y FINALIZACIÓN DE UN SERVICIO (3 INFORMES → 1 HOJA)
 * =================================================================================
 * Cada día llegan por correo tres informes Excel distintos relacionados con el
 * mismo servicio:
 *   1. USOS: quién ha usado el servicio ayer (archivo "maestro"/driver).
 *   2. CLIENTE: datos de contacto (email, cuenta) de cada cliente.
 *   3. ASISTENCIA: datos de póliza (ramo, número de póliza) de cada cliente.
 *
 * El script cruza los tres informes por un identificador común ("Mapfre ID"),
 * añade el resultado como filas nuevas en una hoja de Google Sheets histórica,
 * y envía un correo de resumen con una tabla HTML de lo añadido ese día.
 *
 * ACTIVADOR RECOMENDADO (Editor de Apps Script > Activadores):
 *   • procesarCruceDiario → diario, después de la hora a la que suelen llegar
 *     los tres informes del día anterior.
 */
/**
 * CONFIGURACIÓN GLOBAL — RELLENA AQUÍ TUS DATOS
 */
const CONFIG = {
  EMAIL_DESTINO: "tu_correo@tudominio.com",
  EMAIL_CC: "correo_area_interna@tudominio.com",
  SHEET_DESTINO_ID: "TU_SHEET_ID_DESTINO",
  SHEET_TAB_NAME: "Consolidado_Diario",

  FILES: {
    // ARCHIVO MAESTRO (DRIVER): registros de uso del servicio
    USOS: {
      SUBJECT: "TU_ASUNTO_INFORME_USOS",
      PREFIX: "TU_PREFIJO_ADJUNTO_USOS",
      SKIP_ROWS: 14,               // Nº de filas de cabecera a saltar antes de los datos
      COL_FECHA_ORIGINAL: 1,       // Columna con la fecha del uso (índice base 0)
      COL_ID_MAPFRE_ORIGINAL: 6,   // Columna con el identificador común (Mapfre ID)
      COL_CHECK_ORIGINAL: 9,       // Columna con el valor a comprobar (p. ej. estado/servicio)
      COL_SERVICIO_ORIGINAL: 12,   // Columna con el nombre del servicio
      TARGET_VAL: "TU_VALOR_OBJETIVO" // Valor que debe tener COL_CHECK_ORIGINAL para incluir la fila
    },
    // Datos de contacto del cliente
    CLIENTE: {
      SUBJECT: "TU_ASUNTO_INFORME_CLIENTE",
      PREFIX: "TU_PREFIJO_ADJUNTO_CLIENTE",
      SKIP_ROWS: 17,
      COL_ID_MAPFRE_ORIGINAL: 5,
      COL_EMAIL_ORIGINAL: 7,
      COL_CUENTA_ORIGINAL: 8
    },
    // Datos de póliza asociados a cada cliente
    ASISTENCIA: {
      SUBJECT: "TU_ASUNTO_INFORME_ASISTENCIA",
      PREFIX: "TU_PREFIJO_ADJUNTO_ASISTENCIA",
      SKIP_ROWS: 13,
      COL_ID_MAPFRE_ORIGINAL: 6,
      HEADER_RAMO: "Ramo",
      HEADER_POLIZA: "Número de póliza"
    }
  }
};

function procesarCruceDiario() {
  const yesterday = getYesterday();

  const files = getDailyFiles(yesterday.date);

  if (!files.usos || !files.cliente || !files.asistencia) {
    sendErrorEmail("Falta uno o más archivos adjuntos del día de ayer.");
    return;
  }
  // --- PROCESAR USOS ---
  const rawUsos = convertExcelToValues(files.usos);
  const dataUsos = processUsosFile(rawUsos, yesterday.date);

  if (dataUsos.length === 0) {
    Logger.log("ALERTA: Finalizado sin registros. Revisa los logs de arriba.");
    sendReportEmail([], yesterday.str); // Pasamos array vacío
    return;
  }
  const rawCliente = convertExcelToValues(files.cliente);
  const mapCliente = createClienteMap(rawCliente);
  const rawAsistencia = convertExcelToValues(files.asistencia);
  const mapAsistencia = createAsistenciaMap(rawAsistencia);
  const finalData = [];

  dataUsos.forEach(row => {
    const mapfreId = row.joinKey;
    const datosCliente = mapCliente.get(mapfreId) || { email: "No encontrado", cuenta: "No encontrado" };
    const datosAsistencia = mapAsistencia.get(mapfreId) || { ramo: "No encontrado", poliza: "No encontrado" };
    finalData.push([
      yesterday.str,
      mapfreId,
      row.promoCode,
      row.servicio,
      datosCliente.email,
      datosCliente.cuenta,
      datosAsistencia.ramo,
      datosAsistencia.poliza
    ]);
  });
  saveToSheet(finalData);
  sendReportEmail(finalData, yesterday.str); // Pasamos los datos completos
}

/**
 * LÓGICA DE FILTRADO CON LOGS
 */
function processUsosFile(rawMatrix, dateTargetObj) {
  if (!rawMatrix || rawMatrix.length <= CONFIG.FILES.USOS.SKIP_ROWS) return [];
  const config = CONFIG.FILES.USOS;
  const cleanData = [];

  let logsPrinted = 0;
  for (let i = config.SKIP_ROWS; i < rawMatrix.length; i++) {
    const row = rawMatrix[i];

    if (!row || row.length < 10) continue;
    // 1. OBTENER VALORES ORIGINALES
    const rawDate = row[config.COL_FECHA_ORIGINAL];
    const rawCheck = row[config.COL_CHECK_ORIGINAL];
    const rawId = row[config.COL_ID_MAPFRE_ORIGINAL];
    const rawServicio = row[config.COL_SERVICIO_ORIGINAL] || "";
    // 2. NORMALIZAR FECHA
    const isSameDay = compareDates(rawDate, dateTargetObj);
    // 3. NORMALIZAR TEXTO
    const isTargetText = String(rawCheck).trim().toUpperCase() === config.TARGET_VAL.toUpperCase();
    // --- BLOQUE DE LOGS (solo las primeras 5 filas con ID, para depurar sin saturar el log) ---
    if (logsPrinted < 5 && rawId !== "") {
      Logger.log(
        `[Fila ${i + 1}] ` +
        `Fecha: "${rawDate}" -> ¿Ayer?: ${isSameDay} | ` +
        `Check: "${rawCheck}" -> ¿OK?: ${isTargetText} | ` +
        `Servicio: "${rawServicio}"`
      );
      logsPrinted++;
    }
    // -----------------------------------------------------------
    if (!isSameDay) continue;
    if (!isTargetText) continue;
    cleanData.push({
      joinKey: String(rawId).trim(),
      promoCode: rawCheck,
      servicio: rawServicio
    });
  }

  return cleanData;
}

// --- COMPARADOR DE FECHAS ROBUSTO ---
function compareDates(rawVal, targetDateObj) {
  if (!rawVal) return false;

  let d = null;

  if (rawVal instanceof Date) {
    d = rawVal;
  } else {
    // Si es texto "24/11/2025, 10:52"
    const datePart = String(rawVal).split(',')[0].trim();
    const parts = datePart.split('/');
    if (parts.length === 3) {
      // Mes base 0
      d = new Date(parts[2], parts[1] - 1, parts[0]);
    } else {
      d = new Date(datePart);
    }
  }
  if (!d || isNaN(d.getTime())) return false;
  return d.getDate() === targetDateObj.getDate() &&
         d.getMonth() === targetDateObj.getMonth() &&
         d.getFullYear() === targetDateObj.getFullYear();
}

/**
 * RESTO DE FUNCIONES (Cliente, Asistencia, Utilidades...)
 */
function createClienteMap(rawMatrix) {
  const map = new Map();
  const config = CONFIG.FILES.CLIENTE;
  if (!rawMatrix || rawMatrix.length <= config.SKIP_ROWS) return map;
  for (let i = config.SKIP_ROWS; i < rawMatrix.length; i++) {
    const row = rawMatrix[i];
    // Protección por si este archivo también está desplazado
    if (row.length <= config.COL_CUENTA_ORIGINAL) continue;
    const key = String(row[config.COL_ID_MAPFRE_ORIGINAL]).trim();
    if (key) {
      map.set(key, {
        email: row[config.COL_EMAIL_ORIGINAL],
        cuenta: row[config.COL_CUENTA_ORIGINAL]
      });
    }
  }
  return map;
}

function createAsistenciaMap(rawMatrix) {
  const map = new Map();
  const config = CONFIG.FILES.ASISTENCIA;
  if (!rawMatrix || rawMatrix.length <= config.SKIP_ROWS) return map;
  const headerRow = rawMatrix[config.SKIP_ROWS];
  let idxRamo = -1;
  let idxPoliza = -1;

  if (headerRow) {
    headerRow.forEach((cell, index) => {
      const txt = String(cell).toLowerCase();
      if (txt.includes("ramo")) idxRamo = index;
      if (txt.includes("número de póliza") || txt.includes("numero de poliza")) idxPoliza = index;
    });
  }
  for (let i = config.SKIP_ROWS + 1; i < rawMatrix.length; i++) {
    const row = rawMatrix[i];
    const key = String(row[config.COL_ID_MAPFRE_ORIGINAL]).trim();
    if (key) {
      map.set(key, {
        ramo: idxRamo !== -1 ? row[idxRamo] : "N/A",
        poliza: idxPoliza !== -1 ? row[idxPoliza] : "N/A"
      });
    }
  }
  return map;
}

function getDailyFiles(dateObj) {
  const findFile = (subject, prefix) => {
    const query = `subject:"${subject}" has:attachment newer_than:3d`;
    const threads = GmailApp.search(query);
    for (const thread of threads) {
      const messages = thread.getMessages();
      const attachments = messages[messages.length - 1].getAttachments();
      for (const att of attachments) {
        if (att.getName().startsWith(prefix)) return att;
      }
    }
    return null;
  };
  return {
    usos: findFile(CONFIG.FILES.USOS.SUBJECT, CONFIG.FILES.USOS.PREFIX),
    cliente: findFile(CONFIG.FILES.CLIENTE.SUBJECT, CONFIG.FILES.CLIENTE.PREFIX),
    asistencia: findFile(CONFIG.FILES.ASISTENCIA.SUBJECT, CONFIG.FILES.ASISTENCIA.PREFIX)
  };
}

function convertExcelToValues(blob) {
  if (!blob) return [];
  try {
    const resource = { title: blob.getName(), mimeType: MimeType.GOOGLE_SHEETS };
    const tempFile = Drive.Files.insert(resource, blob);
    const ss = SpreadsheetApp.openById(tempFile.id);
    const data = ss.getSheets()[0].getDataRange().getValues();
    Drive.Files.remove(tempFile.id);
    return data;
  } catch (e) {
    Logger.log("Error convirtiendo Excel: " + e.toString());
    return [];
  }
}

function saveToSheet(data) {
  const ss = SpreadsheetApp.openById(CONFIG.SHEET_DESTINO_ID);
  let sheet = ss.getSheetByName(CONFIG.SHEET_TAB_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(CONFIG.SHEET_TAB_NAME);
    sheet.appendRow(["Fecha", "Mapfre ID", "PromoCode", "Servicio", "Email Cliente", "Cuenta Cliente", "Ramo", "Póliza"]);
    sheet.getRange(1, 1, 1, 8).setFontWeight("bold");
  }
  if (data.length > 0) {
    sheet.getRange(sheet.getLastRow() + 1, 1, data.length, data[0].length).setValues(data);
  }
}

/**
 * ENVÍA EL REPORTE CON TABLA HTML
 */
function sendReportEmail(dataRows, dateStr) {
  const count = dataRows.length;
  let htmlBody = `
    <h3>Resumen Cruce Automático</h3>
    <p>El proceso ha finalizado para la fecha: <b>${dateStr}</b>.</p>
    <p>Registros procesados y añadidos: <b>${count}</b></p>
  `;
  if (count > 0) {
    // Construir tabla HTML (Limitada a 50 filas para no romper el email si hay demasiados)
    const limit = 50;
    const rowsToShow = dataRows.slice(0, limit);

    htmlBody += `
      <hr>
      <h4>Detalle de registros añadidos ${count > limit ? '(Mostrando primeros 50)' : ''}:</h4>
      <table border="1" cellpadding="5" style="border-collapse: collapse; width: 100%; font-size: 12px;">
        <tr style="background-color: #f2f2f2;">
          <th>Fecha</th>
          <th>ID</th>
          <th>PromoCode</th>
          <th>Servicio</th>
          <th>Email</th>
          <th>Cuenta</th>
          <th>Ramo</th>
          <th>Póliza</th>
        </tr>
    `;
    rowsToShow.forEach(row => {
      htmlBody += `
        <tr>
          <td>${row[0]}</td>
          <td>${row[1]}</td>
          <td>${row[2]}</td>
          <td>${row[3]}</td>
          <td>${row[4]}</td>
          <td>${row[5]}</td>
          <td>${row[6]}</td>
          <td>${row[7]}</td>
        </tr>
      `;
    });
    htmlBody += `</table>`;

    if (count > limit) {
      htmlBody += `<p><i>... y ${count - limit} filas más ver en la hoja de cálculo.</i></p>`;
    }
  } else {
    htmlBody += `<p>No se encontraron coincidencias para añadir hoy.</p>`;
  }
  GmailApp.sendEmail(
    CONFIG.EMAIL_DESTINO,
    `Resumen Cruce Automático - ${dateStr}`,
    "", // Cuerpo texto plano vacío
    {
      htmlBody: htmlBody,
      cc: CONFIG.EMAIL_CC
    }
  );
}

function sendErrorEmail(msg) {
  GmailApp.sendEmail(CONFIG.EMAIL_DESTINO, "ERROR: Automatización de cruce diario", msg);
}

function getYesterday() {
  const d = new Date();
  d.setDate(d.getDate() - 1);
  return {
    date: d,
    str: Utilities.formatDate(d, Session.getScriptTimeZone(), "dd/MM/yyyy")
  };
}
