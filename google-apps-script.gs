/**
 * CONEXIÓN DEL FORMULARIO DE INSCRIPCIÓN CON GOOGLE SHEETS — YA CONFIGURADA Y EN VIVO
 * ===================================================================================
 *
 * Este archivo documenta la integración que ya quedó funcionando. index.html y
 * code.html ya tienen la URL real cargada en REGISTRO_ENDPOINT_URL y el roster de la
 * página ya lee datos reales de esta hoja.
 *
 * Hoja de cálculo ("tu Excel en vivo"):
 *   https://docs.google.com/spreadsheets/d/1PMEGEb8apyDPI7QUFfdgavCy2_V3q7o85XK-nHt3ydc/edit
 *   Cada inscripción del formulario aparece ahí como una fila nueva: Fecha, Nombre
 *   Completo, Epic Games ID, WhatsApp, Comprobante de Pago, Estado de Verificación,
 *   Detalle de Verificación. Se puede exportar a .xlsx en cualquier momento desde
 *   Archivo > Descargar.
 *
 * Proyecto de Apps Script (el "backend" gratuito que conecta el formulario con la hoja):
 *   https://script.google.com/home/projects/1UjvzUOwS4XhBLQG2RVFImQACsAju_uEJTozjRs9mz3mVH_aUXhSm9LcT/edit
 *
 * URL de la aplicación web desplegada (la que está pegada en REGISTRO_ENDPOINT_URL):
 *   https://script.google.com/macros/s/AKfycbwXqybOQw6kcjeX7mrmzN4AfHcBwza48ux8uFesmqZjnru2Y5aaHwWaDxH8c_WDqFZoWA/exec
 *
 * Comprobantes de pago subidos por los jugadores se guardan automáticamente en una
 * carpeta de Google Drive llamada "Comprobantes Colombia Cup", y el enlace de cada
 * uno queda en la columna E de la hoja.
 *
 * === NUEVO: VERIFICACIÓN AUTOMÁTICA DE COMPROBANTES (con IA de Google) ===
 *   Cada vez que alguien se inscribe, el script le pide a la IA de Google (Gemini)
 *   que mire la imagen del comprobante y responda si de verdad parece una transferencia
 *   Nequi/Bre-B/Daviplata y por cuánto fue. Con eso llena dos columnas nuevas:
 *     F. Estado de Verificación -> "Verificado" o "Revisar Manualmente"
 *     G. Detalle de Verificación -> explica por qué (monto detectado, o el motivo de duda)
 *   Esto NO bloquea el cupo del jugador (su fila se guarda igual, para no arriesgar
 *   dejar a alguien real por fuera si la IA se equivoca o falla). Es una ayuda para que,
 *   antes de agregar a alguien al grupo de WhatsApp / lobby, revises solo la lista de
 *   "Revisar Manualmente" en vez de las 100 filas completas.
 *
 *   PASO OBLIGATORIO para activar la verificación (una sola vez):
 *     1. Abre el link del "Proyecto de Apps Script" de arriba.
 *     2. Ve a https://aistudio.google.com/apikey e inicia sesión con tu cuenta de Google.
 *     3. Crea una API key gratuita (no pide tarjeta) y cópiala.
 *     4. En el editor de Apps Script: ícono de engranaje (Configuración del proyecto)
 *        > "Propiedades del script" > "Añadir propiedad de script".
 *        Nombre: GEMINI_API_KEY   Valor: (pega tu API key)  > Guardar.
 *     5. Vuelve al editor de código, selecciona la función "configurarTodo" en el
 *        desplegable de funciones (arriba, junto a "Depurar") y presiona "Ejecutar" una
 *        sola vez. Esto crea las dos pestañas filtradas (ver abajo) y el formato de
 *        colores. La primera vez te pedirá autorizar permisos: acepta.
 *   Si no configuras la API key, todas las inscripciones quedarán marcadas como
 *   "Revisar Manualmente" (no se rompe nada, simplemente no hay verificación automática).
 *
 * === NUEVO: DOS PESTAÑAS QUE SE FILTRAN SOLAS ===
 *   La función "configurarTodo" (ver arriba) crea dos pestañas en esta misma hoja:
 *     "Verificados"          -> solo las filas marcadas como Verificado
 *     "Revisar Manualmente"  -> solo las filas que necesitan que tú las mires
 *   Son fórmulas QUERY en vivo: se actualizan solas cada vez que entra una inscripción
 *   nueva, no hay que volver a ejecutar nada.
 *
 * Qué hace el script (doGet / doPost más abajo):
 *   1. doPost  -> valida que no se hayan alcanzado los 100 cupos, guarda la inscripción
 *                 como fila nueva (con lock para evitar que dos envíos simultáneos
 *                 exactamente en el cupo 100 lo sobrepasen), sube el comprobante a Drive
 *                 y le pide a la IA que lo verifique.
 *   2. doGet   -> le devuelve a la página web la lista de inscritos reales en JSON,
 *                 para pintar la tabla de roster y los contadores 0-100 en vivo.
 *   3. Cierre automático -> si ya hay 100 filas, doPost rechaza inscripciones nuevas
 *                 aunque alguien intente saltarse el formulario (curl, etc.).
 *
 * SI NECESITAS MODIFICAR EL SCRIPT:
 *   1. Abre el link del "Proyecto de Apps Script" de arriba.
 *   2. Edita el código, guarda (Ctrl+S).
 *   3. Implementar > Gestionar implementaciones > ícono de lápiz (editar) >
 *      en "Versión" elige "Nueva versión" > Implementar.
 *      (Si creas una implementación totalmente nueva en vez de editar la existente,
 *      la URL cambia y tendrías que actualizar REGISTRO_ENDPOINT_URL en el sitio).
 *
 * ===================================================================================
 * CÓDIGO ACTUALMENTE DESPLEGADO:
 * ===================================================================================
 */

var MAX_SLOTS = 100;
var SHEET_ID = '1PMEGEb8apyDPI7QUFfdgavCy2_V3q7o85XK-nHt3ydc';
var SHEET_NAME = 'Inscripciones';
var MONTO_ESPERADO_COP = 5000;
var HEADERS = ['Fecha', 'Nombre Completo', 'Epic Games ID', 'WhatsApp', 'Comprobante de Pago', 'Estado de Verificación', 'Detalle de Verificación'];
var ESTADO_VERIFICADO = 'Verificado';
var ESTADO_REVISAR = 'Revisar Manualmente';

function getSheet_() {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) sheet = ss.insertSheet(SHEET_NAME);
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(HEADERS);
  } else {
    var lastCol = sheet.getLastColumn();
    if (lastCol < HEADERS.length) {
      sheet.getRange(1, lastCol + 1, 1, HEADERS.length - lastCol).setValues([HEADERS.slice(lastCol)]);
    }
  }
  return sheet;
}

function doGet(e) {
  var sheet = getSheet_();
  var lastRow = sheet.getLastRow();
  var jugadores = [];
  if (lastRow > 1) {
    var rows = sheet.getRange(2, 1, lastRow - 1, 4).getValues();
    jugadores = rows.map(function (row) {
      return { fecha: row[0], nombre: row[1], epicId: row[2] };
    });
  }
  return ContentService
    .createTextOutput(JSON.stringify({ jugadores: jugadores, maxSlots: MAX_SLOTS }))
    .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var sheet = getSheet_();
    var currentCount = sheet.getLastRow() - 1;
    if (currentCount >= MAX_SLOTS) {
      return ContentService
        .createTextOutput(JSON.stringify({ result: 'error', message: 'Registro cerrado: ya se alcanzaron los ' + MAX_SLOTS + ' jugadores.' }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    var data = JSON.parse(e.postData.contents);
    var comprobanteUrl = '';
    var verificacion = { estado: ESTADO_REVISAR, detalle: 'No se recibió comprobante de pago.' };

    if (data.comprobanteBase64 && data.comprobanteNombre) {
      var folderName = 'Comprobantes Colombia Cup';
      var folders = DriveApp.getFoldersByName(folderName);
      var folder = folders.hasNext() ? folders.next() : DriveApp.createFolder(folderName);

      var bytes = Utilities.base64Decode(data.comprobanteBase64);
      var blob = Utilities.newBlob(bytes, data.comprobanteTipo || 'application/octet-stream', data.comprobanteNombre);
      var file = folder.createFile(blob);
      file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      comprobanteUrl = file.getUrl();

      verificacion = verificarComprobante_(data.comprobanteBase64, data.comprobanteTipo);
    }

    sheet.appendRow([
      new Date(),
      data.nombre || '',
      data.epicId || '',
      data.whatsapp || '',
      comprobanteUrl,
      verificacion.estado,
      verificacion.detalle
    ]);

    return ContentService
      .createTextOutput(JSON.stringify({ result: 'success' }))
      .setMimeType(ContentService.MimeType.JSON);
  } finally {
    lock.releaseLock();
  }
}

/**
 * Le pide a Gemini (IA gratuita de Google) que mire la imagen del comprobante y diga
 * si de verdad parece una transferencia Nequi/Bre-B/Daviplata y por cuánto fue.
 * Si algo falla (sin API key, error de red, respuesta rara) siempre devuelve
 * ESTADO_REVISAR en vez de fallar la inscripción: el jugador nunca pierde su cupo
 * por un problema de la verificación automática.
 */
function verificarComprobante_(base64, mimeType) {
  var apiKey = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
  if (!apiKey) {
    return { estado: ESTADO_REVISAR, detalle: 'Verificación automática no configurada (falta GEMINI_API_KEY).' };
  }

  try {
    var prompt = 'Analiza esta imagen de un comprobante de pago colombiano (transferencia Nequi, Bre-B o Daviplata). ' +
      'Responde EXCLUSIVAMENTE con un JSON valido, sin texto adicional ni bloques de codigo, con este formato exacto: ' +
      '{"es_comprobante": true o false, "monto": numero_en_pesos_colombianos_o_null, "razon": "breve explicacion en español"}. ' +
      '"es_comprobante" debe ser true solo si la imagen muestra con claridad una confirmacion o comprobante de una transferencia o pago exitoso. ' +
      '"monto" es el valor numerico transferido en pesos colombianos (COP), solo digitos sin puntos ni simbolos, o null si no se puede leer con certeza.';

    var payload = {
      contents: [{
        parts: [
          { text: prompt },
          { inline_data: { mime_type: mimeType || 'image/jpeg', data: base64 } }
        ]
      }],
      generationConfig: { temperature: 0, responseMimeType: 'application/json' }
    };

    var response = UrlFetchApp.fetch(
      'https://generativelanguage.googleapis.com/v1beta/models/gemini-2.0-flash:generateContent?key=' + apiKey,
      {
        method: 'post',
        contentType: 'application/json',
        payload: JSON.stringify(payload),
        muteHttpExceptions: true
      }
    );

    if (response.getResponseCode() !== 200) {
      return { estado: ESTADO_REVISAR, detalle: 'Error del servicio de verificación (HTTP ' + response.getResponseCode() + ').' };
    }

    var json = JSON.parse(response.getContentText());
    var parts = json.candidates && json.candidates[0] && json.candidates[0].content && json.candidates[0].content.parts;
    var textoRespuesta = parts && parts[0] && parts[0].text;
    if (!textoRespuesta) {
      return { estado: ESTADO_REVISAR, detalle: 'Respuesta vacía del servicio de verificación.' };
    }

    var resultado = JSON.parse(textoRespuesta);

    if (!resultado.es_comprobante) {
      return { estado: ESTADO_REVISAR, detalle: 'La imagen no parece un comprobante de pago. ' + (resultado.razon || '') };
    }
    if (resultado.monto === null || resultado.monto === undefined || resultado.monto === '') {
      return { estado: ESTADO_REVISAR, detalle: 'No se pudo leer el monto con certeza. ' + (resultado.razon || '') };
    }
    if (Number(resultado.monto) !== MONTO_ESPERADO_COP) {
      return { estado: ESTADO_REVISAR, detalle: 'Monto detectado: $' + resultado.monto + ' COP (se esperaban $' + MONTO_ESPERADO_COP + ').' };
    }

    return { estado: ESTADO_VERIFICADO, detalle: 'Monto confirmado: $' + resultado.monto + ' COP.' };
  } catch (err) {
    return { estado: ESTADO_REVISAR, detalle: 'Error al verificar automáticamente: ' + err.message };
  }
}

/**
 * Ejecuta esto UNA SOLA VEZ desde el editor de Apps Script (menú de funciones de
 * arriba > elegir "configurarTodo" > botón Ejecutar) después de guardar tu
 * GEMINI_API_KEY en Propiedades del script. Crea las pestañas filtradas y el color
 * de estado. Se puede volver a ejecutar sin problema si algo se borra sin querer.
 */
function configurarTodo() {
  var sheet = getSheet_();
  aplicarFormatoCondicional_(sheet);
  configurarPestanasFiltradas_();
}

function aplicarFormatoCondicional_(sheet) {
  var rango = sheet.getRange('F2:F1000');
  var reglasExistentes = sheet.getConditionalFormatRules().filter(function (regla) {
    return !regla.getRanges().some(function (r) { return r.getA1Notation() === rango.getA1Notation(); });
  });
  var reglaVerificado = SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo(ESTADO_VERIFICADO)
    .setBackground('#d9ead3')
    .setRanges([rango])
    .build();
  var reglaRevisar = SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo(ESTADO_REVISAR)
    .setBackground('#fce5cd')
    .setRanges([rango])
    .build();
  reglasExistentes.push(reglaVerificado, reglaRevisar);
  sheet.setConditionalFormatRules(reglasExistentes);
}

/**
 * Hojas en idiomas como español usan coma como separador decimal, por lo que Sheets
 * exige punto y coma entre los argumentos de una función (en vez de coma). Esta función
 * detecta eso a partir del idioma configurado en la hoja para armar la fórmula bien.
 */
function separadorDeArgumentos_() {
  var idioma = (SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetLocale() || '').split(/[-_]/)[0].toLowerCase();
  var idiomasConComaDecimal = ['es', 'fr', 'de', 'it', 'pt', 'nl', 'pl', 'ru', 'tr', 'sv', 'fi', 'da', 'nb', 'no', 'cs', 'sk', 'ro', 'hu', 'el', 'uk'];
  return idiomasConComaDecimal.indexOf(idioma) !== -1 ? ';' : ',';
}

function configurarPestanasFiltradas_() {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sep = separadorDeArgumentos_();

  var verificadosSheet = ss.getSheetByName('Verificados') || ss.insertSheet('Verificados');
  verificadosSheet.clear();
  verificadosSheet.getRange('A1').setFormula(
    "=QUERY('" + SHEET_NAME + "'!A:G" + sep + "\"select A,B,C,D,E,F,G where F = '" + ESTADO_VERIFICADO + "'\"" + sep + "1)"
  );

  var revisarSheet = ss.getSheetByName('Revisar Manualmente') || ss.insertSheet('Revisar Manualmente');
  revisarSheet.clear();
  revisarSheet.getRange('A1').setFormula(
    "=QUERY('" + SHEET_NAME + "'!A:G" + sep + "\"select A,B,C,D,E,F,G where F = '" + ESTADO_REVISAR + "'\"" + sep + "1)"
  );
}
