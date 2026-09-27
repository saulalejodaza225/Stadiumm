/**
 * CONEXIÓN DEL FORMULARIO DE INSCRIPCIÓN CON GOOGLE SHEETS — YA CONFIGURADA Y EN VIVO
 * ===================================================================================
 *
 * Este archivo documenta la integración que ya quedó funcionando. No necesitas hacer
 * nada más: index.html y code.html ya tienen la URL real cargada en
 * REGISTRO_ENDPOINT_URL y el roster de la página ya lee datos reales de esta hoja.
 *
 * Hoja de cálculo ("tu Excel en vivo"):
 *   https://docs.google.com/spreadsheets/d/1PMEGEb8apyDPI7QUFfdgavCy2_V3q7o85XK-nHt3ydc/edit
 *   Nombre: "Inscripciones Colombia Cup". Cada inscripción del formulario aparece ahí
 *   como una fila nueva (Fecha, Nombre Completo, Epic Games ID, WhatsApp, Comprobante
 *   de Pago). Se puede exportar a .xlsx en cualquier momento desde Archivo > Descargar.
 *
 * Proyecto de Apps Script (el "backend" gratuito que conecta el formulario con la hoja):
 *   https://script.google.com/home/projects/1UjvzUOwS4XhBLQG2RVFImQACsAju_uEJTozjRs9mz3mVH_aUXhSm9LcT/edit
 *
 * URL de la aplicación web desplegada (la que está pegada en REGISTRO_ENDPOINT_URL):
 *   https://script.google.com/macros/s/AKfycbwXqybOQw6kcjeX7mrmzN4AfHcBwza48ux8uFesmqZjnru2Y5aaHwWaDxH8c_WDqFZoWA/exec
 *
 * Comprobantes de pago subidos por los jugadores se guardan automáticamente en una
 * carpeta de Google Drive llamada "Comprobantes Colombia Cup", y el enlace de cada
 * uno queda en la última columna de la hoja.
 *
 * Qué hace el script (doGet / doPost más abajo):
 *   1. doPost  -> valida que no se hayan alcanzado los 100 cupos y, si hay espacio,
 *                 guarda la inscripción como fila nueva (con lock para evitar que dos
 *                 envíos simultáneos exactamente en el cupo 100 lo sobrepasen).
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

function getSheet_() {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) sheet = ss.insertSheet(SHEET_NAME);
  if (sheet.getLastRow() === 0) sheet.appendRow(['Fecha', 'Nombre Completo', 'Epic Games ID', 'WhatsApp', 'Comprobante de Pago']);
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

    if (data.comprobanteBase64 && data.comprobanteNombre) {
      var folderName = 'Comprobantes Colombia Cup';
      var folders = DriveApp.getFoldersByName(folderName);
      var folder = folders.hasNext() ? folders.next() : DriveApp.createFolder(folderName);

      var bytes = Utilities.base64Decode(data.comprobanteBase64);
      var blob = Utilities.newBlob(bytes, data.comprobanteTipo || 'application/octet-stream', data.comprobanteNombre);
      var file = folder.createFile(blob);
      file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      comprobanteUrl = file.getUrl();
    }

    sheet.appendRow([
      new Date(),
      data.nombre || '',
      data.epicId || '',
      data.whatsapp || '',
      comprobanteUrl
    ]);

    return ContentService
      .createTextOutput(JSON.stringify({ result: 'success' }))
      .setMimeType(ContentService.MimeType.JSON);
  } finally {
    lock.releaseLock();
  }
}
