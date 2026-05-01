var CORREO_AVISO = "rotoinovaqueretaro@gmail.com";
var SHEET_NAME = "Solicitudes";
var COLUMNAS = ["Fecha","Origen","Nombre","WhatsApp","Instagram","Correo","Canal","Figura","Tamano","Diseno","Lote","Molde","Cantidad","Acabado","Descripcion","Fecha Limite","Referencias","Confidencialidad"];

function doPost(e) {
  try {
    var datos = {};

    if (e.parameter && e.parameter.payload) {
      datos = JSON.parse(e.parameter.payload);
    } else if (e.postData && e.postData.type == "application/json") {
      datos = JSON.parse(e.postData.contents);
    } else if (e.postData && e.postData.contents) {
      try {
        datos = JSON.parse(e.postData.contents);
      } catch(ex) {
        datos = e.parameter;
      }
    } else {
      datos = e.parameter;
    }

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(SHEET_NAME);

    if (!sheet) {
      sheet = ss.insertSheet(SHEET_NAME);
      sheet.appendRow(COLUMNAS);
      var h = sheet.getRange(1, 1, 1, COLUMNAS.length);
      h.setBackground("#0a8ab8");
      h.setFontColor("#ffffff");
      h.setFontWeight("bold");
      sheet.setFrozenRows(1);
    }

    sheet.appendRow([
      datos.timestamp || new Date().toLocaleString(),
      datos.origen || "",
      datos.nombre || "",
      datos.whatsapp || "",
      datos.instagram || "",
      datos.correo || "",
      datos.canal || "",
      datos.figura || "",
      datos.tamano || "",
      datos.diseno || "",
      datos.lote_mixto || "",
      datos.molde || "",
      datos.cantidad || "",
      datos.acabado || "",
      datos.descripcion || "",
      datos.fecha_limite || "",
      datos.referencias || "",
      datos.confidencial || ""
    ]);

    sheet.autoResizeColumns(1, COLUMNAS.length);

    GmailApp.sendEmail(
      CORREO_AVISO,
      "Nueva solicitud Roto-Inova: " + (datos.nombre || "Cliente"),
      "Nombre: " + (datos.nombre || "") + "\nWhatsApp: " + (datos.whatsapp || "") + "\nFigura: " + (datos.figura || "") + "\nDescripcion: " + (datos.descripcion || "")
    );

    return ContentService
      .createTextOutput(JSON.stringify({status: "ok"}))
      .setMimeType(ContentService.MimeType.JSON);

  } catch(err) {
    return ContentService
      .createTextOutput(JSON.stringify({status: "error", msg: err.toString()}))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function doGet(e) {
  return ContentService
    .createTextOutput(JSON.stringify({status: "activo", app: "Roto-Inova"}))
    .setMimeType(ContentService.MimeType.JSON);
}
