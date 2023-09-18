function abrirFormulario() {
  var html = HtmlService.createHtmlOutputFromFile('gui')
      .setWidth(400)
      .setHeight(300);
  SpreadsheetApp.getUi().showModalDialog(html, 'Crear Borrador Gmail');
}

function generarBorradorGmail(rangoNombre) {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet();
    var range = sheet.getRangeByName(rangoNombre);
    var data = range.getValues();
    var fecha = Utilities.formatDate(new Date(), "GMT-03:00", "yyyy-MM-dd HH:mm:ss");

    var leerColumna = function(hoja, separador = "<br>") {
      return sheet.getSheetByName(hoja).getRange("A:A").getValues().flat().filter(String).join(separador);
    };

    var asunto = leerColumna("Asunto") + " " + fecha;
    var destinatarios = leerColumna("Destinatarios", ", ");
    var cc = leerColumna("cc", ", ");
    var cuerpoPre = leerColumna("cuerpo_del_mensaje_pre_tabla");
    var cuerpoPost = leerColumna("cuerpo_del_mensaje_post_tabla");

    var mensaje = `<div style='font-family: sans-serif; font-size: 12px;'>`;
    mensaje += `<p>${cuerpoPre.replace(/\n/g, '<br>')}</p>`;
    mensaje += "<table style='border-collapse: collapse;'>";

    for (var i = 0; i < data.length; i++) {
      mensaje += "<tr>";
      
      var bgColor = i === 0 ? "#2E748C" : "white";
      var fontColor = i === 0 ? "white" : "black";

      for (var j = 0; j < data[i].length; j++) {
        var style = `color: ${fontColor}; background-color: ${bgColor}; font-family: sans-serif; font-size: 12px; padding: 8px; border: 1px solid #ccc;`;
        mensaje += `<td style='${style}'>` + data[i][j] + "</td>";
      }

      mensaje += "</tr>";
    }

    mensaje += "</table>";
    mensaje += `<p>${cuerpoPost.replace(/\n/g, '<br>')}</p>`;
    mensaje += "</div>";

    GmailApp.createDraft(destinatarios, asunto, "", {htmlBody: mensaje, cc: cc});
    
    return "Borrador creado con éxito! 🎉";
  } catch (e) {
    console.error("Ocurrió un error: " + e);
    var mensajeError = "Error en el documento " + SpreadsheetApp.getActiveSpreadsheet().getName();
    mensajeError += " y en el rango " + rangoNombre + ". ";
    mensajeError += "Detalle del error: " + e.toString();
    return mensajeError;
  }
}
