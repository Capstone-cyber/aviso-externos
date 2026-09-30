/*
 * Aviso de destinatarios externos - Smart Alerts (OnMessageSend)
 * Dominios internos de la organización (sin @):
 */
var DOMINIOS_INTERNOS = ["capstonemx.com", "capstonemexico.onmicrosoft.com", "keystonemortgage.mx"];

function esInterno(correo) {
  var c = (correo || "").toLowerCase().trim();
  if (c.indexOf("@") === -1) return true;
  for (var i = 0; i < DOMINIOS_INTERNOS.length; i++) {
    var d = "@" + DOMINIOS_INTERNOS[i].toLowerCase();
    if (c.length >= d.length && c.substr(c.length - d.length) === d) return true;
  }
  return false;
}

function obtenerDestinatarios(campo, callback) {
  if (!campo) { callback([]); return; }
  campo.getAsync(function (resultado) {
    if (resultado.status === Office.AsyncResultStatus.Succeeded) {
      callback(resultado.value || []);
    } else {
      callback([]);
    }
  });
}

function onMessageSendHandler(event) {
  var item = Office.context.mailbox.item;
  var campos = [item.to, item.cc, item.bcc];
  var todos = [];
  var pendientes = campos.length;

  for (var i = 0; i < campos.length; i++) {
    obtenerDestinatarios(campos[i], function (lista) {
      todos = todos.concat(lista);
      pendientes--;
      if (pendientes === 0) evaluar(todos, event);
    });
  }
}

function recortar(texto) {
  return texto.length > 500 ? texto.substr(0, 497) + "..." : texto;
}

function evaluar(destinatarios, event) {
  var externos = [];
  for (var i = 0; i < destinatarios.length; i++) {
    var correo = (destinatarios[i].emailAddress || "").toLowerCase().trim();
    if (correo && !esInterno(correo) && externos.indexOf(correo) === -1) {
      externos.push(correo);
    }
  }

  if (externos.length === 0) {
    event.completed({ allowEvent: true });
    return;
  }

  var MAX = 8;
  var mostrar = externos.slice(0, MAX);
  var resto = externos.length - mostrar.length;
  var cuantos = externos.length === 1 ? "1 destinatario externo" : externos.length + " destinatarios externos";

  // Versión con formato (negritas y lista)
  var md = "**⚠️ ATENCIÓN: AVISO DE CORREO EXTERNO**\n\n" +
           "Este correo incluye " + cuantos + ":\n\n";
  for (var j = 0; j < mostrar.length; j++) md += "- " + mostrar[j] + "\n";
  if (resto > 0) md += "- y " + resto + " más\n";
  md += "\nVerifica que la información y los adjuntos puedan compartirse fuera de la empresa.";

  // Versión sencilla para versiones de Outlook que no admiten formato
  var simple = "⚠️ ATENCIÓN: AVISO DE CORREO EXTERNO. Este correo incluye " + cuantos + ": " +
               mostrar.join(", ") + (resto > 0 ? " y " + resto + " más" : "") +
               ". Verifica que la información y los adjuntos puedan compartirse fuera de la empresa.";

  event.completed({
    allowEvent: false,
    errorMessage: recortar(simple),
    errorMessageMarkdown: recortar(md)
  });
}

Office.actions.associate("onMessageSendHandler", onMessageSendHandler);
