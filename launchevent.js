/*
 * Aviso de destinatarios externos - Smart Alerts (OnMessageSend)
 * Dominios internos de la organización (sin @):
 */
var DOMINIOS_INTERNOS = ["capstonemx.com"];

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

  var MAX = 5;
  var lista = externos.slice(0, MAX).join(", ");
  if (externos.length > MAX) lista += " y " + (externos.length - MAX) + " más";

  var mensaje =
    "⚠️ ATENCIÓN: este correo va a " + externos.length +
    (externos.length === 1 ? " destinatario EXTERNO" : " destinatarios EXTERNOS") +
    " a la organización: " + lista +
    ". Verifica que la información y los adjuntos puedan compartirse fuera de la empresa.";

  if (mensaje.length > 500) mensaje = mensaje.substr(0, 497) + "...";

  event.completed({ allowEvent: false, errorMessage: mensaje });
}

Office.actions.associate("onMessageSendHandler", onMessageSendHandler);
