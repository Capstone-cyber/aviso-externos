/*
 * Aviso de destinatarios externos - Smart Alerts (OnMessageSend)
 * Dominios internos de la organización (sin @):
 */
var DOMINIOS_INTERNOS = ["capstonemx.com", "capstonemexico.onmicrosoft.com", "keystonemortgage.mx"];

Office.onReady(function () {});

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
  try {
    if (!campo || !campo.getAsync) { callback([]); return; }
    campo.getAsync(function (resultado) {
      if (resultado && resultado.status === Office.AsyncResultStatus.Succeeded) {
        callback(resultado.value || []);
      } else {
        callback([]);
      }
    });
  } catch (e) {
    callback([]);
  }
}

function onMessageSendHandler(event) {
  var terminado = false;

  function terminar(opciones) {
    if (terminado) return;
    terminado = true;
    event.completed(opciones);
  }

  // Seguro: si Outlook no responde en 4 segundos, muestra un aviso general
  setTimeout(function () {
    terminar({
      allowEvent: false,
      errorMessage: "⚠️ No fue posible verificar los destinatarios. Revisa que no haya correos externos antes de enviar."
    });
  }, 4000);

  try {
    var item = Office.context.mailbox.item;
    var campos = [item.to, item.cc, item.bcc];
    var todos = [];
    var pendientes = campos.length;

    for (var i = 0; i < campos.length; i++) {
      obtenerDestinatarios(campos[i], function (lista) {
        todos = todos.concat(lista);
        pendientes--;
        if (pendientes === 0) evaluar(todos, terminar);
      });
    }
  } catch (e) {
    terminar({
      allowEvent: false,
      errorMessage: "⚠️ No fue posible verificar los destinatarios. Revisa que no haya correos externos antes de enviar."
    });
  }
}

function evaluar(destinatarios, terminar) {
  var externos = [];
  for (var i = 0; i < destinatarios.length; i++) {
    var correo = (destinatarios[i].emailAddress || "").toLowerCase().trim();
    if (correo && !esInterno(correo) && externos.indexOf(correo) === -1) {
      externos.push(correo);
    }
  }

  if (externos.length === 0) {
    terminar({ allowEvent: true });
    return;
  }

  var MAX = 8;
  var lista = externos.slice(0, MAX).join(", ");
  if (externos.length > MAX) lista += " y " + (externos.length - MAX) + " más";

  var mensaje =
    "⚠️ ATENCIÓN: AVISO DE CORREO EXTERNO. Este correo incluye " +
    (externos.length === 1 ? "1 destinatario externo" : externos.length + " destinatarios externos") +
    ": " + lista +
    ". Verifica que la información y los adjuntos puedan compartirse fuera de la empresa.";

  if (mensaje.length > 500) mensaje = mensaje.substr(0, 497) + "...";

  terminar({ allowEvent: false, errorMessage: mensaje });
}

Office.actions.associate("onMessageSendHandler", onMessageSendHandler);
