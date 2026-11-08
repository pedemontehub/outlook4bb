/* Brew Brother para Outlook: el boton "New case" del correo abierto. Lee el asunto, quien lo mando, la fecha y el
   texto, y abre New case de Brew Brother con eso escrito (new-case.html#correo=…; leerCorreo en app.js lo lee).
   El correo va en el hash de la direccion: no pasa por ningun servidor, ni por el de la app. Va comprimido
   (deflate-raw) para que quepa y, si aun asi no cabe, va la primera parte del texto y la app lo dice. */
var APP = 'https://brewbrother101-1382170586462442.2.azure.databricksapps.com';
// lo largo que puede ser la direccion: Outlook de escritorio se la pasa al navegador y una mas larga se corta
var TOPE_URL = 2000;

Office.onReady(function () {});

function b64url(bytes) {
  var s = '';
  for (var i = 0; i < bytes.length; i++) s += String.fromCharCode(bytes[i]);
  return btoa(s).replace(/\+/g, '-').replace(/\//g, '_').replace(/=+$/, '');
}

function comprimir(texto) {
  var datos = new TextEncoder().encode(texto);
  if (typeof CompressionStream === 'undefined') return Promise.resolve('p' + b64url(datos));
  var cs = new CompressionStream('deflate-raw');
  var w = cs.writable.getWriter();
  w.write(datos);
  w.close();
  return new Response(cs.readable).arrayBuffer().then(function (ab) { return 'z' + b64url(new Uint8Array(ab)); });
}

function enlace(m) {
  return comprimir(JSON.stringify(m)).then(function (c) { return APP + '/new-case.html#correo=' + c; });
}

// el correo entero si cabe; si no, el mas largo comienzo del texto que cabe (busqueda binaria), marcado como cortado
function enlaceQueQuepa(m) {
  return enlace(m).then(function (url) {
    if (url.length <= TOPE_URL || !m.b) return url;
    var lo = 0, hi = m.b.length, mejor = null;
    function cortado(n) { return { s: m.s, f: m.f, d: m.d, b: m.b.slice(0, n), c: 1 }; }
    function paso() {
      if (hi - lo <= 20) return mejor ? Promise.resolve(mejor) : enlace(cortado(0));
      var mid = Math.floor((lo + hi) / 2);
      return enlace(cortado(mid)).then(function (u) {
        if (u.length <= TOPE_URL) { lo = mid; mejor = u; } else { hi = mid; }
        return paso();
      });
    }
    return paso();
  });
}

function aviso(texto) {
  try {
    Office.context.mailbox.item.notificationMessages.replaceAsync('brewbrother', {
      type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
      message: texto, icon: 'Icono.16', persistent: false
    });
  } catch (e) { /* sin barra de avisos no hay donde decirlo */ }
}

function abrir(url) {
  try {
    if (Office.context.ui && typeof Office.context.ui.openBrowserWindow === 'function') {
      Office.context.ui.openBrowserWindow(url);
      return;
    }
  } catch (e) { /* abajo, la ventana del navegador */ }
  window.open(url, '_blank');
}

function nuevoCaso(event) {
  var item = Office.context.mailbox.item;
  function listo() { if (event && typeof event.completed === 'function') event.completed(); }
  item.body.getAsync(Office.CoercionType.Text, function (r) {
    var de = item.from || {};
    var m = {
      s: item.subject || '',
      f: de.emailAddress ? (de.displayName ? de.displayName + ' <' + de.emailAddress + '>' : de.emailAddress) : '',
      d: item.dateTimeCreated ? new Date(item.dateTimeCreated).toISOString() : '',
      b: r.status === Office.AsyncResultStatus.Succeeded ? String(r.value || '') : ''
    };
    enlaceQueQuepa(m).then(function (url) {
      abrir(url);
      aviso('Brew Brother: the new case opens in your browser. Choose its source and create it.');
      listo();
    }, function () {
      aviso('Brew Brother could not read this email. Copy its text into New case instead.');
      listo();
    });
  });
}

// el archivo de funciones de un manifiesto XML llama a la funcion global; el runtime nuevo, a la asociada
window.nuevoCaso = nuevoCaso;
if (Office.actions && typeof Office.actions.associate === 'function') Office.actions.associate('nuevoCaso', nuevoCaso);
