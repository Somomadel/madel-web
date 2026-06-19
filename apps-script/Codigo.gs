/**
 * Sistema de seguimiento de clientes — Madel
 * Alojar en la cuenta Somosmadel@gmail.com (Apps Script).
 */

// ===== CONFIGURACIÓN =====
const CONFIG = {
  SHEET_ID: '1OTxriDlNgafxJOzAeEeQXeOi0eNtqfc2w_RejlEVt-8',
  SHEET_NAME: 'Leads',
  EMAIL_MADEL: 'Somosmadel@gmail.com',
  URL_SITIO: 'https://somomadel.github.io/madel-web/',
  REMITENTE: 'Botrigo'
};

// Orden EXACTO de columnas en la hoja (fila 1 = encabezados).
const COLS = [
  'Fecha y Hora',           // A
  'Nombre',                 // B
  'Correo',                 // C
  'Celular',                // D
  'Estado',                 // E
  'Email enviado',          // F
  'Respuesta del cliente',  // G
  'Notas internas',         // H
  'Fuente'                  // I
];

/**
 * Valida los datos del lead. Devuelve {ok:true} o {ok:false, error:'...'}.
 */
function validarLead(datos) {
  if (!datos || typeof datos !== 'object') return { ok: false, error: 'sin-datos' };
  const nombre = (datos.nombre || '').trim();
  const correo = (datos.correo || '').trim();
  if (!nombre) return { ok: false, error: 'nombre-vacio' };
  if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(correo)) return { ok: false, error: 'correo-invalido' };
  return { ok: true };
}

function test_validarLead() {
  const casos = [
    [{ nombre: 'Ana', correo: 'ana@x.com' }, true],
    [{ nombre: '', correo: 'ana@x.com' }, false],
    [{ nombre: 'Ana', correo: 'malo' }, false],
    [{ nombre: 'Ana', correo: 'ana@x' }, false],
    [null, false]
  ];
  casos.forEach(([entrada, esperado], i) => {
    const r = validarLead(entrada).ok;
    Logger.log('caso %s: %s', i, r === esperado ? 'PASS' : 'FAIL (got ' + r + ')');
  });
}

/**
 * Devuelve la pestaña de leads, creándola con sus encabezados si no existe.
 * Garantiza que la fila 1 tenga siempre los encabezados de COLS.
 */
function getHoja() {
  const ss = SpreadsheetApp.openById(CONFIG.SHEET_ID);
  let hoja = ss.getSheetByName(CONFIG.SHEET_NAME);
  if (!hoja) {
    // Si solo hay una pestaña vacía por defecto, la reutilizamos; si no, creamos una.
    const todas = ss.getSheets();
    if (todas.length === 1 && todas[0].getLastRow() === 0) {
      hoja = todas[0].setName(CONFIG.SHEET_NAME);
    } else {
      hoja = ss.insertSheet(CONFIG.SHEET_NAME);
    }
  }
  // Asegurar encabezados en la fila 1.
  const primera = hoja.getRange(1, 1, 1, COLS.length).getValues()[0];
  const tieneEncabezados = primera[0] === COLS[0];
  if (!tieneEncabezados) {
    hoja.getRange(1, 1, 1, COLS.length).setValues([COLS]);
    hoja.getRange(1, 1, 1, COLS.length).setFontWeight('bold');
    hoja.setFrozenRows(1);
  }
  // Forzar la columna Celular a texto para que números como "+57 300..."
  // no se interpreten como fórmula (evita #ERROR!).
  const colCel = COLS.indexOf('Celular') + 1;
  hoja.getRange(2, colCel, Math.max(hoja.getMaxRows() - 1, 1), 1).setNumberFormat('@');
  return hoja;
}

/**
 * Inserta una fila de lead y devuelve el número de fila escrita.
 */
function registrarLead(datos) {
  const hoja = getHoja();
  const fila = [
    new Date(),                                   // Fecha y Hora
    (datos.nombre || '').trim(),                  // Nombre
    (datos.correo || '').trim(),                  // Correo
    (datos.celular || '').trim() || 'No proporcionado', // Celular
    '🟡 Nuevo',                                   // Estado
    '',                                           // Email enviado (se llena tras enviar)
    'No',                                         // Respuesta del cliente
    '',                                           // Notas internas
    datos.fuente || 'Modal inicio'                // Fuente
  ];
  hoja.appendRow(fila);
  return hoja.getLastRow();
}

/**
 * Escribe el resultado del envío del correo al cliente en la columna F.
 */
function marcarEmailEnviado(numFila, exito) {
  const hoja = getHoja();
  const colEmailEnviado = COLS.indexOf('Email enviado') + 1; // 1-based
  hoja.getRange(numFila, colEmailEnviado).setValue(exito ? 'Sí' : 'Fallo');
}

/**
 * Envía el correo de bienvenida (consultivo) al cliente. Lanza si falla.
 */
function enviarBienvenida(datos) {
  const nombre = (datos.nombre || '').trim();
  const cuerpo =
    'Hola ' + nombre + ', soy ' + CONFIG.REMITENTE + '. Recibimos tu mensaje y nos ' +
    'encantaría conocer mejor tu proyecto para ayudarte de la mejor forma.\n\n' +
    'Cuéntanos: ¿ya tienes página web o partes desde cero? ¿Atiendes a tus clientes ' +
    'por WhatsApp? Con eso podemos proponerte algo a tu medida.\n\n' +
    'Mientras tanto, te invitamos a conocernos más a fondo aquí:\n' +
    CONFIG.URL_SITIO + '\n\n' +
    'Quedamos atentos a tu respuesta.\n— Equipo Madel';

  GmailApp.sendEmail(
    (datos.correo || '').trim(),
    '¡Gracias por contactarnos, ' + nombre + '! 👋',
    cuerpo,
    { name: 'Madel · ' + CONFIG.REMITENTE, replyTo: CONFIG.EMAIL_MADEL }
  );
}

/**
 * Notifica a Madel del nuevo lead, con enlace directo a la hoja.
 */
function notificarMadel(datos) {
  const nombre = (datos.nombre || '').trim();
  const urlHoja = 'https://docs.google.com/spreadsheets/d/' + CONFIG.SHEET_ID;
  const html =
    '<h2>🔔 Nuevo lead — ' + nombre + '</h2>' +
    '<p><strong>👤 Nombre:</strong> ' + nombre + '<br>' +
    '<strong>📧 Correo:</strong> ' + (datos.correo || '').trim() + '<br>' +
    '<strong>📱 Celular:</strong> ' + ((datos.celular || '').trim() || 'No proporcionado') + '<br>' +
    '<strong>🕐 Fecha:</strong> ' + new Date().toLocaleString('es-CO') + '</p>' +
    '<p><a href="' + urlHoja + '" ' +
    'style="background:#1e7d4f;color:#fff;padding:10px 18px;border-radius:6px;' +
    'text-decoration:none;display:inline-block;">→ Ver en Google Sheets</a></p>';

  GmailApp.sendEmail(CONFIG.EMAIL_MADEL, '🔔 Nuevo lead — ' + nombre, '', {
    htmlBody: html
  });
}

/**
 * Punto de entrada del Web App. Recibe el POST del formulario.
 * Orden: validar -> registrar en Sheets -> notificar Madel -> bienvenida cliente.
 * El lead se registra ANTES de enviar correos para no perderlo si Gmail falla.
 */
function doPost(e) {
  try {
    const datos = JSON.parse(e.postData.contents);

    const v = validarLead(datos);
    if (!v.ok) return _json({ ok: false, error: v.error });

    const numFila = registrarLead(datos);

    // Alerta a Madel (no debe bloquear la respuesta al cliente si falla).
    try { notificarMadel(datos); } catch (err) { Logger.log('alerta Madel falló: ' + err); }

    // Bienvenida al cliente; marcar resultado en la hoja.
    let exito = false;
    try { enviarBienvenida(datos); exito = true; } catch (err) { Logger.log('bienvenida falló: ' + err); }
    marcarEmailEnviado(numFila, exito);

    return _json({ ok: true });
  } catch (err) {
    Logger.log('doPost error: ' + err);
    return _json({ ok: false, error: 'server-error', detalle: String(err) });
  }
}

/** Respuesta JSON estándar del Web App. */
function _json(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

/**
 * Prueba doPost con datos simulados. Revisar la hoja tras ejecutar:
 * debe aparecer una fila nueva de "Prueba Bot" y llegar los 2 correos.
 */
function test_doPost() {
  const fake = { postData: { contents: JSON.stringify({
    nombre: 'Prueba Bot', correo: CONFIG.EMAIL_MADEL, celular: '+57 300 000 0000',
    fuente: 'Prueba manual'
  })}};
  const res = doPost(fake);
  Logger.log(res.getContent()); // esperado: {"ok":true}
}

/**
 * Revisa Gmail y marca "Respuesta del cliente = Sí" para los leads que ya
 * respondieron por correo. Solo procesa filas con "Respuesta del cliente = No".
 */
function revisarRespuestas() {
  const hoja = getHoja();
  const datos = hoja.getDataRange().getValues(); // incluye encabezados en datos[0]

  const cCorreo = COLS.indexOf('Correo');
  const cFecha = COLS.indexOf('Fecha y Hora');
  const cEstado = COLS.indexOf('Estado');
  const cResp = COLS.indexOf('Respuesta del cliente');

  for (let i = 1; i < datos.length; i++) {
    if (String(datos[i][cResp]).normalize('NFC').trim().toLowerCase() === 'sí') continue;

    const correo = String(datos[i][cCorreo]).trim();
    if (!correo || correo.indexOf('@') === -1) continue;

    const fechaContacto = new Date(datos[i][cFecha]);
    if (respondio(correo, fechaContacto)) {
      const fila = i + 1; // 1-based
      hoja.getRange(fila, cResp + 1).setValue('Sí');
      hoja.getRange(fila, cEstado + 1).setValue('🟢 Respondió');
    }
  }
}

/**
 * ¿Hay un correo ENTRANTE de `correo` posterior a `desde`?
 */
function respondio(correo, desde) {
  // El filtro `after:` de Gmail es por día y depende de la zona horaria de la cuenta,
  // así que ampliamos la ventana un día hacia atrás para no perder respuestas cerca de
  // medianoche o por desfase de zona horaria. El filtro estricto `m.getDate() > desde`
  // de abajo es la verdadera compuerta y evita falsos positivos del histórico.
  const ventana = new Date(desde.getTime() - 24 * 60 * 60 * 1000);
  const yyyy = ventana.getFullYear();
  const mm = ('0' + (ventana.getMonth() + 1)).slice(-2);
  const dd = ('0' + ventana.getDate()).slice(-2);
  const query = 'from:' + correo + ' after:' + yyyy + '/' + mm + '/' + dd;
  const objetivo = correo.toLowerCase();
  const hilos = GmailApp.search(query, 0, 10);
  for (const hilo of hilos) {
    const mensajes = hilo.getMessages();
    for (const m of mensajes) {
      // Extrae solo la dirección de "Nombre <correo@dominio>" para comparar exacto
      // y evitar que ana@x.com coincida con ana@x.com.attacker.net.
      const match = m.getFrom().toLowerCase().match(/[^\s<>]+@[^\s<>]+/);
      const fromAddr = match ? match[0] : '';
      if (fromAddr === objetivo && m.getDate() > desde) {
        return true;
      }
    }
  }
  return false;
}

/**
 * Prueba `respondio` contra un correo real de tu bandeja.
 * Cambiar CORREO_PRUEBA por un remitente que SÍ exista en la bandeja.
 */
function test_respondio() {
  const CORREO_PRUEBA = 'algun-remitente-conocido@gmail.com';
  const hace30dias = new Date(Date.now() - 30 * 24 * 60 * 60 * 1000);
  Logger.log('respondio() = %s (esperado true si ese remitente escribió en 30 días)',
    respondio(CORREO_PRUEBA, hace30dias));
}
