/**
 * Sistema de seguimiento de clientes — Madel
 * Alojar en la cuenta Somosmadel@gmail.com (Apps Script).
 */

// ===== CONFIGURACIÓN =====
const CONFIG = {
  SHEET_ID: 'PEGAR_AQUI_EL_ID_DE_LA_HOJA',   // se llena durante la configuración en Google
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
 * Inserta una fila de lead y devuelve el número de fila escrita.
 */
function registrarLead(datos) {
  const ss = SpreadsheetApp.openById(CONFIG.SHEET_ID);
  const hoja = ss.getSheetByName(CONFIG.SHEET_NAME);
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
  const ss = SpreadsheetApp.openById(CONFIG.SHEET_ID);
  const hoja = ss.getSheetByName(CONFIG.SHEET_NAME);
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
    return _json({ ok: false, error: 'server-error' });
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
  const ss = SpreadsheetApp.openById(CONFIG.SHEET_ID);
  const hoja = ss.getSheetByName(CONFIG.SHEET_NAME);
  const datos = hoja.getDataRange().getValues(); // incluye encabezados en datos[0]

  const cCorreo = COLS.indexOf('Correo');
  const cFecha = COLS.indexOf('Fecha y Hora');
  const cEstado = COLS.indexOf('Estado');
  const cResp = COLS.indexOf('Respuesta del cliente');

  for (let i = 1; i < datos.length; i++) {
    if (String(datos[i][cResp]).trim().toLowerCase() === 'sí') continue;

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
  // Filtra por remitente y fecha; after evita reprocesar histórico.
  const yyyy = desde.getFullYear();
  const mm = ('0' + (desde.getMonth() + 1)).slice(-2);
  const dd = ('0' + desde.getDate()).slice(-2);
  const query = 'from:' + correo + ' after:' + yyyy + '/' + mm + '/' + dd;
  const hilos = GmailApp.search(query, 0, 10);
  for (const hilo of hilos) {
    const mensajes = hilo.getMessages();
    for (const m of mensajes) {
      const fromAddr = m.getFrom().toLowerCase();
      if (fromAddr.indexOf(correo.toLowerCase()) !== -1 && m.getDate() > desde) {
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
