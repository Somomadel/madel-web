# Sistema de Seguimiento y Contacto de Clientes — Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Capturar leads del formulario del sitio, registrarlos en Google Sheets, notificar a Madel, enviar correo de bienvenida al cliente, y detectar automáticamente si el cliente respondió.

**Architecture:** El formulario estático hace `POST` a un Apps Script Web App alojado en la cuenta Somosmadel@gmail.com. El Web App registra en Sheets y envía dos correos en tiempo real. Un trigger temporal cada 15 min revisa Gmail y marca las respuestas de clientes en la hoja.

**Tech Stack:** Google Apps Script (`.gs`, V8 runtime), Google Sheets, GmailApp / MailApp, HTML+JS (fetch) en `index.html`.

---

## Nota sobre el método de pruebas

Apps Script no se puede probar con un runner local (pytest/jest). Las pruebas de este plan
son de dos tipos:

1. **Funciones de auto-prueba dentro del `.gs`** (prefijo `test_`) que se ejecutan desde el
   editor de Apps Script y registran resultados con `Logger.log`.
2. **Pruebas de extremo a extremo manuales** (enviar el formulario real, revisar la hoja y
   el correo).

El código `.gs` se versiona en el repo (`apps-script/Codigo.gs`) como fuente de verdad,
aunque se ejecute pegándolo en el editor de Apps Script.

## Estructura de archivos

- **Crear** `apps-script/Codigo.gs` — todo el código del Web App: `doPost`, registro en
  Sheets, envío de correos, `revisarRespuestas`, y funciones `test_*`.
- **Crear** `apps-script/README-configuracion.md` — pasos manuales de configuración en la
  cuenta de Google (crear hoja, pegar código, publicar Web App, crear trigger).
- **Modificar** `index.html` (~líneas 2648-2661) — cambiar el destino del `fetch` de
  FormSubmit a la URL del Web App y ajustar el cuerpo del POST.

---

## Task 1: Esqueleto del Apps Script con constantes de configuración

**Files:**
- Create: `apps-script/Codigo.gs`

- [ ] **Step 1: Crear el archivo con la configuración y constantes de encabezados**

```javascript
/**
 * Sistema de seguimiento de clientes — Madel
 * Alojar en la cuenta Somosmadel@gmail.com (Apps Script).
 */

// ===== CONFIGURACIÓN =====
const CONFIG = {
  SHEET_ID: 'PEGAR_AQUI_EL_ID_DE_LA_HOJA',   // se llena en Task 6
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
```

- [ ] **Step 2: Commit**

```bash
git add apps-script/Codigo.gs
git commit -m "feat(apps-script): add config skeleton and column schema"
```

---

## Task 2: Validación de entrada

**Files:**
- Modify: `apps-script/Codigo.gs`

- [ ] **Step 1: Añadir la función de validación y su auto-prueba**

```javascript
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
```

- [ ] **Step 2: (Manual) Ejecutar la prueba — diferido a la fase de configuración**

Cuando el código esté pegado en el editor de Apps Script (Task 6), ejecutar
`test_validarLead` y verificar en el registro: 5 líneas `PASS`. Marcar este step al
validar en Task 6.

- [ ] **Step 3: Commit**

```bash
git add apps-script/Codigo.gs
git commit -m "feat(apps-script): add lead validation with self-test"
```

---

## Task 3: Registro en Google Sheets

**Files:**
- Modify: `apps-script/Codigo.gs`

- [ ] **Step 1: Añadir la función que escribe la fila del lead**

```javascript
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
```

- [ ] **Step 2: Añadir helper para escribir el estado de "Email enviado"**

```javascript
/**
 * Escribe el resultado del envío del correo al cliente en la columna F.
 */
function marcarEmailEnviado(numFila, exito) {
  const ss = SpreadsheetApp.openById(CONFIG.SHEET_ID);
  const hoja = ss.getSheetByName(CONFIG.SHEET_NAME);
  const colEmailEnviado = COLS.indexOf('Email enviado') + 1; // 1-based
  hoja.getRange(numFila, colEmailEnviado).setValue(exito ? 'Sí' : 'Fallo');
}
```

- [ ] **Step 3: Commit**

```bash
git add apps-script/Codigo.gs
git commit -m "feat(apps-script): write lead rows to Sheets"
```

---

## Task 4: Correos (cliente + alerta a Madel)

**Files:**
- Modify: `apps-script/Codigo.gs`

- [ ] **Step 1: Añadir el correo de bienvenida al cliente**

```javascript
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
```

- [ ] **Step 2: Añadir el correo de alerta a Madel con link a la hoja**

```javascript
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
```

- [ ] **Step 3: Commit**

```bash
git add apps-script/Codigo.gs
git commit -m "feat(apps-script): add client welcome and Madel alert emails"
```

---

## Task 5: Endpoint doPost (orquestación en tiempo real)

**Files:**
- Modify: `apps-script/Codigo.gs`

- [ ] **Step 1: Añadir `doPost` que orquesta validación, registro y correos**

```javascript
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
```

- [ ] **Step 2: Añadir una auto-prueba de extremo a extremo (sin enviar correos reales)**

```javascript
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
```

- [ ] **Step 3: Commit**

```bash
git add apps-script/Codigo.gs
git commit -m "feat(apps-script): add doPost orchestration endpoint"
```

---

## Task 6: Configuración en Google + validación de Tasks 1-5

**Files:**
- Create: `apps-script/README-configuracion.md`

- [ ] **Step 1: Escribir la guía de configuración**

Contenido de `apps-script/README-configuracion.md`:

```markdown
# Configuración del sistema de seguimiento — cuenta Somosmadel@gmail.com

## 1. Crear la hoja de cálculo
1. En Google Drive (logueado como Somosmadel@gmail.com) crear una hoja nueva
   llamada "Madel — Leads".
2. Renombrar la pestaña a `Leads`.
3. En la fila 1, escribir los encabezados EXACTAMENTE en este orden (A→I):
   `Fecha y Hora | Nombre | Correo | Celular | Estado | Email enviado | Respuesta del cliente | Notas internas | Fuente`
4. Copiar el ID de la hoja desde la URL:
   `https://docs.google.com/spreadsheets/d/ESTE_ES_EL_ID/edit`

## 2. Crear el proyecto de Apps Script
1. En la misma hoja: Extensiones → Apps Script.
2. Borrar el contenido por defecto y pegar TODO `apps-script/Codigo.gs`.
3. En `CONFIG.SHEET_ID`, pegar el ID copiado en el paso 1.4.
4. Guardar (Ctrl/Cmd+S).

## 3. Autorizar permisos
1. Seleccionar la función `test_validarLead` y pulsar Ejecutar.
2. Autorizar los permisos solicitados (Sheets + Gmail).
3. Ver → Registros: deben aparecer 5 líneas `PASS`.

## 4. Probar el flujo completo
1. Ejecutar `test_doPost`.
2. Verificar: una fila nueva en la hoja + correo de alerta + correo de bienvenida
   (ambos llegan a Somosmadel@gmail.com porque la prueba usa ese correo).
3. Borrar la fila de prueba de la hoja.

## 5. Publicar como Web App
1. Implementar → Nueva implementación → tipo "Aplicación web".
2. Ejecutar como: "Yo (Somosmadel@gmail.com)".
3. Quién tiene acceso: "Cualquier persona".
4. Implementar y COPIAR la URL (termina en `/exec`).

## 6. Crear el trigger de seguimiento
1. En el editor: Activadores (reloj) → Añadir activador.
2. Función: `revisarRespuestas`. Evento: "Según tiempo" → "Cada 15 minutos".
3. Guardar.

## 7. Conectar el formulario
Pegar la URL `/exec` en `index.html` (ver Task 7 del plan).
```

- [ ] **Step 2: Ejecutar la configuración real en Google**

Seguir el README recién escrito, pasos 1 a 5. Esto valida de forma efectiva los
steps diferidos de Task 2 (`test_validarLead` → 5 PASS) y Task 5 (`test_doPost` →
`{"ok":true}` + fila + 2 correos). Marcar aquellos steps al confirmar.

- [ ] **Step 3: Commit**

```bash
git add apps-script/README-configuracion.md
git commit -m "docs(apps-script): add Google setup guide"
```

---

## Task 7: Conectar el formulario en index.html

**Files:**
- Modify: `index.html:2648-2661`

- [ ] **Step 1: Reemplazar el bloque fetch de FormSubmit por el Web App**

Buscar en `index.html` (≈línea 2648) el bloque que empieza con
`const res = await fetch('https://formsubmit.co/ajax/...` y reemplazarlo por:

```javascript
      const res = await fetch('PEGAR_AQUI_LA_URL_DEL_WEB_APP/exec', {
        method:'POST',
        headers:{'Content-Type':'text/plain;charset=utf-8'},
        body: JSON.stringify({
          nombre: nombre,
          correo: correo,
          celular: celular,
          fuente: 'Modal inicio · ' + currentLang.toUpperCase()
        })
      });
      const data = await res.json().catch(() => ({ ok: res.ok }));
      if(!data.ok) throw new Error('send-failed');
```

Notas:
- `Content-Type: text/plain` evita el preflight CORS; Apps Script igual lee el cuerpo con
  `JSON.parse(e.postData.contents)`.
- Las claves del body (`nombre`, `correo`, `celular`, `fuente`) coinciden con las que lee
  `validarLead` y `registrarLead`.

- [ ] **Step 2: (Manual) Prueba de extremo a extremo desde el sitio**

1. Servir el sitio localmente o usar la versión publicada.
2. Abrir el modal, enviar el formulario con un correo propio de prueba.
3. Verificar: pantalla de éxito en el modal, fila nueva en la hoja, correo de alerta a
   Madel, y correo de bienvenida en el correo de prueba.

- [ ] **Step 3: Commit**

```bash
git add index.html
git commit -m "feat(form): point lead form to Apps Script Web App"
```

---

## Task 8: Seguimiento automático de respuestas del cliente

**Files:**
- Modify: `apps-script/Codigo.gs`

- [ ] **Step 1: Añadir `revisarRespuestas` (disparada por el trigger cada 15 min)**

```javascript
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
  // Filtra por remitente y fecha; afterEpoch evita reprocesar histórico.
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
```

- [ ] **Step 2: Añadir auto-prueba de la búsqueda**

```javascript
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
```

- [ ] **Step 3: (Manual) Validar el seguimiento de extremo a extremo**

1. Desde un correo de prueba ya registrado en la hoja (Respuesta = No), responder al
   correo de bienvenida.
2. En el editor de Apps Script, ejecutar `revisarRespuestas` manualmente.
3. Verificar en la hoja: esa fila pasa a `Respuesta del cliente = Sí` y
   `Estado = 🟢 Respondió`.

- [ ] **Step 4: Commit**

```bash
git add apps-script/Codigo.gs
git commit -m "feat(apps-script): detect client email replies via scheduled scan"
```

---

## Self-Review (completado por el autor del plan)

- **Cobertura del spec:**
  - Recepción de datos del formulario → Task 7.
  - Registro en Sheets → Task 3.
  - Email de alerta a Madel con link a la hoja → Task 4.
  - Email de bienvenida consultivo ("Hola [Nombre], soy Botrigo") + link al sitio → Task 4.
  - Operación 24/7 reactiva → Task 5 (`doPost`).
  - Seguimiento de respuestas + columna "Respuesta del cliente" → Tasks 3 y 8.
  - Esquema de 9 columnas → Tasks 1, 3, 6.
  - Manejo de errores (registro antes que correos, fallos no bloqueantes) → Task 5.
  - CORS desde GitHub Pages (text/plain, sin preflight) → Task 7.
  - Configuración manual en la cuenta de Google → Task 6.
- **Placeholders intencionales:** `SHEET_ID`, la URL del Web App y `CORREO_PRUEBA` son
  valores que solo existen tras crear los recursos en Google; se rellenan en Task 6/7/8.
  No son omisiones del plan.
- **Consistencia de tipos:** claves `nombre/correo/celular/fuente` idénticas entre el body
  del fetch (Task 7), `validarLead`, `registrarLead` y `doPost`. Orden de `COLS` usado de
  forma consistente vía `indexOf` en Tasks 3 y 8.
```

