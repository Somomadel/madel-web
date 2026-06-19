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
Pegar la URL `/exec` en `index.html` (reemplaza `PEGAR_AQUI_LA_URL_DEL_WEB_APP/exec`).
