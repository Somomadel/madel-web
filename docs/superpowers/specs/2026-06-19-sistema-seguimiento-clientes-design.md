# Sistema de Seguimiento y Contacto de Clientes — Madel

**Fecha:** 2026-06-19
**Estado:** Diseño aprobado
**Proyecto:** Sitio web Madel (estático, GitHub Pages)

## Objetivo

Capturar los datos de contacto del modal interactivo del sitio, registrarlos en una
hoja de cálculo de Google, notificar a Madel de cada nuevo contacto, enviar un correo
de bienvenida al cliente, y hacer seguimiento automático de si el cliente respondió.
El sistema opera 24/7 sin intervención manual.

## Arquitectura — Opción A (directa, sin intermediarios)

```
Formulario (index.html) --POST--> Apps Script Web App --> Google Sheets
                                          |--> Email de alerta a Madel
                                          |--> Email de bienvenida al cliente
                                          |
Trigger por tiempo (cada 15 min) --> revisa Gmail --> actualiza "Respuesta del cliente"
```

Tres piezas:

1. **Formulario (index.html):** ya existe. Único cambio: el `POST` deja de ir a
   `formsubmit.co/ajax/Somosmadel@gmail.com` y va a la URL del Web App de Apps Script.
   Sigue enviando `nombre`, `correo`, `celular`.

2. **Apps Script Web App** (alojado en la cuenta Somosmadel@gmail.com): publicado como
   aplicación web con acceso "cualquiera". Dos responsabilidades:
   - `doPost(e)` — reacciona en tiempo real (~1-2 s) a cada envío del formulario.
   - `revisarRespuestas()` — función disparada por un trigger temporal cada 15 min.

3. **Google Sheets:** base de datos de leads, en el Drive de Somosmadel@gmail.com.

**Decisión clave:** el registro del lead es *reactivo* (instantáneo, no hay polling).
La detección de respuestas sí es *periódica* porque la respuesta del cliente llega en
cualquier momento. El sistema vive en los servidores de Google → funciona 24/7.

## Flujo de datos

### Al enviar el formulario — `doPost(e)`

Orden de operaciones (registro primero, correos después, para no perder leads):

1. Validar: correo con formato válido y nombre no vacío. Si falla, devolver error →
   el formulario muestra su mensaje de error existente.
2. Escribir una fila nueva en Sheets con `Respuesta del cliente = No`.
3. Enviar email de alerta a Somosmadel@gmail.com (datos + link a la hoja).
4. Enviar email de bienvenida al cliente.
5. Marcar la columna `Email enviado` (Sí / fallo) según el resultado del paso 4.

### Seguimiento de respuestas — `revisarRespuestas()` (cada 15 min)

1. Recorrer las filas con `Respuesta del cliente = No`.
2. Para cada una, buscar en Gmail un correo entrante desde ese `Correo`, posterior a la
   `Fecha y Hora` de contacto.
3. Si existe, actualizar la fila: `Respuesta del cliente = Sí` y `Estado = 🟢 Respondió`.

Solo cuenta la respuesta por **email**. Respuestas por otros medios (WhatsApp) no se
detectan — comportamiento aceptado.

## Estructura de Google Sheets

Columnas:

| Columna | Origen | Valor inicial |
|---|---|---|
| Fecha y Hora | automático | timestamp del envío |
| Nombre | formulario | — |
| Correo | formulario | — |
| Celular | formulario (opcional) | — |
| Estado | seguimiento | 🟡 Nuevo |
| Email enviado | sistema | Sí / fallo |
| Respuesta del cliente | seguimiento | No → Sí |
| Notas internas | manual | (vacío) |
| Fuente | sistema | Modal inicio |

## Textos de los correos

### Email de bienvenida al cliente (consultivo + link)

> **Asunto:** ¡Gracias por contactarnos, [Nombre]! 👋
>
> Hola [Nombre], soy Botrigo. Recibimos tu mensaje y nos encantaría conocer mejor tu
> proyecto para ayudarte de la mejor forma. Cuéntanos: ¿ya tienes página web o partes
> desde cero? ¿Atiendes a tus clientes por WhatsApp? Con eso podemos proponerte algo a
> tu medida.
>
> Mientras tanto, te invitamos a conocernos más a fondo aquí:
> **https://somomadel.github.io/madel-web/**
>
> Quedamos atentos a tu respuesta. — Equipo Madel

Tono: cordial, amigable y profesional. Busca generar conversación y que el cliente
responda y elija a Madel.

### Email de alerta a Madel

Incluye: nombre, correo, celular, fecha/hora del contacto, y un botón/enlace
"Ver en Google Sheets" que abre la hoja directamente.

## Manejo de errores y confiabilidad

- **Validación** de correo y nombre en `doPost` antes de procesar.
- **Registro antes que correos:** si Gmail falla, el lead queda en la hoja; `Email
  enviado` refleja el fallo para reenvío manual.
- **Cuota de Gmail:** ~100 correos/día en cuenta gratuita; suficiente para el volumen
  esperado. Límite conocido, no bloqueante.
- **Anti-duplicados:** `revisarRespuestas()` solo procesa filas con `Respuesta = No`.
- **CORS:** Web App publicado con acceso "cualquiera"; el formulario envía sin headers
  que disparen preflight, para funcionar desde GitHub Pages.

## Pruebas

- Envío de formulario de prueba → verificar fila en Sheets + ambos correos recibidos.
- Correo inválido → verificar rechazo y mensaje de error en el formulario.
- Simular respuesta del cliente → ejecutar `revisarRespuestas()` manualmente →
  verificar que la columna pasa a `Sí`.
- Simular fallo de envío de correo → verificar que el lead persiste en la hoja.

## Configuración (manual, por el dueño de la cuenta)

Requiere acceso a Somosmadel@gmail.com:

1. Crear la hoja de Google Sheets con los encabezados.
2. Crear el proyecto de Apps Script, pegar el código, autorizar permisos
   (Sheets + Gmail).
3. Publicar como Web App (ejecutar como: yo; acceso: cualquiera) → copiar la URL.
4. Crear el trigger temporal de `revisarRespuestas()` cada 15 min.
5. Pegar la URL del Web App en el formulario de `index.html`.

## Fuera de alcance (YAGNI)

- Detección de respuestas por canales distintos a email.
- Panel/dashboard propio (la hoja de Sheets es la interfaz de gestión).
- Secuencias de seguimiento automatizadas (recordatorios, drip campaigns).
