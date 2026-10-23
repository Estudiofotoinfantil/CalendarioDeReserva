# Servidor de reservas v2 (Google Apps Script)

`Servidor.gs` es el servidor que usan el formulario (`index.html`) y el panel (`Panel.html`).
La dirección del servidor se configura en un solo lugar: `config.js` (en la raíz del repo).

## Qué hace

- El formulario público solo ve horarios ocupados (sin nombres, teléfonos ni montos) y solo puede crear reservas.
- El panel entra con un código de acceso que valida el servidor (uno solo para el estudio) y se recupera por mail.
- Seña: la reserva hecha desde la web se guarda 24 h para transferir; si no llega la seña, el horario se libera solo.
  Lo que carga o toca el estudio no se libera nunca.
- Cada cambio queda anotado en la pestaña **Registro** con el valor anterior.
- Las reservas existentes no se reescriben: solo cambia la celda que se pide desde el panel.
- Copia de seguridad diaria en Drive (carpeta "Respaldos Estudio Infantil", últimas 30).

## Instalación (resumen)

La guía completa, paso a paso, está en el PDF "Guía de instalación · Servidor de reservas v2".

1. Proyecto nuevo en script.google.com → pegar `Servidor.gs`.
2. Propiedades del script: `SHEET_ID`, `CODIGO_INICIAL`, `EMAILS_ADMIN` y, si hay, `GEMINI_API_KEY` (opcional `GEMINI_MODEL`).
3. Ejecutar `revisarInstalacion`, `instalarCopiaDiaria` y `probarMail`.
4. Implementar → Nueva implementación → Aplicación web (Ejecutar como: Yo · Acceso: Cualquier persona).
5. Pegar la dirección `/exec` en `config.js` y publicar.
6. Cuando todo funcione, archivar las implementaciones del servidor viejo.

Nunca se guardan claves ni correos en este repositorio: todo eso va en las propiedades del script.
