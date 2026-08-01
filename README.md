# SITI · Tickets de Mantenimiento

Proyecto mínimo para publicar en Vercel y guardar tickets en Google Sheets / Drive mediante Apps Script.

## Publicación

1. Crea un proyecto de **Google Apps Script vinculado** al Sheet de SITI y pega `Code.gs`.
2. Despliega como **Web App**: ejecutar como tu cuenta; acceso: cualquier persona con la liga. Copia la URL que termina en `/exec`.
3. Sube esta carpeta a un repositorio privado de GitHub.
4. Importa ese repositorio en Vercel y agrega la variable de entorno `APPS_SCRIPT_URL` con la URL `/exec` del paso 2.
5. Publica Vercel y sustituye `CONFIG.appUrl` en `Code.gs` por la URL definitiva de Vercel. Despliega una nueva versión del Web App.
6. Genera los QR: `https://TU-PROYECTO.vercel.app/?maquina=BIS1`, cambiando `BIS1` por cada ID de `Máquinas-LI`.

## Notas

- La API de Vercel evita el problema de CORS al conectar el HTML con Apps Script. El botón de compartir genera además una tarjeta PNG local para que el teléfono la adjunte al mensaje de WhatsApp.
- Nunca subas una URL de Apps Script ni PINs a variables visibles del navegador. Los PINs se validan sólo en Apps Script contra el Sheet.
- Los folios se bloquean con `LockService`, para evitar duplicados cuando dos operadores abren un ticket al mismo tiempo.
