# Prototipo web de inventario DASE

Este repositorio contiene el código fuente del prototipo web desarrollado en Google Apps Script para el registro de inventario de equipos de TI mediante captura de imágenes, integración con Gemini-3-Flash-Preview y almacenamiento en Google Sheets.

## Archivos principales

- Code.gs: lógica del servidor, conexión con Google Sheets, validación de acceso e integración con Gemini.
- Index.html: estructura principal de la interfaz web.
- Styles.html: estilos visuales del prototipo.
- Client.html: funciones JavaScript ejecutadas en el navegador.

## Configuración requerida

1. Crear una hoja de cálculo en Google Sheets.
2. Copiar el ID de la hoja.
3. Pegar el ID de la hoja en las propiedades del proyecto con el nombre `SPREADSHEET_ID`.
4. Configurar la clave de Gemini en las propiedades del proyecto con el nombre `GEMINI_API_KEY`.
5. Publicar el proyecto como aplicación web en Apps Script.
