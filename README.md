# Biblioteca Cristiana 031

Catálogo web de la biblioteca del grupo. El sitio contiene metadatos y enlaces a Google Drive; no contiene los libros.

## Publicación

GitHub Pages sirve `index.html` desde la rama `main`, carpeta raíz.

## Actualizar el catálogo

1. Actualiza el archivo `Catálogo - abrir en el navegador.html` de la carpeta Libreria.
2. Desde esta carpeta, ejecuta `powershell -ExecutionPolicy Bypass -File .\build.ps1`.
3. Revisa `index.html` y publícalo en GitHub.

Los libros permanecen en Google Drive. La página y sus títulos son públicos cuando GitHub Pages está activado; el acceso a los archivos depende de los permisos de Drive.