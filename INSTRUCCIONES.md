# Instrucciones de mantenimiento — ALEX

Cómo agregar libros, corregir datos y publicar el catálogo. Sirve tanto para una persona como para un asistente de IA (ver la sección final).

## Dónde está cada cosa

Todo vive en la carpeta `Libreria` de Google Drive (en la computadora: `G:\My Drive\Libreria`).

| Qué | Dónde | Notas |
|---|---|---|
| Los libros | Una carpeta por tema (`Liderazgo, Pastoral e Iglesia`, …) | Nombre: `LID-208 · Título — Autor.pdf` |
| Libros nuevos por clasificar | `DropBox\` | Se vacía al procesarlos |
| **La fuente del catálogo** | `website\data\catalogo.json` | Lo único que se edita |
| Sitio web | `website\index.html` | Publicado en https://rabelo83.github.io/ALEX_707/ |
| Página local | `Catálogo - abrir en el navegador.html` | La misma página que el sitio |
| Hoja de cálculo | `Catálogo de libros.xlsx` | Se genera sola |
| Diseño de la página | `website\index.template.html` | |
| Repositorio | `website\` → https://github.com/Rabelo83/ALEX_707 | Solo el catálogo, **nunca los libros** |

**Regla de oro:** los datos se cambian solo en `catalogo.json`. La página, el sitio y la hoja se regeneran desde ahí con `build.ps1`. Si alguien edita la hoja o la página a mano, `build.ps1` se detiene y avisa, para no borrar ese trabajo.

Todos los comandos se ejecutan en PowerShell desde la carpeta `website`:

```powershell
cd "G:\My Drive\Libreria\website"
```

## Agregar libros nuevos

1. **Copia los archivos a `DropBox\`.** Espera a que Google Drive termine de subirlos (el ícono de la nube deja de girar).

2. **Crea `DropBox\_clasificar.csv`** con una fila por archivo. Se puede hacer en Excel (Guardar como → *CSV UTF-8*) o en el Bloc de notas:

   ```csv
   archivo,tema,titulo,autor,idioma,escaneado,traduccion,reemplaza,drive
   CÉLULAS  EXITOSAS .-  Joel Comiskey con Jim Egli.pdf,LID,Células exitosas,Joel Comiskey y Jim Egli,es,,,,https://drive.google.com/file/d/15Ch1NN.../view
   ```

   | Columna | Qué poner |
   |---|---|
   | `archivo` | El nombre exacto del archivo en DropBox |
   | `tema` | El prefijo del tema (tabla de abajo), p. ej. `LID` |
   | `titulo` | Título limpio, con acentos y `:` o `?` si los lleva. El nombre del archivo se ajusta solo |
   | `autor` | Vacío si no se sabe |
   | `idioma` | `es`, `en`, `pt`… Vacío = `es` |
   | `escaneado` | Vacío = se detecta solo. `si` / `no` para forzarlo |
   | `traduccion` | `si` si es traducción automática |
   | `reemplaza` | Un código existente (p. ej. `GUE-211`) si este archivo sustituye a otro. El viejo va a `DropBox\_reemplazados` |
   | `drive` | El enlace de Drive del archivo (en Drive: clic derecho → Compartir → Copiar enlace). Puede quedar vacío y agregarse después; mientras tanto el libro sale como «Pendiente» |

   Si un archivo es **copia exacta** de un libro que ya está, basta con poner su nombre en `archivo`: el script lo detecta y lo aparta en `DropBox\_duplicados`.

3. **Prueba primero** (no cambia nada):

   ```powershell
   powershell -ExecutionPolicy Bypass -File .\agregar-libros.ps1 -Prueba
   ```

   Revisa el código, la carpeta y el nombre que recibirá cada libro.

4. **Hazlo de verdad:**

   ```powershell
   powershell -ExecutionPolicy Bypass -File .\agregar-libros.ps1
   ```

   El script da el siguiente código libre del tema, renombra y mueve el archivo, cuenta las páginas, lo anota en `catalogo.json`, guarda el CSV en `DropBox\_procesados` y regenera la página y la hoja. Mover o renombrar no cambia el enlace de Drive.

5. **Comprueba y publica** (ver abajo).

## Corregir un título, autor o tema

1. Busca el código en `data\catalogo.json` y cambia el campo (tabla de campos al final).
2. Regenera y deja los archivos con el nombre nuevo:

   ```powershell
   powershell -ExecutionPolicy Bypass -File .\build.ps1
   powershell -ExecutionPolicy Bypass -File .\verificar.ps1 -Corregir
   ```

   Para **cambiar de tema** un libro hay que cambiar también su código, porque el prefijo indica el tema: usa el siguiente número libre del tema nuevo.

## Comprobar que todo coincide

```powershell
powershell -ExecutionPolicy Bypass -File .\verificar.ps1
```

Revisa que cada libro del catálogo tenga su archivo con el nombre y la carpeta correctos, que no haya archivos sin catalogar, y que la página y la hoja estén al día. `-Corregir` renombra y mueve lo que no coincida.

## Publicar en GitHub

```powershell
git add -A
git commit -m "Agrega 5 libros"
git push
```

GitHub Pages actualiza el sitio en uno o dos minutos. La página local y la hoja no se suben: están fuera de `website\`.

## Temas y prefijos

| Prefijo | Letra | Tema (= carpeta) |
|---|---|---|
| APO | A | Apologética, Sectas y Religiones |
| BIB | B | Biblias y Herramientas Bíblicas |
| CON | K | Consejería y Sanidad Interior |
| ESP | S | Espíritu Santo y lo Sobrenatural |
| EST | C | Estudio Bíblico y Comentarios |
| EVA | X | Evangelismo, Misiones y Discipulado |
| FIN | D | Finanzas y Mayordomía |
| GUE | G | Guerra Espiritual y Liberación |
| HIS | H | Historia de la Iglesia y Biografías |
| JOV | J | Jóvenes y Adolescentes |
| LID | L | Liderazgo, Pastoral e Iglesia |
| FAM | F | Matrimonio y Familia |
| MUJ | M | Mujeres |
| NIN | N | Niños y Escuela Dominical |
| NOV | R | Novelas, Testimonios y Cielo e Infierno |
| ORA | O | Oración, Ayuno y Adoración |
| PRE | P | Predicación y Homilética |
| ESC | E | Profecía y Escatología |
| TEO | T | Teología y Doctrina |
| VID | V | Vida Cristiana y Devocionales |
| OTR | Z | Otros |

Clasifica por **el asunto principal del libro**, no por palabras sueltas del título: un libro sobre el ayuno de Daniel va en Oración, no en Profecía.

## Nombres de archivo

`CÓDIGO · Título — Autor.ext` (sin `— Autor` si no se conoce). Como Windows no admite ciertos caracteres, el nombre del archivo se ajusta solo: `:` pasa a ` -`, se quitan `? * < > |`, `"` pasa a `'`, y se quita el punto final. El catálogo conserva el título correcto. Los nombres muy largos se recortan con `…`.

## Campos de `catalogo.json`

| Campo | Significado | Ejemplo |
|---|---|---|
| `k` | Código | `"LID-208"` |
| `t` | Título | `"Células exitosas"` |
| `c` | Letra del tema | `"L"` |
| `f` | Formato (extensión) | `"pdf"` |
| `s` | Tamaño en bytes (debe coincidir con el archivo) | `2932548` |
| `a` | Autor (se omite si no se sabe) | `"Joel Comiskey y Jim Egli"` |
| `l` | Idioma; `""` o ausente = sin confirmar | `"es"` |
| `sc` | `1` si es escaneado (el texto no se puede buscar) | `1` |
| `mt` | `1` si es traducción automática | `1` |
| `p` | Páginas | `145` |
| `d` | Fecha en que se agregó | `"2026-09-23"` |
| `id` | ID del archivo en Google Drive | `"15Ch1NN…"` |

## Acceso a los libros

El sitio es público (títulos y autores), pero **los archivos en Drive deben seguir restringidos** a las personas autorizadas. Tener un libro para uso personal no da permiso para distribuirlo al público; compartir en abierto libros con derechos de autor puede infringir la ley en EE. UU. y hacer que Google suspenda la cuenta. No cambies el uso compartido de las carpetas a «Cualquier persona con el enlace».

## Para asistentes de IA

Si te piden «procesar el DropBox» o «agregar libros»:

1. Lee este archivo completo. Trabaja desde `G:\My Drive\Libreria\website` con PowerShell 5.1. Los `.ps1` deben guardarse como **UTF-8 con BOM**, o los acentos se rompen.
2. Para cada archivo de `DropBox\`: busca título y autor en los metadatos del PDF, en su primera página o en el nombre. Busca en `catalogo.json` si ya existe (mismo título o autor). Si es otra versión de un libro que ya está, **pregunta** si se agrega, se reemplaza (`reemplaza`) o se descarta.
3. Si tienes el conector de Google Drive, busca la carpeta `DropBox` y lista sus archivos para obtener el `id` de cada uno y ponerlo en la columna `drive`. El ID se conserva al mover o renombrar.
4. Escribe `DropBox\_clasificar.csv`, ejecuta `agregar-libros.ps1 -Prueba`, revisa el plan y luego ejecútalo de verdad.
5. Ejecuta `verificar.ps1`, y luego `git add -A`, `git commit` y `git push` (confirma con el dueño antes de publicar).
6. Nunca cambies los permisos de Drive ni subas libros al repositorio.
