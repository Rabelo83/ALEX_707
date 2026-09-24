# Genera, a partir de data/catalogo.json:
#   website/index.html                          (el sitio en GitHub Pages)
#   ../Catálogo - abrir en el navegador.html    (la misma página, para abrir en la computadora)
#   ../Catálogo de libros.xlsx                  (la hoja de cálculo)
# Uso:  powershell -ExecutionPolicy Bypass -File .\build.ps1 [-Forzar]
param([switch]$Forzar)
. (Join-Path $PSScriptRoot 'scripts\lib.ps1')

$cat = Read-Catalogo
$libros = Sort-Libros $cat.categorias $cat.libros

# --- validar
$dup = $libros | Group-Object k | Where-Object Count -gt 1
if ($dup) { throw "Códigos repetidos en catalogo.json: $(($dup | ForEach-Object Name) -join ', ')" }
foreach ($b in $libros) {
  if (-not $b.k -or -not $b.t -or -not $b.f) { throw "Libro incompleto en catalogo.json: $(ConvertTo-LibroJson $b)" }
  $c = $cat.categorias | Where-Object { $_.c -eq $b.c }
  if (-not $c) { throw "$($b.k): categoría '$($b.c)' no existe" }
  if (-not $b.k.StartsWith($c.prefijo + '-')) { throw "$($b.k): el prefijo no corresponde a la categoría $($c.nombre)" }
}

# --- no pisar cambios hechos a mano en la hoja o la página local
$estadoPath = Join-Path $script:Web 'data\.generado.json'
if ((Test-Path $estadoPath) -and -not $Forzar) {
  $estado = [IO.File]::ReadAllText($estadoPath) | ConvertFrom-Json
  $cambiados = @()
  if ((Get-Hash $script:HojaLocal) -ne $estado.hoja) { $cambiados += 'Catálogo de libros.xlsx' }
  if ((Get-Hash $script:PaginaLocal) -ne $estado.pagina) { $cambiados += 'Catálogo - abrir en el navegador.html' }
  if ($cambiados) {
    throw ("Estos archivos se editaron a mano desde la última generación: $($cambiados -join ', ').`n" +
      "Pasa esos cambios a website\data\catalogo.json (la fuente) y vuelve a ejecutar con -Forzar, o se perderán.")
  }
}

# --- página web
$template = [IO.File]::ReadAllText((Join-Path $script:Web 'index.template.html'), $script:Utf8)
$booksJs = '[' + (($libros | ForEach-Object { ConvertTo-LibroJson $_ }) -join ',') + ']'
$catsJs = '[' + (($cat.categorias | ForEach-Object { '[' + (ConvertTo-JsonString $_.c) + ',' + (ConvertTo-JsonString $_.nombre) + ']' }) -join ',') + ']'
$html = $template.Replace('__BOOKS__', $booksJs).Replace('__CATS__', $catsJs)
[IO.File]::WriteAllText((Join-Path $script:Web 'index.html'), $html, $script:Utf8)

# La copia local lleva el ícono incrustado porque favicon.svg no está a su lado
$svg = [Convert]::ToBase64String([IO.File]::ReadAllBytes((Join-Path $script:Web 'favicon.svg')))
$local = $html.Replace('href="favicon.svg"', "href=""data:image/svg+xml;base64,$svg""")
[IO.File]::WriteAllText($script:PaginaLocal, $local, $script:Utf8)

# Se anota cada archivo apenas se genera, para que un fallo más adelante no parezca una edición a mano
function Save-Estado {
  $e = @{ hoja = (Get-Hash $script:HojaLocal); pagina = (Get-Hash $script:PaginaLocal) } | ConvertTo-Json
  [IO.File]::WriteAllText($estadoPath, $e, $script:Utf8)
}
Save-Estado

# --- hoja de cálculo
function Esc([string]$s) { $s.Replace('&', '&amp;').Replace('<', '&lt;').Replace('>', '&gt;') }
function Txt($ref, $v, $style = '') {
  $s = if ($style) { " s=""$style""" } else { '' }
  if ($null -eq $v -or "$v" -eq '') { "<c r=""$ref""$s t=""inlineStr"" />" } else { "<c r=""$ref""$s t=""inlineStr""><is><t>$(Esc $v)</t></is></c>" }
}
function Num($ref, $v) { "<c r=""$ref"" t=""n""><v>$(([double]$v).ToString([Globalization.CultureInfo]::InvariantCulture))</v></c>" }

$nombres = @{}; foreach ($c in $cat.categorias) { $nombres[$c.c] = $c.nombre }
$rel = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
$ultima = $libros.Count + 1
$rows = New-Object System.Text.StringBuilder
$links = New-Object System.Text.StringBuilder
$rels = New-Object System.Text.StringBuilder
$enc = 'Código', 'Título', 'Autor', 'Tema', 'Formato', 'Idioma', 'Escaneado', 'Traducción automática', 'Páginas', 'Tamaño (MB)', 'Compartido', 'Enlace'
$cols = 'A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L'
[void]$rows.Append('<row r="1">' + (-join (0..11 | ForEach-Object { Txt "$($cols[$_])1" $enc[$_] '1' })) + '</row>')
$n = 0
foreach ($b in $libros) {
  $r = $n + 2
  $cells = (Txt "A$r" $b.k) + (Txt "B$r" $b.t)
  if ($b.a) { $cells += Txt "C$r" $b.a }
  $cells += (Txt "D$r" $nombres[$b.c]) + (Txt "E$r" $b.f.ToUpper())
  $cells += Txt "F$r" $(if ($null -ne $b.l) { $script:Idiomas["$($b.l)"] } else { '' })
  $cells += (Txt "G$r" $(if ($b.sc) { 'Sí' })) + (Txt "H$r" $(if ($b.mt) { 'Sí' }))
  if ($b.p) { $cells += Num "I$r" $b.p }
  $cells += Num "J$r" ([math]::Round($b.s / 1048576, 1))
  if ($b.d) { $cells += Txt "K$r" $b.d }
  if ($b.id) {
    $n++
    $cells += Txt "L$r" 'Abrir' '2'
    [void]$links.Append("<hyperlink xmlns:r=""$rel"" ref=""L$r"" r:id=""rId$n"" />")
    [void]$rels.Append("<Relationship Type=""$rel/hyperlink"" Target=""https://drive.google.com/file/d/$($b.id)/view"" TargetMode=""External"" Id=""rId$n"" />")
  } else {
    $cells += Txt "L$r" 'Pendiente'
  }
  [void]$rows.Append("<row r=""$r"">$cells</row>")
}
$plantilla = Join-Path $script:Web 'plantilla-xlsx'
$sheet = '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetPr><outlinePr summaryBelow="1" summaryRight="1" /><pageSetUpPr /></sheetPr>' +
  "<dimension ref=""A1:L$ultima"" />" +
  '<sheetViews><sheetView workbookViewId="0"><pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen" /><selection pane="bottomLeft" activeCell="A1" sqref="A1" /></sheetView></sheetViews><sheetFormatPr baseColWidth="8" defaultRowHeight="15" />' +
  '<cols><col width="10" customWidth="1" min="1" max="1" /><col width="60" customWidth="1" min="2" max="2" /><col width="28" customWidth="1" min="3" max="3" /><col width="34" customWidth="1" min="4" max="4" /><col width="9" customWidth="1" min="5" max="5" /><col width="11" customWidth="1" min="6" max="6" /><col width="11" customWidth="1" min="7" max="7" /><col width="12" customWidth="1" min="8" max="8" /><col width="9" customWidth="1" min="9" max="9" /><col width="12" customWidth="1" min="10" max="10" /><col width="12" customWidth="1" min="11" max="11" /><col width="11" customWidth="1" min="12" max="12" /></cols>' +
  "<sheetData>$rows</sheetData><autoFilter ref=""A1:L$ultima"" />" +
  $(if ($links.Length) { "<hyperlinks>$links</hyperlinks>" } else { '' }) +
  '<pageMargins left="0.75" right="0.75" top="1" bottom="1" header="0.5" footer="0.5" /></worksheet>'
$partes = [ordered]@{
  '[Content_Types].xml' = $null; '_rels/.rels' = $null; 'docProps/app.xml' = $null; 'docProps/core.xml' = $null
  'xl/workbook.xml' = $null; 'xl/_rels/workbook.xml.rels' = $null; 'xl/styles.xml' = $null; 'xl/theme/theme1.xml' = $null
  'xl/worksheets/sheet1.xml' = $sheet
  'xl/worksheets/_rels/sheet1.xml.rels' = "<Relationships xmlns=""http://schemas.openxmlformats.org/package/2006/relationships"">$rels</Relationships>"
}
Add-Type -AssemblyName System.IO.Compression
$tmp = "$($script:HojaLocal).tmp"
$fs = [IO.File]::Open($tmp, [IO.FileMode]::Create)
$zip = New-Object IO.Compression.ZipArchive($fs, [IO.Compression.ZipArchiveMode]::Create)
foreach ($name in $partes.Keys) {
  $content = $partes[$name]
  if ($null -eq $content) { $content = [IO.File]::ReadAllText((Join-Path $plantilla $name), $script:Utf8) }
  if ($name -eq 'xl/workbook.xml') { $content = $content.Replace('__ULTIMA__', "$ultima") }
  $w = New-Object IO.StreamWriter($zip.CreateEntry($name).Open(), $script:Utf8)
  $w.Write($content); $w.Close()
}
$zip.Dispose(); $fs.Close()
# Google Drive a veces bloquea el archivo unos segundos mientras sincroniza: reintentar
for ($i = 1; ; $i++) {
  try { Move-Item -LiteralPath $tmp -Destination $script:HojaLocal -Force; break }
  catch {
    if ($i -ge 6) { Remove-Item -LiteralPath $tmp; throw "No se pudo guardar la hoja ($($_.Exception.Message)). ¿Está abierta en Excel? Ciérrala y vuelve a ejecutar build.ps1." }
    Start-Sleep -Seconds 3
  }
}

Save-Estado

$pend = @($libros | Where-Object { -not $_.id })
Write-Output "Listo: $($libros.Count) libros -> index.html, página local y hoja de cálculo."
if ($pend) { Write-Output "Aviso: $($pend.Count) libro(s) sin enlace de Drive (se muestran como 'Pendiente'): $(($pend | ForEach-Object k) -join ', ')" }
