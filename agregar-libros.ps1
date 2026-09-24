# Agrega a la biblioteca los libros de la carpeta DropBox, según DropBox\_clasificar.csv.
# Uso:  powershell -ExecutionPolicy Bypass -File .\agregar-libros.ps1 [-Prueba]
#   -Prueba  muestra lo que haría sin mover ni cambiar nada.
#
# Columnas de _clasificar.csv (ver INSTRUCCIONES.md):
#   archivo, tema, titulo, autor, idioma, escaneado, traduccion, reemplaza, drive
param([switch]$Prueba)
. (Join-Path $PSScriptRoot 'scripts\lib.ps1')

$manifiesto = Join-Path $script:Buzon '_clasificar.csv'
if (-not (Test-Path -LiteralPath $manifiesto)) { throw "No existe $manifiesto. Créalo primero (ver INSTRUCCIONES.md)." }
$columnas = 'archivo', 'tema', 'titulo', 'autor', 'idioma', 'escaneado', 'traduccion', 'reemplaza', 'drive'
$filas = @(Import-Csv -LiteralPath $manifiesto -Encoding UTF8 | ForEach-Object {
  $fila = $_; $o = [ordered]@{}
  foreach ($c in $columnas) { $o[$c] = "$($fila.$c)".Trim() }  # columnas ausentes = vacías
  [pscustomobject]$o
})
if (-not $filas) { throw '_clasificar.csv está vacío.' }

$cat = Read-Catalogo
$hoy = Get-Date -Format 'yyyy-MM-dd'
function Test-Si([string]$v) { $v.Trim().ToLower() -in 'si', 'sí', 's', 'x', '1', 'yes', 'true' }
function Test-No([string]$v) { $v.Trim().ToLower() -in 'no', 'n', '0', 'false' }

# --- 1. revisar todas las filas antes de tocar nada
$plan = @()
foreach ($f in $filas) {
  $origen = Join-Path $script:Buzon $f.archivo
  if (-not $f.archivo -or -not (Test-Path -LiteralPath $origen)) { throw "No encuentro en DropBox el archivo '$($f.archivo)'" }
  $item = Get-Item -LiteralPath $origen

  # duplicado exacto de un libro que ya está en la biblioteca
  $mismos = @($cat.libros | Where-Object { $_.s -eq $item.Length })
  $copiaDe = $null
  foreach ($m in $mismos) {
    $ruta = Get-RutaLibro $cat $m
    if ((Test-Path -LiteralPath $ruta) -and (Get-Hash $ruta) -eq (Get-Hash $origen)) { $copiaDe = $m; break }
  }
  if ($copiaDe) { $plan += [pscustomobject]@{ accion = 'duplicado'; origen = $item; libro = $copiaDe }; continue }

  if (-not $f.titulo) { throw "Falta el título para '$($f.archivo)'" }
  $viejo = $null
  if ($f.reemplaza) {
    $viejo = $cat.libros | Where-Object { $_.k -eq $f.reemplaza.Trim().ToUpper() }
    if (-not $viejo) { throw "'$($f.archivo)': el código a reemplazar $($f.reemplaza) no existe" }
    $categoria = $cat.categorias | Where-Object { $_.c -eq $viejo.c }
  } else {
    if (-not $f.tema) { throw "Falta el tema para '$($f.archivo)'" }
    $categoria = Get-Categoria $cat $f.tema
  }

  $ext = $item.Extension.TrimStart('.').ToLower()
  $b = [pscustomobject]@{ k = $null; t = $f.titulo.Trim(); c = $categoria.c; f = $ext; s = $item.Length; a = $f.autor.Trim()
    l = $(if ($f.idioma) { $f.idioma.Trim().ToLower() } else { 'es' }); sc = $false; mt = (Test-Si $f.traduccion); p = $null; d = $hoy; id = (Get-DriveId $f.drive) }
  if ($ext -eq 'pdf') {
    $info = Get-PdfInfo $origen
    $b.p = $info.paginas
    $b.sc = $info.escaneado
  }
  if (Test-Si $f.escaneado) { $b.sc = $true } elseif (Test-No $f.escaneado) { $b.sc = $false }
  if ($b.sc -and -not $f.idioma) { $b.l = $null }  # como el resto del catálogo: escaneado sin idioma confirmado
  $plan += [pscustomobject]@{ accion = $(if ($viejo) { 'reemplazo' } else { 'nuevo' }); origen = $item; libro = $b; viejo = $viejo; categoria = $categoria }
}

# --- 2. asignar códigos y mostrar el plan
foreach ($p in $plan | Where-Object accion -ne 'duplicado') {
  $p.libro.k = if ($p.viejo) { $p.viejo.k } else { Get-SiguienteCodigo $cat $p.categoria }
  if (-not $p.viejo) { [void]$cat.libros.Add($p.libro) }  # para que el siguiente del mismo tema reciba otro número
}
foreach ($p in $plan) {
  $b = $p.libro
  switch ($p.accion) {
    'duplicado' { Write-Output "= $($p.origen.Name)`n    copia exacta de $($b.k) · $($b.t) → DropBox\_duplicados" }
    default {
      $extra = @($(if ($b.p) { "$($b.p) págs" }), $(if ($b.sc) { 'escaneado' }), $(if (-not $b.id) { 'SIN enlace de Drive' })) | Where-Object { $_ }
      $verbo = if ($p.accion -eq 'reemplazo') { "reemplaza $($b.k) (el anterior → DropBox\_reemplazados)" } else { 'nuevo' }
      Write-Output "+ $($p.origen.Name)`n    $verbo → $($p.categoria.nombre)\$(Get-NombreArchivo $b)  [$($extra -join ', ')]"
    }
  }
}
if ($Prueba) { Write-Output "`n(Prueba: no se cambió nada.)"; return }

# --- 3. mover archivos y actualizar el catálogo
function Mover($item, [string]$subcarpeta) {
  $dir = Join-Path $script:Buzon $subcarpeta
  New-Item -ItemType Directory -Force $dir | Out-Null
  Move-Item -LiteralPath $item.FullName -Destination (Join-Path $dir $item.Name)
}
foreach ($p in $plan) {
  if ($p.accion -eq 'duplicado') { Mover $p.origen '_duplicados'; continue }
  if ($p.viejo) {
    $rutaVieja = Get-RutaLibro $cat $p.viejo
    if (Test-Path -LiteralPath $rutaVieja) { Mover (Get-Item -LiteralPath $rutaVieja) '_reemplazados' }
    $i = $cat.libros.IndexOf($p.viejo); $cat.libros[$i] = $p.libro
  }
  Move-Item -LiteralPath $p.origen.FullName -Destination (Get-RutaLibro $cat $p.libro)
}
Write-Catalogo $cat

$hecho = Join-Path $script:Buzon '_procesados'
New-Item -ItemType Directory -Force $hecho | Out-Null
Move-Item -LiteralPath $manifiesto -Destination (Join-Path $hecho ("clasificar_{0}.csv" -f (Get-Date -Format 'yyyy-MM-dd_HHmm')))

& (Join-Path $script:Web 'build.ps1')
