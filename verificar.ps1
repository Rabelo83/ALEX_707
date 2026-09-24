# Comprueba que catalogo.json, los archivos reales, la página y la hoja coinciden.
# Uso:  powershell -ExecutionPolicy Bypass -File .\verificar.ps1 [-Corregir]
#   -Corregir  renombra/mueve los archivos cuyo nombre o carpeta no coincide con el catálogo.
param([switch]$Corregir)
. (Join-Path $PSScriptRoot 'scripts\lib.ps1')

$cat = Read-Catalogo
$problemas = 0
function Problema([string]$msg) { $script:problemas++; Write-Output "  ✗ $msg" }

# Índice de archivos reales por código
$carpetas = $cat.categorias | ForEach-Object { $_.nombre }
$archivos = @{}
foreach ($nombre in $carpetas) {
  $dir = Join-Path $script:Libreria $nombre
  if (-not (Test-Path -LiteralPath $dir)) { Problema "Falta la carpeta '$nombre'"; continue }
  foreach ($f in Get-ChildItem -LiteralPath $dir -File) {
    if ($f.Name -match '^([A-Z]{3}-\d{3}) · ') {
      if ($archivos.ContainsKey($matches[1])) { Problema "Código repetido en disco: $($matches[1])" }
      $archivos[$matches[1]] = $f
    } else { Problema "Archivo sin código en '$nombre': $($f.Name)" }
  }
}

Write-Output 'Archivos:'
$codigos = @{}
foreach ($b in $cat.libros) {
  $codigos[$b.k] = $true
  $esperado = Get-RutaLibro $cat $b
  $real = $archivos[$b.k]
  if (-not $real) { Problema "$($b.k) no tiene archivo (se esperaba '$esperado')"; continue }
  if ($real.FullName -cne $esperado) {
    Problema "$($b.k) se llama '$(Split-Path (Split-Path $real.FullName) -Leaf)\$($real.Name)'`n      debería ser '$(Split-Path (Split-Path $esperado) -Leaf)\$(Split-Path $esperado -Leaf)'"
    if ($Corregir) { Move-Item -LiteralPath $real.FullName -Destination $esperado; Write-Output '      → corregido' }
  }
  elseif ($real.Length -ne $b.s) { Problema "$($b.k): el tamaño del archivo ($($real.Length)) no coincide con el catálogo ($($b.s))" }
}
foreach ($k in $archivos.Keys) { if (-not $codigos[$k]) { Problema "Archivo que no está en el catálogo: $($archivos[$k].FullName)" } }

Write-Output 'Página y hoja:'
$booksJs = '[' + ((Sort-Libros $cat.categorias $cat.libros | ForEach-Object { ConvertTo-LibroJson $_ }) -join ',') + ']'
foreach ($p in (Join-Path $script:Web 'index.html'), $script:PaginaLocal) {
  if (-not ([IO.File]::ReadAllText($p, $script:Utf8)).Contains("const BOOKS = $booksJs;")) { Problema "$(Split-Path $p -Leaf) está desactualizada: ejecuta build.ps1" }
}
$estadoPath = Join-Path $script:Web 'data\.generado.json'
if (-not (Test-Path $estadoPath)) { Problema 'La hoja nunca se generó con build.ps1' }
else {
  $estado = [IO.File]::ReadAllText($estadoPath) | ConvertFrom-Json
  if ((Get-Hash $script:HojaLocal) -ne $estado.hoja) { Problema 'Catálogo de libros.xlsx se editó a mano (ver INSTRUCCIONES.md)' }
  elseif ((Get-Item $script:CatalogoJson).LastWriteTime -gt (Get-Item $script:HojaLocal).LastWriteTime) { Problema 'Catálogo de libros.xlsx es más vieja que catalogo.json: ejecuta build.ps1' }
}

$sinId = @($cat.libros | Where-Object { -not $_.id })
if ($sinId) { Write-Output "Aviso: sin enlace de Drive: $(($sinId | ForEach-Object k) -join ', ')" }
$enBuzon = @(Get-ChildItem -LiteralPath $script:Buzon -File -ErrorAction SilentlyContinue | Where-Object { $_.Name -notlike '_*' -and $_.Name -ne 'LEEME.txt' })
if ($enBuzon) { Write-Output "Aviso: $($enBuzon.Count) archivo(s) esperando en DropBox" }

if ($problemas) { Write-Output "`n$problemas problema(s) encontrados."; exit 1 }
Write-Output "`nTodo coincide: $($cat.libros.Count) libros."
