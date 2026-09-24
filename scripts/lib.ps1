# Funciones compartidas por build.ps1, agregar-libros.ps1 y verificar.ps1.
# Guardar este archivo como UTF-8 con BOM (PowerShell 5.1 lo necesita para leer los acentos).

$ErrorActionPreference = 'Stop'
$script:Utf8 = New-Object System.Text.UTF8Encoding($false)
$script:Web = Split-Path $PSScriptRoot -Parent
$script:Libreria = Split-Path $script:Web -Parent
$script:CatalogoJson = Join-Path $script:Web 'data\catalogo.json'
$script:PaginaLocal = Join-Path $script:Libreria 'Catálogo - abrir en el navegador.html'
$script:HojaLocal = Join-Path $script:Libreria 'Catálogo de libros.xlsx'
$script:Buzon = Join-Path $script:Libreria 'DropBox'
$script:Idiomas = @{ es = 'Español'; en = 'Inglés'; pt = 'Portugués'; fr = 'Francés'; it = 'Italiano'; '' = '' }

# ---------- catálogo (data/catalogo.json) ----------

function Read-Catalogo {
  $raw = [IO.File]::ReadAllText($script:CatalogoJson, $script:Utf8)
  $data = $raw | ConvertFrom-Json
  $books = New-Object System.Collections.ArrayList
  foreach ($b in $data.libros) { [void]$books.Add($b) }
  [pscustomobject]@{ categorias = @($data.categorias); libros = $books }
}

function ConvertTo-JsonString([string]$s) {
  $sb = New-Object System.Text.StringBuilder
  [void]$sb.Append('"')
  foreach ($ch in $s.ToCharArray()) {
    switch ($ch) {
      '"' { [void]$sb.Append('\"') }
      '\' { [void]$sb.Append('\\') }
      default { if ([int]$ch -lt 32) { [void]$sb.AppendFormat('\u{0:x4}', [int]$ch) } else { [void]$sb.Append($ch) } }
    }
  }
  [void]$sb.Append('"')
  $sb.ToString()
}

# Un libro en una línea, con el orden de campos de siempre y sin campos vacíos.
function ConvertTo-LibroJson($b) {
  $parts = @()
  foreach ($key in 'k', 't', 'c', 'f', 's', 'a', 'l', 'sc', 'mt', 'p', 'd', 'id') {
    $v = $b.$key
    if ($null -eq $v) { continue }
    if ($key -eq 'l') { $parts += '"l":' + (ConvertTo-JsonString $v); continue }  # "l" vacío = idioma desconocido
    if ($v -is [string] -and $v -eq '') { continue }
    if ($key -in 'sc', 'mt') { if ($v) { $parts += """${key}"":1" }; continue }
    if ($v -is [string]) { $parts += """${key}"":" + (ConvertTo-JsonString $v) } else { $parts += """${key}"":$v" }
  }
  '{' + ($parts -join ',') + '}'
}

function Get-OrdenCategoria($cats, [string]$c) {
  for ($i = 0; $i -lt $cats.Count; $i++) { if ($cats[$i].c -eq $c) { return $i } }
  throw "Categoría desconocida: $c"
}

function Sort-Libros($cats, $books) {
  $order = @{}; for ($i = 0; $i -lt $cats.Count; $i++) { $order[$cats[$i].c] = $i }
  @($books | Sort-Object @{ Expression = { $order[$_.c] } }, @{ Expression = { $_.k } })
}

function Write-Catalogo($cat) {
  $sorted = Sort-Libros $cat.categorias $cat.libros
  $lines = @('{', '  "categorias": [')
  $catLines = foreach ($c in $cat.categorias) { '    {"c":' + (ConvertTo-JsonString $c.c) + ',"nombre":' + (ConvertTo-JsonString $c.nombre) + ',"prefijo":' + (ConvertTo-JsonString $c.prefijo) + '}' }
  $lines += ($catLines -join ",`n")
  $lines += '  ],'
  $lines += '  "libros": ['
  $lines += (($sorted | ForEach-Object { '    ' + (ConvertTo-LibroJson $_) }) -join ",`n")
  $lines += '  ]'
  $lines += '}'
  [IO.File]::WriteAllText($script:CatalogoJson, ($lines -join "`n") + "`n", $script:Utf8)
}

function Get-Categoria($cat, [string]$valor) {
  $v = $valor.Trim()
  $m = @($cat.categorias | Where-Object { $_.c -eq $v -or $_.prefijo -eq $v.ToUpper() -or $_.nombre -eq $v })
  if ($m.Count -ne 1) { throw "No reconozco el tema '$valor'. Usa el prefijo (p. ej. LID), la letra o el nombre exacto de la carpeta." }
  $m[0]
}

function Get-SiguienteCodigo($cat, $categoria) {
  $max = 0
  foreach ($b in $cat.libros) {
    if ($b.k -like "$($categoria.prefijo)-*") { $n = [int]($b.k.Substring(4)); if ($n -gt $max) { $max = $n } }
  }
  '{0}-{1:D3}' -f $categoria.prefijo, ($max + 1)
}

# ---------- nombres de archivo ----------

# Windows no admite : ? * < > | " / \ en nombres, ni un punto o espacio al final.
function Get-NombreSeguro([string]$s) {
  $s = $s -replace '\s*:\s*', ' - ' -replace '[?*<>|]', '' -replace '"', "'" -replace '[/\\]', '-'
  $s = ($s -replace '\s+', ' ').Trim()
  $s.TrimEnd('.', ' ')
}

function Get-NombreArchivo($b) {
  $titulo = Get-NombreSeguro $b.t
  $autor = if ($b.a) { Get-NombreSeguro $b.a } else { '' }
  $cuerpo = if ($autor) { "$titulo — $autor" } else { $titulo }
  $max = 150
  if ($cuerpo.Length -gt $max) { $cuerpo = $cuerpo.Substring(0, $max - 1).TrimEnd() + '…' }
  "$($b.k) · $cuerpo.$($b.f)"
}

function Get-RutaLibro($cat, $b) {
  $nombreCat = ($cat.categorias | Where-Object { $_.c -eq $b.c }).nombre
  Join-Path (Join-Path $script:Libreria $nombreCat) (Get-NombreArchivo $b)
}

# ---------- PDF: páginas y si es escaneado ----------

function Expand-PdfStreams([string]$raw) {
  # Descomprime los flujos FlateDecode para encontrar el árbol de páginas y las fuentes
  $sb = New-Object System.Text.StringBuilder
  $latin = [Text.Encoding]::GetEncoding(28591)
  foreach ($m in [regex]::Matches($raw, '/FlateDecode[^>]*>>\s*stream\r?\n')) {
    $start = $m.Index + $m.Length
    $end = $raw.IndexOf('endstream', $start)
    if ($end -lt 0 -or $end - $start -gt 8MB) { continue }
    try {
      $bytes = $latin.GetBytes($raw.Substring($start + 2, $end - $start - 2))
      $ms = New-Object IO.MemoryStream(, $bytes)
      $ds = New-Object IO.Compression.DeflateStream($ms, [IO.Compression.CompressionMode]::Decompress)
      $sr = New-Object IO.StreamReader($ds, $latin)
      [void]$sb.Append($sr.ReadToEnd())
    } catch { }
  }
  $sb.ToString()
}

function Get-PdfInfo([string]$path) {
  $raw = [Text.Encoding]::GetEncoding(28591).GetString([IO.File]::ReadAllBytes($path))
  $pat = '/Type\s*/Pages\b[^>]*?/Count\s+(\d+)|/Count\s+(\d+)[^>]*?/Type\s*/Pages\b'
  $texto = $raw
  $counts = @([regex]::Matches($texto, $pat) | ForEach-Object { [int]($_.Groups[1].Value + $_.Groups[2].Value) })
  if (-not $counts.Count -or $raw -match '/ObjStm') {
    Add-Type -AssemblyName System.IO.Compression
    $texto = $raw + (Expand-PdfStreams $raw)
    $counts = @([regex]::Matches($texto, $pat) | ForEach-Object { [int]($_.Groups[1].Value + $_.Groups[2].Value) })
  }
  $paginas = if ($counts.Count) { ($counts | Measure-Object -Maximum).Maximum } else { $null }
  # Sin fuentes = solo imágenes = escaneado
  $escaneado = -not ($texto -match '/Font\b')
  [pscustomobject]@{ paginas = $paginas; escaneado = $escaneado }
}

# ---------- Google Drive ----------

function Get-DriveId([string]$valor) {
  if (-not $valor) { return $null }
  $m = [regex]::Match($valor, '/d/([A-Za-z0-9_-]{20,})|[?&]id=([A-Za-z0-9_-]{20,})|^([A-Za-z0-9_-]{20,})$')
  if (-not $m.Success) { throw "No reconozco el enlace o ID de Drive: $valor" }
  $m.Groups[1].Value + $m.Groups[2].Value + $m.Groups[3].Value
}

function Get-Hash([string]$path) {
  if (-not (Test-Path -LiteralPath $path)) { return '' }
  (Get-FileHash -LiteralPath $path -Algorithm SHA256).Hash
}
