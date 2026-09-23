param(
  [string]$CatalogPath = (Join-Path $PSScriptRoot '..\Catálogo - abrir en el navegador.html')
)
$ErrorActionPreference = 'Stop'
$source = [System.IO.File]::ReadAllText((Resolve-Path -LiteralPath $CatalogPath).Path, [System.Text.Encoding]::UTF8)
$template = [System.IO.File]::ReadAllText((Join-Path $PSScriptRoot 'index.template.html'), [System.Text.Encoding]::UTF8)
$books = [regex]::Match($source, '(?m)^const BOOKS = (.*);\r?$')
$categories = [regex]::Match($source, '(?m)^const CATS = (.*);\r?$')
if (-not $books.Success -or -not $categories.Success) {
  throw 'No se encontraron los datos del catálogo en el HTML original.'
}
$output = $template.Replace('__BOOKS__', $books.Groups[1].Value).Replace('__CATS__', $categories.Groups[1].Value)
[System.IO.File]::WriteAllText((Join-Path $PSScriptRoot 'index.html'), $output, [System.Text.UTF8Encoding]::new($false))
Write-Output "Generado index.html"