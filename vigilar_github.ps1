$ErrorActionPreference="Stop"

$RepoRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$Parent = "C:\Users\Natural Point Progra\OneDrive\NATURAL POINT"
$ExcelPath = "C:\Users\Natural Point Progra\OneDrive\PUNTO DE VENTA NATURAL POINT GENERAL.xlsm"
$SheetName = "Base"
$StartRow = 3
$StockColumn = 15   # O
$PathColumn = 17    # Q

$Source = Get-ChildItem $Parent -Directory |
  Where-Object { $_.Name -like "CAT*LOGO DE PRODUCTOS" } |
  Select-Object -First 1

if (-not $Source) {
  Write-Host "ERROR: No se encontro CATÁLOGO DE PRODUCTOS." -ForegroundColor Red
  pause
  exit 1
}
if (-not (Test-Path $ExcelPath)) {
  Write-Host "ERROR: No se encontro el Excel." -ForegroundColor Red
  pause
  exit 1
}

$SourceFolder = $Source.FullName
$ImagesDir = Join-Path $RepoRoot "images"
$Ext = @(".jpg",".jpeg",".png",".gif",".webp",".bmp",".tif",".tiff",".heic")

function Normalize-Path([string]$p) {
  if ([string]::IsNullOrWhiteSpace($p)) { return "" }
  $p = $p.Trim().Trim('"').Replace('/','\')
  try { return [IO.Path]::GetFullPath($p).TrimEnd('\').ToLowerInvariant() }
  catch { return $p.TrimEnd('\').ToLowerInvariant() }
}

function Get-ExcelData {
  $excel = $null
  $wb = $null
  $ws = $null
  $startedExcel = $false

  try {
    # Si el libro ya está abierto, intentamos leer ESA misma instancia.
    try {
      $excel = [Runtime.InteropServices.Marshal]::GetActiveObject("Excel.Application")
    } catch {
      $excel = New-Object -ComObject Excel.Application
      $excel.Visible = $false
      $excel.DisplayAlerts = $false
      try { $excel.AutomationSecurity = 3 } catch {}
      $startedExcel = $true
    }

    # Buscar el libro ya abierto por nombre/ruta.
    foreach ($book in @($excel.Workbooks)) {
      try {
        if ($book.FullName -eq $ExcelPath -or $book.Name -eq [IO.Path]::GetFileName($ExcelPath)) {
          $wb = $book
          break
        }
      } catch {}
    }

    # Si no está abierto en esa instancia, abrirlo solo lectura.
    if (-not $wb) {
      $wb = $excel.Workbooks.Open($ExcelPath, 0, $true)
    }

    $ws = $wb.Worksheets.Item($SheetName)
    $lastRow = $ws.Cells($ws.Rows.Count, $PathColumn).End(-4162).Row

    $rows = @()
    if ($lastRow -ge $StartRow) {
      for ($r=$StartRow; $r -le $lastRow; $r++) {
        $rp = $ws.Cells.Item($r, $PathColumn).Value2
        if ([string]::IsNullOrWhiteSpace([string]$rp)) { continue }
        $st = $ws.Cells.Item($r, $StockColumn).Value2
        $rows += [PSCustomObject]@{
          Row = $r
          Path = [string]$rp
          Stock = $st
        }
      }
    }
    return $rows
  }
  finally {
    # Solo cerrar lo que abrió el script; nunca cerrar el Excel del usuario.
    if ($startedExcel) {
      if ($wb) { try { $wb.Close($false) } catch {} }
      if ($excel) { try { $excel.Quit() } catch {} }
    }
    foreach ($o in @($ws)) {
      if ($o) { try { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($o) } catch {} }
    }
    if ($startedExcel) {
      foreach ($o in @($wb,$excel)) {
        if ($o) { try { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($o) } catch {} }
      }
    }
  }
}

function Get-CombinedSignature {
  $parts = @()

  Get-ChildItem $SourceFolder -File -Recurse -ErrorAction SilentlyContinue |
    Where-Object { $Ext -contains $_.Extension.ToLower() } |
    Sort-Object FullName |
    ForEach-Object {
      $parts += "IMG|$($_.FullName)|$($_.Length)|$($_.LastWriteTimeUtc.Ticks)"
    }

  # La clave de esta corrección: la firma incluye directamente O y Q.
  try {
    $excelRows = Get-ExcelData
    foreach ($row in $excelRows) {
      $parts += "XLS|$($row.Row)|$($row.Path)|$($row.Stock)"
    }
  }
  catch {
    Write-Host ("Aviso al leer Excel para detectar cambios: " + $_.Exception.Message) -ForegroundColor DarkYellow
  }

  $sha = [Security.Cryptography.SHA256]::Create()
  try {
    $text = $parts -join "`n"
    return ([BitConverter]::ToString(
      $sha.ComputeHash([Text.Encoding]::UTF8.GetBytes($text))
    ) -replace "-","")
  }
  finally { $sha.Dispose() }
}

function Read-StockMap {
  $rows = Get-ExcelData
  $exact = @{}
  $byName = @{}

  foreach ($row in $rows) {
    $np = Normalize-Path $row.Path
    if (-not $np) { continue }

    $exact[$np] = $row.Stock

    $fn = [IO.Path]::GetFileName($np).ToLowerInvariant()
    if ($fn) {
      if (-not $byName.ContainsKey($fn)) { $byName[$fn] = @() }
      $byName[$fn] += ,([PSCustomObject]@{ Path=$np; Stock=$row.Stock })
    }
  }

  $uniqueByName = @{}
  foreach ($fn in $byName.Keys) {
    if ($byName[$fn].Count -eq 1) {
      $uniqueByName[$fn] = $byName[$fn][0].Stock
    }
  }

  Write-Host ("Registros leidos del Excel: " + $exact.Count) -ForegroundColor Cyan
  return @{ Exact=$exact; Names=$uniqueByName }
}

function Build-Catalog {
  Write-Host ""
  Write-Host "Actualizando catalogo y existencias..." -ForegroundColor Cyan

  $stockData = Read-StockMap
  $exact = $stockData.Exact
  $names = $stockData.Names

  if (Test-Path $ImagesDir) { Remove-Item $ImagesDir -Recurse -Force }
  New-Item -ItemType Directory -Path $ImagesDir | Out-Null

  $out = @()
  $matched = 0
  $unmatched = 0

  Get-ChildItem $SourceFolder -File -Recurse |
    Where-Object { $Ext -contains $_.Extension.ToLower() } |
    ForEach-Object {
      $f = $_
      $rel = $f.FullName.Substring($SourceFolder.TrimEnd('\').Length).TrimStart('\')
      $dir = Split-Path $rel -Parent
      $td = if ($dir) { Join-Path $ImagesDir $dir } else { $ImagesDir }

      New-Item -ItemType Directory -Path $td -Force | Out-Null
      Copy-Item $f.FullName (Join-Path $td $f.Name) -Force

      $stock = $null
      $linked = $false
      $key = Normalize-Path $f.FullName

      if ($exact.ContainsKey($key)) {
        $stock = $exact[$key]
        $linked = $true
      } else {
        $fn = $f.Name.ToLowerInvariant()
        if ($names.ContainsKey($fn)) {
          $stock = $names[$fn]
          $linked = $true
        }
      }

      if ($linked) { $matched++ } else { $unmatched++ }

      $out += [PSCustomObject]@{
        name = $f.BaseName
        folder = if ($dir) { $dir } else { "General" }
        src = "images/" + ($rel -replace "\\","/")
        stock = $stock
        stockLinked = $linked
      }
    }

  [PSCustomObject]@{
    generatedAt = (Get-Date).ToUniversalTime().ToString("o")
    images = $out
  } | ConvertTo-Json -Depth 6 | Set-Content (Join-Path $RepoRoot "catalog.json") -Encoding UTF8

  Write-Host ("Imagenes: " + $out.Count) -ForegroundColor Green
  Write-Host ("Vinculadas con Excel: " + $matched) -ForegroundColor Green
  if ($unmatched -gt 0) {
    Write-Host ("Sin coincidencia: " + $unmatched) -ForegroundColor Yellow
  }
}

function Publish-GitHub {
  Push-Location $RepoRoot
  try {
    git add -A
    git diff --cached --quiet

    if ($LASTEXITCODE -eq 0) {
      Write-Host "Sin cambios para publicar." -ForegroundColor Gray
      return
    }

    git commit -m ("Actualizar catalogo y existencias " + (Get-Date -Format "yyyy-MM-dd HH:mm:ss"))
    if ($LASTEXITCODE -ne 0) { throw "No se pudo crear commit." }

    git push origin main
    if ($LASTEXITCODE -ne 0) { throw "No se pudo publicar en GitHub." }

    Write-Host "PUBLICACION COMPLETADA." -ForegroundColor Green
  }
  finally { Pop-Location }
}

Write-Host ""
Write-Host "CATALOGO PRODUCTOS + EXISTENCIAS (DETECCION DIRECTA EXCEL)" -ForegroundColor Cyan
Write-Host "Base: O = existencia | Q = ruta | desde fila 3" -ForegroundColor Yellow
Write-Host "Revisa directamente O/Q cada 15 segundos." -ForegroundColor Yellow
Write-Host "Publica 60 segundos despues del ultimo cambio." -ForegroundColor Yellow
Write-Host ""

$last = Get-CombinedSignature
$pending = $false
$lastChange = Get-Date

while ($true) {
  Start-Sleep -Seconds 15
  $now = Get-CombinedSignature

  if ($now -ne $last) {
    $last = $now
    $pending = $true
    $lastChange = Get-Date
    Write-Host ("Cambio detectado " + (Get-Date -Format "HH:mm:ss") + ". Esperando 60 segundos...") -ForegroundColor Yellow
  }

  if ($pending -and ((Get-Date)-$lastChange).TotalSeconds -ge 60) {
    $pending = $false
    try {
      Build-Catalog
      Publish-GitHub
      Write-Host "Vigilancia activa - productos y existencias..." -ForegroundColor Green
    }
    catch {
      Write-Host ("ERROR: " + $_.Exception.Message) -ForegroundColor Red
    }
  }
}
