$ErrorActionPreference="Stop"

$RepoRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$Parent = "C:\Users\Natural Point Progra\OneDrive\NATURAL POINT"
$ExcelPath = "C:\Users\Natural Point Progra\OneDrive\PUNTO DE VENTA NATURAL POINT GENERAL.xlsm"
$SheetName = "Base"
$StartRow = 3
$StockColumn = 15
$PathColumn = 17

$Source = Get-ChildItem $Parent -Directory |
  Where-Object { $_.Name -like "CAT*LOGO DE PRODUCTOS" } |
  Select-Object -First 1

if (-not $Source) { Write-Host "ERROR: No se encontro CATÁLOGO DE PRODUCTOS." -ForegroundColor Red; pause; exit 1 }
if (-not (Test-Path $ExcelPath)) { Write-Host "ERROR: No se encontro el Excel." -ForegroundColor Red; pause; exit 1 }

$SourceFolder = $Source.FullName
$ImagesDir = Join-Path $RepoRoot "images"
$Ext = @(".jpg",".jpeg",".png",".gif",".webp",".bmp",".tif",".tiff",".heic")

function Normalize-Text([string]$s) {
  if ([string]::IsNullOrWhiteSpace($s)) { return "" }
  $s = $s.Trim().Trim('"').Replace('/','\').ToLowerInvariant()
  $s = $s -replace '\s+',' '
  return $s
}

function Relative-Catalog-Key([string]$p) {
  $p = Normalize-Text $p
  if (-not $p) { return "" }

  # Conserva solo lo que viene después de "CATÁLOGO DE PRODUCTOS\"
  # para ignorar C:\Users\Administrador vs C:\Users\Natural Point Progra
  $markers = @(
    "\catálogo de productos\",
    "\catalogo de productos\"
  )
  foreach ($m in $markers) {
    $i = $p.IndexOf($m)
    if ($i -ge 0) {
      return $p.Substring($i + $m.Length).TrimStart('\')
    }
  }
  return [IO.Path]::GetFileName($p)
}

function Stem-Key([string]$p) {
  $name = [IO.Path]::GetFileNameWithoutExtension((Normalize-Text $p))
  if (-not $name) { return "" }
  # Normaliza pequeños detalles de espacios/guiones para una coincidencia más tolerante
  $name = $name -replace '[_\-]+',' '
  $name = $name -replace '\s+',' '
  return $name.Trim()
}

function Get-ExcelData {
  $excel=$null;$wb=$null;$ws=$null;$started=$false
  try {
    try { $excel=[Runtime.InteropServices.Marshal]::GetActiveObject("Excel.Application") }
    catch {
      $excel=New-Object -ComObject Excel.Application
      $excel.Visible=$false
      $excel.DisplayAlerts=$false
      try{$excel.AutomationSecurity=3}catch{}
      $started=$true
    }

    foreach($book in @($excel.Workbooks)) {
      try {
        if($book.FullName -eq $ExcelPath -or $book.Name -eq [IO.Path]::GetFileName($ExcelPath)) {
          $wb=$book; break
        }
      } catch {}
    }
    if(-not $wb){$wb=$excel.Workbooks.Open($ExcelPath,0,$true)}

    $ws=$wb.Worksheets.Item($SheetName)
    $last=$ws.Cells($ws.Rows.Count,$PathColumn).End(-4162).Row
    $rows=@()

    for($r=$StartRow;$r-le$last;$r++){
      $rp=$ws.Cells.Item($r,$PathColumn).Value2
      if([string]::IsNullOrWhiteSpace([string]$rp)){continue}
      $st=$ws.Cells.Item($r,$StockColumn).Value2
      $rows += [PSCustomObject]@{Row=$r;Path=[string]$rp;Stock=$st}
    }
    return $rows
  }
  finally {
    if($started){
      if($wb){try{$wb.Close($false)}catch{}}
      if($excel){try{$excel.Quit()}catch{}}
    }
    if($ws){try{[void][Runtime.InteropServices.Marshal]::ReleaseComObject($ws)}catch{}}
    if($started){
      foreach($o in @($wb,$excel)){if($o){try{[void][Runtime.InteropServices.Marshal]::ReleaseComObject($o)}catch{}}}
    }
  }
}

function Get-CombinedSignature {
  $parts=@()
  Get-ChildItem $SourceFolder -File -Recurse -ErrorAction SilentlyContinue |
    Where-Object {$Ext -contains $_.Extension.ToLower()} |
    Sort-Object FullName |
    ForEach-Object {$parts += "IMG|$($_.FullName)|$($_.Length)|$($_.LastWriteTimeUtc.Ticks)"}

  try {
    foreach($row in (Get-ExcelData)) {
      $parts += "XLS|$($row.Row)|$($row.Path)|$($row.Stock)"
    }
  } catch {}

  $sha=[Security.Cryptography.SHA256]::Create()
  try {
    return([BitConverter]::ToString($sha.ComputeHash([Text.Encoding]::UTF8.GetBytes(($parts-join "`n"))))-replace "-","")
  } finally {$sha.Dispose()}
}

function Read-StockMap {
  $rows=Get-ExcelData
  $relative=@{}
  $stemBuckets=@{}

  foreach($row in $rows){
    $rk=Relative-Catalog-Key $row.Path
    if($rk){ $relative[$rk]=$row.Stock }

    $sk=Stem-Key $row.Path
    if($sk){
      if(-not $stemBuckets.ContainsKey($sk)){ $stemBuckets[$sk]=@() }
      $stemBuckets[$sk] += ,([PSCustomObject]@{Path=$row.Path;Stock=$row.Stock})
    }
  }

  # Solo usar coincidencia por nombre base si ese nombre aparece una sola vez en Excel.
  $uniqueStem=@{}
  foreach($k in $stemBuckets.Keys){
    if($stemBuckets[$k].Count -eq 1){
      $uniqueStem[$k]=$stemBuckets[$k][0].Stock
    }
  }

  Write-Host ("Registros leidos del Excel: "+$rows.Count) -ForegroundColor Cyan
  return @{Relative=$relative;Stem=$uniqueStem}
}

function Build-Catalog {
  Write-Host ""
  Write-Host "Actualizando catalogo y existencias..." -ForegroundColor Cyan
  $sd=Read-StockMap
  $relMap=$sd.Relative
  $stemMap=$sd.Stem

  if(Test-Path $ImagesDir){Remove-Item $ImagesDir -Recurse -Force}
  New-Item -ItemType Directory $ImagesDir|Out-Null

  $out=@();$matched=0;$unmatched=0

  Get-ChildItem $SourceFolder -File -Recurse |
    Where-Object {$Ext -contains $_.Extension.ToLower()} |
    ForEach-Object {
      $f=$_
      $rel=$f.FullName.Substring($SourceFolder.TrimEnd('\').Length).TrimStart('\')
      $dir=Split-Path $rel -Parent
      $td=if($dir){Join-Path $ImagesDir $dir}else{$ImagesDir}
      New-Item -ItemType Directory $td -Force|Out-Null
      Copy-Item $f.FullName (Join-Path $td $f.Name) -Force

      $stock=$null;$linked=$false;$method=""
      $rk=Normalize-Text $rel

      if($relMap.ContainsKey($rk)){
        $stock=$relMap[$rk];$linked=$true;$method="ruta relativa"
      } else {
        $sk=Stem-Key $f.Name
        if($stemMap.ContainsKey($sk)){
          $stock=$stemMap[$sk];$linked=$true;$method="nombre sin extension"
        }
      }

      if($linked){$matched++}else{$unmatched++}

      $out += [PSCustomObject]@{
        name=$f.BaseName
        folder=if($dir){$dir}else{"General"}
        src="images/"+($rel-replace "\\","/")
        stock=$stock
        stockLinked=$linked
        stockMatch=$method
      }
    }

  [PSCustomObject]@{
    generatedAt=(Get-Date).ToUniversalTime().ToString("o")
    images=$out
  } | ConvertTo-Json -Depth 6 | Set-Content (Join-Path $RepoRoot "catalog.json") -Encoding UTF8

  Write-Host ("Imagenes: "+$out.Count) -ForegroundColor Green
  Write-Host ("Vinculadas con Excel: "+$matched) -ForegroundColor Green
  if($unmatched -gt 0){Write-Host ("Sin coincidencia: "+$unmatched) -ForegroundColor Yellow}
}

function Publish-GitHub {
  Push-Location $RepoRoot
  try {
    git add -A
    git diff --cached --quiet
    if($LASTEXITCODE -eq 0){Write-Host "Sin cambios para publicar.";return}

    git commit -m ("Mejorar vinculacion de existencias "+(Get-Date -Format "yyyy-MM-dd HH:mm:ss"))
    if($LASTEXITCODE -ne 0){throw "No se pudo crear commit."}

    git push origin main
    if($LASTEXITCODE -ne 0){throw "No se pudo publicar en GitHub."}
    Write-Host "PUBLICACION COMPLETADA." -ForegroundColor Green
  } finally {Pop-Location}
}

Write-Host ""
Write-Host "CATALOGO PRODUCTOS + EXISTENCIAS (RUTAS FLEXIBLES)" -ForegroundColor Cyan
Write-Host "Ignora el usuario de Windows y la extension de imagen cuando es necesario." -ForegroundColor Yellow
Write-Host "Base: O = existencia | Q = ruta | desde fila 3" -ForegroundColor Yellow
Write-Host "Revisa Excel cada 15 segundos y publica 60 segundos despues del ultimo cambio." -ForegroundColor Yellow
Write-Host ""

$last=Get-CombinedSignature
$pending=$false
$lastChange=Get-Date

while($true){
  Start-Sleep 15
  $now=Get-CombinedSignature

  if($now-ne$last){
    $last=$now;$pending=$true;$lastChange=Get-Date
    Write-Host ("Cambio detectado "+(Get-Date -Format "HH:mm:ss")+". Esperando 60 segundos...") -ForegroundColor Yellow
  }

  if($pending -and ((Get-Date)-$lastChange).TotalSeconds -ge 60){
    $pending=$false
    try {
      Build-Catalog
      Publish-GitHub
      Write-Host "Vigilancia activa - productos y existencias..." -ForegroundColor Green
    } catch {
      Write-Host ("ERROR: "+$_.Exception.Message) -ForegroundColor Red
    }
  }
}
