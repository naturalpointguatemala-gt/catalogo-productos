$ErrorActionPreference = "Stop"

$RepoRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$Parent = "C:\Users\Natural Point Progra\OneDrive\NATURAL POINT"
$ExcelPath = "C:\Users\Natural Point Progra\OneDrive\PUNTO DE VENTA NATURAL POINT GENERAL.xlsm"
$ExcelFileName = [IO.Path]::GetFileName($ExcelPath)
$SheetName = "Base"
$StartRow = 3
$StockColumn = 15
$PathColumn = 16

$Source = Get-ChildItem $Parent -Directory | Where-Object { $_.Name -like "CAT*LOGO DE PRODUCTOS" } | Select-Object -First 1
if (-not $Source) { Write-Host "ERROR: No se encontro CATÁLOGO DE PRODUCTOS." -ForegroundColor Red; pause; exit 1 }

$SourceFolder = $Source.FullName
$ImagesDir = Join-Path $RepoRoot "images"
$Ext = @(".jpg",".jpeg",".png",".gif",".webp",".bmp",".tif",".tiff",".heic")

if (-not ("ExcelRotSnapshotFinder" -as [type])) {
Add-Type -Language CSharp -TypeDefinition @"
using System;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;

public static class ExcelRotSnapshotFinder
{
    [DllImport("ole32.dll")]
    private static extern int GetRunningObjectTable(int reserved, out IRunningObjectTable pprot);

    private static object GetProp(object obj, string name)
    {
        if (obj == null) return null;
        try { return obj.GetType().InvokeMember(name, BindingFlags.GetProperty, null, obj, null); }
        catch { return null; }
    }

    private static object GetIndexed(object obj, string name, object index)
    {
        if (obj == null) return null;
        try { return obj.GetType().InvokeMember(name, BindingFlags.GetProperty, null, obj, new object[] { index }); }
        catch { return null; }
    }

    private static string S(object v)
    {
        try { return Convert.ToString(v); }
        catch { return ""; }
    }

    public static object FindWorkbook(string wantedFullName, string wantedFileName)
    {
        IRunningObjectTable rot = null;
        IEnumMoniker en = null;

        try {
            GetRunningObjectTable(0, out rot);
            rot.EnumRunning(out en);
            en.Reset();

            IMoniker[] m = new IMoniker[1];

            while (en.Next(1, m, IntPtr.Zero) == 0) {
                object obj = null;
                bool keep = false;

                try {
                    try { rot.GetObject(m[0], out obj); } catch { obj = null; }
                    if (obj == null) continue;

                    string full = S(GetProp(obj, "FullName"));
                    string name = S(GetProp(obj, "Name"));

                    if ((!String.IsNullOrEmpty(full) && String.Equals(full, wantedFullName, StringComparison.OrdinalIgnoreCase)) ||
                        (!String.IsNullOrEmpty(name) && String.Equals(name, wantedFileName, StringComparison.OrdinalIgnoreCase))) {
                        keep = true;
                        return obj;
                    }

                    object books = GetProp(obj, "Workbooks");
                    if (books != null) {
                        int count = 0;
                        try { count = Convert.ToInt32(GetProp(books, "Count")); } catch {}

                        for (int i = 1; i <= count; i++) {
                            object wb = GetIndexed(books, "Item", i);
                            if (wb == null) continue;

                            string wf = S(GetProp(wb, "FullName"));
                            string wn = S(GetProp(wb, "Name"));

                            if ((!String.IsNullOrEmpty(wf) && String.Equals(wf, wantedFullName, StringComparison.OrdinalIgnoreCase)) ||
                                (!String.IsNullOrEmpty(wn) && String.Equals(wn, wantedFileName, StringComparison.OrdinalIgnoreCase))) {
                                keep = true;
                                return wb;
                            }

                            try { Marshal.ReleaseComObject(wb); } catch {}
                        }

                        try { Marshal.ReleaseComObject(books); } catch {}
                    }
                }
                finally {
                    if (!keep && obj != null) { try { Marshal.ReleaseComObject(obj); } catch {} }
                    if (m[0] != null) { try { Marshal.ReleaseComObject(m[0]); } catch {} }
                }
            }
        }
        finally {
            if (en != null) { try { Marshal.ReleaseComObject(en); } catch {} }
            if (rot != null) { try { Marshal.ReleaseComObject(rot); } catch {} }
        }

        return null;
    }
}
"@
}

function Normalize-Text([string]$s) {
    if ([string]::IsNullOrWhiteSpace($s)) { return "" }
    return (($s.Trim().Trim('"').Replace('/','\').ToLowerInvariant()) -replace '\s+',' ')
}

function Relative-Catalog-Key([string]$p) {
    $p = Normalize-Text $p
    if (-not $p) { return "" }
    foreach ($marker in @("\catálogo de productos\", "\catalogo de productos\")) {
        $i = $p.IndexOf($marker)
        if ($i -ge 0) { return $p.Substring($i + $marker.Length).TrimStart('\') }
    }
    return [IO.Path]::GetFileName($p)
}

function Stem-Key([string]$p) {
    $name = [IO.Path]::GetFileNameWithoutExtension((Normalize-Text $p))
    if (-not $name) { return "" }
    $name = $name -replace '[_\-]+',' '
    $name = $name -replace '\s+',' '
    return $name.Trim()
}

$script:LiveWorkbook = $null

function Connect-LiveWorkbook {
    $script:LiveWorkbook = $null
    try { $script:LiveWorkbook = [ExcelRotSnapshotFinder]::FindWorkbook($ExcelPath, $ExcelFileName) } catch { $script:LiveWorkbook = $null }
    if (-not $script:LiveWorkbook) { throw "No se encontro el libro abierto. NO se usara una copia guardada." }
    Write-Host "Excel conectado: LIBRO ABIERTO EN VIVO" -ForegroundColor Magenta
}

function Get-LiveRows {
    for ($attempt = 1; $attempt -le 3; $attempt++) {
        $ws = $null
        $rng = $null
        try {
            if (-not $script:LiveWorkbook) { Connect-LiveWorkbook }

            $ws = $script:LiveWorkbook.Worksheets.Item($SheetName)
            if (-not $ws) { throw "Hoja Base no disponible." }

            $lastRow = $ws.Cells($ws.Rows.Count, $PathColumn).End(-4162).Row
            if ($lastRow -lt $StartRow) { return @() }

            # IMPORTANTE: Base!O contiene formulas. Forzar el recalculo de O antes de leer.
            $calcRange = $null
            $app = $null
            try {
                $app = $script:LiveWorkbook.Application
                $calcRange = $ws.Range($ws.Cells.Item($StartRow, $StockColumn), $ws.Cells.Item($lastRow, $StockColumn))
                $calcRange.Calculate()

                # Esperar a que Excel termine de calcular (xlDone = 0).
                $limit = (Get-Date).AddSeconds(10)
                while ((Get-Date) -lt $limit) {
                    try { if ($app.CalculationState -eq 0) { break } } catch { break }
                    Start-Sleep -Milliseconds 100
                }
                Start-Sleep -Milliseconds 250
            } finally {
                if ($calcRange) { try { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($calcRange) } catch {} }
                if ($app) { try { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($app) } catch {} }
            }

            $rng = $ws.Range($ws.Cells.Item($StartRow, $StockColumn), $ws.Cells.Item($lastRow, $PathColumn))
            $vals = $rng.Value2
            if ($null -eq $vals) { throw "Excel devolvio un rango vacio." }

            $rows = @()
            $count = $lastRow - $StartRow + 1

            for ($i = 1; $i -le $count; $i++) {
                $stock = $vals[$i,1]
                $path = $vals[$i,2]
                if ([string]::IsNullOrWhiteSpace([string]$path)) { continue }

                $rows += [PSCustomObject]@{
                    Row = $StartRow + $i - 1
                    Path = [string]$path
                    Stock = $stock
                }
            }
            return @($rows)
        }
        catch {
            if ($attempt -eq 3) { throw }
            Write-Host ("Lectura Excel fallo; reconectando (intento " + $attempt + "/3)...") -ForegroundColor DarkYellow
            $script:LiveWorkbook = $null
            Start-Sleep -Milliseconds 700
            try { Connect-LiveWorkbook } catch {}
        }
        finally {
            if ($rng) { try { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($rng) } catch {} }
            if ($ws)  { try { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($ws) } catch {} }
        }
    }
}

function Get-RowsState($rows) {
    $parts = New-Object System.Collections.Generic.List[string]
    foreach ($r in @($rows)) {
        $stockText = if ($null -eq $r.Stock) { "" } else { [string]$r.Stock }
        $pathText  = if ($null -eq $r.Path)  { "" } else { [string]$r.Path }
        [void]$parts.Add(("{0}|{1}|{2}" -f $r.Row, $stockText, $pathText))
    }
    return [string]::Join("`n", $parts.ToArray())
}

function Get-ImageState {
    $parts = New-Object System.Collections.Generic.List[string]
    $files = @(Get-ChildItem $SourceFolder -File -Recurse -ErrorAction SilentlyContinue |
        Where-Object { $Ext -contains $_.Extension.ToLowerInvariant() } |
        Sort-Object FullName)
    foreach ($f in $files) {
        [void]$parts.Add(("{0}|{1}|{2}" -f $f.FullName, $f.Length, $f.LastWriteTimeUtc.Ticks))
    }
    return [string]::Join("`n", $parts.ToArray())
}

function Build-CatalogFromSnapshot($rows) {
    Write-Host ""
    Write-Host "Construyendo catalogo desde SNAPSHOT de existencias..." -ForegroundColor Cyan

    $rows = @($rows)
    if ($rows.Count -eq 0) { throw "El snapshot de Excel esta vacio. Se cancela para proteger existencias." }

    Write-Host ("Filas de stock en snapshot: " + $rows.Count) -ForegroundColor Cyan

    $relative = @{}
    $stemBuckets = @{}

    foreach ($row in $rows) {
        $rk = Relative-Catalog-Key $row.Path
        if ($rk) { $relative[$rk] = $row.Stock }

        $sk = Stem-Key $row.Path
        if ($sk) {
            if (-not $stemBuckets.ContainsKey($sk)) { $stemBuckets[$sk] = @() }
            $stemBuckets[$sk] += ,$row.Stock
        }
    }

    $uniqueStem = @{}
    foreach ($k in $stemBuckets.Keys) {
        if ($stemBuckets[$k].Count -eq 1) { $uniqueStem[$k] = $stemBuckets[$k][0] }
    }

    if (Test-Path $ImagesDir) { Remove-Item $ImagesDir -Recurse -Force }
    New-Item -ItemType Directory $ImagesDir | Out-Null

    $out = @()
    $matched = 0
    $unmatched = 0

    Get-ChildItem $SourceFolder -File -Recurse |
        Where-Object { $Ext -contains $_.Extension.ToLowerInvariant() } |
        ForEach-Object {
            $f = $_
            $rel = $f.FullName.Substring($SourceFolder.TrimEnd('\').Length).TrimStart('\')
            $dir = Split-Path $rel -Parent
            $td = if ($dir) { Join-Path $ImagesDir $dir } else { $ImagesDir }

            New-Item -ItemType Directory $td -Force | Out-Null
            Copy-Item $f.FullName (Join-Path $td $f.Name) -Force

            $stock = $null
            $linked = $false
            $rk = Normalize-Text $rel

            if ($relative.ContainsKey($rk)) {
                $stock = $relative[$rk]
                $linked = $true
            } else {
                $sk = Stem-Key $f.Name
                if ($uniqueStem.ContainsKey($sk)) {
                    $stock = $uniqueStem[$sk]
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
    if ($unmatched -gt 0) { Write-Host ("Sin coincidencia: " + $unmatched) -ForegroundColor Yellow }
}

function Publish-GitHub {
    Push-Location $RepoRoot
    try {
        git add -A
        git diff --cached --quiet
        if ($LASTEXITCODE -eq 0) { Write-Host "Sin cambios para publicar." -ForegroundColor Gray; return }

        git commit -m ("Actualizar existencias snapshot " + (Get-Date -Format "yyyy-MM-dd HH:mm:ss"))
        if ($LASTEXITCODE -ne 0) { throw "No se pudo crear el commit." }

        git push origin main
        if ($LASTEXITCODE -ne 0) { throw "No se pudo publicar en GitHub." }

        Write-Host "PUBLICACION COMPLETADA." -ForegroundColor Green
    }
    finally { Pop-Location }
}

Write-Host ""
Write-Host "CATALOGO PRODUCTOS + EXISTENCIAS (BASE RECALCULADA)" -ForegroundColor Cyan
Write-Host "Fuente unica: Base!O = existencia | Base!P = ruta | desde fila 3" -ForegroundColor Yellow
Write-Host "Antes de cada lectura se recalcula Base!O para obtener el resultado actual de sus formulas." -ForegroundColor Yellow
Write-Host "NO lee existencias directamente de la hoja Existencia." -ForegroundColor Yellow
Write-Host ""

Connect-LiveWorkbook

$initialRows = @(Get-LiveRows)
if ($initialRows.Count -eq 0) { throw "No se pudieron leer las existencias iniciales." }

Write-Host ("Filas con ruta detectadas: " + $initialRows.Count) -ForegroundColor Cyan
Write-Host "Vigilancia activa: recalculando Base!O cada 15 segundos..." -ForegroundColor Green

$excelState = Get-RowsState $initialRows
$imageState = Get-ImageState

$script:LastGoodRows = @($initialRows)
$script:PendingRows = @($initialRows)

$pending = $false
$lastChange = Get-Date

while ($true) {
    Start-Sleep -Seconds 15

    try {
        $newRows = @(Get-LiveRows)
        if ($newRows.Count -eq 0) {
            Write-Host "Lectura vacia ignorada: se conserva el ultimo snapshot bueno." -ForegroundColor DarkYellow
            continue
        }

        $newExcelState = Get-RowsState $newRows
        $newImageState = Get-ImageState
        $changed = $false

        $script:LastGoodRows = @($newRows)

        if ($newExcelState -ne $excelState) {
            $excelState = $newExcelState
            $script:PendingRows = @($newRows)
            $changed = $true
            Write-Host ("CAMBIO DE EXISTENCIA/RUTA detectado " + (Get-Date -Format "HH:mm:ss")) -ForegroundColor Green
            Write-Host ("Snapshot guardado: " + $script:PendingRows.Count + " filas") -ForegroundColor Cyan
        }

        if ($newImageState -ne $imageState) {
            $imageState = $newImageState
            $script:PendingRows = @($script:LastGoodRows)
            $changed = $true
            Write-Host ("CAMBIO DE IMAGEN detectado " + (Get-Date -Format "HH:mm:ss")) -ForegroundColor Green
        }

        if ($changed) {
            $pending = $true
            $lastChange = Get-Date
            Write-Host "Esperando 60 segundos sin nuevos cambios..." -ForegroundColor Yellow
        }

        if ($pending -and ((Get-Date) - $lastChange).TotalSeconds -ge 60) {
            $pending = $false

            $rowsToPublish = @($script:PendingRows)
            if ($rowsToPublish.Count -eq 0) {
                Write-Host "PUBLICACION CANCELADA: snapshot vacio." -ForegroundColor Red
                continue
            }

            Build-CatalogFromSnapshot $rowsToPublish
            Publish-GitHub

            Write-Host "Vigilancia activa - esperando cambios..." -ForegroundColor Green
        }
    }
    catch {
        Write-Host ("LECTURA EXCEL OMITIDA: " + $_.Exception.Message) -ForegroundColor Red
        Write-Host "No se publicara nada hasta obtener otra lectura valida en vivo." -ForegroundColor DarkYellow

        $script:LiveWorkbook = $null
        try { Connect-LiveWorkbook } catch { Write-Host "Esperando a poder reconectar con Excel..." -ForegroundColor DarkYellow }
    }
}
