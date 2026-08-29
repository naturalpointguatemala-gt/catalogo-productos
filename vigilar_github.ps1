$ErrorActionPreference = "Stop"

$RepoRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$Parent = "C:\Users\Natural Point Progra\OneDrive\NATURAL POINT"
$ExcelPath = "C:\Users\Natural Point Progra\OneDrive\PUNTO DE VENTA NATURAL POINT GENERAL.xlsm"
$ExcelFileName = [IO.Path]::GetFileName($ExcelPath)
$SheetName = "Base"
$StartRow = 3
$StockColumn = 15   # O
$PathColumn = 16    # P

$Source = Get-ChildItem $Parent -Directory |
    Where-Object { $_.Name -like "CAT*LOGO DE PRODUCTOS" } |
    Select-Object -First 1

if (-not $Source) {
    Write-Host "ERROR: No se encontro CATÁLOGO DE PRODUCTOS." -ForegroundColor Red
    pause
    exit 1
}

$SourceFolder = $Source.FullName
$ImagesDir = Join-Path $RepoRoot "images"
$Ext = @(".jpg",".jpeg",".png",".gif",".webp",".bmp",".tif",".tiff",".heic")

# ============================================================
# ROT FINDER SIN Microsoft.CSharp.RuntimeBinder / SIN dynamic
# ============================================================
if (-not ("ExcelRotFinder2" -as [type])) {
Add-Type -Language CSharp -TypeDefinition @"
using System;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;

public static class ExcelRotFinder2
{
    [DllImport("ole32.dll")]
    private static extern int GetRunningObjectTable(int reserved, out IRunningObjectTable pprot);

    [DllImport("ole32.dll")]
    private static extern int CreateBindCtx(int reserved, out IBindCtx ppbc);

    private static object GetProp(object obj, string name)
    {
        if (obj == null) return null;
        try
        {
            return obj.GetType().InvokeMember(
                name,
                BindingFlags.GetProperty,
                null,
                obj,
                null
            );
        }
        catch { return null; }
    }

    private static object GetIndexed(object obj, string name, object index)
    {
        if (obj == null) return null;
        try
        {
            return obj.GetType().InvokeMember(
                name,
                BindingFlags.GetProperty,
                null,
                obj,
                new object[] { index }
            );
        }
        catch { return null; }
    }

    private static string AsString(object v)
    {
        try { return Convert.ToString(v); }
        catch { return ""; }
    }

    private static bool IsWantedWorkbook(object obj, string wantedFullName, string wantedFileName)
    {
        if (obj == null) return false;

        string full = AsString(GetProp(obj, "FullName"));
        string name = AsString(GetProp(obj, "Name"));

        if (!String.IsNullOrEmpty(full) &&
            String.Equals(full, wantedFullName, StringComparison.OrdinalIgnoreCase))
            return true;

        if (!String.IsNullOrEmpty(name) &&
            String.Equals(name, wantedFileName, StringComparison.OrdinalIgnoreCase))
            return true;

        return false;
    }

    public static object FindWorkbook(string wantedFullName, string wantedFileName)
    {
        IRunningObjectTable rot = null;
        IEnumMoniker enumMoniker = null;

        try
        {
            GetRunningObjectTable(0, out rot);
            rot.EnumRunning(out enumMoniker);
            enumMoniker.Reset();

            IMoniker[] monikers = new IMoniker[1];

            while (enumMoniker.Next(1, monikers, IntPtr.Zero) == 0)
            {
                IBindCtx ctx = null;
                object obj = null;
                bool keepObj = false;

                try
                {
                    CreateBindCtx(0, out ctx);
                    try { rot.GetObject(monikers[0], out obj); } catch { obj = null; }

                    if (obj == null) continue;

                    // El propio objeto puede ser un Workbook.
                    if (IsWantedWorkbook(obj, wantedFullName, wantedFileName))
                    {
                        keepObj = true;
                        return obj;
                    }

                    // O puede ser Excel.Application. Revisar Workbooks.
                    object books = GetProp(obj, "Workbooks");
                    if (books != null)
                    {
                        int count = 0;
                        try { count = Convert.ToInt32(GetProp(books, "Count")); }
                        catch { count = 0; }

                        for (int i = 1; i <= count; i++)
                        {
                            object wb = GetIndexed(books, "Item", i);
                            if (wb == null) continue;

                            if (IsWantedWorkbook(wb, wantedFullName, wantedFileName))
                            {
                                keepObj = true;
                                return wb;
                            }

                            try { Marshal.ReleaseComObject(wb); } catch {}
                        }

                        try { Marshal.ReleaseComObject(books); } catch {}
                    }
                }
                finally
                {
                    if (ctx != null) {
                        try { Marshal.ReleaseComObject(ctx); } catch {}
                    }

                    if (!keepObj && obj != null) {
                        try { Marshal.ReleaseComObject(obj); } catch {}
                    }

                    if (monikers[0] != null) {
                        try { Marshal.ReleaseComObject(monikers[0]); } catch {}
                    }
                }
            }
        }
        finally
        {
            if (enumMoniker != null) {
                try { Marshal.ReleaseComObject(enumMoniker); } catch {}
            }
            if (rot != null) {
                try { Marshal.ReleaseComObject(rot); } catch {}
            }
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
        if ($i -ge 0) {
            return $p.Substring($i + $marker.Length).TrimStart('\')
        }
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

function Get-LiveWorkbook {
    $wb = $null

    try {
        $wb = [ExcelRotFinder2]::FindWorkbook($ExcelPath, $ExcelFileName)
    }
    catch {
        $wb = $null
    }

    if ($wb) {
        return [PSCustomObject]@{
            Workbook = $wb
            Excel = $null
            OpenedByScript = $false
            Source = "LIBRO ABIERTO EN VIVO"
        }
    }

    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    try { $excel.AutomationSecurity = 3 } catch {}

    $wb = $excel.Workbooks.Open($ExcelPath, 0, $true)

    return [PSCustomObject]@{
        Workbook = $wb
        Excel = $excel
        OpenedByScript = $true
        Source = "COPIA GUARDADA EN DISCO"
    }
}

function Get-ExcelRows {
    $ctx = $null
    $wb = $null
    $excel = $null
    $ws = $null
    $rng = $null

    try {
        $ctx = Get-LiveWorkbook
        $wb = $ctx.Workbook
        $excel = $ctx.Excel

        $ws = $wb.Worksheets.Item($SheetName)
        $lastRow = $ws.Cells($ws.Rows.Count, $PathColumn).End(-4162).Row
        $rows = @()

        if ($lastRow -ge $StartRow) {
            $rng = $ws.Range(
                $ws.Cells.Item($StartRow, $StockColumn),
                $ws.Cells.Item($lastRow, $PathColumn)
            )

            $vals = $rng.Value2

            for ($i = 1; $i -le ($lastRow - $StartRow + 1); $i++) {
                $stock = $vals[$i,1]
                $path = $vals[$i,2]

                if ([string]::IsNullOrWhiteSpace([string]$path)) { continue }

                $rows += [PSCustomObject]@{
                    Row = $StartRow + $i - 1
                    Path = [string]$path
                    Stock = $stock
                }
            }
        }

        return [PSCustomObject]@{
            Rows = $rows
            Source = $ctx.Source
        }
    }
    finally {
        if ($rng) { try { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($rng) } catch {} }

        if ($ctx -and $ctx.OpenedByScript) {
            if ($wb) { try { $wb.Close($false) } catch {} }
            if ($excel) { try { $excel.Quit() } catch {} }
        }

        if ($ws) { try { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($ws) } catch {} }
        if ($wb) { try { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($wb) } catch {} }

        if ($ctx -and $ctx.OpenedByScript -and $excel) {
            try { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($excel) } catch {}
        }

        [GC]::Collect()
        [GC]::WaitForPendingFinalizers()
    }
}

function Get-ExcelSignature {
    $data = Get-ExcelRows
    $parts = @()

    foreach ($r in $data.Rows) {
        $parts += "$($r.Row)|$($r.Stock)|$($r.Path)"
    }

    $sha = [Security.Cryptography.SHA256]::Create()
    try {
        $hash = ([BitConverter]::ToString(
            $sha.ComputeHash([Text.Encoding]::UTF8.GetBytes(($parts -join "`n")))
        ) -replace "-","")

        return [PSCustomObject]@{
            Hash = $hash
            Source = $data.Source
            Count = $data.Rows.Count
        }
    }
    finally { $sha.Dispose() }
}

function Get-ImageSignature {
    $parts = @()

    Get-ChildItem $SourceFolder -File -Recurse -ErrorAction SilentlyContinue |
        Where-Object { $Ext -contains $_.Extension.ToLower() } |
        Sort-Object FullName |
        ForEach-Object {
            $parts += "$($_.FullName)|$($_.Length)|$($_.LastWriteTimeUtc.Ticks)"
        }

    $sha = [Security.Cryptography.SHA256]::Create()
    try {
        return ([BitConverter]::ToString(
            $sha.ComputeHash([Text.Encoding]::UTF8.GetBytes(($parts -join "`n")))
        ) -replace "-","")
    }
    finally { $sha.Dispose() }
}

function Read-StockMap {
    $data = Get-ExcelRows

    Write-Host ("Fuente Excel: " + $data.Source) -ForegroundColor Magenta
    Write-Host ("Registros leidos del Excel: " + $data.Rows.Count) -ForegroundColor Cyan

    $relative = @{}
    $stemBuckets = @{}

    foreach ($row in $data.Rows) {
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
        if ($stemBuckets[$k].Count -eq 1) {
            $uniqueStem[$k] = $stemBuckets[$k][0]
        }
    }

    return @{ Relative=$relative; Stem=$uniqueStem }
}

function Build-Catalog {
    Write-Host ""
    Write-Host "Actualizando catalogo y existencias..." -ForegroundColor Cyan

    $sd = Read-StockMap
    $relMap = $sd.Relative
    $stemMap = $sd.Stem

    if (Test-Path $ImagesDir) { Remove-Item $ImagesDir -Recurse -Force }
    New-Item -ItemType Directory $ImagesDir | Out-Null

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

            New-Item -ItemType Directory $td -Force | Out-Null
            Copy-Item $f.FullName (Join-Path $td $f.Name) -Force

            $stock = $null
            $linked = $false
            $rk = Normalize-Text $rel

            if ($relMap.ContainsKey($rk)) {
                $stock = $relMap[$rk]
                $linked = $true
            }
            else {
                $sk = Stem-Key $f.Name
                if ($stemMap.ContainsKey($sk)) {
                    $stock = $stemMap[$sk]
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
    } | ConvertTo-Json -Depth 6 |
        Set-Content (Join-Path $RepoRoot "catalog.json") -Encoding UTF8

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

        git commit -m ("Actualizar existencias " + (Get-Date -Format "yyyy-MM-dd HH:mm:ss"))
        if ($LASTEXITCODE -ne 0) { throw "No se pudo crear el commit." }

        git push origin main
        if ($LASTEXITCODE -ne 0) { throw "No se pudo publicar en GitHub." }

        Write-Host "PUBLICACION COMPLETADA." -ForegroundColor Green
    }
    finally { Pop-Location }
}

Write-Host ""
Write-Host "CATALOGO PRODUCTOS + EXISTENCIAS (HUELLA O/P - FIX)" -ForegroundColor Cyan
Write-Host "Base: O = existencia | P = ruta | desde fila 3" -ForegroundColor Yellow
Write-Host "Compatible con Windows PowerShell sin RuntimeBinder." -ForegroundColor Yellow
Write-Host "Compara directamente O/P cada 15 segundos." -ForegroundColor Yellow
Write-Host "Publica 60 segundos despues del ultimo cambio." -ForegroundColor Yellow
Write-Host ""

$excelSig = Get-ExcelSignature
$imageSig = Get-ImageSignature

Write-Host ("Fuente inicial Excel: " + $excelSig.Source) -ForegroundColor Magenta
Write-Host ("Filas con ruta detectadas: " + $excelSig.Count) -ForegroundColor Cyan

$pending = $false
$lastChange = Get-Date

while ($true) {
    Start-Sleep -Seconds 15

    try {
        $newExcelSig = Get-ExcelSignature
        $newImageSig = Get-ImageSignature
        $changed = $false

        if ($newExcelSig.Hash -ne $excelSig.Hash) {
            $excelSig = $newExcelSig
            $changed = $true
            Write-Host ("CAMBIO DE EXISTENCIA/RUTA detectado " + (Get-Date -Format "HH:mm:ss")) -ForegroundColor Green
        }

        if ($newImageSig -ne $imageSig) {
            $imageSig = $newImageSig
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
            Build-Catalog
            Publish-GitHub
            Write-Host "Vigilancia activa - productos y existencias..." -ForegroundColor Green
        }
    }
    catch {
        Write-Host ("ERROR DE VIGILANCIA: " + $_.Exception.Message) -ForegroundColor Red
    }
}
