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

if (-not ("ExcelRotFinder3" -as [type])) {
Add-Type -Language CSharp -TypeDefinition @"
using System;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;

public static class ExcelRotFinder3
{
    [DllImport("ole32.dll")]
    private static extern int GetRunningObjectTable(int reserved, out IRunningObjectTable pprot);

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

    private static string S(object v)
    {
        try { return Convert.ToString(v); }
        catch { return ""; }
    }

    public static object FindWorkbook(string wantedFullName, string wantedFileName)
    {
        IRunningObjectTable rot = null;
        IEnumMoniker en = null;

        try
        {
            GetRunningObjectTable(0, out rot);
            rot.EnumRunning(out en);
            en.Reset();

            IMoniker[] m = new IMoniker[1];

            while (en.Next(1, m, IntPtr.Zero) == 0)
            {
                object obj = null;
                bool keep = false;

                try
                {
                    try { rot.GetObject(m[0], out obj); } catch { obj = null; }
                    if (obj == null) continue;

                    string full = S(GetProp(obj, "FullName"));
                    string name = S(GetProp(obj, "Name"));

                    if ((!String.IsNullOrEmpty(full) && String.Equals(full, wantedFullName, StringComparison.OrdinalIgnoreCase)) ||
                        (!String.IsNullOrEmpty(name) && String.Equals(name, wantedFileName, StringComparison.OrdinalIgnoreCase)))
                    {
                        keep = true;
                        return obj;
                    }

                    object books = GetProp(obj, "Workbooks");
                    if (books != null)
                    {
                        int count = 0;
                        try { count = Convert.ToInt32(GetProp(books, "Count")); } catch {}

                        for (int i = 1; i <= count; i++)
                        {
                            object wb = GetIndexed(books, "Item", i);
                            if (wb == null) continue;

                            string wf = S(GetProp(wb, "FullName"));
                            string wn = S(GetProp(wb, "Name"));

                            if ((!String.IsNullOrEmpty(wf) && String.Equals(wf, wantedFullName, StringComparison.OrdinalIgnoreCase)) ||
                                (!String.IsNullOrEmpty(wn) && String.Equals(wn, wantedFileName, StringComparison.OrdinalIgnoreCase)))
                            {
                                keep = true;
                                return wb;
                            }

                            try { Marshal.ReleaseComObject(wb); } catch {}
                        }

                        try { Marshal.ReleaseComObject(books); } catch {}
                    }
                }
                finally
                {
                    if (!keep && obj != null) {
                        try { Marshal.ReleaseComObject(obj); } catch {}
                    }
                    if (m[0] != null) {
                        try { Marshal.ReleaseComObject(m[0]); } catch {}
                    }
                }
            }
        }
        finally
        {
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

# ------------------------------------------------------------
# CONEXION PERSISTENTE AL EXCEL ABIERTO
# ------------------------------------------------------------
$script:LiveWorkbook = $null

function Connect-LiveWorkbook {
    Write-Host "Buscando el Excel abierto..." -ForegroundColor Cyan

    try {
        $script:LiveWorkbook = [ExcelRotFinder3]::FindWorkbook($ExcelPath, $ExcelFileName)
    }
    catch {
        $script:LiveWorkbook = $null
    }

    if (-not $script:LiveWorkbook) {
        throw "No pude conectar con el libro abierto: $ExcelFileName. Verifique que Excel este abierto."
    }

    Write-Host "Fuente Excel: LIBRO ABIERTO EN VIVO" -ForegroundColor Magenta
}

function Ensure-LiveWorkbook {
    if (-not $script:LiveWorkbook) {
        Connect-LiveWorkbook
        return
    }

    try {
        $null = $script:LiveWorkbook.Name
    }
    catch {
        Write-Host "Se perdio la conexion con Excel. Reconectando..." -ForegroundColor Yellow
        $script:LiveWorkbook = $null
        Connect-LiveWorkbook
    }
}

function Get-ExcelRows {
    Ensure-LiveWorkbook

    $ws = $null
    $rng = $null

    try {
        $ws = $script:LiveWorkbook.Worksheets.Item($SheetName)
        if (-not $ws) { throw "No se encontro la hoja Base." }

        $lastRow = $ws.Cells($ws.Rows.Count, $PathColumn).End(-4162).Row
        $rows = @()

        if ($lastRow -lt $StartRow) {
            return $rows
        }

        $rng = $ws.Range(
            $ws.Cells.Item($StartRow, $StockColumn),
            $ws.Cells.Item($lastRow, $PathColumn)
        )

        $vals = $rng.Value2
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

        return $rows
    }
    finally {
        if ($rng) {
            try { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($rng) } catch {}
        }
        if ($ws) {
            try { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($ws) } catch {}
        }
    }
}

function Get-ExcelSignature {
    $rows = Get-ExcelRows
    $parts = @()

    foreach ($r in $rows) {
        $parts += "$($r.Row)|$($r.Stock)|$($r.Path)"
    }

    $sha = [Security.Cryptography.SHA256]::Create()
    try {
        return ([BitConverter]::ToString(
            $sha.ComputeHash([Text.Encoding]::UTF8.GetBytes(($parts -join "`n")))
        ) -replace "-","")
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
    $rows = Get-ExcelRows
    Write-Host ("Registros leidos del Excel: " + $rows.Count) -ForegroundColor Cyan

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
Write-Host "CATALOGO PRODUCTOS + EXISTENCIAS (CONEXION PERSISTENTE)" -ForegroundColor Cyan
Write-Host "Base: O = existencia | P = ruta | desde fila 3" -ForegroundColor Yellow
Write-Host "Lee directamente el Excel abierto cada 15 segundos." -ForegroundColor Yellow
Write-Host "Publica 60 segundos despues del ultimo cambio." -ForegroundColor Yellow
Write-Host ""

Connect-LiveWorkbook

$excelRows = Get-ExcelRows
Write-Host ("Filas con ruta detectadas: " + $excelRows.Count) -ForegroundColor Cyan

$excelSig = Get-ExcelSignature
$imageSig = Get-ImageSignature
$pending = $false
$lastChange = Get-Date

while ($true) {
    Start-Sleep -Seconds 15

    try {
        $newExcelSig = Get-ExcelSignature
        $newImageSig = Get-ImageSignature
        $changed = $false

        if ($newExcelSig -ne $excelSig) {
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
        Write-Host ("Linea aproximada: " + $_.InvocationInfo.ScriptLineNumber) -ForegroundColor DarkYellow
    }
}
