$ErrorActionPreference="Stop"
$RepoRoot=Split-Path -Parent $MyInvocation.MyCommand.Path
$Parent="C:\Users\Natural Point Progra\OneDrive\NATURAL POINT"
$Source=Get-ChildItem $Parent -Directory | Where-Object {$_.Name -like "CAT*LOGO DE PRODUCTOS"} | Select-Object -First 1
if(-not $Source){Write-Host "No se encontro la carpeta del catalogo" -ForegroundColor Red; pause; exit 1}
$SourceFolder=$Source.FullName
$ImagesDir=Join-Path $RepoRoot "images"
$Ext=@(".jpg",".jpeg",".png",".gif",".webp",".bmp",".tif",".tiff",".heic")
function Sig{
 $f=Get-ChildItem $SourceFolder -File -Recurse | Where-Object {$Ext -contains $_.Extension.ToLower()} | Sort-Object FullName
 $t=($f|ForEach-Object{"$($_.FullName)|$($_.Length)|$($_.LastWriteTimeUtc.Ticks)"})-join "`n"
 $s=[Security.Cryptography.SHA256]::Create()
 try{return ([BitConverter]::ToString($s.ComputeHash([Text.Encoding]::UTF8.GetBytes($t)))-replace "-","")}finally{$s.Dispose()}
}
function Build{
 if(Test-Path $ImagesDir){Remove-Item $ImagesDir -Recurse -Force}
 New-Item -ItemType Directory $ImagesDir|Out-Null
 $out=@()
 Get-ChildItem $SourceFolder -File -Recurse | Where-Object {$Ext -contains $_.Extension.ToLower()} | ForEach-Object{
   $f=$_;$rel=$f.FullName.Substring($SourceFolder.TrimEnd('\').Length).TrimStart('\');$dir=Split-Path $rel -Parent
   $td=if($dir){Join-Path $ImagesDir $dir}else{$ImagesDir}
   New-Item -ItemType Directory $td -Force|Out-Null
   Copy-Item $f.FullName (Join-Path $td $f.Name) -Force
   $out += [PSCustomObject]@{name=$f.BaseName;folder=if($dir){$dir}else{"General"};src="images/"+($rel-replace "\\","/")}
 }
 [PSCustomObject]@{generatedAt=(Get-Date).ToUniversalTime().ToString("o");images=$out}|ConvertTo-Json -Depth 5|Set-Content (Join-Path $RepoRoot "catalog.json") -Encoding UTF8
 Write-Host ("Catalogo actualizado. Imagenes: "+$out.Count) -ForegroundColor Green
}
function Publish{
 Push-Location $RepoRoot
 try{
  & git add -A
  & git diff --cached --quiet
  if($LASTEXITCODE -eq 0){Write-Host "Sin cambios para publicar.";return}
  & git commit -m ("Actualizar catalogo "+(Get-Date -Format "yyyy-MM-dd HH:mm:ss"))
  & git push origin main
  if($LASTEXITCODE -eq 0){Write-Host "PUBLICACION COMPLETADA." -ForegroundColor Green}
 }finally{Pop-Location}
}
Write-Host ("Vigilando: "+$SourceFolder) -ForegroundColor Green
Write-Host "Se publicara 60 segundos despues del ultimo cambio." -ForegroundColor Yellow
$last=Sig;$pending=$false;$lastChange=Get-Date
while($true){
 Start-Sleep 5
 $n=Sig
 if($n-ne$last){$last=$n;$pending=$true;$lastChange=Get-Date;Write-Host "Cambio detectado. Esperando 60 segundos..."}
 if($pending -and ((Get-Date)-$lastChange).TotalSeconds -ge 60){$pending=$false;Build;Publish;Write-Host "Vigilancia activa..."}
}
