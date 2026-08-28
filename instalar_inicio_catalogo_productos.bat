@echo off
title Instalar inicio automatico - Catalogo Productos
set "TARGET=%~dp0iniciar_catalogo.bat"
set "STARTUP=%APPDATA%\Microsoft\Windows\Start Menu\Programs\Startup"

if not exist "%TARGET%" (
  echo.
  echo ERROR: Este archivo debe estar dentro de la misma carpeta que iniciar_catalogo.bat
  echo.
  pause
  exit /b 1
)

powershell.exe -NoProfile -ExecutionPolicy Bypass -Command "$ws=New-Object -ComObject WScript.Shell; $s=$ws.CreateShortcut('%STARTUP%\Catalogo Productos GitHub.lnk'); $s.TargetPath='%TARGET%'; $s.WorkingDirectory='%~dp0'; $s.WindowStyle=1; $s.Save()"

if errorlevel 1 (
  echo.
  echo No se pudo crear el acceso directo.
  pause
  exit /b 1
)

echo.
echo ================================================
echo   INICIO AUTOMATICO INSTALADO CORRECTAMENTE
echo ================================================
echo.
echo El catalogo se iniciara automaticamente al entrar a Windows.
echo.
pause
