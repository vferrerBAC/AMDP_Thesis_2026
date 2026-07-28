@echo off
echo Setting up IJET App shortcut...

powershell -NoProfile -ExecutionPolicy Bypass -Command "$appFolder = '%~dp0'.TrimEnd('\'); $desktop = [Environment]::GetFolderPath('Desktop'); $WshShell = New-Object -comObject WScript.Shell; $Shortcut = $WshShell.CreateShortcut((Join-Path $desktop 'IJET App.lnk')); $Shortcut.TargetPath = (Join-Path $appFolder 'launch.vbs'); $Shortcut.WorkingDirectory = $appFolder; $Shortcut.IconLocation = (Join-Path $appFolder 'IJET_logo.ico'); $Shortcut.Description = 'Launch IJET Streamlit App'; $Shortcut.Save(); Write-Host 'Shortcut created on Desktop.'"

echo.
echo Done! You should now see "IJET App" on your Desktop.
pause
