Dim WshShell, appFolder
appFolder = CreateObject("Scripting.FileSystemObject").GetParentFolderName(WScript.ScriptFullName)
Set WshShell = CreateObject("WScript.Shell")
WshShell.Run "cmd /c cd /d """ & appFolder & """ && python -m streamlit run app.py", 0, False