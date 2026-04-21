' Interfaz sin consola negra. Usa el venv del proyecto (donde están selenium, etc.).
Option Explicit
Dim sh, fso, dir, py, venvPyw
Set sh = CreateObject("WScript.Shell")
Set fso = CreateObject("Scripting.FileSystemObject")
dir = fso.GetParentFolderName(WScript.ScriptFullName)
py = dir & "\facturador_ui.py"
venvPyw = dir & "\.venv\Scripts\pythonw.exe"
sh.CurrentDirectory = dir

If fso.FileExists(venvPyw) Then
  ' 1 = ventana normal; False = no bloquear
  sh.Run """" & venvPyw & """ """ & py & """", 1, False
Else
  sh.Run "pyw -3 """ & py & """", 1, False
End If
