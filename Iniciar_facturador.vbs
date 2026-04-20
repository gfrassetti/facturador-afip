' Solo interfaz, sin consola negra (pyw no abre terminal).
Set sh = CreateObject("WScript.Shell")
Set fso = CreateObject("Scripting.FileSystemObject")
dir = fso.GetParentFolderName(WScript.ScriptFullName)
sh.CurrentDirectory = dir
py = dir & "\facturador_ui.py"
' 1 = ventana normal (se ve la app); False = no esperar a que termine
sh.Run "pyw -3 """ & py & """", 1, False
