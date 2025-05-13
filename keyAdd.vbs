Dim objShell, chaveLicenca, comando

If Not WScript.Arguments.Named.Exists("elevate") Then
  CreateObject("Shell.Application").ShellExecute WScript.FullName _
    , WScript.ScriptFullName & " /elevate", "", "runas", 1
  WScript.Quit
End If

chaveLicenca = InputBox("Input Office key:", "Input Office key")

If chaveLicenca = "" Then
    MsgBox "No key inserted, stopping.", vbExclamation, "Stopped"
    WScript.Quit
End If

Set objShell = CreateObject("WScript.Shell")

' Assuming it's a Office x64 installation
scriptPath = """C:\Program Files\Microsoft Office\Office16\OSPP.VBS"""

comando = "cscript //nologo " & scriptPath & " /inpkey:" & chaveLicenca

objShell.Run comando, 1, True

MsgBox "Done. Open Office to check the validation status", vbInformation, "Finished"
