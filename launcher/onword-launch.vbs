' Hidden wrapper for onword-launch.ps1 - double-clicking this (or the desktop
' shortcut created by Install-Shortcut.ps1) starts Word + the MCP server + the
' reverse SSH tunnel without flashing a console window.
'
' Any arguments are passed through to the PowerShell script, so the shortcut
' can also be used as a drop target / "Open with" handler for .docx files:
'   wscript onword-launch.vbs "C:\docs\report.docx"

Option Explicit

Dim fso, shell, scriptDir, psScript, cmd, i, arg

Set fso = CreateObject("Scripting.FileSystemObject")
Set shell = CreateObject("WScript.Shell")

scriptDir = fso.GetParentFolderName(WScript.ScriptFullName)
psScript = fso.BuildPath(scriptDir, "onword-launch.ps1")

If Not fso.FileExists(psScript) Then
    MsgBox "onword-launch.ps1 not found next to this script:" & vbCrLf & psScript, _
           vbCritical, "onword-mcp"
    WScript.Quit 1
End If

cmd = "powershell.exe -NoProfile -NonInteractive -ExecutionPolicy Bypass -File """ & psScript & """"

' First argument (if it looks like a path) becomes -Document, the rest are
' passed verbatim so named parameters like -NoCache keep working.
For i = 0 To WScript.Arguments.Count - 1
    arg = WScript.Arguments(i)
    If i = 0 And Left(arg, 1) <> "-" Then
        cmd = cmd & " -Document """ & arg & """"
    Else
        cmd = cmd & " " & arg
    End If
Next

' 0 = hidden window, False = do not wait for it to finish
shell.Run cmd, 0, False
