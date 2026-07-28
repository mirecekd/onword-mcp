Attribute VB_Name = "onwordAutoStart"
' ---------------------------------------------------------------------------
' Optional: start onword-mcp automatically whenever Word starts.
'
' Install:
'   1. Word -> Alt+F11 (VBA editor)
'   2. File -> Import File... -> this AutoExec.bas
'      (it lands in Normal -> Modules, so it applies to every document)
'   3. Adjust LAUNCHER_PATH below to where you copied the launcher folder.
'   4. Save Normal (Ctrl+S in the VBA editor) and restart Word.
'
' Uninstall: delete the "onwordAutoStart" module from Normal.
'
' Notes:
'   - AutoExec runs on Word startup. It fires and forgets: the launcher itself
'     waits for the server and supervises the SSH tunnel, and it exits on its
'     own when Word is closed.
'   - The launcher is idempotent - if port 18347 already serves, it does
'     nothing, so opening a second Word window is harmless.
'   - Corporate macro policies may block macros in Normal.dotm entirely. If so,
'     use the desktop shortcut (Install-Shortcut.ps1) instead.
' ---------------------------------------------------------------------------

Option Explicit

Private Const LAUNCHER_PATH As String = "C:\tools\onword-mcp\launcher\onword-launch.vbs"

Public Sub AutoExec()
    StartOnwordMcp
End Sub

Public Sub StartOnwordMcp()
    Dim fso As Object
    Dim shell As Object

    On Error GoTo Fail

    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FileExists(LAUNCHER_PATH) Then Exit Sub

    Set shell = CreateObject("WScript.Shell")
    ' 0 = hidden, False = do not wait
    shell.Run "wscript.exe """ & LAUNCHER_PATH & """", 0, False
    Exit Sub

Fail:
    ' never block Word startup because of this
End Sub
