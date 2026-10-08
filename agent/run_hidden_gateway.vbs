Option Explicit

Dim shell, fso, agentRoot, projectRoot, powershellExe, scriptPath, command
Set shell = CreateObject("WScript.Shell")
Set fso = CreateObject("Scripting.FileSystemObject")

agentRoot = fso.GetParentFolderName(WScript.ScriptFullName)
projectRoot = fso.GetParentFolderName(agentRoot)
powershellExe = shell.ExpandEnvironmentStrings("%SystemRoot%\System32\WindowsPowerShell\v1.0\powershell.exe")
scriptPath = fso.BuildPath(agentRoot, "start_hidden_gateway.ps1")
command = Quote(powershellExe) & " -NoLogo -NoProfile -NonInteractive -ExecutionPolicy Bypass -File " & _
          Quote(scriptPath) & " -ProjectRoot " & Quote(projectRoot)

' pythonw.exe is selected by start_hidden_gateway.ps1, so Windows Terminal never owns the process.
shell.Run command, 0, False

Function Quote(ByVal value)
    Quote = Chr(34) & value & Chr(34)
End Function

