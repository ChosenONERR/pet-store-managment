Option Explicit

Dim shell, fso, agentRoot, projectRoot, powershellExe, watchdogScript, command, exitCode

Set shell = CreateObject("WScript.Shell")
Set fso = CreateObject("Scripting.FileSystemObject")

agentRoot = fso.GetParentFolderName(WScript.ScriptFullName)
projectRoot = fso.GetParentFolderName(agentRoot)
powershellExe = shell.ExpandEnvironmentStrings("%SystemRoot%\System32\WindowsPowerShell\v1.0\powershell.exe")
watchdogScript = fso.BuildPath(agentRoot, "ensure_lark_event_bridge.ps1")

command = Quote(powershellExe) & " -NoLogo -NoProfile -NonInteractive -ExecutionPolicy Bypass -File " & _
          Quote(watchdogScript) & " -ProjectRoot " & Quote(projectRoot)

' Window style 0 keeps the five-minute watchdog entirely off the interactive desktop.
exitCode = shell.Run(command, 0, True)
WScript.Quit exitCode

Function Quote(ByVal value)
    Quote = Chr(34) & value & Chr(34)
End Function
