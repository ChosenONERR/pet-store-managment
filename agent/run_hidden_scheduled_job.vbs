Option Explicit

Dim shell, fso, agentRoot, projectRoot, powershellExe, scriptPath, job, command
If WScript.Arguments.Count < 1 Then WScript.Quit 2
job = WScript.Arguments(0)

Set shell = CreateObject("WScript.Shell")
Set fso = CreateObject("Scripting.FileSystemObject")
agentRoot = fso.GetParentFolderName(WScript.ScriptFullName)
projectRoot = fso.GetParentFolderName(agentRoot)
powershellExe = shell.ExpandEnvironmentStrings("%SystemRoot%\System32\WindowsPowerShell\v1.0\powershell.exe")
scriptPath = fso.BuildPath(agentRoot, "run_scheduled_job.ps1")
command = Quote(powershellExe) & " -NoLogo -NoProfile -NonInteractive -ExecutionPolicy Bypass -File " & _
          Quote(scriptPath) & " -Job " & Quote(job) & " -ProjectRoot " & Quote(projectRoot)

' Keep the window hidden, but wait for the PowerShell child so Task Scheduler
' records the real exit code and applies its execution limit to the real job.
Dim exitCode
exitCode = shell.Run(command, 0, True)
WScript.Quit exitCode

Function Quote(ByVal value)
    Quote = Chr(34) & value & Chr(34)
End Function
