Set sh = CreateObject("WScript.Shell")
Set fso = CreateObject("Scripting.FileSystemObject")
ps  = sh.ExpandEnvironmentStrings("%SystemRoot%\System32\WindowsPowerShell\v1.0\powershell.exe")
ps1 = fso.GetAbsolutePathName("LabDash-Setup-Host.ps1")
cmd = """" & ps & """ -NoProfile -ExecutionPolicy Bypass -Sta -NoExit -File " & """" & ps1 & """""
' Show console and wait, so you can see any errors
sh.Run cmd, 1, True

