Set sh = CreateObject("WScript.Shell")
Set fso = CreateObject("Scripting.FileSystemObject")
ps  = sh.ExpandEnvironmentStrings("%SystemRoot%\System32\WindowsPowerShell\v1.0\powershell.exe")
ps1 = fso.GetAbsolutePathName("LabDash-Setup-Host.ps1")
cmd = """" & ps & """ -NoProfile -ExecutionPolicy Bypass -Sta -WindowStyle Hidden -File " & """" & ps1 & """""
sh.Run cmd, 0, False

