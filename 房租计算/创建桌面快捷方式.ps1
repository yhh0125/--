# 创建桌面快捷方式脚本
$WshShell = New-Object -comObject WScript.Shell
$Shortcut = $WshShell.CreateShortcut("$env:USERPROFILE\Desktop\房租计算程序.lnk")
$Shortcut.TargetPath = "C:\Users\lenovo\Desktop\房租计算\一键启动房租计算.bat"
$Shortcut.WorkingDirectory = "C:\Users\lenovo\Desktop\房租计算"
$Shortcut.Description = "一键启动房租计算程序"
$Shortcut.IconLocation = "C:\Users\lenovo\AppData\Local\Programs\Python\Python313\python.exe,0"
$Shortcut.Save()

Write-Host "桌面快捷方式创建成功！" -ForegroundColor Green
Write-Host "您可以在桌面上找到 '房租计算程序' 快捷方式" -ForegroundColor Yellow
Read-Host "按回车键退出"

