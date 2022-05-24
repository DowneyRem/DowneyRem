# 重新启用休眠并休眠

# 获取管理员权限执行
if (!([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) { Start-Process powershell.exe "-NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`"" -Verb RunAs; exit }

chcp 65001
Write-Host "进入休眠"
powercfg.exe /hibernate off
powercfg.exe /hibernate on
Write-Output "休眠中"
shutdown -h 