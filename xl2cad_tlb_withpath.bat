@echo off
REM 把excel中COM加载项的dll路径改为C:\Windows\System\xl2cad.dll
REM 默认是在项目的Debug目录下
echo 把excel中COM加载项的dll路径改为C:\Windows\System\xl2cad.dll...
"C:\Program Files (x86)\Microsoft SDKs\Windows\v7.0A\Bin\TlbExp.exe" C:\Windows\System\xl2cad.dll
"C:\Windows\Microsoft.NET\Framework\v2.0.50727\RegAsm.exe" C:\Windows\System\xl2cad.dll /codebase
pause