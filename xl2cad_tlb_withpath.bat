@echo off
REM ��excel��COM�������dll·����ΪC:\Windows\System\xl2cad.dll
REM Ĭ��������Ŀ��DebugĿ¼��
echo ��excel��COM�������dll·����ΪC:\Windows\System\xl2cad.dll...
"C:\Program Files (x86)\Microsoft SDKs\Windows\v7.0A\Bin\TlbExp.exe" C:\Windows\System\xl2cad.dll
"C:\Windows\Microsoft.NET\Framework\v2.0.50727\RegAsm.exe" C:\Windows\System\xl2cad.dll /codebase
pause