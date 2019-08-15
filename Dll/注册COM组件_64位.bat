@ECHO OFF & CD /D %~DP0 & TITLE 注册COM组件
>NUL 2>&1 REG.exe query "HKU\S-1-5-19" || (
    ECHO SET UAC = CreateObject^("Shell.Application"^) > "%TEMP%\Getadmin.vbs"
    ECHO UAC.ShellExecute "%~f0", "%1", "", "runas", 1 >> "%TEMP%\Getadmin.vbs"
    "%TEMP%\Getadmin.vbs"
    DEL /f /q "%TEMP%\Getadmin.vbs" 2>NUL
    Exit /b
)

SPF_PROXY.exe /regserver

CLS && ECHO. & ECHO 注册COM组件完成，请按任意键退出！ &&PAUSE >NUL & EXIT