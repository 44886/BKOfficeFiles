@echo off

%1 mshta vbscript:CreateObject("Shell.Application").ShellExecute("cmd.exe","/c %~s0 ::","","runas",1)(window.close)&&exit

cd /d "%~dp0"

echo:
echo =============================
echo 现在是在卸载老版的“不坑盒子2024”，不要惊慌
echo =============================
echo 作   者:不坑老师
echo 公众号:不坑老师
echo =============================
echo:
echo 正在卸载32位【不坑盒子】...
echo:
C:\Windows\Microsoft.NET\Framework\v4.0.30319\regasm.exe %~dp0BKOffice.dll /u
echo:
echo 正在卸载64位【不坑盒子】...
echo:
C:\Windows\Microsoft.NET\Framework64\v4.0.30319\regasm.exe %~dp0BKOffice.dll /u
echo:
echo 正在移除Office注册表...
echo:
reg delete "HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\Word\Addins\BKOffice.Word" /f
reg delete HKLM\SOFTWARE\Microsoft\Office\Word\Addins\BKOffice.Word /f
reg delete "HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\Excel\Addins\BKOffice.Excel" /f
reg delete HKLM\SOFTWARE\Microsoft\Office\Excel\Addins\BKOffice.Excel /f
reg delete "HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\PowerPoint\Addins\BKOffice.PPT" /f
reg delete HKLM\SOFTWARE\Microsoft\Office\PowerPoint\Addins\BKOffice.PPT /f
reg delete HKLM\SOFTWARE\WOW6432Node\Microsoft\Office\Word\Addins\BKOffice.Word /f
reg delete HKLM\SOFTWARE\WOW6432Node\Microsoft\Office\Excel\Addins\BKOffice.Excel /f
reg delete HKLM\SOFTWARE\WOW6432Node\Microsoft\Office\PowerPoint\Addins\BKOffice.PPT /f
echo:
echo 正在移除WPS注册表...
echo:
reg delete HKEY_CURRENT_USER\Software\Kingsoft\Office\WPS\AddinsWL\  /v BKOffice.Word /f
reg delete HKEY_CURRENT_USER\Software\Kingsoft\Office\ET\AddinsWL\ /v BKOffice.Excel /f
reg delete HKEY_CURRENT_USER\Software\Kingsoft\Office\WPP\AddinsWL\ /v BKOffice.PPT /f

reg delete HKEY_LOCAL_MACHINE\Software\Kingsoft\Office\WPS\AddinsWL\  /v BKOffice.Word /f
reg delete HKEY_LOCAL_MACHINE\Software\Kingsoft\Office\ET\AddinsWL\ /v BKOffice.Excel /f
reg delete HKEY_LOCAL_MACHINE\Software\Kingsoft\Office\WPP\AddinsWL\ /v BKOffice.PPT /f

reg delete HKEY_LOCAL_MACHINE\Software\WOW6432Node\Kingsoft\Office\WPS\AddinsWL\  /v BKOffice.Word /f
reg delete HKEY_LOCAL_MACHINE\Software\WOW6432Node\Kingsoft\Office\ET\AddinsWL\ /v BKOffice.Excel /f
reg delete HKEY_LOCAL_MACHINE\Software\WOW6432Node\Kingsoft\Office\WPP\AddinsWL\ /v BKOffice.PPT /f
echo:
echo 移除程序列表...
echo:
reg delete "HKEY_CURRENT_USER\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\BKOffice" /f


echo 老版的“不坑盒子2024”，卸载完毕,欢迎使用不坑盒子2025...
echo:
pause>nul



