@echo off
%1 mshta vbscript:CreateObject("Shell.Application").ShellExecute("cmd.exe","/c %~s0 ::","","runas",1)(window.close)&&exit

reg copy HKCU\SOFTWARE\Microsoft\Office\Word\Addins\BKOffice.Word HKLM\SOFTWARE\Microsoft\Office\Word\Addins\BKOffice.Word /s /f
reg copy HKCU\SOFTWARE\Software\Kingsoft\Office\WPS\AddinsWL HKLM\Software\Kingsoft\Office\WPS\AddinsWL /s /f

reg copy HKCU\SOFTWARE\Microsoft\Office\Excel\Addins\BKOffice.Excel HKLM\SOFTWARE\Microsoft\Office\Excel\Addins\BKOffice.Excel /s /f
reg copy HKCU\SOFTWARE\Software\Kingsoft\Office\ET\AddinsWL HKLM\Software\Kingsoft\Office\ET\AddinsWL /s /f

reg copy HKCU\SOFTWARE\Microsoft\Office\PowerPoint\Addins\BKOffice.PPT HKLM\SOFTWARE\Microsoft\Office\PowerPoint\Addins\BKOffice.PPT /s /f
reg copy HKCU\SOFTWARE\Software\Kingsoft\Office\WPP\AddinsWL HKLM\Software\Kingsoft\Office\WPP\AddinsWL /s /f