@echo off

%1 mshta vbscript:CreateObject("Shell.Application").ShellExecute("cmd.exe","/c %~s0 ::","","runas",1)(window.close)&&exit

cd /d "%~dp0"

echo:
echo =============================
echo ��������ж���ϰ�ġ����Ӻ���2024������Ҫ����
echo =============================
echo ��   ��:������ʦ
echo ���ں�:������ʦ
echo =============================
echo:
echo ����ж��32λ�����Ӻ��ӡ�...
echo:
C:\Windows\Microsoft.NET\Framework\v4.0.30319\regasm.exe %~dp0BKOffice.dll /u
echo:
echo ����ж��64λ�����Ӻ��ӡ�...
echo:
C:\Windows\Microsoft.NET\Framework64\v4.0.30319\regasm.exe %~dp0BKOffice.dll /u
echo:
echo �����Ƴ�Officeע���...
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
echo �����Ƴ�WPSע���...
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
echo �Ƴ������б�...
echo:
reg delete "HKEY_CURRENT_USER\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\BKOffice" /f


echo �ϰ�ġ����Ӻ���2024����ж�����,��ӭʹ�ò��Ӻ���2025...
echo:
pause>nul



