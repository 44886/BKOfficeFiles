@echo off

%1 mshta vbscript:CreateObject("Shell.Application").ShellExecute("cmd.exe","/c %~s0 ::","","runas",1)(window.close)&&exit

cd /d "%~dp0"

chcp 936

echo:
echo =============================
echo ��ӭʹ�á����Ӻ��ӡ�Office���
echo =============================
echo ��   ��:������ʦ
echo ���ں�:������ʦ
echo =============================
echo:
echo ���ڼ�������л���...
setlocal enabledelayedexpansion
set existnet=false 
for /f "tokens=7 delims=\" %%a in ('REG QUERY HKLM\SOFTWARE\Microsoft\.NETFramework\v4.0.30319\SKUs') do (
	if not "%%a" == "Client" if not "%%a" == "Default" (
		if "%1" == "" (
			SET gg=%%a			
			SET ss=!gg:~22,6!
			SET ss=!ss:,=!
			SET ss=!ss:P=!
			rem ��ӡ.net framework�汾��
			rem echo !ss!
			if "!ss:v4.8=!" NEQ "!ss!" (
				set existnet=true
			)
		) else (			
			if "!%%a:v4.8=!" NEQ "!%%a!" (
			 	set existnet=true
			)	
			goto exit			
		)
	)
)
:exit
if %existnet% == true (
	echo [32m���л���������,�Ѱ�װ .net framework v4.8���л���[0m
) else (
	echo [31mȱ�����л�����������˳�,��װ���������ִ�а�װ...[0m
	start https://dotnet.microsoft.com/zh-cn/download/dotnet-framework/thank-you/net48-offline-installer
	pause>nul
exit
 )


echo:
echo �������Officeע���...
echo:
rem ע������: /v �������� /t �������� /d ���� /f ����ʾǿ���޸�
echo ע��Word���...
reg add "HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\Word\Addins\BKOffice2025.Word" /f
reg add "HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\Word\Addins\BKOffice2025.Word" /v "FriendlyName" /t REG_SZ /d "���Ӻ���2025" /f
reg add "HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\Word\Addins\BKOffice2025.Word" /v "Description" /t REG_SZ /d "һ��ȫ�ܵ�Office�����ӵ�ж�������͹���" /f
reg add "HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\Word\Addins\BKOffice2025.Word" /v "LoadBehavior" /t REG_DWORD /d "3" /f
reg add "HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\Word\Addins\BKOffice2025.Word" /v "CommandLineSafe" /t REG_DWORD /d "1" /f
reg add "HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\Word\Addins\BKOffice2025.Word" /v "Manifest" /t REG_SZ /d "file:///%~dp0BKOffice2025.Word.vsto|vstolocal" /f
reg copy "HKEY_CURRENT_USER\Software\Microsoft\Office\Word\Addins\BKOffice2025.Word" "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Office\Word\Addins\BKOffice2025.Word" /s /f
echo:
echo ע��Excel���...
reg add "HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\Excel\Addins\BKOffice2025.Excel" /f
reg add "HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\Excel\Addins\BKOffice2025.Excel" /v "FriendlyName" /t REG_SZ /d "���Ӻ���2025" /f
reg add "HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\Excel\Addins\BKOffice2025.Excel" /v "Description" /t REG_SZ /d "һ��ȫ�ܵ�Office�����ӵ�ж�������͹���" /f
reg add "HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\Excel\Addins\BKOffice2025.Excel" /v "LoadBehavior" /t REG_DWORD /d "3" /f
reg add "HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\Excel\Addins\BKOffice2025.Excel" /v "CommandLineSafe" /t REG_DWORD /d "1" /f
reg add "HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\Excel\Addins\BKOffice2025.Excel" /v "Manifest" /t REG_SZ /d "file:///%~dp0BKOffice2025.Excel.vsto|vstolocal" /f
reg copy "HKEY_CURRENT_USER\Software\Microsoft\Office\Excel\Addins\BKOffice2025.Excel" "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Office\Excel\Addins\BKOffice2025.Excel" /s /f
echo:
echo ע��PPT���...
reg add "HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\PowerPoint\Addins\BKOffice2025.PPT" /f
reg add "HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\PowerPoint\Addins\BKOffice2025.PPT" /v "FriendlyName" /t REG_SZ /d "���Ӻ���2025" /f
reg add "HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\PowerPoint\Addins\BKOffice2025.PPT" /v "Description" /t REG_SZ /d "һ��ȫ�ܵ�Office�����ӵ�ж�������͹���" /f
reg add "HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\PowerPoint\Addins\BKOffice2025.PPT" /v "LoadBehavior" /t REG_DWORD /d "3" /f
reg add "HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\PowerPoint\Addins\BKOffice2025.PPT" /v "CommandLineSafe" /t REG_DWORD /d "1" /f
reg add "HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\PowerPoint\Addins\BKOffice2025.PPT" /v "Manifest" /t REG_SZ /d "file:///%~dp0BKOffice2025.PPT.vsto|vstolocal" /f
reg copy "HKEY_CURRENT_USER\Software\Microsoft\Office\PowerPoint\Addins\BKOffice2025.PPT" "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Office\PowerPoint\Addins\BKOffice2025.PPT" /s /f

echo:
echo �������WPSע���...
reg add "HKEY_CURRENT_USER\Software\Kingsoft\Office\WPS\AddinsWL" /f
reg add "HKEY_CURRENT_USER\Software\Kingsoft\Office\WPS\AddinsWL" /v "BKOffice2025.Word" /t REG_SZ /d "" /f
reg copy "HKEY_CURRENT_USER\Software\Kingsoft\Office\WPS\AddinsWL" "HKEY_LOCAL_MACHINE\Software\Kingsoft\Office\WPS\AddinsWL" /s /f

reg add "HKEY_CURRENT_USER\Software\Kingsoft\Office\ET\AddinsWL" /f
reg add "HKEY_CURRENT_USER\Software\Kingsoft\Office\ET\AddinsWL" /v "BKOffice2025.Excel" /t REG_SZ /d "" /f
reg copy "HKEY_CURRENT_USER\Software\Kingsoft\Office\ET\AddinsWL" "HKEY_LOCAL_MACHINE\Software\Kingsoft\Office\ET\AddinsWL" /s /f

reg add "HKEY_CURRENT_USER\Software\Kingsoft\Office\WPP\AddinsWL" /f
reg add "HKEY_CURRENT_USER\Software\Kingsoft\Office\WPP\AddinsWL" /v "BKOffice2025.PPT" /t REG_SZ /d "" /f
reg copy "HKEY_CURRENT_USER\Software\Kingsoft\Office\WPP\AddinsWL" "HKEY_LOCAL_MACHINE\Software\Kingsoft\Office\WPP\AddinsWL" /s /f

echo:
echo ����ڳ����б�...
echo:
reg add "HKEY_CURRENT_USER\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\BKOFfice2025" /f
reg add "HKEY_CURRENT_USER\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\BKOFfice2025" /v "DisplayName" /t REG_SZ /d "���Ӻ���2025" /f
reg add "HKEY_CURRENT_USER\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\BKOFfice2025" /v "DisplayIcon" /t REG_SZ /d %~dp0logo.ico /f
reg add "HKEY_CURRENT_USER\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\BKOFfice2025" /v "Publisher" /t REG_SZ /d "������ʦ" /f
reg add "HKEY_CURRENT_USER\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\BKOFfice2025" /v "UninstallString" /t REG_SZ /d %~dp0_uninstall.bat /f
reg add "HKEY_CURRENT_USER\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\BKOFfice2025" /v "InstallLocation" /t REG_SZ /d %~dp0 /f
reg add "HKEY_CURRENT_USER\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\BKOFfice2025" /v "DisplayVersion" /t REG_SZ /d "2024.1022" /f
reg add "HKEY_CURRENT_USER\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\BKOFfice2025" /v "EstimatedSize" /t REG_DWORD /d "28045" /f
reg add "HKEY_CURRENT_USER\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\BKOFfice2025" /v "URLInfoAbout" /t REG_SZ /d "https://www.bukenghezi.com" /f

echo:
echo ��װ���塭��

:: ���������ļ�·��
set "font_folder=%~dp0\fonts"

:: �����ļ����е����������ļ�
for %%f in ("%font_folder%\*.ttf") do (
    echo ���ڰ�װ����: %%~nxf
    copy "%%f" "C:\Windows\Fonts"
    reg add "HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Fonts" /v "%%~nxf" /t REG_SZ /d "%%~nxf" /f
)

echo ���尲װ���

echo:
echo ������ݷ�ʽ...
set path=%~dp0

set SHORTCUT_PATH="%userprofile%\Desktop\���Ӻ���˵����.lnk"
set SHORTCUT_PATH2="%APPDATA%\Microsoft\Windows\Start Menu\Programs\���Ӻ���˵����.lnk"
set TARGET_FILE=%path%���Ӻ���.exe
set VBS_SCRIPT=""%path%\CreateShortcut.vbs"

echo Set oWS = WScript.CreateObject("WScript.Shell") > "%VBS_SCRIPT%"
echo sLinkFile = %SHORTCUT_PATH% >> "%VBS_SCRIPT%"
echo Set oLink = oWS.CreateShortcut(sLinkFile) >> "%VBS_SCRIPT%"
echo oLink.TargetPath = "%TARGET_FILE%" >> "%VBS_SCRIPT%"
echo oLink.Save >> "%VBS_SCRIPT%"

echo Set oWS2 = WScript.CreateObject("WScript.Shell") >> "%VBS_SCRIPT%"
echo sLinkFile2 = %SHORTCUT_PATH2% >> "%VBS_SCRIPT%"
echo Set oLink2 = oWS2.CreateShortcut(sLinkFile2) >> "%VBS_SCRIPT%"
echo oLink2.TargetPath = "%TARGET_FILE%" >> "%VBS_SCRIPT%"
echo oLink2.Save >> "%VBS_SCRIPT%"

"%SystemRoot%\System32\WScript.exe" //NoLogo %VBS_SCRIPT%

del %VBS_SCRIPT%

echo:
echo �����װ���!!
echo:
pause>nul