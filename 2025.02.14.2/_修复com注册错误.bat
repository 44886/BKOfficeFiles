@echo off
%1 mshta vbscript:CreateObject("Shell.Application").ShellExecute("cmd.exe","/c %~s0 ::","","runas",1)(window.close)&&exit

reg delete "HKEY_CLASSES_ROOT\TypeLib\{00020905-0000-0000-C000-000000000046}\8.7" /f