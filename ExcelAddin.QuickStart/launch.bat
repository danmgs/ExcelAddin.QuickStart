REM https://msdn.microsoft.com/fr-fr/library/bb772100.aspx
"C:\Program Files (x86)\Common Files\Microsoft Shared\VSTO\10.0\VSTOInstaller.exe" /Install C:\MyEXCELSetup\ExcelAddin.QuickStart.vsto /Silent 
echo %errorLevel% 
start excel /r /x 
pause