@echo off
cls

rem Declare LOG file
set logDir=C:\
set logFile=%logDir%ExcelAttachment.log

rem Declare global variables
set JavaDir=C:\Users\DIEGOPC\Documents\NetBeansProjects\createExcelAttachmentNoMaven\dist\
set JarName=createExcelAttachmentNoMaven.jar

rem Declare parameters
set fileRead="C:/base_ejemplo.xlsx"
set srcTemp="C:/CreateAttachment/Excel"
set srcDestiny="C:/CreateAttachment/Zip"

rem Creating attachment
echo =========================================================== >> %logFile%
echo Start Create Excel Attachment date: %date%, time: %time% >> %logFile%
echo. >> %logFile%
echo Params: %fileRead% - %filter% >> %logFile%
echo. >> %logFile%
call java -jar %JavaDir%%JarName% %fileRead% %srcTemp% %srcDestiny% >> %logFile%
echo. >> %logFile%
echo Log file in %logFile% >> %logFile%
echo. >> %logFile%

pause

exit
