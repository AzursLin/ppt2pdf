@echo off & title 
color 0a
#taskkill /F /IM soffice.exe >nul
java -jar ppt2pdf.jar
pause