@echo off
REM Get dir of script
   SET MYDIR=%~dp0

REM Taken from the Stack Exchange Network. Based on http://stackoverflow.com/a/9872111 by VonC (https://stackoverflow.com/users/73070)
REM Licensed under cc-by-sa 3.0 https://creativecommons.org/licenses/by-sa/3.0/
REM Gets UTC time and saves them in %Hour%, etc.
   for /f %%x in ('wmic path win32_utctime get /format:list ^| findstr "="') do SET %%x

REM set 12 hour offset as the comic names are in UTC + 12 hours
   SET /a Hour=%Hour% + 12

REM choose right minute parameter based on elapsed seconds 
   SET m=00
   SET /a Second=%Second% + (%Minute% * 60)
   REM If more than 52.5 minutes have elapsed, choose next hour picture
   IF %Second% GTR 3150 SET /A Hour=%Hour% + 1
   REM If less than 52.5 minutes have elapsed, choose 45 minute parameter
   IF %Second% LSS 3150 SET m=45
   REM If less than 37.5 minutes have elapsed, choose 30 minute parameter
   IF %Second% LSS 2250 SET m=30
   REM If less than 22.5 minutes have elapsed, choose 15 minute parameter
   IF %Second% LSS 1350 SET m=15
   REM If less than  7.5 minutes have elapsed, choose 15 minute parameter
   IF %Second% LSS  450 SET m=00

REM Check Hour to be smaller than 24 and format it to hh with leading zeroes
   IF %Hour% GTR 23 SET /a Hour-=24
   SET h=00%Hour%
   SET h=%h:~-2%

REM actually set wallpaper
   ECHO Set wallpaper to: %MYDIR%images\%h%h%m%m.bmp
   powershell.exe -File "%MYDIR%ChangeWallpaper.ps1" "%MYDIR%images\%h%h%m%m.jpg"