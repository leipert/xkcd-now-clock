@echo off
setlocal enableExtensions enableDelayedExpansion

REM Check whether imagemagick is installed and find path for convert.exe
REM We are doing this with the help of imagemagicks mogrify.exe, because Windows has a native convert.exe
   for %%X in (mogrify.exe) do (set IMMOGRIFY=%%~$PATH:X)
   if not defined IMMOGRIFY (
      echo "ERROR: ImageMagick is not installed! Please do so and add the path to the ImageMagick binaries to the PATH Variable."
      goto eof
   )
   set IMCONVERT=%IMMOGRIFY:mogrify.exe=%convert.exe

REM Initialise Variables
   SET downloadParam=false
   SET convertParam=false
   SET label=false
   SET /A mode=0
   SET /A degSig=0
   SET /A degDec=0
   SET counter=0
   SET ampm=1
   FOR %%A IN (%*) DO (
      IF /I "%%A"=="download" SET downloadParam=true
      IF /I "%%A"=="convert" SET convertParam=true
      IF /I "%%A"=="clock" SET mode=2
      IF /I "%%A"=="center" SET mode=4
      IF /I "%%A"=="label" SET label=true
      IF /I "%%A"=="ampm" SET ampm=2
   )
   IF /I "%label%"=="true" SET /A mode=!mode!+1
      
REM Get offset to UTC and save it in variable offsett
   set /a tzYear=%date:~-4% * 1
   set /a tzMonth=%date:~3,2% * 1
   set /a tzDay=%date:~0,2% * 1

   REM Taken from the Stack Exchange Network. Based on http://stackoverflow.com/a/9872111 by VonC (https://stackoverflow.com/users/73070)
   REM Licensed under cc-by-sa 3.0 https://creativecommons.org/licenses/by-sa/3.0/
   REM Gets UTC time and saves them in %Hour%, etc.
      for /f %%x in ('wmic path win32_utctime get /format:list ^| findstr "="') do SET %%x

   REM Taken from the Stack Exchange Network. Based on http://stackoverflow.com/a/19160337 by foxidrive (https://stackoverflow.com/users/2299431/foxidrive)
   REM Licensed under cc-by-sa 3.0 https://creativecommons.org/licenses/by-sa/3.0/
      echo Wscript.Echo DateDiff^("n",CDate^(#%tzYear%-%tzMonth%-%tzDay% %time:~0,8%#^),CDate^(#%Year%-%Month%-%Day% %Hour%:%Minute%:%Second%#^)^) > tmp.vbs
      echo Wscript.Echo DateDiff^("n",CDate^(#%Year%-%Month%-%Day% %Hour%:%Minute%:%Second%#^),CDate^(#%tzYear%-%tzMonth%-%tzDay% %time:~0,8%#^)^) > tmp.vbs
      for /f %%a in ('cscript /nologo tmp.vbs') do set /A offset=%%a
      DEL tmp.vbs

REM Check whether the last png and last jpg exist. If not, set parameters accordingly
   if not exist images\23h45m.png (
      SET downloadParam=true
   )
   if not exist images\23h45m.jpg (
      SET convertParam=true
   )

REM Remove files if they will be created again
   if "!downloadParam!"=="true" (
      DEL images\*.png
   )
   if "!convertParam!"=="true" (
      DEL images\*.jpg
   )

REM Download all images and convert them to jpg
REM Loop through all hours (0-23) of a day
   for %%x in (0 1 2 3 4 5 6 7 8 9 10 11 12 13 14 15 16 17 18 19 20 21 22 23) do set x=%%x& call :outer_loop
   REM for %%x in ( 4 5) do set x=%%x& call :outer_loop
   goto :outer_end
   :outer_loop
      REM format hour to hh with leading zeroes
         SET h=00%x%
         SET h=%h:~-2%

      REM Loop through all four 15 min intervals of an hour
         for %%y in (0 15 30 45) do set y=%%y& call :inner_loop
         goto :inner_end
         :inner_loop
            REM format minutes to mm with leading zeroes
SET /a counter+=1
            REM format minutes to mm with leading zeroes
               SET m=00%y%
               SET m=%m:~-2%

            REM Logging
               ECHO.
               set /p "=%h%h%m%m ^(!counter!/96^)" <NUL
               set /p "=!ASCII_8!..." <NUL

            REM Download src images if needed
               if "!downloadParam!"=="true" (

                  cscript /b /nologo download.vbs "http://imgs.xkcd.com/comics/now/" "images\" "%h%h%m%m.png"

                  REM Exit if download failed
                     IF NOT EXIST images\%h%h%m%m.png (
                        ECHO.
                        ECHO ERROR: Could not download images\%h%h%m%m.png. Will exit now. Please check if you have an active internet connection
                        GOTO eof
                     )

                  REM Logging
                     set /p "=!ASCII_13!downloaded" <NUL
                     set /p "=!ASCII_13!..." <NUL
               )

            REM Convert pictures if needed
               if "!convertParam!"=="true" (
                  REM Exit if src does not exist
                     IF NOT EXIST images\%h%h%m%m.png (
                        ECHO.
                        ECHO ERROR: Could not access images\%h%h%m%m.png. Please make sure this file exists!
                        GOTO eof
                     )

                  REM Calculate rotation degree of current picture based on filename and UTC offset. Degrees will be normalized (0 <= degree <= 360)
                     SET /A degSig=15 * %x% * %ampm% + 375 * %y% * %ampm% / 100 / 15 + %offset% * %ampm% * 1 / 4
                     SET /A degDec=75 * %y% * %ampm% / 15
                     SET /A degDec=!degDec:~-2!
                     :increase_loop
                     if not !degSig! LSS 0 goto :increase_end
                        SET /A degSig=!degSig!+360
                        goto :increase_loop
                     :increase_end
                     set /p "=!ASCII_13!" <NUL
                     
                  REM Calculate UTC hour based on filename
                     SET /A UTCHour=12 + %x%
                     if !UTCHour! GTR 23 SET /A UTCHour=!UTCHour! - 24
                     set UTCHour=00!UTCHour!
                     
                  REM local time based on filename and offset
                     SET /A localHour= %x% + 12 + %offset% / 60
                     SET /A localMinute= %y% + ^( %offset% - ^(^(%offset% / 60^) * 60^)^)
                     if !localMinute! GTR 60 SET /A localHour=!localHour! + 1
                     if !localMinute! GTR 60 SET /A localMinute=!localMinute! - 60
                     if !localHour! GTR 23 SET /A localHour=!localHour! - 24
                     set localHour=00!localHour!
                     set localMinute=00!localMinute!
                  REM Set text for conversions with time annotation
                     set annotation=-pointsize 16 -gravity SouthEast -annotate 0 "UTC: ~ !UTCHour:~-2!:%m%\nLocal: ~ !localHour:~-2!:!localMinute:~-2!"

                  REM Choose conversion based on selected mode
                     REM mode=0: Plain conversion png -> jpg
                        IF "%mode%"=="0" "%IMCONVERT%" images\%h%h%m%m.png images\%h%h%m%m.jpg
                     REM mode=1: Conversion png -> jpg with times annotation in bottom right corner
                        IF "%mode%"=="1" "%IMCONVERT%" images\%h%h%m%m.png !annotation! images\%h%h%m%m.jpg
                     REM mode=2: Conversion png -> jpg with addition of little dot indicating local time
                        IF "%mode%"=="2" "%IMCONVERT%" dot.png ^( +clone -background "rgba(0,0,0,0)" -rotate !degSig!.!degDec! ^) -gravity center -compose Src -composite miff:- | "%IMCONVERT%" images\%h%h%m%m.png - -composite images\%h%h%m%m.jpg
                     REM mode=3: Conversion png -> jpg with addition of little dot indicating local time and times annotation in bottom right corner
                        IF "%mode%"=="3" "%IMCONVERT%" dot.png ^( +clone -background "rgba(0,0,0,0)" -rotate !degSig!.!degDec! ^) -gravity center -compose Src -composite miff:- | "%IMCONVERT%" images\%h%h%m%m.png !annotation! - -composite images\%h%h%m%m.jpg
                     REM mode=4: Conversion png -> jpg. Local timezone will always be at 0 degrees (12 o'clock location)
                        IF "%mode%"=="4" "%IMCONVERT%" images\%h%h%m%m.png ^( +clone -background white -rotate -!degSig!.!degDec! ^) -gravity center -compose Src -composite images\%h%h%m%m.jpg
                     REM mode=5: Conversion png -> jpg. Local timezone will always be at 0 degrees (12 o'clock location) and times annotation in bottom right corner
                        IF "%mode%"=="5" "%IMCONVERT%" images\%h%h%m%m.png ^( +clone -background white -rotate -!degSig!.!degDec! ^) -gravity center -compose Src -composite miff:- | "%IMCONVERT%" - !annotation! images\%h%h%m%m.jpg

                  REM Exit if file couldn't be created
                     IF NOT EXIST images\%h%h%m%m.jpg (
                        ECHO.
                        ECHO ERROR: Could not create images\%h%h%m%m.jpg. Please check if images\%h%h%m%m.png exists.
                        GOTO eof
                     )

                  REM Logging
                     set /p "=!ASCII_13!converted" <NUL
                     set /p "=!ASCII_13!..." <NUL
               )

            REM Logging
               set /p "=!ASCII_13!finished" <NUL
               
         goto :eof
         :inner_end
   goto :eof
   :outer_end

REM Creation of template file for Task Scheduler
   type templates\template1.xml > XKCD-Now-Clock.xml
   echo ^<Author^>%COMPUTERNAME%\%username%^</Author^> >> XKCD-Now-Clock.xml
   type templates\template2.xml >> XKCD-Now-Clock.xml
   echo ^<UserId^>%COMPUTERNAME%\%username%^</UserId^> >> XKCD-Now-Clock.xml
   type templates\template2a.xml >> XKCD-Now-Clock.xml
   echo ^<UserId^>%COMPUTERNAME%\%username%^</UserId^> >> XKCD-Now-Clock.xml
   type templates\template2b.xml >> XKCD-Now-Clock.xml
   echo ^<UserId^>%COMPUTERNAME%\%username%^</UserId^> >> XKCD-Now-Clock.xml
   type templates\template3.xml >> XKCD-Now-Clock.xml
   echo ^<Arguments^>"%~dp0invis.vbs" powershell.exe -File "%~dp0RotateWallpaper.ps1"^</Arguments^> >> XKCD-Now-Clock.xml
   type templates\template4.xml >> XKCD-Now-Clock.xml

REM Logging
   ECHO.
   ECHO Assumed Offset to UTC is %offset% minutes
   echo Finished download and conversion
   echo Just import the created task file ^("XKCD-Now-Clock.xml"^) with the Task Scheduler
   
   call powershell.exe -File "%~dp0RotateWallpaper.ps1"
endlocal