@echo off
setlocal enabledelayedexpansion

:: Terminal Initialization (applies for entire session)
mode con: cols=80 lines=25
set "SCRIPT_DIR=%~dp0"
set "BG_PROGRAM=cmdbkg.exe"
set "BG_IMAGE=123.bmp"
set "api_url=https://api.github.com/repos/Parth-Sancheti-5/WaBulkSender/tags"
set "url=https://github.com/Parth-Sancheti-5/WaBulkSender/releases"
set "Current=v3.5"
set "Maintainer=Parth-Sancheti-5"
set "name=WaBulkSender"
set "WELCOME_SCRIPT=wn.vbs"
set "last_scan=%SCRIPT_DIR%last_scan.txt"
set "LOGO_SCRIPT=logo.bat"
set "Dots=..."

title WaBulkMessageSender
color E

:: Disable standby and monitor timeout
powercfg /change standby-timeout-ac 0
powercfg /change standby-timeout-dc 0
powercfg /change monitor-timeout-ac 0
powercfg /change monitor-timeout-dc 0

:: Start background program with image
"%BG_PROGRAM%" "%BG_IMAGE%" /t 5 /c Spacer

:: -------------------------------
:: Continue with logo/welcome scripts
call "%LOGO_SCRIPT%"
start "" "%WELCOME_SCRIPT%"
timeout /t 2 /nobreak >nul

:: Auto Update Check
call :Auto_Update

:: -------------------------------
:: First-run Check (installation/login only once)
if not exist "%SCRIPT_DIR%first_run.flag" (
    call :Header "Opening Github"
    start "" "https://github.com/Parth-Sancheti-5/"

    call :Header "Requirement"
    echo "Do you have the WhatsApp Beta app installed? (Y/N)"
    set /p "has_app=Choice: "
    if /i "!has_app!"=="N" (
        call :Header "Install WhatsApp Beta"
        echo Opening Microsoft Store to install WhatsApp Beta...
        start "" "https://apps.microsoft.com/detail/9nbdxk71nk08?hl=en-US&gl=US"
        pause
    )
    call :Header "WhatsApp Login"
    echo Opening WhatsApp Beta for login...
    call :Formater
    echo Please log in to WhatsApp Beta if not already logged in.
    call :Formater
    start "" "WhatsApp_Beta.lnk"
    echo.
    pause
    rem Create flag file so subsequent runs skip these prompts.
    type nul > "%SCRIPT_DIR%first_run.flag"
)

:: Welcome Message
call :Header "Instructions"
echo - Ensure WhatsApp Beta App Is Installed On Your System.
call :Formater
echo - Keep This Window Open During Execution.
call :Formater
echo - Phone Numbers Must Include Country Code (e.g., +919876543210).
call :Formater
echo - Don't Touch/Use the System When the Script IS WORKING.
call :Formater
timeout /t 3 /nobreak >nul

:: -------------------------------
:: Attachment Options: Ask for Image or Video
call :Header "Attachment Option"
echo Does your message contain an IMAGE? (Y/N)
set /p "img_choice=Choice: "
if /i "!img_choice!"=="Y" (
    for /f "delims=" %%I in ('powershell -command "Add-Type -AssemblyName System.Windows.Forms; $ofd = New-Object System.Windows.Forms.OpenFileDialog; $ofd.Filter = 'Image Files (*.bmp;*.jpg;*.jpeg;*.png)|*.bmp;*.jpg;*.jpeg;*.png|All Files (*.*)|*.*'; if($ofd.ShowDialog() -eq 'OK'){Write-Output $ofd.FileName}"') do set "imgfile=%%I"
    echo Selected image: !imgfile!
) else (
    set "imgfile="
)
echo.

:: Enforce only one attachment type: if both selected, only use image.
if defined imgfile if defined vidfile (
    echo Both image and video selected. Only image will be used.
    set "vidfile="
)
timeout /t 2 /nobreak >nul

:: -------------------------------
:: Prepare Inputs: Edit Phone Numbers and Message
call :Header "Prepare Your Inputs"
echo Opening Notepad to edit phone numbers (num.txt)...
if not exist num.txt type nul > num.txt
start /wait notepad.exe num.txt
echo Opening Notepad to edit your message (msg.txt)...
if not exist msg.txt type nul > msg.txt
start /wait notepad.exe msg.txt

:: Read Message from msg.txt (first line)
set /p "message="<msg.txt

:: Count Total Phone Numbers
set "count=0"
for /f "usebackq delims=" %%a in ("num.txt") do (
    set /a count+=1
)
set "total=%count%"
call :Formater
echo Total phone numbers found: !total!

:: Get Delay Range from User
call :Header "Set Random Delay Range"
echo Please specify the delay range for chat loading.
set /p "min_delay=Enter Min Delay in Sec (Default 1): "
call :Formater
if "%min_delay%"=="" set "min_delay=1"
set /p "max_delay=Enter Max Delay in Sec (Default 10): "
if "%max_delay%"=="" set "max_delay=10"

:: ================================
:: Create VBScript for Sending Text Message (sendmessage.vbs)
(
    echo Set WshShell = WScript.CreateObject("WScript.Shell"^)
    echo Do
    echo     WScript.Sleep 100
    echo     success = WshShell.AppActivate("WhatsApp - Google Chrome"^)
    echo Loop Until success = True
    echo WScript.Sleep 1000
    echo WshShell.SendKeys "^v"
    echo WScript.Sleep 500
    echo WshShell.SendKeys "{ENTER}"
) > sendmessage.vbs

:: Create VBScript for Sending Image (sendimage.vbs)
(
    echo Set WshShell = WScript.CreateObject("WScript.Shell"^)
    echo Do
    echo     WScript.Sleep 100
    echo     success = WshShell.AppActivate("WhatsApp - Google Chrome"^)
    echo Loop Until success = True
    echo WScript.Sleep 1000
    echo WshShell.SendKeys "^v"
) > sendimage.vbs

:: ================================
:: Process Each Phone Number & Send Message/Attachment
call :Header "Sending Messages"
set "count=0"
for /f "tokens=*" %%a in (num.txt) do (
    set /a count+=1
    echo.
    echo Sending message !count! of !total! to %%a
    :: Open chat in WhatsApp Web using wa.me link
    start "" "https://wa.me/%%a"
    :: Generate a random delay within specified range
    set /a "delay=!random! %% (!max_delay! - !min_delay! + 1) + !min_delay!"
    
    echo Waiting !delay! seconds for chat to load...
    timeout /t !delay! /nobreak >nul

    :: If image attachment defined, paste image first
    if defined imgfile (
        echo Copying image to clipboard...
        powershell -command "Add-Type -AssemblyName System.Windows.Forms; Add-Type -AssemblyName System.Drawing; [System.Windows.Forms.Clipboard]::SetImage([System.Drawing.Image]::FromFile('!imgfile!'))"
        timeout /t 2 /nobreak >nul
        echo Pasting image...
        cscript //nologo sendimage.vbs
        timeout /t 2 /nobreak >nul
    )

    :: Now send the text message (copy text, wait 2 seconds, then paste and press Enter)
    echo Copying text message to clipboard...
    echo !message! | clip
    timeout /t 2 /nobreak >nul
    echo Pasting text message...
    cscript //nologo sendmessage.vbs

    echo Waiting 7 seconds before next message...
    timeout /t 7 /nobreak >nul
)

:: -------------------------------
:: Completion Message and Reset Timeouts
:Quit_Toolbox
echo Total messages sent: !total!
powercfg /change standby-timeout-ac 30
powercfg /change standby-timeout-dc 10
powercfg /change monitor-timeout-ac 10
powercfg /change monitor-timeout-dc 5
del "%SCRIPT_DIR%sendmessage.vbs"
del "%SCRIPT_DIR%sendimage.vbs"

call :Spacer 6
echo      "+----------------------------------------+"
echo      "|                                        |"
call :Formater
echo      "|   Thank-You-For-Using-WaBulkSender     |"
echo      "|            By-@Parth-Sancheti          |"
echo      "|       Have-A-Productive-Day-Ahead      |"
echo      "|                                        |"
call :Formater
echo      "+----------------------------------------+"
call :Spacer 6
timeout /t 3 >nul
exit /b

:: -------------------------------
:: Subroutine: Header
:Header
cls
echo ------------------------------------------------------------------
call :Formater
call :CenterText "%~1"
call :CenterText "Created By: @Parth-Sancheti"
call :Formater
echo ------------------------------------------------------------------
call :Formater
goto :EOF

:: Subroutine: CenterText - Center text within 80 columns
:CenterText
set "text=%~1"
set /a console_width=80
set /a text_length=0
for /l %%i in (0,1,4095) do (
    if "!text:~%%i,1!"=="" (
        set /a text_length=%%i
        goto :done_length
    )
)
:done_length
set /a padding=(console_width - text_length) / 2
set "spaces="
for /l %%i in (1,1,!padding!) do set "spaces=!spaces! "
echo !spaces!!text!
goto :EOF

:: Subroutine: Spacer - Print blank lines
:Spacer
set "i=0"
set "n=%~1"
:SpacerLoop
if !i! lss %n% (
    echo.
    set /a i+=1
    goto SpacerLoop
)
goto :EOF

:: Subroutine: Formater - Print a blank line with a short pause
:Formater
echo.
ping 127.0.0.1 -n 1 -w 200 >nul
goto :EOF

:Auto_Update
if not exist "%last_scan%" (
    powershell -command "Get-Date -Format 'yyyy-MM-dd'" > "%last_scan%"
)
for /f "tokens=* delims=" %%a in ('powershell -command "$lastScanDate = Get-Content -Path '%last_scan%'; $today = Get-Date; $diff = ($today - [datetime]$lastScanDate).Days; $diff"') do set "days_difference=%%a"
if !days_difference! gtr 7 (
    powershell -command "Get-Date -Format 'yyyy-MM-dd'" > "%last_scan%"
    timeout /t 1 >nul
    goto :Check_For_Update
)
goto :EOF

:Check_For_Update
call :Header "Check_For_Update" %name%
        echo Checking for updates
        curl -s %api_url% > data.json
            for /f %%i in ('powershell -ExecutionPolicy Bypass -Command "(Get-Content -Raw -Path data.json | ConvertFrom-Json)[0].name"') do (
                set "last=%%i"
            )
REM Extract major version number (assuming the format is "vXX.XX-XXXX")
    for /f "tokens=1 delims=v.-" %%a in ("%Current%") do set "num1=%%a"
    for /f "tokens=1 delims=v.-" %%a in ("%last%") do set "num2=%%a"
    if !num1! gtr !num2! (
        call :Header Dev_Mod %name%
        echo You're using version: %Current%
        echo The latest version available is: %last%
    ) else if !num1! lss !num2! (
        echo You're using version: %Current%
        echo The latest version available is: %last%
        call :Download "%last%" "%last%"
    ) else (
        echo You're on the latest version: %Current%
        echo %last%
        echo Thank you for staying updated!
    )
    goto :EOF

:Download
call :Header "Downloading" %name%
set "B1=%~1"
set "B2=%~2"
echo Updating from version %Current% to %B2%
echo You're about to download the update.
call :Choice_Maker "Do you wish to continue? (Y/N)"
call :case Y y
call :case N n
if /i "%user_choice%"=="N" goto :Quit_Toolbox
set "downloadUrl=https://github.com/%Maintainer%/%name%/releases/download/%B1%/%name%-!B2!.exe"
set "outputFile=Update_Setup.exe"
echo Starting The Download %Dots%
echo Please Wait %Dots%
curl -L --progress-bar -o "%temp%\%outputFile%" "%downloadUrl%"
if errorlevel 1 (
    echo Download failed. Please check your internet connection or try again later.
    goto :Quit_Toolbox
)
echo File downloaded to: %temp%\%outputFile%
echo To start setup, press any key %Dots%
pause >nul
start "" "%temp%\%outputFile%"
goto :Quit_Toolbox

:continue_or_quit
set /p "user_continue=Do you want to continue? (Y/N): "
if /i "!user_continue!"=="N" goto :Quit_Toolbox
goto :EOF

:case
if /i "%user_choice%"=="%2" set "user_choice=%1"
goto :EOF

:Choice_Maker
set /p "user_choice=%~1 :"
goto :EOF

