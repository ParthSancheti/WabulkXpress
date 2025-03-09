@echo off
setlocal enabledelayedexpansion

:: Initialization
:: Loct()
set BG_PROGRAM=cmdbkg.exe
set BG_IMAGE=123.bmp
set WELCOME_SCRIPT=wn.vbs
set border_line=------------------------------------------------------------------
set LOGO_SCRIPT=logo.bat

:: Config()
REM Disabling Screen Timeouts
powercfg /change standby-timeout-ac 0
powercfg /change standby-timeout-dc 0
powercfg /change monitor-timeout-ac 0
powercfg /change monitor-timeout-dc 0
%BG_PROGRAM% %BG_IMAGE% /t 5 /c Spacer
mode con: cols=66 lines=25
title WaBulkMessageSender
color E
call %LOGO_SCRIPT%
start "" %WELCOME_SCRIPT%
timeout /t 2 /nobreak >nul

:: Welcome Message (Displayed for 2 Seconds)
call :Header Instructions
echo - Ensure WhatsApp Beta app is installed on your system.
echo - Keep this window open during execution.
echo - Phone numbers must include country code (e.g., +919876543210).
echo.
timeout /t 3 /nobreak >nul

:: Check WhatsApp Beta Installation
call :Header Requirement 
echo Do you have the WhatsApp Beta app installed? (Y/N)
set /p "has_app=Choice: "
if /i "!has_app!"=="N" (
    call :Header "Install WhatsApp Beta"
    echo Opening Microsoft Store to install WhatsApp Beta...
    start "" "https://apps.microsoft.com/detail/9nbdxk71nk08?hl=en-US&gl=US"
    pause
)

:: WhatsApp Web Login Prompt
call :Header "WhatsApp Login"
echo Opening WhatsApp Web for login...
echo Please scan the QR code with your phone using WhatsApp Beta.
start "" "https://web.whatsapp.com/"
echo.
set /p "login=Press Enter after logging in: "

:: Create VBScript for Sending Messages
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
    echo WScript.Sleep 100
    echo WshShell.SendKeys "{ENTER}"
) > sendmessage.vbs

:: Edit Phone Numbers and Message via Notepad
call :Header "Prepare Your Inputs"
echo Opening Notepad to edit phone numbers (num.txt)...
if not exist num.txt echo. > num.txt
start /wait notepad.exe num.txt
echo Opening Notepad to edit your message (msg.txt)...
if not exist msg.txt echo. > msg.txt
start /wait notepad.exe msg.txt

:: Read Message from msg.txt
set /p "message="<msg.txt

:: Count Total Phone Numbers
for /f %%a in ('find /c /v "" ^< num.txt') do set "total=%%a"
echo.
echo Total phone numbers found: !total!

:: Get Delay Range from User
call :Header "Set Random Delay Range"
echo Please specify the delay range for chat loading.
set /p "min_delay=Enter minimum delay (seconds): "
set /p "max_delay=Enter maximum delay (seconds): "

:: Process Each Phone Number
call :Header "Sending Messages"
set "count=0"
for /f "tokens=*" %%a in (num.txt) do (
    set /a count+=1
    echo Sending message !count! of !total! to %%a
    :: Copy message to clipboard
    echo !message! | clip
    :: Open chat in WhatsApp Web
    start "" "https://wa.me/%%a"
    :: Generate random delay within user-specified range
    set /a "delay=!random! %% (!max_delay! - !min_delay! + 1) + !min_delay!"
    echo Waiting !delay! seconds for chat to load...
    timeout /t !delay! /nobreak >nul
    :: Send message using VBScript
    cscript //nologo sendmessage.vbs
    :: Fixed 7-second delay after sending
    echo Waiting 7 seconds before next message...
    timeout /t 7 /nobreak >nul
)

:: Completion Message
cls
echo Total messages sent: !total!
powercfg /change standby-timeout-ac 30
powercfg /change standby-timeout-dc 10
powercfg /change monitor-timeout-ac 10
powercfg /change monitor-timeout-dc 5
del sendmessage.vbs
:QuitScript
call :Spacer 6
echo      "+----------------------------------------+"
echo      "|                                        |"
ping 127.0.0.1 -n 1 -w 200 >nul
echo      "|   Thank-You-For-Using-WaBulkSender       |"
echo      "|            By-@Parth_Sancheti            |"
echo      "|       Have-A-Productive-Day-Ahead        |"
echo      "|                                        |"
ping 127.0.0.1 -n 1 -w 200 >nul
echo      "+----------------------------------------+"
call :Spacer 6
timeout /t 3 >nul
exit /b

:Header
cls
echo %border_line%
call :Formater
call :CenterText %~1
call :CenterText "Created By: @Parth_Sancheti"
call :Formater
echo %border_line%
call :Formater
goto :EOF

:: Utility Function to Center Text
:CenterText
set "text=%~1"
set /a "console_width=66"
set /a "text_length=0"
for /l %%i in (0,1,4095) do (
    if "!text:~%%i,1!"=="" (
        set /a "text_length=%%i"
        goto :done_length
    )
)
:done_length
set /a "padding=(console_width - text_length) / 2"
set "spaces="
for /l %%i in (1,1,!padding!) do set "spaces=!spaces! "
echo !spaces!!text!
goto :EOF

:Spacer
set i=0
set n=%~1
:SpacerLoop
if !i! lss %n% (
    echo.
    set /a i+=1
    goto SpacerLoop
)
goto :EOF

:Formater
echo.
ping 127.0.0.1 -n 1 -w 200 >nul
goto :EOF

