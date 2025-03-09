Set WshShell = WScript.CreateObject("WScript.Shell")
Do
    WScript.Sleep 100
    success = WshShell.AppActivate("WhatsApp - Google Chrome")
Loop Until success = True
WScript.Sleep 1000
WshShell.SendKeys "^v"
WScript.Sleep 500
WshShell.SendKeys "{ENTER}"
WScript.Sleep 100
WshShell.SendKeys "{ENTER}"
