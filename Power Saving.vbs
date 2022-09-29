'  Power Saving Protocol V1
' Author Brandon Roff
'
' Prompts a box to ask user to shutdown but allow rejection
' Code VB
'
' Date of Birth 21/09/22

Option Explicit

Dim objShell, initShutdown
Dim strShutdown, strAbort, wshShell, strRegValue, strHostname

Set wshShell = CreateObject ( "WScript.Shell" )

'Set Values For Script

strRegValue = "HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Services\Tcpip\Parameters\Hostname"
strHostname = wshShell.RegRead ( strRegValue )

'Echo Hostname for Debugging Reasons
'WScript.Echo "Hostname: " & strHostname'

if strHostname = "LAP123" OR strHostname = "SERVER123" then
    'Debug Match if Hostame equals above.
    WScript.Echo "Match Found for: " & strHostname'    
    Wscript.Quit

Else

    'Confirm Device is not on elcusion list, Debig Only,
    'Wscript.Echo "No Match Found of: " & strHostname "Continuing"'

    ' -s =Shutdown, -t 600 = 10 Minutes, -f = force' 

    strShutdown = "Shutdown.exe -s -t 600 -f"
    set objShell = CreateObject ( "WScript.Shell" )
    objShell.Run strShutdown, 0, false

    ' go to sleep so message apears at top of the screen'
    Wscript.Sleep 1000

    'Input Box to abort Shutdown'
    initShutdown = ( MsgBox ("This Computer will automatically shutdown in 10 minutes, Do you wish to cancel the shutdown?", vbYesNo+vbExclamation+vbSystemModal, "Automatic Shutdown"))

    if initShutdown = vbYes then
        'Abort Shutdown'
        strAbort = "Shutdown.exe -a"
        set objShell = CreateObject ( "Wscript.Shell" )
        objShell.Run strAbort, 0, false
        Wscript.Quit
    Elseif initShutdown = vbNo then
        MsgBox "Make Sure to save your Work you have 10 minutes before shutdown, there is no further warning.",0, "Shutdown Confirmed"
    End if
    Wscript.Quit
End if

