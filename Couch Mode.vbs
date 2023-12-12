Set WshShell = WScript.CreateObject ("WScript.Shell")

' Retrieve the Windows build number
buildNumber = WshShell.RegRead("HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\CurrentBuild")

' Draw input window
strInput = MsgBox("Activate couch mode ?", vbYesNoCancel, "Couch Mode " & buildNumber)
Dim WshShell

' Keystrokes that are executed
Function Payload(activate)

    if activate = true Then
        cmd1 = "Down"
        cmd2 = "Up"
    Else
        cmd1 = "Up"
        cmd2 = "Down"
    End if

    ' Change display scalling
    WshShell.Run "ms-settings:display", 0
    Wscript.Sleep 1000
    WshShell.SendKeys "+{TAB} "
    Wscript.Sleep 500
    WshShell.SendKeys "{TAB 3} {"& cmd1 &" 2}~"
    Wscript.Sleep 500
    ' Change mouse primary click to right
    WshShell.Run "ms-settings:mousetouchpad", 0
    Wscript.Sleep 500
    WshShell.SendKeys "~{"& cmd1 &" 1}~"
    ' Switch audio output
    WshShell.Run "ms-settings:easeofaccess-audio", 0
    Wscript.Sleep 500
    WshShell.SendKeys "~"
    Wscript.Sleep 500
    WshShell.SendKeys "~{"& cmd2 &" 1}~"
    ' Regain focus of settings window
    Wscript.Sleep 500
    WshShell.Run "ms-settings:mousetouchpad", 0
    WshShell.SendKeys "%{F4}"
End Function

' Check if its windows 10 or 11
if InStr (1, buildNumber, "190")=1 Then
    If strInput = 6 Or strInput = 7 Then
        ' If yes is clicked
        If strInput = 6 Then
            Payload(true)
        Else
            Payload(false)
        End if
    End if
' If its windows 11
Elseif InStr(1, buildNumber, "22")=1 Then
    WshShell.Run "ms-settings:display", 0
End if
