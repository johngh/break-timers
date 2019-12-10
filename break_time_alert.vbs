Dim objArgs, objShell, messageText, messageTitle, result
Set objArgs = WScript.Arguments
Set objShell = CreateObject ("WScript.Shell")

Dim time_now, title, line_1, line_2
title = Array("Stop and think...","Time!","Hold on a minute...","Halt!","Whoa, hold up!","Ring... Ring... Ring...","Beep! Beep! Beep!","Freedom!")
line_1 = Array("Time for a break...","Another hour has passed...","Think of your heart... and mind.","Is that the time already?","Eleven! Eleven!","Take some time off")
line_2 = Array("Have a break (no KitKat for you though!)","Stand up and have a walk.","Stretch your legs...","Step away from the keyboard","Can braincells regenerate?")

dim dtmValue
dtmValue = Now()
TimeString = CStr(Year(dtmValue)) _
    & "-" & Right("0" & CStr(Month(dtmValue)), 2) _
    & "-" & Right("0" & CStr(Day(dtmValue)), 2) _
    & " " & Right("0" & CStr(Hour(dtmValue)), 2) _
    & ":" & Right("0" & CStr(Minute(dtmValue)), 2)

Randomize

Function RandomWithinRange(min, max)
    RandomWithinRange = Int((max - min + 1) * Rnd() + min)
End Function

Function rnd_str(arr)
    rnd_str = arr(RandomWithinRange(LBound(arr), UBound(arr)))
End Function

messageTitle = TimeString & " " & rnd_str(title)
messageText = rnd_str(line_1) & vbCrLf & rnd_str(line_2)
'' result = MsgBox (messageText, vbOKOnly + vbExclamation, messageTitle)
result = objShell.Popup(messageText,60,messageTitle, 48)
'' Popup: https://www.vbsedit.com/html/f482c739-3cf9-4139-a6af-3bde299b8009.asp
