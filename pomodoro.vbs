'Author: Ilari Aarnio
'Date: 2014-Oct-02
'
'A simple pomodoro timer
'
'Description:
'  The script aims to help you on your pomodoro sprints by giving you a non-disturbing timer.
'
'  Starting the script activates the timer. There are no bells or whistles or running counters.
'  Just a brief popup windows informs you that pomodoro time has begun. The popup will auto-close too.
'  After that the script stays out of your way until your pomodoro time is finished.
'  Then a new popup window is displayed and options are given to finish or snooze the pomodoro.
'
'Usage:  wscript.exe pomodoro.vbs [y]
'  (or just double click the icon if the default app for .vbs is MS Script Host)
'  Optional parameter y or yes enables sound alarm when popup appears.
'
'For details on the Pomodoro technique, please refer http://en.wikipedia.org/wiki/Pomodoro_Technique

Option Explicit

'Configure these for your likings
Dim alarm, minute, pomMinutes, pomodoroTime, snoozeTime, soundFile, alarmPlayer
pomMinutes = 25
alarm        = False
minute       = 60 * 1000
pomodoroTime = pomMinutes * minute
snoozeTime   =  5 * minute

Dim lockFile
lockFile = "c:\john\bin\pomodoro.lock"

Const DateLastModified = 1

dim objShell, fso

Set objShell = CreateObject("WScript.Shell")
Set alarmPlayer = CreateObject("WMPlayer.OCX")
Call Main
alarmPlayer.close
Set objShell = Nothing

Function FileExists(FilePath)
  Set fso = CreateObject("Scripting.FileSystemObject")
  If fso.FileExists(FilePath) Then
    FileExists=CBool(1)
  Else
    FileExists=CBool(0)
  End If
End Function

Sub CreateLockFile
   Dim fso, MyFile
   Set fso = CreateObject("Scripting.FileSystemObject")
   Set MyFile = fso.CreateTextFile(lockFile, True)
   MyFile.WriteLine("Pomodoro is running...")
   MyFile.Close
End Sub

Sub DeleteLockFile
   Dim fso
   Set fso = CreateObject("Scripting.FileSystemObject")
   if fso.FileExists(lockFile) then
      fso.DeleteFile lockFile
   end if
End Sub

Sub Main

    Call CheckArgs

    If alarm Then
        Set WshSysEnv = objShell.Environment("PROCESS")
        soundFile = WshSysEnv("WINDIR") & "\Media\notify.wav"  'Alarm ringtone
        Set WshSysEnv = Nothing
    End If

    If FileExists(lockFile) Then

        Dim LastModified, FSO, DateDifference, TimeLeft

        Set FSO = CreateObject("Scripting.FileSystemObject")

        LastModified = FSO.GetFile(lockFile).DateLastModified

        DateDifference = DateDiff("n",LastModified, Now()) + 1

        If DateDifference < pomMinutes Then
            TimeLeft = pomMinutes - DateDifference
            WScript.Echo ("You have " & TimeLeft & " minutes left.")
        Else
            WScript.Echo ( "Found old file: " & lockFile & " - Is Pomodoro already running. Exiting" )
        End If

        WScript.Quit(1)

    End If

    Dim messageText
    messageText = "Start Pomodoro session?" 

    Select Case MsgBox (messageText, vbYesNo + vbQuestion, "Pomodoro")
    Case vbYes
        CreateLockFile
    Case vbNo
        WScript.Quit(0)
    End Select

    Dim dummy
    dummy = objShell.Popup("Pomodoro started", 2, "", 0 + 64 + 4096)
    WScript.Sleep pomodoroTime
    Call Finished
End Sub

'Check if 'y' or 'yes' is provided to enable alarm
Sub CheckArgs
    dim args, arg
    Set args = WScript.Arguments
    For Each arg In args
        If Lcase(arg) = "y" Or Lcase(arg) = "yes" Then
            alarm = True
        End If
    Next
End Sub

'Pomodoro time done - give options to finish or snooze
Sub Finished
    Dim btnCode, title, postponed
    Do Until btnCode = 6
        If alarm Then
            Call PlayAlarm
        End If
        title = "Time's up"
        If postponed > 0 Then
            title = title + "! Postponed " + Cstr(postponed) + " time"
            If postponed > 1 Then
                title = title + "s"'
            End If
        End If
        btnCode = objShell.Popup("Pomodoro finished?", 0, title, 4 + 32 + 4096)
        If btnCode = 7 Then
            postponed = postponed + 1
            WScript.Sleep snoozeTime
        End If
    Loop
    If FileExists(lockFile) Then
       DeleteLockFile
    Else
       WScript.Echo lockFile & " not found"
    End If
End Sub

'Play alarm sound
Sub PlayAlarm
    alarmPlayer.URL = soundFile
End Sub

