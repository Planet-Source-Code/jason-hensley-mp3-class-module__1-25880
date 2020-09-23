VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MP3Class"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'
Private Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long
'Used to store the mp3 filename
Public MP3File As String
Dim TheFile As String
'Used to find the last backslash of the file
Public Function GetLastBackSlash(text As String) As String
    Dim i, pos As Integer
    Dim lastslash As Integer


    For i = 1 To Len(text)
        pos = InStr(i, text, "\", vbTextCompare)
        If pos <> 0 Then lastslash = pos
    Next i
    GetLastBackSlash = Right(text, Len(text) - lastslash)
End Function
'Take the path and .mp3 off the file
Public Sub ListNoChar(List1 As ListBox, List2 As Control)
Dim X
Dim NoChar As String
Dim NoEnd As String
For X = 0 To List2.ListCount - 1
NoChar = GetLastBackSlash(List2.List(X))
NoEnd = RightLeft(NoChar, ".")
'NoEnd = Mid(NoChar, 1, 1)
List1.AddItem NoEnd
Next X
End Sub
Private Function RightLeft(source As String, token As String) As String
Dim i As Integer
RightLeft = ""
'
For i = Len(source) To 1 Step -1
'
If Mid(source, i, 1) = token Then
RightLeft = Left(source, i - 1)
Exit Function
End If
Next i
End Function
'Take the .mp3 off the end of file
Private Function NoEndChar(List1 As ListBox, List2 As ListBox) As String
Dim N
'Dim NoEndChar As String
For N = 0 To List2.ListCount - 1
NoEndChar = Left(List2.List(N), 1)
Next N
End Function
'Convert seconds to minutes
Private Function SecondsToMinutes(Secs As Long) As String
 On Error Resume Next
    Dim mins As Integer
    mins = Int(Secs / 60)
    Secs = Secs Mod 60
    SecondsToMinutes = mins & ":" & Format(Secs, "0#")
End Function
'Check to see if the mp3 is playing or not
Public Function MP3Playing() As Boolean
On Error Resume Next
Static s As String * 30
mciSendString "status " & TheFile & " mode", s, Len(s), 0
MP3Playing = (Mid$(s, 1, 7) = "playing")
End Function
'Change the mp3 pistion to a certain position
Public Function MP3ChangePositionTo(Second)
On Error Resume Next
TheFile$ = Chr$(34) + Trim(MP3File$) + Chr$(34)
Second = Second * 1000
mciSendString "set MP3Play time format milliseconds", 0, 0, 0
If MP3Playing = True Then mciSendString "play " & TheFile & " from " & Second, 0, 0, 0
If MP3Playing = False Then mciSendString "seek " & TheFile & " to " & Second, 0, 0, 0
End Function
'Get the current mp3 position in seconds
Public Function MP3PositionInSec()
On Error Resume Next
TheFile$ = Chr$(34) + Trim(MP3File$) + Chr$(34)
Static s As String * 30
    mciSendString "set " & TheFile & " time format milliseconds", 0, 0, 0
    mciSendString "status " & TheFile & " position", s, Len(s), 0
    MP3PositionInSec = Round(Mid$(s, 1, Len(s)) / 1000)
    If MP3PositionInSec <= 9 Then
    MP3PositionInSec = "0" & MP3PositionInSec
    End If
End Function
'Get the current mp3 position in Milli-sec.
Public Function MP3PositionInMS()
On Error Resume Next
Static s As String * 30
TheFile$ = Chr$(34) + Trim(MP3File$) + Chr$(34)
mciSendString "set " & TheFile & " time format milliseconds", 0, 0, 0
mciSendString "status " & TheFile & " position", s, Len(s), 0
MP3PositionInMS = Val(s)
End Function
'Get the formatted duration of the mp3
Public Function MP3Duration()
On Error Resume Next
Dim TotalTime As String * 128
Dim T As String
Dim lTotalTime As Long
TheFile$ = Chr$(34) + Trim(MP3File$) + Chr$(34)
    mciSendString "set " & TheFile & " time format ms", TotalTime, 128, 0&
    mciSendString "status " & TheFile & " length", TotalTime, 128, 0&

    mciSendString "set " & TheFile & " time format frames", 0&, 0&, 0&
    
    lTotalTime = Val(TotalTime)
   T = GetThisTime(lTotalTime)
    MP3Duration = T
End Function
'Get the duration of the mp3 in seconds
Public Function MP3DurationInSec()
On Error Resume Next
Dim TotalTime As String * 128
Dim T As String
TheFile$ = Chr$(34) + Trim(MP3File$) + Chr$(34)
    mciSendString "set " & TheFile & " time format ms", TotalTime, 128, 0&
    mciSendString "status " & TheFile & " length", TotalTime, 128, 0&
    mciSendString "set " & TheFile & " time format frames", 0&, 0&, 0&
    MP3DurationInSec = Round(Mid$(TotalTime, 1, Len(TotalTime)) / 1000)
End Function
'Get the mp3 duration in Milli-sec.
Public Function MP3DurationInMs()
On Error Resume Next
Dim TotalTime As String * 128
Dim T As String
 TheFile$ = Chr$(34) + Trim(MP3File$) + Chr$(34)
    mciSendString "set " & TheFile & " time format ms", TotalTime, 128, 0&
    mciSendString "status " & TheFile & " length", TotalTime, 128, 0&
    mciSendString "set " & TheFile & " time format frames", 0&, 0&, 0&
    DurationInMs = Val(TotalTime)
End Function
Private Function GetThisTime(ByVal timein As Long) As String
    On Error Resume Next
    Dim conH As Integer
    Dim conM As Integer
    Dim conS As Integer
    Dim remTime As Long
    Dim strRetTime As String
    TheFile$ = Chr$(34) + Trim(MP3File$) + Chr$(34)
    remTime = timein / 1000
    conH = Int(remTime / 3600)
    remTime = remTime Mod 3600
    conM = Int(remTime / 60)
    remTime = remTime Mod 60
    conS = remTime
    
    If conH > 0 Then
        strRetTime = Trim(Str(conH)) & ":"
    Else
        strRetTime = ""
    End If
    
    If conM >= 10 Then
        strRetTime = strRetTime & Trim(Str(conM))
    ElseIf conM > 0 Then
        strRetTime = strRetTime & "0" & Trim(Str(conM))
    Else
        strRetTime = strRetTime & "00"
    End If
    
    strRetTime = strRetTime & ":"
    
    If conS >= 10 Then
        strRetTime = strRetTime & Trim(Str(conS))
    ElseIf conS > 0 Then
        strRetTime = strRetTime & "0" & Trim(Str(conS))
    Else
        strRetTime = strRetTime & "00"
    End If
    
    GetThisTime = strRetTime
End Function
'Get the formatted mp3 position
Public Function MP3Position()
Dim Sec As Long
Dim mins As Long
On Error Resume Next
Static s As String * 30
TheFile$ = Chr$(34) + Trim(MP3File$) + Chr$(34)
    mciSendString "set " & TheFile & " time format milliseconds", 0, 0, 0
    mciSendString "status " & TheFile & " position", s, Len(s), 0
    Sec = Round(Mid$(s, 1, Len(s)) / 1000)
    If Sec < 60 Then MP3Position = "0:" & Format(Sec, "00")
    If Sec > 59 Then
        mins = Int(Sec / 60)
        Sec = Sec - (mins * 60)
        MP3Position = Format(mins, "00") & ":" & Format(Sec, "00")
    End If
End Function
'Open and load a .m3u playlist
Public Function MP3OpenPlayList(TheFile As String, TheListBox)
Dim test As String
On Error Resume Next
If TheFile = "" Then Exit Function
Open TheFile For Input As 1
While Not EOF(1)
Line Input #1, test
TheListBox.AddItem RTrim(test)
Wend
Close 1
End Function
'Save a playlist in .m3u format
Public Function MP3SavePlayList(TheFile As String, TheListBox)
Dim i As Integer
Dim a As String
On Error Resume Next
Open TheFile For Output As #1
For i = 0 To TheListBox.ListCount - 1
a$ = TheListBox.List(i)
Print #1, a$
Next
Close 1
End Function
'Start playing the mp3
Sub MP3Play()
'On Error Resume Next
mciSendString "close " & TheFile, 0&, 0&, 0&
TheFile$ = Chr$(34) + Trim(MP3File$) + Chr$(34)
mciSendString "open " & TheFile, 0&, 0&, 0&
mciSendString "play " & TheFile, "", 0&, 0&
End Sub
'Stop the mp3 from playing
Sub MP3Stop()
On Error Resume Next
TheFile$ = Chr$(34) + Trim(MP3File$) + Chr$(34)
mciSendString "close " & TheFile, 0&, 0&, 0&
End Sub
'Resume the mp3 if paused
Sub MP3Resume()
On Error Resume Next
TheFile$ = Chr$(34) + Trim(MP3File$) + Chr$(34)
mciSendString "play " & TheFile, "", 0&, 0&
End Sub
'Pause the mp3 if playing
Sub MP3Pause()
On Error Resume Next
TheFile$ = Chr$(34) + Trim(MP3File$) + Chr$(34)
Call mciSendString("Stop " & TheFile, 0, 0, 0)
End Sub
