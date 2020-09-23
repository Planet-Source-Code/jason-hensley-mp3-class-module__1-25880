VERSION 5.00
Begin VB.Form Form1 
   Caption         =   " MP3 Player Example"
   ClientHeight    =   3495
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3495
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.HScrollBar HScroll1 
      Height          =   135
      Left            =   240
      Max             =   0
      TabIndex        =   14
      Top             =   2040
      Width           =   4215
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Save List"
      Height          =   255
      Left            =   3600
      TabIndex        =   13
      Top             =   3000
      Width           =   855
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Load List"
      Height          =   255
      Left            =   3600
      TabIndex        =   12
      Top             =   2760
      Width           =   855
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   1920
      TabIndex        =   11
      Top             =   2880
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1920
      TabIndex        =   8
      Top             =   2520
      Width           =   1455
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Add MP3"
      Height          =   255
      Left            =   3600
      TabIndex        =   7
      Top             =   2520
      Width           =   855
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Pause"
      Height          =   255
      Left            =   240
      TabIndex        =   6
      Top             =   3000
      Width           =   855
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Stop"
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   2760
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Play"
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   2520
      Width           =   855
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   3720
      Top             =   120
   End
   Begin VB.Frame Frame2 
      Caption         =   "Controls"
      Height          =   1095
      Left            =   120
      TabIndex        =   3
      Top             =   2280
      Width           =   4455
      Begin VB.Label Label2 
         Caption         =   "Position"
         Height          =   255
         Left            =   1080
         TabIndex        =   10
         Top             =   600
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "Duration:"
         Height          =   255
         Left            =   1080
         TabIndex        =   9
         Top             =   240
         Width           =   735
      End
   End
   Begin VB.ListBox List1 
      Height          =   1620
      Left            =   240
      TabIndex        =   2
      Top             =   360
      Width           =   4215
   End
   Begin VB.ListBox List2 
      Height          =   1620
      Left            =   240
      TabIndex        =   1
      Top             =   360
      Width           =   4215
   End
   Begin VB.Frame Frame1 
      Caption         =   "Playlist"
      Height          =   2175
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4455
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim MP3 As New MP3Class

Private Sub Command1_Click()
'Make sure no mp3 is playing
MP3.MP3Stop
'load mp3 and play
MP3.MP3File = List2
MP3.MP3Play
Text1 = MP3.MP3Duration
Timer1.Enabled = True
HScroll1.Max = MP3.MP3DurationInSec
End Sub

Private Sub Command2_Click()
MP3.MP3Stop
Timer1.Enabled = False
End Sub

Private Sub Command3_Click()
If Command3.Caption = "Pause" Then
MP3.MP3Pause: Command3.Caption = "Resume"
Timer1.Enabled = False
Else
Command3.Caption = "Pause"
MP3.MP3Resume
Timer1.Enabled = True
End If
End Sub

Private Sub Command4_Click()
List2.AddItem "c:\mp3 music\ac dc - back in black.mp3"
MP3.ListNoChar List1, List2
End Sub

Private Sub Command5_Click()
List1.Clear: List2.Clear
MP3.MP3OpenPlayList "c:\my documents\mylist.m3u", List2
MP3.ListNoChar List1, List2
End Sub

Private Sub Command6_Click()
MP3.MP3SavePlayList "c:\my documents\mylist.m3u", List2
End Sub

Private Sub Form_Load()

End Sub

Private Sub HScroll1_Scroll()
MP3.MP3ChangePositionTo HScroll1.Value
End Sub

Private Sub List1_Click()
List2.ListIndex = List1.ListIndex
End Sub

Private Sub Timer1_Timer()
HScroll1.Value = MP3.MP3PositionInSec
Text2.text = MP3.MP3Position
End Sub
