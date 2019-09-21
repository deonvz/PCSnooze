VERSION 5.00
Begin VB.Form fileselect 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Pc Snooze (Beta)"
   ClientHeight    =   2595
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3810
   Icon            =   "fileselect.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2595
   ScaleWidth      =   3810
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command4 
      Caption         =   "E&xit"
      Height          =   255
      Left            =   2520
      TabIndex        =   6
      Top             =   2280
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "&Media Player 7"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   2280
      Width           =   2175
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&No Program to Shutdown"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   1920
      Width           =   2175
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Select"
      Height          =   255
      Left            =   2520
      TabIndex        =   3
      Top             =   1920
      Width           =   1215
   End
   Begin VB.FileListBox File1 
      Height          =   1650
      Left            =   1920
      TabIndex        =   2
      Top             =   120
      Width           =   1815
   End
   Begin VB.DirListBox Dir1 
      Height          =   1215
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   1695
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1695
   End
   Begin VB.Label programactive 
      Height          =   135
      Left            =   240
      TabIndex        =   7
      Top             =   0
      Visible         =   0   'False
      Width           =   255
   End
End
Attribute VB_Name = "fileselect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()

programactive.Caption = Dir1.Path & "\" & File1.FileName

pcsnooze.Show
Unload Me

End Sub

Private Sub Command2_Click()

programactive.Caption = "none"

pcsnooze.Show
Unload Me


End Sub

Private Sub Command3_Click()

programactive.Caption = "C:\Program Files\Windows Media Player\wmplayer.exe"

pcsnooze.Show
Unload Me

End Sub

Private Sub Command4_Click()

End

End Sub

Private Sub Dir1_Change()

File1.Path = Dir1.Path

End Sub

Private Sub Drive1_Change()

Dim answer As String

On Error GoTo handler
Dir1.Path = Drive1.Drive
Exit Sub

handler:
        answer = MsgBox("Insert Disk", vbRetryCancel, "PC Snooze")
        If answer = vbRetry Then
        Resume
        Else
        Drive1.Drive = "c:"
        End If
        
End Sub

Private Sub Form_Load()

MsgBox "Select program that would be active at Shutdown", vbInformation, "Pc Snooze"

End Sub
