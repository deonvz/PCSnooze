VERSION 5.00
Begin VB.Form pcsnooze 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "PC Snooze"
   ClientHeight    =   1800
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2400
   DrawStyle       =   5  'Transparent
   FillColor       =   &H80000001&
   FillStyle       =   0  'Solid
   BeginProperty Font 
      Name            =   "Courier New"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "pcsnooze.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Palette         =   "pcsnooze.frx":0442
   ScaleHeight     =   1800
   ScaleWidth      =   2400
   StartUpPosition =   2  'CenterScreen
   Begin VB.OptionButton options3 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Log Off"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1440
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   1320
      Width           =   855
   End
   Begin VB.OptionButton options2 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Restart"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1440
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   960
      Width           =   855
   End
   Begin VB.Timer Timer2 
      Interval        =   1000
      Left            =   -120
      Top             =   0
   End
   Begin VB.TextBox txtminute 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   7177
         SubFormatType   =   1
      EndProperty
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   840
      MaxLength       =   2
      TabIndex        =   2
      Top             =   720
      Width           =   375
   End
   Begin VB.OptionButton options1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Shutdown"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1440
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   600
      Width           =   855
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   0
      Top             =   1440
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Caption         =   "Alarm"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   1335
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   1215
      Begin VB.CommandButton Command1 
         Caption         =   "&Apply"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   3
         Top             =   840
         Width           =   975
      End
      Begin VB.TextBox txthour 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   7177
            SubFormatType   =   1
         EndProperty
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   120
         MaxLength       =   2
         TabIndex        =   1
         Top             =   360
         Width           =   375
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         Caption         =   ":"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   480
         TabIndex        =   10
         Top             =   360
         Width           =   255
      End
   End
   Begin VB.Label Label3 
      Caption         =   "?"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2160
      TabIndex        =   11
      Top             =   0
      Width           =   135
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Function"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1440
      TabIndex        =   9
      Top             =   360
      Width           =   735
   End
   Begin VB.Label lbltime 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "HH.mm"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   7177
         SubFormatType   =   4
      EndProperty
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   1440
      TabIndex        =   8
      Top             =   0
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "   Current Time: "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   0
      Width           =   1095
   End
   Begin VB.Menu mnuSystray 
      Caption         =   "mnuSystray"
      Visible         =   0   'False
      Begin VB.Menu mnuRestore 
         Caption         =   "Restore"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
   End
End
Attribute VB_Name = "pcsnooze"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function ExitWindowsEx Lib "user32" (ByVal uFlags As Long, ByVal dwReserved As Long) As Long
Private Declare Function Shell_NotifyIcon Lib "shell32" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, pnid As NOTIFYICONDATA) As Boolean
Private Declare Function SetForegroundWindow Lib "user32" (ByVal hwnd As Long) As Long

'//UDT required by Shell_NotifyIcon API call
Private Type NOTIFYICONDATA
 cbSize As Long             '//size of this UDT
 hwnd As Long               '//handle of the app
 uId As Long                '//unused (set to vbNull)
 uFlags As Long             '//Flags needed for actions
 uCallBackMessage As Long   '//WM we are going to subclass
 hIcon As Long              '//Icon we're going to use for the systray
 szTip As String * 64       '//ToolTip for the mouse_over of the icon.
End Type


'//Constants required by Shell_NotifyIcon API call:
Private Const NIM_ADD = &H0             '//Flag : "ALL NEW nid"
Private Const NIM_MODIFY = &H1          '//Flag : "ONLY MODIFYING nid"
Private Const NIM_DELETE = &H2          '//Flag : "DELETE THE CURRENT nid"
Private Const NIF_MESSAGE = &H1         '//Flag : "Message in nid is valid"
Private Const NIF_ICON = &H2            '//Flag : "Icon in nid is valid"
Private Const NIF_TIP = &H4             '//Flag : "Tip in nid is valid"
Private Const WM_MOUSEMOVE = &H200      '//This is our CallBack Message
Private Const WM_LBUTTONDOWN = &H201    '//LButton down
Private Const WM_LBUTTONUP = &H202      '//LButton up
Private Const WM_LBUTTONDBLCLK = &H203  '//LDouble-click
Private Const WM_RBUTTONDOWN = &H204    '//RButton down
Private Const WM_RBUTTONUP = &H205      '//RButton up
Private Const WM_RBUTTONDBLCLK = &H206  '//RDouble-click
Private nid As NOTIFYICONDATA

Dim programtime
Dim flicker As Boolean
Dim flagger As Boolean

Private Sub Form_Activate()
'//////////////////////////////////////////////////////////////////
'//Purpose:         Load up the UDT for the Systray Function.  This
'//                 must be done after the form is fully visable.
'//                 The Form_Activate is a perfect place for that.
'//////////////////////////////////////////////////////////////////
 
  With nid
    .cbSize = Len(nid)
    .hwnd = Me.hwnd
    .uId = vbNull
    .uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
    .uCallBackMessage = WM_MOUSEMOVE
    .hIcon = Me.Icon
    .szTip = "PCSnooze" & vbNullChar
  End With
 
  Shell_NotifyIcon NIM_ADD, nid
End Sub

Private Sub Form_Resize()
'//////////////////////////////////////////////////////////////////
'//Purpose:         This is just to check to make sure that, if
'//                 indeed the application is minimized (hence on
'//                 the systray) to also hide the form.
'//////////////////////////////////////////////////////////////////
  If (Me.WindowState = vbMinimized) Then Me.Hide
  
End Sub



Private Sub mnuRestore_Click()
'//////////////////////////////////////////////////////////////////
'//Purpose:         When the application is minimized on the systray
'//                 this will restore it.
'//////////////////////////////////////////////////////////////////
  Me.WindowState = vbNormal
  Call SetForegroundWindow(Me.hwnd)
  Me.Show
End Sub
Private Sub mnuExit_Click()
'//////////////////////////////////////////////////////////////////
'//Purpose:         When the application is minimized on the systray
'//                 this will close the application.
'//////////////////////////////////////////////////////////////////
   Shell_NotifyIcon NIM_DELETE, nid
   Set pcsnooze = Nothing
  Unload Me
End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, _
                                X As Single, Y As Single)
'//////////////////////////////////////////////////////////////////
'//Purpose:         This is the callback function of icon in the
'//                 system tray.  This is where will will process
'//                 what the application will do when Mouse Input
'//                 is given to the icon.
'//
'//Inputs:          What Button was clicked (this is button & shift),
'//                 also, the X & Y coordinates of the mouse.
'//////////////////////////////////////////////////////////////////

  Dim msg As Long     '//The callback value
  
  '//The value of X will vary depending
  '//upon the ScaleMode setting.  Here
  '//we are using that fact to determine
  '//what the value of 'msg' should really be
  If (Me.ScaleMode = vbPixels) Then
    msg = X
  Else
    msg = X / Screen.TwipsPerPixelX
  End If

  Select Case msg
    Case WM_LBUTTONDBLCLK    '515 restore form window
      Me.WindowState = vbNormal
      Call SetForegroundWindow(Me.hwnd)
      Me.Show
      
    Case WM_RBUTTONUP        '517 display popup menu
      Call SetForegroundWindow(Me.hwnd)
      Me.PopupMenu Me.mnuSystray
    
    Case WM_LBUTTONUP        '514 restore form window
      '//commonly an application on the
      '//systray will do nothing on a
      '//single mouse_click, so nothing
  End Select

  '//small note:  I just learned that when using a Select Case
  '//structure you always want to place the most commonly anticipated
  '//action highest. Saves CPU cycles becuase of less evaluations.
End Sub

Private Sub Command1_Click()

'---- Change Button Caption --

If Command1.Caption = "&Apply" Then
    Call checkhours
    Call checkminutes
    Timer1.Enabled = True
    Command1.Caption = "&Cancel"
        ElseIf Command1.Caption = "&Cancel" Then
        Timer1.Enabled = False
        Command1.Caption = "&Apply"
        txthour.Enabled = True
        txtminute.Enabled = True
End If

' ===================

End Sub

Private Sub Form_Unload(Cancel As Integer)
'//////////////////////////////////////////////////////////////////
'//Purpose:         Deletes the systray icon, and makes the application
'//                 "safe" to unload.
'//////////////////////////////////////////////////////////////////
   Shell_NotifyIcon NIM_DELETE, nid
   Set pcsnooze = Nothing
End Sub
Private Sub Form_Load()

options1.Value = 1
txthour.Text = Format(Time, "hh")
txtminute.Text = Format(Time, "nn")

flicker = True

If Format(Date, "dd mm") = "01 11" Then
MsgBox ("It`s my Birthday !!")
End If



End Sub


Private Sub Label3_Click()

MsgBox ("Created by Deon van Zyl (pcsnooze@webmail.co.za)"), vbInformation, "PC Snooze ver 1.1"


End Sub

Private Sub Timer1_Timer()

programtime = txthour.Text & ":" & txtminute.Text

' Select a Function to do

If flagger = False Then
    If Format(Time, "hh:mm") = programtime Then
           flagger = True
        If options1.Value = True Then
         X = ExitWindowsEx(1, 0)
        ElseIf options2.Value = True Then
         X = ExitWindowsEx(2, 0)
        ElseIf options3.Value = True Then
         X = ExitWindowsEx(0, 0)
        End If
        
      End
    End If
End If

End Sub


Private Sub Timer2_Timer()

If flicker = True Then
    lbltime.Caption = Format(Time, "hh mm")
    flicker = False
Else
    lbltime.Caption = Format(Time, "hh:mm")
    flicker = True
End If


End Sub

Private Sub checkhours()

    Call fill_missing_field

If txthour.Text >= 24 Or txthour.Text < 0 Then
    txthour.Text = 23
End If

End Sub

Private Sub checkminutes()

    Call fill_missing_field

If txtminute.Text >= 60 Or txtminute.Text < 0 Then
    txtminute.Text = 59

End If

End Sub


Private Sub txthour_Change()

Command1.Enabled = True ' Enable apply button


End Sub

Private Sub txtminute_Change()

Command1.Enabled = True ' Enable apply button

End Sub

Public Sub fill_missing_field()

' Check if a field is blank

If txthour.Text = "" Then
    txthour.Text = "00"
ElseIf txtminute.Text = "" Then
    txtminute.Text = "00"
End If

' Check if the value is less than 10

If IsNumeric(txthour.Text) = False Then

    txthour.Text = "00"

End If

If IsNumeric(txtminute.Text) = False Then

    txtminute.Text = "00"

End If

If (Len(txthour.Text) < 2) Then

    txthour.Text = "0" & txthour.Text

End If

If (Len(txtminute.Text) < 2) Then

    txtminute.Text = "0" & txtminute.Text

End If

' =========

End Sub
