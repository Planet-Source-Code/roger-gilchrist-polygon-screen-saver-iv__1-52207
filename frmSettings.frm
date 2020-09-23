VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Begin VB.Form frmSettings 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Polygon III Settings"
   ClientHeight    =   6720
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   4695
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6720
   ScaleWidth      =   4695
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.HScrollBar hscSettings 
      Height          =   135
      Index           =   8
      Left            =   120
      Max             =   6
      TabIndex        =   28
      Top             =   3120
      Width           =   2055
   End
   Begin VB.HScrollBar hscSettings 
      Height          =   135
      Index           =   7
      Left            =   120
      Max             =   17
      TabIndex        =   25
      Top             =   2760
      Width           =   2055
   End
   Begin VB.HScrollBar hscSettings 
      Height          =   135
      Index           =   6
      Left            =   120
      Max             =   200
      Min             =   1
      TabIndex        =   24
      Top             =   2400
      Value           =   1
      Width           =   2055
   End
   Begin RichTextLib.RichTextBox rtbHelp 
      Height          =   2895
      Left            =   0
      TabIndex        =   22
      Top             =   3720
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   5106
      _Version        =   393217
      Enabled         =   -1  'True
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      FileName        =   "C:\downloads2004\04-03\03\Polygon_III_screensaver\Polygon Help.rtf"
      TextRTF         =   $"frmSettings.frx":0000
   End
   Begin VB.CommandButton cmdPreferences 
      Caption         =   "Rnd User"
      Height          =   255
      Index           =   4
      Left            =   3600
      TabIndex        =   21
      Top             =   2280
      Width           =   975
   End
   Begin VB.CommandButton cmdPreferences 
      Caption         =   "Random"
      Height          =   255
      Index           =   3
      Left            =   2640
      TabIndex        =   20
      Top             =   2280
      Width           =   975
   End
   Begin VB.ListBox lstPreferences 
      Height          =   1815
      Left            =   2640
      Sorted          =   -1  'True
      TabIndex        =   19
      Top             =   480
      Width           =   1935
   End
   Begin VB.CommandButton cmdPreferences 
      Caption         =   "Delete"
      Height          =   255
      Index           =   1
      Left            =   3600
      TabIndex        =   17
      Top             =   240
      Width           =   975
   End
   Begin VB.CommandButton cmdPreferences 
      Caption         =   "Save"
      Height          =   255
      Index           =   0
      Left            =   2640
      TabIndex        =   16
      Top             =   240
      Width           =   975
   End
   Begin VB.HScrollBar hscSettings 
      Height          =   150
      Index           =   5
      Left            =   120
      Max             =   5
      Min             =   1
      TabIndex        =   15
      Top             =   2040
      Value           =   1
      Width           =   2055
   End
   Begin VB.HScrollBar hscSettings 
      Height          =   135
      Index           =   4
      Left            =   120
      Max             =   100
      TabIndex        =   13
      Top             =   1680
      Width           =   2055
   End
   Begin VB.CheckBox chkRndColour 
      Caption         =   "Collision Colour Change"
      Height          =   195
      Left            =   2280
      TabIndex        =   11
      Top             =   2760
      Width           =   2295
   End
   Begin VB.CommandButton cmdOther 
      Caption         =   "Help"
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   10
      Top             =   3360
      Width           =   975
   End
   Begin VB.HScrollBar hscSettings 
      Height          =   135
      Index           =   2
      Left            =   120
      Max             =   20
      Min             =   1
      TabIndex        =   9
      Top             =   960
      Value           =   1
      Width           =   2055
   End
   Begin VB.CommandButton cmdOther 
      Caption         =   "Hide Settings"
      Height          =   255
      Index           =   1
      Left            =   1320
      TabIndex        =   7
      Top             =   3360
      Width           =   1695
   End
   Begin VB.CommandButton cmdOther 
      Caption         =   "Exit"
      Height          =   255
      Index           =   0
      Left            =   3600
      TabIndex        =   6
      Top             =   3360
      Width           =   975
   End
   Begin VB.HScrollBar hscSettings 
      Height          =   135
      Index           =   3
      Left            =   120
      Max             =   20
      Min             =   1
      TabIndex        =   5
      Top             =   1320
      Value           =   3
      Width           =   2055
   End
   Begin VB.HScrollBar hscSettings 
      Height          =   135
      Index           =   1
      Left            =   120
      Max             =   1000
      Min             =   1
      TabIndex        =   3
      Top             =   600
      Value           =   1
      Width           =   2055
   End
   Begin VB.HScrollBar hscSettings 
      Height          =   135
      Index           =   0
      Left            =   120
      Max             =   250
      Min             =   2
      TabIndex        =   0
      Top             =   240
      Value           =   2
      Width           =   2055
   End
   Begin VB.Label lblSettings 
      Caption         =   "Rnd Timer Off"
      Height          =   255
      Index           =   8
      Left            =   120
      TabIndex        =   27
      Top             =   2880
      Width           =   1935
   End
   Begin VB.Label lblSettings 
      Caption         =   "Colour set 0"
      Height          =   255
      Index           =   7
      Left            =   120
      TabIndex        =   26
      Top             =   2520
      Width           =   1935
   End
   Begin VB.Label lblSettings 
      Caption         =   "Fat 1"
      Height          =   255
      Index           =   6
      Left            =   120
      TabIndex        =   23
      Top             =   2160
      Width           =   1935
   End
   Begin VB.Label lblSettings 
      Caption         =   "Preferences"
      Height          =   255
      Index           =   9
      Left            =   2640
      TabIndex        =   18
      Top             =   0
      Width           =   1935
   End
   Begin VB.Label lblSettings 
      Caption         =   "Motion "
      Height          =   255
      Index           =   5
      Left            =   120
      TabIndex        =   14
      Top             =   1800
      Width           =   1935
   End
   Begin VB.Label lblSettings 
      Caption         =   "spin (0 -100)"
      Height          =   255
      Index           =   4
      Left            =   120
      TabIndex        =   12
      Top             =   1440
      Width           =   1935
   End
   Begin VB.Label lblSettings 
      Caption         =   "Min points (1 -20)"
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   8
      Top             =   720
      Width           =   1935
   End
   Begin VB.Label lblSettings 
      Caption         =   "Max points (1 -20)"
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   4
      Top             =   1080
      Width           =   1935
   End
   Begin VB.Label lblSettings 
      Caption         =   "Speed (1 - 1000)"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   2
      Top             =   360
      Width           =   2055
   End
   Begin VB.Label lblSettings 
      Caption         =   "Objects (2- 250)"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   1
      Top             =   0
      Width           =   2055
   End
End
Attribute VB_Name = "frmSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub chkRndColour_Click()

  bRndCol = chkRndColour.Value = vbChecked

End Sub

Private Sub cmdOther_Click(Index As Integer)

  Select Case Index
   Case 0
    Unload Me
    ShowCursor 1
    End
   Case 1
    If SetMode = 0 Then
      frmSettings.Move -frmSettings.Width, 0
      frm_Polygon.SetFocus
      ShowCursor False
    End If
   Case 2
    RandomDisplay
   Case 3
    If frmSettings.Height = SettingHighSmall Then
      frmSettings.Height = SettingHighLarge
     Else
      frmSettings.Height = SettingHighSmall
    End If
   Case 4
    PreferenceSave
  End Select

End Sub

Private Sub cmdPreferences_Click(Index As Integer)

  DoPreferences Index

End Sub

Private Sub Form_Load()

  ShowCursor 1
  cmdOther(1).Visible = SetMode = 0

End Sub

Private Sub Form_Unload(Cancel As Integer)

  Me.Visible = False

End Sub

Private Sub hscSettings_Change(Index As Integer)

  Select Case Index
   Case 0
    SCount = LblScrollValue(0, "Objects (2- 250)")
   Case 1
    MSpeed = LblScrollValue(1, "Speed(1-1000)")
   Case 2
    If hscSettings(2).Value >= VertMax Then
      hscSettings(2).Value = VertMax
    End If
    
    VertMin = LblScrollValue(2, "Min points(1-20)")
   Case 3
    If hscSettings(3).Value <= VertMin Then
      hscSettings(3).Value = VertMin
    End If
    VertMax = LblScrollValue(3, "Max points(1-20)")
   Case 4
    Spinner = LblScrollValue(4, "Spin ( 0 - 100)")
   Case 5
    Motion = hscSettings(5).Value
    lblSettings(5).Caption = "Motion:" & MotionName()
   Case 6

    lngFat = LblScrollValue(6, "Fat")
   Case 7
    lngCol = LblScrollValue(7, "Colour Set")
   Case 8
    lblSettings(8).Caption = "Rnd Timer " & IIf(hscSettings(8).Value = 0, " Off", hscSettings(8).Value * 10)
    lngTimer = hscSettings(8).Value
    
    frm_Polygon.tmr_Polygon.Interval = lngTimer * 10000
    frm_Polygon.tmr_Polygon.Enabled = lngTimer > 0
  End Select
  PolygonUpdateAll

End Sub
Function LblScrollValue(Index As Integer, strCap As String) As Long
    lblSettings(Index).Caption = strCap & " " & hscSettings(Index).Value
    LblScrollValue = hscSettings(Index).Value
End Function
Private Sub lstPreferences_Click()

  Dim strPRef As String

  If BSelecting Then
    Exit Sub
  End If
  BSelecting = True
  If bLoadingList = False Then
    strPRef = lstPreferences.List(lstPreferences.ListIndex)
    Select Case strPRef
     Case "Random"
      PreferenceLoad "Random"
     Case "Random Pref"
      strPRef = "Rnd User"
      PreferenceLoad strPRef
     Case Else
      PreferenceLoad strPRef
    End Select
    ResetControls strPRef
    PolygonCreateAll
  End If
  BSelecting = False

End Sub

':)Roja's VB Code Fixer V1.1.93 (8/03/2004 10:16:06 AM) 1 + 124 = 125 Lines Thanks Ulli for inspiration and lots of code.
