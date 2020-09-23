VERSION 5.00
Begin VB.Form frm_Polygon 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "PolygonIII"
   ClientHeight    =   3195
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4680
   ForeColor       =   &H0000FF00&
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Timer tmr_Polygon 
      Interval        =   20000
      Left            =   0
      Top             =   0
   End
End
Attribute VB_Name = "frm_Polygon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private bMouseMoved      As Boolean
Private MouseMoveInd     As Long    ' Stops small mouse moves from triggering shutdown; only a large Click&Drag will shutdown

Private Sub Form_Activate()

  ShowCursor False

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)

  bMouseMoved = True

End Sub

Private Sub Form_Load()

  If App.PrevInstance Then
    End
  End If
  Randomize Timer
  PreferenceFillList
  loadInitialSettings
  With frmSettings
    .Move -frmSettings.Width, 0
    .Show , Me
  End With
  Show
  DoEvents
  ShowCursor False
  frm_Polygon.SetFocus
  PolygonCreateAll
  Do While Not bMouseMoved
    Select Case Motion
     Case 1
      MoveLinear
     Case 5
      MoveBrownian
     Case Else
      If Rnd * 5 > Motion Then
        MoveLinear
       Else
        MoveBrownian
      End If
    End Select
    ShowShapes
    DoEvents
  Loop
  Unload Me

End Sub

Private Sub Form_MouseMove(Button As Integer, _
                           Shift As Integer, _
                           X As Single, _
                           Y As Single)

  If Button <> 0 Then
    MouseMoveInd = MouseMoveInd + 1
    If MouseMoveInd >= 5 Then
      bMouseMoved = True
    End If
  End If

End Sub

Private Sub Form_MouseUp(Button As Integer, _
                         Shift As Integer, _
                         X As Single, _
                         Y As Single)

  Select Case Button
   Case 2
    DoPreferences 3
    MouseMoveInd = 0
   Case 1
    ShowCursor 1
    If frmSettings.Left < 0 Then
      frmSettings.Move 0, 0, frmSettings.Width, SettingHighSmall
     Else
      If SetMode = 0 Then
        frmSettings.Move -frmSettings.Width, 0
        frm_Polygon.SetFocus
        ShowCursor False
      End If
    End If
    MouseMoveInd = 0
   Case 4
    DoPreferences 4
    MouseMoveInd = 0
  End Select

End Sub

Private Sub Form_Unload(Cancel As Integer)

  On Error Resume Next
  ShowCursor 1
  Unload frmSettings
  On Error GoTo 0

End Sub

Private Sub ShowShapes()

  Dim n As Long

  On Error Resume Next
'you might comment out the next line for a very different effect
'but not particularly pretty so not offered as an option
  Me.Cls
  For n = 1 To SCount
    PolygonCycle Me, n
  Next n
  On Error GoTo 0

End Sub

Private Sub tmr_Polygon_Timer()

  RandomTimeShift

End Sub

':)Roja's VB Code Fixer V1.1.93 (8/03/2004 10:16:05 AM) 3 + 119 = 122 Lines Thanks Ulli for inspiration and lots of code.

