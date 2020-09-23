Attribute VB_Name = "mod_Polygon"
'Thanks to
''Bouncing Polygon screensaver' by Brian Adriance at PSC txtCodeId=52063.
'
Option Explicit
Private Const PolyName             As String = "PolygonIV"
Public lngTimer                    As Long
Private bDoingReset                As Boolean
Public lngCol                      As Long
Public lngFat                      As Long
Public bLoadingList                As Boolean
Public BSelecting                  As Boolean
Public Motion                      As Long
Public Spinner                     As Long
Public Const SettingHighSmall      As Long = 4035
Public Const SettingHighLarge      As Long = 7020
Public bRndCol                     As Boolean
Public VertMin                     As Long                                       ' min = 1 a dot
Public VertMax                     As Long                                       ' max = 20 could be higher but not likely to give visible differences
Public Type SPOLYGON
  XCenter                          As Long
  YCenter                          As Long
  XVertex(1 To 21)                 As Long                                         ' the upper range must maximum value for VertMax + 1
  YVertex(1 To 21)                 As Long
  DispX                            As Long
  DispY                            As Long
  Mass                             As Double
  Angle                            As Double
  RSpeed                           As Double
  AngleR                           As Double
  Color                            As Long
  NVertex                          As Long
  DispVector                       As Double
  Displacement                     As Double
  IncVert                          As Boolean
  SpinC                            As Boolean
End Type
Public MSpeed                      As Long                                       'at very high speeds(2000+) some of the objects disappear as the distance between draws is so great that they spend most of the time off screen
Public SCount                      As Long                                       'no real limit but above about 250 (on a PII) the redraw cycle is too slow for good animation
Public SetMode                     As Long
Private Polygon()                  As SPOLYGON
Private Const PIDiv180             As Double = 3.14159265358979 / 180
Public Declare Function ShowCursor Lib "user32" (ByVal bShow As Long) As Long                                                                    'mouse visible(0 or 1)

Public Function AtBoundry(poly As SPOLYGON) As Boolean
'detect edge of screen
  With poly
    If .XCenter + (.Mass * 2 + lngFat * 1.5) >= Screen.Width Then
      AtBoundry = True
     ElseIf .YCenter + (.Mass * 2 + lngFat * 1.5) >= Screen.Height Then
      AtBoundry = True
     ElseIf .XCenter - (.Mass * 2 + lngFat * 1.5) <= 0 Then
      AtBoundry = True
     Else
      If .YCenter - (.Mass * 2 + lngFat * 1.5) <= 0 Then
        AtBoundry = True
      End If
    End If
  End With

End Function

Private Function Dist(X1 As Long, _
                      Y1 As Long, _
                      X2 As Long, _
                      Y2 As Long, _
                      Optional Z1 As Long = 0, _
                      Optional Z2 As Long = 0) As Double
'detect distance between 2 points
  If Z1 = 0 Then
    Dist = Sqr((X2 - X1) ^ 2 + (Y2 - Y1) ^ 2)
   Else
    Dist = Sqr((X2 - X1) ^ 2 + (Y2 - Y1) ^ 2 + (Z2 - Z1) ^ 2)
  End If

End Function

Private Sub DoCollision(n As Long)
' check for collisions
  Dim M As Long

  On Error Resume Next
  With Polygon(n)
    Do While .AngleR >= 360
      .AngleR = .AngleR - 360
    Loop
    For M = 1 To SCount
      If M <> n Then
        If Dist(.XCenter, .YCenter, Polygon(M).XCenter, Polygon(M).YCenter) <= (.Mass * 2) + (Polygon(M).Mass * 2 + lngFat * 1.5) Then
          PolygonChangeShape n
          PolygonChangeShape M
        End If
      End If
    Next M
    If .DispVector = 0 Then
      .DispVector = Rnd * MSpeed + 0.1
    End If
  End With
  On Error GoTo 0

End Sub

Public Sub DoPreferences(ByVal intIndex As Integer)

  Dim strPRef As String

  Select Case intIndex
   Case 0
    PreferenceSave
   Case 1
    PreferenceDelete frmSettings.lstPreferences.ListIndex
   Case 2
    PreferenceReNumber
   Case 3
    BSelecting = True
    PreferenceLoad "Random"
    ResetControls strPRef
    PolygonCreateAll
    BSelecting = False
   Case 4
    BSelecting = True
    strPRef = "Rnd User"
    PreferenceLoad strPRef
    ResetControls strPRef
    PolygonCreateAll
    BSelecting = False
  End Select

End Sub

Private Function FatFactor(dblMass As Double) As Long
' apply fat value
  FatFactor = lngFat
  If FatFactor < 1 Then
    FatFactor = 1
  End If

Exit Function

  Select Case lngFat
   Case 0
    FatFactor = 1
   Case 1
    If dblMass / 100 > 1 Then
      FatFactor = dblMass / 10
     Else
      FatFactor = 10
    End If
   Case 2
    If dblMass / 100 > 1 Then
      FatFactor = dblMass / 4
     Else
      FatFactor = 50
    End If
   Case 3
    FatFactor = 150
   Case 3
    FatFactor = 250
  End Select

End Function

Private Function GetDoit() As String
' get the saved start up behaviour
  Dim sets As Variant
  Dim I    As Long

  sets = GetAllSettings(PolyName, "Options")
  If Not IsEmpty(sets) Then
    For I = LBound(sets, 1) To UBound(sets, 1)
      If sets(I, 0) = "Doit" Then
        GetDoit = sets(I, 1)
        Exit Function
      End If
    Next I
  End If

End Function

Private Function GetFreeUser() As Long
' get a number of a new user preference set
  Dim sets As Variant
  Dim I    As Long

  GetFreeUser = 1
  sets = GetAllSettings(PolyName, "Options")
  If Not IsEmpty(sets) Then
    For I = LBound(sets, 1) To UBound(sets, 1)
      If Left$(sets(I, 0), 4) = "User" Then
        If CLng(Mid$(sets(I, 0), InStr(sets(I, 0), " ") + 1)) = GetFreeUser Then
          GetFreeUser = GetFreeUser + 1
        End If
      End If
    Next I
  End If

End Function

Private Sub KeepInRange(ByVal lngMin As Long, _
                        lngVal As Long, _
                        ByVal lngMax As Long)

  If lngVal < lngMin Then
    lngVal = lngMin
   ElseIf lngVal > lngMax Then
    lngVal = lngMax
  End If

End Sub

Public Sub loadInitialSettings()

'This will force a random setting on first usage

  Dim strSet As String

  bLoadingList = True
  strSet = GetSetting(PolyName, "Options", "Doit", "Random")
  PreferenceLoad strSet
  ResetControls strSet
  bLoadingList = False

End Sub

Private Sub Main()

  Select Case Mid$(UCase$(Trim$(Command$)), 1, 2)
   Case "/C"
    SetMode = 1
    PreferenceFillList
    loadInitialSettings
    frmSettings.Height = SettingHighSmall
    frmSettings.Show 1
    SetMode = 0
   Case "", "/S"
    SetMode = 0 '
    On Error Resume Next
    frm_Polygon.Show
  End Select
  On Error GoTo 0

End Sub

Public Function MotionName() As String
'caption for the Motion scrollbar
  Select Case Motion
   Case 1
    MotionName = "Linear"
   Case 2
    MotionName = "L3:B1"
   Case 3
    MotionName = "L:B"
   Case 4
    MotionName = "L1:B3"
   Case 5
    MotionName = "Brownian"
  End Select

End Function

Public Sub MoveBrownian()
'simple brownian (jiggle around current location)

  Dim n    As Long
  Dim divF As Single

  For n = 1 To SCount
    With Polygon(n)
      If AtBoundry(Polygon(n)) Then
        .Angle = .Angle + 15#
        .DispX = .XCenter
        .DispY = .YCenter
        .Displacement = 0
      End If
      If .Mass < 300 Then
        .Mass = .Mass * 1.01
      End If
      If .Mass > 10000 Then
        .Mass = .Mass * 0.9
      End If

      divF = Int(Rnd * MSpeed) + 1
      .XCenter = .XCenter + IIf(Rnd > 0.5, .Mass / divF, -.Mass / divF)
      .YCenter = .YCenter + IIf(Rnd > 0.5, -.Mass / divF, .Mass / divF)
      PolygonSetUp n
      DoCollision n
    End With
  Next n

End Sub

Public Sub MoveLinear()
'move in a straight line
  Dim n As Long

  For n = 1 To SCount
    With Polygon(n)
      If AtBoundry(Polygon(n)) Then
        .Angle = .Angle + 15#
        .DispX = .XCenter
        .DispY = .YCenter
        .Displacement = 0
      End If
      If .Mass < 300 Then
        .Mass = .Mass + 0.025
      End If
      If .Mass > 10000 Then
        .Mass = .Mass * 0.9
      End If
      .Displacement = .Displacement + (.DispVector) + IIf(.SpinC, 4, -4)
      .XCenter = Cos(Rad(.Angle)) * .Displacement + .DispX
      .YCenter = Sin(Rad(.Angle)) * .Displacement + .DispY
      PolygonSetUp n
      DoCollision n
    End With
  Next n

End Sub

Public Sub PolygonChangeShape(ByVal PNum As Long)
'apply changes to object
  Dim I As Long

  On Error Resume Next
  With Polygon(PNum)
    For I = 1 To .NVertex
      .XVertex(I) = 0
      .YVertex(I) = 0
    Next I
    If VertMax <> VertMin Then        'min & max are different
      If .IncVert Then                'due to increase sides
        If .NVertex < VertMax Then    'if not already at max
          .NVertex = .NVertex + 1
          .Mass = .Mass - 1
         Else
          .IncVert = False            'else reverse side increase/decrease
          .NVertex = .NVertex - 1     ' and decrease
          .Mass = .Mass + 1
        End If
       Else                           'due to decrease size
        If .NVertex > VertMin Then    'if not at min
          .NVertex = .NVertex - 1
          .Mass = .Mass - 1
         Else
          .IncVert = True             'else reverse side increase/decrease
          .NVertex = .NVertex + 1     ' and increase
          .Mass = .Mass + 1
        End If
      End If
      .Mass = .Mass * IIf(.IncVert, 1.1, 0.9) ' resize according to side count
     Else                                      'same Min & Max
      .NVertex = VertMax                       ' force to value just in case
      .Mass = .Mass * IIf(Rnd > 0.5, 0.9, 1.1) ' randomly choose size
    End If
    .Angle = Rnd * 361
    .AngleR = .Angle
    .Displacement = 0
    .DispX = .XCenter
    .DispY = .YCenter
    .RSpeed = Rnd * 0.125 - 0.0625
    .DispVector = Rnd * MSpeed + 0.1
    If bRndCol Then
      .Color = RandomColour
    End If
    PolygonSetUp PNum
  End With
  On Error GoTo 0

End Sub

Private Sub PolygonCreate(ByVal num As Long)

'create a new object

  With Polygon(num)
    .IncVert = True
    .XCenter = Int(Rnd * (Screen.Width - 400)) + 200
    .YCenter = Int(Rnd * (Screen.Height - 400)) + 200
    If VertMax = VertMin Then
      .NVertex = VertMax
     Else
      .NVertex = Int(Rnd * (VertMax)) + VertMin
    End If
    .Mass = Int(Rnd * 300) + 100
    .Angle = Rnd * 361
    .AngleR = .Angle
    .Displacement = 0
    .DispX = .XCenter
    .DispY = .YCenter
    .SpinC = Rnd > 0.5
    .RSpeed = Rnd * 0.125 - 0.0625
    .DispVector = Rnd * MSpeed + 0.1
    .Color = RandomColour
  End With

End Sub

Public Sub PolygonCreateAll()

  Dim n As Long

  ReDim Polygon(1 To SCount) As SPOLYGON
  For n = 1 To SCount
    PolygonCreate n
    PolygonSetUp n
  Next n

End Sub

Public Sub PolygonCycle(frm As Form, _
                        ByVal num As Long)
'main loop for the screensaver
  Dim I As Long

  With Polygon(num)
    frm.DrawWidth = lngFat
    For I = 1 To UBound(.XVertex)
      If UBound(.XVertex) = VertMax Then
        If I = VertMax Then
          GoTo MaxSize
        End If
      End If
      If .XVertex(I + 1) <> 0 Then
        frm.Line (.XVertex(I), .YVertex(I))-(.XVertex(I + 1), .YVertex(I + 1)), .Color
       Else
MaxSize:
        frm.Line (.XVertex(I), .YVertex(I))-(.XVertex(1), .YVertex(1)), .Color
        Exit For
      End If
    Next I
  End With

End Sub

Private Sub PolygonSetUp(n As Long)

  Dim M          As Long
  Dim SpinFactor As Double

  On Error Resume Next ' copes with objects that escape the frame by jumping beyound the limits of screen witout triggering the boundry tests
  With Polygon(n)
    For M = 1 To .NVertex
      .XVertex(M) = .Mass * 2 * Cos(Rad(.AngleR)) + .XCenter
      .YVertex(M) = .Mass * 2 * Sin(Rad(.AngleR)) + .YCenter
      SpinFactor = .DispVector * (Spinner / 5) / 100
      If Not .SpinC Then
        SpinFactor = -SpinFactor
      End If
      .AngleR = .AngleR + (360 / .NVertex) + SpinFactor
    Next M
  End With
  If Err.Number <> 0 Then
    PolygonCreate n
  End If
  On Error GoTo 0

End Sub

Private Sub PolygonUpdate(ByVal num As Long)

'change object shape/colour

  With Polygon(num)
    If VertMax = VertMin Then
      .NVertex = VertMax
     Else
      .NVertex = Int(Rnd * (VertMax)) + VertMin
    End If
    .Color = RandomColour
    PolygonChangeShape num
  End With

End Sub

Public Sub PolygonUpdateAll()
' reset only the appearance of objects
  Dim n As Long

  If Not bDoingReset Then
    ReDim Preserve Polygon(1 To SCount) As SPOLYGON
    For n = 1 To SCount
      PolygonUpdate n
      PolygonSetUp n
    Next n
  End If

End Sub

Public Sub PreferenceDelete(ByVal intIndex As Integer)

  Dim sets    As Variant
  Dim I       As Long
  Dim strPRef As String
  Dim lngPos  As Long

  lngPos = frmSettings.lstPreferences.ListIndex
  strPRef = frmSettings.lstPreferences.List(intIndex)
  If strPRef = "Random" Then
    Exit Sub
  End If
  If strPRef = "Random Pref" Then
    Exit Sub
  End If
  sets = GetAllSettings(PolyName, "Options")
  For I = LBound(sets, 1) To UBound(sets, 1)
    If sets(I, 0) = strPRef Then
      If GetDoit = strPRef Then
        SaveSetting PolyName, "Options", "Doit", "Random"
      End If
      DeleteSetting PolyName, "Options", strPRef
      PreferenceReNumber
      PreferenceFillList
      With frmSettings
        If lngPos < .lstPreferences.ListCount Then
          .lstPreferences.ListIndex = -1
          .lstPreferences.ListIndex = lngPos
         Else
          .lstPreferences.ListIndex = 0
        End If
      End With 'frmSettings
      Exit Sub
    End If
  Next I

End Sub

Public Sub PreferenceFillList()

  Dim sets As Variant
  Dim I    As Long
  Dim Prev As Long

  Prev = frmSettings.lstPreferences.ListIndex
  frmSettings.lstPreferences.Clear
  bLoadingList = True
  sets = GetAllSettings(PolyName, "Options")
  frmSettings.lstPreferences.AddItem "Random"
  frmSettings.cmdPreferences(4).Enabled = False
  If Not IsEmpty(sets) Then
    frmSettings.lstPreferences.AddItem "Rnd User"
    frmSettings.cmdPreferences(4).Enabled = True
    For I = LBound(sets, 1) To UBound(sets, 1)
      If sets(I, 0) <> "Doit" Then
        frmSettings.lstPreferences.AddItem sets(I, 0)
       Else
      End If
    Next I
  End If
  If Prev > -1 Then
    If Prev < frmSettings.lstPreferences.ListCount Then
      frmSettings.lstPreferences.ListIndex = Prev
     Else
      frmSettings.lstPreferences.Text = GetDoit
    End If
   Else
    frmSettings.lstPreferences.Text = GetDoit
  End If
  bLoadingList = False

End Sub

Public Sub PreferenceLoad(strName As String)

  Dim sets     As Variant
  Dim I        As Long
  Dim tmpA     As Variant
  Dim bRndUSer As Boolean

  Select Case strName
   Case "Random"
    RandomDisplay
    Exit Sub
   Case Else
    sets = GetAllSettings(PolyName, "Options")
    If strName = "Rnd User" Then
      bRndUSer = True
      SaveSetting PolyName, "Options", "Doit", strName
      strName = PreferenceRandom
      If LenB(strName) = 0 Then
        strName = "Random"
      End If
    End If
    For I = LBound(sets, 1) To UBound(sets, 1)
      If sets(I, 0) = strName Then
        If bRndUSer = False Then
          SaveSetting PolyName, "Options", "Doit", strName
        End If
        tmpA = Split(sets(I, 1), "|")
        SCount = tmpA(0)
        MSpeed = tmpA(1)
        VertMax = tmpA(2)
        VertMin = tmpA(3)
        Spinner = tmpA(4)
        lngFat = tmpA(5)
        lngCol = tmpA(6)
        bRndCol = IIf(tmpA(7) = "True", True, False)
        Motion = tmpA(8)
        lngTimer = tmpA(9)
      End If
    Next I
  End Select

End Sub

Private Function PreferenceRandom() As String

  Dim sets   As Variant
  Dim arrTmp As Variant
  Dim I      As Long
  Dim strTmp As String

  sets = GetAllSettings(PolyName, "Options")
  For I = LBound(sets, 1) To UBound(sets, 1)
    If Left$(sets(I, 0), 5) = "User " Then
      strTmp = strTmp & "|" & sets(I, 0)
    End If
  Next I
  If Len(strTmp) Then
    strTmp = Mid$(strTmp, 2)
    arrTmp = Split(strTmp, "|")
    PreferenceRandom = arrTmp(Int(Rnd * (UBound(arrTmp) + 1)))
  End If

End Function

Public Sub PreferenceReNumber()

  Dim sets      As Variant
  Dim I         As Long
  Dim lngGuard  As Long
  Dim lngNewNum As Long

  sets = GetAllSettings(PolyName, "Options")
  For I = LBound(sets, 1) To UBound(sets, 1)
    If Left$(sets(I, 0), 4) = "Doit" Then
      lngGuard = I
    End If
    If Left$(sets(I, 0), 5) = "User " Then
      lngNewNum = lngNewNum + 1
      If I = lngGuard Then 'ensure that Doit is updated
        sets(lngGuard, 1) = "User " & lngNewNum
      End If
      sets(I, 0) = "User " & lngNewNum
    End If
  Next I
  DeleteSetting PolyName, "Options"
  For I = LBound(sets, 1) To UBound(sets, 1)
    SaveSetting PolyName, "Options", sets(I, 0), sets(I, 1)
  Next I
  PreferenceFillList

End Sub

Public Sub PreferenceSave()

  Dim sets As Variant
  Dim I    As Long

  If frmSettings.lstPreferences.ListCount < 32768 Then ' not likely but you never know
    sets = GetAllSettings(PolyName, "Options")
    GetFreeUser
    For I = LBound(sets, 1) To UBound(sets, 1)
      If sets(I, 0) <> "Doit" Then
        If StorageString = sets(I, 1) Then ' skip duplicates
          MsgBox sets(I, 0) & " already has these settings", vbOKOnly, PolyName
          Exit Sub
        End If
      End If
    Next I
    SaveSetting PolyName, "Options", "User " & GetFreeUser, StorageString
    PreferenceFillList
  End If

End Sub

Public Function Rad(Degrees As Double) As Double

  Rad = Degrees * PIDiv180

End Function

Private Function RandomColour() As Long
' generate a random colour within the various colour sets
  Dim vG As Long

  Select Case lngCol
   Case 0 ' B/W dark
    vG = Int(Rnd * 30) + 16
    RandomColour = RGB(vG, vG, vG)
   Case 1 ' B/W average
    vG = Int(Rnd * 127) + 6
    RandomColour = RGB(vG, vG, vG)
   Case 2 ' B/W bright
    vG = Int(Rnd * 230) + 26
    RandomColour = RGB(vG, vG, vG)
   Case 3 'dull
    RandomColour = RGB(Int(Rnd * 100) + 6, Int(Rnd * 100) + 6, Int(Rnd * 100) + 6)
   Case 4 'pastel
    RandomColour = RGB(Int(Rnd * 100) + 156, Int(Rnd * 100) + 156, Int(Rnd * 100) + 156)
   Case 5
    RandomColour = RGB(Int(Rnd * 200) + 56, 0, 0)
   Case 6
    RandomColour = RGB(0, Int(Rnd * 200) + 56, 0)
   Case 7
    RandomColour = RGB(0, 0, Int(Rnd * 200) + 56)
   Case 8
    RandomColour = RGB(Int(Rnd * 200) + 56, Int(Rnd * 200) + 56, 0)
   Case 9
    RandomColour = RGB(Int(Rnd * 200) + 56, 0, Int(Rnd * 200) + 56)
   Case 10
    RandomColour = RGB(0, Int(Rnd * 200) + 56, Int(Rnd * 200) + 56)
   Case 11
    vG = Int(Rnd * 230) + 26
    RandomColour = RGB(0, vG, vG)
   Case 12
    vG = Int(Rnd * 230) + 26
    RandomColour = RGB(vG, 0, vG)
   Case 13
    vG = Int(Rnd * 230) + 26
    RandomColour = RGB(vG, vG, 0)
   Case 14
    RandomColour = RGB(0, Int(Rnd * 230) + 26, 0)
   Case 15
    RandomColour = RGB(Int(Rnd * 130) + 26, Int(Rnd * 130) + 26, Int(Rnd * 130) + 26)
   Case 16
    RandomColour = RGB(Int(Rnd * 200) + 56, Int(Rnd * 200) + 56, Int(Rnd * 200) + 56)
   Case 17
    RandomColour = RGB(Int(Rnd * 230) + 26, Int(Rnd * 230) + 26, Int(Rnd * 230) + 26)
  End Select

End Function

Public Sub RandomDisplay()

  SaveSetting PolyName, "Options", "Doit", "Random"
  RandomValues
  If SetMode = 0 Then
    PolygonCreateAll
  End If

End Sub

Public Sub RandomTimeShift()
'randomly change one property of polygon sets
'NOTE if the selected action is at the limits of its range no change may occur
  Dim OCount As Long ' used to test wheater a new object needs t be created

  Select Case Int(Rnd * 7)
   Case 0
    OCount = SCount
    SCount = SCount + IIf(Int(Rnd > 0.5), 1, -1)
    ReDim Preserve Polygon(1 To SCount) As SPOLYGON
    If OCount > SCount Then
      PolygonCreate UBound(Polygon)
    End If
    KeepInRange 2, SCount, 250
   Case 1
    MSpeed = MSpeed + IIf(Int(Rnd > 0.5), 1, -1)
    KeepInRange 1, MSpeed, 1000
   Case 2
    VertMin = VertMin + IIf(Int(Rnd > 0.5), 1, -1)
    KeepInRange 1, VertMin, VertMax
   Case 3
    VertMax = VertMax + IIf(Int(Rnd > 0.5), 1, -1)
    KeepInRange VertMin, VertMax, 20
   Case 4
    Spinner = Spinner + IIf(Int(Rnd > 0.5), 1, -1)
    KeepInRange 1, Spinner, 100
   Case 5
    lngFat = lngFat + IIf(Int(Rnd > 0.5), 1, -1)
    KeepInRange 1, lngFat, 200
   Case 6
    lngCol = lngCol + IIf(Int(Rnd > 0.5), 1, -1)
    KeepInRange 0, lngCol, 17
   Case 7
    bRndCol = Rnd > 0.5
  End Select
  ResetControls ""

End Sub

Public Sub RandomValues()
' generate random values for all settings
' note lngTimer and Spinner are skewed to prefer the off position
  If frmSettings.lstPreferences.Text <> "Random" Then
    frmSettings.lstPreferences.Text = "Random"
  End If
  lngTimer = IIf(Rnd > 0.8, Int(Rnd * 6), 0)
  bRndCol = Rnd > 0.5
  lngCol = Int(Rnd * 17)
  lngFat = Int(Rnd * 100) + 1
  SCount = Int(Rnd * 100) + 10
  MSpeed = Int(Rnd * 500) + 30
  VertMax = Int(Rnd * 10) + 3
  Spinner = IIf(Rnd > 0.5, 0, Int(Rnd * 100))
  Do ' this makes sure that VertMin is less than or equal to VertMax
    VertMin = Int(Rnd * 10) + 1
  Loop While VertMin > VertMax
  Motion = Int(Rnd * 5) + 1

End Sub

Public Sub ResetControls(ByVal strName As String)
'resets the controls without causing cascades
  bDoingReset = True
  With frmSettings
    If Not bLoadingList Then
      .lstPreferences.Text = strName
    End If
    .hscSettings(0).Value = SCount
    .hscSettings(1).Value = MSpeed
    .hscSettings(2).Value = VertMin
    .hscSettings(3).Value = VertMax
    .hscSettings(4).Value = Spinner
    .hscSettings(5).Value = Motion
    .hscSettings(6).Value = lngFat
    .hscSettings(7).Value = lngCol
    bDoingReset = False ' turn guard off so final reset fires the changes
    .hscSettings(8) = lngTimer
    .chkRndColour.Value = IIf(bRndCol, vbChecked, vbUnchecked)
  End With

End Sub

Private Function StorageString() As String
'single source makes recoding easier
  StorageString = SCount & "|" & MSpeed & "|" & VertMax & "|" & VertMin & "|" & Spinner & "|" & lngFat & "|" & lngCol & "|" & bRndCol & "|" & Motion & "|" & lngTimer

End Function

':)Roja's VB Code Fixer V1.1.93 (8/03/2004 10:16:13 AM) 42 + 784 = 826 Lines Thanks Ulli for inspiration and lots of code.

