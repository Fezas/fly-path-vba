VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FormInsert 
   ClientHeight    =   6570
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4545
   OleObjectBlob   =   "FormInsert.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FormInsert"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


'Option Explicit
Option Compare Text
Dim errorArray() As String
Dim subDivisions As New Collection ' podrazdeleniya
Dim otmetki As New Collection
Dim errorMessages As New Collection
Dim currentSubDivision As Subdivision

Private Sub ComboBoxSubdiv_Change() ' vibor tehniki
    Dim a() As String
    Set currentSubDivision = subDivisions(ComboBoxSubdiv.ListIndex + 1)
    ComboBoxRT.Clear
    a = Split(currentSubDivision.Tech, ",")
    For i = 0 To UBound(a)
        ComboBoxRT.AddItem a(i)
    Next i
    CommandButtonLoad.Enabled = True
End Sub
Private Sub CommandButtonLoad_Click()
    OpenFileAndLoadData
End Sub
Public Sub CommandButtonCreate_Click()
    CreateCirclePoint
End Sub
Sub OpenFileAndLoadData()
    Set errorMessages = New Collection
    sFileName = CorelScriptTools.GetFileBox("All Files (*.*)|*.*", "Select a file", 0)
    errorMessages.Add "Загрузка файла " & sFileName
    errorMessages.Add "***********************************"
    If sFileName <> "" Then
        Dim errorMessage As String
        Dim lineIndex As Integer
        Dim a() As String
        Set otmetki = New Collection
        Open sFileName For Input As #1
            While Not EOF(1)
                Line Input #1, sLine
                lineIndex = lineIndex + 1
                If sLine <> "" Then
                    a = Split(sLine)
                    If ValidateData(a, lineIndex) Then
                            Set otm = New Otmetka
                            otm.Number = a(1)
                            otm.Hour = Left(a(0), 2)
                            otm.Minute = Right(a(0), 2)
                            otm.CoordX = a(2)
                            otm.CoordY = a(3)
                            otm.Height = a(4)
                            otmetki.Add otm
                    End If
                End If
            Wend
        Close #1
        
        If errorMessages.Count = 2 Then
            errorMessages.Add (otmetki.Count) & " записей успешно создано!"
        Else
            errorMessages.Add "***********************************"
            errorMessages.Add "С ошибками :(" & vbCrLf & (otmetki.Count) & " записей успешно создано!"
        End If
        For Each errMsg In errorMessages
            errorMessage = errorMessage + errMsg & vbLf
        Next errMsg
        With textBoxErrors
            .MultiLine = True
            .Text = errorMessage
        End With
        CommandButtonCreate.Enabled = True
        
    Else
        'Wait
    End If
End Sub
Function ValidateData(data() As String, lineIndex As Integer) As Boolean
    ValidateData = True
    If ArrayLen(data) < 5 Then 'examination line with data
        errorMessages.Add "Алярм! Недостаточно значений в строке " & lineIndex
        ValidateData = False
    ElseIf ArrayLen(data) > 5 Then
        errorMessages.Add "Алярм! Слишком много значений в строке или есть пробел в конце строки " & lineIndex
        ValidateData = False
    End If
    If Len(data(0)) > 4 And ValidateNumbersWithLike(data(0)) Then
        errorMessages.Add "Ошибка! Неправильное время " & lineIndex
        ValidateData = False
    End If
    If Len(data(1)) > 5 And ValidateNumbersWithLike(data(1)) Then
        errorMessages.Add "Ошибка! Неправильный номер цели " & lineIndex
        ValidateData = False
    End If
    If Len(data(2)) > 3 And ValidateNumbersWithLike(data(2)) Then
        errorMessages.Add "Ошибка! Неправильный азимут " & lineIndex
        ValidateData = False
    End If
    If Len(data(3)) > 3 And ValidateNumbersWithLike(data(3)) Then
        errorMessages.Add "Ошибка! Неправильная дальность " & lineIndex
        ValidateData = False
    End If
    If Len(data(4)) > 5 And ValidateNumbersWithLike(data(4)) Then
        errorMessages.Add "Ошибка! Неправильная высота " & lineIndex
        ValidateData = False
    End If
End Function

'Array size
Public Function ArrayLen(arr As Variant) As Integer
    ArrayLen = UBound(arr) - LBound(arr) + 1
End Function

Function UniqueTargetCreate() As Collection
    Dim otm As Otmetka
    Set UniqueTarget = New Collection: On Error Resume Next
    For Each otm In otmetki
        Set uniqueOtmetka = New Otmetka
        uniqueOtmetka.Number = otm.Number
        uniqueOtmetka.Minute = "0"
        uniqueOtmetka.CoordX = 0
        uniqueOtmetka.CoordY = 0
        uniqueOtmetka.setReloadFly False
        UniqueTarget.Add uniqueOtmetka, uniqueOtmetka.Number
    Next otm
    Set UniqueTargetCreate = UniqueTarget
End Function
'Create point
Sub CreateCirclePoint()
    Dim nameLayer As String
    Dim pointMinute As Shape, textMinute As Shape, heightMinute As Shape, heightTextMinute As Shape
    Dim arShapes() As Variant
    Dim groupShapes As New ShapeRange
    Dim CoordFlyX As Double, CoordFlyY As Double
    Dim CoordSubdivX As Double, CoordSubdivY As Double
    Dim Azimut As Double, Dalnost As Double, AzimutFly As Double
    Dim otm As Otmetka, otmCurrent As Otmetka
    Dim timeEpisode As Integer
    Dim crvHeight As Curve
    Dim linePath As Shape
    Dim UniqueAir As New Collection
    Set UniqueAir = UniqueTargetCreate
    Const pi = 3.14159265358979
    nameLayer = ComboBoxSubdiv.Value & " " & ComboBoxRT.Value
    ActivePage.CreateLayer (nameLayer)
    Select Case True 'cvet provodki
        Case OptionButton1.Value
            cmyk_1 = 0
            cmyk_2 = 0
            cmyk_3 = 0
            cmyk_4 = 0
        Case OptionButton2.Value
            cmyk_1 = 0
            cmyk_2 = 100
            cmyk_3 = 100
            cmyk_4 = 0
        Case OptionButton3.Value
            cmyk_1 = 100
            cmyk_2 = 100
            cmyk_3 = 100
            cmyk_4 = 100
        Case OptionButton4.Value
            cmyk_1 = 0
            cmyk_2 = 3
            cmyk_3 = 100
            cmyk_4 = 0
        Case OptionButton5.Value
            cmyk_1 = 88
            cmyk_2 = 80
            cmyk_3 = 0
            cmyk_4 = 0
        Case OptionButton6.Value
            cmyk_1 = 0
            cmyk_2 = 75
            cmyk_3 = 10
            cmyk_4 = 0
        Case OptionButton7.Value
            cmyk_1 = 0
            cmyk_2 = 40
            cmyk_3 = 88
            cmyk_4 = 0
    End Select
    
    For Each otm In otmetki
        Azimut = CDbl(otm.CoordX)
        Dalnost = CDbl(otm.CoordY)
        CoordSubdivX = CDbl(currentSubDivision.CoordX)
        CoordSubdivY = CDbl(currentSubDivision.CoordY)
        AzimutFly = Azimut - CDbl(currentSubDivision.North)
        CoordFlyX = Dalnost * Sin(AzimutFly * pi / 180) + CoordSubdivX
        CoordFlyY = Dalnost * Cos(AzimutFly * pi / 180) + CoordSubdivY
        For Each otmCurrent In UniqueAir
            If StrComp(otm.Number, otmCurrent.Number) = 0 Then
                    If otmCurrent.KeyReloadFly = False Then
                        otmCurrent.KeyReloadFly = True
                        otmCurrent.CoordX = CoordFlyX
                        otmCurrent.CoordY = CoordFlyY
                    Else
                        timeEpisode = CInt(otm.Minute) - CInt(otmCurrent.Minute)
                        If timeEpisode > 1 Or timeEpisode = 0 Then
                            otmCurrent.KeyReloadFly = False
                        Else
                            Dim line As Shape
                            Set line = ActiveLayer.CreateLineSegment(otmCurrent.CoordX, otmCurrent.CoordY, CoordFlyX, CoordFlyY)
                            line.Fill.UniformColor.CMYKAssign 0, 0, 0, 0
                            line.Outline.SetPropertiesEx 0.5, OutlineStyles(0), CreateCMYKColor(0, 0, 0, 100), ArrowHeads(0), ArrowHeads(0), cdrFalse, cdrFalse, cdrOutlineButtLineCaps, cdrOutlineMiterLineJoin, 0#, 100, MiterLimit:=5#, Justification:=cdrOutlineJustificationMiddle
                            line.OrderToBack
                        End If
                    End If
                    otmCurrent.Number = otm.Number
                    otmCurrent.Minute = otm.Minute
                    otmCurrent.CoordX = CoordFlyX
                    otmCurrent.CoordY = CoordFlyY
                    Exit For
             End If
        Next otmCurrent
        Set groupShapes = New ShapeRange
        If otm.Height <> "0" Then
            Set heightMinute = ActiveLayer.CreateLineSegment(CoordFlyX, CoordFlyY, CoordFlyX + 10.283901, CoordFlyY + 29.194506)
            heightMinute.Fill.ApplyNoFill
            heightMinute.Outline.SetPropertiesEx 0.3, OutlineStyles(0), CreateCMYKColor(0, 0, 0, 100), ArrowHeads(0), ArrowHeads(0), cdrFalse, cdrFalse, cdrOutlineButtLineCaps, cdrOutlineMiterLineJoin, 0#, 100, MiterLimit:=5#, Justification:=cdrOutlineJustificationMiddle
            Set crvHeight = ActiveDocument.CreateCurve
            With crvHeight.CreateSubPath(CoordFlyX, CoordFlyY)
                .AppendLineSegment CoordFlyX + 10.283901, CoordFlyY + 29.194506
                .AppendLineSegment CoordFlyX + 40.459304, CoordFlyY + 29.194506
            End With
            heightMinute.Curve.CopyAssign crvHeight
            groupShapes.Add heightMinute
            Set heightTextMinute = ActiveLayer.CreateArtisticText(CoordFlyX + 21.392311, CoordFlyY + 30.880202, otm.Height)
            heightTextMinute.Fill.UniformColor.CMYKAssign 0, 0, 0, 100
            heightTextMinute.Outline.SetNoOutline
            groupShapes.Add heightTextMinute
        End If
        Set pointMinute = ActiveLayer.CreateEllipse2(CoordFlyX, CoordFlyY, 2.5)
        pointMinute.Fill.UniformColor.CMYKAssign cmyk_1, cmyk_2, cmyk_3, cmyk_4
        pointMinute.Outline.SetPropertiesEx 0.5, OutlineStyles(0), CreateCMYKColor(0, 0, 0, 100), ArrowHeads(0), ArrowHeads(0), cdrFalse, cdrFalse, cdrOutlineButtLineCaps, cdrOutlineMiterLineJoin, 0#, 100, MiterLimit:=5#, Justification:=cdrOutlineJustificationMiddle
        groupShapes.Add pointMinute
        Set textMinute = ActiveLayer.CreateArtisticText(CoordFlyX + 3, CoordFlyY - 3, otm.Minute, , , , 12)
        textMinute.Fill.UniformColor.CMYKAssign 0, 0, 0, 100
        groupShapes.Add textMinute
        groupShapes.Group

        
    Next otm
End Sub
'validate data
Function ValidateNumbers(str As String) As Boolean
    Dim myRegExp As New RegExp
    Dim aMatch As Match
    Dim colMatches As MatchCollection
    Dim strTest As String
    With myRegExp
        .Global = True
        .Pattern = "^[0-9]*$"
        Set colMatches = .Execute(str)
    End With
    For Each aMatch In colMatches
        a = aMatch.FirstIndex
        b = aMatch.Length
        c = aMatch.Value
    Next aMatch
    If b > 0 Then
        ValidateNumbers = True
    Else
        ValidateNumbers = False
    End If
End Function

Function ValidateNumbersWithLike(str As String) As Boolean
    Dim strTest As String, lenStr As Integer
    lenStr = Len(Trim(str))
    For i = 1 To lenStr
        If Not Mid(str, i, 1) Like "[0-9]" Then
            ValidateNumbersWithLike = False
        Else
            ValidateNumbersWithLike = True
        End If
    Next i
End Function

Private Sub LoadSubdivision()
    Dim a() As String, Tech() As String
    Dim subdiv As Subdivision
    sFileName = Application.ActiveDocument.FilePath & "subdiv.txt"
    If Dir(sFileName) = "" Then
        MsgBox "Нет файла с точками стояния"
        Exit Sub
    Else
        If subDivisions.Count = 0 Then
            Open sFileName For Input As #1
            While Not EOF(1)
                Line Input #1, sLine
                a = Split(sLine)
                Set subdiv = New Subdivision
                subdiv.Name = a(0) & " " & a(1)
                subdiv.CoordX = Val(a(2))
                subdiv.CoordY = Val(a(3))
                subdiv.North = Val(a(4))
                subdiv.TypeSub = a(5)
                subdiv.Tech = a(6)
                subDivisions.Add subdiv
            Wend
        Close #1
        End If
    End If
End Sub
Function FileOrDirExists(ByVal PathName As String) As Boolean
    Dim iTemp As Long
    On Error Resume Next
    iTemp = GetAttr(PathName)
    Select Case Err.Number
    Case Is = 0
        FileOrDirExists = True
    Case Else
        FileOrDirExists = False
    End Select
    On Error GoTo 0
End Function
Private Sub UserForm_Initialize()
    LoadSubdivision
    ComboBoxSubdiv.Clear
    'ActiveDocument.Unit = cdrMillimeter
    For Each sdv In subDivisions
        ComboBoxSubdiv.AddItem sdv.Name
    Next
End Sub
