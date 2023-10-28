VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FormPositionCircle 
   Caption         =   "Выставить азимутальный круг"
   ClientHeight    =   945
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   OleObjectBlob   =   "FormPositionCircle.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FormPositionCircle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim subDivisions As New Collection ' podrazdeleniya
Dim AzCircle As Shape
Dim currentSdv As Subdivision
Dim ang As Double
Private Sub UserForm_Initialize()
    Dim sdv As Subdivision
    ComboBoxWithSdv.Clear
    ActiveDocument.Unit = cdrMillimeter
    ActiveDocument.ReferencePoint = cdrCenter
    
    LoadSubdivision
    For Each sdv In subDivisions
        ComboBoxWithSdv.AddItem sdv.Name
    Next
End Sub
Private Sub ComboBoxWithSdv_Change()
    Set currentSdv = subDivisions.Item(ComboBoxWithSdv.Value)
    With ActivePage.Layers("Circle")
        .Activate
        .Visible = True
    End With
    Set AzCircle = ActiveLayer.Shapes.FindShape(, , 6849)
    ang = AzCircle.RotationAngle
    If ang <> 0 Then
        AzCircle.Rotate (360 - ang + currentSdv.North)
    Else
        AzCircle.Rotate (currentSdv.North)
    End If
    AzCircle.SetPosition currentSdv.CoordX, currentSdv.CoordY
    ActivePage.Layers("Circle").Editable = False
End Sub

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
                    subDivisions.Add subdiv, subdiv.Name
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

