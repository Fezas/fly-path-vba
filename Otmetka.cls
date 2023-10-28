VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Otmetka"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public Number As String
Public Hour As String
Public Minute As String
Public CoordX As Double
Public CoordY As Double
Public Height As String
Public KeyReloadFly As Boolean
Public crv As Curve
Public Sp As SubPath

Public Sub createCurveFly(CoordPathStartX As Double, CoordPathStartY As Double)
    Set crv = createCurve(ActiveDocument)
    Set Sp = crv.CreateSubPath(CoordPathStartX, CoordPathStartY)
End Sub

Public Sub addPathFly(CoordPathX As Double, CoordPathY As Double)
    Sp.AppendLineSegment CoordPathX, CoordPathY, False
End Sub
