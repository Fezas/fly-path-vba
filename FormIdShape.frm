VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FormIdShape 
   Caption         =   "Узнать ID выбранного объекта"
   ClientHeight    =   1065
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   OleObjectBlob   =   "FormIdShape.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FormIdShape"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton1_Click()
    Dim sr As ShapeRange
    Dim s As Shape
    Set sr = ActiveSelectionRange
    For Each s In sr
    MsgBox "Идентификатор выбранного объекта: " & s.StaticID
    Next s
End Sub
