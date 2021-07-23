VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Turtle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private isPenDown As Boolean
Private x As Double
Private y As Double
Private degree As Double

Sub Forward(length)
    Dim new_x As Double
    Dim new_y As Double
    new_x = x + Sin(Radian(degree)) * length
    new_y = y + Cos(Radian(degree)) * length
    If isPenDown Then
        Call CanvasSheet.Shapes.AddConnector(msoConnectorStraight, x, y, new_x, new_y)
    End If
    x = new_x
    y = new_y
End Sub

Sub PenDown()
    isPenDown = True
End Sub

Sub PenUP()
    isPenDown = False
End Sub

Sub TurnLeft(d)
    degree = degree + d
End Sub

Sub TurnRight(d)
    degree = degree - d
End Sub

Private Sub Class_Initialize()
    With CanvasSheet
        With Cells(.Rows.Count \ 2, .Columns.Count \ 2)
            x = .Left + Application.UsableWidth / 2
            y = .Top + Application.UsableHeight / 2
            Application.Goto .Item(1), True
        End With
    End With
    degree = 90
    CanvasSheet.Clear
End Sub

Function Radian(degree As Double) As Double
    Radian = degree / 45 * Atn(1)
End Function

