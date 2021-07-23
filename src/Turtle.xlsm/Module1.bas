Attribute VB_Name = "Module1"
Sub TurtleSample()
    Dim t As New Turtle
    For i = 1 To 36
        t.PenDown
        Call Square(t, i * 5)
        t.PenUP
        t.TurnRight 10
        t.Forward 10
    Next
End Sub

Sub Square(t As Turtle, size)
    For i = 1 To 4
        t.Forward CDbl(size)
        t.TurnRight 90
    Next
End Sub

