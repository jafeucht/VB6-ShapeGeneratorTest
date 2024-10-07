Attribute VB_Name = "mdShapes"
Option Explicit

Public Type AngData
    Degrees As Integer
    Length As Long
End Type

Public Type Pt
    X As Currency
    Y As Currency
End Type

Const pi = 3.14159265358979

Public AutoClear As Boolean
Public Sets() As AngData
Public SHeight As Long

Sub Main()
    Randomize Timer
    SetNum (3)
    frmShapes.Show
End Sub

Public Function SetValues(X As Long, Y As Long) As Pt
    SetValues.X = X
    SetValues.Y = Y
End Function

Sub SetNum(NewNum As Integer)
    ReDim Sets(1 To NewNum)
End Sub

Public Function Sine(ByVal i As Double) As Double
    Sine = Sin(i * (pi / 180))
End Function

Public Function Cosine(ByVal i As Double) As Double
    Cosine = Cos(i * (pi / 180))
End Function

Public Function ISine(ByVal i As Double) As Double
    On Error Resume Next
    ISine = 90
    ISine = Atn(-i / Sqr(-i * i + 1)) + 2 * Atn(1) * 180 / pi
End Function

Public Function ICosine(ByVal i As Double) As Double
    ICosine = Atn(-i / Sqr(-i * i + 1)) + 2 * Atn(1) * 180 / pi
End Function
