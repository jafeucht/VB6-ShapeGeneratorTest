VERSION 5.00
Begin VB.Form frmShapes 
   BackColor       =   &H00000000&
   ClientHeight    =   3135
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   4680
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   3135
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Timer tmTimer 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   1800
      Top             =   1320
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileSettings 
         Caption         =   "&Settings..."
      End
      Begin VB.Menu mnuFilePause 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "E&xit"
      End
   End
End
Attribute VB_Name = "frmShapes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Terminate As Integer

Private Sub Angle()
Dim pPoint As Pt, Ang As Currency, Obj As Currency
Dim LLen As Currency, i As Currency
Dim pStart As Pt, j As Long
Dim Cur As Integer, Degrees As Integer
    Ang = 30
    pPoint = SetValues(ScaleWidth / 2, ScaleHeight / 2 - SHeight / 2)
    pStart = pPoint
    LLen = Cosine(Ang) * SHeight
    If AutoClear Then Cls
    Do
        If Terminate Then
            Terminate = False
            Exit Sub
        End If
        Cur = Cur + 1
        Degrees = Sets(Cur).Degrees
        LLen = Cosine(Degrees) * Sets(Cur).Length
        If Cur >= UBound(Sets) - 1 Then Cur = 0
        DoEvents
        i = Cosine(Ang) * LLen
        Obj = Sine(Ang) * LLen
        Line (pPoint.X, pPoint.Y)-(pPoint.X + Obj, pPoint.Y + i)
        pPoint = SetValues(pPoint.X + Obj, pPoint.Y + i)
        Ang = Ang + Abs(Degrees + 180)
        If Ang > 360 Then Ang = Ang - 360
        j = j + 1
        If j > (360 * UBound(Sets)) Then Exit Sub
    Loop Until IsSimilar(pPoint, pStart)
End Sub

Private Function IsSimilar(Point1 As Pt, Point2 As Pt) As Boolean
    IsSimilar = Abs(ScaleX(Point1.Y, 1, 3) - ScaleX(Point2.Y, 1, 3)) < 2 And Abs(ScaleX(Point1.X, 1, 3) - ScaleX(Point2.X, 1, 3)) < 2
End Function

Private Function MakeRandom() As AngData
    MakeRandom.Degrees = Rnd * 360
    MakeRandom.Length = Rnd * (UBound(Sets) * -30) + 5030
End Function

Sub RandomSets()
Dim i As Integer
    For i = 1 To UBound(Sets)
        Sets(i) = MakeRandom
    Next i
End Sub

Private Sub Form_Click()
    NextPhase
End Sub

Private Sub Form_DblClick()
Static Clicked As Boolean
    If Clicked Then End
    Clicked = True
    mnuFile.Visible = False
    Do
        DoEvents
        SetNum Int(Rnd * 19) + 1
        NextPhase
    Loop
End Sub

Private Sub mnuFileExit_Click()
    End
End Sub

Private Sub mnuFileSettings_Click()
    Terminate = True
    frmSettings.Show vbModal
End Sub

Function RandomElement() As Integer
    RandomElement = Int(Rnd * 150) + 105
End Function

Private Sub tmTimer_Timer()
    NextPhase
End Sub

Sub NextPhase()
    ForeColor = RGB(RandomElement, RandomElement, RandomElement)
    RandomSets
    Angle
End Sub
