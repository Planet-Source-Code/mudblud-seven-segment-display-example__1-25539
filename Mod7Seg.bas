Attribute VB_Name = "mod7Seg"
Type Linenfo
x As Integer
y As Integer
h As Boolean
End Type
Const ONColor = vbGreen
Const OFFColor = &H4000&
Public Declare Function SetPixel Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal crColor As Long) As Long
Function LineInfo(Index) As Linenfo
If Index = 1 Then LineInfo.x = 5: LineInfo.y = 1: LineInfo.h = True
If Index = 2 Then LineInfo.x = 17: LineInfo.y = 2: LineInfo.h = False
If Index = 3 Then LineInfo.x = 17: LineInfo.y = 14: LineInfo.h = False
If Index = 4 Then LineInfo.x = 5: LineInfo.y = 23: LineInfo.h = True
If Index = 5 Then LineInfo.x = 1: LineInfo.y = 14: LineInfo.h = False
If Index = 6 Then LineInfo.x = 5: LineInfo.y = 12: LineInfo.h = True
If Index = 7 Then LineInfo.x = 1: LineInfo.y = 2: LineInfo.h = False
End Function

Sub DrawSegment(x, y, OFFColor, ONColor, frm As PictureBox)
For i = 1 To 7
DrawLine x, y, i, False, ONColor, OFFColor, frm
Next
frm.Refresh
End Sub

Sub DrawLine(x, y, Number, Lit As Boolean, ONColor, OFFColor, frm As PictureBox)
If Lit = True Then
    If LineInfo(Number).h = True Then
    HLine ONColor, x + LineInfo(Number).x, y + LineInfo(Number).y, frm
    Else
    VLine ONColor, x + LineInfo(Number).x, y + LineInfo(Number).y, frm
    End If
Else
    If LineInfo(Number).h = True Then
    HLine OFFColor, x + LineInfo(Number).x, y + LineInfo(Number).y, frm
    Else
    VLine OFFColor, x + LineInfo(Number).x, y + LineInfo(Number).y, frm
    End If
End If
frm.Refresh
End Sub

Sub HLine(Color, x, y, frm As PictureBox)
For i = 1 To 9
SetPixel frm.hdc, x + i, y + 0, Color
SetPixel frm.hdc, x + i, y + 2, Color
Next
For i = 0 To 10
SetPixel frm.hdc, x + i, y + 1, Color
Next
frm.Refresh
End Sub

Sub VLine(Color, x, y, frm As PictureBox)
SetPixel frm.hdc, x + 1, y + 0, Color
SetPixel frm.hdc, x + 1, y + 10, Color
For i = 1 To 9
SetPixel frm.hdc, x + 0, y + i, Color
SetPixel frm.hdc, x + 1, y + i, Color
SetPixel frm.hdc, x + 2, y + i, Color
Next
frm.Refresh
End Sub

Sub Number2Segment(x, y, Number, pic As PictureBox)
Select Case Number
Case 0
DrawLine x, y, 1, True, ONColor, OFFColor, pic
DrawLine x, y, 2, True, ONColor, OFFColor, pic
DrawLine x, y, 3, True, ONColor, OFFColor, pic
DrawLine x, y, 4, True, ONColor, OFFColor, pic
DrawLine x, y, 5, True, ONColor, OFFColor, pic
DrawLine x, y, 6, False, ONColor, OFFColor, pic
DrawLine x, y, 7, True, ONColor, OFFColor, pic
Case 1
DrawLine x, y, 1, False, ONColor, OFFColor, pic
DrawLine x, y, 2, True, ONColor, OFFColor, pic
DrawLine x, y, 3, True, ONColor, OFFColor, pic
DrawLine x, y, 4, False, ONColor, OFFColor, pic
DrawLine x, y, 5, False, ONColor, OFFColor, pic
DrawLine x, y, 6, False, ONColor, OFFColor, pic
DrawLine x, y, 7, False, ONColor, OFFColor, pic
Case 2
DrawLine x, y, 1, True, ONColor, OFFColor, pic
DrawLine x, y, 2, True, ONColor, OFFColor, pic
DrawLine x, y, 3, False, ONColor, OFFColor, pic
DrawLine x, y, 4, True, ONColor, OFFColor, pic
DrawLine x, y, 5, True, ONColor, OFFColor, pic
DrawLine x, y, 6, True, ONColor, OFFColor, pic
DrawLine x, y, 7, False, ONColor, OFFColor, pic
Case 3
DrawLine x, y, 1, True, ONColor, OFFColor, pic
DrawLine x, y, 2, True, ONColor, OFFColor, pic
DrawLine x, y, 3, True, ONColor, OFFColor, pic
DrawLine x, y, 4, True, ONColor, OFFColor, pic
DrawLine x, y, 5, False, ONColor, OFFColor, pic
DrawLine x, y, 6, True, ONColor, OFFColor, pic
DrawLine x, y, 7, False, ONColor, OFFColor, pic
Case 4
DrawLine x, y, 1, False, ONColor, OFFColor, pic
DrawLine x, y, 2, True, ONColor, OFFColor, pic
DrawLine x, y, 3, True, ONColor, OFFColor, pic
DrawLine x, y, 4, False, ONColor, OFFColor, pic
DrawLine x, y, 5, False, ONColor, OFFColor, pic
DrawLine x, y, 6, True, ONColor, OFFColor, pic
DrawLine x, y, 7, True, ONColor, OFFColor, pic
Case 5
DrawLine x, y, 1, True, ONColor, OFFColor, pic
DrawLine x, y, 2, False, ONColor, OFFColor, pic
DrawLine x, y, 3, True, ONColor, OFFColor, pic
DrawLine x, y, 4, True, ONColor, OFFColor, pic
DrawLine x, y, 5, False, ONColor, OFFColor, pic
DrawLine x, y, 6, True, ONColor, OFFColor, pic
DrawLine x, y, 7, True, ONColor, OFFColor, pic
Case 6
DrawLine x, y, 1, False, ONColor, OFFColor, pic
DrawLine x, y, 2, False, ONColor, OFFColor, pic
DrawLine x, y, 3, True, ONColor, OFFColor, pic
DrawLine x, y, 4, True, ONColor, OFFColor, pic
DrawLine x, y, 5, True, ONColor, OFFColor, pic
DrawLine x, y, 6, True, ONColor, OFFColor, pic
DrawLine x, y, 7, True, ONColor, OFFColor, pic
Case 7
DrawLine x, y, 1, True, ONColor, OFFColor, pic
DrawLine x, y, 2, True, ONColor, OFFColor, pic
DrawLine x, y, 3, True, ONColor, OFFColor, pic
DrawLine x, y, 4, False, ONColor, OFFColor, pic
DrawLine x, y, 5, False, ONColor, OFFColor, pic
DrawLine x, y, 6, False, ONColor, OFFColor, pic
DrawLine x, y, 7, False, ONColor, OFFColor, pic
Case 8
DrawLine x, y, 1, True, ONColor, OFFColor, pic
DrawLine x, y, 2, True, ONColor, OFFColor, pic
DrawLine x, y, 3, True, ONColor, OFFColor, pic
DrawLine x, y, 4, True, ONColor, OFFColor, pic
DrawLine x, y, 5, True, ONColor, OFFColor, pic
DrawLine x, y, 6, True, ONColor, OFFColor, pic
DrawLine x, y, 7, True, ONColor, OFFColor, pic
Case 9
DrawLine x, y, 1, True, ONColor, OFFColor, pic
DrawLine x, y, 2, True, ONColor, OFFColor, pic
DrawLine x, y, 3, True, ONColor, OFFColor, pic
DrawLine x, y, 4, False, ONColor, OFFColor, pic
DrawLine x, y, 5, False, ONColor, OFFColor, pic
DrawLine x, y, 6, True, ONColor, OFFColor, pic
DrawLine x, y, 7, True, ONColor, OFFColor, pic
End Select
End Sub
