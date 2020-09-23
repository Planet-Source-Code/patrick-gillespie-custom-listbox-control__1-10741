Attribute VB_Name = "Module1"
Option Explicit

' position/size functions
Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Public Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long

' drawing functions
Public Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
Public Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Public Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Public Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Public Declare Function Rectangle Lib "gdi32" (ByVal hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long

Public Const PS_SOLID = 0

Public Type POINTAPI
        x As Long
        y As Long
End Type

Public Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type

Public Sub DrawRectangle(TheDC As Long, X1 As Long, Y1 As Long, X2 As Long, Y2 As Long, TheColor As Long)
    ' This sub draws the select box rectangle
    ' (the box that shows an item is highlighted)
    Dim NewPen As Long, OldPen As Long
    Dim NewBrush As Long, OldBrush As Long
    
    NewPen = CreatePen(PS_SOLID, 1, TheColor)
    NewBrush = CreateSolidBrush(TheColor)
    OldPen = SelectObject(TheDC, NewPen)
    OldBrush = SelectObject(TheDC, NewBrush)
    
    Call Rectangle(TheDC, X1, Y1, X2, Y2)
    
    Call SelectObject(TheDC, OldPen)
    Call SelectObject(TheDC, OldBrush)
    Call DeleteObject(NewPen)
    Call DeleteObject(NewBrush)
End Sub

Public Sub Timeout(Dur As Double)
    Dim TimeNow As Double
    TimeNow = Timer
    Do While Timer < (TimeNow + Dur)
        DoEvents
    Loop
End Sub
