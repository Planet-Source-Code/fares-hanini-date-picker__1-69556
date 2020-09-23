Attribute VB_Name = "Global"
Option Explicit

'*************** API
Private Declare Function GetCursorPos Lib "user32" (lppoint As POINTAPI) As Long
Private Type POINTAPI
    X As Long
    Y As Long
End Type


Public strDate As String 'variable globale

Public Sub GetMousePos(X As Long, Y As Long)
    Dim lppoint As POINTAPI
    GetCursorPos lppoint
    X = lppoint.X
    Y = lppoint.Y
End Sub

