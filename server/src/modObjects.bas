Attribute VB_Name = "modObjects"
Option Explicit

Sub CheckObject(ByVal index As Long, ByVal x As Long, ByVal y As Long)
Dim MapNum As Long
    
    MapNum = GetPlayerMap(index)
    
    Select Case MapNum
        Case 1
            If x = 10 And y = 10 Then
                Call PlayerWarp(index, 2, 10, 10)
            End If
    End Select
End Sub
