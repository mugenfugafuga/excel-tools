VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "EditGraph"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private xs_ As Scripting.Dictionary

Public Function Add(xx As Long, yy As Long, drct As EditDirection)
    GetX_(xx).Add yy, drct
End Function

Public Function GetDirection(xx As Long, yy As Long)
    GetDirection = GetX_(xx).Item(yy)
End Function

Private Function GetX_(xx As Long) As Scripting.Dictionary
    If xs_.Exists(xx) Then
        Set GetX_ = xs_(xx)
    Else
        Dim ret As New Scripting.Dictionary
        xs_.Add xx, ret
        Set GetX_ = xs_(xx)
    End If
End Function

Private Sub Class_Initialize()
    Set xs_ = New Scripting.Dictionary
End Sub

Private Sub Class_Terminate()
    Set xs_ = Nothing
End Sub
