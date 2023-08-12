Attribute VB_Name = "Utilities"
Option Explicit

Function Assert(condition As Boolean)
    If Not condition Then
        Stop
    End If
End Function
