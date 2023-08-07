Attribute VB_Name = "RangeHashSetFactory"
Option Explicit

Public Function CreateRowHashSet(table As Range) As RangeHashSet
    Dim hs As New RangeHashSet
    Dim rw As Range
    
    For Each rw In table.Rows
        hs.Add rw
    Next rw
    
    Set CreateRowHashSet = hs
    Set hs = Nothing
End Function
