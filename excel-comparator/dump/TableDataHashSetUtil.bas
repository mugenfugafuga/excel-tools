Attribute VB_Name = "TableDataHashSetUtil"
Option Explicit

Public Function CreateTableDataHashSet(tbl As Range) As TableDataHashSet
    Dim ret As New TableDataHashSet
    
    Dim record As Range
    
    Dim firstRecord As Boolean
    firstRecord = True
    
    With ret
        For Each record In tbl.Rows
            If firstRecord Then
                .Initialize record
                firstRecord = False
            Else
                .Add record
            End If
        Next record
    End With
    
    Set CreateTableDataHashSet = ret
    Set ret = Nothing
End Function

Public Function CreateColumnIndexMap(rng As Range) As Scripting.Dictionary
    Dim ret As New Scripting.Dictionary
    
    Dim i As Long
    
    With rng.Cells
        For i = 1 To .Count
            ret.Add CStr(.Item(i).Value2), i
        Next i
    End With
    
    Set CreateColumnIndexMap = ret
End Function
