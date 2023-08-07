Attribute VB_Name = "VariantMatrix"
Option Explicit

Public Function MatchResultsToMatrix(ByRef result As MatchResults, Optional headers As Boolean = True) As Variant
    With result
        Dim ret As Variant
        If headers Then
            ReDim ret(0 To .Count, 1 To 3)
            ret(0, 1) = "target"
            ret(0, 2) = "match.number"
            ret(0, 3) = "match.values"
        Else
            ReDim ret(1 To .Count, 1 To 3)
        End If
        
        Dim i As Integer
        For i = 1 To .Count
            With .Results(i)
                
                ret(i, 1) = GetSheetAddress_(.Value)
                ret(i, 2) = UBound(.Matchs)
                ret(i, 3) = ConcatSheetAddresses_(.Matchs)
            End With
        Next i
        
        MatchResultsToMatrix = ret
    End With
End Function

Private Function GetSheetAddress_(val As Range) As String
    GetSheetAddress_ = "'" & val.Worksheet.Name & "'!" & val.Address
End Function

Private Function ConcatSheetAddresses_(vals() As Range) As String
    Dim cnt As Integer
    cnt = UBound(vals)
    
    If cnt = 0 Then
        ConcatSheetAddresses_ = ""
        Exit Function
    End If
    
    If cnt = 1 Then
        ConcatSheetAddresses_ = GetSheetAddress_(vals(1))
        Exit Function
    End If
    
    ConcatSheetAddresses_ = GetSheetAddress_(vals(1))
        
    Dim i As Integer
    For i = 2 To cnt
        ConcatSheetAddresses_ = ConcatSheetAddresses_ & "," & GetSheetAddress_(vals(i))
    Next i
End Function
