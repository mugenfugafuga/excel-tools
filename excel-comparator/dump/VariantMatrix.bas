Attribute VB_Name = "VariantMatrix"
Option Explicit

Public Function RangesToMatrix(ByRef rngs() As Range) As Variant
    Dim ret As Variant
    
    Dim rwCnt As Long
    rwCnt = UBound(rngs)
    
    If rwCnt = 0 Then
        ReDim ret(0, 0)
        RangesToMatrix = ret
        Exit Function
    End If
    
    Dim clCnt As Long
    clCnt = rngs(1).Cells.Count
    
    ReDim ret(1 To rwCnt, 1 To clCnt)
    
    Dim r As Long, c As Long
    Dim rng As Range
    
    For r = 1 To rwCnt
        Set rng = rngs(r)
        
        With rng.Cells
            If rng.Cells.Count < clCnt Then
                For c = 1 To rng.Cells.Count
                    ret(r, c) = .Item(c)
                Next c
            Else
                For c = 1 To clCnt
                    ret(r, c) = .Item(c)
                Next c
            End If
        End With
        
        Set rng = Nothing
    Next r
End Function

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
        
        Dim i As Long
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
    Dim cnt As Long
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
        
    Dim i As Long
    For i = 2 To cnt
        ConcatSheetAddresses_ = ConcatSheetAddresses_ & "," & GetSheetAddress_(vals(i))
    Next i
End Function
