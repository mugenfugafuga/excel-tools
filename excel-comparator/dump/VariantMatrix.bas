Attribute VB_Name = "VariantMatrix"
Option Explicit

Public Type TwoMatrices
    Left As Variant
    Right As Variant
End Type

Public Function PrintVariantOnSheet(vals As Variant, sht As Worksheet) As Range
    Dim lrow As Long, urow As Long
    Dim lcol As Long, ucol As Long
    
    With sht
        .Cells.Clear
        
        If IsEmpty(vals) Then
            Exit Function
        End If
        
        lrow = LBound(vals, 1): urow = UBound(vals, 1)
        lcol = LBound(vals, 2): ucol = UBound(vals, 2)
    
        .Range(.Cells(1, 1), .Cells(urow - lrow + 1, ucol - lcol + 1)) = vals
        
        Set PrintVariantOnSheet = .Range(.Cells(1, 1), .Cells(urow - lrow + 1, ucol - lcol + 1))
    End With
End Function

Public Function PrintVariantOnRange(vals As Variant, rng As Range) As Range
    Dim lrow As Long, urow As Long
    Dim lcol As Long, ucol As Long
    
    rng.Cells.Clear
    
    If IsEmpty(vals) Then
        Exit Function
    End If
    
    With rng.Cells.Item(1)
        lrow = LBound(vals, 1): urow = UBound(vals, 1)
        lcol = LBound(vals, 2): ucol = UBound(vals, 2)
        
        Range(.Offset(0, 0), .Offset(urow - lrow, ucol - lcol)) = vals
        
        Set PrintVariantOnRange = Range(.Offset(0, 0), .Offset(urow - lrow, ucol - lcol))
    End With
End Function

Public Function EditPointsTo2Matrices(ByRef editPnts() As EditPoint) As TwoMatrices
    If UBound(editPnts) = 0 Then
        Dim ret As Variant
        ReDim ret(0, 0)
        
        With EditPointsTo2Matrices
            .Left = ret
            .Right = ret
        End With
        
        Exit Function
    End If

    Dim alen As Long, blen As Long
    alen = GetARowLen_(editPnts)
    blen = GetBRowLen_(editPnts)
    
    Dim num As Long, colnum As Long
    num = UBound(editPnts)
    If alen > blen Then
        colnum = alen
    Else
        colnum = blen
    End If


    Dim lvals As Variant, rvals As Variant
    ReDim lvals(0 To num, 0 To colnum)
    ReDim rvals(0 To num, 0 To colnum)
    
    Dim r As Long, c As Long
    
    lvals(0, 0) = "address"
    For c = 1 To colnum
        lvals(0, c) = "column" & c
    Next c
    
    For r = 1 To num
        With editPnts(r)
            If Not .BBRow Is Nothing Then
                lvals(r, 0) = GetSheetAddress_(.BBRow)
                For c = 1 To blen
                    lvals(r, c) = .BBRow.Cells.Item(c).Value2
                Next c
            End If
        End With
    Next r
        
    rvals(0, 0) = "address"
    For c = 1 To colnum
        rvals(0, c) = "column" & c
    Next c
    
    For r = 1 To num
        With editPnts(r)
            If Not .AARow Is Nothing Then
                rvals(r, 0) = GetSheetAddress_(.AARow)
                For c = 1 To alen
                    rvals(r, c) = .AARow.Cells.Item(c).Value2
                Next c
            End If
        End With
    Next r
    
    With EditPointsTo2Matrices
        .Left = lvals
        .Right = rvals
    End With
End Function

Private Function GetARowLen_(ByRef editPnts() As EditPoint) As Long
    Dim i As Long
    
    For i = 1 To UBound(editPnts)
        With editPnts(i)
            If Not .AARow Is Nothing Then
                GetARowLen_ = .AARow.Cells.Count
                Exit Function
            End If
        End With
    Next i
    
End Function

Private Function GetBRowLen_(ByRef editPnts() As EditPoint) As Long
    Dim i As Long
    
    For i = 1 To UBound(editPnts)
        With editPnts(i)
            If Not .BBRow Is Nothing Then
                GetBRowLen_ = .BBRow.Cells.Count
                Exit Function
            End If
        End With
    Next i
    
End Function
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
    
    ReDim ret(1 To rwCnt, 1 To clCnt + 1)
    
    Dim r As Long, c As Long
    Dim rng As Range
    
    For r = 1 To rwCnt
        Set rng = rngs(r)
        
        ret(r, 1) = GetSheetAddress_(rng)
        
        With rng.Cells
            If rng.Cells.Count < clCnt Then
                For c = 1 To rng.Cells.Count
                    ret(r, c + 1) = .Item(c)
                Next c
            Else
                For c = 1 To clCnt
                    ret(r, c + 1) = .Item(c)
                Next c
            End If
        End With
        
        Set rng = Nothing
    Next r
    
    RangesToMatrix = ret
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
