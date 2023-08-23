Attribute VB_Name = "ExcelDifference"
Option Explicit

Public Enum ARowOrBRowType
    Both
    AARow
    BBRow
End Enum

Public Type EditPoint
    ABType As ARowOrBRowType
    AARow As Range
    BBRow As Range
End Type

Public Enum EditDirection
    None_
    Same
    Delete ' down  x: +1
    Insert ' right y: +1
End Enum

Public Function DiffExcel( _
    table0 As Range, _
    table1 As Range, _
    Optional ignoreNum As Boolean = False, _
    Optional tolernce As Long = 0 _
    ) As EditPoint()
    
    Dim a As Range, m As Long
    Dim b As Range, n As Long
    
    If table0.Rows.Count > table1.Rows.Count Then
        Set b = table0: n = b.Rows.Count
        Set a = table1: m = a.Rows.Count
    Else
        Set b = table1: n = b.Rows.Count
        Set a = table0: m = a.Rows.Count
    End If
    
    Dim delta As Long
    Dim fp() As Long
    
    delta = n - m
        
    Dim i As Long
    ReDim fp(-m - 1 To n + 1)
    For i = -m - 1 To n + 1: fp(i) = -1: Next i
    
    Dim editGrph As New EditGraph
    
    Dim p As Long, k As Long, y As Long
    For p = 0 To m
        For k = -p To delta - 1
            y = MaxAndAddDirection_(fp(k - 1) + 1, fp(k + 1), k, editGrph)
            fp(k) = Snake_(a, b, k, y, editGrph, ignoreNum, tolernce)
        Next k
        For k = delta + p To delta + 1 Step -1
            y = MaxAndAddDirection_(fp(k - 1) + 1, fp(k + 1), k, editGrph)
            fp(k) = Snake_(a, b, k, y, editGrph, ignoreNum, tolernce)
        Next k
        y = MaxAndAddDirection_(fp(delta - 1) + 1, fp(delta + 1), k, editGrph)
        fp(delta) = Snake_(a, b, delta, y, editGrph, ignoreNum, tolernce)
        
        If fp(delta) = n Then
            DiffExcel = GetEditPoints_(a, b, editGrph)
            Exit Function
        End If
    Next p

    Set editGrph = Nothing
End Function

Private Function Snake_( _
    a As Range, _
    b As Range, _
    k As Long, _
    y As Long, _
    editGrph As EditGraph, _
    ignoreNum As Boolean, _
    tolernce As Long _
    ) As Long
    Dim xx As Long, yy As Long
    yy = y
    xx = y - k
    
    While xx < a.Rows.Count And yy < b.Rows.Count And IsSame_(a.Rows(xx + 1), b.Rows(yy + 1), ignoreNum, tolernce)
        xx = xx + 1
        yy = yy + 1
        editGrph.Add xx, yy, Same
    Wend
    
    Snake_ = yy
End Function

Private Function IsSame_(arow As Range, brow As Range, ignoreNum As Boolean, tolernce As Long) As Boolean
    Dim num As Long
    If arow.Cells.Count < brow.Cells.Count Then
        num = arow.Cells.Count
    Else
        num = brow.Cells.Count
    End If
    
    Dim i As Long, errNum As Long
    
    errNum = 0
    If ignoreNum Then
        For i = 1 To num
            If _
                (IsNumeric(arow.Cells.Item(i).Value2) And IsNumeric(brow.Cells.Item(i).Value2)) = False And _
                arow.Cells.Item(i).Value2 <> brow.Cells.Item(i).Value2 _
            Then
                errNum = errNum + 1
                If errNum > tolernce Then
                    IsSame_ = False
                    Exit Function
                End If
            End If
        Next i
    Else
        For i = 1 To num
            If arow.Cells.Item(i).Value2 <> brow.Cells.Item(i).Value2 Then
                errNum = errNum + 1
                If errNum > tolernce Then
                    IsSame_ = False
                    Exit Function
                End If
            End If
        Next i
    End If
    
    IsSame_ = True
End Function

Private Function MaxAndAddDirection_(v0 As Long, v1 As Long, k As Long, editGrph As EditGraph) As Long
    If v0 > v1 Then
        MaxAndAddDirection_ = v0
        editGrph.Add v0 - k, v0, Insert
    Else
        MaxAndAddDirection_ = v1
        editGrph.Add v1 - k, v1, Delete
    End If
End Function

Private Function GetEditPoints_(a As Range, b As Range, editGrph As EditGraph) As EditPoint()
    Dim xx As Long, yy As Long
    xx = a.Rows.Count
    yy = b.Rows.Count
    
    Dim rvrs() As EditPoint
    ReDim rvrs(1 To xx + yy)
    
    Dim drct As EditDirection
    Dim i As Integer
    i = 0
    
    While xx <> 0 Or yy <> 0
        i = i + 1
        
        With rvrs(i)
            drct = editGrph.GetDirection(xx, yy)
            
            If drct = Delete Then
                .ABType = AARow
                Set .AARow = a.Rows(xx)
                xx = xx - 1
            ElseIf drct = Insert Then
                .ABType = BBRow
                Set .BBRow = b.Rows(yy)
                yy = yy - 1
            Else 'ElseIf drct = Same Then
                .ABType = Both
                Set .AARow = a.Rows(xx)
                Set .BBRow = b.Rows(yy)
                xx = xx - 1
                yy = yy - 1
            End If
        End With
    Wend
    
    Dim num As Long
    num = i
    
    Dim eps() As EditPoint
    ReDim eps(1 To num)
    
    
    For i = 1 To num
        eps(i) = rvrs(num - i + 1)
    Next i
    
    GetEditPoints_ = eps
End Function
