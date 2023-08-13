Attribute VB_Name = "Module1"
Option Explicit

Private Const compSheet_ As String = "ÉfÅ[É^î‰är"

Private Const file1Range_ As String = "FILE1"
Private Const file2Range_ As String = "FILE2"

Private Const withColumnNames_ As String = "w/ column name"
Private Const allRows_ As String = "all row"

Private Type FileItem_
    Path As String
    Charset As String
End Type

Private Type ComResult_
    Comparison As Worksheet
    Rest As Worksheet
End Type

Private Enum CompType_
    Both
    No1WithNo2
    No2WithNo1
End Enum

Sub File1_UTF8_W_Columns_Click()
    Dim fileName As String
    fileName = SelectFile("á@UTF8ÇÃÉtÉ@ÉCÉãÇëIëÇµÇƒÇ≠ÇæÇ≥Ç¢ÅB")
    
    If fileName = "" Then
        Exit Sub
    End If
    
    Dim rw As Long
    
    With Range(file1Range_)
        rw = .CurrentRegion.Rows.Count
        
        .Offset(rw, -1) = withColumnNames_
        .Offset(rw, 0) = fileName
        .Offset(rw, 1) = UTF8Charset
    End With
End Sub

Sub File1_SJIS_W_Columns_Click()
    Dim fileName As String
    fileName = SelectFile("á@SJISÇÃÉtÉ@ÉCÉãÇëIëÇµÇƒÇ≠ÇæÇ≥Ç¢ÅB")
    
    If fileName = "" Then
        Exit Sub
    End If
    
    Dim rw As Long
    
    With Range(file1Range_)
        rw = .CurrentRegion.Rows.Count
        
        .Offset(rw, -1) = withColumnNames_
        .Offset(rw, 0) = fileName
        .Offset(rw, 1) = SJISCharset
    End With
End Sub

Sub File1_UTF8_All_Rows_Click()
    Dim fileName As String
    fileName = SelectFile("á@UTF8ÇÃÉtÉ@ÉCÉãÇëIëÇµÇƒÇ≠ÇæÇ≥Ç¢ÅB")
    
    If fileName = "" Then
        Exit Sub
    End If
    
    Dim rw As Long
    
    With Range(file1Range_)
        rw = .CurrentRegion.Rows.Count
        
        .Offset(rw, -1) = allRows_
        .Offset(rw, 0) = fileName
        .Offset(rw, 1) = UTF8Charset
    End With
End Sub

Sub File1_SJIS_All_Rows_Click()
    Dim fileName As String
    fileName = SelectFile("á@SJISÇÃÉtÉ@ÉCÉãÇëIëÇµÇƒÇ≠ÇæÇ≥Ç¢ÅB")
    
    If fileName = "" Then
        Exit Sub
    End If
    
    Dim rw As Long
    
    With Range(file1Range_)
        rw = .CurrentRegion.Rows.Count
        
        .Offset(rw, -1) = allRows_
        .Offset(rw, 0) = fileName
        .Offset(rw, 1) = SJISCharset
    End With
End Sub

Sub File2_UTF8_Click()
    Dim fileName As String
    fileName = SelectFile("áAUTF8ÇÃÉtÉ@ÉCÉãÇëIëÇµÇƒÇ≠ÇæÇ≥Ç¢ÅB")
    
    If fileName = "" Then
        Exit Sub
    End If
    
    Dim rw As Long
    
    With Range(file2Range_)
        rw = .CurrentRegion.Rows.Count
        
        .Offset(rw, 0) = fileName
        .Offset(rw, 1) = UTF8Charset
    End With
End Sub

Sub File2_SJIS_Click()
    Dim fileName As String
    fileName = SelectFile("áASJISÇÃÉtÉ@ÉCÉãÇëIëÇµÇƒÇ≠ÇæÇ≥Ç¢ÅB")
    
    If fileName = "" Then
        Exit Sub
    End If
    
    Dim rw As Long
    
    With Range(file2Range_)
        rw = .CurrentRegion.Rows.Count
        
        .Offset(rw, 0) = fileName
        .Offset(rw, 1) = UTF8Charset
    End With
End Sub

Sub Do_Compare_Both_Click()
    DoCompareFiles_ Both
End Sub

Sub Do_Compare_No1_With_No2_Click()
    DoCompareFiles_ No1WithNo2
End Sub

Sub Do_Compare_No2_With_No1_Click()
    DoCompareFiles_ No2WithNo1
End Sub

Sub Clear_Comparison_Sheet_Click()
    With ThisWorkbook.Sheets(compSheet_)
        .Rows("7:" & .UsedRange.Rows.Count + 7).ClearContents
    End With
End Sub

Private Function DoCompareFiles_(ctype As CompType_)
    Dim no1cnt As Long, no2cnt As Long
    Dim cnt As Long
    
    Dim i As Long
    Dim empties As Boolean
    
    Dim f1 As FileItem_, f2 As FileItem_
    Dim fileType As String
    
    With ThisWorkbook.Sheets(compSheet_)
        no1cnt = .Range(file1Range_).CurrentRegion.Rows.Count
        no2cnt = .Range(file2Range_).CurrentRegion.Rows.Count
        
        If no1cnt < no2cnt Then
            cnt = no1cnt
        Else
            cnt = no2cnt
        End If
        
        For i = 2 To cnt
            With .Range(file1Range_).Cells(i, 1)
                empties = IsEmpty(.Offset(0, 0).Value) And IsEmpty(.Offset(0, 1).Value)
            End With
            With .Range(file2Range_).Cells(i, 1)
                empties = empties And IsEmpty(.Offset(0, 0).Value) And IsEmpty(.Offset(0, 1).Value)
            End With
            
            If Not empties Then
                With .Range(file1Range_).Cells(i, 1)
                    If IsEmpty(.Offset(0, -1)) Then
                        fileType = allRows_
                    Else
                        fileType = .Offset(0, -1).Value2
                    End If
                
                    f1.Path = CStr(.Offset(0, 0).Value2)
                    f1.Charset = CStr(.Offset(0, 1).Value2)
                End With
                
                With .Range(file2Range_).Cells(i, 1)
                    f2.Path = CStr(.Offset(0, 0).Value2)
                    f2.Charset = CStr(.Offset(0, 1).Value2)
                End With
                
                DoCompare_ fileType, f1, f2, ctype
            End If
        Next i
    End With
    
    TryDeleteTempFiles
End Function

Private Function DoCompare_(fileType As String, file1 As FileItem_, file2 As FileItem_, ctype As CompType_)
    Dim wb1 As Workbook, wb2 As Workbook
    Dim resultbook As Workbook
    
    Dim no1 As Worksheet, no2 As Worksheet
    Dim run As Worksheet
    
    Dim no1rng As Range, no2rng As Range
    
    Set resultbook = NewWorkBook(1)
    
    Dim com1With2Result As ComResult_
    Dim com2With1Result As ComResult_
    
    Set wb1 = OpenBook_(file1)
    With wb1
        Set no1 = CopySheetAfterTarget(.Sheets(1), resultbook.Sheets(1))
        no1.Name = "á@"
        no1.Cells.EntireColumn.AutoFit
        Set no1rng = no1.UsedRange
        
        .Close
        Set wb1 = Nothing
    End With
    
    Set wb2 = OpenBook_(file2)
    With wb2
        Set no2 = CopySheetAfterTarget(.Sheets(1), no1)
        no2.Name = "áA"
        no2.Cells.EntireColumn.AutoFit
        Set no2rng = no2.UsedRange
    
        .Close
        Set wb2 = Nothing
    End With
        
    Set run = no2
    
    If ctype <> No2WithNo1 Then
        com1With2Result = AddComparisonResultAfterTarget_( _
            fileType, _
            run, no1rng, no2rng, _
            "á@ÇáAÇ…ìÀçáÇπÇΩåãâ ", "á@Ç…Ç†Ç¡ÇƒáAÇ…Ç»Ç¢", xlThemeColorAccent6)
        Set run = com1With2Result.Rest
    End If
    
    If ctype <> No1WithNo2 Then
        com2With1Result = AddComparisonResultAfterTarget_( _
            fileType, _
            run, no2rng, no1rng, _
            "áAÇá@Ç…ìÀçáÇπÇΩåãâ ", "áAÇ…Ç†Ç¡Çƒá@Ç…Ç»Ç¢", xlThemeColorAccent5)
    End If
    
    With resultbook
        With .Sheets(1)
            .Activate
            .Name = "åãâ Ç‹Ç∆Çﬂ"
            With .Tab
                .ThemeColor = xlThemeColorAccent4
                .TintAndShade = 0.399975585192419
            End With
        End With
        
        ResultSummary_ fileType, .Sheets(1), file1, no1, file2, no2, com1With2Result, com2With1Result
    End With
    
End Function

Private Function ResultSummary_( _
    fileType As String, _
    sht As Worksheet, _
    file1 As FileItem_, sheet1 As Worksheet, _
    file2 As FileItem_, sheet2 As Worksheet, _
    com1w2result As ComResult_, com2w1result As ComResult_)
    
    With sht.Cells(2, 2)
        With .Offset(0, 0)
            .Offset(0, 0) = "ÉeÅ[ÉuÉã": .Offset(0, 1) = "ÉåÉRÅ[Éhêî": .Offset(0, 2) = "ÉtÉ@ÉCÉã"
        
            If fileType = withColumnNames_ Then
                .Offset(1, 0) = "á@": .Offset(1, 1) = sheet1.UsedRange.Rows.Count - 1: .Offset(1, 2) = file1.Path
                .Offset(2, 0) = "áA": .Offset(2, 1) = sheet2.UsedRange.Rows.Count - 1: .Offset(2, 2) = file2.Path
            Else 'if fileType = allRows_ Then
                .Offset(1, 0) = "á@": .Offset(1, 1) = sheet1.UsedRange.Rows.Count: .Offset(1, 2) = file1.Path
                .Offset(2, 0) = "áA": .Offset(2, 1) = sheet2.UsedRange.Rows.Count: .Offset(2, 2) = file2.Path
            End If
        End With
        
        With .Offset(4, 0)
            .Offset(0, 0) = "á@ÇáAÇ…ìÀçáÇπÇΩåãâ "
            
            .Offset(1, 0) = "á@Ç…Ç†Ç¡ÇƒáAÇ…Ç»Ç¢"
            .Offset(2, 0) = "1 ëŒ 1"
            .Offset(3, 0) = "1 ëŒ ëΩ"
            
            If Not com1w2result.Rest Is Nothing Then
                .Offset(1, 1).Formula = "=COUNTA('" & com1w2result.Rest.Name & "'!A:A)"
            End If
            
            If Not com1w2result.Comparison Is Nothing Then
                .Offset(2, 1).Formula = "=COUNTIF('" & com1w2result.Comparison.Name & "'!B:B,1)"
                .Offset(3, 1).Formula = "=COUNTA('" & com1w2result.Comparison.Name & "'!B:B)-" & .Offset(2, 1).Address & "-1"
            End If
        End With
        
        With .Offset(9, 0)
            .Offset(0, 0) = "áAÇá@Ç…ìÀçáÇπÇΩåãâ "
            
            .Offset(1, 0) = "áAÇ…Ç†Ç¡Çƒá@Ç…Ç»Ç¢"
            .Offset(2, 0) = "1 ëŒ 1"
            .Offset(3, 0) = "1 ëŒ ëΩ"
            
            If Not com2w1result.Rest Is Nothing Then
                .Offset(1, 1).Formula = "=COUNTA('" & com2w1result.Rest.Name & "'!A:A)"
            End If
            
            If Not com2w1result.Comparison Is Nothing Then
                .Offset(2, 1).Formula = "=COUNTIF('" & com2w1result.Comparison.Name & "'!B:B,1)"
                .Offset(3, 1).Formula = "=COUNTA('" & com2w1result.Comparison.Name & "'!B:B)-" & .Offset(2, 1).Address & "-1"
            End If
        End With
    End With
    
    With sht
        .Columns("C:C").NumberFormatLocal = "#,##0_ "
        .Cells.EntireColumn.AutoFit
    End With
    
End Function

Private Function AddComparisonResultAfterTarget_( _
    fileType As String, _
    target As Worksheet, _
    rng1 As Range, _
    rng2 As Range, _
    comparisonSheetName As String, _
    restSheetName As String, _
    tabColor As XlThemeColor _
    ) As ComResult_
    
    With AddComparisonResultAfterTarget_
        Dim rngComResult As RangeComparisonResult
        
        If fileType = withColumnNames_ Then
            rngComResult = CompareTableWithTableDataHashSet(rng1, CreateTableDataHashSet(rng2))
        Else 'if fileType = allRows_ Then
            rngComResult = CompareTableWithRangeHasSet(rng1, CreateRowHashSet(rng2))
        End If
        
        Set .Comparison = AddSheetAfter(target)
        PrintVariantOnSheet MatchResultsToMatrix(rngComResult.Matchs), .Comparison
        
        With .Comparison
            .Cells.EntireColumn.AutoFit
            
            .Name = comparisonSheetName
            With .Tab
                .ThemeColor = tabColor
                .TintAndShade = 0.399975585192419
            End With
        End With
        
        
        Set .Rest = AddSheetAfter(.Comparison)
        PrintVariantOnSheet RangesToMatrix(rngComResult.Rest), .Rest
        
        With .Rest
            .Cells.EntireColumn.AutoFit
        
            .Name = restSheetName
            With .Tab
                .ThemeColor = tabColor
                .TintAndShade = 0.399975585192419
            End With
        End With
        
    End With
End Function

Private Function OpenBook_(file As FileItem_) As Workbook
    With file
        If .Charset = SJISCharset Then
            Set OpenBook_ = Workbooks.Open(.Path)
            Exit Function
        End If
        
        Dim tempFile As String
        tempFile = GetTempFileName(GetExtention(.Path))
        
        ConvertWithNewCharset .Path, .Charset, tempFile, SJISCharset
        Set OpenBook_ = Workbooks.Open(tempFile)
    End With
End Function
