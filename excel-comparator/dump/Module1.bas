Attribute VB_Name = "Module1"
Option Explicit

Private Const compSheet_ As String = "データ比較"

Private Const file1Range_ As String = "FILE1"
Private Const file2Range_ As String = "FILE2"

Private Const withColumnNames_ As String = "w/ column name"
Private Const allRows_ As String = "all row"
Private Const diff_ As String = "diff"

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
    fileName = SelectFile("�@UTF8のファイルを選択してください。")
    
    If fileName = "" Then
        Exit Sub
    End If
    
    Dim rw As Long
    
    With Range(file1Range_)
        rw = .CurrentRegion.Rows.Count
        
        .Offset(rw, -3) = withColumnNames_
        .Offset(rw, -2) = "-"
        .Offset(rw, -1) = "完全一致"
        .Offset(rw, 0) = fileName
        .Offset(rw, 1) = UTF8Charset
    End With
End Sub

Sub File1_SJIS_W_Columns_Click()
    Dim fileName As String
    fileName = SelectFile("�@SJISのファイルを選択してください。")
    
    If fileName = "" Then
        Exit Sub
    End If
    
    Dim rw As Long
    
    With Range(file1Range_)
        rw = .CurrentRegion.Rows.Count
        
        .Offset(rw, -3) = withColumnNames_
        .Offset(rw, -2) = "-"
        .Offset(rw, -1) = "完全一致"
        .Offset(rw, 0) = fileName
        .Offset(rw, 1) = SJISCharset
    End With
End Sub

Sub File1_UTF8_All_Rows_Click()
    Dim fileName As String
    fileName = SelectFile("�@UTF8のファイルを選択してください。")
    
    If fileName = "" Then
        Exit Sub
    End If
    
    Dim rw As Long
    
    With Range(file1Range_)
        rw = .CurrentRegion.Rows.Count
        
        .Offset(rw, -3) = allRows_
        .Offset(rw, -2) = "-"
        .Offset(rw, -1) = "完全一致"
        .Offset(rw, 0) = fileName
        .Offset(rw, 1) = UTF8Charset
    End With
End Sub

Sub File1_SJIS_All_Rows_Click()
    Dim fileName As String
    fileName = SelectFile("�@SJISのファイルを選択してください。")
    
    If fileName = "" Then
        Exit Sub
    End If
    
    Dim rw As Long
    
    With Range(file1Range_)
        rw = .CurrentRegion.Rows.Count
        
        .Offset(rw, -3) = allRows_
        .Offset(rw, -2) = "-"
        .Offset(rw, -1) = "完全一致"
        .Offset(rw, 0) = fileName
        .Offset(rw, 1) = SJISCharset
    End With
End Sub

Sub File1_UTF8_Diff_Click()
    Dim fileName As String
    fileName = SelectFile("�@UTF8のファイルを選択してください。")
    
    If fileName = "" Then
        Exit Sub
    End If
    
    Dim rw As Long
    
    With Range(file1Range_)
        rw = .CurrentRegion.Rows.Count
        
        .Offset(rw, -3) = diff_
        .Offset(rw, -2) = True
        .Offset(rw, -1) = 0
        .Offset(rw, 0) = fileName
        .Offset(rw, 1) = UTF8Charset
    End With
End Sub

Sub File1_SJIS_Diff_Click()
    Dim fileName As String
    fileName = SelectFile("�@SJISのファイルを選択してください。")
    
    If fileName = "" Then
        Exit Sub
    End If
    
    Dim rw As Long
    
    With Range(file1Range_)
        rw = .CurrentRegion.Rows.Count
        
        .Offset(rw, -3) = diff_
        .Offset(rw, -2) = True
        .Offset(rw, -1) = 0
        .Offset(rw, 0) = fileName
        .Offset(rw, 1) = SJISCharset
    End With
End Sub

Sub File2_UTF8_Click()
    Dim fileName As String
    fileName = SelectFile("�AUTF8のファイルを選択してください。")
    
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
    fileName = SelectFile("�ASJISのファイルを選択してください。")
    
    If fileName = "" Then
        Exit Sub
    End If
    
    Dim rw As Long
    
    With Range(file2Range_)
        rw = .CurrentRegion.Rows.Count
        
        .Offset(rw, 0) = fileName
        .Offset(rw, 1) = SJISCharset
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
    Dim cnt As Long, ignoreNum As Boolean, tolernce As Long
    
    Dim i As Long
    Dim empties As Boolean
    
    Dim f1 As FileItem_, f2 As FileItem_
    Dim compType As String
    
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
                With .Range(file2Range_).Cells(i, 1)
                    f2.Path = CStr(.Offset(0, 0).Value2)
                    f2.Charset = CStr(.Offset(0, 1).Value2)
                End With
                
                With .Range(file1Range_).Cells(i, 1)
                    f1.Path = CStr(.Offset(0, 0).Value2)
                    f1.Charset = CStr(.Offset(0, 1).Value2)
                    
                    If IsEmpty(.Offset(0, -3)) Then
                        compType = allRows_
                    Else
                        compType = .Offset(0, -3).Value2
                    End If
                
                    If compType = diff_ Then
                        ignoreNum = .Offset(0, -2)
                        tolernce = .Offset(0, -1)
                        
                        DoDiff_ f1, f2, ignoreNum, tolernce
                    Else
                        DoCompare_ compType, f1, f2, ctype
                    End If
                End With
            End If
        Next i
    End With
    
    TryDeleteTempFiles
End Function

Private Function DoDiff_(file1 As FileItem_, file2 As FileItem_, ignoreNum As Boolean, tolernce As Long)
    Dim wb1 As Workbook, wb2 As Workbook
    Dim resultbook As Workbook
    
    Dim no1 As Worksheet, no2 As Worksheet
    
    Dim no1rng As Range, no2rng As Range
    
    Set resultbook = NewWorkBook(1)
    
    Dim com1With2Result As ComResult_
    Dim com2With1Result As ComResult_
    
    Set wb1 = OpenBook_(file1)
    With wb1
        Set no1 = CopySheetAfterTarget(.Sheets(1), resultbook.Sheets(1))
        no1.Name = "�@"
        no1.Cells.EntireColumn.AutoFit
        Set no1rng = no1.UsedRange
        
        .Close
        Set wb1 = Nothing
    End With
    
    Set wb2 = OpenBook_(file2)
    With wb2
        Set no2 = CopySheetAfterTarget(.Sheets(1), no1)
        no2.Name = "�A"
        no2.Cells.EntireColumn.AutoFit
        Set no2rng = no2.UsedRange
    
        .Close
        Set wb2 = Nothing
    End With
    
    Dim twoM As TwoMatrices
    Dim lr As Range, rr As Range
    Dim th1 As Range, th2 As Range
    
    With resultbook.Sheets(1)
        twoM = EditPointsTo2Matrices(DiffExcel(no1.UsedRange, no2.UsedRange, ignoreNum, tolernce))
        
        With .Cells(5, 1)
            .Offset(0, 0) = "閾値1": .Offset(0, 1) = 1000
            .Offset(0, 0).HorizontalAlignment = xlRight
            Set th1 = .Offset(0, 1)
            
            .Offset(0, 2) = "閾値2": .Offset(0, 3) = 0.00001
            .Offset(0, 2).HorizontalAlignment = xlRight
            Set th2 = .Offset(0, 3)
        End With
        
        Set lr = PrintVariantOnRange(twoM.Left, .Cells(7, 1))
        Set rr = PrintVariantOnRange(twoM.Right, lr.Cells(1, lr.Columns.Count).Offset(0, 2))
        
        UpdateFormatConditions_ SkipColumns_(lr), SkipColumns_(rr), th1, th2
        UpdateFormatConditions_ SkipColumns_(rr), SkipColumns_(lr), th1, th2
        
        .Name = "結果"
        .Cells.EntireColumn.AutoFit
        
        With .Cells(2, 1)
            .Offset(0, 0) = "�@": .Offset(0, 0).HorizontalAlignment = xlRight: .Offset(0, 1) = file1.Path
            .Offset(1, 0) = "�A": .Offset(1, 0).HorizontalAlignment = xlRight: .Offset(1, 1) = file2.Path
        End With
        
        .Activate
    End With
End Function

Private Function UpdateFormatConditions_( _
    target As Range, _
    othr As Range, _
    th1 As Range, _
    th2 As Range)
    Dim ttop As String, otop As String
    ttop = target.Cells(1, 1).Address(RowAbsolute:=False, ColumnAbsolute:=False)
    otop = othr.Cells(1, 1).Address(RowAbsolute:=False, ColumnAbsolute:=False)

    With target.FormatConditions
        .Delete
        
        .Add Type:=xlExpression, Formula1:="=ISBLANK(" & ttop & ")"
        .Item(1).Interior.Color = RGB(250, 250, 210)
        
        .Add Type:=xlExpression, Formula1:="=ABS(" & ttop & "-" & otop & ")>" & th1.Cells(1, 1).Address
        .Item(2).Interior.Color = RGB(255, 165, 0)
        
        .Add Type:=xlExpression, Formula1:="=ABS(" & ttop & "-" & otop & ")>" & th2.Cells(1, 1).Address
        .Item(3).Interior.Color = RGB(175, 238, 238)
        
        .Add Type:=xlCellValue, Operator:=xlNotEqual, Formula1:="=" & otop
        .Item(4).Interior.Color = RGB(255, 99, 71)
        
    End With
End Function

Private Function SkipColumns_(rng As Range, Optional cols As Long = 1) As Range
    With rng
        Set SkipColumns_ = Range(.Cells(1, cols + 1), .Cells(.Rows.Count, .Columns.Count))
    End With
End Function

Private Function DoCompare_(compType As String, file1 As FileItem_, file2 As FileItem_, ctype As CompType_)
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
        no1.Name = "�@"
        no1.Cells.EntireColumn.AutoFit
        Set no1rng = no1.UsedRange
        
        .Close
        Set wb1 = Nothing
    End With
    
    Set wb2 = OpenBook_(file2)
    With wb2
        Set no2 = CopySheetAfterTarget(.Sheets(1), no1)
        no2.Name = "�A"
        no2.Cells.EntireColumn.AutoFit
        Set no2rng = no2.UsedRange
    
        .Close
        Set wb2 = Nothing
    End With
        
    Set run = no2
    
    If ctype <> No2WithNo1 Then
        com1With2Result = AddComparisonResultAfterTarget_( _
            compType, _
            run, no1rng, no2rng, _
            "�@を�Aに突合せた結果", "�@にあって�Aにない", xlThemeColorAccent6)
        Set run = com1With2Result.Rest
    End If
    
    If ctype <> No1WithNo2 Then
        com2With1Result = AddComparisonResultAfterTarget_( _
            compType, _
            run, no2rng, no1rng, _
            "�Aを�@に突合せた結果", "�Aにあって�@にない", xlThemeColorAccent5)
    End If
    
    With resultbook
        With .Sheets(1)
            .Activate
            .Name = "結果まとめ"
            With .Tab
                .ThemeColor = xlThemeColorAccent4
                .TintAndShade = 0.399975585192419
            End With
        End With
        
        ResultSummary_ compType, .Sheets(1), file1, no1, file2, no2, com1With2Result, com2With1Result
    End With
    
End Function

Private Function ResultSummary_( _
    compType As String, _
    sht As Worksheet, _
    file1 As FileItem_, sheet1 As Worksheet, _
    file2 As FileItem_, sheet2 As Worksheet, _
    com1w2result As ComResult_, com2w1result As ComResult_)
    
    With sht.Cells(2, 2)
        With .Offset(0, 0)
            .Offset(0, 0) = "テーブル": .Offset(0, 1) = "レコード数": .Offset(0, 2) = "ファイル"
        
            If compType = withColumnNames_ Then
                .Offset(1, 0) = "�@": .Offset(1, 1) = sheet1.UsedRange.Rows.Count - 1: .Offset(1, 2) = file1.Path
                .Offset(2, 0) = "�A": .Offset(2, 1) = sheet2.UsedRange.Rows.Count - 1: .Offset(2, 2) = file2.Path
            Else 'if compType = allRows_ Then
                .Offset(1, 0) = "�@": .Offset(1, 1) = sheet1.UsedRange.Rows.Count: .Offset(1, 2) = file1.Path
                .Offset(2, 0) = "�A": .Offset(2, 1) = sheet2.UsedRange.Rows.Count: .Offset(2, 2) = file2.Path
            End If
        End With
        
        With .Offset(4, 0)
            .Offset(0, 0) = "�@を�Aに突合せた結果"
            
            .Offset(1, 0) = "�@にあって�Aにない"
            .Offset(2, 0) = "1 対 1"
            .Offset(3, 0) = "1 対 多"
            
            If Not com1w2result.Rest Is Nothing Then
                .Offset(1, 1).Formula = "=COUNTA('" & com1w2result.Rest.Name & "'!A:A)"
            End If
            
            If Not com1w2result.Comparison Is Nothing Then
                .Offset(2, 1).Formula = "=COUNTIF('" & com1w2result.Comparison.Name & "'!B:B,1)"
                .Offset(3, 1).Formula = "=COUNTA('" & com1w2result.Comparison.Name & "'!B:B)-" & .Offset(2, 1).Address & "-1"
            End If
        End With
        
        With .Offset(9, 0)
            .Offset(0, 0) = "�Aを�@に突合せた結果"
            
            .Offset(1, 0) = "�Aにあって�@にない"
            .Offset(2, 0) = "1 対 1"
            .Offset(3, 0) = "1 対 多"
            
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
    compType As String, _
    target As Worksheet, _
    rng1 As Range, _
    rng2 As Range, _
    comparisonSheetName As String, _
    restSheetName As String, _
    tabColor As XlThemeColor _
    ) As ComResult_
    
    With AddComparisonResultAfterTarget_
        Dim rngComResult As RangeComparisonResult
        
        If compType = withColumnNames_ Then
            rngComResult = CompareTableWithTableDataHashSet(rng1, CreateTableDataHashSet(rng2))
        Else 'if compType = allRows_ Then
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
