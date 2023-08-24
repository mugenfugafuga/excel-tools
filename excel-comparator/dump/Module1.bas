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
    fileName = SelectFile("①UTF8のファイルを選択してください。")
    
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
    fileName = SelectFile("①SJISのファイルを選択してください。")
    
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
    fileName = SelectFile("①UTF8のファイルを選択してください。")
    
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
    fileName = SelectFile("①SJISのファイルを選択してください。")
    
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
    fileName = SelectFile("①UTF8のファイルを選択してください。")
    
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
    fileName = SelectFile("①SJISのファイルを選択してください。")
    
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
    fileName = SelectFile("②UTF8のファイルを選択してください。")
    
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
    fileName = SelectFile("②SJISのファイルを選択してください。")
    
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
    
    Const baseRow = 7
    
    Set wb1 = OpenBook_(file1)
    With wb1
        Set no1 = CopySheetAfterTarget(.Sheets(1), resultbook.Sheets(1))
        no1.Name = "①"
        no1.Cells.EntireColumn.AutoFit
        Set no1rng = no1.UsedRange
        
        .Close
        Set wb1 = Nothing
    End With
    
    Set wb2 = OpenBook_(file2)
    With wb2
        Set no2 = CopySheetAfterTarget(.Sheets(1), no1)
        no2.Name = "②"
        no2.Cells.EntireColumn.AutoFit
        Set no2rng = no2.UsedRange
    
        .Close
        Set wb2 = Nothing
    End With
    
    Dim twoM As TwoMatrices
    Dim lr As Range, ldt As Range
    Dim rr As Range, rdt As Range
    Dim mtchfuncs As Range, numdiffunc As Range
    Dim tlrnc As Range
    Dim unmtch As Range, witlr As Range
    Dim ootlr As Range, lorr As Range
    
    With resultbook.Sheets(1)
        twoM = EditPointsTo2Matrices(DiffExcel(no1.UsedRange, no2.UsedRange, ignoreNum, tolernce))
        
        With .Cells(5, 6)
            .Offset(0, 0) = "許容値": .Offset(0, 1) = 0.00001
            Set tlrnc = .Offset(0, 1)
        End With
        
        Set lr = PrintVariantOnRange(twoM.Left, .Cells(baseRow, 6))
        GetAddressArea_(lr).Interior.Color = RGB(252, 228, 214)
        Set ldt = GetDataArea_(lr)
        
        
        Set rr = PrintVariantOnRange(twoM.Right, lr.Cells(1, lr.Columns.Count).Offset(0, 2))
        GetAddressArea_(rr).Interior.Color = RGB(214, 220, 228)
        Set rdt = GetDataArea_(rr)
        
        Set mtchfuncs = PrintVariantOnRange(GetMatchFunctions_(ldt, rdt), rr.Cells(1, rr.Columns.Count).Offset(0, 2))
        Set numdiffunc = PrintVariantOnRange(GetNumericDiffFunctions_(ldt, rdt), mtchfuncs.Cells(1, mtchfuncs.Columns.Count).Offset(0, 2))
        
        Set unmtch = PrintVariantOnRange(GetCountUnmatchFunctions_(mtchfuncs), .Cells(baseRow, 1))
        Set witlr = PrintVariantOnRange(GetCountWithinToleranceFunctions_(numdiffunc, tlrnc), .Cells(baseRow, 2))
        Set ootlr = PrintVariantOnRange(GetCountOutOfToleranceFunctions_(unmtch, witlr), .Cells(baseRow, 3))
        Set lorr = PrintVariantOnRange(GetLeftOrRightFunctions_(lr, rr), .Cells(baseRow, 4))
        
        SetDataBar_ witlr, RGB(200, 255, 255)
        SetDataBar_ ootlr, RGB(255, 0, 0)
        
        UpdateFormatConditions_ ldt, rdt, tlrnc, lorr
        
        With .Cells(4, 1)
            .Offset(0, 0) = "完全一致":     .Offset(0, 1) = "=IF(SUM(" & SkipFirstRow_(unmtch).Address(RowAbsolute:=False, ColumnAbsolute:=False) & ")=0,""OK"",""NG"")"
            .Offset(1, 0) = "許容値内一致": .Offset(1, 1) = "=IF(SUM(" & SkipFirstRow_(ootlr).Address(RowAbsolute:=False, ColumnAbsolute:=False) & ")=0,""OK"",""NG"")"
        End With
        
        .Name = "結果"
        .Cells.EntireColumn.AutoFit
        
        With .Cells(1, 1)
            .Offset(0, 0) = "①": .Offset(0, 0).HorizontalAlignment = xlRight: .Offset(0, 1) = file1.Path
            .Offset(1, 0) = "②": .Offset(1, 0).HorizontalAlignment = xlRight: .Offset(1, 1) = file2.Path
        End With
        
        .Columns(5).ColumnWidth = 2
        .Rows(baseRow).Font.Bold = True
         
        .Activate
        
        .Cells(baseRow + 1, 6).Select
        ActiveWindow.FreezePanes = True
    End With
End Function

Private Function SetDataBar_(rng As Range, rgbnum As Long)
    With rng.FormatConditions
        .Delete
        
        .AddDatabar
        With .Item(1)
            .BarBorder.Type = xlDataBarBorderSolid
            .BarColor.Color = rgbnum
            .BarBorder.Color.Color = rgbnum
        End With
    End With
End Function

Private Function SkipFirstRow_(rng As Range) As Range
    With rng
        Set SkipFirstRow_ = Range(.Cells(2, 1), .Cells(.Rows.Count, .Columns.Count))
    End With
End Function

Private Function GetLeftOrRightFunctions_(lft As Range, rght As Range) As Variant
    Dim ret() As Variant
    
    Dim rnum As Long, cnum As Long
    rnum = lft.Rows.Count
    cnum = lft.Columns.Count
    
    ReDim ret(1 To rnum, 1 To 1)
    
    Dim ladd As String
    Dim radd As String
    
    Dim r As Long
    
    ret(1, 1) = "比較結果"
    
    For r = 2 To rnum
        With lft.Rows(r)
            ladd = Range(.Cells(1, 2), .Cells(1, cnum)).Address(RowAbsolute:=False, ColumnAbsolute:=False)
        End With
        
        With rght.Rows(r)
            radd = Range(.Cells(1, 2), .Cells(1, cnum)).Address(RowAbsolute:=False, ColumnAbsolute:=False)
        End With
        
        ret(r, 1) = "=IF(COUNTA(" & ladd & ")=0,IF(COUNTA(" & radd & ")=0,""値なし"",""Right""),IF(COUNTA(" & radd & ")=0,""Left"",""Same""))"
    Next r
    
    GetLeftOrRightFunctions_ = ret
End Function

Private Function GetCountUnmatchFunctions_(mtchRange As Range) As Variant
    Dim ret() As Variant
    
    Dim rnum As Long, r As Long
    rnum = mtchRange.Rows.Count
    
    ReDim ret(1 To rnum, 1 To 1)
    
    ret(1, 1) = "不一致"
    
    With mtchRange
        For r = 2 To rnum
            ret(r, 1) = "=COUNTIF(" & .Rows(r).Address(RowAbsolute:=False, ColumnAbsolute:=False) & ",FALSE)"
        Next r
    End With
    
    GetCountUnmatchFunctions_ = ret
End Function

Private Function GetCountWithinToleranceFunctions_(numdifRange As Range, tlrnc As Range) As Variant
    Dim ret() As Variant
    
    Dim rnum As Long, r As Long
    rnum = numdifRange.Rows.Count
    
    ReDim ret(1 To rnum, 1 To 1)
    
    ret(1, 1) = "許容値内"
    
    Dim adrs As String
    With numdifRange
        For r = 2 To rnum
            adrs = .Rows(r).Address(RowAbsolute:=False, ColumnAbsolute:=False)
            ret(r, 1) = "=COUNTIFS(" & adrs & ","">0""," & adrs & ",""<=""&" & tlrnc.Address & ")"
        Next r
    End With
    
    GetCountWithinToleranceFunctions_ = ret
End Function

Private Function GetCountOutOfToleranceFunctions_(unmatchRange As Range, numdifRange As Range) As Variant
    Dim ret() As Variant
    
    Dim rnum As Long, r As Long
    rnum = numdifRange.Rows.Count
    
    ReDim ret(1 To rnum, 1 To 1)
    
    ret(1, 1) = "許容値外"
    
    Dim adrs As String
    With numdifRange
        For r = 2 To rnum
            ret(r, 1) = "=" & _
                unmatchRange.Cells(r, 1).Address(RowAbsolute:=False, ColumnAbsolute:=False) & "-" & _
                .Cells(r, 1).Address(RowAbsolute:=False, ColumnAbsolute:=False)
        Next r
    End With
    
    GetCountOutOfToleranceFunctions_ = ret
End Function
Private Function GetMatchFunctions_(ldata As Range, rdata As Range) As Variant
    Dim rnum As Long, cnum As Long
    
    rnum = ldata.Rows.Count
    cnum = ldata.Columns.Count
    
    Dim ret() As Variant
    ReDim ret(0 To rnum, 1 To cnum)
    
    Dim r As Long, c As Long
    
    For c = 1 To cnum
        ret(0, c) = "match" & c
    Next c
    
    For r = 1 To rnum
        For c = 1 To cnum
            ret(r, c) = "=" & _
                ldata.Cells(r, c).Address(RowAbsolute:=False, ColumnAbsolute:=False) & "=" & _
                rdata.Cells(r, c).Address(RowAbsolute:=False, ColumnAbsolute:=False)
        Next c
    Next r
    
    GetMatchFunctions_ = ret
End Function

Private Function GetNumericDiffFunctions_(ldata As Range, rdata As Range) As Variant
    Dim rnum As Long, cnum As Long
    
    rnum = ldata.Rows.Count
    cnum = ldata.Columns.Count
    
    Dim ret() As Variant
    ReDim ret(0 To rnum, 1 To cnum)
    
    Dim r As Long, c As Long
    
    Dim rad As String, lad As String
    
    For c = 1 To cnum
        ret(0, c) = "diff" & c
    Next c
    
    For r = 1 To rnum
        For c = 1 To cnum
            lad = ldata.Cells(r, c).Address(RowAbsolute:=False, ColumnAbsolute:=False)
            rad = rdata.Cells(r, c).Address(RowAbsolute:=False, ColumnAbsolute:=False)
            
            ret(r, c) = "=IF(AND(ISNUMBER(" & lad & "),ISNUMBER(" & rad & ")),ABS(" & lad & "-" & rad & "),0)"
        Next c
    Next r
    
    GetNumericDiffFunctions_ = ret
End Function

Private Function UpdateFormatConditions_( _
    lft As Range, _
    rght As Range, _
    tlrnc As Range, _
    lorr As Range)
    
    Dim ltop As String, rtop As String
    Dim trl As String, lr As String
    
    ltop = lft.Cells(1, 1).Address(RowAbsolute:=False, ColumnAbsolute:=False)
    rtop = rght.Cells(1, 1).Address(RowAbsolute:=False, ColumnAbsolute:=False)
    trl = tlrnc.Cells(1, 1).Address
    lr = lorr.Cells(2, 1).Address(RowAbsolute:=False)
    
    With lft.FormatConditions
        .Delete
        
        .Add _
            Type:=xlExpression, _
            Formula1:="=" & lr & "=""Right"""
        .Item(1).Interior.Color = RGB(255, 255, 0)
        
        .Add _
            Type:=xlExpression, _
            Formula1:="=AND(ISNUMBER(" & ltop & "),ISNUMBER(" & rtop & "),0<ABS(" & ltop & "-" & rtop & "),ABS(" & ltop & "-" & rtop & ")<=" & trl & ")"
        .Item(2).Interior.Color = RGB(200, 255, 255)
        
        .Add _
            Type:=xlExpression, _
            Formula1:="=AND(ISNUMBER(" & ltop & "),ISNUMBER(" & rtop & ")," & trl & "<ABS(" & ltop & "-" & rtop & "))"
        .Item(3).Interior.Color = RGB(255, 0, 0)
        
        .Add Type:=xlCellValue, Operator:=xlNotEqual, Formula1:="=" & rtop
        .Item(4).Interior.Color = RGB(240, 100, 100)
    End With

    With rght.FormatConditions
        .Delete
        .Add _
            Type:=xlExpression, _
            Formula1:="=" & lr & "=""Left"""
        .Item(1).Interior.Color = RGB(255, 255, 0)

        .Add _
            Type:=xlExpression, _
            Formula1:="=AND(ISNUMBER(" & ltop & "),ISNUMBER(" & rtop & "),0<ABS(" & ltop & "-" & rtop & "),ABS(" & ltop & "-" & rtop & ")<=" & trl & ")"
        .Item(2).Interior.Color = RGB(173, 255, 47)
        
        .Add _
            Type:=xlExpression, _
            Formula1:="=AND(ISNUMBER(" & ltop & "),ISNUMBER(" & rtop & ")," & trl & "<ABS(" & ltop & "-" & rtop & "))"
        .Item(3).Interior.Color = RGB(255, 69, 0)
        
        .Add Type:=xlCellValue, Operator:=xlNotEqual, Formula1:="=" & ltop
        .Item(4).Interior.Color = RGB(255, 140, 0)
    End With

End Function

Private Function GetAddressArea_(rng As Range) As Range
    With rng
        Set GetAddressArea_ = Range(.Cells(2, 1), .Cells(.Rows.Count, 1))
    End With
End Function

Private Function GetDataArea_(rng As Range) As Range
    With rng
        Set GetDataArea_ = Range(.Cells(2, 2), .Cells(.Rows.Count, .Columns.Count))
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
        no1.Name = "①"
        no1.Cells.EntireColumn.AutoFit
        Set no1rng = no1.UsedRange
        
        .Close
        Set wb1 = Nothing
    End With
    
    Set wb2 = OpenBook_(file2)
    With wb2
        Set no2 = CopySheetAfterTarget(.Sheets(1), no1)
        no2.Name = "②"
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
            "①を②に突合せた結果", "①にあって②にない", xlThemeColorAccent6)
        Set run = com1With2Result.Rest
    End If
    
    If ctype <> No1WithNo2 Then
        com2With1Result = AddComparisonResultAfterTarget_( _
            compType, _
            run, no2rng, no1rng, _
            "②を①に突合せた結果", "②にあって①にない", xlThemeColorAccent5)
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
                .Offset(1, 0) = "①": .Offset(1, 1) = sheet1.UsedRange.Rows.Count - 1: .Offset(1, 2) = file1.Path
                .Offset(2, 0) = "②": .Offset(2, 1) = sheet2.UsedRange.Rows.Count - 1: .Offset(2, 2) = file2.Path
            Else 'if compType = allRows_ Then
                .Offset(1, 0) = "①": .Offset(1, 1) = sheet1.UsedRange.Rows.Count: .Offset(1, 2) = file1.Path
                .Offset(2, 0) = "②": .Offset(2, 1) = sheet2.UsedRange.Rows.Count: .Offset(2, 2) = file2.Path
            End If
        End With
        
        With .Offset(4, 0)
            .Offset(0, 0) = "①を②に突合せた結果"
            
            .Offset(1, 0) = "①にあって②にない"
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
            .Offset(0, 0) = "②を①に突合せた結果"
            
            .Offset(1, 0) = "②にあって①にない"
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
