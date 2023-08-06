Attribute VB_Name = "ExcelDumpExecution"
Option Explicit

Public Sub DoExcelDumpAll()
Attribute DoExcelDumpAll.VB_ProcData.VB_Invoke_Func = "d\n14"
    Call Dump_((DumpType.VbModule Or DumpType.SheetValue Or DumpType.SheetFormula))
End Sub

Public Sub DoExcelDumpWithoutValue()
Attribute DoExcelDumpWithoutValue.VB_ProcData.VB_Invoke_Func = "D\n14"
    Call Dump_((DumpType.VbModule Or DumpType.SheetFormula))
End Sub

Private Function Dump_(excelDumpType As DumpType)
    Dim outputDir As String
    
    outputDir = SelectFolder("dump�t�@�C���̏o�̓t�H���_��I�����Ă��������B")
    
    If outputDir <> "" Then
        Call DumpExcel(ActiveWorkbook, outputDir, excelDumpType)
    End If
End Function
