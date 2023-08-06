Attribute VB_Name = "ExcelDumper"
Option Explicit

Public Enum DumpType
    None = 0
    VbModule = 1
    SheetValue = 2
    SheetFormula = 4
End Enum

Public Function DumpExcel( _
    book As Workbook, _
    outputDir As String, _
    Optional excelDumpType As DumpType = (VbModule Or DumpType.SheetValue Or SheetFormula))
    
    If ShouldOutputSheets_(excelDumpType) Then
        Call DumpSheets_(book, outputDir, excelDumpType)
    End If
    
    If (excelDumpType And VbModule) Then
        Call DumpModules_(book, outputDir)
    End If
End Function

Private Function DumpSheet_( _
    sht As Worksheet, _
    filePath As String, _
    excelDumpType As DumpType)
    
    Dim rw As Range
    Dim cll As Range
    Dim addrss As String
    Dim v As String
    
    Dim fn As Integer
    fn = FreeFile
    
    Open filePath For Output As fn
    Rem On Error GoTo FILE_CLOSE_

    For Each rw In sht.UsedRange.Rows
        For Each cll In rw.Cells
            addrss = cll.Address
        
            If (excelDumpType And SheetValue) Then
                Print #fn, addrss & ".Value2:=" & CStr(cll.Value2)
            End If
            
            If (excelDumpType And SheetFormula) Then
                Print #fn, addrss & ".Formula:=" & CStr(cll.Formula)
            End If
        Next cll
    Next rw

FILE_CLOSE_:
    Close fn
    On Error GoTo 0
End Function

Private Function DumpSheets_( _
    book As Workbook, _
    outputDir As String, _
    excelDumpType As DumpType)
    Dim sht As Worksheet
    Dim filePath As String
    
    Dim suffix As String
    suffix = GetSheetFileSuffix_(excelDumpType)
    
    For Each sht In book.Worksheets
        filePath = outputDir & "\" & sht.Name & suffix
        Call DumpSheet_(sht, filePath, excelDumpType)
    Next sht
    
End Function

Private Function GetSheetFileSuffix_(excelDumpType As DumpType) As String
    Dim suffix As String
    
    suffix = "_sheet"
    
    If (excelDumpType And SheetValue) Then suffix = suffix & "_value"
    If (excelDumpType And SheetFormula) Then suffix = suffix & "_function"
    
    GetSheetFileSuffix_ = suffix & ".txt"
End Function

Private Function ShouldOutputSheets_(excelDumpType As DumpType) As Boolean
    ShouldOutputSheets_ = (excelDumpType And SheetValue) Or (excelDumpType And SheetFormula)
End Function

Private Function DumpModules_(book As Workbook, outputDir As String)
    Dim module As VBComponent
    Dim extension As String
    Dim filePath As String
    
    For Each module In book.VBProject.VBComponents
        extension = GetExtension_(module)
        
        If extension <> "" Then
            filePath = outputDir & "\" & module.Name & "." & extension
            Call module.Export(filePath)
        End If
    Next module
    
    Set module = Nothing
End Function

Private Function GetExtension_(module As VBComponent) As String
    If (module.Type = vbext_ct_ClassModule) Then
        GetExtension_ = "cls"
    ElseIf (module.Type = vbext_ct_MSForm) Then
        GetExtension_ = "frm"
    ElseIf (module.Type = vbext_ct_StdModule) Then
        GetExtension_ = "bas"
    ElseIf (module.Type = vbext_ct_Document) Then
        GetExtension_ = "cls"
    Else
        GetExtension_ = ""
    End If
End Function
