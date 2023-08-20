Attribute VB_Name = "ExcelUtilities"
Option Explicit

Public Function NewWorkBook(Optional sheetNum As Integer = 1) As Workbook
    Dim num As Integer
    num = Application.SheetsInNewWorkbook
    
    Application.SheetsInNewWorkbook = sheetNum
    
    Set NewWorkBook = Workbooks.Add
    
    Application.SheetsInNewWorkbook = num
End Function

Public Function AddSheetAfter(sht As Worksheet) As Worksheet
    Worksheets.Add after:=sht
    Set AddSheetAfter = sht.Next
End Function

Public Function CopySheetAfterTarget(copySheet As Worksheet, targetSheet As Worksheet) As Worksheet
    copySheet.Copy after:=targetSheet
    Set CopySheetAfterTarget = targetSheet.Next
End Function

Public Function SheetExists(sheetName As String, Optional book As Workbook = Nothing) As Boolean
    Dim bk As Workbook
    Dim Sheet As Worksheet
    
    If book Is Nothing Then
        Set bk = ThisWorkbook
    Else
        Set bk = book
    End If
    
    On Error Resume Next
    Set Sheet = bk.Worksheets(sheetName)
    On Error GoTo 0
    If Not Sheet Is Nothing Then
        SheetExists = True
    Else
        SheetExists = False
    End If
End Function
