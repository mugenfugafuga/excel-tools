Attribute VB_Name = "FileUtilities"
Option Explicit

Private fs_ As Scripting.FileSystemObject
Private initilized_ As Boolean

Private Const TempFileSheet_ As String = "TempFileList"

Public Function GetExtention(filePath As String) As String
    init_
    GetExtention = fs_.GetExtensionName(filePath)
End Function

Public Function GetTempFileName(Optional extension As String = "") As String
    init_
    
    Dim fn As String
    
    If extension = "" Then
        fn = fs_.GetSpecialFolder(TemporaryFolder) & "\" & fs_.GetTempName
    Else
        fn = fs_.GetSpecialFolder(TemporaryFolder) & "\" & fs_.GetTempName & "." & extension
    End If
    
    AppendValueOnBottom_ TempFileSheet_, fn
    
    GetTempFileName = fn
End Function

Private Function AppendValueOnBottom_(sht As String, val As String)
    If SheetExists(sht) = False Then
        Exit Function
    End If
    
    With ThisWorkbook.Worksheets(sht).Range("A1")
        If IsEmpty(.Value) Then
            .Value2 = val
        Else
            .Offset(.CurrentRegion.Rows.Count, 0) = val
        End If
    End With
End Function

Public Function TryDeleteTempFiles()
    Dim undeleteds As New Collection
    Dim rng As Range
    Dim fn As Variant
    
    init_
    
    If SheetExists(TempFileSheet_) = False Then
        Exit Function
    End If
    
    With ThisWorkbook.Worksheets(TempFileSheet_).Range("A1")
        If IsEmpty(.Value) Then
            Exit Function
        End If
        
        For Each rng In .CurrentRegion
            fn = rng.Value2
            If TryDeleteIfExists_(CStr(fn)) = False Then
                undeleteds.Add fn
            End If
        Next rng
        
        .CurrentRegion.Clear
        
        For Each fn In undeleteds
            AppendValueOnBottom_ TempFileSheet_, CStr(fn)
        Next fn
    End With
    
End Function

Private Function TryDeleteIfExists_(fileName As String) As Boolean
    
    If fs_.FileExists(fileName) Then
        On Error Resume Next
        
        fs_.DeleteFile fileName
        
        If Err.Number <> 0 Then
            On Error GoTo 0
        
            Err.Clear
            TryDeleteIfExists_ = False
            
            Exit Function
        End If
        
        On Error GoTo 0
    End If
    
    TryDeleteIfExists_ = True
End Function

Private Function init_()
    If initilized_ = False Then
        Set fs_ = New Scripting.FileSystemObject
        initilized_ = True
    End If
End Function
