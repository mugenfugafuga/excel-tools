Attribute VB_Name = "CharsetConverter"
Option Explicit

Public Const UTF8Charset As String = "UTF-8"
Public Const SJISCharset As String = "Shift-JIS"

Public Function ConvertWithNewCharset( _
    sourceFile As String, sourceCharset, _
    targetFile As String, targetCharset)
    
    Dim sourceStream As New ADODB.Stream
    With sourceStream
        .Type = adTypeText
        .Charset = sourceCharset
        .Open
        
        .LoadFromFile (sourceFile)
    End With
    
    Dim targetStream As New ADODB.Stream
    With targetStream
        .Type = adTypeText
        .Charset = targetCharset
        .Open
        .WriteText (sourceStream.ReadText)
        
        .SaveToFile targetFile, adSaveCreateOverWrite
    End With
    
    sourceStream.Close
    Set sourceStream = Nothing
    
    targetStream.Close
    Set targetStream = Nothing
End Function
