Attribute VB_Name = "FileTools"
Option Explicit

Private selectedFolder_ As String

Public Function SelectFolder( _
        Optional title As String = "select folder" _
    ) As String
    
    With Application.FileDialog(msoFileDialogFolderPicker)
        .title = title
        
        If selectedFolder_ = "" Then
            .InitialFileName = ActiveWorkbook.path
        Else
            .InitialFileName = selectedFolder_
        End If
        
        If .Show Then
            selectedFolder_ = .SelectedItems(1)
            SelectFolder = selectedFolder_
        Else
            SelectFolder = ""
        End If
    End With
    
End Function

