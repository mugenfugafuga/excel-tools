Attribute VB_Name = "FileTools"
Option Explicit

Private selectedFolder_ As String
Private selectedFile_ As String

Public Function SelectFolder( _
        Optional title As String = "select folder" _
    ) As String
    
    With Application.FileDialog(msoFileDialogFolderPicker)
        .title = title
        
        If selectedFolder_ = "" Then
            .InitialFileName = ActiveWorkbook.Path
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

Public Function SelectFile( _
    Optional title As String = "select file" _
    ) As String

    With Application.FileDialog(msoFileDialogFilePicker)
        .title = title
        
        If SelectFile <> "" Then
            .InitialFileName = selectedFile_
        End If
        
        If .Show Then
            selectedFile_ = .SelectedItems(1)
            SelectFile = selectedFile_
        Else
            SelectFile = ""
        End If
    End With
    
End Function

