VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "RangeHashSet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private count_ As Long
Private hash_ As Scripting.Dictionary

Public Property Get Count() As Long
    Count = count_
End Property

Public Function Add(val As Range)
    Dim key As String
    key = RangeToMD5(val)
        
    Dim list As RangeList
    
    If hash_.Exists(key) Then
        Set list = hash_(key)
    Else
        Set list = New RangeList
        hash_.Add key, list
    End If
    
    list.Add val
    count_ = count_ + 1
End Function

Public Function GetValues(val As Range) As RangeList
    Dim ret As New RangeList
    
    Dim key As String
    Dim vs As RangeList
    Dim i As Long
    
    key = RangeToMD5(val)
    If hash_.Exists(key) Then
        Set vs = hash_(key)
        
        With vs
            For i = 1 To .Count
                If IsSame_(val, .Item(i)) Then
                    ret.Add .Item(i)
                End If
            Next i
        End With
        
    End If
    
    Set GetValues = ret
    Set ret = Nothing
End Function

Private Function IsSame_(l As Range, r As Range) As Boolean
    If l.Cells.Count <> r.Cells.Count Then
        IsSame_ = False
        Exit Function
    End If
    
    Dim i As Long
    
    For i = 1 To l.Cells.Count
        If l.Cells.Item(i).Value2 <> r.Cells.Item(i).Value2 Then
        IsSame_ = False
        Exit Function
        End If
    Next i
    
    IsSame_ = True
End Function

Private Sub Class_Initialize()
    count_ = 0
    Set hash_ = New Scripting.Dictionary
End Sub

Private Sub Class_Terminate()
    Set hash_ = Nothing
End Sub
