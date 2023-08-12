VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TableDataHashSet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private count_ As Long
Private columns_() As String
Private hash_ As Scripting.Dictionary

Public Function Initialize(columnNames As Range)
    Dim ret() As String
    
    Dim i As Long
    
    With columnNames.Cells
        ReDim ret(1 To .Count)
    
        For i = 1 To .Count
            ret(i) = CStr(.Item(i).Value2)
        Next i
    End With
    
    columns_ = ret
End Function

Public Function Add(val As Range)
    
    Assert UBound(columns_) = val.Cells.Count

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

Public Function GetValues(colIndexMap As Scripting.Dictionary, val As Range) As RangeList
    Dim ret As New RangeList
    
    Dim key As String
    Dim vs As RangeList
    Dim i As Long
    
    If _
        UBound(columns_) = colIndexMap.Count _
        And UBound(columns_) = val.Cells.Count _
    Then
        key = ToMD5_(colIndexMap, val)
        If hash_.Exists(key) Then
            Set vs = hash_(key)
            
            With vs
                For i = 1 To .Count
                    If IsSame_(val, colIndexMap, .Item(i)) Then
                        ret.Add .Item(i)
                    End If
                Next i
            End With
            
        End If
    End If
    
    Set GetValues = ret
    Set ret = Nothing
End Function

Private Function ToMD5_(colIndexMap As Scripting.Dictionary, val As Range) As String
    Dim ret As String
    
    Dim indx As Long
    Dim i As Long
    
    With val.Cells
        For i = 1 To UBound(columns_)
            indx = colIndexMap(columns_(i))
            ret = ret & CStr(.Item(indx).Value2)
        Next i
        
        ToMD5_ = StringToMD5(ret)
    End With
End Function

Private Function IsSame_(l As Range, colIndexMap As Scripting.Dictionary, r As Range) As Boolean
    If l.Cells.Count <> r.Cells.Count Then
        IsSame_ = False
        Exit Function
    End If
    
    Dim indx As Long
    Dim i As Long
    
    With l.Cells
        For i = 1 To UBound(columns_)
            indx = colIndexMap(columns_(i))
            If .Item(indx).Value2 <> r.Cells.Item(i).Value2 Then
                IsSame_ = False
                Exit Function
            End If
        Next i
    End With
    
    IsSame_ = True
End Function

Private Sub Class_Initialize()
    count_ = 0
    Set hash_ = New Scripting.Dictionary
End Sub

Private Sub Class_Terminate()
    Set hash_ = Nothing
End Sub

