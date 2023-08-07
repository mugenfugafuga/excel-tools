VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "RangeList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private count_ As Integer
Private reserved_ As Integer
Private values_() As Range

Public Property Get Count() As Integer
    Count = count_
End Property

Public Function Add(val As Range)
    If reserved_ < count_ + 1 Then
        If reserved_ < 10 Then
            reserved_ = reserved_ + 3
        ElseIf reserved_ < 30 Then
            reserved_ = reserved_ + 10
        ElseIf reserved_ < 300 Then
            reserved_ = reserved_ + 100
        ElseIf reserved_ < 1500 Then
            reserved_ = reserved_ + 300
        ElseIf reserved_ < 3000 Then
            reserved_ = reserved_ + 1000
        Else
            reserved_ = reserved_ + 10000
        End If
        
        ReDim Preserve values_(1 To reserved_)
    End If
    
    count_ = count_ + 1
    Set values_(count_) = val
End Function

Public Function Item(index As Integer) As Range
    Set Item = values_(index)
End Function

Public Property Get Items() As Range()
    Dim vs() As Range
    ReDim vs(1 To count_)
    
    Dim i As Integer
    For i = 1 To Count
        Set vs(i) = values_(i)
    Next i
    
    Items = vs
End Property

Private Sub Class_Initialize()
    count_ = 0
    reserved_ = 1
    ReDim values_(1 To reserved_)
End Sub

Private Sub Class_Terminate()
    Dim i As Integer
    For i = 1 To count_
        Set values_(i) = Nothing
    Next i
End Sub
