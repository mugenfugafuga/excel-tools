Attribute VB_Name = "RangeComparator"
Option Explicit

Public Type MatchResult
    Value As Range
    Matchs() As Range
End Type

Public Type MatchResults
    Reserved As Long
    Count As Long
    Results() As MatchResult
End Type

Public Type RangeComparisonResult
    Matchs As MatchResults
    Rest() As Range
End Type

Public Function CompareTableWithRangeHasSet(table As Range, hashSet As RangeHashSet) As RangeComparisonResult
    Dim vs As RangeList
    Dim rw As Range
    
    Dim rst As New RangeList
    
    With CompareTableWithRangeHasSet
        For Each rw In table.Rows
            Set vs = hashSet.GetValues(rw)
            
            If vs.Count = 0 Then
                rst.Add rw
            Else
                Reserve_ .Matchs
                
                With .Matchs
                    .Count = .Count + 1
                    With .Results(.Count)
                        Set .Value = rw
                        .Matchs = vs.Items
                    End With
                End With
            End If
            
            Set vs = Nothing
        Next rw
        
        .Rest = rst.Items
    End With
End Function

Public Function CompareTableWithTableDataHashSet(table As Range, hashSet As TableDataHashSet) As RangeComparisonResult
    Dim vs As RangeList
    Dim rw As Range
    
    Dim rst As New RangeList
    
    Dim columIndexMap As Scripting.Dictionary
    
    Dim firstRecord As Boolean
    firstRecord = True

    With CompareTableWithTableDataHashSet
        For Each rw In table.Rows
            If firstRecord Then
                Set columIndexMap = CreateColumnIndexMap(rw)
                firstRecord = False
            Else
                Set vs = hashSet.GetValues(columIndexMap, rw)
                
                If vs.Count = 0 Then
                    rst.Add rw
                Else
                    Reserve_ .Matchs
                    
                    With .Matchs
                        .Count = .Count + 1
                        With .Results(.Count)
                            Set .Value = rw
                            .Matchs = vs.Items
                        End With
                    End With
                End If
                
                Set vs = Nothing
            End If
        Next rw
        
        .Rest = rst.Items
    End With
End Function

Private Function Reserve_(ByRef Results As MatchResults)
    With Results
        If .Reserved = 0 Then
            .Reserved = 3
            ReDim .Results(1 To .Reserved)
            Exit Function
        End If
        
        If .Reserved < .Count + 1 Then
            If .Reserved < 30 Then
                .Reserved = .Reserved + 10
            ElseIf .Reserved < 300 Then
                .Reserved = .Reserved + 100
            ElseIf .Reserved < 1500 Then
                .Reserved = .Reserved + 300
            ElseIf .Reserved < 3000 Then
                .Reserved = .Reserved + 1000
            Else
                .Reserved = .Reserved + 10000
            End If
        
            ReDim Preserve .Results(1 To .Reserved)
        End If
    End With
End Function

