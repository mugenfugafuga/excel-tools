Attribute VB_Name = "MD5"
Option Explicit

Private utf8 As Object
Private md5 As Object

Private initialized As Boolean

Public Function RangeToMD5(rng As Range) As String
    Dim v As String
    Dim cll As Range
    
    v = ""
    For Each cll In rng.Cells
        v = v & CStr(cll.Value2)
    Next cll
    
    RangeToMD5 = StringToMD5(v)
End Function

Public Function StringToMD5(str As String) As String
    Dim bytes() As Byte
    Dim hash() As Byte
    Dim i As Integer
    Dim res As String

    init_
    
    bytes = utf8.GetBytes_4(str)
    hash = md5.ComputeHash_2(bytes)

    For i = LBound(hash) To UBound(hash)
        res = res & LCase(Right("0" & Hex(hash(i)), 2))
    Next i

    StringToMD5 = LCase(res)
End Function

Private Function init_()
    If Not initialized Then
        Set utf8 = CreateObject("System.Text.UTF8Encoding")
        Set md5 = CreateObject("System.Security.Cryptography.MD5CryptoServiceProvider")
        initialized = True
    End If
End Function

