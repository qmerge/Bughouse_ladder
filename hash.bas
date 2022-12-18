Attribute VB_Name = "hash"
Option Explicit

' Size of HashTable '
Private m_lHashTableSize As Long

' Pseudorandom cryptographic array used to mix up '
' the input data and create a hash function with  '
' good chance of preventing clashes:              '
Private rand8(0 To 255) As Long
Const hashsize As Integer = 2047
Public hasharray(0 To hashsize) As String
Public hashindex(0 To hashsize) As Integer
Private Sub Class_Initialize()
    Dim i As Long, j As Long
    Dim s As Byte, k As Long
    ' Create a pseudorandom array using the                 '
    ' alleged RC4 algorithm.                                '
    ' See:                                                  '
    ' http://burtleburtle.net/bob/hash/pearson.html         '
    ' This initialisation is to prevent a low repeat count: '
    For i = 0 To 255
        rand8(i) = i
    Next i
    ' Here we go: '
    k = 7
    For j = 0 To 3
        For i = 0 To 255
            s = rand8(i)
            k = (k + s) Mod 256
            rand8(i) = rand8(k)
            rand8(k) = s
        Next i
    Next j
    ' Hash table is 65k in size.  This should remain '
    ' very efficient for > 200,000 items.            '
    m_lHashTableSize = hashsize + 1
End Sub
'method 0==add
'method 1==check
'method 2==delete
Public Function hash(ByVal skey As String, sval As String) As String
    Dim b() As Byte
    Dim i As Long
    Dim lKeyVal As Long
    Dim h1 As Long, h2 As Long
    Dim method As Integer
    If sval = "" Then method = 1
    If sval = "delete" Then method = 2
    ' using hashing algorithm "Variable String Exclusive-Or '
    ' method (tablesize up to 65,536)".  See                '
    ' Dictionaries ->Hash Tables section at                 '
    ' http://members.xoom.com/thomasn/s_man.htm             '
    b = skey
    h1 = b(0): h2 = h1 + 1
    For i = 1 To UBound(b)
        h1 = rand8(h1 Xor b(i))
        h2 = rand8(h2 Xor b(i))
    Next
    lKeyVal = h1 * &HFFFF& + h2
    If lKeyVal = 0 Then lKeyVal = 1
    i = lKeyVal Mod m_lHashTableSize
    Do
        If i = hashindex(lKeyVal) Then
            If method = 2 Then
                hashindex(lKeyVal) = 0
                hasharray(lKeyVal) = ""
            End If
            Exit Do
        End If
        If 0 = hashindex(lKeyVal) Then
            If method = 0 Then
                hashindex(lKeyVal) = i
                hasharray(lKeyVal) = sval
            End If
            Exit Do
        End If
        lKeyVal = lKeyVal + 1
        If lKeyVal = m_lHashTableSize Then lKeyVal = 0
    Loop While 1
    If 0 = hashindex(lKeyVal) Then
        hash = ""
        Exit Function
    End If
    hash = hasharray(lKeyVal)
End Function
