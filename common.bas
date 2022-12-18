Attribute VB_Name = "common"
Option Explicit
Global Const grows_max As Integer = 200
Global Const gcols As Integer = 44
Global Const group_field As Integer = 0
Global Const last_name_field As Integer = 1
Global Const first_name_field As Integer = 2
Global Const rating_field As Integer = 3
Global Const ranking_field As Integer = 4
Global Const nrating_field As Integer = 5
Global Const grade_field As Integer = 6
Global Const Games_field  As Integer = 7
Global Const attendance_field  As Integer = 8
Global Const phone_field  As Integer = 9
Global Const info_field  As Integer = 10
Global Const school_field  As Integer = 11
Global Const room_field  As Integer = 12
Global Const last_param_field  As Integer = 12
Global password_set As Integer
Global Const rating_field_size As Integer = 500
Global Const ascz As Integer = 122
Global Const ascsp As Integer = 32
Global Const charqt As String * 1 = """"
Global Const GROUP_CODES As String = "A1xAxBxCxDxExFxGxHxIxZx   "
Global Game_Result(0 To 2) As String
Global error_count As Integer
Global Const error_max As Integer = 10
Global Const OtherSize As Integer = 10
Global Const A1size As Integer = 8
Global reports(0 To 1) As Form
Global last_commit As String
' Size of HashTable '
Private m_lHashTableSize As Long
' Pseudorandom cryptographic array used to mix up '
' the input data and create a hash function with  '
' good chance of preventing clashes:              '
Private rand8(0 To 255) As Long
' Global Const begin_hashsize As Integer = 255 '
Global Const begin_hashsize As Integer = 16383
Global hashsize As Integer
Public hasharray() As String
Public hashkeyarray() As String
Public hashindex() As Long
' This program is free software; you can redistribute it and/or               '
' modify it under the terms of the GNU General Public License                 '
' as published by the Free Software Foundation; either version 2              '
' of the License, or (at your option) any later version.                      '
'                                                                             '
' This program is distributed in the hope that it will be useful,             '
' but WITHOUT ANY WARRANTY; without even the implied warranty of              '
' MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the               '
' GNU General Public License for more details.                                '
'                                                                             '
' You should have received a copy of the GNU General Public License           '
' along with this program; if not, write to the Free Software                 '
' Foundation, Inc., 59 Temple Place - Suite 330, Boston, MA  02111-1307, USA. '
' Copyright Matt Mahowald 2008                                                    '
 Global Idle_Timer_Tag As Integer
 Global in_recalc As Integer
 Global entry_size As Integer
 Global group_mapping(10) As Integer
 Global rating(grows_max) As Double
 Global last_entry(grows_max) As Integer
 Global ispresent() As Integer
 Global row2RR() As Integer
 Global row2RR_count As Integer
 ' Dim rating_sort(grows_max) As Integer '
 Global sort_by As Integer
 Global needs_resize As Boolean
Global merr As Double
Global k_val_base As Double
Global AppZoom  As Double
Global drag_from As Integer
Global drag_shift As Integer
Global quick As Integer
Global save_flag As Integer
Global num_games_added As Integer
Global grows As Integer
Global mpassword$, opassword$
Global Show_Ratings_value(10) As Boolean
Global p2r(grows_max) As Integer
Global r2p(grows_max) As Integer
Global Const player_size As Long = 128
Global Const game_size As Long = 4
Global Const sort_rank As Integer = 0
Global Const Sort_Name As Integer = 1
Global Const Sort_First_Name As Integer = 2
Global Const Sort_Rating As Integer = 3
' Global Const result_string As String = "0123456789" '
Global Const result_string As String = "OLDWXYZ__________"
Global players(-1 To 4) As Integer, scores(0 To 2) As Integer, quick_entry As Integer
Global playerOrRow As Boolean
Public Sub player2row(pl() As Integer)
    Dim i As Integer
    For i = 0 To 4
        pl(i) = p2r(pl(i))
    Next
    playerOrRow = False
End Sub
Public Sub row2player(pl() As Integer)
    Dim i As Integer
    For i = 0 To 4
        pl(i) = r2p(pl(i))
    Next
    playerOrRow = True
End Sub
Public Function long2string(ByVal game As Long) As String
    Dim res As String
    Dim s As String
    res = Str$(game Mod player_size)
    game = game \ player_size
    res = res + ":" + Str$(game Mod player_size)
    game = game \ player_size
    res = res + Mid$(result_string, (game Mod game_size) + 1, 1)
    game = game \ game_size
    s = Mid$(result_string, (game Mod game_size) + 1, 1)
    If (s <> "O") Then res = res + s
    game = game \ game_size
    res = res + Str$(game Mod player_size)
    game = game \ player_size
    res = res + ":" + Str$(game Mod player_size)
    res = Replace(res, " ", "")
    res = Replace(res, ":0", "")
    long2string = res
End Function
Public Function string2long(game As String) As Long
    string2long = parse_entry(game, players, scores, quick_entry)
End Function
Public Function formula(my_rating As Double, opponents_rating As Double) As Double
    formula = 1# / (1# + 10# ^ ((Abs(opponents_rating) - Abs(my_rating)) / 400#))
End Function
Function get_ladder_name() As String
    Dim my_file$
    my_file$ = CurDir
    my_file$ = Right(my_file$, Len(my_file$) - InStrRev(my_file$, "\"))
    get_ladder_name = my_file$
End Function
Function entry2string(players() As Integer, score() As Integer, quick_entry As Integer) As String
    ' Stop                           ' not tested '
    If players(0) > players(1) Then
        swapint players(0), players(1)
    End If
    If players(3) > players(4) Then
        swapint players(3), players(4)
    End If
    Dim res As String
    res = Str$(players(0))
    res = res + ":" + Str$(players(1))
    res = res + Mid$(result_string, score(0), 1)
    If score(1) > 0 Then res = res + Mid$(result_string, score(1), 1)
    res = res + Str$(players(3))
    res = res + ":" + Str$(players(4))
    entry2string = res
End Function
Function parse_entry(my_text$, players() As Integer, score() As Integer, quick_entry As Integer) As Long
    Dim i As Integer
    Dim num_or_char As Integer     ' num is 0, char is 1 '
    Dim strlen As Integer
    Dim mychar As String * 1
    Dim myasc As Integer
    Dim entry As Integer
    Dim is_num As Integer
    Dim entry_string As String
    Dim results$
    Dim error_num As Integer
    quick_entry = 0
    For i = 0 To 4
        players(i) = 0
    Next
    playerOrRow = True
    ' start with number '
    strlen = Len(my_text$)
    If (strlen < 2) Then Exit Function
    For i = 1 To strlen
        mychar = Mid$(my_text$, i, 1)
        myasc = Asc(mychar)
        If myasc > 33 Then
            If myasc = 95 Then
                If strlen = i Then Exit For
                error_num = 1
                'Stop
                ' Ret = 1            ' we need to be called again '
                Exit For
            End If
            If (myasc >= 48) And (myasc <= 57) Then
                is_num = 0
            Else
                If myasc = 58 Then
                    entry = entry + 1
                    entry_string = ""
                    GoTo continue_for
                ElseIf mychar = "W" Or mychar = "L" Or mychar = "D" Then
                    If entry < 1 Then
                        entry = 1
                    End If
                    If entry > 2 Then
                        entry = 2
                    End If
                Else
                    error_num = 2
                End If
                is_num = 1
            End If
            If num_or_char <> is_num Then
                entry = entry + 1
                entry_string = ""
                num_or_char = is_num
            End If
            entry_string = entry_string + mychar
            If is_num = 0 Then
                players(entry) = Val(entry_string)
                If players(entry) > grows_max Then
                    error_num = 9
                    Exit For
                End If
            Else
                results$ = entry_string
             End If
        End If    '}-> If myasc > 33 Then
continue_for:
    Next i    '}-> For i = 1 To strlen
    ' Stop '
    score(0) = 0
    score(1) = 0
    score(0) = InStr(result_string, Mid$(results$, 1, 1)) - 1
    score(1) = InStr(result_string, Mid$(results$, 2, 1)) - 1
    If score(0) < 0 Then score(0) = 0
    If score(1) < 0 Then score(1) = 0
    If (players(1) > 0) Then
        If players(0) > players(1) Then
            swapint players(0), players(1)
        End If
        If players(3) > players(4) Then
            swapint players(3), players(4)
        End If
    End If
    If players(0) > players(3) Then
        swapint players(0), players(3)
        swapint players(1), players(4)
        score(0) = 4 - score(0)
        If score(1) > 0 Then score(1) = 4 - score(1)
    End If
    Dim res As Long
    res = players(4)
    res = res * player_size
    res = res + players(3)
    res = res * game_size
    res = res + score(1)
    res = res * game_size
    res = res + score(0)
    res = res * player_size
    res = res + players(1)
    res = res * player_size
    res = res + players(0)
    If players(1) > 0 And players(4) = 0 Then
        error_num = 7
    End If
    If players(0) = players(3) Then
        parse_entry = -4 'duplicate player
        Exit Function
    End If
    If error_num <> 0 Or players(0) = 0 Or players(3) = 0 Or score(0) < 0 Or score(1) < 0 Then
        If error_num = 0 Then
            parse_entry = -3
        Else
            parse_entry = -error_num
        End If
    Else
        parse_entry = res
    End If
End Function
Public Sub hash_Initialize()
    Dim i As Long, j As Long
    Dim s As Byte, k As Long
    reset_hash (begin_hashsize)
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
Public Sub reset_hash(newhashsize As Integer)
    If (newhashsize > 0) Then hashsize = newhashsize
    ReDim hasharray(0 To hashsize) As String
    ReDim hashkeyarray(0 To hashsize) As String
    ReDim hashindex(0 To hashsize) As Long
End Sub
' Public Function hash_function(ByVal skey As String) As Long '
' Dim b() As Byte                                             '
' Dim i As Long                                               '
' Dim lKeyVal As Long                                         '
' Dim h1 As Long, h2 As Long                                  '
' ' using hashing algorithm "Variable String Exclusive-Or     '
' ' method (tablesize up to 65,536)".  See                    '
' ' Dictionaries ->Hash Tables section at                     '
' ' http://members.xoom.com/thomasn/s_man.htm                 '
' b = skey                                                    '
' h1 = b(0): h2 = h1 + 1                                      '
' For i = 0 To UBound(b)                                      '
' If b(i) >= 48 And b(i) <= 57 Then                           '
' h1 = rand8(h2 Xor b(i))                                     '
' h2 = rand8(h1 Xor b(i))                                     '
' End If                                                      '
' Next                                                        '
' lKeyVal = h1 * &HFFFF& + h2                                 '
' hash_function = lKeyVal                                     '
' End Function                                                '
' hash_method 0==add                                          '
' hash_method 1==check                                        '
' hash_method 2==delete                                       '
Public Function DataHash(ByVal skey As String, sval As String, hash_method As Integer) As String
    Dim b() As Byte
    Dim i As Long
    Dim lKeyVal As Long
    Dim h1 As Long, h2 As Long
    ' using hashing algorithm "Variable String Exclusive-Or '
    ' method (tablesize up to 65,536)".  See                '
    ' Dictionaries ->Hash Tables section at                 '
    ' http://members.xoom.com/thomasn/s_man.htm             '
    b = skey
    h1 = b(0): h2 = h1 + 1
    For i = 0 To UBound(b)
        If b(i) >= 48 And b(i) <= 57 Then
            lKeyVal = lKeyVal * 10 + b(i) - 48
        End If
    Next
    i = lKeyVal Mod m_lHashTableSize
    Do
        If lKeyVal = hashindex(i) Then
            ' If skey <> hasharray(i) Then Stop '
            If hash_method = 2 Then
                hashindex(i) = 0
                hasharray(i) = ""
            End If
            Exit Do
        End If
        If 0 = hashindex(i) Then
            If hash_method = 0 Then
                hashindex(i) = lKeyVal
                hasharray(i) = sval
            End If
            Exit Do
        End If
        i = i + 1
        If i = m_lHashTableSize Then i = 0
    Loop While 1
    If 0 = hashindex(i) Then
        DataHash = ""
        Exit Function
    End If
    DataHash = hasharray(i)
End Function
Public Sub swapint(a As Integer, b As Integer)
    Dim c As Integer
    c = a
    a = b
    b = c
End Sub
Sub reset_placement()
    Dim i As Integer
    For i = 0 To grows
        last_entry(i) = 0
    Next
End Sub

