VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form History_Form 
   Caption         =   "History"
   ClientHeight    =   8580
   ClientLeft      =   165
   ClientTop       =   1785
   ClientWidth     =   15105
   LinkTopic       =   "Form1"
   ScaleHeight     =   8580
   ScaleWidth      =   15105
   Begin VB.CommandButton Add_File 
      Caption         =   "Add Record File"
      Height          =   855
      Left            =   12000
      TabIndex        =   2
      Top             =   7200
      Width           =   1935
   End
   Begin VB.FileListBox File1 
      Height          =   6720
      Left            =   11400
      MultiSelect     =   2  'Extended
      Pattern         =   "*.txt"
      TabIndex        =   1
      Top             =   120
      Width           =   3495
   End
   Begin MSFlexGridLib.MSFlexGrid Chess 
      DragIcon        =   "history.frx":0000
      Height          =   8535
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11295
      _ExtentX        =   19923
      _ExtentY        =   15055
      _Version        =   393216
      Rows            =   500
      Cols            =   200
      FixedCols       =   7
      FocusRect       =   2
      HighLight       =   2
   End
   Begin VB.Menu MNU_SAVE 
      Caption         =   "Save"
   End
   Begin VB.Menu MNU_Recalc 
      Caption         =   "Recalc Ratings"
      Index           =   0
   End
   Begin VB.Menu MNU_Recalc 
      Caption         =   "ReRank all Players"
      Index           =   1
      Visible         =   0   'False
   End
   Begin VB.Menu MNU_Recalc 
      Caption         =   "Setup for new day"
      Index           =   2
   End
   Begin VB.Menu MNU_Edit_Player 
      Caption         =   "Edit Player"
   End
   Begin VB.Menu MNU_Settings 
      Caption         =   "Settings"
   End
   Begin VB.Menu MNU_History 
      Caption         =   "History"
      Visible         =   0   'False
   End
End
Attribute VB_Name = "History_Form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
DefLng A-Z
Option Explicit





Private Sub Add_File_Click()
Dim i As Long, j As Long
Dim sp As Long 'string position
Dim lsp As Long 'last string position
Dim ll As Long 'line length
Dim last_char As Long 'the last char read
Dim line$
Dim cell$
Dim f_idx As Long
Dim extra_col As Long
On Error Resume Next
For f_idx = 0 To File1.ListCount
    Close
    If File1.Selected(f_idx) = True Then
        Open File1.List(f_idx) For Input As #1
        For i = 0 To Chess.rows - 1
            Chess.Row = i
            If EOF(1) Then Exit For
            Line Input #1, line$ 'get rid of header
If i = 0 Then If (InStr(line$, "Games") > 10) Then extra_col = 0 Else extra_col = 1
            
            lsp = 1
            sp = 1
            ll = Len(line$)
            For j = 0 To Chess.Cols - 1
                If j > 5 Then
                    Chess.Col = j + extra_col
                Else
                    Chess.Col = j
                End If
                Do
                    last_char = Asc(Mid$(line$, sp, 1))
                    If last_char = 9 Then Exit Do
                    sp = sp + 1
                Loop While sp <= ll
                Chess.Text = LTrim$(RTrim$(Mid$(line, lsp, sp - lsp)))
                If sp >= ll Then Exit For
                    sp = sp + 1
                lsp = sp
            Next j
        Next i
        End If
Next
End Sub

Private Sub Chess_DragDrop(Source As Control, x As Single, y As Single)
On Error Resume Next
Dim i As Long, j As Long
Dim drag_to As Long
    drag_to = Chess.MouseRow
save_flag = save_flag + 1
    
    Chess.Row = drag_to
    If drag_to = drag_from Then Exit Sub
For j = Chess.FixedCols To Chess.Cols - 1
    Chess.Col = j
    If Chess.Text = "" Then
        If drag_shift Then
            Chess.Text = "D" + Str$(drag_from)
        Else
            Chess.Text = "W" + Str$(drag_from)
        End If
        Exit For
    End If
Next
End Sub


Private Sub Chess_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
If Chess.Col > 3 Then
    Select Case KeyCode
        Case vbKeyEscape
            Chess.Text = ""
        Case vbKeyCancel:
            Chess.Text = ""
        Case vbKeyBack:
            If Len(Chess.Text) Then
                Chess.Text = Left$(Chess.Text, Len(Chess.Text) - 1)
            End If
        Case vbKeyDelete:
            If Len(Chess.Text) Then
                Chess.Text = Left$(Chess.Text, Len(Chess.Text) - 1)
            End If
        Case vbKeyTab:
            Chess.Col = Chess.Col + 1
        Case vbKeyClear:
            Chess.Text = ""
        Case vbKeyReturn:
        Chess.Row = Chess.Row + 1
        Case vbKeyHome:
            Chess.Row = 1
            Chess.Col = 4
        Case vbKeyEnd:
            Chess.Row = Chess.rows
            Chess.Col = 4
        Case vbKeyInsert:
        'vbKeyLeft 37 LEFT ARROW key
        'vbKeyUp 38 UP ARROW key
        'vbKeyRight 39 RIGHT ARROW key
        'vbKeyDown
        Case 0:

        Case Else:
            If KeyCode < Asc("z") And KeyCode > Asc(" ") Then
            If Shift < 4 Then
                Chess.Text = Chess.Text + Chr$(KeyCode)
                save_flag = save_flag + 1
            End If
            End If
    End Select
End If
    
End Sub


Private Sub Chess_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim dragx As Long
On Error Resume Next
drag_shift = Shift
dragx = x / Chess.CellWidth
drag_from = Chess.MouseRow
drag_col = Chess.MouseCol
If Button <> 1 Then
    Chess.Drag 1
End If
End Sub


Private Sub Form_Load()
Dim i As Long, j As Long
Dim sp As Long 'string position
Dim lsp As Long 'last string position
Dim ll As Long 'line length
Dim last_char As Long 'the last char read
Dim line$
Dim cell$
On Error Resume Next
Close
Open "ladder.ini" For Input As #1
line$ = "32"
Line Input #1, line$
Settings!K_Factor.Text = line$
k_val = Val(Settings!K_Factor.Text)
On Error GoTo done_reading
Ladder_Form.Top = -100
Ladder_Form.Left = -100
'Chess.rows
'Chess.Cols
Close
Open "Players.xls" For Input As #1
For i = 0 To Chess.rows - 1
    Chess.Row = i
    If EOF(1) Then Exit For
    Line Input #1, line$ 'get rid of header
    lsp = 1
    sp = 1
    ll = Len(line$)
    For j = 0 To Chess.Cols - 1
        Chess.Col = j
        Do
            last_char = Asc(Mid$(line$, sp, 1))
            If last_char = 9 Then Exit Do
            sp = sp + 1
        Loop While sp <= ll
        Chess.Text = LTrim$(RTrim$(Mid$(line, lsp, sp - lsp)))
        If sp >= ll Then Exit For
            sp = sp + 1
        lsp = sp
    Next j
Next i
Chess.Col = ranking_field
Chess.ColSel = ranking_field
Chess.Sort = 1
exit_reading_sub:
Close
Chess.Row = 0

Chess.ColWidth(ranking_field) = 500
Chess.ColWidth(nrating_field) = 700
Chess.ColWidth(rating_field) = 700
Chess.ColWidth(0) = 500
Chess.ColWidth(Games_field) = 700
Chess.Col = 0: Chess.Text = "Group"
Chess.Col = 1: Chess.Text = "Last Name"
Chess.Col = 2: Chess.Text = "First Name"
Chess.Col = 3: Chess.Text = "Rating"
Chess.Col = 4: Chess.Text = "Rank"
Chess.Col = 5: Chess.Text = "N Rating"
Chess.Col = 6: Chess.Text = "Games"
              
For j = Chess.FixedCols To Chess.Cols - 1
    Chess.ColWidth(j) = 500
    Chess.Col = j
    Chess.Text = Str$(j - Chess.FixedCols + 1)
Next j

Chess.Col = Chess.FixedCols
Chess.Row = 1
Exit Sub
done_reading:
Resume exit_reading_sub
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
On Error Resume Next
Dim i As Long
If vbFormCode = UnloadMode Then End
If save_flag > 0 Then
    i = MsgBox("Save your work??", vbYesNoCancel)
    Select Case i
        Case vbYes:
            Call MNU_SAVE_Click
            End
        Case vbNo:
            End
        Case vbCancel
            Exit Sub
    End Select
End If
End
End Sub

Private Sub Form_Resize()
On Error Resume Next
Chess.Width = Ladder_Form.Width - 70
Chess.Height = Ladder_Form.Height - 660
End Sub



Private Sub MNU_Edit_Player_Click()
On Error Resume Next
Edit_Player!Player_Rank.Text = Str$(Chess.Row)
Edit_Player.Show
save_flag = save_flag + 1
End Sub


Private Sub MNU_Recalc_Click(Index As Integer)
On Error Resume Next
Dim i As Long, i2 As Long, j As Long, j2 As Long
Dim score$, player$, my_text$, my_file$
Dim opponents(grows, gcols) As Double
Dim scores(grows, gcols) As Long
Dim s(grows) As Long 'sort index
Dim rs(grows) As Long 'reverse_sort index
Dim match As Long
Dim g_count As Long
Dim prob As Double
Dim games As Double
Dim results As Double
Dim temp As Long
'first fill array
For i = 1 To grows
    s(i) = i
    rs(i) = i
Next
For i = 1 To Chess.rows - 1
    opponents(i, rating_field) = Val(Chess.TextMatrix(i, rating_field))
    For j = Chess.FixedCols To Chess.Cols - 1
        my_text$ = Chess.TextMatrix(i, j)
        If Len(my_text$) Then
            score$ = UCase$(Left$(my_text$, 1))
            player$ = Mid$(my_text$, 2)
            opponents(i, j) = Val(player$)
            Select Case score$
                Case "W":
                    scores(i, j) = 2
                Case "D":
                    scores(i, j) = 1
                Case "L":
                    scores(i, j) = 0
            End Select
        End If
    Next j
Next i
On Error GoTo 0
'next make sure that each one makes sense
For i = 1 To Chess.rows - 1
    For j = Chess.FixedCols To Chess.Cols - 1
        If opponents(i, j) <= 0 Or opponents(i, j) >= grows Then
            If opponents(i, j) Then
                MsgBox ("Player " + Str$(i) + ":" + get_name$(i) + " Has a problem")
                Exit Sub
            End If
        Else
            'We need to make sure that they only played once
             i2 = opponents(i, j)
             For j2 = j + 1 To Chess.Cols - 1
                 If i2 = opponents(i, j2) Then
                     MsgBox ("Player " + Str$(i) + ":" + get_name$(i) + " has allready played" + Str$(i2) + ":" + get_name$(i2))
                     opponents(i, j2) = 0
                     Exit For
                 End If
             Next
             'I guess we need to match this with the one on the other line
             match = 0
             For j2 = Chess.FixedCols To Chess.Cols - 1
                 If opponents(i2, j2) = i Then
                     match = 1
                     If scores(i2, j2) + scores(i, j) <> 2 Then
                         MsgBox ("Player " + Str$(i) + ":" + get_name$(i) + " vs " + Str$(i2) + ":" + get_name$(i2) + " Result missmatch")
                         Exit Sub
                     End If
                 End If
            Next j2
             'if there was no match then we get to fill it in
             If match = 0 Then
             'find an empty spot
                 For j2 = Chess.FixedCols To Chess.Cols - 1
                     If opponents(i2, j2) = 0 Then
                     'fill it in
                         opponents(i2, j2) = i
                         scores(i2, j2) = 2 - scores(i, j)
                         Exit For
                     End If
                 Next
             End If
        End If
    Next j
Next i

'last calculate ratings
For i = 1 To Chess.rows - 1
    prob = 0
    games = 0
    results = 0
    For j = Chess.FixedCols To Chess.Cols - 1
        If opponents(i, j) Then
            games = games + 1
            prob = prob + formula(opponents(i, rating_field), opponents(opponents(i, j), rating_field))
            results = results + scores(i, j) / 2#
        End If
    Next j
    opponents(i, nrating_field) = opponents(i, rating_field) + k_val * (results - prob)
Next i

'if sort commanded then sort it
If Index Then
    'just do a simple bubble sort
    For i = 1 To Chess.rows - 1
        For i2 = i + 1 To Chess.rows - 1
            If opponents(s(i), nrating_field) < opponents(s(i2), nrating_field) Then
                temp = s(i)
                s(i) = s(i2)
                s(i2) = temp
            End If
        Next
    Next
End If
For i = 1 To Chess.rows - 1
    rs(s(i)) = i
Next
'now put it back
For i = 1 To Chess.rows - 1
    Chess.TextMatrix(i, nrating_field) = Str$(Int(opponents(i, nrating_field)))
    Chess.TextMatrix(i, ranking_field) = Str$(rs(i))
    For j = Chess.FixedCols To Chess.Cols - 1
        my_text$ = ""
        If opponents(i, j) Then
            score$ = Mid$("LDW", scores(i, j) + 1, 1)
            player$ = Str$(rs(opponents(i, j)))
            my_text$ = score$ + player$
        End If
        Chess.TextMatrix(i, j) = my_text$
    Next j
Next i
If Index = 2 Then
    my_file$ = Replace(Date$ + " " + Time$, ":", "_")
    
    Call write_file("Players" + my_file$ + ".txt")
'Games_field
'add game count
For i = 1 To Chess.rows - 1
    g_count = Val(Chess.TextMatrix(i, Games_field))
    For j = Chess.FixedCols To Chess.Cols - 1
        If opponents(i, j) Then
            g_count = g_count + 1
        End If
    Next j
    Chess.TextMatrix(i, Games_field) = Str$(g_count)
Next i
    For i = 1 To Chess.rows - 1
        Chess.TextMatrix(i, rating_field) = Str$(Int(opponents(i, nrating_field)))
        Chess.TextMatrix(i, nrating_field) = ""
        Chess.TextMatrix(i, ranking_field) = Str$(rs(i))
        For j = Chess.FixedCols To Chess.Cols - 1
            my_text$ = ""
            Chess.TextMatrix(i, j) = my_text$
        Next j
    Next i
    Call MNU_SAVE_Click
End If
normal_exit_for_recalc:
    Chess.RowSel = Chess.Row
    Chess.ColSel = ranking_field
    Chess.Col = ranking_field
    Chess.Sort = flexSortNumericAscending
End Sub
Public Function get_name$(Index As Long)
get_name$ = Chess.TextMatrix(Index, first_name_field) + Chess.TextMatrix(Index, last_name_field)

End Function


Public Sub write_file(finame$)
Dim i As Long, j As Long
Dim my_text$
Open finame$ For Output As #1
For i = 0 To Chess.rows - 1
    For j = 0 To Chess.Cols - 1
        my_text$ = Chess.TextMatrix(i, j)
        Print #1, my_text$; Chr$(9);
    Next j
    Print #1, ""
Next i
Close #1
End Sub
Private Sub MNU_Settings_Click()
On Error Resume Next
Settings.Show 1
End Sub
