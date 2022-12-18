VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Ladder_Form 
   Caption         =   "Bughouse"
   ClientHeight    =   6255
   ClientLeft      =   7395
   ClientTop       =   3255
   ClientWidth     =   11550
   LinkTopic       =   "Ladder"
   ScaleHeight     =   6255
   ScaleWidth      =   11550
   Begin VB.TextBox cmd 
      Height          =   375
      Left            =   6720
      TabIndex        =   3
      Top             =   3240
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.ListBox MiniGameType 
      Height          =   1425
      Left            =   2040
      TabIndex        =   2
      Top             =   720
      Visible         =   0   'False
      Width           =   3855
   End
   Begin VB.Timer Idle_Timer 
      Interval        =   1000
      Left            =   11520
      Top             =   360
   End
   Begin MSFlexGridLib.MSFlexGrid Chess 
      DragIcon        =   "ladder.frx":0000
      Height          =   3336
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   4332
      _ExtentX        =   7646
      _ExtentY        =   5874
      _Version        =   393216
      Rows            =   50
      Cols            =   20
      FixedCols       =   13
      RowHeightMin    =   50
      FocusRect       =   2
      HighLight       =   2
      AllowUserResizing=   1
      MousePointer    =   1
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   11400
      TabIndex        =   1
      Top             =   1680
      Width           =   615
   End
   Begin VB.Menu MNU_SAVE 
      Caption         =   "Save"
   End
   Begin VB.Menu MNU_Recalc 
      Caption         =   "Recalc Ratings"
      Index           =   0
   End
   Begin VB.Menu Enable_admin_functions 
      Caption         =   "Disable Admin Functions"
   End
   Begin VB.Menu MNU_ADMIN 
      Caption         =   "Admin"
      Begin VB.Menu MNU_Edit_Player 
         Caption         =   "Edit Player"
      End
      Begin VB.Menu Blank4 
         Caption         =   ""
         Enabled         =   0   'False
      End
      Begin VB.Menu MNU_New_Day 
         Caption         =   "New day"
         Index           =   2
      End
      Begin VB.Menu blank1 
         Caption         =   ""
         Enabled         =   0   'False
      End
      Begin VB.Menu MNU_New_Day_wo_ReRank 
         Caption         =   "New Day wo ReRank"
         Index           =   3
      End
      Begin VB.Menu blank3 
         Caption         =   ""
         Enabled         =   0   'False
         Index           =   1
      End
      Begin VB.Menu MNU_Read_misc_file 
         Caption         =   "Read_misc_file"
      End
      Begin VB.Menu Blank5 
         Caption         =   ""
         Enabled         =   0   'False
      End
      Begin VB.Menu MNU_Print_Lables 
         Caption         =   "Print Lables 20/sheet"
         Index           =   0
      End
      Begin VB.Menu MNU_Print_Lables 
         Caption         =   "Print Lables 30/sheet"
         Index           =   1
      End
      Begin VB.Menu MNU_Print_Lables 
         Caption         =   "Print N Lables 20/sheet"
         Index           =   2
      End
      Begin VB.Menu MNU_Print_Lables 
         Caption         =   "Print N Lables 30/sheet"
         Index           =   3
      End
      Begin VB.Menu MNU_Print_Lables 
         Caption         =   "Print MiniGame Lables 20/Sheet"
         Index           =   4
      End
      Begin VB.Menu MNU_Print_Lables 
         Caption         =   "Print MiniGame Lables 30/Sheet"
         Index           =   5
      End
      Begin VB.Menu blank32 
         Caption         =   ""
      End
      Begin VB.Menu Print_Room_Sheet_MNU 
         Caption         =   "Print Room Sheet"
      End
      Begin VB.Menu Blank_report 
         Caption         =   ""
      End
      Begin VB.Menu write_html 
         Caption         =   "Write_Html"
      End
      Begin VB.Menu MNU_Student_Report 
         Caption         =   "Student Report"
      End
      Begin VB.Menu Blank_Settings 
         Caption         =   ""
         Enabled         =   0   'False
      End
      Begin VB.Menu MNU_Settings 
         Caption         =   "Settings"
      End
      Begin VB.Menu MNU_Blank4 
         Caption         =   ""
      End
      Begin VB.Menu MNU_Trophies 
         Caption         =   "Trophies"
      End
      Begin VB.Menu MNU_blank7 
         Caption         =   ""
      End
      Begin VB.Menu MENU_League 
         Caption         =   "League Names"
      End
      Begin VB.Menu MNU_Auto_Letter 
         Caption         =   "Auto_Letter"
         Visible         =   0   'False
      End
      Begin VB.Menu MNU_Paste_Games 
         Caption         =   "Paste Games"
      End
      Begin VB.Menu MNU_Copy_Games 
         Caption         =   "Copy Games"
      End
      Begin VB.Menu MNU_blank8 
         Caption         =   ""
      End
      Begin VB.Menu MNUSetNumGamesToZero 
         Caption         =   "SetupForNextMiniGame"
      End
   End
   Begin VB.Menu MNU_Sort 
      Caption         =   "Sort_Name"
   End
   Begin VB.Menu MNU_Zoom 
      Caption         =   "Zoom"
   End
   Begin VB.Menu MNU_Wide 
      Caption         =   "Narrow"
   End
   Begin VB.Menu MNU_Net 
      Caption         =   "Networking"
      Begin VB.Menu MNU_Net_stat 
         Caption         =   "Networking Off"
         Checked         =   -1  'True
         Index           =   0
      End
      Begin VB.Menu MNU_Net_stat 
         Caption         =   "Master"
         Index           =   1
      End
      Begin VB.Menu MNU_Net_stat 
         Caption         =   "Slave 1"
         Index           =   2
      End
      Begin VB.Menu MNU_Net_stat 
         Caption         =   "Slave 2"
         Index           =   3
      End
   End
   Begin VB.Menu MNU_MiniGameMode 
      Caption         =   "MiniGame"
      Begin VB.Menu MNU_MiniGame 
         Caption         =   "MiniGame"
      End
      Begin VB.Menu MNU_BugHouse 
         Caption         =   "BugHouse"
      End
   End
End
Attribute VB_Name = "Ladder_Form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 Option Explicit
 ' This program is free software; you can redistribute it and/or               '
 ' modify it under the terms of the GNU General Public License                 '
 ' as published by the Free Software Foundation; version 2  of the License     '
 '                                                                             '
 ' This program is distributed in the hope that it will be useful,             '
 ' but WITHOUT ANY WARRANTY; without even the implied warranty of              '
 ' MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the               '
 ' GNU General Public License for more details.                                '
 '                                                                             '
 ' You should have received a copy of the GNU General Public License           '
 ' along with this program; if not, write to the Free Software                 '
 ' Foundation, Inc., 59 Temple Place - Suite 330, Boston, MA  02111-1307, USA. '
 '                                                                             '
 ' Copyright 2008 Matt Mahowald                                                '
Function find_empty(Player As Integer) As Integer
    Dim i As Integer
    For i = last_entry(Player) To gcols - last_param_field
        If Chess.TextMatrix(Player, i + last_param_field + 1) = "" Then
            find_empty = i + last_param_field + 1
            Exit Function
        Else
            last_entry(Player) = i
        End If
    Next
    MsgBox ("Too many games by one player")
    find_empty = 0
End Function
Sub Auto_letter(this_player As Integer, al_mode As Integer)
    ' Exit Sub    '
    ' rating_sort '
    ' last_rank   '
    If Show_Ratings_value(3) = False Then Exit Sub
    Dim i As Integer
    Dim nr As Double
    Dim xr As Double
    Dim r As Integer
    Dim num_students As Integer
    Dim new_letter As String
    Dim old_letter As String
    If this_player > grows Then Exit Sub
    If this_player <= 0 Then Exit Sub
    If Right(Ladder_Form!Chess.TextMatrix(i, group_field), 1) = "x" Then Exit Sub
    ' nr = opponents(this_player, nrating_field)             '
    ' nr = Val(Chess.TextMatrix(this_player, nrating_field)) '
    new_letter = Chess.TextMatrix(this_player, last_name_field)
    If Len(new_letter) < 1 Then Exit Sub
    If nr < 0 Then Exit Sub
    r = 0
    #If 0 Then
        For i = 1 To grows
            xr = opponents(i, nrating_field)
            ' xr = Val(Chess.TextMatrix(i, nrating_field)) '
            If (opponents(i, rating_field) >= 0) Then
                num_students = num_students + 1
                If xr > nr Then
                    r = r + 1
                ' If i > this_player Then MsgBox ("error") '
                Else
                ' If i < this_player Then MsgBox ("error") '
                End If
            End If
        Next
        r = Int((r - A1size) / OtherSize) + 1
        If r > Int((num_students - A1size) / OtherSize) Then r = Int((num_students - A1size) / OtherSize)
        If r < 0 Then r = 0
    #Else
        r = 8
        If nr > 1 Then r = 5
        If nr > 100 Then r = 4
        If nr > 200 Then r = 3
        If nr > 350 Then r = 2
        If nr > 600 Then r = 1
        If nr > 900 Then r = 0
    #End If
    old_letter = Chess.TextMatrix(this_player, group_field)
    new_letter = Mid$(GROUP_CODES, group_mapping(r), group_mapping(r + 1) - group_mapping(r) - 1)
    If new_letter <> old_letter Then
        If al_mode < 10 Then
            MsgBox ("Player changed Letters" + vbCr + "Player (" + Str$(this_player) + ") " _
                    + Chess.TextMatrix(this_player, first_name_field) _
                    + " " + Chess.TextMatrix(this_player, last_name_field) & vbCr + " To " + new_letter)
        End If
        Chess.TextMatrix(this_player, group_field) = new_letter
    End If
End Sub

Private Sub Chess_Compare(ByVal Row1 As Long, ByVal Row2 As Long, Cmp As Integer)
Dim sort_column1 As Integer
Dim sort_column2 As Integer
sort_column1 = last_name_field
sort_column2 = first_name_field

If Chess.ColSel = first_name_field Then
    sort_column2 = last_name_field
    sort_column1 = first_name_field
End If
If Chess.TextMatrix(Row2, sort_column1) = "" Then
    Cmp = -1
    Exit Sub
End If
If Chess.TextMatrix(Row2, sort_column2) = "" Then
    Cmp = -1
    Exit Sub
End If
If Chess.TextMatrix(Row1, sort_column1) = "" Then
    Cmp = 1
    Exit Sub
End If
If Chess.TextMatrix(Row1, sort_column2) = "" Then
    Cmp = 1
    Exit Sub
End If
    If Chess.TextMatrix(Row1, sort_column1) > Chess.TextMatrix(Row2, sort_column1) Then
        Cmp = 1
        Exit Sub
    End If
    If Chess.TextMatrix(Row1, sort_column1) < Chess.TextMatrix(Row2, sort_column1) Then
        Cmp = -1
        Exit Sub
    End If
    If Chess.TextMatrix(Row1, sort_column2) > Chess.TextMatrix(Row2, sort_column2) Then
        Cmp = 1
        Exit Sub
    End If
    If Chess.TextMatrix(Row1, sort_column2) < Chess.TextMatrix(Row2, sort_column2) Then
        Cmp = -1
        Exit Sub
    End If

End Sub
Private Sub Chess_DragDrop(Source As Control, x As Single, y As Single)
    Exit Sub
    On Error Resume Next
    Dim j As Integer
    Dim drag_to As Integer
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
    Static zero_set As Integer
    Idle_Timer_Tag = 0
    error_count = 0
    Dim Number As Integer
    Dim this_player As Integer
    Dim gscore As Integer
    Dim opponent As Integer
    Dim test_opponent As Integer
    Dim quick_entry As Integer
    Idle_Timer.Enabled = False
    On Error Resume Next
    If Shift = 0 And ((Chess.Col < last_param_field + 1) Or (Right(Chess.Text, 1) = "_" And KeyCode <> vbKeyBack)) Then
        Chess.Col = last_param_field
        Do
            Chess.Col = Chess.Col + 1
        Loop While Chess.Text <> ""
    End If
    Select Case KeyCode
        Case vbKeyReturn:
            If MNU_Wide.Caption = "RoundRobin" Then
                Call resize_chess
            End If
            Call recalc(0)
        ' Chess.Row = Chess.Row + 1 '
        Case vbKeyHome:
            Chess.Row = 1
            Chess.Col = 4
        Case vbKeyEnd:
            Chess.Row = Chess.Rows
            Chess.Col = 4
        Case vbKeyInsert:
        ' vbKeyLeft 37 LEFT ARROW key   '
        ' vbKeyUp 38 UP ARROW key       '
        ' vbKeyRight 39 RIGHT ARROW key '
        ' vbKeyDown                     '
        Case vbKeySubtract:
            Chess.Text = Chess.Text + "L"
            save_flag = save_flag + 1
        Case vbKeyAdd:
            Chess.Text = Chess.Text + "W"
            save_flag = save_flag + 1
        Case vbKeyMultiply:
            Chess.Text = Chess.Text + "D"
            save_flag = save_flag + 1
        Case vbKeyNumpad0:
            Chess.Text = Chess.Text + "0"
            Number = 1
        Case vbKeyNumpad1:
            Chess.Text = Chess.Text + "1"
            Number = 1
        Case vbKeyNumpad2:
            Chess.Text = Chess.Text + "2"
            Number = 1
        Case vbKeyNumpad3:
            Chess.Text = Chess.Text + "3"
            Number = 1
        Case vbKeyNumpad4:
            Chess.Text = Chess.Text + "4"
            Number = 1
        Case vbKeyNumpad5:
            Chess.Text = Chess.Text + "5"
            Number = 1
        Case vbKeyNumpad6:
            Chess.Text = Chess.Text + "6"
            Number = 1
        Case vbKeyNumpad7:
            Chess.Text = Chess.Text + "7"
            Number = 1
        Case vbKeyNumpad8:
            Chess.Text = Chess.Text + "8"
            Number = 1
        Case vbKeyNumpad9:
            Chess.Text = Chess.Text + "9"
            Number = 1
        Case vbKeyF10:
            If Right$(Chess.TextMatrix(Chess.Row, group_field), 1) = "x" Then
                Chess.TextMatrix(Chess.Row, group_field) = Left$(Chess.TextMatrix(Chess.Row, group_field), Len(Chess.TextMatrix(Chess.Row, group_field)) - 1)
            Else
                Chess.TextMatrix(Chess.Row, group_field) = Chess.TextMatrix(Chess.Row, group_field) + "x"
            End If
            save_flag = save_flag + 1
        Case vbKeyF11:
            If password_set Then
                If Chess.TextMatrix(Chess.Row, attendance_field) = "X" Then
                    Chess.TextMatrix(Chess.Row, attendance_field) = ""
                Else
                    Chess.TextMatrix(Chess.Row, attendance_field) = "X"
                End If
            End If
        Case vbKeyF1:
            If password_set Then
                Chess.TextMatrix(Chess.Row, group_field) = "A1"
            End If
            save_flag = save_flag + 1
        Case vbKeyF2:
            If password_set Then Chess.TextMatrix(Chess.Row, group_field) = "A"
            save_flag = save_flag + 1
        Case vbKeyF3:
            If password_set Then Chess.TextMatrix(Chess.Row, group_field) = "B"
            save_flag = save_flag + 1
        Case vbKeyF4:
            If password_set Then Chess.TextMatrix(Chess.Row, group_field) = "C"
            save_flag = save_flag + 1
        Case vbKeyF5:
            If password_set Then Chess.TextMatrix(Chess.Row, group_field) = "D"
            save_flag = save_flag + 1
        Case vbKeyF6:
            If password_set Then Chess.TextMatrix(Chess.Row, group_field) = "E"
            save_flag = save_flag + 1
        Case vbKeyF7:
            If password_set Then Chess.TextMatrix(Chess.Row, group_field) = "F"
            save_flag = save_flag + 1
        Case vbKeyF12:
            If password_set Then
                Chess.TextMatrix(Chess.Row, group_field) = ""
                Chess.TextMatrix(Chess.Row, last_name_field) = ""
                Chess.TextMatrix(Chess.Row, first_name_field) = ""
                Chess.TextMatrix(Chess.Row, rating_field) = ""
            End If
            save_flag = save_flag + 1
        Case 0:
            zero_set = 2
        Case vbKeyTab:
            Chess.Col = Chess.Col + 1
        Case vbKeyEscape
            Chess.Text = ""
        Case vbKeyCancel:
            Chess.Text = ""
        Case vbKeyBack:
            save_flag = save_flag + 1
            If Len(Chess.Text) Then
                Chess.Text = Left$(Chess.Text, Len(Chess.Text) - 1)
            End If
        Case vbKeyDelete:
            save_flag = save_flag + 1
            If Len(Chess.Text) Then
                Chess.Text = Left$(Chess.Text, Len(Chess.Text) - 1)
            End If
        Case vbKeyClear:
            Chess.Text = ""
        Case 186:
            Chess.Text = Chess.Text + ":"
        Case vbKeyControl:
        Case 18:
        Case vbKeyShift:           ' shift    '
        Case Else:                 ' Asc("z") '
            If Shift And 2 Then
                If Shift = 6 And KeyCode = 68 Then Call delete_matching_Cell("")   ' ctrl_alt_d '
                If KeyCode = 86 Then
                    Chess.Text = Chess.Text + Clipboard.GetText
                End If
            Else
                If (KeyCode < ascz And KeyCode > ascsp) Then
                    If Shift < 4 Then
                        Chess.Text = Chess.Text + Chr$(KeyCode)
                        Number = 1
                        save_flag = save_flag + 1
                    End If
                End If
            End If
    End Select    '}-> Select Case KeyCode
    ' If number = 1 Then                                                                                                     '
    ' save_flag = save_flag + 1                                                                                              '
    ' Dim Ret As Long                                                                                                        '
    ' Ret = parse_entry(Chess.Text, players, scores, quick_entry)                                                            '
    ' If (Ret) Then Chess.Text = Chess.Text + "_"                                                                            '
    ' 'Call parse_entry(Chess.Text, this_player, gscore, opponent, quick_entry)                                              '
    ' If (opponent > Chess.Rows) Then Chess.Text = InputBox("You have entered a illegal player", "Bad Entry", Chess.Text)    '
    ' If (this_player > Chess.Rows) Then Chess.Text = InputBox("You have entered a illegal player", "Bad Entry", Chess.Text) '
    ' End If                                                                                                                 '
    Idle_Timer.Enabled = True
    If zero_set > 0 Then zero_set = zero_set - 1
End Sub
Public Sub delete_matching_Cell(this_entry As String)
    ' we need to delete the other game as well '
    On Error GoTo delete_error
    Dim i As Integer
    If this_entry = "" Then
        this_entry = Chess.Text
    End If
    Dim i2 As Integer
    Call string2long(this_entry)
    Call player2row(players)
    For i2 = 0 To 4
        If players(i2) > 0 Then
            For i = last_param_field + 1 To gcols - 1
                If this_entry = Chess.TextMatrix((players(i2)), i) Then
                    Chess.TextMatrix((players(i2)), i) = ""
                    Exit For
                End If
            Next
            If i = gcols - 1 Then
                MsgBox ("error")
            End If
        End If
    Next
    Call DataHash(this_entry, this_entry, 2)
delete_resume:
    Exit Sub
delete_error:
    error_count = error_count + 1
    GoTo delete_resume
End Sub
Private Sub Chess_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    Idle_Timer_Tag = 0
    ' Dim dragx as integer '
    On Error Resume Next
    drag_shift = Shift
    ' dragx = x / Chess.CellWidth '
    drag_from = Chess.MouseRow
    ' drag_col = Chess.MouseCol '
    If Button <> 1 Then
        Chess.Drag 1
    End If
End Sub

Private Sub cmd_Change()
On Error Resume Next
Dim i As Integer
Dim j As Integer
If cmd.Text <> "" Then
    If Left$(cmd.Text, 1) = "-" Then
        Dim this_entry As String
        this_entry = Mid$(cmd.Text, 2) + "_"
        Call delete_matching_Cell(this_entry)
        Exit Sub
    End If
    For i = 1 To Chess.Rows - 1
        For j = Chess.Cols - 1 To Chess.FixedCols Step -1
            If Chess.TextMatrix(i, j) = "" Then
                Chess.TextMatrix(i, j) = cmd.Text
                cmd.Text = ""
                Call recalc(0)
                Exit Sub
            End If
        Next
    Next
End If
End Sub

Private Sub Enable_admin_functions_Click()
    Dim i As Integer
    ' On Error GoTo enable_admin_error_resume '
    MNU_ADMIN.Visible = False
    password_set = 0
    If UCase$(Left$(Enable_admin_functions.Caption, 4)) = "DISA" Then
        Enable_admin_functions.Caption = "Enable Admin Functions"
        Call recalc(0)
        Call resize_chess
        Exit Sub
    End If
    Enable_admin_functions.Caption = "Enable Admin Functions"
    frmLogin.Show 1
    If frmLogin.LoginSucceeded = True Then password_set = 1 Else Exit Sub
    MNU_ADMIN.Visible = True
    Enable_admin_functions.Caption = "Disable Admin Functions"
    Call recalc(0)
    Call resize_chess
    Exit Sub
enable_admin_error_resume:
    error_count = error_count + 1
    Exit Sub
End Sub
Public Sub load_file(filename$)
    Dim i As Integer, j As Integer
    Dim sp As Integer              ' string position      '
    Dim lsp As Integer             ' last string position '
    Dim ll As Integer              ' line length          '
    Dim last_char As Integer       ' the last char read   '
    Dim line$
    ' Dim cell$ '
    Dim computer_name$
    Dim li(gcols + 1) As String
    Dim my_rank As Integer
    Dim ladder_name$
    ' Dim debug_ptr as integer '
    Dim version_no As Double
    Dim state$
    Dim games_read  As Integer
    Dim last_rank As Integer
    group_mapping(0) = 1           ' a1 '
    group_mapping(1) = 4           ' a  '
    group_mapping(2) = 6
    group_mapping(3) = 8
    group_mapping(4) = 10
    group_mapping(5) = 12
    group_mapping(6) = 14
    group_mapping(7) = 16
    group_mapping(8) = 19
    Game_Result(0) = "Lost"
    Game_Result(1) = "Draw"
    Game_Result(2) = "Won"
    ladder_name$ = get_ladder_name()
    Me.Caption = ladder_name$ + " Rated Ladder"
    On Error GoTo form_load_open_error
    Close
    Open "ladder.ini" For Input As #1
    line$ = "32"
    Line Input #1, line$
    Settings!K_Factor.Text = line$
    line$ = "100"
    Line Input #1, line$
    On Error GoTo form_load_error
    Settings!set_grows = line$
    grows = Val(line$)
    k_val_base = Val(Settings!K_Factor.Text)
    On Error GoTo form_load_error
    line$ = "0"
    Line Input #1, line$
    Settings!Show_Ratings(0).Value = Val(line$)
    Show_Ratings_value(0) = Settings!Show_Ratings(0).Value
    On Error Resume Next
    mpassword$ = "Matt"
    Line Input #1, mpassword$
    opassword$ = mpassword$
    line$ = "0"
    Line Input #1, line$
    Settings!Show_Ratings(1).Value = Val(line$)
    Show_Ratings_value(1) = Settings!Show_Ratings(1).Value
    Line Input #1, line$
    Settings!Show_Ratings(2).Value = Val(line$)
    Show_Ratings_value(2) = Settings!Show_Ratings(2).Value
    line$ = ""
    Line Input #1, line$:    Settings!Place_Trophies.Text = line$
    line$ = ""
    Line Input #1, line$: Settings!coaches(0).Text = line$
    Line Input #1, line$: Settings!coaches(1).Text = line$
    Line Input #1, line$: Settings!coaches(2).Text = line$
    line$ = ""
    Line Input #1, line$: Settings!print_offset.Text = line$
    line$ = ""
    Line Input #1, line$: Settings!pzoom.Text = line$
    line$ = "0"
    Line Input #1, line$
    Show_Ratings_value(3) = False
    Settings!Show_Ratings(3).Value = Val(line$)
    Show_Ratings_value(3) = Settings!Show_Ratings(3).Value
    line$ = ""
    Line Input #1, line$: Settings!MasterIP.Text = line$
    On Error GoTo check_drive_error
    #If 1 Then
        If ladder_name$ = "ladder_Program" Then mpassword$ = ""
        ' check for spefific computers '
        If Left$(CurDir$, 12) = "F:\@\a_chess" Then mpassword$ = ""
            Drive1.Drive = "c:"
            computer_name$ = Drive1.Drive
            i = InStr(computer_name$, "[")
            If i > 0 Then computer_name$ = Mid$(computer_name, i + 1)
            i = InStr(computer_name$, "]")
            If i > 0 Then computer_name$ = Left$(computer_name, i - 1)
            If UCase$(computer_name$) = "TUTOR" Then mpassword$ = ""
            If Left$(UCase$(computer_name$), 5) = "MATTS" Then mpassword$ = ""
    #End If
    On Error Resume Next
    Ladder_Form.Top = -100
    Ladder_Form.Left = -100
    On Error GoTo done_reading
    Chess.Clear
    Chess.Sort = 0
    Chess.Rows = grows + 1
    Chess.Cols = gcols
    ReDim ispresent(gcols)
    ReDim row2RR(gcols)
    row2RR_count = 0
    ' ReDim chess_matrix(0 To grows + 1, 0 To gcols) '
    last_rank = grows
    Close
    Chess.Redraw = False
    state$ = "opening file"
    Open filename$ For Input As #1
    state$ = "opened file"
    version_no = 1
    For i = 0 To Chess.Rows - 1
        If EOF(1) Then Exit For
        state$ = "Ready_to_Read"
read_line_again:
        Line Input #1, line$       ' get rid of header '
        If Len(line$) < 2 Then GoTo read_line_again
        ll = InStr(line$, "Version")
        If ll > 0 Then
            version_no = Val(Mid$(line$, ll + 8, 4))
        End If
        lsp = 1
        sp = 1
        For j = 0 To Chess.Cols - 1
            li(j) = ""
        Next
        ll = Len(line$)
        For j = 0 To Chess.Cols - 1
            Do
                last_char = Asc(Mid$(line$, sp, 1))
                If last_char = 9 Then
                    Exit Do
                End If
                If last_char = 44 Then
                    Exit Do
                End If
                sp = sp + 1
            Loop While sp <= ll
            li(j) = LTrim$(RTrim$(Mid$(line, lsp, sp - lsp)))
            If version_no < 1.05 And j = attendance_field - 1 Then
                j = j + 2
            End If
            If version_no < 1.2 And j = room_field - 1 Then
                j = j + 1
            End If
            If sp >= ll Then Exit For
            sp = sp + 1
            lsp = sp
        Next j
        my_rank = Val(li(ranking_field))
        state$ = "Doing Ranks"
        If li(last_name_field) <> "" Then
            If my_rank >= grows Then
                my_rank = last_rank
                last_rank = last_rank - 1
            End If
            Chess.Row = my_rank
            For j = 0 To Chess.Cols - 1
                Chess.Col = j
                If j > last_param_field Then
                    Chess.CellAlignment = flexAlignCenterCenter
                End If
                Chess.Text = li(j)
                If (li(j) <> "") And (my_rank > 0) And (j > last_param_field) Then
                    games_read = games_read + 1
                    If Left$(Chess.TextMatrix(my_rank, attendance_field), 1) = "X" Then Chess.TextMatrix(my_rank, attendance_field) = "X"
                End If
            Next j
        End If
    Next i    '}-> For i = 0 To Chess.Rows - 1
exit_reading_sub:
    ' repair rating field '
    For i = 1 To Chess.Rows - 1
        Chess.TextMatrix(i, ranking_field) = i
        If Chess.TextMatrix(i, last_name_field) = "" Then
            Chess.TextMatrix(i, rating_field) = "-1"
        End If
    Next
    Call Set_sort_rank
    Chess.Redraw = True
    Close
    Chess.Row = 0
    Chess.ColWidth(ranking_field) = 500 * AppZoom
    Chess.ColWidth(group_field) = 300 * AppZoom
    Chess.ColWidth(Games_field) = 400 * AppZoom
    Chess.ColWidth(ranking_field) = 300 * AppZoom
    Chess.ColWidth(grade_field) = 400 * AppZoom
    Chess.ColWidth(attendance_field) = 200 * AppZoom
    Chess.Col = group_field: Chess.Text = "Group"
    Chess.Col = last_name_field: Chess.Text = "Last Name"
    Chess.Col = first_name_field: Chess.Text = "First Name"
    Chess.Col = rating_field: Chess.Text = "Rating"
    Chess.Col = ranking_field: Chess.Text = "Rnk"
    Chess.Col = nrating_field: Chess.Text = "N Rate"
    Chess.Col = Games_field: Chess.Text = "Gms"
    Chess.Col = grade_field: Chess.Text = "Gr"
    Chess.Col = info_field: Chess.Text = "Info"
    Chess.Col = phone_field: Chess.Text = "Phone"
    Chess.Col = school_field: Chess.Text = "School"
    Chess.Col = room_field: Chess.Text = "Room"
    Chess.Col = attendance_field: Chess.Text = "X"
    For j = Chess.FixedCols To Chess.Cols - 1
        Chess.Col = j
        Chess.Text = Str$(j - Chess.FixedCols + 1)
    Next j
    Chess.Col = Chess.FixedCols
    Chess.Row = 1
    ' If games_read > 0 Then                   '
    ' recalc (3) 'auto new day without re-rank '
    ' End If                                   '
    ' Stop                                     '
    If Files.Sort_By_Name.ListIndex = Sort_Name Then
        Call Ladder_Form.Set_Sort_Name
    End If
    Exit Sub
done_reading:
    error_count = error_count + 1
    If error_count > error_max Then Resume Next
    Select Case MsgBox(Error$(merr) & line$ + Chr$(38) + state$, 2)
        Case vbAbort:
            Resume exit_reading_sub
        Case vbRetry:
            Resume
        Case vbIgnore:
            Resume Next
    End Select
    Resume exit_reading_sub
    Resume
form_load_open_error:
    MsgBox ("Could not open input file!!" + vbLf + "Exiting")
    End
form_load_error:
check_drive_error:
    error_count = error_count + 1
    If error_count > error_max Then Resume Next
    merr = Err
    Select Case MsgBox(Error$(merr), 2)
        Case vbAbort:
            Resume Next
        Case vbRetry:
            Resume
        Case vbIgnore:
            Resume Next
    End Select
    ' Resume '
    Resume Next
End Sub
Public Sub reset_r2p()
    Dim i As Integer
    For i = 1 To Chess.Rows - 1
        r2p(i) = Chess.TextMatrix(i, ranking_field)
    Next i
    For i = 1 To Chess.Rows - 1
        p2r(r2p(i)) = i
    Next i
End Sub

Private Sub Form_Load()
On Error Resume Next
    Call hash_Initialize
    If (Command <> "") Then
        ChDir (Command)
        ChDrive (Command)
    End If
    ' Dim i As Integer        '
    ' Dim j As Integer        '
    ' Dim s As String         '
    ' For i = 1 To 100        '
    ' For j = 1 To 100        '
    ' s = Str(i) + Str(j)     '
    ' s = Replace(s, " ", "") '
    ' Call hash(s, s, 0)      '
    ' Next j                  '
    ' Next                    '
    MiniGameType.AddItem "B-G_Game"
    MiniGameType.AddItem "Bishop_Game"
    MiniGameType.AddItem "Pillar_Game"
    MiniGameType.AddItem "Kings_Cross"
    MiniGameType.AddItem "Pawn_Game"
    MiniGameType.AddItem "Queen_Game"
    Call MNU_Sort_Click
    #If 0 Then
        Dim test As Long
        Dim tests As String
        Dim tests2 As String
        tests = "23:29LW31:28"
        ' tests = "23:99:34:31:28" '
        test = string2long(tests)
        tests2 = long2string(test)
        If "23:29LW28:31" <> tests2 Then
            Stop
        End If
    ' Stop                                     '
    ' tests = hash("hash", "hash_string", 0)   '
    ' tests = hash("hasher", "hash_string", 0) '
    ' tests = hash("hash", "", 1)              '
    ' tests = hash("hasher", "delete", 2)      '
    ' tests = hash("hash", "", 1)              '
    #End If
    AppZoom = 1
    Dim file_loaded As Boolean
    If (GetAttr("ladder.ini") And 1) <> 1 Then
    Else
        Shell "attrib  -R -S  -H *"
    End If
    If Dir("Saved_Players.xls") <> "" Then
        If vbYes = MsgBox("Use recovered file (or discard)", vbYesNo) Then
            file_loaded = True
            load_file "Saved_Players.xls"
        End If
    End If
    If file_loaded = False Then
        load_file "Players.xls"
    End If
    If merr Then Exit Sub
    Call MNU_MiniGame_Click
    Call Enable_admin_functions_Click
    'Ladder_Form.WindowState = 2
    If Screen.Width > 20000 Then Call MNU_Zoom_Click
    Set reports(0) = New Report
    reports(0).Tag = 0
    reports(0).Show
    Set reports(1) = New Report
    reports(1).Tag = 1
    reports(1).Show
    'Report.Show 0
End Sub
Private Sub MENU_League_Click()
    Dim sres As Integer
    Dim i As Integer
    sres = MsgBox("This will room number column with leauge names, are you sure", vbOKCancel)
    If sres <> vbOK Then Exit Sub
    Dim s As String
    For i = 1 To Chess.Rows - 1
        s = Chess.TextMatrix(i, last_name_field) + ", " + Chess.TextMatrix(i, first_name_field)
        If Len(s) > 5 Then Chess.TextMatrix(i, room_field) = s
    Next i
End Sub
Private Sub MiniGameType_Click()
    On Error Resume Next
    Call Ladder_Form.Set_Sort_Rating
    write_file (MiniGameType + ".xls")
    Call Ladder_Form.Set_sort_rank
    Call recalc(3)
    Dim i As Integer
    For i = 0 To Chess.Rows - 1
        Chess.TextMatrix(i, Games_field) = ""
    Next
    MiniGameType.Visible = False
End Sub
Private Sub MNU_Auto_Letter_Click()
' Basic flow for spectaculars:                                                                                                                                                               '
' Known players have ratings & groups, with games set to 200 (special for spectaculars).                                                                                                     '
' Results are entered one at a time (with quick entry).                                                                                                                                      '
' Program indicates whether a letter has changed for ONLY the two contestants.  (A single entry can change as many as half the players).  Players are given new stickers to apply to labels. '
' Players with games set to 200 will automatically drop to 0 when 5 games have been reached.                                                                                                 '
' Problems:                                                                                                                                                                                  '
' Winner of fast, then slow game is going to be fairly uncomfortable on their third game (cause they will be an A1 and not deserve it).                                                      '
' Many players will need 1 sticker per round.                                                                                                                                                '
' Advantages:                                                                                                                                                                                '
' Never need to reprint labels.  We can hand write replacements.                                                                                                                             '
' Winner will be more of a surprise.                                                                                                                                                         '
' I think the following rules:                                                                                                                                                               '
' A.  8 A1                                                                                                                                                                                   '
' B.  10 for each letter                                                                                                                                                                     '
' C.  Last letter gets extras.                                                                                                                                                               '
' ?D.  If a player misses a week, automatically don’t count them the next week.                                                                                                              '
' A:                                                                                                                                                                                         '
' A1 is a special category created for the players too good to play the general population.  By limiting it to the top 5, we make it more special.                                           '
' If we limit it to the top 8, its still special, but it has enough players to make a group.                                                                                                 '
' B:  We know from SH that less than 10 does not work (we had problems with small groups).                                                                                                   '
' I think that larger than 15 also causes problems.                                                                                                                                          '
' C:  I think that the lowest group should be the largest, because they play the fewest games.                                                                                               '
' D:  Does not seem to work with spectaculars, (the primary purpose for this code).                                                                                                          '
End Sub

Private Sub MNU_BugHouse_Click()
MNU_MiniGameMode.Caption = "BugHouse"
setMiniGame (False)
End Sub
Public Sub setMiniGame(mini As Boolean)
On Error Resume Next
Dim isTrue As Boolean
Dim isFalse As Boolean
isTrue = mini
isFalse = Not mini
MNU_BugHouse.Checked = isFalse
MNU_MiniGame.Checked = isTrue
MNUSetNumGamesToZero.Visible = isTrue
MNU_New_Day(2).Visible = isFalse
MNU_New_Day_wo_ReRank(3) = isFalse
MNU_Read_misc_file = isFalse
End Sub
Private Sub MNU_MiniGame_Click()
MNU_MiniGameMode.Caption = "MiniGame"
setMiniGame (True)
End Sub

Private Sub MNU_Net_stat_Click(Index As Integer)
On Error Resume Next
Dim i As Integer
For i = 0 To 3
    If i = Index Then
        MNU_Net_stat(i).Checked = True
    Else
        MNU_Net_stat(i).Checked = False
    End If
Next
End Sub

Private Sub MNU_Paste_Games_Click()
    Dim p As Integer
    Dim i As Integer
    Dim s As String
    Dim c As Integer
    Dim r As Integer
    Dim result As String
    Dim my_text$
    p = 2
    s = Clipboard.GetText
    Do While Len(s) > 0
        i = InStr(s, Chr$(9))
        If i < 1 Then Exit Do
        If i = 1 Then
            s = Mid$(s, i + 1)
            i = InStr(s, Chr$(9))
        End If
        result = Left$(s, i - 1)
        s = Mid$(s, i + 1)
        If i > 1 Then
            ' find blank square '
            c = Ladder_Form!ChessCols - (p Mod (Ladder_Form!Chess.Rows - 2))
            r = p / (Ladder_Form!Chess.Rows - 2)
            Chess.TextMatrix(r, c) = result
        End If
    Loop
    ' On Error Resume Next '
    Clipboard.Clear
End Sub
Private Sub MNU_Copy_Games_Click()
    Dim i As Integer
    Dim s As String
    Dim my_text$
    On Error Resume Next
    Clipboard.Clear
    For i = 1 To hashsize
        my_text$ = hasharray(i)
        If Len(my_text$) > 0 Then s = s + Chr$(9) + my_text$
    Next
    Clipboard.SetText s
End Sub
Public Sub MNU_Print_Lables_Click(Index As Integer)
    Dim i As Integer, j As Integer, k As Integer
    Dim my_text$, l_table$, r_table$, end_of_line$, start_of_line$
    Dim x As Integer
    Dim y As Integer
    Dim label_number As Integer
    Dim print_style As Integer
    Dim last_name$
    Dim file1 As String
    Dim ladder_name$
    Dim labels_per_page As Integer
    Dim xscale As Double
    Dim yscale As Double
    Dim font_scale As Double
    Dim num_ys As Integer
    Dim num_xs As Integer
    Dim n As Integer
    Dim poffset As Double
    Dim pzoom As Double
    poffset = Val(Settings!print_offset) * 1600
    pzoom = Val(Settings!pzoom)
    If pzoom < 0.5 Or pzoom > 2 Then
        pzoom = 1
    End If
    n = 1
    xscale = 1
    num_ys = 10
    num_xs = 2
    If Index = 2 Or Index = 3 Then
        n = Val(InputBox("Number of lables"))
    End If
    Select Case Index Mod 2
        Case 0:
        Case 1:
            num_xs = 3
            xscale = 2# / 3#
    End Select
    labels_per_page = num_ys * num_xs
    font_scale = Sqr(xscale)
    font_scale = font_scale * Sqr(pzoom)
    yscale = pzoom
    ladder_name$ = get_ladder_name()
    Call Set_Sort_First_Name
    'Call Set_Sort_Name
    For k = 1 To n
        For i = 1 To Ladder_Form!Chess.Rows - 1
            If Ladder_Form!Chess.TextMatrix(i, last_name_field) <> "" Then
                If Right(Ladder_Form!Chess.TextMatrix(i, group_field), 1) <> "x" Then
                    label_number = label_number + 1
                    If label_number = labels_per_page + 1 Then
                        Printer.NewPage
                        label_number = 1
                    End If
                    x = ((label_number - 1) Mod num_xs) * 6057 + 250   '
                    y = Int((label_number - 1) / num_xs) * 1446 + 500 + poffset
                    x = x * xscale
                    Printer.PSet (x + 150 * xscale, y + 200 * yscale)
                    Printer.fontsize = 6 * font_scale
                    Printer.Print ladder_name$
                    If Index > 4 Then
                        j = 1          ' Last Name '
                        my_text$ = Ladder_Form!Chess.TextMatrix(i, j)
                        Printer.PSet (x + 1500 * xscale, y + 1000 * yscale)
                        Printer.fontsize = 12 * font_scale
                        Printer.Print my_text$
                        j = 2          ' First Name '
                        my_text$ = Ladder_Form!Chess.TextMatrix(i, j)
                        Printer.PSet (x + 1000 * xscale, y + 400 * yscale)
                        Printer.fontsize = 30 * font_scale
                        Printer.Print my_text$
                        j = 4          ' Rank '
                        my_text$ = Ladder_Form!Chess.TextMatrix(i, j)
                        Printer.PSet (x + 50 * xscale, y + 400 * yscale)
                        Printer.fontsize = 30 * font_scale
                        Printer.Print my_text$
                        j = 4          ' Rank '
                        my_text$ = Ladder_Form!Chess.TextMatrix(i, j)
                        Printer.PSet (x + 4300 * xscale, y + 400 * yscale)
                        Printer.fontsize = 30 * font_scale
                        Printer.Print my_text$
                        j = grade_field   ' Grade '
                        my_text$ = Ladder_Form!Chess.TextMatrix(i, j)
                        Printer.PSet (x + 1300 * xscale, y + 1000 * yscale)
                        Printer.fontsize = 14 * font_scale
                        Printer.FontBold = True
                        Printer.Print my_text$
                    Else
                        j = 0          ' group '
                        my_text$ = Ladder_Form!Chess.TextMatrix(i, j)
                        Printer.PSet (x + 4500 * xscale, y + 500 * yscale)
                        Printer.fontsize = 24 * font_scale
                        Printer.Print my_text$
                        j = 1          ' Last Name '
                        my_text$ = Ladder_Form!Chess.TextMatrix(i, j)
                        Printer.PSet (x + 1500 * xscale, y + 1000 * yscale)
                        Printer.fontsize = 12 * font_scale
                        Printer.Print my_text$
                        j = 2          ' First Name '
                        my_text$ = Ladder_Form!Chess.TextMatrix(i, j)
                        Printer.PSet (x + 50 * xscale, y + 400 * yscale)
                        Printer.fontsize = 30 * font_scale
                        Printer.Print my_text$
                        j = 3          ' Rating '
                        my_text$ = Ladder_Form!Chess.TextMatrix(i, j)
                        Printer.PSet (x + 2500 * xscale, y + 200 * yscale)
                        Printer.fontsize = 12 * font_scale
                        If Show_Ratings_value(0) = True Then Printer.Print my_text$
                        j = 4          ' Rank '
                        my_text$ = Ladder_Form!Chess.TextMatrix(i, j)
                        Printer.PSet (x + 4500 * xscale, y + 50 * yscale)
                        Printer.fontsize = 17 * font_scale
                        Printer.FontBold = True
                        Printer.Print my_text$
                        j = grade_field   ' Grade '
                        my_text$ = Ladder_Form!Chess.TextMatrix(i, j)
                        Printer.PSet (x + 150 * xscale, y + 1000 * yscale)
                        Printer.fontsize = 13 * font_scale
                        Printer.FontBold = False
                        Printer.Print my_text$
                    End If
                    Printer.FontBold = False
                    j = school_field   ' school_field '
                    my_text$ = Ladder_Form!Chess.TextMatrix(i, j)
                    Printer.PSet (x + 3500 * xscale, y + 1000 * yscale)
                    Printer.fontsize = 10 * font_scale
                    If Show_Ratings_value(1) = True Then Printer.Print my_text$
                    j = room_field   ' room_field '
                    my_text$ = Ladder_Form!Chess.TextMatrix(i, j)
                    Printer.PSet (x + 3500 * xscale, y + 1000 * yscale)
                    Printer.fontsize = 10 * font_scale
                    If Show_Ratings_value(1) = False Then Printer.Print my_text$
                End If    '}-> If Right(Ladder_Form!Chess.TextMatrix(i, group_field), 1) <> "x" Then
            Else
                If Index = 4 Then
                    label_number = label_number + 1
                    If label_number = labels_per_page + 1 Then
                        Printer.NewPage
                        label_number = 1
                    End If
                    x = ((label_number - 1) Mod num_xs) * 6057 + 250   '
                    y = Int((label_number - 1) / num_xs) * 1446 + 500 + poffset
                    x = x * xscale
                    Printer.PSet (x + 150 * xscale, y + 200 * yscale)
                    Printer.fontsize = 6 * font_scale
                    Printer.Print ladder_name$
                    j = 4          ' Rank '
                    my_text$ = Ladder_Form!Chess.TextMatrix(i, j)
                    Printer.PSet (x + 50 * xscale, y + 400 * yscale)
                    Printer.fontsize = 30 * font_scale
                    Printer.Print my_text$
                    j = 4          ' Rank '
                    my_text$ = Ladder_Form!Chess.TextMatrix(i, j)
                    Printer.PSet (x + 4300 * xscale, y + 400 * yscale)
                    Printer.fontsize = 30 * font_scale
                    Printer.Print my_text$
                End If
            End If    '}-> If Ladder_Form!Chess.TextMatrix(i, last_name_field) <> "" Then
        Next i    '}-> For i = 1 To Ladder_Form!Chess.Rows - 1
    Next k    '}-> For k = 1 To n
    Printer.EndDoc
    Call Set_sort_rank
end_of_table:
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    On Error Resume Next
    Dim i As Integer
    If vbFormCode = UnloadMode Then End
    If save_flag > 0 Then
        i = MsgBox("Save your work??", vbYesNoCancel)
        Select Case i
            Case vbYes:
                Call do_save
                End
            Case vbNo:
                End
            Case vbCancel
                Cancel = True
                Exit Sub
        End Select
    End If
    End
End Sub
Private Sub Form_Resize()
    If quick Then Exit Sub
    quick = 1
    On Error Resume Next
    Chess.Width = Ladder_Form.Width - 70
    Chess.Height = Ladder_Form.Height - 660
    quick = 0
End Sub
Private Sub Idle_Timer_Timer()
    If save_flag = 0 Then
        Idle_Timer_Tag = 0
    End If
    If Idle_Timer_Tag > 30 Then
        Idle_Timer_Tag = 0
        Idle_Timer.Enabled = False
        If in_recalc = 0 Then
            Call write_file("Saved_Players.xls")
            Call recalc(0)
        End If
    End If
    Idle_Timer_Tag = Idle_Timer_Tag + 1
End Sub
Private Sub Menu_ReRank_all_Players_Click(Index As Integer)
    MNU_Recalc_Click (Index)
End Sub
Private Sub MNU_Edit_Player_Click()
    On Error Resume Next
    Dim i As Integer
    Dim change_player  As Boolean
    Edit_Player!Player_Rank.Text = Str$(Chess.Row)
    Edit_Player.Hide
    If password_set = 0 Then
        change_player = False
    Else
        change_player = True
    End If
    Edit_Player!Save_Next(0).Visible = change_player
    Edit_Player!Save_Next(1).Visible = change_player
    Edit_Player!Clear_All.Visible = change_player
    For i = Edit_Player!Letter.LBound To Edit_Player!Letter.UBound
        Edit_Player!Letter(i).Visible = change_player
    Next
    Edit_Player.Show 1
    save_flag = save_flag + 1
End Sub
Private Sub MNU_New_Day_Click(Index As Integer)
    MNU_Recalc_Click (Index)
End Sub
Private Sub MNU_New_Day_wo_ReRank_Click(Index As Integer)
    MNU_Recalc_Click (Index)
End Sub
Private Sub MNU_Read_misc_file_Click()
    Files.Show
End Sub
Public Sub resize_chess()
    needs_resize = False
    On Error Resume Next
    Dim i As Integer
    Dim j As Integer
    Chess.Font.Size = 8.25 * AppZoom
    If Files.Sort_By_Name.ListIndex Then
        Call Ladder_Form.Set_Sort_Name
    Else
        Call Ladder_Form.Set_sort_rank
    End If
    j = last_param_field
    If MNU_Wide.Caption = "RoundRobin" Then
        For i = 1 To Chess.Rows - 1
            If Chess.TextMatrix(i, attendance_field) = "X" Then
                Chess.RowHeight(i) = 230 * Sqr(AppZoom)
                j = j + 1
                Chess.TextMatrix(0, j) = i
            Else
                Chess.RowHeight(i) = 1
            End If
        Next
        If j < last_param_field + 1 Then
            For i = 1 To Chess.Rows - 1
                Chess.RowHeight(i) = 230 * Sqr(AppZoom)
                j = j + 1
                Chess.TextMatrix(0, j) = i
            Next
        End If
    Else
        For i = 1 To Chess.Rows - 1
            Chess.RowHeight(i) = 230 * Sqr(AppZoom)
            j = j + 1
            Chess.TextMatrix(0, j) = i
        Next
    End If
    Chess.ColWidth(group_field) = 300 * AppZoom
    Chess.ColWidth(ranking_field) = 300 * AppZoom
    Chess.ColWidth(grade_field) = 200 * AppZoom
    Chess.ColWidth(attendance_field) = 200 * AppZoom
    If entry_size < 300 Then
        entry_size = 300
    End If
    ' column width for entries '
    For i = 1 To gcols - last_param_field - 1
        Chess.ColWidth(last_param_field + i) = entry_size * AppZoom * 2
    Next
    If password_set Then
        Chess.ColWidth(info_field) = 500 * AppZoom
        Chess.ColWidth(phone_field) = 1100 * AppZoom
        Chess.ColWidth(room_field) = 500 * AppZoom
    Else
        Chess.ColWidth(info_field) = 0
        Chess.ColWidth(phone_field) = 0
        Chess.ColWidth(room_field) = 0
    End If
    If Show_Ratings_value(1) Or password_set Then
        Chess.ColWidth(school_field) = 600 * AppZoom
    Else
        Chess.ColWidth(school_field) = 0
    End If
    If Show_Ratings_value(0) Or password_set Then
        Chess.ColWidth(nrating_field) = rating_field_size * AppZoom
        Chess.ColWidth(rating_field) = rating_field_size * AppZoom
    Else
        Chess.ColWidth(nrating_field) = 0
        Chess.ColWidth(rating_field) = 0
    End If
End Sub
Private Sub MNU_Recalc_Click(Index As Integer)
    merr = 0
    in_recalc = 0
    Call recalc(Index)
    in_recalc = 0
End Sub
Public Sub recalc(Index As Integer)
    If merr Then Exit Sub
    On Error Resume Next
    in_recalc = 1
    Dim i As Integer
    Dim j As Integer
    Dim j2 As Integer
    Dim myplayer As Integer
    Dim myside As Integer
    Dim my_text$
    Dim k_val As Double
    Dim perf As Double
    Dim o_id As Integer
    Dim issport As Integer
    ReDim isvalid(grows) As Integer
    ReDim num_games(grows) As Integer
    ReDim nrating(grows) As Double
    ReDim ispresent(grows) As Integer
    Dim sides(2) As Double
    Dim perf_rating(2) As Double
    Dim perfs(2) As Double
    Dim ret As Long
    Dim pl As Integer
    Dim opl As Integer
    Dim sres As String
    ' first fill array '
    k_val = Val(Settings!K_Factor.Text)
    Call reset_r2p
    row2RR_count = 0
    For i = 1 To Chess.Rows - 1
        ' If i = 37 Then Stop '
        If Len(Ladder_Form!Chess.TextMatrix(i, last_name_field)) > 1 Then
            isvalid(i) = 1
        End If
        ispresent(i) = 0
        If Left$(Chess.TextMatrix(i, attendance_field), 1) = "X" Then
            ispresent(i) = -1
            row2RR_count = row2RR_count + 1
            row2RR(i) = row2RR_count
        Else
            ispresent(i) = Val(Chess.TextMatrix(i, attendance_field)) + 1
        End If
        num_games(i) = Val(Chess.TextMatrix(i, Games_field))
        perf = Val(Chess.TextMatrix(i, rating_field))
        If num_games(i) = 0 Then
            perf = Val(Chess.TextMatrix(i, nrating_field))
            If perf > 1200 Then perf = 1200
            nrating(i) = Abs(perf)
        Else
            nrating(i) = Abs(perf)
        End If
    ' If nrating(i) > 1800 Then nrating(i) = 1800 '
    Next
    For i = 1 To Chess.Rows - 1
        For j = Chess.FixedCols To Chess.Cols - 1
parse_try_again:
            my_text$ = UCase(Chess.TextMatrix(i, j))
            ' If Right(my_text$, 1) <> "_" Then '
            If my_text$ <> "" Then
                If my_text$ = "S" Then
                    issport = 1
                    issport = Val(Mid$(Chess.TextMatrix(i, attendance_field), 2)) + 1
                    Chess.TextMatrix(i, j) = ""
                    Chess.TextMatrix(i, attendance_field) = "X" + Str$(issport)
                Else
                    ret = parse_entry(my_text$, players, scores, quick_entry)
                    Call player2row(players)
                    If (ret > 0) Then
                        Chess.TextMatrix(i, j) = ""   ' delete the entry '
                        my_text$ = long2string(ret)
                        Call DataHash(my_text$, my_text$, 0)   ' add it to the hash '
                    End If
                End If
            ' Stop '
            End If
        Next
    Next
    For i = 1 To hashsize
        my_text$ = hasharray(i)
        If Len(my_text$) Then
            ret = parse_entry(my_text$, players, scores, quick_entry)
            Call player2row(players)
            If players(1) > 0 Then
                sides(0) = (nrating(players(0)) + nrating(players(1))) / 2
                sides(1) = (nrating(players(4)) + nrating(players(3))) / 2
            Else
                sides(0) = nrating(players(0))
                sides(1) = nrating(players(3))
            End If
            perf = formula(sides(0), sides(1))
            perfs(0) = 0
            perfs(1) = 0
            For myplayer = 0 To 1
                If scores(myplayer) > 0 Then
                    Select Case scores(myplayer)
                        Case 3:
                            perfs(0) = perfs(0) + 0.5
                            perfs(1) = perfs(1) - 0.5
                        Case 2:
                        Case 1:
                            perfs(0) = perfs(0) - 0.5
                            perfs(1) = perfs(1) + 0.5
                    End Select
                End If
            Next
            If scores(1) > 0 Then
                sides(0) = sides(0) * 2
                sides(1) = sides(1) * 2
            End If
            sides(0) = sides(0) + 800 * perfs(1)
            sides(1) = sides(1) + 800 * perfs(0)
            If scores(1) > 0 Then
                sides(0) = sides(0) / 2
                sides(1) = sides(1) / 2
            End If
            If sides(0) < 0 Then sides(0) = 0
            If sides(1) < 0 Then sides(1) = 0
            For myplayer = 0 To 1
                If scores(myplayer) > 0 Then
                    perfs(0) = perfs(0) + (0.5 - perf)
                    perfs(1) = perfs(1) + (perf - 0.5)
                End If
            Next
            ' Stop '
            For myside = 0 To 1
                For myplayer = 0 To 1
                    pl = myside * 3 + myplayer
                    ' If nrating((players(pl))) = 0 Then Stop '
                    If num_games(players(pl)) > 9 Then
                        nrating((players(pl))) = nrating((players(pl))) + perfs(myside) * k_val
                    Else
                        nrating(players(pl)) = (nrating(players(pl)) * num_games(players(pl)) + sides(1 - myside)) / (num_games(players(pl)) + 1)
                    End If
                    nrating(players(pl)) = Abs(nrating(players(pl)))
                    num_games(players(pl)) = num_games(players(pl)) + 1
                Next
            Next
        End If    '}-> If Len(my_text$) Then
    Next    '}-> For i = 1 To hashsize
    ' Stop '
normal_exit_for_recalc:
    Call reset_placement
    For i = 1 To hashsize
        my_text$ = hasharray(i)
        If Len(my_text$) Then
            ' Stop '
            ret = parse_entry(my_text$, players, scores, quick_entry)
            Call player2row(players)
            If Index = 0 Then
                For myside = 0 To 1
                    For myplayer = 0 To 1
                        pl = myside * 3 + myplayer
                        opl = (1 - myside) * 3 + myplayer
                        If players(pl) And (players(pl - 1) <> players(pl)) Then
                            If MNU_Wide.Caption = "RoundRobin" Then
                                j = last_param_field + (row2RR(players(opl)) Mod (gcols - last_param_field))
                                If j <= last_param_field Then j = last_param_field + 1
                                While Chess.TextMatrix(players(pl), j) <> ""
                                    j = j + 1
                                    If j >= Chess.Cols Then j = last_param_field + 1
                                Wend
                            Else
                                j = find_empty(players(pl))
                            End If
                            Chess.TextMatrix(players(pl), j) = my_text$ + "_"
                            If Left$(Chess.TextMatrix(players(pl), attendance_field), 1) <> "X" Then Chess.TextMatrix(players(pl), attendance_field) = "X"
                            If ispresent(pl) > -1 Then
                                needs_resize = True
                                ispresent(pl) = -1
                                row2RR_count = row2RR_count + 1
                                row2RR(pl) = row2RR_count
                            End If
                        Else
                        ' Stop '
                        End If
                    Next
                Next
            End If
            If Index >= 2 Then     ' new day '
                ' we just need to increment the number of games '
                For myside = 0 To 1
                    For myplayer = 0 To 1
                        pl = myside * 3 + myplayer
                        ' If players(pl) And (players(pl - 1) <> players(pl)) Then '
                        Chess.TextMatrix(players(pl), Games_field) = Val(Chess.TextMatrix(players(pl), Games_field)) + 1
                    ' End If '
                    Next
                Next
            End If
        End If    '}-> If Len(my_text$) Then
    Next    '}-> For i = 1 To hashsize
    For myside = 0 To 1
        For myplayer = 0 To 1
            pl = myside * 3 + myplayer
            If players(pl) <> 0 And isvalid(players(pl)) = 0 Then
                On Error Resume Next
                Chess.Row = players(pl)
                Ladder_Form.Show
                DoEvents
                Call MNU_Edit_Player_Click
            End If
        Next
    Next
    For i = 1 To Chess.Rows - 1
        If isvalid(i) Then
            If (Val(Chess.TextMatrix(i, rating_field)) < 0) Then
                Chess.TextMatrix(i, nrating_field) = Str$(-Int(nrating(i)))
            Else
                Chess.TextMatrix(i, nrating_field) = Str$(Int(nrating(i)))
            End If
        Else
            Chess.TextMatrix(i, nrating_field) = "0"
        End If
    Next
    If Index >= 2 Then             ' new day '
        For i = 1 To Chess.Rows - 1
            If nrating(i) < 1 Then nrating(i) = 1
            If (Val(Chess.TextMatrix(i, rating_field)) < 0) Then
                Chess.TextMatrix(i, rating_field) = Str$(-Int(nrating(i)))
            Else
                Chess.TextMatrix(i, rating_field) = Str$(Int(nrating(i)))
            End If
            ret = Val(Chess.TextMatrix(i, attendance_field))
            If Chess.TextMatrix(i, attendance_field) = "X" Then
                Chess.TextMatrix(i, attendance_field) = " "
            Else
                Chess.TextMatrix(i, attendance_field) = Str(ret + 1)
            End If
        Next
        Call reset_hash(0)
    End If
    If Index = 2 Then              ' ==rerank '
        For i = 1 To Chess.Rows - 1
            Chess.TextMatrix(i, ranking_field) = Str$(i)
        Next
    End If
    ' Call sort_rank '
    If Index >= 2 Then             ' new day '
        Dim my_file$
        my_file$ = Replace(get_ladder_name(), ":", "_")
        my_file$ = Replace(my_file$, " ", "_")
        Call write_html_file(my_file$, ".html", 1)   '
        Call do_save
        Call reset_hash(0)
    End If
    If merr Then
error_exit_for_recalc:
        Chess.Row = i
        Chess.Col = j
    End If
    in_recalc = 0
    ' If Files.Sort_By_Name.ListIndex Then '
    ' Call Ladder_Form.Sort_Name           '
    ' End If                               '
    If needs_resize Then
        Call resize_chess
    End If
    '
    Chess.Redraw = True
End Sub
Public Function get_name$(Index As Integer)
    get_name$ = Chess.TextMatrix(Index, first_name_field) + Chess.TextMatrix(Index, last_name_field)
End Function
Public Sub Set_Sort_Name()
    Files.Sort_By_Name.ListIndex = Sort_Name
    Chess.Redraw = False
    Chess.Col = first_name_field
    Chess.ColSel = first_name_field
    Chess.Sort = 9
    Chess.Col = last_name_field
    Chess.ColSel = last_name_field
    Chess.Sort = 9
    Chess.Redraw = True
End Sub
Public Sub Set_Sort_First_Name()
    Files.Sort_By_Name.ListIndex = Sort_First_Name
    Chess.Redraw = False
    Chess.Col = last_name_field
    Chess.ColSel = last_name_field
    Chess.Sort = 9
    Chess.Col = first_name_field
    Chess.ColSel = first_name_field
    Chess.Sort = 9
    Chess.Redraw = True
End Sub
Public Sub Set_Sort_Rating()
    Files.Sort_By_Name.ListIndex = Sort_Rating
    Chess.Redraw = False
    Chess.Col = grade_field
    Chess.ColSel = grade_field
    Chess.Sort = 4
    Chess.Col = rating_field
    Chess.ColSel = rating_field
    Chess.Sort = 4
    Chess.Col = nrating_field
    Chess.ColSel = nrating_field
    Chess.Sort = 4
    Chess.Redraw = True
End Sub
Public Sub Set_sort_rank()
    Files.Sort_By_Name.ListIndex = sort_rank
    Chess.Redraw = False
    Chess.Col = grade_field
    Chess.ColSel = grade_field
    Chess.Sort = 2
    Chess.Col = ranking_field
    Chess.ColSel = ranking_field
    Chess.Sort = flexSortNumericAscending
    Chess.Redraw = True
End Sub
Public Sub Set_sort_room()
    Chess.Redraw = False
    Chess.Col = grade_field
    Chess.ColSel = grade_field
    Chess.Sort = 2
    Chess.Col = room_field
    Chess.ColSel = room_field
    Chess.Sort = flexSortStringAscending
    Chess.Redraw = True
End Sub
Private Sub MNU_SAVE_Click()
    On Error Resume Next
    Close
    save_flag = 0
    Call do_save
End Sub
Public Sub do_save()
    On Error GoTo save_error
    Call Set_Sort_Name
    Call write_file("Players.txt")
    Call write_file("Players.xls")
    Call Set_sort_rank
    Call write_file(get_ladder_name + ".xls")
    save_flag = 0
    On Error Resume Next
    Kill "Saved_Players.xls"
    Exit Sub
save_error:
    MsgBox ("save error")
    Resume
End Sub
Public Sub write_file(finame$)
    Dim i As Integer, j As Integer
    Dim my_text$
    Dim rank_array(grows_max) As Integer
    Dim lines_output As Integer, extra_lines As Integer
    On Error Resume Next
    Kill "Backup" & finame$
    Name finame$ As "Backup" & finame$
    Open finame$ For Output As #1
    For i = 0 To Chess.Rows - 1
        If Chess.TextMatrix(i, last_name_field) <> "" Then
            For j = 0 To Chess.Cols - 1
                my_text$ = Chess.TextMatrix(i, j)
                If Asc(Right$(my_text$, 1)) = 160 Then
                    my_text$ = Left$(my_text$, Len(my_text$) - 1)
                End If
                Print #1, my_text$; vbTab;
            Next j
            rank_array(Val(Chess.TextMatrix(i, ranking_field))) = i + 1
            If i = 0 Then
                Print #1, "Version 1.21"   ' must have 4 char to version num '
            Else
                Print #1, ""
            End If
            lines_output = lines_output + 1
        End If
    Next i
    For i = 0 To Chess.Rows - 1
        If rank_array(i) = 0 Then
            lines_output = lines_output + 1
            extra_lines = extra_lines + 1
            If extra_lines > 3 Then If lines_output / 20 = Int(lines_output / 20) Then Exit For
            For j = 0 To ranking_field - 1
                Print #1, vbTab;
            Next j
            Print #1, i, ""
        End If
    Next
    Close #1
    Close
End Sub
Private Sub MNU_Settings_Click()
    On Error Resume Next
    Settings.Show 1
    Call resize_chess
End Sub
Public Sub update_sorts()
    Select Case (Files.Sort_By_Name.ListIndex)
        Case Sort_Name
            Call Ladder_Form.Set_Sort_Name
        Case Sort_Rating
            Call Ladder_Form.Set_Sort_Rating
        Case sort_rank
            Call Ladder_Form.Set_sort_rank
        Case Sort_First_Name
            Call Ladder_Form.Set_Sort_First_Name
    End Select
    MNU_Sort.Caption = Files.Sort_By_Name.Text
    If MNU_Wide.Caption = "RoundRobin" Then
        Call resize_chess
        Call recalc(0)
    End If
End Sub
Private Sub MNU_Sort_Click()
    On Error Resume Next
    If Files.Sort_By_Name.ListIndex < 1 Then
        If Files.Sort_By_Name.ListIndex = -1 Then
            Files.Sort_By_Name.ListIndex = 1
        Else
            Files.Sort_By_Name.ListIndex = Files.Sort_By_Name.ListCount - 1
        End If
    Else
        Files.Sort_By_Name.ListIndex = Files.Sort_By_Name.ListIndex - 1
    End If
    Call update_sorts
    MNU_Sort.Caption = Files.Sort_By_Name.Text
End Sub
Private Sub MNU_Student_Report_Click()
    On Error GoTo resume_point
    Dim i As Integer, j As Integer
    Dim my_text$, l_table$, r_table$, end_of_line$, start_of_line$
    Dim last_name$, ladder_name$, src_body$, src_header$, src_tail$, working_body$, this_room$, my_room$
    Dim dataIn As String, this_player$
    Dim file1 As String, pos As Integer
    Dim root_files$, report_filename$, src_page_break$, adv_body$
    Dim vbquote$, rooms_printed As Integer
    ladder_name$ = get_ladder_name()
    vbquote$ = Chr$(34)
    src_page_break$ = "<p class=" + vbquote$ + "breakhere" + vbquote$ + "><br></p><p class=" + vbquote$ + "pagestart" + vbquote$ + "></p>"
    root_files$ = "..\Matthew.a.mahowald\"
    report_filename$ = "Student Enrolment.html"
    start_of_line$ = "<TR VALIGN=" + charqt + "bottom" + charqt + ">"
    end_of_line$ = "</TR>"
    ladder_name$ = get_ladder_name()
    Open "Student_Enrolment.html" For Input As #2
    Do
        Line Input #2, dataIn
        src_header$ = src_header$ + dataIn + vbCrLf
    Loop While InStr(LCase$(dataIn), "<body>") = 0
    Do
        Line Input #2, dataIn
        If InStr(LCase$(dataIn), "<\body>") Then Exit Do
        src_body$ = src_body$ + dataIn + vbCrLf
    Loop While Not EOF(2)
    src_tail$ = dataIn
    Do Until EOF(2)
        Line Input #2, dataIn
        src_tail$ = src_tail$ + dataIn + vbCrLf
    Loop
    Close #2
    Call Set_sort_room
    src_body$ = Replace(src_body$, "SCHOOL", ladder_name$)
    adv_body$ = Replace(src_body$, "ROOM", "Adventure Club")
    Open root_files$ + "report_" + ladder_name$ + ".html" For Output As #1
    Print #1, src_header$
    working_body$ = ""
    my_room$ = ""
    For i = 1 To Chess.Rows - 1
        this_room$ = LCase$(Chess.TextMatrix(i, room_field))
        this_player$ = Chess.TextMatrix(i, first_name_field) + " " + Chess.TextMatrix(i, last_name_field)
        If (InStr(this_room$, "adv")) Then
            adv_body$ = Replace(adv_body$, "STUDENT", this_player$ + "<br>STUDENT")
        End If
        pos = InStr(this_room$, ";")
        If (pos) Then this_room$ = Left$(this_room$, pos - 1)
        If this_room$ <> my_room$ Then
            ' we have a new  rooom '
            If rooms_printed Then
                Print #1, Replace(working_body$, "STUDENT", "")
                Print #1, src_page_break$
            End If
            rooms_printed = rooms_printed + 1
            working_body$ = src_body$
            working_body$ = Replace(working_body$, "ROOM", "Room " + this_room$)
        End If
        my_room$ = this_room$
        working_body$ = Replace(working_body$, "STUDENT", this_player$ + "<br>STUDENT")
    Next i
    Print #1, Replace(working_body$, "STUDENT", "")
    Print #1, src_page_break$
    Print #1, Replace(adv_body$, "STUDENT", "")
    Print #1, src_tail$
    Close #1
    Call Set_sort_rank
    Exit Sub
resume_point:
    Close
    Exit Sub
End Sub
Private Sub MNU_Trophies_Click()
    Dim sres As Integer
    Dim i As Integer
    Dim place As Integer
    Dim grades(15) As Integer
    Dim max As Integer
    Dim place_name(10) As String
    Dim my_grade As Integer
    place_name(0) = "Kindergarten"
    place_name(1) = "1st"
    place_name(2) = "2nd"
    place_name(3) = "3rd"
    place_name(4) = "4th"
    place_name(5) = "5th"
    place_name(6) = "6th"
    place_name(7) = "7th"
    place_name(8) = "8th"
    place_name(9) = "9th"
    place_name(10) = "10th"
    max = Val(Settings!Place_Trophies.Text)
    sres = MsgBox("This will room number column with trophy winners, are you sure", vbOKCancel)
    If sres <> vbOK Then Exit Sub
    place = 0
    ' grade_field,room_field '
    For i = 1 To Chess.Rows - 1
        If (Chess.TextMatrix(i, grade_field) <> "") And (Right(Ladder_Form!Chess.TextMatrix(i, group_field), 1) <> "x") Then
            If place >= max Then
                my_grade = Val(Chess.TextMatrix(i, grade_field))
                grades(my_grade) = grades(my_grade) + 1
                If (grades(my_grade) < 10) Then
                    Chess.TextMatrix(i, room_field) = place_name(grades(my_grade)) + " Place "
                Else
                    Chess.TextMatrix(i, room_field) = ""
                End If
                If (my_grade = 0) Then
                    Chess.TextMatrix(i, room_field) = Chess.TextMatrix(i, room_field) + place_name(my_grade)
                Else
                    Chess.TextMatrix(i, room_field) = Chess.TextMatrix(i, room_field) + place_name(my_grade) + " Grade"
                End If
            Else
                place = place + 1
                Chess.TextMatrix(i, room_field) = place_name(place) + " Overall"
            End If
        Else
            Chess.TextMatrix(i, room_field) = ""
        End If
    Next i
End Sub
Private Sub MNU_Wide_Click()
    Select Case MNU_Wide.Caption
        Case "Wide":
            entry_size = 300
            MNU_Wide.Caption = "Narrow"
        Case "RoundRobin"
            entry_size = 600
            MNU_Wide.Caption = "Wide"
        Case "Narrow"
            entry_size = 300
            MNU_Wide.Caption = "RoundRobin"
    End Select
    Call recalc(0)
    Call resize_chess
End Sub
Private Sub MNU_Zoom_Click()
    If AppZoom < 1.0001 Then
        AppZoom = 1.4
    Else
        AppZoom = 1
    End If
    Call resize_chess
End Sub
Private Sub MNUSetNumGamesToZero_Click()
    On Error Resume Next
    MiniGameType.Visible = True
End Sub

Private Sub slave_Error(Index As Integer, ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
'send entry to master
'Call DataHash(my_text$, my_text$, 0)
End Sub

Private Sub write_html_Click()
    If num_games_added < 1 Then
        num_games_added = 1
        Beep
        Exit Sub
    End If
    num_games_added = 0
    On Error GoTo no_write
    Close
    Dim my_file$
    my_file$ = Replace(get_ladder_name(), ":", "_")
    my_file$ = Replace(my_file$, " ", "_")
    ' we need one file for ,o; future weeks, and one for the present '
    Call write_html_file("..\Matthew.a.mahowald\" + my_file$, Replace(Date$ + "_" + Time$, ":", "_") + ".html", 0)   '
    Call write_html_file("..\Matthew.a.mahowald\" + my_file$, ".html", 0)                                            '
resume_point:
    Close
    Exit Sub
no_write:
    Resume resume_point
End Sub
Public Sub write_html_file(root_files$, my_file$, print_style As Integer)
    Dim i As Integer, j As Integer
    Dim my_text$, l_table$, r_table$, end_of_line$, start_of_line$
    Dim last_name$, ladder_name$
    Dim file1 As String
    l_table$ = "<TD><SMALL>"
    r_table$ = "</SMALL></TD>"
    start_of_line$ = "<TR VALIGN=" + charqt + "bottom" + charqt + ">"
    end_of_line$ = "</TR>"
    ladder_name$ = get_ladder_name()
    Open root_files$ + my_file$ For Output As #1
    Print #1, "<HTML><HEAD><TITLE>" + ladder_name$ + "</TITLE></HEAD>"
    Print #1, "<BODY>"
    Print #1, "<H1><CENTER>" + ladder_name$ + "<BR><SMALL>"
    Print #1, Date$ + " " + Time$
    Print #1, "</SMALL></CENTER></H1>"
    Print #1, "<a href=" + charqt + "index.html" + charqt + ">Home</a>"
    Print #1, "<FONT FACE=" + charqt + "Arial" + charqt + "><SMALL><Table border>"
    i = 0
    For i = 0 To Chess.Rows - 1
        If Chess.TextMatrix(i, last_name_field) = "" Then
            If print_style = 1 Then GoTo end_of_table
        End If
        Print #1, start_of_line$
        For j = 0 To Chess.Cols - 1
            If j = room_field Then
                j = j + 1
            End If
            my_text$ = Chess.TextMatrix(i, j)
            If j = last_name_field Then
                last_name$ = Left$(my_text$, last_name_field)
                j = j + 1
                my_text$ = Chess.TextMatrix(i, j) + " " + last_name$
            End If
            If j = Games_field Then
                j = school_field
            End If
            If my_text$ = "" Then my_text$ = " "
            If Right$(my_text$, 1) = "_" Then my_text$ = Left$(my_text$, Len(my_text$) - 1)
            If j = grade_field Then my_text$ = Left$(my_text$, 1)
            Print #1, l_table$
            If i = 0 And j > school_field Then
                my_text$ = "&nbsp;&nbsp;" + my_text + "&nbsp;&nbsp;"
            End If
            Print #1, Replace(my_text$, " ", "&nbsp;") + r_table$
        Next j
        Print #1, end_of_line$
    Next i
end_of_table:
    Print #1, "</Table></FONT><FONT SIZE=-1><I><BR>Last Updated on " + Date$
    Print #1, "<BR>By Matt Mahowald"
    Print #1, "<BR>Email: <A HREF = " + charqt + "mailto:qmerge@hotmail.com" + charqt + ">qmerge@hotmail.com</A></BODY></HTML>"
    Print #1, "<BR>Older weeks"
    file1 = Dir(root_files$ + "*.html")
    While file1 <> ""
        Print #1, "<BR><a href=" + charqt + file1 + charqt + ">" + file1 + "</a>"
        file1 = Dir()
    Wend
    Print #1, "</i></font></small></body></html>"
    Close #1
End Sub
Private Sub Print_Room_Sheet_MNU_Click()
    Dim i As Integer, j As Integer, k As Integer
    Dim my_text$, l_table$, r_table$, end_of_line$, start_of_line$
    Dim x As Integer
    Dim y As Integer
    Dim label_number As Integer
    Dim print_style As Integer
    Dim last_name$
    Dim file1 As String
    Dim ladder_name$
    Dim labels_per_page As Integer
    Dim xscale As Double
    Dim font_scale As Double
    Dim num_ys As Integer
    Dim num_xs As Integer
    Dim n As Integer
    Dim poffset As Double
    Dim pzoom As Double
    poffset = Val(Settings!print_offset) * 1600
    pzoom = Val(Settings!pzoom)
    If pzoom < 0.7 Then pzoom = 1
    If pzoom > 1.3 Then pzoom = 1
    pzoom = Sqr(pzoom)
    n = 1
    xscale = 2
    num_ys = 30
    num_xs = 1
    labels_per_page = num_ys * num_xs
    font_scale = Sqr(xscale)
    ladder_name$ = get_ladder_name()
    ' Call Sort_Name '
    Printer.fontsize = 9 * font_scale
    x = 100
    For k = 1 To n
        For i = 1 To Ladder_Form!Chess.Rows - 1
            If Ladder_Form!Chess.TextMatrix(i, last_name_field) <> "" Then
                If Right(Ladder_Form!Chess.TextMatrix(i, group_field), 1) <> "x" Then   '
                    ' If 1 Then '
                    label_number = label_number + 1
                    If label_number = labels_per_page + 1 Then
                        Printer.NewPage
                        label_number = 1
                    End If
                    ' x = ((label_number - 1) Mod num_xs) * 6057 + 250 '
                    y = pzoom * (Int((label_number - 1) / num_xs) * 1446 / 3 * 1.02) + 600
                    Printer.Line (x + 500 * xscale, y - 150)-(x + 5700 * xscale, y - 150)
                    j = 0          ' group '
                    my_text$ = Ladder_Form!Chess.TextMatrix(i, j)
                    Printer.PSet (x + 500 * xscale, y)
                    Printer.Print my_text$
                    j = 1          ' Last Name '
                    my_text$ = Ladder_Form!Chess.TextMatrix(i, j)
                    Printer.PSet (x + 800 * xscale, y)
                    Printer.Print my_text$
                    j = 2          ' First Name '
                    my_text$ = Ladder_Form!Chess.TextMatrix(i, j)
                    Printer.PSet (x + 1500 * xscale, y)
                    Printer.Print my_text$
                    j = 3          ' Rating '
                    my_text$ = Ladder_Form!Chess.TextMatrix(i, j)
                    Printer.PSet (x + 2200 * xscale, y)
                    If Show_Ratings_value(0) = True Then Printer.Print my_text$
                    j = 4          ' Rank '
                    my_text$ = Ladder_Form!Chess.TextMatrix(i, j)
                    Printer.PSet (x + 2500 * xscale, y)
                    Printer.Print my_text$
                    j = grade_field   ' Grade '
                    my_text$ = Ladder_Form!Chess.TextMatrix(i, j)
                    Printer.PSet (x + 2800 * xscale, y)
                    Printer.Print my_text$
                ' j = school_field   ' school_field                            '
                ' my_text$ = Ladder_Form!Chess.TextMatrix(i, j)                '
                ' Printer.PSet (x + 3000 * xscale, y)                          '
                ' If Show_Ratings_value(1) = True Then Printer.Print my_text$  '
                ' j = room_field   ' room_field                                '
                ' my_text$ = Ladder_Form!Chess.TextMatrix(i, j)                '
                ' Printer.PSet (x + 3400 * xscale, y)                          '
                ' If Show_Ratings_value(1) = False Then Printer.Print my_text$ '
                End If    '}-> If Right(Ladder_Form!Chess.TextMatrix(i, group_field), 1) <> "x" Then   '
            End If    '}-> If Ladder_Form!Chess.TextMatrix(i, last_name_field) <> "" Then
        Next i    '}-> For i = 1 To Ladder_Form!Chess.Rows - 1
    Next k    '}-> For k = 1 To n
    Printer.EndDoc
    Call Set_sort_rank
end_of_table:
End Sub



