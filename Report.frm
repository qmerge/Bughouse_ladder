VERSION 5.00
Begin VB.Form Report 
   Caption         =   "Pairing Report   <CTRL><ALT><d> to delete last entry"
   ClientHeight    =   7035
   ClientLeft      =   10320
   ClientTop       =   6435
   ClientWidth     =   7425
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   7035
   ScaleWidth      =   7425
   Begin VB.TextBox EntryLine 
      Alignment       =   2  'Center
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   7.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   732
      Left            =   120
      TabIndex        =   3
      ToolTipText     =   "<CTRL><ALT><d> to delete last entry"
      Top             =   5760
      Width           =   6012
   End
   Begin VB.Label Player 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   1452
      Index           =   2
      Left            =   120
      TabIndex        =   2
      Top             =   3600
      Width           =   5892
      WordWrap        =   -1  'True
   End
   Begin VB.Label Player 
      Alignment       =   2  'Center
      BackColor       =   &H0080C0FF&
      BorderStyle     =   1  'Fixed Single
      Height          =   1452
      Index           =   1
      Left            =   120
      TabIndex        =   1
      Top             =   1920
      Width           =   5892
      WordWrap        =   -1  'True
   End
   Begin VB.Label Player 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   1452
      Index           =   0
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   5892
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "Report"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit



Private Sub EntryLine_Change()
Call Form_KeyDown(0, 0)
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
    Static zero_set As Integer
    Idle_Timer_Tag = 0
    error_count = 0
    Dim Number As Integer
    Dim this_player As Integer
    Dim gscore As Integer
    Dim opponent As Integer
    Dim test_opponent As Integer
    Dim quick_entry As Integer
    Dim retVal As Long
    Dim display_string As String
    Dim my_text$
    display_string = "_WDL__"
    Ladder_Form!Idle_Timer.Enabled = False
    On Error Resume Next
    Select Case KeyCode
        Case vbKeyReturn:
'            If Ladder_Form!MNU_Wide.Caption = "RoundRobin" Then
'                Call resize_chess
'            End If
'            Call recalc(0)
        ' Chess.Row = Chess.Row + 1 '

        Case vbKeySubtract:
            EntryLine.Text = EntryLine.Text + "L"
            save_flag = save_flag + 1
        Case vbKeyAdd:
            EntryLine.Text = EntryLine.Text + "W"
            save_flag = save_flag + 1
        Case vbKeyMultiply:
            EntryLine.Text = EntryLine.Text + "D"
            save_flag = save_flag + 1
        Case vbKeyNumpad0:
            EntryLine.Text = EntryLine.Text + "0"
            Number = 1
        Case vbKeyNumpad1:
            EntryLine.Text = EntryLine.Text + "1"
            Number = 1
        Case vbKeyNumpad2:
            EntryLine.Text = EntryLine.Text + "2"
            Number = 1
        Case vbKeyNumpad3:
            EntryLine.Text = EntryLine.Text + "3"
            Number = 1
        Case vbKeyNumpad4:
            EntryLine.Text = EntryLine.Text + "4"
            Number = 1
        Case vbKeyNumpad5:
            EntryLine.Text = EntryLine.Text + "5"
            Number = 1
        Case vbKeyNumpad6:
            EntryLine.Text = EntryLine.Text + "6"
            Number = 1
        Case vbKeyNumpad7:
            EntryLine.Text = EntryLine.Text + "7"
            Number = 1
        Case vbKeyNumpad8:
            EntryLine.Text = EntryLine.Text + "8"
            Number = 1
        Case vbKeyNumpad9:
            EntryLine.Text = EntryLine.Text + "9"
            Number = 1
        Case 0:
            zero_set = 2
        Case vbKeyEscape
            EntryLine.Text = ""
        Case vbKeyCancel:
            EntryLine.Text = ""
        Case vbKeyBack:
            save_flag = save_flag + 1
            If Len(EntryLine.Text) Then
                EntryLine.Text = Left$(EntryLine.Text, Len(EntryLine.Text) - 1)
            End If
        Case vbKeyDelete:
            save_flag = save_flag + 1
            If Len(EntryLine.Text) Then
                EntryLine.Text = Left$(EntryLine.Text, Len(EntryLine.Text) - 1)
            End If
        Case vbKeyClear:
            EntryLine.Text = ""
        Case 186:
            EntryLine.Text = EntryLine.Text + ":"
        Case vbKeyControl:
        Case 18:
        Case vbKeyShift:           ' shift    '
        Case Else:                 ' Asc("z") '
            If Shift And 2 Then 'ctrl
                If Shift = 6 And KeyCode = 68 Then
                    Ladder_Form!cmd.Text = "-" + last_commit
                    last_commit = ""
                    'Call delete_matching_Cell   ' ctrl_alt_d '
                End If
                If KeyCode = 86 Then
                    EntryLine.Text = EntryLine.Text + Clipboard.GetText
                End If
            Else
                If (KeyCode < ascz And KeyCode > ascsp) Then
                    If Shift < 4 Then
                        EntryLine.Text = EntryLine.Text + Chr$(KeyCode)
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
    Ladder_Form!Idle_Timer.Enabled = True
    If zero_set > 0 Then zero_set = zero_set - 1
    my_text$ = EntryLine.Text
    reports(1 - Me.Tag)!EntryLine.Text = EntryLine.Text
    retVal = parse_entry(my_text$, players, scores, quick_entry)
    
    If (players(3) > 0) Then
        Player(0).Caption = get_fullName$(p2r(players(3))) + "(" + LTrim$(Str(players(3))) + ")"
    Else
        Player(0).Caption = ""
    End If
    If (players(0) > 0) Then
        Player(2).Caption = get_fullName$(p2r(players(0))) + "(" + LTrim$(Str(players(0))) + ")"
    Else
        Player(2).Caption = ""
    End If
    Player(1).Caption = ""
    If Len(my_text$) < 2 Then Exit Sub
    Player(1).Caption = Player(1).Caption + Mid$(display_string, scores(0) + 1, 1)
    Player(1).Caption = Player(1).Caption + Mid$(display_string, scores(1) + 1, 1)
    If retVal > 0 Then
        Me.Caption = "Valid Entry"
        If vbKeyReturn = KeyCode Then
            my_text$ = long2string(retVal)
            last_commit = my_text
            Ladder_Form!cmd.Text = my_text$
            EntryLine.Text = ""
            Exit Sub
        End If
        Call setAllGood
        Exit Sub
    End If
    Me.Caption = ""
    Call setAllNeutral
    Select Case retVal
        Case -1:
        Case -2:
        Case -3:
            Player(2).BackColor = &H80C0FF
        Case -4:
            Me.Caption = "Duplicate Player 1"
            Player(2).BackColor = &H80C0FF
        Case -5:
        Case -6:
        Case -7:
    End Select
End Sub
Sub setAllGood()
On Error Resume Next
Dim i As Integer
For i = 0 To 2
    Player(i).BackColor = &H80FF80
Next i
End Sub
Sub setAllNeutral()
On Error Resume Next
Dim i As Integer
For i = 0 To 2
    Player(i).BackColor = &H8000000F
Next
End Sub
Private Sub Form_Resize()
On Error Resume Next
Dim i As Integer
Dim fontsize As Double
Dim heightspace As Double
Dim widthspace As Double
heightspace = 200
widthspace = 300
fontsize = Me.Height - heightspace * 6
'Player(0).Caption = Str(fontsize)
If fontsize > Me.Width / 4 Then fontsize = Me.Width / 4
fontsize = fontsize / 66
'Player(1).Caption = Str(fontsize)
For i = 0 To 2
    Player(i).FontName = EntryLine.FontName
    Player(i).Left = widthspace / 8
    Player(i).Width = (Me.Width - widthspace)
    Player(i).Top = i / 4 * (Me.Height - heightspace * 3) + heightspace / 8
    Player(i).Height = 1 / 4 * (Me.Height) - heightspace
    Player(i).fontsize = fontsize
Next
'Player(2).Caption = Player(2).fontsize
EntryLine.Left = widthspace / 8
EntryLine.Width = (Me.Width - widthspace)
EntryLine.Top = i / 4 * (Me.Height - heightspace * 3) + heightspace / 8
EntryLine.Height = 1 / 4 * (Me.Height) - heightspace
EntryLine.fontsize = fontsize
End Sub
Public Function get_fullName$(Index As Integer)
    get_fullName$ = Ladder_Form!Chess.TextMatrix(Index, first_name_field) + " " + Ladder_Form!Chess.TextMatrix(Index, last_name_field)
End Function

