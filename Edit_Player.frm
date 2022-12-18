VERSION 5.00
Begin VB.Form Edit_Player 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Edit_Player"
   ClientHeight    =   6915
   ClientLeft      =   5640
   ClientTop       =   2340
   ClientWidth     =   4920
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6915
   ScaleWidth      =   4920
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox Edit_Box 
      Height          =   300
      Index           =   12
      Left            =   3840
      TabIndex        =   37
      Text            =   "Edit_Box"
      Top             =   5040
      Width           =   1032
   End
   Begin VB.TextBox Edit_Box 
      Height          =   300
      Index           =   11
      Left            =   2520
      TabIndex        =   8
      Text            =   "Edit_Box"
      Top             =   5880
      Width           =   2355
   End
   Begin VB.TextBox Edit_Box 
      Height          =   300
      Index           =   10
      Left            =   120
      TabIndex        =   6
      Text            =   "Edit_Box"
      Top             =   6480
      Width           =   4635
   End
   Begin VB.TextBox Edit_Box 
      Height          =   300
      Index           =   5
      Left            =   120
      TabIndex        =   5
      Text            =   "Edit_Box"
      Top             =   5760
      Width           =   1275
   End
   Begin VB.CommandButton Save_Next 
      Caption         =   "Save && Exit"
      Height          =   495
      Index           =   2
      Left            =   2640
      TabIndex        =   13
      Top             =   2520
      Width           =   1800
   End
   Begin VB.CommandButton Letter 
      Caption         =   "E"
      Height          =   375
      Index           =   4
      Left            =   1920
      TabIndex        =   32
      Top             =   120
      Width           =   375
   End
   Begin VB.CommandButton Letter 
      Caption         =   "D"
      Height          =   375
      Index           =   3
      Left            =   1560
      TabIndex        =   31
      Top             =   120
      Width           =   375
   End
   Begin VB.CommandButton Letter 
      Caption         =   "C"
      Height          =   375
      Index           =   2
      Left            =   1200
      TabIndex        =   30
      Top             =   120
      Width           =   375
   End
   Begin VB.CommandButton Letter 
      Caption         =   "B"
      Height          =   375
      Index           =   1
      Left            =   840
      TabIndex        =   29
      Top             =   120
      Width           =   375
   End
   Begin VB.CommandButton Letter 
      Caption         =   "A"
      Height          =   375
      Index           =   0
      Left            =   480
      TabIndex        =   28
      Top             =   120
      Width           =   375
   End
   Begin VB.TextBox Edit_Box 
      Height          =   300
      Index           =   6
      Left            =   2400
      TabIndex        =   7
      Text            =   "Edit_Box"
      Top             =   4200
      Width           =   1875
   End
   Begin VB.TextBox Edit_Box 
      Height          =   300
      Index           =   9
      Left            =   120
      TabIndex        =   4
      Text            =   "Edit_Box"
      Top             =   5040
      Width           =   3672
   End
   Begin VB.TextBox Edit_Box 
      Height          =   300
      Index           =   8
      Left            =   3840
      TabIndex        =   9
      Text            =   "Edit_Box"
      Top             =   3360
      Width           =   915
   End
   Begin VB.TextBox Edit_Box 
      Height          =   300
      Index           =   7
      Left            =   120
      TabIndex        =   14
      Text            =   "Edit_Box"
      Top             =   4200
      Width           =   2000
   End
   Begin VB.CommandButton Revert 
      Caption         =   "Revert"
      Height          =   495
      Left            =   2640
      TabIndex        =   23
      Top             =   1920
      Width           =   1800
   End
   Begin VB.CommandButton Save_Next 
      Caption         =   "Save && Go to Previous Player"
      Height          =   495
      Index           =   1
      Left            =   2640
      TabIndex        =   11
      Top             =   1320
      Width           =   1800
   End
   Begin VB.CommandButton Save_Next 
      Caption         =   "Save && Go to Next Player"
      Height          =   495
      Index           =   0
      Left            =   2640
      TabIndex        =   10
      Top             =   720
      Width           =   1800
   End
   Begin VB.TextBox Player_Rank 
      Height          =   375
      Left            =   2400
      TabIndex        =   12
      Text            =   "Player_Rank"
      Top             =   3360
      Width           =   975
   End
   Begin VB.CommandButton Clear_All 
      Caption         =   "       Clear All       (Detetes Player)"
      Height          =   495
      Left            =   2640
      TabIndex        =   21
      Top             =   120
      Width           =   1800
   End
   Begin VB.TextBox Edit_Box 
      Height          =   285
      Index           =   4
      Left            =   6240
      TabIndex        =   19
      Text            =   "Edit_Box"
      Top             =   1800
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.TextBox Edit_Box 
      Height          =   300
      Index           =   3
      Left            =   100
      TabIndex        =   3
      Text            =   "Edit_Box"
      Top             =   3480
      Width           =   2000
   End
   Begin VB.TextBox Edit_Box 
      Height          =   300
      Index           =   2
      Left            =   100
      TabIndex        =   2
      Text            =   "Edit_Box"
      Top             =   2655
      Width           =   2000
   End
   Begin VB.TextBox Edit_Box 
      Height          =   300
      Index           =   1
      Left            =   120
      TabIndex        =   1
      Text            =   "Edit_Box"
      Top             =   1830
      Width           =   2000
   End
   Begin VB.TextBox Edit_Box 
      Height          =   300
      Index           =   0
      Left            =   100
      TabIndex        =   0
      Text            =   "Edit_Box"
      Top             =   1005
      Width           =   2000
   End
   Begin VB.CommandButton Letter 
      Caption         =   "A1"
      Height          =   375
      Index           =   5
      Left            =   120
      TabIndex        =   36
      Top             =   120
      Width           =   375
   End
   Begin VB.Label Edit_Label 
      Caption         =   "Edit_Label"
      Height          =   300
      Index           =   12
      Left            =   3840
      TabIndex        =   38
      Top             =   4680
      Width           =   1752
   End
   Begin VB.Label Edit_Label 
      Caption         =   "Edit_Label"
      Height          =   300
      Index           =   11
      Left            =   2520
      TabIndex        =   35
      Top             =   5520
      Width           =   1995
   End
   Begin VB.Label Edit_Label 
      Caption         =   "Edit_Label"
      Height          =   300
      Index           =   10
      Left            =   120
      TabIndex        =   34
      Top             =   6120
      Width           =   1995
   End
   Begin VB.Label Edit_Label 
      Caption         =   "Edit_Label"
      Height          =   300
      Index           =   5
      Left            =   120
      TabIndex        =   33
      Top             =   5400
      Width           =   1995
   End
   Begin VB.Label Edit_Label 
      Caption         =   "Edit_Label"
      Height          =   300
      Index           =   6
      Left            =   2400
      TabIndex        =   27
      Top             =   3960
      Width           =   1995
   End
   Begin VB.Label Edit_Label 
      Caption         =   "Edit_Label"
      Height          =   300
      Index           =   9
      Left            =   120
      TabIndex        =   26
      Top             =   4680
      Width           =   1995
   End
   Begin VB.Label Edit_Label 
      Caption         =   "Edit_Label"
      Height          =   300
      Index           =   8
      Left            =   3960
      TabIndex        =   25
      Top             =   3120
      Width           =   1155
   End
   Begin VB.Label Edit_Label 
      Caption         =   "Edit_Label"
      Height          =   300
      Index           =   7
      Left            =   120
      TabIndex        =   24
      Top             =   3840
      Width           =   1995
   End
   Begin VB.Label Label1 
      Caption         =   "Player Rank"
      Height          =   255
      Left            =   2400
      TabIndex        =   22
      Top             =   3120
      Width           =   1215
   End
   Begin VB.Label Edit_Label 
      Caption         =   "Edit_Label"
      Height          =   375
      Index           =   4
      Left            =   6720
      TabIndex        =   20
      Top             =   600
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.Label Edit_Label 
      Caption         =   "Edit_Label"
      Height          =   300
      Index           =   3
      Left            =   105
      TabIndex        =   18
      Top             =   3060
      Width           =   1995
   End
   Begin VB.Label Edit_Label 
      Caption         =   "Edit_Label"
      Height          =   300
      Index           =   2
      Left            =   105
      TabIndex        =   17
      Top             =   2250
      Width           =   1995
   End
   Begin VB.Label Edit_Label 
      Caption         =   "Edit_Label"
      Height          =   300
      Index           =   1
      Left            =   105
      TabIndex        =   16
      Top             =   1425
      Width           =   1995
   End
   Begin VB.Label Edit_Label 
      Caption         =   "Edit_Label"
      Height          =   300
      Index           =   0
      Left            =   105
      TabIndex        =   15
      Top             =   600
      Width           =   1995
   End
End
Attribute VB_Name = "Edit_Player"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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
'                                                                             '
Option Explicit
Private Sub Clear_All_Click()
    On Error Resume Next
    Dim i As Integer
    For i = 0 To Ladder_Form!Chess.FixedCols
        If i <> ranking_field Then
            Edit_Box(i).Text = ""
        End If
    Next
End Sub



Private Sub Letter_Click(Index As Integer)
    Edit_Box(0).Text = Letter(Index).Caption
    Call Save_Next_Click(0)
End Sub
Private Sub Player_Rank_Change()
    On Error Resume Next
    Dim i As Integer
    Dim rank As Integer
    rank = Val(Player_Rank)
    If rank < 1 Then
        ' oops we need to fix player_rank '
        rank = Ladder_Form!Chess.Rows - 1
        Player_Rank = Str$(rank)
        Exit Sub
    End If
    If rank > Ladder_Form!Chess.Rows - 1 Then
        ' oops we need to fix player_rank '
        rank = 1
        Player_Rank = Str$(rank)
        Exit Sub
    End If
    On Error Resume Next           ' just skip over rank window '
    For i = 0 To Ladder_Form!Chess.FixedCols
        Edit_Label(i).Caption = Ladder_Form!Chess.TextMatrix(0, i)
        Edit_Box(i).Text = Ladder_Form!Chess.TextMatrix(rank, i)
    Next
End Sub
Private Sub Revert_Click()
    On Error Resume Next
    Call Player_Rank_Change
End Sub
Private Sub Save_Next_Click(Index As Integer)
    On Error Resume Next
    Dim i As Integer
    Dim rank As Integer
    If InStr(GROUP_CODES, Edit_Box(0).Text) = 0 Then Edit_Box(0).Text = " "
    rank = Val(Player_Rank)
    For i = 0 To Ladder_Form!Chess.FixedCols
        Ladder_Form!Chess.TextMatrix(rank, i) = LTrim$(RTrim$(Edit_Box(i).Text))
    Next
    If Index = 2 Then
        Call Player_Rank_Change
        Edit_Player.Hide
    End If
    If Index = 0 Then
        Player_Rank = Str$(Val(Player_Rank) + 1)
    Else
        Player_Rank = Str$(Val(Player_Rank) - 1)
    End If
End Sub
