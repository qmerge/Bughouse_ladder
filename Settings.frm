VERSION 5.00
Begin VB.Form Settings 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Settings"
   ClientHeight    =   7875
   ClientLeft      =   2580
   ClientTop       =   1650
   ClientWidth     =   4215
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7875
   ScaleWidth      =   4215
   Begin VB.TextBox MasterIP 
      Height          =   285
      Left            =   1200
      TabIndex        =   26
      Text            =   "127.0.0.1"
      Top             =   7560
      Width           =   1815
   End
   Begin VB.CheckBox Show_Ratings 
      Caption         =   "Auto_Letter"
      Height          =   195
      Index           =   3
      Left            =   1800
      TabIndex        =   25
      Tag             =   "11"
      Top             =   2880
      Value           =   1  'Checked
      Width           =   1455
   End
   Begin VB.TextBox pzoom 
      Height          =   285
      Left            =   1200
      TabIndex        =   23
      Top             =   7080
      Width           =   2652
   End
   Begin VB.TextBox print_offset 
      Height          =   285
      Left            =   1200
      TabIndex        =   21
      Top             =   6600
      Width           =   2652
   End
   Begin VB.CheckBox Show_Ratings 
      Caption         =   "Print_Room"
      Height          =   195
      Index           =   2
      Left            =   1800
      TabIndex        =   20
      Tag             =   "11"
      Top             =   3120
      Value           =   1  'Checked
      Width           =   1455
   End
   Begin VB.TextBox Coach 
      Height          =   285
      Index           =   2
      Left            =   1200
      TabIndex        =   18
      Top             =   6120
      Width           =   2652
   End
   Begin VB.TextBox Coach 
      Height          =   285
      Index           =   1
      Left            =   1200
      TabIndex        =   16
      Top             =   5880
      Width           =   2652
   End
   Begin VB.TextBox Coach 
      Height          =   285
      Index           =   0
      Left            =   1200
      TabIndex        =   14
      Top             =   5640
      Width           =   2652
   End
   Begin VB.TextBox Place_Trophies 
      Height          =   372
      Left            =   2160
      TabIndex        =   12
      Text            =   "5"
      Top             =   5040
      Width           =   1452
   End
   Begin VB.CheckBox Show_Ratings 
      Caption         =   "Show School not room"
      Height          =   312
      Index           =   1
      Left            =   360
      TabIndex        =   11
      Tag             =   "11"
      Top             =   3480
      Value           =   1  'Checked
      Width           =   1332
   End
   Begin VB.TextBox txtPassword 
      Height          =   345
      IMEMode         =   3  'DISABLE
      Left            =   1575
      PasswordChar    =   "*"
      TabIndex        =   9
      Top             =   3960
      Width           =   2325
   End
   Begin VB.CheckBox Show_Ratings 
      Caption         =   "Show_Ratings"
      Height          =   195
      Index           =   0
      Left            =   1800
      TabIndex        =   8
      Tag             =   "5"
      Top             =   3480
      Value           =   1  'Checked
      Width           =   1692
   End
   Begin VB.TextBox set_grows 
      Height          =   285
      Left            =   2880
      TabIndex        =   6
      Text            =   "100"
      Top             =   2400
      Width           =   855
   End
   Begin VB.CommandButton Done 
      Caption         =   "Done"
      Height          =   375
      Left            =   2400
      TabIndex        =   5
      Top             =   4560
      Width           =   1095
   End
   Begin VB.CommandButton Abort 
      Caption         =   "Abort"
      Height          =   375
      Left            =   240
      TabIndex        =   4
      Top             =   4560
      Width           =   1095
   End
   Begin VB.TextBox K_Factor 
      Height          =   375
      Left            =   1440
      TabIndex        =   1
      Text            =   "32"
      Top             =   2160
      Width           =   855
   End
   Begin VB.Label Label_coaches 
      Caption         =   "Print Zoom"
      Height          =   372
      Index           =   4
      Left            =   120
      TabIndex        =   24
      Top             =   7080
      Width           =   972
   End
   Begin VB.Label Label_coaches 
      Caption         =   "Print Offset Inches"
      Height          =   492
      Index           =   3
      Left            =   240
      TabIndex        =   22
      Top             =   6480
      Width           =   972
   End
   Begin VB.Label Label_coaches 
      Caption         =   "Coach3"
      Height          =   252
      Index           =   2
      Left            =   240
      TabIndex        =   19
      Top             =   6120
      Width           =   1092
   End
   Begin VB.Label Label_coaches 
      Caption         =   "Coach2"
      Height          =   252
      Index           =   1
      Left            =   240
      TabIndex        =   17
      Top             =   5880
      Width           =   1092
   End
   Begin VB.Label Label_coaches 
      Caption         =   "Coach1"
      Height          =   252
      Index           =   0
      Left            =   240
      TabIndex        =   15
      Top             =   5640
      Width           =   1092
   End
   Begin VB.Label Label5 
      Caption         =   "Place Trophies"
      Height          =   252
      Left            =   480
      TabIndex        =   13
      Top             =   5040
      Width           =   1452
   End
   Begin VB.Label lblLabels 
      Caption         =   "&Password:"
      Height          =   276
      Index           =   1
      Left            =   360
      TabIndex        =   10
      Top             =   3972
      Width           =   1080
   End
   Begin VB.Label Label4 
      Caption         =   "Num of Rows"
      Height          =   372
      Left            =   2760
      TabIndex        =   7
      Top             =   2040
      Width           =   1092
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "K factor   "
      Height          =   252
      Left            =   240
      TabIndex        =   3
      Top             =   2280
      Width           =   1332
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Amount of points won or lost in a game"
      Height          =   732
      Left            =   120
      TabIndex        =   2
      Top             =   2640
      Width           =   1332
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "     Rated Ladder                 By              Matt Mahowald"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2055
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   3735
   End
End
Attribute VB_Name = "Settings"
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
Private Sub Abort_Click()
    On Error Resume Next
    Settings.Hide
End Sub
Private Sub Done_Click()
    On Error Resume Next
    k_val_base = Val(Settings!K_Factor.Text)
    Open "ladder.ini" For Output As #1
    Print #1, Str$(k_val_base)
    Print #1, Str$(grows)
    Print #1, Str$(Show_Ratings(0).Value)
    Print #1, opassword$
    Print #1, Str$(Show_Ratings(1).Value)
    Print #1, Str$(Show_Ratings(2).Value)
    Print #1, Place_Trophies.Text
    Print #1, Coach(0).Text
    Print #1, Coach(1).Text
    Print #1, Coach(2).Text
    Print #1, print_offset.Text
    Print #1, pzoom.Text
    Print #1, Str$(Show_Ratings(3).Value)
    Print #1, MasterIP.Text
    Close
    Settings.Hide
End Sub
Private Sub Form_Load()
    On Error Resume Next
    K_Factor.Text = Str$(k_val_base)
End Sub
Private Sub set_grows_Change()
    grows = set_grows.Text
End Sub
Private Sub Show_Ratings_Click(Index As Integer)
    Show_Ratings_value(Index) = Show_Ratings(Index).Value
End Sub
Private Sub txtPassword_Change()
    opassword$ = txtPassword.Text
End Sub
