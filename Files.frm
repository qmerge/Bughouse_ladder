VERSION 5.00
Begin VB.Form Files 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "files"
   ClientHeight    =   5196
   ClientLeft      =   6096
   ClientTop       =   3516
   ClientWidth     =   3444
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5196
   ScaleWidth      =   3444
   Begin VB.ComboBox Sort_By_Name 
      Height          =   288
      Left            =   240
      TabIndex        =   3
      Text            =   "Combo1"
      Top             =   4080
      Width           =   2172
   End
   Begin VB.CommandButton sort_rank 
      Caption         =   "Sort Rank"
      Height          =   255
      Left            =   360
      TabIndex        =   2
      Top             =   4920
      Width           =   1335
   End
   Begin VB.CommandButton Sort_Name 
      Caption         =   "Sort Name"
      Height          =   255
      Left            =   360
      TabIndex        =   1
      Top             =   4560
      Width           =   1215
   End
   Begin VB.FileListBox filelist 
      Height          =   3336
      Left            =   0
      Pattern         =   "*.txt"
      TabIndex        =   0
      Top             =   120
      Width           =   3015
   End
End
Attribute VB_Name = "Files"
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
Private Sub Combo1_Change()
End Sub
Private Sub Form_Load()
    Sort_By_Name.AddItem "Sort Rank"
    Sort_By_Name.AddItem "Sort Name"
    Sort_By_Name.AddItem "Sort First Name"
    Sort_By_Name.AddItem "Sort Rating"
    Sort_By_Name.ListIndex = 1
End Sub
Private Sub Sort_Name_Click()
    Call Ladder_Form.Set_Sort_Name
End Sub
Private Sub sort_rank_Click()
    Call Ladder_Form.Set_sort_rank
End Sub
Private Sub filelist_Click()
    Ladder_Form.load_file filelist.List(filelist.ListIndex)
End Sub

