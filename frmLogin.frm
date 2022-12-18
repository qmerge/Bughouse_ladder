VERSION 5.00
Begin VB.Form frmLogin 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Login"
   ClientHeight    =   1545
   ClientLeft      =   3405
   ClientTop       =   6060
   ClientWidth     =   3750
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   912.837
   ScaleMode       =   0  'User
   ScaleWidth      =   3521.047
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   390
      Left            =   495
      TabIndex        =   2
      Top             =   1020
      Width           =   1140
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   390
      Left            =   2100
      TabIndex        =   3
      Top             =   1020
      Width           =   1140
   End
   Begin VB.TextBox txtPassword 
      Height          =   345
      IMEMode         =   3  'DISABLE
      Left            =   1320
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   525
      Width           =   2325
   End
   Begin VB.Label lblLabels 
      Caption         =   "&Password:"
      Height          =   270
      Index           =   1
      Left            =   105
      TabIndex        =   0
      Top             =   540
      Width           =   1080
   End
End
Attribute VB_Name = "frmLogin"
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
Public LoginSucceeded As Boolean
Private Sub cmdCancel_Click()
    ' set the global var to false '
    ' to denote a failed login    '
    LoginSucceeded = False
    Me.Hide
End Sub
Private Sub cmdOK_Click()
    ' check for correct password '
    If txtPassword = mpassword$ Then
        ' place code to here to pass the      '
        ' success to the calling sub          '
        ' setting a global var is the easiest '
        LoginSucceeded = True
        Me.Hide
    Else
        MsgBox "Invalid Password, try again!", , "Login"
        txtPassword.SetFocus
        On Error Resume Next
        SendKeys "{Home}+{End}"
        LoginSucceeded = False
    End If
    txtPassword.Text = ""
End Sub
Private Sub Form_Activate()
    If txtPassword = mpassword$ Then Call cmdOK_Click
End Sub
Private Sub txtPassword_Change()
    If txtPassword = mpassword$ Then Call cmdOK_Click
End Sub
