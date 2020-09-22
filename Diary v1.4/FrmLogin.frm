VERSION 5.00
Begin VB.Form FrmLogin 
   Appearance      =   0  'Flat
   BackColor       =   &H8000000A&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Password Protected"
   ClientHeight    =   1470
   ClientLeft      =   5100
   ClientTop       =   5280
   ClientWidth     =   2895
   Icon            =   "FrmLogin.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1470
   ScaleWidth      =   2895
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CmdCancle 
      Caption         =   "Cancel"
      Height          =   255
      Left            =   1560
      TabIndex        =   3
      Top             =   1080
      Width           =   855
   End
   Begin VB.CommandButton CmdOk 
      Caption         =   "Ok"
      Height          =   255
      Left            =   360
      TabIndex        =   2
      Top             =   1080
      Width           =   735
   End
   Begin VB.TextBox txtPassword 
      BeginProperty Font 
         Name            =   "System"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   120
      MaxLength       =   12
      PasswordChar    =   "*"
      TabIndex        =   0
      Top             =   600
      Width           =   2655
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   2280
      Picture         =   "FrmLogin.frx":0442
      Top             =   120
      Width           =   480
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Password:"
      BeginProperty Font 
         Name            =   "System"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   240
      Width           =   1095
   End
End
Attribute VB_Name = "FrmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
    '***************************************************************'
    '                       "Diary 1.4"                            '
    '                        Written by                             '
    '                         Cooliced                              '
    '                                                               '
    '  You are free to use the source code in your private,         '
    '  non-commercial, projects with permission.    If you want     '
    '  to use this code in commercial projects EXPLICIT permission  '
    '  from the author is required.                                 '
    '                                                               '
    '                                                               '
    '        Copyright © Cooliced - Cooliced.c.uk 1999-2002         '
    '***************************************************************'



Private Sub CmdCancle_Click()
End
End Sub

Private Sub cmdOk_Click()
    Dim strTest As String
    
    strTest = GetSetting("Diary", "Main", "Password", passval)
    
     If LCase(txtPassword.Text) = Converter(strTest) Then
        ' show
        FrmMain.Show
        ' The name of the main application
        Me.Hide
        ' Hides the login dialog box
        
    Else 'incorrect password!
        MsgBox "Enter a Valid Password for this System", 8, "Password Error"
        txtPassword.SetFocus
        Exit Sub
        
    End If
End Sub

Private Sub Form_Load()
If App.PrevInstance Then
        AppActivate "Diary 1.4"
        End
    End If
If FrmLogin.Caption <> "Diary 1.4 - Login" Then
FrmLogin.Caption = "Diary 1.4 - Login"
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    End ' end program
End Sub


Private Sub txtPassword_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
cmdOk_Click
End If
End Sub
