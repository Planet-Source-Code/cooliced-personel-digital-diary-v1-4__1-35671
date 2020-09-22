VERSION 5.00
Begin VB.Form FrmFirstRun 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "First Run"
   ClientHeight    =   3735
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   Icon            =   "FrmFirstRun.frx":0000
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3735
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   4080
      Picture         =   "FrmFirstRun.frx":0442
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   9
      Top             =   0
      Width           =   495
   End
   Begin VB.Frame Frame1 
      Caption         =   "Options"
      Height          =   2055
      Left            =   0
      TabIndex        =   2
      Top             =   1680
      Width           =   4695
      Begin VB.CommandButton CmdCancel 
         Caption         =   "Cancel"
         Height          =   255
         Left            =   3600
         TabIndex        =   8
         Top             =   1680
         Width           =   975
      End
      Begin VB.CommandButton CmdOk 
         Caption         =   "OK"
         Height          =   255
         Left            =   2400
         TabIndex        =   7
         Top             =   1680
         Width           =   1095
      End
      Begin VB.CheckBox Chk2 
         Caption         =   "Always startup when Windows starts"
         Height          =   255
         Left            =   240
         TabIndex        =   6
         Top             =   720
         Width           =   4215
      End
      Begin VB.CheckBox Chk1 
         Caption         =   "Always login to Dairy V1.4"
         Height          =   195
         Left            =   240
         TabIndex        =   5
         Top             =   360
         Width           =   3855
      End
      Begin VB.TextBox TxtNewPass 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   120
         MaxLength       =   12
         PasswordChar    =   "*"
         TabIndex        =   3
         Top             =   1320
         Width           =   4455
      End
      Begin VB.Label Label4 
         Caption         =   "Note: Max length 12 characters"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   1080
         Width           =   4455
      End
   End
   Begin VB.Label Label2 
      Caption         =   "NOTE: You MUST enter a password even if you do not wish to login!"
      BeginProperty Font 
         Name            =   "System"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   10
      Top             =   1200
      Width           =   4455
   End
   Begin VB.Label Label3 
      Caption         =   $"FrmFirstRun.frx":0884
      Height          =   615
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   4455
   End
   Begin VB.Label Label1 
      Caption         =   "Welcome to Diary V1.4"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3735
   End
End
Attribute VB_Name = "FrmFirstRun"
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
    '        Copyright © Cooliced - Cooliced.co.uk 1999-2002         '
    '***************************************************************'

Private Sub CmdCancel_Click()
End
End Sub

Private Sub cmdOk_Click()
Dim strNewPW As String
Dim strEncryptNewPW As String
Dim BoxVal As String


    If TxtNewPass.text = "" Then
    MsgBox "You MUST enter a password!", vbError Or vbOKOnly, "Diary 1.4"
    Exit Sub
    End If

        SaveSetting "Diary", "Main", "FirstRun", 1 'make it not first run
' ********************* Password *************************
        
        strNewPW = LCase(TxtNewPass.text)
        strEncryptNewPW = Converter(strNewPW)
        SaveSetting "Diary", "Main", "Password", strEncryptNewPW
' *********************  Login  **************************
        If Chk1.Value = 0 Then
         BoxVal = "no"
         SaveSetting "Diary", "Main", "Login", BoxVal
        ElseIf Chk1.Value = 1 Then
         BoxVal = "yes"
         SaveSetting "Diary", "Main", "Login", BoxVal
        End If
' ********************* Starup ***************************
        If Chk2.Value = 0 Then
         SaveSetting "Diary", "Main", "Startup", 0
         SaveString HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Run", "Diary", "Diary"
         SaveString HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Run", "Diary", "0"
        ElseIf Chk2.Value = 1 Then
         SaveSetting "Diary", "Main", "Startup", 1
         SaveString HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Run", "Diary", "Diary"
         SaveString HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Run", "Diary", App.path & "\Diary.exe"
        End If
' ********************* assoc ****************************
        Associate "Diary", ".ccd", "Diary file", App.path & "\net14.ico"
' ********************************************************
      MsgBox "Saving settings and reseting the program", vbExclamation Or vbOKOnly, "Diary 1.4"
      frmSplash.Show
      Unload Me
End Sub

Private Sub Form_Load()
Chk1.Value = 1
If FrmFirstRun.Caption <> "Diary 1.4 - First Run" Then
FrmFirstRun.Caption = "Diary 1.4 - First Run"
End If
End Sub
