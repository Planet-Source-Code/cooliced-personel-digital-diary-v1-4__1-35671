VERSION 5.00
Begin VB.Form FrmOptions 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Options"
   ClientHeight    =   2175
   ClientLeft      =   6795
   ClientTop       =   4320
   ClientWidth     =   2910
   FillColor       =   &H00404040&
   FillStyle       =   0  'Solid
   Icon            =   "FrmOptions.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2175
   ScaleWidth      =   2910
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox Check2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
      Caption         =   "Always Startup With Windows"
      BeginProperty Font 
         Name            =   "System"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   120
      TabIndex        =   3
      Top             =   600
      Width           =   2655
   End
   Begin VB.CommandButton CmdChangePass 
      Caption         =   "Change Password"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   1200
      Width           =   2655
   End
   Begin VB.CommandButton CmdOk 
      Caption         =   "Ok"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   1680
      Width           =   2655
   End
   Begin VB.CheckBox Check1 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000A&
      Caption         =   "Always use Password startup."
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
      Height          =   495
      Left            =   120
      MaskColor       =   &H00000000&
      MouseIcon       =   "FrmOptions.frx":0442
      TabIndex        =   0
      Top             =   120
      Width           =   2655
   End
End
Attribute VB_Name = "FrmOptions"
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


Private Sub CmdChangePass_Click()
FrmPassChange.Show
    Me.Hide
End Sub

Private Sub cmdOk_Click()
Dim BoxVal As String

If Check1.Value = 0 Then
 BoxVal = "no"
 SaveSetting "Diary", "Main", "Login", BoxVal
End If
If Check1.Value = 1 Then
 BoxVal = "yes"
 SaveSetting "Diary", "Main", "Login", BoxVal
End If

Unload Me
FrmMain.Show
End Sub

Private Sub Form_Load()
If FrmOptions.Caption <> "Diary 1.4 - Options" Then
FrmOptions.Caption = "Diary 1.4 - Options"
End If
Dim Startup As Integer
    
  Startup = GetSetting("Diary", "Main", "Startup", 0)
    If Startup = 0 Then
     Check2.Value = 0
    ElseIf Startup = 1 Then
     Check2.Value = 1
    End If
 
 
 Dim ChkVal As String
  ChkVal = GetSetting("Diary", "Main", "Login", BoxVal)
    If ChkVal = "no" Then
    Check1.Value = 0
    End If
    If ChkVal = "yes" Then
    Check1.Value = 1
    End If
    
End Sub

Private Sub Check2_Click()
If Check2.Value = 1 Then
On Error Resume Next
SaveSetting "Diary", "Main", "Startup", 1
SaveString HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Run", "Diary", "Diary"
SaveString HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Run", "Diary", App.Path & "\Diary.exe"
ElseIf Check2.Value = 0 Then
SaveSetting "Diary", "Main", "Startup", 0
SaveString HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Run", "Diary", "Diary"
SaveString HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Run", "Diary", "0"
End If
End Sub
