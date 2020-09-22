VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form FrmTodo 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "To-Do"
   ClientHeight    =   1560
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   1935
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1560
   ScaleWidth      =   1935
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin RichTextLib.RichTextBox RTB1 
      Height          =   1575
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   2778
      _Version        =   393217
      Enabled         =   -1  'True
      ScrollBars      =   2
      MaxLength       =   500
      TextRTF         =   $"FrmTodo.frx":0000
   End
   Begin RichTextLib.RichTextBox RTB2 
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   960
      Visible         =   0   'False
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   661
      _Version        =   393217
      BorderStyle     =   0
      Appearance      =   0
      TextRTF         =   $"FrmTodo.frx":006E
   End
End
Attribute VB_Name = "FrmTodo"
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
    

Private Sub Form_Load()
Dim FontVal As String
Dim passval As String
Dim Pass As String
Dim FVal As String

 Path = App.Path & "\Data\" 'set the path
 extention = ".ccd"
 FileName = "Todo" ' to match the date
 Pass = GetSetting("Diary", "Main", "Password", passval)
 
 If FileExists(Path & FileName & extention) = False Then
    RTB1.Text = ""
 Else  'if does exist
  RTB2.LoadFile Path & FileName & extention                      'load file
  RTB1.Text = Decrypt(RTB2.Text, Pass)
  RTB2.Text = ""                                                 'clear text box
  FVal = GetSetting("Diary", "Main", "Text", FontVal)
  RTB1.Font = FVal          'set font
 End If

End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim passval As String
Dim Pass As String
Dim IniFont As String

 Path = App.Path & "\Data\"
 extention = ".ccd"
 FileName = "Todo"
 
 Pass = GetSetting("Diary", "Main", "Password", passval)
    
    If RTB1.Text = "" Then
     Exit Sub
    Else
     IniFont = RTB1.SelFontName
     SaveSetting "Diary", "Main", "Text", IniFont

     RTB2.Text = Encrypt(RTB1.Text, Pass)
     RTB2.SaveFile Path & FileName & extention 'save the file
     RTB2.Text = "" 'after crypt and save clear the crypt box
    End If

FrmMain.Show

End Sub
