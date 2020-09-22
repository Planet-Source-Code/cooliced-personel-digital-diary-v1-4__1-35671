VERSION 5.00
Begin VB.Form frmSplash 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   ClientHeight    =   2775
   ClientLeft      =   210
   ClientTop       =   1365
   ClientWidth     =   4860
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2775
   ScaleWidth      =   4860
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CmdOk 
      Caption         =   "Start"
      Enabled         =   0   'False
      Height          =   255
      Left            =   3000
      TabIndex        =   1
      Top             =   2400
      Width           =   1695
   End
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      BackColor       =   &H80000007&
      Height          =   2760
      Left            =   0
      Picture         =   "FrmSplash.frx":0000
      ScaleHeight     =   2700
      ScaleWidth      =   4785
      TabIndex        =   0
      Top             =   0
      Width           =   4845
      Begin VB.Timer Timer1 
         Interval        =   5000
         Left            =   2520
         Top             =   600
      End
   End
End
Attribute VB_Name = "frmSplash"
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

Private Sub cmdOk_Click()
Startup 'Call startup sub
End Sub

Private Sub Form_Load()
   
    If App.PrevInstance Then
        ActivatePrevInstance
    End If


 If SoftICELoaded Then ' check if softice is loaded
  MsgBox "SoftICE is detected! Closing now!", vbMsgBoxSetForeground + vbInformation, "Diary 1.4"
  End ' if true finish the app
 End If
End Sub

Private Sub Timer1_Timer()
CmdOk.Enabled = True
CmdOk.SetFocus
End Sub
