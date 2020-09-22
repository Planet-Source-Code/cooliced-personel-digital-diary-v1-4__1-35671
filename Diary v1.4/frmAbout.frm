VERSION 5.00
Begin VB.Form FrmAbout 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "About"
   ClientHeight    =   4860
   ClientLeft      =   5550
   ClientTop       =   4245
   ClientWidth     =   4320
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4860
   ScaleWidth      =   4320
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picOut 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4455
      Left            =   0
      ScaleHeight     =   4395
      ScaleWidth      =   4275
      TabIndex        =   1
      Top             =   0
      Width           =   4335
      Begin VB.Timer ReDrawTimer 
         Interval        =   1
         Left            =   0
         Top             =   0
      End
   End
   Begin VB.CommandButton CmdOk 
      Caption         =   "Ok"
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   4560
      Width           =   4335
   End
   Begin VB.PictureBox picBackBuffer 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00C0C0C0&
      FillColor       =   &H00404040&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   4455
      Left            =   0
      ScaleHeight     =   4395
      ScaleWidth      =   4275
      TabIndex        =   3
      Top             =   0
      Visible         =   0   'False
      Width           =   4335
   End
   Begin VB.PictureBox picBuffer 
      BackColor       =   &H00C0C0C0&
      FillColor       =   &H00404040&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   4455
      Left            =   0
      ScaleHeight     =   4395
      ScaleWidth      =   4275
      TabIndex        =   2
      Top             =   0
      Visible         =   0   'False
      Width           =   4335
   End
End
Attribute VB_Name = "FrmAbout"
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
Option Explicit

Private Declare Function BitBlt Lib "gdi32" ( _
   ByVal hdcDest As Long, ByVal XDest As Long, _
   ByVal YDest As Long, ByVal nWidth As Long, _
   ByVal nHeight As Long, ByVal hDCSrc As Long, _
   ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) _
   As Long

Private Const SRCCOPY = &HCC0020

Dim Tempstring(1 To 3000) As Variant
Dim ipicHeight As Integer
Dim ipicWidth As Integer
Dim lYOffset As Integer
Dim iColorCur As Single
Dim iColorStep As Single
Dim NumLines As Integer
Dim lX As Long
Dim lY As Long
Dim strRead As String

Private Sub Form_Load()
If FrmAbout.Caption <> "Diary 1.4 - About" Then
FrmAbout.Caption = "Diary 1.4 - About"
End If
Dim iLine As Integer
    
    NumLines = 1
    
    FrmAbout.ScaleMode = vbPixels
    
    picBuffer.ScaleMode = vbPixels
    
    picBuffer.ForeColor = vbWhite
    picBuffer.BackColor = vbBlack
    picBuffer.AutoRedraw = True
    
    picBuffer.Visible = False
    
    Open (App.path & "\Diary-Credits.cdc") For Input As #1
    
    Do Until EOF(1)
        Line Input #1, Tempstring(NumLines)
        NumLines = NumLines + 1
    Loop
    Close #1
    
    NumLines = NumLines - 1
    
    lX = picBuffer.ScaleLeft
    lY = picBuffer.ScaleHeight
    
    GradiantBackground picBackBuffer
    
    ReDrawTimer.Interval = 5
    ReDrawTimer.Enabled = True

End Sub


Private Function GradiantBackground(picBox As PictureBox)
    ipicWidth = picBox.ScaleWidth
    ipicHeight = picBox.ScaleHeight
    
    iColorCur = 255
    iColorStep = 5 * (0 - 255) / ipicHeight

    For lYOffset = 0 To ipicHeight Step 5
        picBox.Line (-1, lYOffset - 1)-(ipicWidth, lYOffset + 5), RGB(0, 0, iColorCur), BF
        iColorCur = iColorCur + iColorStep
    Next lYOffset
    
End Function

Private Sub RedrawTimer_Timer()
Dim l As Long
Dim j As Long

On Error Resume Next
        
    ' Draw the background to the buffer. It's only had to be written once, so we'll just re-blit it over again and agin.
    l = BitBlt(picBuffer.hDC, 0, picBuffer.ScaleTop, picBuffer.ScaleWidth, picBuffer.ScaleHeight, picBackBuffer.hDC, 0, 0, SRCCOPY)
    
    ' Do the following for each line of text in our credits message...
    For j = 1 To NumLines Step 1
        
        ' Set the starting location of where to print the text. Starts off below the bottom of the buffer.
        picBuffer.CurrentY = lY + (j * picBuffer.FontSize + (6 * j))
        picBuffer.CurrentX = (picBuffer.ScaleWidth / 2) - (picBuffer.TextWidth(Tempstring(j)) / 2)
        
        ' Preset the forground color to white
        picBuffer.ForeColor = vbWhite
       
        ' Once the current line of text reaches this point, begin the color shift. This is done for each line
        ' of text in your message
        If picBuffer.CurrentY < 245 Then
            
            ' If a piece of text is color shifting, but not quite to the top yet...
            If picBuffer.CurrentY > 15 Then
                
                ' This changes the forground color to a shade of whatever color (in this case...blue). As
                ' it nears the top, the rate of the R,G anf B values shift differently, to allow a gradual
                ' color shift.
                picBuffer.ForeColor = RGB((((255 / 235) * picBuffer.CurrentY)), (((255 / 235) * picBuffer.CurrentY)), (((255 / 25) * picBuffer.CurrentY)))
            Else
                
                ' We've reached the top...just paint it black and get it over with....
                picBuffer.ForeColor = vbBlack
                
                If j = NumLines And picBuffer.CurrentY < -25 Then
                ' If we've painted the last line, and it's above the top, there's no more text to scroll
                ' and we exit.
                    ReDrawTimer.Enabled = False
                    FrmMain.Show
                    Unload Me
                End If
            End If
        End If
        
        ' Send the text directly into the buffer hDC
        picBuffer.Print Tempstring(j)
        
    Next
    
    ' Ok, now that we have painted the entire buffer as we see fit for this pass, we blast the entire
    ' finished image directly to our output picturebox control.
    l = BitBlt(picOut.hDC, 0, picOut.ScaleTop, picOut.ScaleWidth, picOut.ScaleHeight, picBuffer.hDC, 0, 0, SRCCOPY)
    
    picOut.Refresh
    
    ' Change the offset for the location of where the text will display next turn
    lY = lY - 1

End Sub

Private Sub cmdOk_Click()
 ReDrawTimer.Enabled = False
 FrmMain.Show
 Unload Me
End Sub
