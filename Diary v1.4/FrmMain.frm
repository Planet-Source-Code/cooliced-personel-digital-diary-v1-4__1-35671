VERSION 5.00
Object = "{8E27C92E-1264-101C-8A2F-040224009C02}#7.0#0"; "MSCAL.OCX"
Begin VB.Form FrmMain 
   ClientHeight    =   2775
   ClientLeft      =   5430
   ClientTop       =   4065
   ClientWidth     =   4245
   Icon            =   "FrmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2775
   ScaleWidth      =   4245
   StartUpPosition =   2  'CenterScreen
   Begin MSACAL.Calendar Calendar1 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "dd-MM-yyyy"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   2057
         SubFormatType   =   3
      EndProperty
      Height          =   2775
      Left            =   0
      TabIndex        =   0
      ToolTipText     =   "Double click on the desired date for input."
      Top             =   0
      Width           =   4215
      _Version        =   524288
      _ExtentX        =   7435
      _ExtentY        =   4895
      _StockProps     =   1
      BackColor       =   -2147483633
      Year            =   2001
      Month           =   5
      Day             =   8
      DayLength       =   1
      MonthLength     =   2
      DayFontColor    =   0
      FirstDay        =   2
      GridCellEffect  =   1
      GridFontColor   =   10485760
      GridLinesColor  =   -2147483632
      ShowDateSelectors=   -1  'True
      ShowDays        =   -1  'True
      ShowHorizontalGrid=   -1  'True
      ShowTitle       =   -1  'True
      ShowVerticalGrid=   -1  'True
      TitleFontColor  =   10485760
      ValueIsNull     =   -1  'True
      BeginProperty DayFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty GridFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty TitleFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Menu mPopupSys 
      Caption         =   "&SysTray"
      Visible         =   0   'False
      Begin VB.Menu mPopRestore 
         Caption         =   "&Restore"
      End
      Begin VB.Menu mPopTodo 
         Caption         =   "&To-Do"
      End
      Begin VB.Menu mPopAddress 
         Caption         =   "&Address book"
      End
      Begin VB.Menu mPopOptions 
         Caption         =   "&Options"
      End
      Begin VB.Menu mPopAbout 
         Caption         =   "&About"
      End
      Begin VB.Menu mPopExit 
         Caption         =   "&Exit"
      End
   End
End
Attribute VB_Name = "FrmMain"
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

    Private Sub Calendar1_DblClick()
     FrmText.Show
     Me.Hide
    End Sub

    Private Sub Form_Load()
    
    
    If App.PrevInstance Then
        ActivatePrevInstance
    End If

    
    If FrmMain.Caption <> "Diary 1.4" Then
    FrmMain.Caption = "Diary 1.4"
    End If
    
    Calendar1.Value = Date   ' give the Calendar todays date
    DoAboutTxt
     
     'the form must be fully visible before calling Shell_NotifyIcon
       Me.Show
       Me.Refresh
       With nid
        .cbSize = Len(nid)
        .hwnd = Me.hwnd
        .uId = vbNull
        .uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
        .uCallBackMessage = WM_MOUSEMOVE
        .hIcon = Me.Icon
        .szTip = "Diary 1.4" & vbNullChar
       End With
       
       Shell_NotifyIcon NIM_ADD, nid
       
     End Sub

     Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
      'this procedure receives the callbacks from the System Tray icon.
      Dim result As Long
      Dim msg As Long
       'the value of X will vary depending upon the scalemode setting
       If Me.ScaleMode = vbPixels Then
        msg = x
       Else
        msg = x / Screen.TwipsPerPixelX
       End If
       Select Case msg
       
        '****************** un'coment if you want to just click once **********
        '  Case WM_LBUTTONUP        '514 restore form window
        ' Me.WindowState = vbNormal
        ' Result = SetForegroundWindow(Me.hwnd)
        ' Me.Show
        '**********************************************************************
        
        ' dblclick left button
        Case WM_LBUTTONDBLCLK    '515 restore form window
         Me.WindowState = vbNormal
         result = SetForegroundWindow(Me.hwnd)
         Me.Show
        ' right click
        Case WM_RBUTTONUP        '517 display popup menu
         result = SetForegroundWindow(Me.hwnd)
         Me.PopupMenu Me.mPopupSys ' popup menu
         End Select
      End Sub

      Private Sub Form_Resize()
       'this is necessary to assure that the minimized window is hidden
       If Me.WindowState = vbMinimized Then Me.Hide
      End Sub
      Private Sub Form_Unload(Cancel As Integer)
      Kill App.path & "\Diary-Credits.cdc"
       'this removes the icon from the system tray
       Shell_NotifyIcon NIM_DELETE, nid
       Shell_NotifyIcon NIM_DELETE, nid
       End
      End Sub

      Private Sub mPopAbout_Click()
      'called when the user clicks the popup menu about command
       Me.Hide
       FrmAbout.Show
      End Sub

      Private Sub mPopAddress_Click()
      'called when the user clicks the popup menu address command
       Form1.Show
       Me.Hide
      End Sub

      Private Sub mPopExit_Click()
      'called when user clicks the popup menu Exit command
       Unload Me
      End Sub

      Private Sub mPopOptions_Click()
      'called when the user clicks the popup menu Options command
       Me.Hide
       FrmOptions.Show
      End Sub

      Private Sub mPopRestore_Click()

      'called when the user clicks the popup menu Restore command
       Me.WindowState = vbNormal
       result = SetForegroundWindow(Me.hwnd)
       Me.Show

      End Sub
                
      Private Sub mPopTodo_Click()
       Me.Hide
       FrmTodo.Show
      End Sub
