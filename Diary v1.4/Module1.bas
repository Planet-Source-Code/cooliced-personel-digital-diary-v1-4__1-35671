Attribute VB_Name = "mod1"
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

Dim path As String
Dim FileName As String
Dim extention As String
Dim MyData As Database
Dim MyRecord As Recordset
Dim SQL As String

'user defined type required by Shell_NotifyIcon API call
      Public Type NOTIFYICONDATA
       cbSize As Long
       hwnd As Long
       uId As Long
       uFlags As Long
       uCallBackMessage As Long
       hIcon As Long
       szTip As String * 64
      End Type
     
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Const WM_PASTE = &H302

Private Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long         ' Note that if you declare the lpData parameter as String, you must pass it By Value.
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Private Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long

Private Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long

     'constants required by Shell_NotifyIcon API call:
      Public Const NIM_ADD = &H0
      Public Const NIM_MODIFY = &H1
      Public Const NIM_DELETE = &H2
      Public Const NIF_MESSAGE = &H1
      Public Const NIF_ICON = &H2
      Public Const NIF_TIP = &H4
      Public Const WM_MOUSEMOVE = &H200
      Public Const WM_LBUTTONDOWN = &H201     'Button down
      Public Const WM_LBUTTONUP = &H202       'Button up
      Public Const WM_LBUTTONDBLCLK = &H203   'Double-click
      Public Const WM_RBUTTONDOWN = &H204     'Button down
      Public Const WM_RBUTTONUP = &H205       'Button up
      Public Const WM_RBUTTONDBLCLK = &H206   'Double-click

      Public Declare Function SetForegroundWindow Lib "user32" _
      (ByVal hwnd As Long) As Long
      Public Declare Function Shell_NotifyIcon Lib "shell32" _
      Alias "Shell_NotifyIconA" _
      (ByVal dwMessage As Long, pnid As NOTIFYICONDATA) As Boolean

      Public nid As NOTIFYICONDATA
      Private Declare Function ShellExecute Lib "shell32.dll" _
      Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, _
      ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, _
      ByVal nShowCmd As Long) As Long
      
      Public Declare Function CreateFileNS Lib "kernel32" Alias "CreateFileA" (ByVal lpFileName As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, ByVal lpSecurityAttributes As Long, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
Public Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Public Declare Function WriteFileNO Lib "kernel32" Alias "WriteFile" (ByVal hfile As Long, lpBuffer As Any, ByVal nNumberOfBytesToWrite As Long, lpNumberOfBytesWritten As Long, ByVal lpOverlapped As Long) As Long

Public Const GENERIC_READ = &H80000000
Public Const GENERIC_WRITE = &H40000000
Public Const FILE_SHARE_READ = &H1
Public Const FILE_SHARE_WRITE = &H2
Public Const OPEN_EXISTING = 3
Public Const FILE_ATTRIBUTE_NORMAL = &H80
Private Const SW_MAXIMIZE = 3

Function FileExists(FileNa As String) As Boolean
    Dim FRes As String
    On Error GoTo NotFound
    FRes = Dir$(FileNa)
    If FRes = "" Then FileExists = False Else FileExists = True
NotFound:
    If Err = 53 Then Resume Next
End Function

Public Function GetSettings()
Dim Fol As String
On Error GoTo GetSettingsError
  '  make folder for diary files
   Fol = App.path & "\Data"
    MkDir Fol
    
GetSettingsError:
End Function

Public Sub Startup()
Dim ChkVal As String
Dim RunVal As Integer
    GetSettings
    Unload frmSplash
    
    RunVal = GetSetting("Diary", "Main", "FirstRun", 0)
    
    If RunVal = 0 Then             ' if first run
        FrmFirstRun.Show           ' show firstrun
    End If
    
    If RunVal = 1 Then                   'if not first run
    
    ChkVal = GetSetting("Diary", "Main", "Login", BoxVal)
    
    If ChkVal = "no" Then     ' if ChkVal = no then
     FrmMain.Show ' show the main form
    ElseIf ChkVal = "yes" Then    'if ChkVal = yes then
     FrmLogin.Show           'show login form
    End If
    
   End If
End Sub
' About form text, to stop people from changing it :P
Public Sub DoAboutTxt()
Open App.path & "\Diary-Credits.cdc" For Output As #1
Print #1, "Diary 1.4"
Print #1,
Print #1, "CODED BY"
Print #1, "Cooliced"
Print #1,
Print #1, "Licensed to:"
Print #1, UserName
Print #1,
Print #1,
Print #1, "Copyright © Cooliced"
Print #1, "Cooliced.co.uk 1999-2002"
Print #1,
Close #1
End Sub

' i know what you crackers are like!!!
' i know i am one!
' Secure Reasoning
Public Function SoftICELoaded() As Boolean
Dim hfile As Long, RetVal As Long
    hfile = CreateFileNS("\\.\SICE", GENERIC_WRITE Or GENERIC_READ, FILE_SHARE_READ Or FILE_SHARE_WRITE, 0, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, 0)
    If hfile <> -1 Then
        ' SoftICE is detected.
        RetVal = CloseHandle(hfile) ' Close the file handle
        SoftICELoaded = True
    Else
    ' SoftICE is not found.
    SoftICELoaded = False
    End If
End Function

Public Sub SaveString(hKey As Long, StrPath As String, StrValue As String, StrData As String)
   Dim KeyH&
    r = RegCreateKey(hKey, StrPath, KeyH&)
    r = RegSetValueEx(KeyH&, StrValue, 0, 1, ByVal StrData, Len(StrData))
    r = RegCloseKey(KeyH&)
End Sub

Public Function UserName() As String
    Dim lpBuffer As String
    Dim j
    lpBuffer = Space$(255)
    GetUserName lpBuffer, Len(lpBuffer)
        j = InStr(lpBuffer, Chr$(0))
    If j > 0 Then UserName = Left$(lpBuffer, j - 1)
End Function
