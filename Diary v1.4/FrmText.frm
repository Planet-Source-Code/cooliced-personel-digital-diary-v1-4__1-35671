VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form FrmText 
   ClientHeight    =   6150
   ClientLeft      =   5430
   ClientTop       =   3045
   ClientWidth     =   9255
   Icon            =   "FrmText.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   ScaleHeight     =   6150
   ScaleWidth      =   9255
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog CfontDialog 
      Left            =   840
      Top             =   3720
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin RichTextLib.RichTextBox RTB1 
      Height          =   5775
      Left            =   0
      TabIndex        =   0
      Top             =   360
      Width           =   9255
      _ExtentX        =   16325
      _ExtentY        =   10186
      _Version        =   393217
      BorderStyle     =   0
      ScrollBars      =   2
      TextRTF         =   $"FrmText.frx":0442
   End
   Begin MSComctlLib.ImageList imgToolBar 
      Left            =   3600
      Top             =   1320
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   16
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmText.frx":04F0
            Key             =   "New"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmText.frx":0602
            Key             =   "Open"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmText.frx":0714
            Key             =   "Save"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmText.frx":0826
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmText.frx":0938
            Key             =   "Cut"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmText.frx":0A4A
            Key             =   "Copy"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmText.frx":0B5C
            Key             =   "Paste"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmText.frx":0C6E
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmText.frx":0FC2
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmText.frx":1316
            Key             =   "Bold"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmText.frx":1428
            Key             =   "Italic"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmText.frx":153A
            Key             =   "Underline"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmText.frx":164C
            Key             =   "Align Left"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmText.frx":175E
            Key             =   "Center"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmText.frx":1870
            Key             =   "Align Right"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmText.frx":1982
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tbToolBar 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   9255
      _ExtentX        =   16325
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      Appearance      =   1
      Style           =   1
      ImageList       =   "imgToolBar"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   22
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "New"
            Object.ToolTipText     =   "Import"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Open"
            Object.ToolTipText     =   "Export"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Save"
            Object.ToolTipText     =   "Save"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Print"
            Object.ToolTipText     =   "Print"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Cut"
            Object.ToolTipText     =   "Cut"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Copy"
            Object.ToolTipText     =   "Copy"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Paste"
            Object.ToolTipText     =   "Paste"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Undo"
            Object.ToolTipText     =   "Undo"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Redo"
            Object.ToolTipText     =   "Redo"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Bold"
            Object.ToolTipText     =   "Bold"
            ImageIndex      =   10
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Italic"
            Object.ToolTipText     =   "Italic"
            ImageIndex      =   11
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Underline"
            Object.ToolTipText     =   "Underline"
            ImageIndex      =   12
         EndProperty
         BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button18 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Align Left"
            Object.ToolTipText     =   "Align Left"
            ImageIndex      =   13
            Style           =   2
         EndProperty
         BeginProperty Button19 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Center"
            Object.ToolTipText     =   "Center"
            ImageIndex      =   14
            Style           =   2
         EndProperty
         BeginProperty Button20 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Align Right"
            Object.ToolTipText     =   "Align Right"
            ImageIndex      =   15
            Style           =   2
         EndProperty
         BeginProperty Button21 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button22 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Color"
            Object.ToolTipText     =   "Color"
            ImageIndex      =   16
         EndProperty
      EndProperty
      Begin VB.ComboBox size 
         Height          =   315
         ItemData        =   "FrmText.frx":1D86
         Left            =   8640
         List            =   "FrmText.frx":1DBD
         TabIndex        =   3
         Text            =   "10"
         Top             =   0
         Width           =   615
      End
      Begin VB.ComboBox text 
         Height          =   315
         ItemData        =   "FrmText.frx":1E04
         Left            =   6240
         List            =   "FrmText.frx":1E35
         TabIndex        =   2
         Text            =   "Times New Roman"
         Top             =   0
         Width           =   2295
      End
   End
   Begin VB.Menu MnuFile 
      Caption         =   "&File"
      Begin VB.Menu DoSave 
         Caption         =   "&Save"
      End
      Begin VB.Menu MnuDel 
         Caption         =   "Delete File"
      End
      Begin VB.Menu line5 
         Caption         =   "-"
      End
      Begin VB.Menu MnuImport 
         Caption         =   "Import"
      End
      Begin VB.Menu MnuExport 
         Caption         =   "Export"
      End
      Begin VB.Menu line12 
         Caption         =   "-"
      End
      Begin VB.Menu MnuPrint 
         Caption         =   "Print"
      End
      Begin VB.Menu line3 
         Caption         =   "-"
      End
      Begin VB.Menu DoClose 
         Caption         =   "C&lose"
      End
   End
   Begin VB.Menu MnuEdit 
      Caption         =   "&Edit"
      Begin VB.Menu MnuUndo 
         Caption         =   "Undo"
      End
      Begin VB.Menu MnuRedo 
         Caption         =   "Redo"
      End
      Begin VB.Menu Line1 
         Caption         =   "-"
      End
      Begin VB.Menu MnuCut 
         Caption         =   "Cut"
      End
      Begin VB.Menu MnuCopy 
         Caption         =   "Copy"
      End
      Begin VB.Menu MnuPaste 
         Caption         =   "Paste"
      End
      Begin VB.Menu line2 
         Caption         =   "-"
      End
      Begin VB.Menu MnuSelect 
         Caption         =   "Select all"
      End
   End
End
Attribute VB_Name = "FrmText"
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
    

'These are the variables
Dim gblnIgnoreChange As Boolean
Dim gintIndex As Integer
Dim gstrStack(1000) As String
Dim path As String
Dim FileName As String
Dim extention As String
Dim s(0 To 255) As Integer 'S-Box
Dim kep(0 To 255) As Integer
Dim i As Integer, j As Integer

Private Sub DoClose_Click()
    FrmMain.Show
    Unload Me
End Sub

Private Sub DoSave_Click()

 path = App.path & "\Data\"
 extention = ".ccd"
 FileName = Format(FrmMain.Calendar1.Value, "Medium Date")

    If RTB1.text = "" Then
     DoSave.Checked = False
     Exit Sub
    Else
     cmdRC4Decrypt
     RTB1.SaveFile path & FileName & extention 'save the file
     cmdRC4Encrypt
    End If
End Sub

Private Sub Form_Load()

 FrmMain.Visible = False
 Me.Visible = True
 Me.SetFocus
 
 path = App.path & "\Data\" 'set the path
 extention = ".ccd"
 FileName = Format(FrmMain.Calendar1.Value, "Medium Date") ' to match the date
 
 'Set the forms date to the diary selected date
 Me.Caption = Format(FrmMain.Calendar1.Value, "Long Date")
 
 If FileExists(path & FileName & extention) = False Then
  MsgBox "This date is empty at the moment", vbExclamation Or vbOKOnly, "Diary"
  RTB1.text = ""
 Else  'if does exist
  cmdRC4Decrypt
  RTB1.LoadFile path & FileName & extention                      'load file
  cmdRC4Encrypt
 End If

Dim i As Integer
   With text
      For i = 0 To Screen.FontCount - 1
         .AddItem Screen.Fonts(i)
      Next i
      ' Set ListIndex to 0.
      .ListIndex = 0
   End With

   With size
      ' Populate the combo with sizes in
      ' increments of 2.
      For i = 8 To 72 Step 2
         .AddItem i
      Next i
      ' Set ListIndex to 0
      .ListIndex = 1 ' size 10.
   End With

End Sub

Private Sub Form_Resize()
RTB1.Width = Me.ScaleWidth
RTB1.Height = Me.ScaleHeight
End Sub

Private Sub Form_Unload(Cancel As Integer)
 FrmMain.Show
 FrmMain.Visible = True
 Unload Me
 End Sub

Private Sub fs1_Click()

End Sub

Private Sub MnuCopy_Click()
 'Clears the Clipboard to put text on it
 Clipboard.Clear
 'Sets the Text from rtb1 onto the Clipboard
 Clipboard.SetText RTB1.SelText
 'Sets the Focus to rtb1
 RTB1.SetFocus
End Sub

Private Sub MnuCut_Click()
 'Clears the Clipboard to put text on it
 Clipboard.Clear
 'Sets the Text from rtb1 onto the Clipboard
 Clipboard.SetText RTB1.SelText
 'Deletes the Selected Text on rtb1
 RTB1.SelText = ""
 'Sets the Focus to rtb1
 RTB1.SetFocus
End Sub

Private Sub MnuDel_Click()
Dim msg, Response
msg = "Do you want to delete this entry?"
path = App.path & "\Data\" 'set the path
 extention = ".ccd"
 FileName = Format(FrmMain.Calendar1.Value, "Medium Date") ' to match the date
 
 If FileExists(path & FileName & extention) = True Then
 Response = MsgBox(msg, vbYesNo + vbExclamation, "Diary")

   Select Case Response ' select a response
    Case vbYes     ' User chose the Yes button .
     Kill path & FileName & extention
     RTB1.text = ""
     RTB2.text = ""
     Unload Me
     Exit Sub
    Case vbNo ' the user chose the No Button
     Exit Sub
    End Select
 Else
    MsgBox "Entry does not exist yet!", vbExclamation Or vbOKOnly, "Diary"
    Exit Sub
 End If

End Sub

Private Sub MnuExport_Click()
On Error GoTo skip
Dim FileName As String
CfontDialog.Filter = "Text Files (*.txt) |*.txt"
CfontDialog.Action = 2
FileName = CfontDialog.FileName
F = FreeFile
Open FileName For Output As #F
Print #F, RTB1.text
Close #F
skip:
End Sub

Private Sub MnuImport_Click()
    ' Set CancelError is True
    CfontDialog.Cancelerror = True
    On Error GoTo ErrHandler
    ' Set flags
    CfontDialog.Flags = cdlOFNHideReadOnly
    ' Set filters
    CfontDialog.Filter = "Text Files(*.txt;*.rtf)|*.txt;*.rtf"
    ' Specify default filter
    CfontDialog.FilterIndex = 2
    ' Display the Open dialog box
    CfontDialog.ShowOpen
    ' Display name of selected file

     RTB1.LoadFile CfontDialog.FileName
    Exit Sub
    
ErrHandler:
    'User pressed the Cancel button
    Exit Sub

End Sub

Private Sub MnuPaste_Click()
    'Puts the Text from the clipboard into rtb1
    RTB1.SelText = Clipboard.GetText
    'Sets the Focus to rtb1
    RTB1.SetFocus
End Sub

Private Sub mnuprint_Click()

CfontDialog.Flags = cdlPDReturnDC + cdlPDNoPageNums
   If RTB1.SelLength = 0 Then
      CfontDialog.Flags = CfontDialog.Flags + cdlPDAllPages
   Else
      CfontDialog.Flags = CfontDialog.Flags + cdlPDSelection
   End If
   CfontDialog.ShowPrinter
   RTB1.SelPrint CfontDialog.hDC



End Sub

Private Sub MnuRedo_Click()
     'This is the basic redo stuff.
    gblnIgnoreChange = True
    gintIndex = gintIndex + 1
    On Error Resume Next
    RTB1.TextRTF = gstrStack(gintIndex)
    gblnIgnoreChange = False
End Sub

Private Sub MnuSelect_Click()
    'Sets the cursors position to zero
    RTB1.SelStart = 0
    'Selects the full length of rtb1
    RTB1.SelLength = Len(RTB1.text)
    'Sets the Focus to rtb1
    RTB1.SetFocus
End Sub

Private Sub MnuUndo_Click()
      'This says that if the Index is = to 0, then It shouldn't undo anymore
    If gintIndex = 0 Then
    Exit Sub
    End If
    
    'This is the basic undo stuff.
    gblnIgnoreChange = True
    gintIndex = gintIndex - 1
    On Error Resume Next
    RTB1.TextRTF = gstrStack(gintIndex)
    gblnIgnoreChange = False
End Sub

Private Sub RTB1_Change()
    'Basically this updates the Undo and Redo
    If Not gblnIgnoreChange Then
        gintIndex = gintIndex + 1
        gstrStack(gintIndex) = RTB1.TextRTF
    End If
End Sub

'******************************************************************************
'##############################################################################
'                      RC4 Stuff
Public Sub RC4ini(Pwd As String)
    Dim temp As Integer, a As Integer, b As Integer
    'Save Passkey in Byte-Array
    b = 0


    For a = 0 To 255
        b = b + 1


        If b > Len(Pwd) Then
            b = 1
        End If
        kep(a) = Asc(Mid$(Pwd, b, 1))
    Next a
    'INI S-Box


    For a = 0 To 255
        s(a) = a
    Next a
    b = 0


    For a = 0 To 255
        b = (b + s(a) + kep(a)) Mod 256
        ' Swap( S(i),S(j) )
        temp = s(a)
        s(a) = s(b)
        s(b) = temp
    Next a
End Sub

'Only use this routine for short texts
Public Function EnDeCrypt(plaintxt As Variant) As Variant
    Dim temp As Integer, a As Long, i As Integer, j As Integer, k As Integer
    Dim cipherby As Byte, cipher As Variant


    For a = 1 To Len(plaintxt)
        i = (i + 1) Mod 256
        j = (j + s(i)) Mod 256
        ' Swap( S(i),S(j) )
        temp = s(i)
        s(i) = s(j)
        s(j) = temp
        'Generate Keybyte k
        k = s((s(i) + s(j)) Mod 256)
        'Plaintextbyte xor Keybyte
        cipherby = Asc(Mid$(plaintxt, a, 1)) Xor k
        cipher = cipher & Chr(cipherby)
    Next a
    EnDeCrypt = cipher
End Function

'Use this routine for really huge files
Public Function EnDeCryptSingle(plainbyte As Byte) As Byte
    Dim temp As Integer, k As Integer
    Dim cipherby As Byte
    i = (i + 1) Mod 256
    j = (j + s(i)) Mod 256
    ' Swap( S(i),S(j) )
    temp = s(i)
    s(i) = s(j)
    s(j) = temp
    'Generate Keybyte k
    k = s((s(i) + s(j)) Mod 256)
    'Plaintextbyte xor Keybyte
    cipherby = plainbyte Xor k
    EnDeCryptSingle = cipherby
End Function

Private Sub cmdRC4Encrypt()
Dim passval As String
Dim Pass As String
Dim DF As String
    Dim inbyte As Byte
    Dim z As Long
    
    path = App.path & "\Data\" 'set the path
    extention = ".ccd"
    FileName = Format(FrmMain.Calendar1.Value, "Medium Date") ' to match the date
    'Set the Set-Box Counter zero
    i = 0: j = 0
    'Ini the S-Boxes only once for a hole file
    
    Pass = GetSetting("Diary", "Main", "Password", passval)
     RC4ini (Pass)

    'Disable the Mousepointer
    If FileExists(path & FileName & extention) = False Then
       Exit Sub
    End If
    DF = path & FileName & extention
    Open DF For Binary As 1
    Open DF For Binary As 2


    For z = 1 To LOF(1)
        Get #1, , inbyte
        Put #2, , EnDeCryptSingle(inbyte)
    Next z
    Close 1
    Close 2
    'Enable the Mousepointer
    MousePointer = vbDefault
End Sub

Private Sub cmdRC4Decrypt()
Dim passval As String
Dim Pass As String
Dim EF As String
    Dim inbyte As Byte
    Dim z As Long
    
    path = App.path & "\Data\" 'set the path
    extention = ".ccd"
    FileName = Format(FrmMain.Calendar1.Value, "Medium Date") ' to match the date
    'Set the Set-Box counter zero
    i = 0: j = 0
    'Ini the S-Boxes only once for a hole fi
    '     le

    Pass = GetSetting("Diary", "Main", "Password", passval)
    
    RC4ini (Pass)
    
    'Disable the Mousepointer
    MousePointer = vbHourglass
    If FileExists(path & FileName & extention) = False Then
        Exit Sub
    End If
    EF = path & FileName & extention
    Open EF For Binary As 1
    
    Open EF For Binary As 2


    For z = 1 To LOF(1)
        Get #1, , inbyte
        Put #2, , EnDeCryptSingle(inbyte)
    Next
    Close 1
    Close 2
    'Enable the Mousepointer
    MousePointer = vbDefault
End Sub

Private Sub size_click()
RTB1.SelFontSize = size.text
   RTB1.SetFocus
End Sub

Private Sub tbToolBar_ButtonClick(ByVal Button As MSComctlLib.Button)
On Error Resume Next
        ' Handle toolbar clicks
        Select Case Button.KEY
                Case "Undo"
                    MnuUndo_Click
                Case "Redo"
                MnuRedo_Click
                Case "Color"
                    CfontDialog.ShowColor
                    RTB1.SelColor = CfontDialog.Color
                    RTB1.SetFocus
                Case "New"
                    MnuImport_Click
                Case "Open"
                    MnuExport_Click
                Case "Save"
                    DoSave_Click
                Case "Print"
                    mnuprint_Click
                Case "Cut"
                    MnuCut_Click
                Case "Copy"
                    MnuCopy_Click
                Case "Paste"
                    MnuPaste_Click
                Case "Bold"
                    If RTB1.SelBold = False Then
                    RTB1.SelBold = True
                    Else
                    RTB1.SelBold = False
                    End If
                Case "Italic"
                    If RTB1.SelItalic = False Then
                    RTB1.SelItalic = True
                    Else
                    RTB1.SelItalic = False
                    End If
                Case "Underline"
                    If RTB1.SelUnderline = False Then
                    RTB1.SelUnderline = True
                    Else
                    RTB1.SelUnderline = False
                    End If
                Case "Align Left"
                    If RTB1.SelLength > 0 Then
                    RTB1.SelAlignment = 0
                    End If
                Case "Center"
                    If RTB1.SelLength > 0 Then
                    RTB1.SelAlignment = 2
                    End If
                Case "Align Right"
                    If RTB1.SelLength > 0 Then
                    RTB1.SelAlignment = 1
                    End If
        End Select
End Sub


Private Sub text_Click()
  RTB1.SelFontName = text
   RTB1.SetFocus
End Sub
