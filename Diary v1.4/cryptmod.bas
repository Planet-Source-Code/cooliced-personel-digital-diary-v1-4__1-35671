Attribute VB_Name = "cryptmod"
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
    
Public Function FindOppAsc(Value As Integer) As Integer
    If Value <> 128 Then
        FindOppAsc = 255 - Value
    Else
        FindOppAsc = Value
    End If
End Function

Public Function Converter(xString As String) As String
    On Error GoTo FinaliseError
    For cCode = 1 To Len(xString)
        conv = conv + (100 / Len(xString))
        Converter = Converter + Chr(FindOppAsc(Asc(Mid(xString, CInt(cCode), 1))))
    Next cCode
    
    Exit Function
FinaliseError:
    MsgBox "Error, the string that was meant be be coded / decoded was too long.", vbCritical, "Error"
End Function

Public Function Encrypt(Text As String, Password As String) As String
On Error GoTo errorhandler
GoSub begin

errorhandler:
Exit Function

begin:
    Dim oTest As AES, sTemp As String, bytIn() As Byte
    Dim bytOut() As Byte, bytPassword() As Byte, bytClear() As Byte
    Dim lCount As Long, lLength As Long
    
    If Text = "" Or Password = "" Then Exit Function
    
    Set oTest = New AES
    
    bytIn = Text
    bytPassword = Password

    bytOut = oTest.EncryptData(bytIn, bytPassword)

    sTemp = ""
    For lCount = 0 To UBound(bytOut)
        sTemp = sTemp & Right("0" & Hex(bytOut(lCount)), 2)
    Next
    Encrypt = sTemp
End Function

Public Function Decrypt(EncryptedString As String, Password As String) As String
On Error GoTo errorhandler
GoSub begin

errorhandler:
Exit Function

begin:
    If EncryptedString = "" Or Password = "" Then Exit Function
    
    Dim oTest As AES, sTemp As String, bytIn() As Byte
    Dim bytOut() As Byte, bytPassword() As Byte, bytClear() As Byte
    Dim lCount As Long, lLength As Long, DC As String
    
    Set oTest = New AES
    
    bytIn = EncryptedString
    bytPassword = Password
    sTemp = EncryptedString
    
    lLength = Len(sTemp)
    ReDim bytOut((lLength \ 2) - 1)
    For lCount = 1 To lLength Step 2
        bytOut(lCount \ 2) = CByte("&H" & Mid(sTemp, lCount, 2))
    Next

    bytClear = oTest.DecryptData(bytOut, bytPassword)
    DC = bytClear
    If DC = vbNullString Then
        MsgBox "Invalid password", vbOKOnly, "Error"
    Else
        Decrypt = bytClear
    End If
End Function

