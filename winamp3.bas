Attribute VB_Name = "Winamp3"
Option Explicit
Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hwnd As Long) As Long
Public Const WM_GETTEXT = &HD                   'Getting text of child window
Public Const WM_GETTEXTLENGTH = &HE
Function GetText(iHwnd As Long) As String
    Dim Textlen As Long
    Dim Text As String
    Textlen = SendMessage(iHwnd, WM_GETTEXTLENGTH, 0, 0)
    If Textlen = 0 Then
        GetText = ">No text for this class<"
        Exit Function
    End If
    Textlen = Textlen + 1
    Text = Space(Textlen)
    Textlen = SendMessage(iHwnd, WM_GETTEXT, Textlen, ByVal Text)
    GetText = Left(Text, Textlen)
End Function
Function title(texx As String) As String
On Error Resume Next
Dim ff, i, ed, beg As Integer
ff = 0
For i = 1 To Len(texx)
If Mid(texx, i, 1) = "." Then
beg = i + 1
Exit For
End If
Next i
For i = Len(texx) To beg Step -1
If Mid(texx, i, 1) = "(" Then ff = ff + 1
If (Mid(texx, i, 1) = "(") And (ff = 2) Then
ed = i
title = Mid(texx, beg, ed - beg - 1)
Exit For
End If
Next i
End Function
