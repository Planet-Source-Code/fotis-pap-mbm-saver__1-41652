Attribute VB_Name = "WinampModule"
'From planet-Source-code
'with some changes
Option Explicit
Public Const WM_COMMAND = 273
Public Const WM_USER = 1024
Public Const WM_WA_IPC = &H400
Public Const WM_COPYDATA = &H4A
Public WinampID As Long
Public WinampPath As String
Public LastWinampCaption As String
Public LastTitle As String
Declare Function FindWindowA Lib "user32" (ByVal lpClassName As String, ByVal lpWindowName As Long) As Long
Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal WndID As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Public Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hwnd As Long) As Long

Type COPYDATASTRUCT
    dwData As Long
    cbData As Long
    lpData As Long
End Type
'=============================

Public Const SONG_LENGTH As Byte = 3            'Returns the length of the song in seconds
Public Const SONG_POSITION As Byte = 4          'Returns the current position in the song, in milliseconds
Public Const SONG_TITLE As Byte = 13            'Returns the title of the song.


Public Function FindWinamp() As Boolean

On Error GoTo err
    
    WinampID = FindWindowA("Winamp v1.x", 0)
    If WinampID = 0 Then
    FindWinamp = False
    Else
    FindWinamp = True
    End If
    
err:
       
End Function




Public Function WM_GET(cmnd As Byte, Optional data As Long) As Variant
    On Error GoTo err
    Dim tmp As String
    Dim Ret As Integer
    Dim isplay As Integer
    isplay = SendMessage(WinampID, WM_USER, 0, 104)
    Select Case cmnd
        Case SONG_LENGTH
            WM_GET = SendMessage(WinampID, WM_WA_IPC, 1, 105) ' returns track length in seconds
        Case SONG_POSITION
            WM_GET = SendMessage(WinampID, WM_WA_IPC, 0, 105) ' returns position in the current track in milliseconds
        Case SONG_TITLE ' returns a string with the song title in it
            Dim strBuffer As String, lngtextlen As Long
            Let lngtextlen& = GetWindowTextLength(WinampID) 'gets the length of the caption
            Let strBuffer$ = String$(lngtextlen&, 0&) 'i dont know why this is necessary, i found it in someone else's API code
            Call GetWindowText(WinampID, strBuffer$, lngtextlen& + 1&) ' reads in the caption text
            If strBuffer$ = LastWinampCaption Then
                WM_GET = LastTitle
            Else
                LastWinampCaption = strBuffer
                If LCase(strBuffer$) Like "*[paused]" = False Then ' queries if the [Paused] string is there and removes it
                    strBuffer$ = Left(strBuffer$, Len(strBuffer) - 9)
                End If
                strBuffer$ = Mid(strBuffer$, 1, Len(strBuffer) - 8) ' removes the -Winamp
                Dim findDot As Integer
                findDot = InStr(1, strBuffer, ".") ' finds the dot in the number at the beginning
                LastTitle = Trim(Mid(strBuffer$, findDot + 1)) 'Returns the final title value
                WM_GET = LastTitle
            End If
    End Select
    Exit Function
err:

End Function
