Attribute VB_Name = "General"
'some lines from planet-Source-code
Option Explicit
Public Const HWND_NOTOPMOST = -2&
Public Const THREAD_BASE_PRIORITY_MAX = 2
Public Const HIGH_PRIORITY_CLASS = &H80
Public Type OSVERSIONINFO
    dwOSVersionInfoSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatformId As Long
    szCSDVersion As String * 128
End Type
Public saverpreview As Boolean
Private Type GUID
   Data1 As Long
   Data2 As Integer
   Data3 As Integer
   Data4(7) As Byte
End Type
Public fixx As Boolean
Private Const WS_CHILD = &H40000000
Private Const GWL_HWNDPARENT = (-8)
Private Const GWL_STYLE = (-16)
Private Const HWND_TOPMOST = -1&
Private Const HWND_TOP = 0&
Private Const HWND_BOTTOM = 1&
Private Const SWP_NOSIZE = &H1&
Private Const SWP_NOMOVE = &H2
Private Const SWP_NOZORDER = &H4
Private Const SWP_NOREDRAW = &H8
Private Const SWP_NOACTIVATE = &H10
Private Const SWP_FRAMECHANGED = &H20
Private Const SWP_SHOWWINDOW = &H40
Private Const SWP_HIDEWINDOW = &H80
Private Const SWP_NOCOPYBITS = &H100
Private Const SWP_NOOWNERZORDER = &H200
Private Const SWP_DRAWFRAME = SWP_FRAMECHANGED
Private Const SWP_NOREPOSITION = SWP_NOOWNERZORDER
Private Const RASTERCAPS As Long = 38
Private Const RC_PALETTE As Long = &H100
Private Const SIZEPALETTE As Long = 104

Public Mnew As Boolean
Public myData As TSharedData
Public DisplayHwnd As Long
Public DispRec As RECT
Private Type RECT
   Left As Long
   Top As Long
   Right As Long
   Bottom As Long
End Type
Public Declare Function SetThreadPriority Lib "kernel32" (ByVal hThread As Long, ByVal nPriority As Long) As Long
Public Declare Function SetPriorityClass Lib "kernel32" (ByVal hProcess As Long, ByVal dwPriorityClass As Long) As Long
Public Declare Function GetCurrentThread Lib "kernel32" () As Long
Public Declare Function GetCurrentProcess Lib "kernel32" () As Long
Public Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Long
Public Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Public Declare Sub SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long)
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function GetClientRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Private Declare Function GetDesktopWindow Lib "user32" () As Long
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function ShowCursor& Lib "user32" (ByVal bShow&)

Public Sub Main()
   Dim cmd As String
   Dim Style As Long
   cmd = LCase$(Trim$(Command$))
   Select Case Left$(UCase$(Command$), 2)
   Case "/A"         'change password
  sett.Show
   Case "/C"         'config
  sett.Show
   Case "/P"         'preview
  DisplayHwnd = GetHwndFromCmd(cmd)       ' ** Get HWND of Preview  DeskTop
            GetClientRect DisplayHwnd, DispRec      ' Get Display Rectangle dimentions
            Dim sizee As Byte
            saverpreview = True
            Load saver                          ' Load Screen saver form
            
            Style = GetWindowLong(saver.hwnd, GWL_STYLE) ' ** Get current window style
            Style = Style Or WS_CHILD                        ' ** Append "WS_CHILD" style to the hWnd window style
            SetWindowLong saver.hwnd, GWL_STYLE, Style   ' ** Add new style to window
            sizee = GetSetting("MBM Saver", "settings", "combo3", "18")
            saver.winamp.FontSize = 3
            saver.wintime.FontSize = 3
            saver.Label1(0).FontSize = sizee / 5
            saver.Label1(1).FontSize = sizee / 5
            saver.Label1(2).FontSize = sizee / 5
            saver.Label1(3).FontSize = sizee / 5
            saver.Label1(4).FontSize = sizee / 5
            saver.Label1(5).FontSize = sizee / 5
            saver.Label1(8).FontSize = sizee / 5
            saver.Label1(6).FontSize = sizee / 5
            saver.Label1(7).FontSize = sizee / 5
            saver.Label1(9).FontSize = sizee / 5
            saver.counte.Visible = False
            saver.counter.Visible = False
            saver.coun.Visible = False
            saver.Label7.Visible = False
            
            CursorVisible = True
            
            saver.wintime.Top = 70
            If GetSetting("MBM Saver", "settings", "stan", "True") = "True" Then
            saver.Label1(0).Left = 0
            saver.Label1(1).Left = 0
            saver.Label1(2).Left = 0
            saver.Label1(3).Left = 0
            saver.Label1(4).Left = 0
            saver.Label1(5).Left = 0
            saver.Label1(6).Left = 0
            saver.Label1(8).Left = 0
            saver.Label1(7).Left = 0
            saver.Label1(9).Left = 0
            Dim tp As Integer
                       
            If saver.Label1(0).Visible = True Then
            saver.Label1(0).Top = tp
            tp = tp + sizee * 4.5
            End If
            If saver.Label1(1).Visible = True Then
            saver.Label1(1).Top = tp
            tp = tp + sizee * 4.5
            End If
            If saver.Label1(2).Visible = True Then
            saver.Label1(2).Top = tp
            tp = tp + sizee * 4.5
            End If
            If saver.Label1(3).Visible = True Then
            saver.Label1(3).Top = tp
            tp = tp + sizee * 4.5
            End If
            If saver.Label1(4).Visible = True Then
            saver.Label1(4).Top = tp
            tp = tp + sizee * 4.5
            End If
            If saver.Label1(5).Visible = True Then
            saver.Label1(5).Top = tp
            tp = tp + sizee * 4.5
            End If
            If saver.Label1(8).Visible = True Then
            saver.Label1(8).Top = tp
            tp = tp + sizee * 4.5
            End If
            If saver.Label1(6).Visible = True Then
            saver.Label1(6).Top = tp
            tp = tp + sizee * 4.5
            End If
            If saver.Label1(7).Visible = True Then
            saver.Label1(7).Top = tp
            tp = tp + sizee * 4.5
            End If
            If saver.Label1(9).Visible = True Then
            saver.Label1(9).Top = tp
            tp = tp + sizee * 4.5
            End If
            Else
saver.Label1(0).Left = GetSetting("MBM Saver", "settings", "temp1l", "0") / 1.36
saver.Label1(1).Left = GetSetting("MBM Saver", "settings", "temp2l", "0") / 1.36
saver.Label1(2).Left = GetSetting("MBM Saver", "settings", "temp3l", "0") / 1.36
saver.Label1(3).Left = GetSetting("MBM Saver", "settings", "fan1l", "0") / 1.36
saver.Label1(4).Left = GetSetting("MBM Saver", "settings", "fan2l", "0") / 1.36
saver.Label1(5).Left = GetSetting("MBM Saver", "settings", "fan3l", "0") / 1.36
saver.Label1(6).Left = GetSetting("MBM Saver", "settings", "lcpul", "0") / 1.36
saver.Label1(8).Left = GetSetting("MBM Saver", "settings", "ltimel", "0") / 1.36
saver.Label1(7).Left = GetSetting("MBM Saver", "settings", "memoryl", "0") / 1.36
saver.Label1(9).Left = GetSetting("MBM Saver", "settings", "wintimel", "600") / 1.36
saver.Label1(0).Top = GetSetting("MBM Saver", "settings", "temp1", "0") / 1.38
saver.Label1(1).Top = GetSetting("MBM Saver", "settings", "temp2", "240") / 1.38
saver.Label1(2).Top = GetSetting("MBM Saver", "settings", "temp3", "480") / 1.38
saver.Label1(3).Top = GetSetting("MBM Saver", "settings", "fan1", "720") / 1.38
saver.Label1(4).Top = GetSetting("MBM Saver", "settings", "fan2", "960") / 1.38
saver.Label1(5).Top = GetSetting("MBM Saver", "settings", "fan3", "1200") / 1.38
saver.Label1(6).Top = GetSetting("MBM Saver", "settings", "lcpu", "1440") / 1.38
saver.Label1(8).Top = GetSetting("MBM Saver", "settings", "ltime", "1680") / 1.38
saver.Label1(7).Top = GetSetting("MBM Saver", "settings", "memory", "1920") / 1.38
saver.Label1(9).Top = GetSetting("MBM Saver", "settings", "wintime", "2040") / 1.38
End If
            
            SetParent saver.hwnd, DisplayHwnd   ' ** Set preview window as parent window
            SetWindowLong saver.hwnd, GWL_HWNDPARENT, DisplayHwnd ' ** Save the hWnd Parent in hWnd's window struct.
            
            ' ** Show screensaver in the preview window...
            SetWindowPos saver.hwnd, _
                         HWND_TOP, 0&, 0&, DispRec.Right, DispRec.Bottom, _
                         SWP_NOZORDER Or SWP_NOACTIVATE Or SWP_SHOWWINDOW
   
saver.Image1.Top = 0
saver.Image1.Left = 0
saver.Image1.Picture = saver.Picture1.Picture
saver.Image1.Width = saver.Width
saver.Image1.Height = saver.Height

If GetSetting("MBM Saver", "settings", "stre", "0") = 1 Then
 saver.Image1.Stretch = True
Else
 saver.Image1.Stretch = True
End If

If (GetSetting("MBM Saver", "settings", "option2", "0") = True) And (GetSetting("MBM Saver", "settings", "check7", "0") = 1) Then
 saver.Picture1.Picture = LoadPicture(saver.File1.Path & "\" & saver.File1.List(saver.coun.Caption - 1))
Else
 If (GetSetting("MBM Saver", "settings", "option1", "0") = True) And (GetSetting("MBM Saver", "settings", "check7", "0") = 1) Then saver.Picture1.Picture = LoadPicture(GetSetting("MBM Saver", "settings", "br", ""))
End If
saver.Image1.Picture = saver.Picture1.Picture
   'stretch the picture
   If saver.Image1.Stretch = True Then
    If saver.Width / saver.Picture1.Width > saver.Height / saver.Picture1.Height Then
      saver.Image1.Height = saver.Height
      saver.Image1.Width = saver.Picture1.Width * saver.Height / saver.Picture1.Height
    Else
      saver.Image1.Width = saver.Width
      saver.Image1.Height = saver.Picture1.Height * saver.Width / saver.Picture1.Width
    End If
   End If
   'center the picture
   If GetSetting("MBM Saver", "settings", "cntr", "0") = 1 Then
     If saver.Image1.Width < saver.Width Then saver.Image1.Left = (saver.Width - saver.Image1.Width) / 2
     If saver.Image1.Height < saver.Height Then saver.Image1.Top = (saver.Height - saver.Image1.Height) / 2
   End If
   
   Case "/S"         'display
     saverpreview = False
     Load saver
     saver.Show
       End Select
End Sub '(Public) Sub Main ()

Property Let CursorVisible(ByVal CursorVisible As Boolean)
 ShowCursor CLng(Abs(CursorVisible))
End Property ' Property Let CursorVisible

Private Function GetHwndFromCmd(cmd As String) As Long
'-----------------------------------------------------------------
    Dim Str As String                           ' substring variable
    Dim lenStr As Long                          ' length of substring
    Dim Idx As Long                             ' Index variable
'-----------------------------------------------------------------
    Str = Trim$(cmd)                            ' copy command line
    lenStr = Len(Str)                           ' get size of string
    
    For Idx = lenStr To 1 Step -1               ' for each char in string
        Str = Right$(Str, Idx)                  ' chop off the rightmost char
        If IsNumeric(Str) Then                  ' if substring is numeric then value is an hWnd
            GetHwndFromCmd = Val(Str)           ' return hWnd value
            Exit For                            ' exit for loop
        End If
    Next
'-----------------------------------------------------------------
End Function
Public Function IsWinNT() As Boolean
    Dim OSInfo As OSVERSIONINFO
    OSInfo.dwOSVersionInfoSize = Len(OSInfo)
    'retrieve OS version info
    GetVersionEx OSInfo
    'if we're on NT, return True
    IsWinNT = (OSInfo.dwPlatformId = 2)
End Function
