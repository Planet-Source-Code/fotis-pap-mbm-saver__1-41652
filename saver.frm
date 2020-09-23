VERSION 5.00
Begin VB.Form saver 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   Caption         =   "saver"
   ClientHeight    =   11520
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15360
   ClipControls    =   0   'False
   Icon            =   "saver.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   11520
   ScaleWidth      =   15360
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.PictureBox picgraph 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      ForeColor       =   &H0000FF00&
      Height          =   1150
      Index           =   6
      Left            =   0
      ScaleHeight     =   122.951
      ScaleMode       =   0  'User
      ScaleWidth      =   215.612
      TabIndex        =   22
      Top             =   7560
      Visible         =   0   'False
      Width           =   7695
   End
   Begin VB.PictureBox picgraph 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      ForeColor       =   &H0000FF00&
      Height          =   1150
      Index           =   5
      Left            =   0
      ScaleHeight     =   122.951
      ScaleMode       =   0  'User
      ScaleWidth      =   215.612
      TabIndex        =   21
      Top             =   6120
      Visible         =   0   'False
      Width           =   7695
   End
   Begin VB.PictureBox picgraph 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      ForeColor       =   &H0000FF00&
      Height          =   1150
      Index           =   4
      Left            =   0
      ScaleHeight     =   122.951
      ScaleMode       =   0  'User
      ScaleWidth      =   215.612
      TabIndex        =   20
      Top             =   4680
      Visible         =   0   'False
      Width           =   7695
   End
   Begin VB.PictureBox picgraph 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      ForeColor       =   &H0000FF00&
      Height          =   1150
      Index           =   3
      Left            =   0
      ScaleHeight     =   122.951
      ScaleMode       =   0  'User
      ScaleWidth      =   215.612
      TabIndex        =   19
      Top             =   3240
      Visible         =   0   'False
      Width           =   7695
   End
   Begin VB.PictureBox picgraph 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      ForeColor       =   &H0000FF00&
      Height          =   1150
      Index           =   2
      Left            =   0
      ScaleHeight     =   122.951
      ScaleMode       =   0  'User
      ScaleWidth      =   215.612
      TabIndex        =   18
      Top             =   1800
      Visible         =   0   'False
      Width           =   7695
   End
   Begin VB.PictureBox picgraph 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      ForeColor       =   &H0000FF00&
      Height          =   1150
      Index           =   1
      Left            =   0
      ScaleHeight     =   122.951
      ScaleMode       =   0  'User
      ScaleWidth      =   215.612
      TabIndex        =   17
      Top             =   360
      Visible         =   0   'False
      Width           =   7695
   End
   Begin VB.Timer bounce 
      Enabled         =   0   'False
      Interval        =   80
      Left            =   360
      Top             =   3240
   End
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   3360
      Top             =   2640
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   3360
      Top             =   1920
   End
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      Height          =   735
      Left            =   3360
      ScaleHeight     =   675
      ScaleWidth      =   915
      TabIndex        =   10
      Top             =   1440
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Timer slide 
      Left            =   6000
      Top             =   720
   End
   Begin VB.FileListBox File1 
      Height          =   1455
      Left            =   5520
      Pattern         =   "*.jpg;*.bmp;*.gif"
      TabIndex        =   6
      Top             =   1920
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   3000
      Top             =   480
   End
   Begin VB.Label gfan3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Fan3"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   435
      Left            =   0
      TabIndex        =   29
      Top             =   7200
      Visible         =   0   'False
      Width           =   810
   End
   Begin VB.Label gfan2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Fan2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   435
      Left            =   0
      TabIndex        =   28
      Top             =   5760
      Visible         =   0   'False
      Width           =   810
   End
   Begin VB.Label gfan1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Fan1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   435
      Left            =   0
      TabIndex        =   27
      Top             =   4320
      Visible         =   0   'False
      Width           =   810
   End
   Begin VB.Label gtemp3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Temperature3"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   435
      Left            =   0
      TabIndex        =   26
      Top             =   2880
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.Label gtemp2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Temperature2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   435
      Left            =   0
      TabIndex        =   25
      Top             =   1440
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.Label gtemp1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Temperature1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   435
      Left            =   0
      TabIndex        =   24
      Top             =   0
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.Label label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   72
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1785
      Index           =   9
      Left            =   0
      TabIndex        =   23
      Top             =   3960
      Visible         =   0   'False
      Width           =   435
   End
   Begin VB.Label label1 
      AutoSize        =   -1  'True
      Height          =   195
      Index           =   6
      Left            =   0
      TabIndex        =   16
      Tag             =   "cpu"
      Top             =   8520
      Visible         =   0   'False
      Width           =   45
   End
   Begin VB.Label label1 
      AutoSize        =   -1  'True
      Height          =   195
      Index           =   7
      Left            =   0
      TabIndex        =   15
      Tag             =   "mem"
      Top             =   8760
      Visible         =   0   'False
      Width           =   45
   End
   Begin VB.Label label1 
      AutoSize        =   -1  'True
      Height          =   195
      Index           =   8
      Left            =   0
      TabIndex        =   14
      Top             =   9000
      Visible         =   0   'False
      Width           =   45
   End
   Begin VB.Label wintime 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   360
      Left            =   13680
      TabIndex        =   13
      Top             =   360
      Width           =   90
   End
   Begin VB.Label winamp 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   405
      Left            =   14040
      TabIndex        =   12
      Top             =   0
      Width           =   90
   End
   Begin VB.Label label7 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   435
      Left            =   0
      TabIndex        =   11
      Tag             =   "time"
      Top             =   11040
      Visible         =   0   'False
      Width           =   90
   End
   Begin VB.Label counte 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   11040
      TabIndex        =   9
      Top             =   8520
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label counter 
      BackColor       =   &H00000000&
      Caption         =   "of"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   10560
      TabIndex        =   8
      Top             =   8520
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label coun 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   9840
      TabIndex        =   7
      Top             =   8520
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   72
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1785
      Index           =   1
      Left            =   0
      TabIndex        =   2
      Top             =   2640
      Width           =   435
   End
   Begin VB.Label label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   72
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1785
      Index           =   5
      Left            =   0
      TabIndex        =   5
      Top             =   7080
      Width           =   435
   End
   Begin VB.Label label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   72
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1785
      Index           =   3
      Left            =   0
      TabIndex        =   4
      Top             =   4200
      Width           =   435
   End
   Begin VB.Label label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   72
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1785
      Index           =   4
      Left            =   0
      TabIndex        =   3
      Top             =   5640
      Width           =   435
   End
   Begin VB.Label label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   72
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1785
      Index           =   2
      Left            =   0
      TabIndex        =   1
      Top             =   2760
      Width           =   435
   End
   Begin VB.Label label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   72
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1785
      Index           =   0
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   435
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      Height          =   8415
      Left            =   0
      Stretch         =   -1  'True
      Top             =   -120
      Width           =   9615
   End
End
Attribute VB_Name = "saver"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'MBM Saver
'by Fotis from Greece
'www.fetix.8m.com
'robot@mail.gr
'i have used many lines from planet-source-code
'to make the screen saver more powerfull
'so i've missed some names of the authors
'in the modules
'I am a beginner and i have some problems with
'the code:when the image1.stretch is enable the
'screen saver takes too much of CPU usage.
'Can anoyone fix this?
'please mail me if you achieve it

'the code is a mess,i know

Option Explicit
'for labels movement
Private rndi As Byte
Private speedX() As Byte
Private speedY() As Byte
Private y(), x() As Integer
Private maxX() As Boolean
Private maxY() As Boolean

Private fntnm As String 'font name
Private farh As Byte
Private linechart As Boolean

Private spaces As Integer
Private showHigh As Boolean, showAver As Boolean, showMin As Boolean
Private refreshrate As Integer
Private TminVal As Long, TmaxVal As Long, FminVal As Long, FmaxVal As Long

Private temp1name As String, temp2name As String, temp3name As String, fan1name As String, fan2name As String, fan3name As String

Private a As Boolean
Private b As Boolean
Private c As Boolean
Private d As Boolean
Private e As Boolean
Private f As Boolean
Private tp As Integer
Private fadein As Byte
Private clr As String
Private fadeout As Byte
Private Ret As Long 'win9x or winnt
Private mng As Boolean
Private QueryObje As Object
Private LastX As Single 'for mouse movement
Private LastY As Single 'for mouse movement
Private Sub bounce_Timer()
For rndi = 0 To Label1.Count - 1
If Label1(rndi).Left > Me.Width - Label1(rndi).Width Then maxX(rndi) = True
If Label1(rndi).Left < 0 Then maxX(rndi) = False
If Label1(rndi).Top > Me.Height - 1.8 * Label1(rndi).Height Then maxY(rndi) = True
If Label1(rndi).Top < 0 Then maxY(rndi) = False
If maxX(rndi) = False Then
x(rndi) = x(rndi) + speedX(rndi) * 30
Else
x(rndi) = x(rndi) - speedX(rndi) * 30
End If
If maxY(rndi) = False Then
y(rndi) = y(rndi) + speedY(rndi) * 30
Else
y(rndi) = y(rndi) - speedY(rndi) * 30
End If
Label1(rndi).Move x(rndi), y(rndi)
Next rndi
End Sub
Private Sub coun_Change()
'slide show
DoEvents
On Error Resume Next
If coun.Caption = counte.Caption + 1 Then coun.Caption = 1
Picture1.Picture = LoadPicture(File1.Path & "\" & File1.List(coun.Caption - 1))
mng = True
Image1.Visible = False
'stretch the picture
If Image1.Stretch = True Then
    If Me.Width / Picture1.Width > Me.Height / Picture1.Height Then
      Image1.Height = Me.Height
      Image1.Width = Picture1.Width * Me.Height / Picture1.Height
    Else
      Image1.Width = Me.Width
      Image1.Height = Picture1.Height * Me.Width / Picture1.Width
    End If
End If
Label7.Caption = File1.List(coun.Caption - 1)
Image1.Top = 0
Image1.Left = 0
Image1.Picture = Picture1.Picture
'center the picture
If GetSetting("MBM Saver", "settings", "cntr", "0") = 1 Then
  If Image1.Width < Me.Width Then
   Image1.Left = (Me.Width - Image1.Width) / 2
   mng = True
  End If
  If Image1.Height < Me.Height Then
   Image1.Top = (Me.Height - Image1.Height) / 2
   mng = True
  End If
End If
Image1.Visible = True
End Sub
Private Sub Form_Initialize()
'fade in
If (GetSetting("MBM Saver", "settings", "fadein", "0") = "1") And (saverpreview = False) Then
Mache_Transparent saver.hwnd, 1
Timer2.Enabled = True
End If
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'exit screen saver
If GetSetting("MBM Saver", "settings", "onlywith", "0") = "1" Then
  If KeyCode = 27 Then 'Esc
    If GetSetting("MBM Saver", "settings", "fadeout", "0") = 1 Then
      Timer3.Enabled = True 'fade out
    Else
      Unload Me
    End If
  End If
Else
  If GetSetting("MBM Saver", "settings", "fadeout", "0") = 1 Then
   Timer3.Enabled = True 'fade out
  Else
   Unload Me
  End If
End If
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
'exit screen saver
  If GetSetting("MBM Saver", "settings", "onlywith", "0") = "1" Then
    If KeyAscii = 27 Then 'Esc
      If GetSetting("MBM Saver", "settings", "fadeout", "0") = 1 Then
        Timer3.Enabled = True 'Fade out
      Else
        Unload Me
      End If
    End If
  Else
    If GetSetting("MBM Saver", "settings", "fadeout", "0") = 1 Then
       Timer3.Enabled = True 'Fade out
    Else
       Unload Me
    End If
  End If
End Sub
Private Sub Form_Load()
'load the settings from registry

If (GetSetting("MBM Saver", "settings", "check11", "0") = "1") And (saverpreview = False) Then
 Call Mache_Transparent(saver.hwnd, GetSetting("MBM Saver", "settings", "Hscroll1", "251"))
 fixx = True
End If
SetThreadPriority GetCurrentThread, THREAD_BASE_PRIORITY_MAX
SetPriorityClass GetCurrentProcess, HIGH_PRIORITY_CLASS
    'Initialize our Query object
    If IsWinNT Then
        Set QueryObje = New clsCPUUsageNT
    Else
        Set QueryObje = New clsCPUUsage
    End If
    QueryObje.Initialize
mng = False
FindWinamp
If FindWinamp = False Then wintime.Visible = False
Image1.Width = Me.Width
Image1.Height = Me.Height
If GetSetting("MBM Saver", "settings", "stre", "0") = 1 Then
Image1.Stretch = True
Else
Image1.Stretch = False
End If
Timer1.Interval = GetSetting("MBM Saver", "settings", "refreshtime", "1000")
'graph
TmaxVal = GetSetting("MBM Saver", "settings", "maxvalue", "80")
TminVal = GetSetting("MBM Saver", "settings", "minvalue", "-20")
FmaxVal = GetSetting("MBM Saver", "settings", "fmaxvalue", "7500")
FminVal = GetSetting("MBM Saver", "settings", "fminvalue", "1000")
refreshrate = GetSetting("MBM Saver", "settings", "refreshtime", "3")
showAver = GetSetting("MBM Saver", "settings", "disaver", "0")
showMin = GetSetting("MBM Saver", "settings", "dislow", "0")
showHigh = GetSetting("MBM Saver", "settings", "dishi", "0")
spaces = GetSetting("MBM Saver", "settings", "spaces", "0")
temp1name = GetSetting("MBM Saver", "settings", "name1", "temp1")
temp2name = GetSetting("MBM Saver", "settings", "name2", "temp2")
temp3name = GetSetting("MBM Saver", "settings", "name3", "temp3")
fan1name = GetSetting("MBM Saver", "settings", "name4", "fan1")
fan2name = GetSetting("MBM Saver", "settings", "name5", "fan2")
fan3name = GetSetting("MBM Saver", "settings", "name6", "fan3")
gtemp1.Caption = temp1name & ":"
gtemp2.Caption = temp2name & ":"
gtemp3.Caption = temp3name & ":"
gfan1.Caption = fan1name & ":"
gfan2.Caption = fan2name & ":"
gfan3.Caption = fan3name & ":"

If GetSetting("MBM Saver", "settings", "check12", "0") = 1 Then Label1(8).Visible = True
If GetSetting("MBM Saver", "settings", "check13", "0") = 1 Then
Label1(7).Visible = True
Label1(6).Visible = True
End If
If GetSetting("MBM Saver", "settings", "winuptm", "0") = 1 Then Label1(9).Visible = True
If GetSetting("MBM Saver", "settings", "farhenheit", "False") = True Then
 farh = 1
Else
 farh = 2
End If

If GetSetting("MBM Saver", "settings", "check10", "0") = 0 Then
 winamp.Visible = False
 wintime.Visible = False
End If
If (GetSetting("MBM Saver", "settings", "check8", "0") = 0) Then
Label1(0).BackStyle = 1
Label1(1).BackStyle = 1
Label1(2).BackStyle = 1
Label1(3).BackStyle = 1
Label1(4).BackStyle = 1
Label1(5).BackStyle = 1
Label1(6).BackStyle = 1
Label1(7).BackStyle = 1
Label1(8).BackStyle = 1
Label1(9).BackStyle = 1
counte.BackStyle = 1
counter.BackStyle = 1
coun.BackStyle = 1
Label7.BackStyle = 1
wintime.BackStyle = 1
winamp.BackStyle = 1
Else
winamp.BackStyle = 0
wintime.BackStyle = 0
Label7.BackStyle = 0
counte.BackStyle = 0
counter.BackStyle = 0
coun.BackStyle = 0
Label1(0).BackStyle = 0
Label1(1).BackStyle = 0
Label1(2).BackStyle = 0
Label1(3).BackStyle = 0
Label1(4).BackStyle = 0
Label1(5).BackStyle = 0
Label1(6).BackStyle = 0
Label1(7).BackStyle = 0
Label1(8).BackStyle = 0
Label1(9).BackStyle = 0
End If
clor GetSetting("MBM Saver", "settings", "backcol", "Black")
Label1(0).BackColor = clr
Label1(1).BackColor = clr
Label1(2).BackColor = clr
Label1(3).BackColor = clr
Label1(4).BackColor = clr
Label1(5).BackColor = clr
Label1(6).BackColor = clr
Label1(7).BackColor = clr
Label1(8).BackColor = clr
Label1(9).BackColor = clr
counte.BackColor = clr
coun.BackColor = clr
counter.BackColor = clr
Label7.BackColor = clr
wintime.BackColor = clr
winamp.BackColor = clr
On Error GoTo nophoto
If (GetSetting("MBM Saver", "settings", "option1", "False") = True) And (GetSetting("MBM Saver", "settings", "check7", "0") = 1) Then Picture1.Picture = LoadPicture(GetSetting("MBM Saver", "settings", "br", ""))
Image1.Picture = Picture1.Picture
If Image1.Stretch = True Then
If Me.Width / Picture1.Width > Me.Height / Picture1.Height Then
Image1.Height = Me.Height - 250
Image1.Width = Picture1.Width * Me.Height / Picture1.Height
Else
Image1.Width = Me.Width
Image1.Height = Picture1.Height * Me.Width / Picture1.Width
End If
End If
  If GetSetting("MBM Saver", "settings", "cntr", "0") = 1 Then
    If Image1.Width < Me.Width Then
    Image1.Left = (Me.Width - Image1.Width) / 2
    mng = True
    End If
    If Image1.Height < Me.Height Then
    Image1.Top = (Me.Height - Image1.Height) / 2
    mng = True
    End If
  End If
If (GetSetting("MBM Saver", "settings", "option2", "False") = "True") And (GetSetting("MBM Saver", "settings", "check7", "0") = 1) Then
 slide.Interval = GetSetting("MBM Saver", "settings", "delay", "4000")
 File1.Path = GetSetting("MBM Saver", "settings", "sli", "C:\")
 If GetSetting("MBM Saver", "settings", "check9", "0") = 1 Then Label7.Visible = True
 coun.Visible = True
 counter.Visible = True
 counte.Visible = True
 coun.Caption = "1"
 counte.Caption = File1.ListCount
 End If
 GoTo fddd
nophoto:
 slide.Interval = 0
fddd:
Me.WindowState = vbMaximized
Dim colo As String

CursorVisible = False
clor GetSetting("MBM Saver", "settings", "combo1", "Black")
saver.BackColor = clr
clor GetSetting("MBM Saver", "settings", "combo2", "White")
Label1(0).ForeColor = clr
Label1(1).ForeColor = clr
Label1(2).ForeColor = clr
Label1(3).ForeColor = clr
Label1(4).ForeColor = clr
Label1(5).ForeColor = clr
Label1(6).ForeColor = clr
Label1(7).ForeColor = clr
Label1(8).ForeColor = clr
Label1(9).ForeColor = clr
coun.ForeColor = clr
counte.ForeColor = clr
counter.ForeColor = clr
Label7.ForeColor = clr
winamp.ForeColor = clr
wintime.ForeColor = clr
colo = GetSetting("MBM Saver", "settings", "combo3", "28")
Label1(0).FontSize = colo
Label1(1).FontSize = colo
Label1(2).FontSize = colo
Label1(3).FontSize = colo
Label1(4).FontSize = colo
Label1(5).FontSize = colo
Label1(6).FontSize = colo
Label1(7).FontSize = colo
Label1(8).FontSize = colo
Label1(9).FontSize = colo
fntnm = GetSetting("MBM Saver", "settings", "fontnamec", winamp.Font)
Label1(0).Font = fntnm
Label1(1).Font = fntnm
Label1(2).Font = fntnm
Label1(3).Font = fntnm
Label1(4).Font = fntnm
Label1(5).Font = fntnm
Label1(6).Font = fntnm
Label1(7).Font = fntnm
Label1(8).Font = fntnm
Label1(9).Font = fntnm
Label1(9).FontBold = False

Label7.Font = fntnm
counte.Font = fntnm
counter.Font = fntnm
coun.Font = fntnm
winamp.Font = fntnm
wintime.Font = fntnm
a = GetSetting("MBM Saver", "settings", "check1", False)
b = GetSetting("MBM Saver", "settings", "check2", False)
c = GetSetting("MBM Saver", "settings", "check3", False)
d = GetSetting("MBM Saver", "settings", "check4", False)
e = GetSetting("MBM Saver", "settings", "check5", False)
f = GetSetting("MBM Saver", "settings", "check6", False)
Mnew = True
linechart = GetSetting("MBM Saver", "settings", "history", "0") = "1"
If linechart = True Then
gtemp1.Visible = a
gtemp2.Visible = b
gtemp3.Visible = c
gfan1.Visible = d
gfan2.Visible = e
gfan3.Visible = f
picgraph(1).Visible = a
picgraph(2).Visible = b
picgraph(3).Visible = c
picgraph(4).Visible = d
picgraph(5).Visible = e
picgraph(6).Visible = f
winamp.Visible = False
wintime.Visible = False
Label7.Visible = False
counte.Visible = False
counter.Visible = False
coun.Visible = False
Label1(0).Visible = False
Label1(1).Visible = False
Label1(2).Visible = False
Label1(3).Visible = False
Label1(4).Visible = False
Label1(5).Visible = False
Label1(6).Visible = False
Label1(7).Visible = False
Label1(8).Visible = False
Label1(9).Visible = False
Else
Label1(0).Visible = a
Label1(1).Visible = b
Label1(2).Visible = c
Label1(3).Visible = d
Label1(4).Visible = e
Label1(5).Visible = f
If GetSetting("MBM Saver", "settings", "stan", "True") = "True" Then
If Label1(0).Visible = True Then
Label1(0).Top = tp
tp = tp + 24 * colo
End If
If Label1(1).Visible = True Then
Label1(1).Top = tp
tp = tp + 24 * colo
End If
If Label1(2).Visible = True Then
Label1(2).Top = tp
tp = tp + 24 * colo
End If
If Label1(3).Visible = True Then
Label1(3).Top = tp
tp = tp + 24 * colo
End If
If Label1(4).Visible = True Then
Label1(4).Top = tp
tp = tp + 24 * colo
End If
If Label1(5).Visible = True Then
Label1(5).Top = tp
tp = tp + 24 * colo
End If
If Label1(6).Visible = True Then
Label1(6).Top = tp
tp = tp + 24 * colo
End If
If Label1(7).Visible = True Then
Label1(7).Top = tp
tp = tp + 24 * colo
End If
If Label1(8).Visible = True Then
Label1(8).Top = tp
tp = tp + 24 * colo
End If
If Label1(9).Visible = True Then
Label1(9).Top = tp
tp = tp + 24 * colo
End If

Else
Label1(0).Left = GetSetting("MBM Saver", "settings", "temp1l", "0") * Me.Width / 3135
Label1(1).Left = GetSetting("MBM Saver", "settings", "temp2l", "0") * Me.Width / 3135
Label1(2).Left = GetSetting("MBM Saver", "settings", "temp3l", "0") * Me.Width / 3135
Label1(3).Left = GetSetting("MBM Saver", "settings", "fan1l", "0") * Me.Width / 3135
Label1(4).Left = GetSetting("MBM Saver", "settings", "fan2l", "0") * Me.Width / 3135
Label1(5).Left = GetSetting("MBM Saver", "settings", "fan3l", "0") * Me.Width / 3135
Label1(6).Left = GetSetting("MBM Saver", "settings", "lcpul", "0") * Me.Width / 3135
Label1(8).Left = GetSetting("MBM Saver", "settings", "ltimel", "0") * Me.Width / 3135
Label1(7).Left = GetSetting("MBM Saver", "settings", "memoryl", "0") * Me.Width / 3135
Label1(9).Left = GetSetting("MBM Saver", "settings", "wintimel", "600") * Me.Width / 3135

If Screen.Width / Screen.TwipsPerPixelX * Screen.Height / Screen.TwipsPerPixelY < 995000 Then
Label1(0).Top = GetSetting("MBM Saver", "settings", "temp1", "0") * Me.Height / 2351
Label1(1).Top = GetSetting("MBM Saver", "settings", "temp2", "240") * Me.Height / 2351
Label1(2).Top = GetSetting("MBM Saver", "settings", "temp3", "480") * Me.Height / 2351
Label1(3).Top = GetSetting("MBM Saver", "settings", "fan1", "720") * Me.Height / 2351
Label1(4).Top = GetSetting("MBM Saver", "settings", "fan2", "960") * Me.Height / 2351
Label1(5).Top = GetSetting("MBM Saver", "settings", "fan3", "1200") * Me.Height / 2351
Label1(6).Top = GetSetting("MBM Saver", "settings", "lcpu", "1440") * Me.Height / 2351
Label1(8).Top = GetSetting("MBM Saver", "settings", "ltime", "1680") * Me.Height / 2351
Label1(7).Top = GetSetting("MBM Saver", "settings", "memory", "1920") * Me.Height / 2351
Label1(9).Top = GetSetting("MBM Saver", "settings", "wintime", "2040") * Me.Height / 2351
End If
If (Screen.Width / Screen.TwipsPerPixelX * Screen.Height / Screen.TwipsPerPixelY > 995000) And (Screen.Width / Screen.TwipsPerPixelX * Screen.Height / Screen.TwipsPerPixelY < 1300000) Then
Label1(0).Top = GetSetting("MBM Saver", "settings", "temp1", "0") * Me.Height / 1950
Label1(1).Top = GetSetting("MBM Saver", "settings", "temp2", "240") * Me.Height / 1950
Label1(2).Top = GetSetting("MBM Saver", "settings", "temp3", "480") * Me.Height / 1950
Label1(3).Top = GetSetting("MBM Saver", "settings", "fan1", "720") * Me.Height / 1950
Label1(4).Top = GetSetting("MBM Saver", "settings", "fan2", "960") * Me.Height / 1950
Label1(5).Top = GetSetting("MBM Saver", "settings", "fan3", "1200") * Me.Height / 1950
Label1(6).Top = GetSetting("MBM Saver", "settings", "lcpu", "1440") * Me.Height / 1950
Label1(8).Top = GetSetting("MBM Saver", "settings", "ltime", "1680") * Me.Height / 1950
Label1(7).Top = GetSetting("MBM Saver", "settings", "memory", "1920") * Me.Height / 1950
Label1(9).Top = GetSetting("MBM Saver", "settings", "wintime", "2040") * Me.Height / 1950
End If
If Screen.Width / Screen.TwipsPerPixelX * Screen.Height / Screen.TwipsPerPixelY > 1300000 Then
Label1(0).Top = GetSetting("MBM Saver", "settings", "temp1", "0") * Me.Height / 1666
Label1(1).Top = GetSetting("MBM Saver", "settings", "temp2", "240") * Me.Height / 1666
Label1(2).Top = GetSetting("MBM Saver", "settings", "temp3", "480") * Me.Height / 1666
Label1(3).Top = GetSetting("MBM Saver", "settings", "fan1", "720") * Me.Height / 1666
Label1(4).Top = GetSetting("MBM Saver", "settings", "fan2", "960") * Me.Height / 1666
Label1(5).Top = GetSetting("MBM Saver", "settings", "fan3", "1200") * Me.Height / 1666
Label1(6).Top = GetSetting("MBM Saver", "settings", "lcpu", "1440") * Me.Height / 1666
Label1(8).Top = GetSetting("MBM Saver", "settings", "ltime", "1680") * Me.Height / 1666
Label1(7).Top = GetSetting("MBM Saver", "settings", "memory", "1920") * Me.Height / 1666
Label1(9).Top = GetSetting("MBM Saver", "settings", "wintime", "2040") * Me.Height / 1666
End If
End If
End If
'bounce
If (GetSetting("MBM Saver", "settings", "moveme", "0") = "1") Then
ReDim speedX(0 To Label1.Count)
ReDim speedY(0 To Label1.Count)
ReDim y(0 To Label1.Count)
ReDim x(0 To Label1.Count)
ReDim maxX(0 To Label1.Count)
ReDim maxY(0 To Label1.Count)
On Error Resume Next
maxX(1) = False
maxY(1) = True
For rndi = 0 To Label1.Count - 1
x(rndi) = Label1(rndi).Left
y(rndi) = Label1(rndi).Top
Randomize
speedX(rndi) = Int(Rnd * 5) + 1
Randomize
speedY(rndi) = Int(Rnd * 5) + 1
maxX(rndi) = Not (maxX(Rnd - 1))
maxY(rndi) = Not (maxY(Rnd - 1))
Next rndi
bounce.Enabled = True
End If
'speed problem
If (fixx = False) And (saver.Image1.Stretch = True) And (saverpreview = False) Then Mache_Transparent saver.hwnd, 255

End Sub
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
If saverpreview = True Then Exit Sub
'exit screen saver
 If GetSetting("MBM Saver", "settings", "onlywith", "0") = "0" Then
   If GetSetting("MBM Saver", "settings", "fadeout", "0") = 1 Then Timer3.Enabled = True
 Else
   Unload Me
 End If
End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
 If saverpreview = True Then Exit Sub
'exit screen saver
If mng = False Then
If (LastX = 0 And LastY = 0) Or (Abs(LastX - x) < 2 And Abs(LastY - y) < 2) Then
   LastX = x
   LastY = y
  Else
  If GetSetting("MBM Saver", "settings", "onlywith", "0") = "0" Then
    If GetSetting("MBM Saver", "settings", "fadeout", "0") = 1 Then
      Timer3.Enabled = True
    Else
      Unload Me
    End If
  End If
  End If
  Else
  mng = False
  End If
End Sub
Private Sub Form_Resize()
Label7.Top = Me.Height - 480
counte.Left = Me.Width - 900
counte.Top = Me.Height - 480
coun.Top = counte.Top
counter.Top = coun.Top
counter.Left = counte.Left - 480
coun.Left = counter.Left - 720
winamp.Left = Me.Width - winamp.Width
wintime.Left = Me.Width - wintime.Width
picgraph(1).Width = Me.Width
picgraph(2).Width = Me.Width
picgraph(3).Width = Me.Width
picgraph(4).Width = Me.Width
picgraph(5).Width = Me.Width
picgraph(6).Width = Me.Width
End Sub
Private Sub Form_Terminate()
CursorVisible = True
End Sub
Private Sub Form_Unload(Cancel As Integer)
CursorVisible = True
End Sub
Private Sub Image1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
'exit screen saver
If saverpreview = True Then Exit Sub
 If GetSetting("MBM Saver", "settings", "onlywith", "0") = "0" Then
   If GetSetting("MBM Saver", "settings", "fadeout", "0") = 1 Then
     Timer3.Enabled = True
   Else
     Unload Me
   End If
 End If
End Sub
Private Sub Image1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
'exit screen saver
If saverpreview = True Then Exit Sub
If mng = False Then
If (LastX = 0 And LastY = 0) Or (Abs(LastX - x) < 2 And Abs(LastY - y) < 2) Then
   LastX = x
   LastY = y
  Else
  If GetSetting("MBM Saver", "settings", "onlywith", "0") = "0" Then
    If GetSetting("MBM Saver", "settings", "fadeout", "0") = 1 Then
      Timer3.Enabled = True
    Else
      Unload Me
    End If
  End If
  End If
  Else
  mng = False
  End If
End Sub
Private Sub slide_Timer()
'slide show timer
If File1.ListCount = 0 Then
coun.Caption = "0"
Exit Sub
End If
coun.Caption = coun.Caption + 1
End Sub
Private Function mini(Seconds As Integer) As String
'format the seconds to time
Dim tmp, sc As String
Dim mn As Integer
If Seconds > 59 Then
mn = Int(Seconds / 60)
sc = Seconds - (mn * 60)
tmp = sc
If tmp < 10 Then sc = "0" & tmp
mini = mn & ":" & sc
Else
If Seconds < 10 Then
tmp = "0" & Seconds
mini = "0" & ":" & tmp
Else
mini = "0:" & Seconds
End If
End If
End Function
Private Sub Timer1_Timer()
'refresh the labels
DoEvents
'winamp
If winamp.Visible = True Then
  If FindWindow("STUDIO", vbNullString) <> 0 Then
   'winamp3
   winamp.Caption = title(GetText(FindWindow("STUDIO", vbNullString)))
   wintime.Visible = False
  Else
   'winamp 2.x
   winamp.Caption = WM_GET(SONG_TITLE)
   wintime.Caption = mini(WM_GET(SONG_POSITION) / 1000) & "/" & mini(WM_GET(SONG_LENGTH))
  End If
End If

  
If Label1(6).Visible = True Then
 Ret = QueryObje.Query
 Label1(6).Caption = "CPU:" & Space(spaces) & CStr(Ret) + "% "
End If

If Label1(7).Visible = True Then Label1(7).Caption = "Mem:" & Space(spaces) & GetMemoryInfo
If Label1(8).Visible = True Then Label1(8).Caption = "Time:" & Space(spaces) & Time
If Label1(9).Visible = True Then Label1(9).Caption = "Windows up time:" & Space(spaces) & FormatCount(GetTickCount, DaysHoursMinutesSeconds)
myData = MBM_GetData
If myData.sdVersion = 0 Then Mnew = False

If Mnew = True Then
 myData = MBM_GetData
 If a = True Then
  If Label1(0).Visible = True Then
   Label1(0).Caption = temp1name & ":" & Space(spaces) & CtoF(myData.sdSensor(0).ssCurrent, farh)
  Else
   linechart = False
   gtemp1.Caption = temp1name & ":" & Space(spaces) & CtoF(myData.sdSensor(0).ssCurrent, farh)
   linechart = True
   AddValue CtoF(myData.sdSensor(0).ssCurrent, farh), 6, 1, picgraph(1), TmaxVal, TminVal, 200, showAver, showMin, showHigh, vbGreen, vbBlack, vbCyan
  End If
 End If
 
 If b = True Then
  If Label1(1).Visible = True Then
   Label1(1).Caption = temp2name & ":" & Space(spaces) & CtoF(myData.sdSensor(1).ssCurrent, farh)
  Else
   linechart = False
   gtemp2.Caption = temp2name & ":" & Space(spaces) & CtoF(myData.sdSensor(1).ssCurrent, farh)
   linechart = True
   AddValue CtoF(myData.sdSensor(1).ssCurrent, farh), 6, 2, picgraph(2), TmaxVal, TminVal, 200, showAver, showMin, showHigh, vbGreen, vbBlack, vbCyan
  End If
 End If

 If c = True Then
  If Label1(2).Visible = True Then
   Label1(2).Caption = temp3name & ":" & Space(spaces) & CtoF(myData.sdSensor(2).ssCurrent, farh)
  Else
   linechart = False
   gtemp3.Caption = temp3name & ":" & Space(spaces) & CtoF(myData.sdSensor(2).ssCurrent, farh)
   linechart = True
   AddValue CtoF(myData.sdSensor(2).ssCurrent, farh), 6, 3, picgraph(3), TmaxVal, TminVal, 200, showAver, showMin, showHigh, vbGreen, vbBlack, vbCyan
  End If
 End If
 
 If d = True Then
  If Label1(3).Visible = True Then
   Label1(3).Caption = fan1name & ":" & Space(spaces) & myData.sdSensor(48).ssCurrent & "rpm"
  Else
   gfan1.Caption = fan1name & ":" & Space(spaces) & myData.sdSensor(48).ssCurrent & "rpm"
   AddValue myData.sdSensor(48).ssCurrent, 6, 4, picgraph(4), FmaxVal, FminVal, 200, showAver, showMin, showHigh, vbGreen, vbBlack, vbCyan
  End If
 End If
 
 If e = True Then
  If Label1(4).Visible = True Then
   Label1(4).Caption = fan2name & ":" & Space(spaces) & myData.sdSensor(49).ssCurrent & "rpm"
  Else
   gfan2.Caption = fan2name & ":" & Space(spaces) & myData.sdSensor(49).ssCurrent & "rpm"
   AddValue myData.sdSensor(49).ssCurrent, 6, 5, picgraph(5), FmaxVal, FminVal, 200, showAver, showMin, showHigh, vbGreen, vbBlack, vbCyan
  End If
 End If
 
 If f = True Then
  If Label1(5).Visible = True Then
   Label1(5).Caption = fan3name & ":" & Space(spaces) & myData.sdSensor(50).ssCurrent & "rpm"
  Else
   gfan3.Caption = fan3name & ":" & Space(spaces) & myData.sdSensor(50).ssCurrent & "rpm"
   AddValue myData.sdSensor(50).ssCurrent, 6, 6, picgraph(6), FmaxVal, FminVal, 200, showAver, showMin, showHigh, vbGreen, vbBlack, vbCyan
  End If
 End If

Else
 
 If a = True Then
  If Label1(0).Visible = True Then
   Label1(0).Caption = temp1name & ":" & Space(spaces) & CtoF(MBM_GetData1.STemperature(1), farh)
  Else
   linechart = False
   gtemp1.Caption = temp1name & ":" & Space(spaces) & CtoF(myData.sdSensor(0).ssCurrent, farh)
   linechart = True
   AddValue CtoF(myData.sdSensor(0).ssCurrent, farh), 6, 1, picgraph(1), TmaxVal, TminVal, 200, showAver, showMin, showHigh, vbGreen, vbBlack, vbCyan
  End If
 End If
 
 If b = True Then
  If Label1(1).Visible = True Then
   Label1(1).Caption = temp2name & ":" & Space(spaces) & CtoF(MBM_GetData1.STemperature(2), farh)
  Else
   linechart = False
   gtemp2.Caption = temp2name & ":" & Space(spaces) & CtoF(MBM_GetData1.STemperature(2), farh)
   linechart = True
   AddValue CtoF(MBM_GetData1.STemperature(2), farh), 6, 2, picgraph(2), TmaxVal, TminVal, 200, showAver, showMin, showHigh, vbGreen, vbBlack, vbCyan
  End If
 End If
 
 If c = True Then
  If Label1(2).Visible = True Then
   Label1(2).Caption = temp3name & ":" & Space(spaces) & CtoF(MBM_GetData1.STemperature(3), farh)
  Else
   linechart = False
   gtemp3.Caption = temp3name & ":" & Space(spaces) & CtoF(MBM_GetData1.STemperature(3), farh)
   linechart = True
   AddValue CtoF(MBM_GetData1.STemperature(3), farh), 6, 3, picgraph(3), TmaxVal, TminVal, 200, showAver, showMin, showHigh, vbGreen, vbBlack, vbCyan
  End If
 End If
 
 Dim da, ea, fa As Integer
 da = MBM_GetData1.SFan(1)
 If da = 255 Then da = 0
 If d = True Then Label1(3).Caption = GetSetting("MBM Saver", "settings", "name4", "fan1") & ":" & Space(spaces) & da & "rpm"
 ea = MBM_GetData1.SFan(2)
 If ea = 255 Then ea = 0
 If e = True Then Label1(4).Caption = GetSetting("MBM Saver", "settings", "name5", "fan2") & ":" & Space(spaces) & ea & "rpm"
 fa = MBM_GetData1.SFan(3)
 If fa = 255 Then fa = 0
 If f = True Then Label1(5).Caption = GetSetting("MBM Saver", "settings", "name6", "fan3") & ":" & Space(spaces) & fa & "rpm"
End If

End Sub
Private Sub timer2_timer()
 'fade in timer
 fadein = fadein + 1
 If (GetSetting("MBM Saver", "settings", "check11", "0") = "1") And (saverpreview = False) Then
  If fadein = GetSetting("MBM Saver", "settings", "Hscroll1", "251") Then Timer2.Enabled = False
  Mache_Transparent saver.hwnd, fadein
 Else
    Mache_Transparent saver.hwnd, fadein
 End If
 If fadein = 255 Then Timer2.Enabled = False
End Sub
Private Sub Timer3_Timer()
'fade out timer
If Timer2.Enabled = True Then Exit Sub
fadeout = fadeout + 1
 If GetSetting("MBM Saver", "settings", "check11", "0") = 1 Then
  If GetSetting("MBM Saver", "settings", "Hscroll1", "251") = fadeout Then
   CursorVisible = True
   End
  End If
  Mache_Transparent saver.hwnd, GetSetting("MBM Saver", "settings", "Hscroll1", "251") - fadeout
 Else
  Mache_Transparent saver.hwnd, 251 - fadeout
  If fadeout = 251 Then
   Timer3.Enabled = False
   CursorVisible = True
  End
 End If
 End If
End Sub
Private Sub winamp_Change()
winamp.Left = Me.Width - winamp.Width
End Sub
Private Sub wintime_Change()
wintime.Left = Me.Width - wintime.Width
End Sub
Private Sub clor(lcolor As String)
'change the string name to color code
Select Case lcolor
Case "Black"
clr = &H0&
Case "Yellow"
clr = &HFFFF&
Case "Red"
clr = &HFF&
Case "Brown"
clr = &H80C0FF
Case "Blue"
clr = &HFF0000
Case "White"
clr = &HFFFFFF
Case "Pink"
clr = &HFF80FF
Case "Green"
clr = &HFF00&
End Select
End Sub
Private Function CtoF(ByVal mTemp As String, mCelFah As Byte) As String
Dim f As Integer, c As Integer
   
   Select Case mCelFah
    Case 1
     If linechart = False Then
      CtoF = (mTemp * 1.8) + 32 & "°F"
     Else
      CtoF = (mTemp * 1.8) + 32
     End If
    Case 2
     If linechart = False Then
      CtoF = mTemp & "°C"
     Else
      CtoF = mTemp
     End If
   End Select
End Function
