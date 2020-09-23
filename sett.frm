VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form sett 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "MBM VERSION:"
   ClientHeight    =   6000
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4770
   Icon            =   "sett.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6000
   ScaleWidth      =   4770
   StartUpPosition =   3  'Windows Default
   Begin TabDlg.SSTab SSTab1 
      Height          =   5415
      Left            =   0
      TabIndex        =   31
      Top             =   0
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   9551
      _Version        =   327680
      Tabs            =   6
      TabHeight       =   520
      TabCaption(0)   =   "General"
      TabPicture(0)   =   "sett.frx":030A
      Tab(0).ControlCount=   11
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "versionMBM"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label24"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label23"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label32"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "spaces"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label34"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "onlywith"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Frame5"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "trans"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "refreshtime"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "VScroll1"
      Tab(0).Control(10).Enabled=   0   'False
      TabCaption(1)   =   "Sensors"
      TabPicture(1)   =   "sett.frx":0326
      Tab(1).ControlCount=   20
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "winuptm"
      Tab(1).Control(0).Enabled=   -1  'True
      Tab(1).Control(1)=   "Timer2"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Check13"
      Tab(1).Control(2).Enabled=   -1  'True
      Tab(1).Control(3)=   "Check12"
      Tab(1).Control(3).Enabled=   -1  'True
      Tab(1).Control(4)=   "Check10"
      Tab(1).Control(4).Enabled=   -1  'True
      Tab(1).Control(5)=   "Frame3"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "Frame4"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "Label22"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "winuptime"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).Control(9)=   "Line1(2)"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).Control(10)=   "Label21"
      Tab(1).Control(10).Enabled=   0   'False
      Tab(1).Control(11)=   "Line1(0)"
      Tab(1).Control(11).Enabled=   0   'False
      Tab(1).Control(12)=   "Line1(1)"
      Tab(1).Control(12).Enabled=   0   'False
      Tab(1).Control(13)=   "winamptitle"
      Tab(1).Control(13).Enabled=   0   'False
      Tab(1).Control(14)=   "Label16(3)"
      Tab(1).Control(14).Enabled=   0   'False
      Tab(1).Control(15)=   "Label17"
      Tab(1).Control(15).Enabled=   0   'False
      Tab(1).Control(16)=   "Label16(2)"
      Tab(1).Control(16).Enabled=   0   'False
      Tab(1).Control(17)=   "Label16(1)"
      Tab(1).Control(17).Enabled=   0   'False
      Tab(1).Control(18)=   "Label16(0)"
      Tab(1).Control(18).Enabled=   0   'False
      Tab(1).Control(19)=   "Label15"
      Tab(1).Control(19).Enabled=   0   'False
      TabCaption(2)   =   "Position"
      TabPicture(2)   =   "sett.frx":0342
      Tab(2).ControlCount=   6
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Label20"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "stan"
      Tab(2).Control(1).Enabled=   -1  'True
      Tab(2).Control(2)=   "custo"
      Tab(2).Control(2).Enabled=   -1  'True
      Tab(2).Control(3)=   "Picture1"
      Tab(2).Control(3).Enabled=   -1  'True
      Tab(2).Control(4)=   "moveme"
      Tab(2).Control(4).Enabled=   -1  'True
      Tab(2).Control(5)=   "resetbut"
      Tab(2).Control(5).Enabled=   -1  'True
      TabCaption(3)   =   "Fonts"
      TabPicture(3)   =   "sett.frx":035E
      Tab(3).ControlCount=   11
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "fontnamec"
      Tab(3).Control(0).Enabled=   -1  'True
      Tab(3).Control(1)=   "Combo2"
      Tab(3).Control(1).Enabled=   -1  'True
      Tab(3).Control(2)=   "backcol"
      Tab(3).Control(2).Enabled=   -1  'True
      Tab(3).Control(3)=   "Combo3"
      Tab(3).Control(3).Enabled=   -1  'True
      Tab(3).Control(4)=   "Check8"
      Tab(3).Control(4).Enabled=   -1  'True
      Tab(3).Control(5)=   "testt"
      Tab(3).Control(5).Enabled=   0   'False
      Tab(3).Control(6)=   "Label19"
      Tab(3).Control(6).Enabled=   0   'False
      Tab(3).Control(7)=   "Label18"
      Tab(3).Control(7).Enabled=   0   'False
      Tab(3).Control(8)=   "Label13"
      Tab(3).Control(8).Enabled=   0   'False
      Tab(3).Control(9)=   "Label9"
      Tab(3).Control(9).Enabled=   0   'False
      Tab(3).Control(10)=   "Label14"
      Tab(3).Control(10).Enabled=   0   'False
      TabCaption(4)   =   "Background"
      TabPicture(4)   =   "sett.frx":037A
      Tab(4).ControlCount=   4
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Combo1"
      Tab(4).Control(0).Enabled=   -1  'True
      Tab(4).Control(1)=   "Check7"
      Tab(4).Control(1).Enabled=   -1  'True
      Tab(4).Control(2)=   "Frame2"
      Tab(4).Control(2).Enabled=   0   'False
      Tab(4).Control(3)=   "Label12"
      Tab(4).Control(3).Enabled=   0   'False
      TabCaption(5)   =   "Line Chart(beta)"
      TabPicture(5)   =   "sett.frx":0396
      Tab(5).ControlCount=   12
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "fminvalue"
      Tab(5).Control(0).Enabled=   -1  'True
      Tab(5).Control(1)=   "fmaxvalue"
      Tab(5).Control(1).Enabled=   -1  'True
      Tab(5).Control(2)=   "Frame1"
      Tab(5).Control(2).Enabled=   0   'False
      Tab(5).Control(3)=   "minvalue"
      Tab(5).Control(3).Enabled=   -1  'True
      Tab(5).Control(4)=   "maxvalue"
      Tab(5).Control(4).Enabled=   -1  'True
      Tab(5).Control(5)=   "historycheck"
      Tab(5).Control(5).Enabled=   -1  'True
      Tab(5).Control(6)=   "Label26(1)"
      Tab(5).Control(6).Enabled=   0   'False
      Tab(5).Control(7)=   "Label25(1)"
      Tab(5).Control(7).Enabled=   0   'False
      Tab(5).Control(8)=   "Label28"
      Tab(5).Control(8).Enabled=   0   'False
      Tab(5).Control(9)=   "Label27"
      Tab(5).Control(9).Enabled=   0   'False
      Tab(5).Control(10)=   "Label26(0)"
      Tab(5).Control(10).Enabled=   0   'False
      Tab(5).Control(11)=   "Label25(0)"
      Tab(5).Control(11).Enabled=   0   'False
      Begin VB.VScrollBar VScroll1 
         Height          =   255
         Left            =   960
         Max             =   10
         TabIndex        =   127
         Top             =   2520
         Value           =   10
         Width           =   135
      End
      Begin VB.TextBox refreshtime 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1680
         TabIndex        =   123
         Text            =   "1000"
         Top             =   2040
         Width           =   615
      End
      Begin VB.TextBox fminvalue 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   -71640
         TabIndex        =   119
         Text            =   "1000"
         Top             =   2400
         Width           =   615
      End
      Begin VB.TextBox fmaxvalue 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   -71640
         TabIndex        =   118
         Text            =   "7500"
         Top             =   1920
         Width           =   615
      End
      Begin VB.Frame Frame1 
         Caption         =   "High Low and Average"
         Height          =   1935
         Left            =   -74760
         TabIndex        =   110
         Top             =   3120
         Width           =   4095
         Begin VB.CheckBox dislow 
            Alignment       =   1  'Right Justify
            Caption         =   "Display lower value"
            Height          =   255
            Left            =   240
            TabIndex        =   113
            Top             =   1320
            Width           =   1935
         End
         Begin VB.CheckBox dishi 
            Alignment       =   1  'Right Justify
            Caption         =   "Display higher value"
            Height          =   255
            Left            =   240
            TabIndex        =   112
            Top             =   840
            Width           =   1935
         End
         Begin VB.CheckBox disaver 
            Alignment       =   1  'Right Justify
            Caption         =   "Display average"
            Height          =   255
            Left            =   240
            TabIndex        =   111
            Top             =   360
            Width           =   1935
         End
         Begin VB.Label Label31 
            Caption         =   "Blue color"
            Height          =   255
            Left            =   2400
            TabIndex        =   122
            Top             =   1320
            Width           =   735
         End
         Begin VB.Label Label30 
            Caption         =   "Red color"
            Height          =   255
            Left            =   2400
            TabIndex        =   121
            Top             =   840
            Width           =   735
         End
         Begin VB.Label Label29 
            Caption         =   "Cyan color"
            Height          =   255
            Left            =   2400
            TabIndex        =   120
            Top             =   360
            Width           =   855
         End
      End
      Begin VB.TextBox minvalue 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   -73560
         TabIndex        =   109
         Text            =   "-20"
         Top             =   2400
         Width           =   495
      End
      Begin VB.TextBox maxvalue 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   -73560
         TabIndex        =   108
         Text            =   "80"
         Top             =   1920
         Width           =   495
      End
      Begin VB.CheckBox historycheck 
         Alignment       =   1  'Right Justify
         Caption         =   "Show temperatures and fans history"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -74640
         TabIndex        =   105
         ToolTipText     =   "If you enable it,you wont be able to use background picture,or any other sensors or position "
         Top             =   840
         Width           =   3495
      End
      Begin VB.CheckBox winuptm 
         Caption         =   "Monitor"
         Height          =   255
         Left            =   -72600
         TabIndex        =   100
         Top             =   4800
         Width           =   855
      End
      Begin VB.Timer Timer2 
         Interval        =   1000
         Left            =   -71040
         Top             =   4200
      End
      Begin VB.CommandButton resetbut 
         Caption         =   "Reset"
         Height          =   255
         Left            =   -74640
         TabIndex        =   99
         Top             =   1920
         Width           =   735
      End
      Begin VB.CheckBox moveme 
         Caption         =   "Labels movement"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -74760
         TabIndex        =   98
         Top             =   3840
         Width           =   1935
      End
      Begin VB.Frame trans 
         Caption         =   "For windows 2000 and XP"
         Height          =   1815
         Left            =   120
         TabIndex        =   91
         Top             =   3000
         Width           =   4215
         Begin VB.CheckBox Check11 
            Caption         =   "Transparent"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   161
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   95
            Top             =   480
            Width           =   1455
         End
         Begin VB.HScrollBar HScroll1 
            Height          =   255
            Left            =   1680
            Max             =   251
            Min             =   10
            TabIndex        =   94
            Top             =   480
            Value           =   10
            Width           =   1935
         End
         Begin VB.CheckBox fadein 
            Caption         =   "Open with fade-in"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   161
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   93
            Top             =   960
            Width           =   1935
         End
         Begin VB.CheckBox fadeout 
            Caption         =   "Close with fade-out"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   161
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   92
            Top             =   1440
            Width           =   2175
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            Caption         =   "%"
            Height          =   195
            Left            =   3960
            TabIndex        =   97
            Top             =   480
            Width           =   120
         End
         Begin VB.Label Label10 
            Caption         =   "0"
            Height          =   195
            Left            =   3720
            TabIndex        =   96
            Top             =   480
            Width           =   210
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "Temperature format:"
         Height          =   615
         Left            =   240
         TabIndex        =   88
         Top             =   720
         Width           =   4095
         Begin VB.OptionButton Option3 
            Caption         =   "Celcious"
            Height          =   255
            Left            =   360
            TabIndex        =   16
            Top             =   240
            Width           =   975
         End
         Begin VB.OptionButton Option4 
            Caption         =   "Fahrenheit"
            Height          =   255
            Left            =   2160
            TabIndex        =   17
            Top             =   240
            Width           =   1095
         End
      End
      Begin VB.PictureBox Picture1 
         Height          =   2351
         Left            =   -73680
         ScaleHeight     =   2295
         ScaleWidth      =   3075
         TabIndex        =   79
         Top             =   720
         Width           =   3135
         Begin VB.Label wintime 
            AutoSize        =   -1  'True
            Caption         =   "windows up time:dd:hh:mm:ss"
            Height          =   195
            Left            =   600
            TabIndex        =   103
            Top             =   2040
            Width           =   2100
         End
         Begin VB.Label memory 
            AutoSize        =   -1  'True
            Caption         =   "mem"
            Height          =   195
            Left            =   0
            TabIndex        =   89
            Top             =   1680
            Width           =   330
         End
         Begin VB.Label lcpu 
            AutoSize        =   -1  'True
            Caption         =   "cpu%"
            Height          =   195
            Left            =   0
            TabIndex        =   87
            Top             =   1440
            Width           =   390
         End
         Begin VB.Label ltime 
            AutoSize        =   -1  'True
            Caption         =   "time"
            Height          =   195
            Left            =   0
            TabIndex        =   86
            Top             =   1920
            Width           =   285
         End
         Begin VB.Label fan3 
            AutoSize        =   -1  'True
            Caption         =   "fan3"
            Height          =   195
            Left            =   0
            TabIndex        =   85
            Top             =   1200
            Width           =   315
         End
         Begin VB.Label fan2 
            AutoSize        =   -1  'True
            Caption         =   "fan2"
            Height          =   195
            Left            =   0
            TabIndex        =   84
            Top             =   960
            Width           =   315
         End
         Begin VB.Label fan1 
            AutoSize        =   -1  'True
            Caption         =   "fan1"
            Height          =   195
            Left            =   0
            TabIndex        =   83
            Top             =   720
            Width           =   315
         End
         Begin VB.Label temp3 
            AutoSize        =   -1  'True
            Caption         =   "temp3"
            Height          =   195
            Left            =   0
            TabIndex        =   82
            Top             =   480
            Width           =   435
         End
         Begin VB.Label temp2 
            AutoSize        =   -1  'True
            Caption         =   "temp2"
            Height          =   195
            Left            =   0
            TabIndex        =   81
            Top             =   240
            Width           =   435
         End
         Begin VB.Label temp1 
            AutoSize        =   -1  'True
            Caption         =   "temp1"
            Height          =   195
            Left            =   0
            TabIndex        =   80
            Top             =   0
            Width           =   435
         End
      End
      Begin VB.OptionButton custo 
         Caption         =   "Manual"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -74880
         TabIndex        =   30
         Top             =   1560
         Width           =   1095
      End
      Begin VB.OptionButton stan 
         Caption         =   "Stantard"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -74880
         TabIndex        =   28
         Top             =   1200
         Value           =   -1  'True
         Width           =   1095
      End
      Begin VB.CheckBox onlywith 
         Alignment       =   1  'Right Justify
         Caption         =   "Close only with Esc"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   18
         Top             =   1560
         Width           =   3495
      End
      Begin VB.ComboBox fontnamec 
         Height          =   315
         ItemData        =   "sett.frx":03B2
         Left            =   -73680
         List            =   "sett.frx":03B4
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   840
         Width           =   3135
      End
      Begin VB.ComboBox Combo2 
         Height          =   315
         ItemData        =   "sett.frx":03B6
         Left            =   -73680
         List            =   "sett.frx":03D2
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   1800
         Width           =   855
      End
      Begin VB.ComboBox backcol 
         Height          =   315
         ItemData        =   "sett.frx":040B
         Left            =   -73680
         List            =   "sett.frx":0427
         Style           =   2  'Dropdown List
         TabIndex        =   14
         Top             =   2280
         Width           =   855
      End
      Begin VB.ComboBox Combo3 
         Height          =   315
         ItemData        =   "sett.frx":0460
         Left            =   -73680
         List            =   "sett.frx":0494
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   1320
         Width           =   855
      End
      Begin VB.CheckBox Check8 
         Caption         =   "Transparent back style"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -74880
         TabIndex        =   15
         Top             =   4680
         Width           =   2655
      End
      Begin VB.CheckBox Check13 
         Caption         =   "Monitor"
         Height          =   255
         Left            =   -72600
         TabIndex        =   27
         Top             =   4200
         Width           =   855
      End
      Begin VB.CheckBox Check12 
         Caption         =   "Monitor"
         Height          =   255
         Left            =   -72600
         TabIndex        =   26
         Top             =   3600
         Width           =   855
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         ItemData        =   "sett.frx":04D6
         Left            =   -72720
         List            =   "sett.frx":04F2
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   4560
         Width           =   1815
      End
      Begin VB.CheckBox Check10 
         Caption         =   "Monitor"
         Height          =   255
         Left            =   -72600
         TabIndex        =   25
         Top             =   3000
         Width           =   855
      End
      Begin VB.Frame Frame3 
         Caption         =   "Temperatures"
         Height          =   1095
         Left            =   -74880
         TabIndex        =   51
         Top             =   720
         Width           =   4335
         Begin VB.CheckBox Check3 
            Caption         =   "Monitor"
            Height          =   255
            Left            =   2160
            TabIndex        =   21
            Top             =   720
            Width           =   855
         End
         Begin VB.CheckBox Check2 
            Caption         =   "Monitor"
            Height          =   255
            Left            =   2160
            TabIndex        =   20
            Top             =   480
            Width           =   855
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Monitor"
            Height          =   255
            Left            =   2160
            TabIndex        =   19
            Top             =   240
            Width           =   855
         End
         Begin VB.TextBox Text3 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   285
            Left            =   1200
            TabIndex        =   54
            Text            =   "temp3"
            Top             =   720
            Width           =   855
         End
         Begin VB.TextBox Text2 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   285
            Left            =   1200
            TabIndex        =   53
            Text            =   "temp2"
            Top             =   480
            Width           =   855
         End
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   285
            Left            =   1200
            TabIndex        =   52
            Text            =   "temp1"
            Top             =   240
            Width           =   855
         End
         Begin VB.Label val3 
            AutoSize        =   -1  'True
            Caption         =   "0"
            Height          =   195
            Left            =   3600
            TabIndex        =   63
            Top             =   720
            Width           =   90
         End
         Begin VB.Label val2 
            AutoSize        =   -1  'True
            Caption         =   "0"
            Height          =   195
            Left            =   3600
            TabIndex        =   62
            Top             =   480
            Width           =   90
         End
         Begin VB.Label val1 
            AutoSize        =   -1  'True
            Caption         =   "0"
            Height          =   195
            Left            =   3600
            TabIndex        =   61
            Top             =   240
            Width           =   90
         End
         Begin VB.Label Label4 
            Caption         =   "Value:"
            Height          =   255
            Index           =   2
            Left            =   3120
            TabIndex        =   60
            Top             =   480
            Width           =   495
         End
         Begin VB.Label Label4 
            Caption         =   "Value:"
            Height          =   255
            Index           =   1
            Left            =   3120
            TabIndex        =   59
            Top             =   720
            Width           =   495
         End
         Begin VB.Label Label4 
            Caption         =   "Value:"
            Height          =   255
            Index           =   0
            Left            =   3120
            TabIndex        =   58
            Top             =   240
            Width           =   495
         End
         Begin VB.Label Label3 
            Caption         =   "Temp3 name:"
            Height          =   255
            Left            =   120
            TabIndex        =   57
            Top             =   720
            Width           =   975
         End
         Begin VB.Label Label2 
            Caption         =   "Temp2 name:"
            Height          =   255
            Left            =   120
            TabIndex        =   56
            Top             =   480
            Width           =   975
         End
         Begin VB.Label Label1 
            Caption         =   "Temp1 name:"
            Height          =   255
            Left            =   120
            TabIndex        =   55
            Top             =   240
            Width           =   975
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Fan sensors"
         Height          =   1095
         Left            =   -74880
         TabIndex        =   38
         Top             =   1800
         Width           =   4335
         Begin VB.CheckBox Check6 
            Caption         =   "Monitor"
            Height          =   255
            Left            =   2160
            TabIndex        =   24
            Top             =   720
            Width           =   855
         End
         Begin VB.CheckBox Check5 
            Caption         =   "Monitor"
            Height          =   255
            Left            =   2160
            TabIndex        =   23
            Top             =   480
            Width           =   855
         End
         Begin VB.CheckBox Check4 
            Caption         =   "Monitor"
            Height          =   255
            Left            =   2160
            TabIndex        =   22
            Top             =   240
            Width           =   855
         End
         Begin VB.TextBox Text6 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   285
            Left            =   1200
            TabIndex        =   41
            Text            =   "fan3"
            Top             =   720
            Width           =   855
         End
         Begin VB.TextBox Text5 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   285
            Left            =   1200
            TabIndex        =   40
            Text            =   "fan2"
            Top             =   480
            Width           =   855
         End
         Begin VB.TextBox Text4 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   285
            Left            =   1200
            TabIndex        =   39
            Text            =   "fan1"
            Top             =   240
            Width           =   855
         End
         Begin VB.Label val6 
            AutoSize        =   -1  'True
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "MS Serif"
               Size            =   8.25
               Charset         =   161
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   3600
            TabIndex        =   50
            Top             =   720
            Width           =   75
         End
         Begin VB.Label val5 
            AutoSize        =   -1  'True
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "MS Serif"
               Size            =   8.25
               Charset         =   161
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   3600
            TabIndex        =   49
            Top             =   480
            Width           =   75
         End
         Begin VB.Label val4 
            AutoSize        =   -1  'True
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "MS Serif"
               Size            =   8.25
               Charset         =   161
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   3600
            TabIndex        =   48
            Top             =   240
            Width           =   75
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Value:"
            Height          =   195
            Index           =   5
            Left            =   3120
            TabIndex        =   47
            Top             =   720
            Width           =   450
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Value:"
            Height          =   195
            Index           =   4
            Left            =   3120
            TabIndex        =   46
            Top             =   480
            Width           =   450
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Value:"
            Height          =   195
            Index           =   3
            Left            =   3120
            TabIndex        =   45
            Top             =   240
            Width           =   450
         End
         Begin VB.Label Label8 
            Caption         =   "Fan2 name:"
            Height          =   255
            Index           =   2
            Left            =   120
            TabIndex        =   44
            Top             =   480
            Width           =   855
         End
         Begin VB.Label Label8 
            Caption         =   "Fan3 name:"
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   43
            Top             =   720
            Width           =   855
         End
         Begin VB.Label Label8 
            Caption         =   "Fan1 name:"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   42
            Top             =   240
            Width           =   855
         End
      End
      Begin VB.CheckBox Check7 
         Caption         =   "Enable picture show"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   161
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -74760
         TabIndex        =   1
         Top             =   720
         Width           =   2415
      End
      Begin VB.Frame Frame2 
         Caption         =   "Backround picture"
         Height          =   3255
         Left            =   -74880
         TabIndex        =   32
         Top             =   1080
         Width           =   4335
         Begin VB.TextBox br 
            Height          =   285
            Left            =   120
            Locked          =   -1  'True
            TabIndex        =   34
            Top             =   1440
            Width           =   2775
         End
         Begin VB.CommandButton bro 
            Caption         =   "Picture select"
            Height          =   255
            Left            =   3000
            TabIndex        =   5
            Top             =   1440
            Width           =   1215
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Backround picture"
            Height          =   255
            Left            =   120
            TabIndex        =   4
            Top             =   840
            Width           =   2175
         End
         Begin VB.OptionButton Option2 
            Caption         =   "Slide show"
            Height          =   255
            Left            =   120
            TabIndex        =   6
            Top             =   1920
            Width           =   1575
         End
         Begin VB.TextBox sli 
            Height          =   285
            Left            =   120
            Locked          =   -1  'True
            TabIndex        =   33
            Top             =   2520
            Width           =   2775
         End
         Begin VB.CommandButton slid 
            Caption         =   "Folder select"
            Height          =   255
            Left            =   3000
            TabIndex        =   8
            Top             =   2520
            Width           =   1215
         End
         Begin VB.TextBox delay 
            BeginProperty Font 
               Name            =   "MS Serif"
               Size            =   6.75
               Charset         =   161
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   1920
            MaxLength       =   6
            TabIndex        =   7
            Text            =   "4000"
            Top             =   2280
            Width           =   975
         End
         Begin VB.CheckBox stre 
            Caption         =   "Stretch the picture to fit in window"
            Height          =   255
            Left            =   120
            TabIndex        =   2
            Top             =   240
            Width           =   2775
         End
         Begin VB.CheckBox cntr 
            Caption         =   "Center picture position"
            Height          =   255
            Left            =   120
            TabIndex        =   3
            Top             =   480
            Width           =   2415
         End
         Begin VB.CheckBox Check9 
            Caption         =   "Display file name"
            Height          =   255
            Left            =   120
            TabIndex        =   9
            Top             =   2880
            Width           =   1575
         End
         Begin VB.Label Label5 
            Caption         =   "Picture filename:"
            Height          =   255
            Left            =   120
            TabIndex        =   37
            Top             =   1200
            Width           =   1215
         End
         Begin VB.Label Label6 
            Caption         =   "Pictures folder:"
            Height          =   255
            Left            =   120
            TabIndex        =   36
            Top             =   2280
            Width           =   1095
         End
         Begin VB.Label Label7 
            Caption         =   "Delay:(msec)"
            Height          =   255
            Left            =   1920
            TabIndex        =   35
            Top             =   2040
            Width           =   975
         End
      End
      Begin VB.Label Label34 
         AutoSize        =   -1  'True
         Caption         =   "spaces between sensorX: and value"
         Height          =   195
         Left            =   1440
         TabIndex        =   129
         Top             =   2520
         Width           =   2580
      End
      Begin VB.Label spaces 
         AutoSize        =   -1  'True
         Caption         =   "0"
         Height          =   195
         Left            =   1200
         TabIndex        =   128
         Top             =   2530
         Width           =   90
      End
      Begin VB.Label Label32 
         Caption         =   "Leave "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   126
         Top             =   2520
         Width           =   615
      End
      Begin VB.Label Label23 
         Caption         =   "Refresh every:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   125
         Top             =   2040
         Width           =   1335
      End
      Begin VB.Label Label24 
         Caption         =   "msec"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2400
         TabIndex        =   124
         Top             =   2040
         Width           =   615
      End
      Begin VB.Label Label26 
         Caption         =   "Min Value:"
         Height          =   255
         Index           =   1
         Left            =   -72600
         TabIndex        =   117
         Top             =   2400
         Width           =   855
      End
      Begin VB.Label Label25 
         Caption         =   "Max Value:"
         Height          =   255
         Index           =   1
         Left            =   -72600
         TabIndex        =   116
         Top             =   1920
         Width           =   855
      End
      Begin VB.Label Label28 
         Caption         =   "Fans"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   161
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -72000
         TabIndex        =   115
         Top             =   1440
         Width           =   615
      End
      Begin VB.Label Label27 
         Caption         =   "Temperatures"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   161
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -74640
         TabIndex        =   114
         Top             =   1440
         Width           =   1575
      End
      Begin VB.Label Label26 
         Caption         =   "Min Value:"
         Height          =   255
         Index           =   0
         Left            =   -74640
         TabIndex        =   107
         Top             =   2400
         Width           =   855
      End
      Begin VB.Label Label25 
         Caption         =   "Max Value:"
         Height          =   255
         Index           =   0
         Left            =   -74640
         TabIndex        =   106
         Top             =   1920
         Width           =   855
      End
      Begin VB.Label versionMBM 
         Caption         =   "MBM Saver"
         Height          =   435
         Left            =   120
         TabIndex        =   104
         Top             =   4920
         Width           =   4320
      End
      Begin VB.Label Label22 
         Caption         =   "Windows up time"
         Height          =   255
         Left            =   -74880
         TabIndex        =   102
         Top             =   4800
         Width           =   1455
      End
      Begin VB.Label winuptime 
         AutoSize        =   -1  'True
         Height          =   195
         Left            =   -74880
         TabIndex        =   101
         Top             =   5040
         Width           =   45
      End
      Begin VB.Line Line1 
         Index           =   2
         X1              =   -74880
         X2              =   -70680
         Y1              =   4680
         Y2              =   4680
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         Caption         =   "Memory:"
         Height          =   195
         Left            =   -73320
         TabIndex        =   90
         Top             =   4440
         Width           =   600
      End
      Begin VB.Label Label20 
         Caption         =   "Position:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   161
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -74880
         TabIndex        =   78
         Top             =   840
         Width           =   975
      End
      Begin VB.Label testt 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "ABCDE"
         Height          =   1335
         Left            =   -74880
         TabIndex        =   77
         Top             =   3120
         Width           =   4215
      End
      Begin VB.Label Label19 
         Caption         =   "Example:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -74880
         TabIndex        =   76
         Top             =   2760
         Width           =   975
      End
      Begin VB.Label Label18 
         Caption         =   "Font"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -74880
         TabIndex        =   75
         Top             =   840
         Width           =   495
      End
      Begin VB.Label Label13 
         Caption         =   "Font color"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -74880
         TabIndex        =   74
         Top             =   1800
         Width           =   975
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "Back color"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -74880
         TabIndex        =   73
         Top             =   2280
         Width           =   975
      End
      Begin VB.Label Label14 
         Caption         =   "Font size"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -74880
         TabIndex        =   72
         Top             =   1320
         Width           =   855
      End
      Begin VB.Line Line1 
         Index           =   0
         X1              =   -76560
         X2              =   -70560
         Y1              =   4080
         Y2              =   4080
      End
      Begin VB.Line Line1 
         Index           =   1
         X1              =   -76560
         X2              =   -70560
         Y1              =   3480
         Y2              =   3480
      End
      Begin VB.Label winamptitle 
         AutoSize        =   -1  'True
         Height          =   195
         Left            =   -74400
         TabIndex        =   71
         Top             =   3240
         Width           =   45
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "CPU:"
         Height          =   195
         Index           =   3
         Left            =   -74880
         TabIndex        =   70
         Top             =   4440
         Width           =   375
      End
      Begin VB.Label Label17 
         Caption         =   "CPU usage and memory load"
         Height          =   255
         Left            =   -74880
         TabIndex        =   69
         Top             =   4200
         Width           =   2175
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "Value:"
         Height          =   195
         Index           =   2
         Left            =   -74880
         TabIndex        =   68
         Top             =   3840
         Width           =   450
      End
      Begin VB.Label Label16 
         Caption         =   "Time"
         Height          =   255
         Index           =   1
         Left            =   -74880
         TabIndex        =   67
         Top             =   3600
         Width           =   1095
      End
      Begin VB.Label Label16 
         Caption         =   "Value:"
         Height          =   255
         Index           =   0
         Left            =   -74880
         TabIndex        =   66
         Top             =   3240
         Width           =   495
      End
      Begin VB.Label Label15 
         Caption         =   "Winamp (artist-song and time)"
         Height          =   255
         Left            =   -74880
         TabIndex        =   65
         Top             =   3000
         Width           =   2175
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "Background color"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   161
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   -74880
         TabIndex        =   64
         Top             =   4560
         Width           =   1845
      End
   End
   Begin VB.Timer Timer1 
      Left            =   4560
      Top             =   4920
   End
   Begin VB.CommandButton Command2 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   2400
      TabIndex        =   29
      Top             =   5520
      Width           =   2175
   End
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   5520
      Width           =   2175
   End
End
Attribute VB_Name = "sett"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim curx, cury As Integer
Private QueryObject As Object

Private Sub backcol_Click()
'change the font preview backgroung color
testt.BackColor = clrr(backcol.Text)
Call chng
End Sub

Private Sub backcol_KeyDown(KeyCode As Integer, Shift As Integer)
'change the font preview backgroung color
testt.BackColor = clrr(backcol.Text)
Call chng
End Sub

Private Sub bro_Click()
'open the dialog to choose a picture
br.Text = DialogFile(Me.hwnd, 1, "Select a picture to load...", "", "jpg bmp gif" & Chr$(0) & "*.jpg;*.bmp;*.gif", pathh(br.Text), "")
End Sub
Private Function pathh(fullpath As String)
'format the path (c:\eexx\fotis\fotis.jpg to c:\eexx\fotis\)
Dim oo As Integer
For oo = Len(fullpath) To 1 Step -1
If Mid(fullpath, oo, 1) = "\" Then
pathh = Left$(fullpath, oo)
Exit Function
End If
Next oo
pathh = App.Path
End Function
Private Sub Check1_Click()
'temp1 enable
If Check1.value = 1 Then
 Text1.Enabled = True
Else
 val1.Caption = "0"
 Text1.Enabled = False
End If
temp1.Visible = Check1.value
End Sub
Private Sub Check11_Click()
'transparent enable
HScroll1.Enabled = Check11.value
Label10.Enabled = Check11.value
Label11.Enabled = Check11.value
End Sub
Private Sub Check12_Click()
'time label enable
ltime.Visible = Check12.value
End Sub
Private Sub Check13_Click()
'cpu and memory label enable
lcpu.Visible = Check13.value
memory.Visible = Check13.value
End Sub
Private Sub Check2_Click()
'temp2 enable
If Check2.value = 1 Then
 Text2.Enabled = True
Else
 val2.Caption = "0"
 Text2.Enabled = False
End If
temp2.Visible = Check2.value
End Sub
Private Sub Check3_Click()
'temp3 enable
If Check3.value = 1 Then
 Text3.Enabled = True
Else
 val3.Caption = "0"
 Text3.Enabled = False
End If
temp3.Visible = Check3.value
End Sub
Private Sub Check4_Click()
'fan1 enable
If Check4.value = 1 Then
 Text4.Enabled = True
Else
 val4.Caption = "0"
 Text4.Enabled = False
End If
fan1.Visible = Check4.value
End Sub
Private Sub Check5_Click()
'fan2 enable
If Check5.value = 1 Then
 Text5.Enabled = True
Else
 val5.Caption = "0"
 Text5.Enabled = False
End If
fan2.Visible = Check5.value
End Sub
Private Sub Check6_Click()
'fan3 enable
If Check6.value = 1 Then
 Text6.Enabled = True
Else
 val6.Caption = "0"
 Text6.Enabled = False
End If
fan3.Visible = Check6.value
End Sub
Private Sub Check7_Click()
'background enable
br.Enabled = Check7.value
bro.Enabled = Check7.value
sli.Enabled = Check7.value
slid.Enabled = Check7.value
Option1.Enabled = Check7.value
Option2.Enabled = Check7.value
delay.Enabled = Check7.value
If Option1.Enabled = True Then Option1_Click
stre.Enabled = Check7.value
cntr.Enabled = Check7.value
Check9.Enabled = Check7.value
End Sub

Private Sub Check8_Click()
'transparent enable
If Check8.value = 1 Then
 Label9.Enabled = False
 backcol.Enabled = False
 testt.BackStyle = 0
Else
 Label9.Enabled = True
 backcol.Enabled = True
 testt.BackStyle = 1
 testt.BackColor = clrr(backcol.Text)
End If
chng
End Sub

Private Sub Combo1_Click()
'background color
Picture1.BackColor = clrr(Combo1.Text)
End Sub

Private Sub Combo1_KeyDown(KeyCode As Integer, Shift As Integer)
'background color
Picture1.BackColor = clrr(Combo1.Text)
End Sub

Private Sub Combo2_Click()
'font color
testt.ForeColor = clrr(Combo2.Text)
Call chng
End Sub

Function clrr(stringname As String) As Long
'change the string to color code
Select Case stringname
Dim clr As Long
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
clrr = clr
End Function

Private Sub Combo2_KeyDown(KeyCode As Integer, Shift As Integer)
'font color
testt.ForeColor = clrr(Combo2.Text)
Call chng
End Sub

Private Sub Combo3_Click()
'font size
testt.FontSize = Combo3.Text
Call chng
End Sub

Private Sub Combo3_KeyDown(KeyCode As Integer, Shift As Integer)
'font size
testt.FontSize = Combo3.Text
Call chng
End Sub

Private Sub Command1_Click()
'save settings to registry
SaveSetting "MBM Saver", "settings", "name1", Text1.Text
SaveSetting "MBM Saver", "settings", "name2", Text2.Text
SaveSetting "MBM Saver", "settings", "name3", Text3.Text
SaveSetting "MBM Saver", "settings", "name4", Text4.Text
SaveSetting "MBM Saver", "settings", "name5", Text5.Text
SaveSetting "MBM Saver", "settings", "name6", Text6.Text
SaveSetting "MBM Saver", "settings", "check1", Check1.value
SaveSetting "MBM Saver", "settings", "check2", Check2.value
SaveSetting "MBM Saver", "settings", "check3", Check3.value
SaveSetting "MBM Saver", "settings", "check4", Check4.value
SaveSetting "MBM Saver", "settings", "check5", Check5.value
SaveSetting "MBM Saver", "settings", "check6", Check6.value
SaveSetting "MBM Saver", "settings", "winuptm", winuptm.value
SaveSetting "MBM Saver", "settings", "combo1", Combo1.Text
SaveSetting "MBM Saver", "settings", "combo2", Combo2.Text
SaveSetting "MBM Saver", "settings", "combo3", Combo3.Text
SaveSetting "MBM Saver", "settings", "option1", Option1.value
SaveSetting "MBM Saver", "settings", "option2", Option2.value
SaveSetting "MBM Saver", "settings", "br", br.Text
SaveSetting "MBM Saver", "settings", "sli", sli.Text
SaveSetting "MBM Saver", "settings", "check7", Check7.value
SaveSetting "MBM Saver", "settings", "delay", delay.Text
SaveSetting "MBM Saver", "settings", "check8", Check8.value
SaveSetting "MBM Saver", "settings", "backcol", backcol.Text
SaveSetting "MBM Saver", "settings", "stre", stre.value
SaveSetting "MBM Saver", "settings", "cntr", cntr.value
SaveSetting "MBM Saver", "settings", "check9", Check9.value
SaveSetting "MBM Saver", "settings", "check10", Check10.value
SaveSetting "MBM Saver", "settings", "HScroll1", HScroll1.value
SaveSetting "MBM Saver", "settings", "check11", Check11.value
SaveSetting "MBM Saver", "settings", "fontnamec", fontnamec.Text
SaveSetting "MBM Saver", "settings", "celcius", Option3.value
SaveSetting "MBM Saver", "settings", "farhenheit", Option4.value
SaveSetting "MBM Saver", "settings", "check12", Check12.value
SaveSetting "MBM Saver", "settings", "check13", Check13.value
SaveSetting "MBM Saver", "settings", "stan", stan.value
SaveSetting "MBM Saver", "settings", "moveme", moveme.value
SaveSetting "MBM Saver", "settings", "temp1", temp1.Top
SaveSetting "MBM Saver", "settings", "temp2", temp2.Top
SaveSetting "MBM Saver", "settings", "temp3", temp3.Top
SaveSetting "MBM Saver", "settings", "fan1", fan1.Top
SaveSetting "MBM Saver", "settings", "fan2", fan2.Top
SaveSetting "MBM Saver", "settings", "fan3", fan3.Top
SaveSetting "MBM Saver", "settings", "ltime", ltime.Top
SaveSetting "MBM Saver", "settings", "lcpu", lcpu.Top
SaveSetting "MBM Saver", "settings", "wintime", wintime.Top
SaveSetting "MBM Saver", "settings", "memory", memory.Top
SaveSetting "MBM Saver", "settings", "temp1l", temp1.Left
SaveSetting "MBM Saver", "settings", "temp2l", temp2.Left
SaveSetting "MBM Saver", "settings", "temp3l", temp3.Left
SaveSetting "MBM Saver", "settings", "fan1l", fan1.Left
SaveSetting "MBM Saver", "settings", "fan2l", fan2.Left
SaveSetting "MBM Saver", "settings", "fan3l", fan3.Left
SaveSetting "MBM Saver", "settings", "ltimel", ltime.Left
SaveSetting "MBM Saver", "settings", "lcpul", lcpu.Left
SaveSetting "MBM Saver", "settings", "memoryl", memory.Left
SaveSetting "MBM Saver", "settings", "wintimel", wintime.Left
SaveSetting "MBM Saver", "settings", "history", historycheck.value
SaveSetting "MBM Saver", "settings", "fadein", fadein.value
SaveSetting "MBM Saver", "settings", "fadeout", fadeout.value
SaveSetting "MBM Saver", "settings", "onlywith", onlywith.value
SaveSetting "MBM Saver", "settings", "refreshtime", refreshtime.Text
SaveSetting "MBM Saver", "settings", "maxvalue", maxvalue.Text
SaveSetting "MBM Saver", "settings", "minvalue", minvalue.Text
SaveSetting "MBM Saver", "settings", "fmaxvalue", fmaxvalue.Text
SaveSetting "MBM Saver", "settings", "fminvalue", fminvalue.Text
SaveSetting "MBM Saver", "settings", "disaver", disaver.value
SaveSetting "MBM Saver", "settings", "dishi", dishi.value
SaveSetting "MBM Saver", "settings", "dislow", dislow.value
SaveSetting "MBM Saver", "settings", "spaces", spaces.Caption

Unload Me
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub custo_Click()
'the manual position button
If stan.value = True Then
 Picture1.Visible = False 'hide preview
Else
 Picture1.Visible = True 'show preview
End If
End Sub

Private Sub delay_LostFocus()
'slide show delay
On Error GoTo dfd
If (delay.Text > 50000) Or (delay.Text < 1001) Then GoTo dfd
Exit Sub
dfd:
MsgBox "Enter a value between 1001 and 50000", vbInformation, "Error"
delay.Text = "4000"
End Sub

Private Sub fan1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
'take coordinate of x and y
If Button = 1 Then
cury = y
curx = x
End If
End Sub

Private Sub fan1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
'move label to current x y
If Button = 1 Then
fan1.Left = fan1.Left + (x - curx)
fan1.Top = fan1.Top + (y - cury)
End If
End Sub

Private Sub fan2_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
'take coordinate of x and y
If Button = 1 Then
cury = y
curx = x
End If
End Sub

Private Sub fan2_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
'move label to current x y
If Button = 1 Then
fan2.Left = fan2.Left + (x - curx)
fan2.Top = fan2.Top + (y - cury)
End If
End Sub

Private Sub fan3_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
'take coordinate of x and y
If Button = 1 Then
cury = y
curx = x
End If
End Sub

Private Sub fan3_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
'move label to current x y
If Button = 1 Then
fan3.Left = fan3.Left + (x - curx)
fan3.Top = fan3.Top + (y - cury)
End If
End Sub

Private Sub fontnamec_Click()
'update preview
testt.Font = fontnamec.List(fontnamec.ListIndex)
testt.FontSize = Combo3.Text
End Sub

Private Sub fontnamec_KeyDown(KeyCode As Integer, Shift As Integer)
'update preview
testt.Font = fontnamec.List(fontnamec.ListIndex)
testt.FontSize = Combo3.Text
End Sub

Private Sub Form_Load()
'load settings
versionMBM.Caption = "MBM Saver v" & App.Major & "." & App.Minor & "." & App.Revision & " by Fotis P.  For any bugs or anything else please mail me at robot@mail.gr"
trans.Enabled = IsWinNT
FindWinamp
SetThreadPriority GetCurrentThread, THREAD_BASE_PRIORITY_MAX
    SetPriorityClass GetCurrentProcess, HIGH_PRIORITY_CLASS
    'Initialize our Query object
    If IsWinNT Then
        Set QueryObject = New clsCPUUsageNT
    Else
        Set QueryObject = New clsCPUUsage
    End If
    QueryObject.Initialize
Mnew = True
Dim a As Integer

For a = 0 To Screen.FontCount - 1
fontnamec.AddItem Screen.Fonts(a)
Next a
Text1.Text = GetSetting("MBM Saver", "settings", "name1", "temp1")
Text2.Text = GetSetting("MBM Saver", "settings", "name2", "temp2")
Text3.Text = GetSetting("MBM Saver", "settings", "name3", "temp3")
Text4.Text = GetSetting("MBM Saver", "settings", "name4", "fan1")
Text5.Text = GetSetting("MBM Saver", "settings", "name5", "fan2")
Text6.Text = GetSetting("MBM Saver", "settings", "name6", "fan3")
Check1.value = GetSetting("MBM Saver", "settings", "check1", "0")
Check2.value = GetSetting("MBM Saver", "settings", "check2", "0")
Check3.value = GetSetting("MBM Saver", "settings", "check3", "0")
Check4.value = GetSetting("MBM Saver", "settings", "check4", "0")
Check5.value = GetSetting("MBM Saver", "settings", "check5", "0")
Check6.value = GetSetting("MBM Saver", "settings", "check6", "0")
winuptm.value = GetSetting("MBM Saver", "settings", "winuptm", "0")
temp1.Visible = Check1.value
temp2.Visible = Check2.value
temp3.Visible = Check3.value
fan1.Visible = Check4.value
fan2.Visible = Check5.value
fan3.Visible = Check6.value
wintime.Visible = winuptm.value
Picture1.BackColor = clrr(Combo1.Text)

Combo1.Text = GetSetting("MBM Saver", "settings", "combo1", "Black")
Combo2.Text = GetSetting("MBM Saver", "settings", "combo2", "White")
Combo3.Text = GetSetting("MBM Saver", "settings", "combo3", "28")
Option1.value = GetSetting("MBM Saver", "settings", "option1", "true")
Option2.value = Not (Option1.value)

br.Text = GetSetting("MBM Saver", "settings", "br", "")
sli.Text = GetSetting("MBM Saver", "settings", "sli", "")
Check7.value = GetSetting("MBM Saver", "settings", "check7", "0")
delay.Text = GetSetting("MBM Saver", "settings", "delay", "4000")
Check8.value = GetSetting("MBM Saver", "settings", "check8", "0")
backcol.Text = GetSetting("MBM Saver", "settings", "backcol", "Black")
stre.value = GetSetting("MBM Saver", "settings", "stre", "0")
cntr.value = GetSetting("MBM Saver", "settings", "cntr", "0")
Check9.value = GetSetting("MBM Saver", "settings", "check9", "0")
Check10.value = GetSetting("MBM Saver", "settings", "check10", "0")
HScroll1.value = GetSetting("MBM Saver", "settings", "HScroll1", "251")
fontnamec.Text = GetSetting("MBM Saver", "settings", "fontnamec", fontnamec.List(1))
Option3.value = GetSetting("MBM Saver", "settings", "celcius", "True")
Option4.value = GetSetting("MBM Saver", "settings", "farhenheit", "False")
Check12.value = GetSetting("MBM Saver", "settings", "check12", "0")
Check13.value = GetSetting("MBM Saver", "settings", "check13", "0")
lcpu.Visible = Check13.value
memory.Visible = Check13.value
ltime.Visible = Check12.value
stan.value = GetSetting("MBM Saver", "settings", "stan", "1")
moveme.value = GetSetting("MBM Saver", "settings", "moveme", "0")
temp1.Left = GetSetting("MBM Saver", "settings", "temp1l", "0")
temp2.Left = GetSetting("MBM Saver", "settings", "temp2l", "0")
temp3.Left = GetSetting("MBM Saver", "settings", "temp3l", "0")
fan1.Left = GetSetting("MBM Saver", "settings", "fan1l", "0")
fan2.Left = GetSetting("MBM Saver", "settings", "fan2l", "0")
fan3.Left = GetSetting("MBM Saver", "settings", "fan3l", "0")
lcpu.Left = GetSetting("MBM Saver", "settings", "lcpul", "0")
ltime.Left = GetSetting("MBM Saver", "settings", "ltimel", "0")
memory.Left = GetSetting("MBM Saver", "settings", "memoryl", "0")
wintime.Left = GetSetting("MBM Saver", "settings", "wintimel", "600")

temp1.Top = GetSetting("MBM Saver", "settings", "temp1", "0")
temp2.Top = GetSetting("MBM Saver", "settings", "temp2", "240")
temp3.Top = GetSetting("MBM Saver", "settings", "temp3", "480")
fan1.Top = GetSetting("MBM Saver", "settings", "fan1", "720")
fan2.Top = GetSetting("MBM Saver", "settings", "fan2", "960")
fan3.Top = GetSetting("MBM Saver", "settings", "fan3", "1200")
lcpu.Top = GetSetting("MBM Saver", "settings", "lcpu", "1440")
ltime.Top = GetSetting("MBM Saver", "settings", "ltime", "1680")
memory.Top = GetSetting("MBM Saver", "settings", "memory", "1920")
wintime.Top = GetSetting("MBM Saver", "settings", "wintime", "2040")
historycheck.value = GetSetting("MBM Saver", "settings", "history", "0")
fadein.value = GetSetting("MBM Saver", "settings", "fadein", "0")
fadeout.value = GetSetting("MBM Saver", "settings", "fadeout", "0")
onlywith.value = GetSetting("MBM Saver", "settings", "onlywith", "0")
maxvalue.Text = GetSetting("MBM Saver", "settings", "maxvalue", "80")
minvalue.Text = GetSetting("MBM Saver", "settings", "minvalue", "-20")
fmaxvalue.Text = GetSetting("MBM Saver", "settings", "fmaxvalue", "7500")
fminvalue.Text = GetSetting("MBM Saver", "settings", "fminvalue", "1000")
refreshtime.Text = GetSetting("MBM Saver", "settings", "refreshtime", "1000")
disaver.value = GetSetting("MBM Saver", "settings", "disaver", "0")
dislow.value = GetSetting("MBM Saver", "settings", "dislow", "0")
dishi.value = GetSetting("MBM Saver", "settings", "dishi", "0")
Option1_Click
spaces.Caption = 10 - GetSetting("MBM Saver", "settings", "spaces", "0")
VScroll1.value = spaces.Caption

historycheck_Click
Call chng
custo.value = Not (stan.value)
If stan.value = True Then
Picture1.Visible = False
Else
Picture1.Visible = True
End If
Label10.Caption = Int(100 - 100 / 255 * HScroll1.value)
Check8_Click
Check11.value = GetSetting("MBM Saver", "settings", "check11", "0")
HScroll1.Enabled = Check11.value
Label10.Enabled = Check11.value
Label11.Enabled = Check11.value

br.Enabled = Option1.value
bro.Enabled = Option1.value
sli.Enabled = Option2.value
slid.Enabled = Option2.value
Check9.Enabled = Option2.value

Check7_Click
myData = MBM_GetData
If myData.sdVersion = 0 Then Mnew = False
If Option1.Enabled = True Then Option1_Click
 If Mnew = True Then
  If Left$(myData.sdVersion, 1) = "5" Then
   Me.Caption = "MBM VERSION:" & telia(myData.sdVersion)
   Timer1.Interval = 1000
   Exit Sub
  End If
 Else
  If Left$(MBM_GetInfo.SMBM_Version, 1) = "5" Then
   Me.Caption = "MBM VERSION:" & MBM_GetInfo.SMBM_Version
   Timer1.Interval = 1000
   Exit Sub
  End If
 End If
Me.Caption = "MBM 5.x is not running..."
End Sub
Function telia(vers) As String
'format the version
Dim j As Integer
Dim aaa As String
For j = 1 To Len(vers)
aaa = Mid(vers, j, 1)
If j = Len(vers) Then
telia = telia & aaa
Else
telia = telia & aaa & "."
End If
Next j
End Function

Private Sub historycheck_Click()
'graphical history
Frame1.Enabled = historycheck.value
Label25(0).Enabled = historycheck.value
Label26(0).Enabled = historycheck.value
Label25(1).Enabled = historycheck.value
Label26(1).Enabled = historycheck.value
minvalue.Enabled = historycheck.value
maxvalue.Enabled = historycheck.value
fminvalue.Enabled = historycheck.value
fmaxvalue.Enabled = historycheck.value


End Sub

Private Sub HScroll1_Change()
'transparent value
Label10.Caption = Int(100 - 100 / 255 * HScroll1.value)
End Sub

Private Sub HScroll1_Scroll()
'transparent value
Label10.Caption = Int(100 - 100 / 255 * HScroll1.value)
End Sub

Private Sub lcpu_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
'take coordinate of x and y
If Button = 1 Then
cury = y
curx = x
End If
End Sub

Private Sub lcpu_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
'move label to current x y
If Button = 1 Then
lcpu.Left = lcpu.Left + (x - curx)
lcpu.Top = lcpu.Top + (y - cury)
End If
End Sub

Private Sub ltime_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
'take coordinate of x and y
If Button = 1 Then
cury = y
curx = x
End If
End Sub

Private Sub ltime_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
'move label to current x y
If Button = 1 Then
ltime.Left = ltime.Left + (x - curx)
ltime.Top = ltime.Top + (y - cury)
End If
End Sub

Private Sub memory_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
'take coordinate of x and y
If Button = 1 Then
cury = y
curx = x
End If
End Sub

Private Sub memory_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
'move label to current x y
If Button = 1 Then
memory.Left = memory.Left + (x - curx)
memory.Top = memory.Top + (y - cury)
End If
End Sub

Private Sub Option1_Click()
'backgroung picture enable
br.Enabled = Option1.value
bro.Enabled = Option1.value
sli.Enabled = Not (Option1.value)
slid.Enabled = Not (Option1.value)
delay.Enabled = Not (Option1.value)
Check9.Enabled = Not (Option1.value)
End Sub

Private Sub Option2_Click()
Option1_Click
End Sub

Private Sub resetbut_Click()
'reset button
wintime.Top = 2040
temp1.Top = 0
temp2.Top = 240
temp3.Top = 480
fan1.Top = 720
fan2.Top = 960
fan3.Top = 1200
lcpu.Top = 1440
memory.Top = 1680
ltime.Top = 1920
temp1.Left = 0
temp2.Left = 0
temp3.Left = 0
fan1.Left = 0
fan2.Left = 0
fan3.Left = 0
lcpu.Left = 0
memory.Left = 0
ltime.Left = 0
wintime.Left = 600
End Sub

Private Sub slid_Click()
'slide show pictures folder
sli.Text = BrowseFolder(Me.hwnd, "Select a folder with photos")
End Sub

Private Sub stan_Click()
'hide/show preview window
If stan.value = True Then
 Picture1.Visible = False
Else
 Picture1.Visible = True
End If
End Sub

Private Sub temp1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
'take coordinate of x and y
If Button = 1 Then
cury = y
curx = x
End If
End Sub

Private Sub temp1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
'move label to current x y
If Button = 1 Then
temp1.Left = temp1.Left + (x - curx)
temp1.Top = temp1.Top + (y - cury)
End If
End Sub

Private Sub temp2_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
'take coordinate of x and y
If Button = 1 Then
cury = y
curx = x
End If
End Sub

Private Sub temp2_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
'move label to current x y
If Button = 1 Then
temp2.Left = temp2.Left + (x - curx)
temp2.Top = temp2.Top + (y - cury)
End If
End Sub

Private Sub temp3_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
'take coordinate of x and y
If Button = 1 Then
cury = y
curx = x
End If
End Sub

Private Sub temp3_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
'move label to current x y
If Button = 1 Then
temp3.Left = temp3.Left + (x - curx)
temp3.Top = temp3.Top + (y - cury)
End If
End Sub

Private Sub Timer1_Timer()
'refresh the labels
Dim CorF As Byte
If Option3.value = True Then
CorF = 2
Else
CorF = 1
End If

myData = MBM_GetData
If Check1.value = 1 Then
If Mnew = True Then
val1.Caption = CtoF(myData.sdSensor(0).ssCurrent, CorF)
Else
val1.Caption = CtoF(MBM_GetData1.STemperature(1), CorF)
End If
End If
If Check2.value = 1 Then
If Mnew = True Then
val2.Caption = CtoF(myData.sdSensor(1).ssCurrent, CorF)
Else
val2.Caption = CtoF(MBM_GetData1.STemperature(2), CorF)
End If
End If

If Check3.value = 1 Then
If Mnew = True Then
val3.Caption = CtoF(myData.sdSensor(2).ssCurrent, CorF)
Else
val3.Caption = CtoF(MBM_GetData1.STemperature(3), CorF)
End If
End If
Dim aa As Integer

If Check4.value = 1 Then
If Mnew = True Then
 aa = myData.sdSensor(48).ssCurrent
 If aa <> 255 Then val4.Caption = aa & " rpm"
Else
 aa = MBM_GetData1.SFan(1)
 If aa <> 255 Then val4.Caption = aa & " rpm"
End If
End If
If Check5.value = 1 Then
If Mnew = True Then
 aa = myData.sdSensor(49).ssCurrent
 If aa <> 255 Then val5.Caption = aa & " rpm"
Else
 aa = MBM_GetData1.SFan(2)
 If aa <> 255 Then val5.Caption = aa & " rpm"
End If
End If

If Check6.value = 1 Then
If Mnew = True Then
 aa = myData.sdSensor(50).ssCurrent
 If aa <> 255 Then val6.Caption = aa & " rpm"
Else
 aa = MBM_GetData1.SFan(3)
 If aa <> 255 Then val6.Caption = aa & " rpm"
End If
End If



End Sub
Private Sub timer2_timer()
Dim Ret As Long
    'query the CPU usage
    Ret = QueryObject.Query
'cpu and memory usage
If Check13.value = 1 Then
 Label16(3).Caption = "Value: " & CStr(Ret) + "%"
 Label21.Caption = "Memory: " & GetMemoryInfo
End If
'windows up time
If winuptm.value = 1 Then
winuptime.Caption = FormatCount(GetTickCount, DaysHoursMinutesSeconds)
Else
winuptime.Caption = ""
End If

If Check10.value = 1 Then
 If FindWindow("STUDIO", vbNullString) <> 0 Then
  'winamp3
  winamptitle.Caption = title(GetText(FindWindow("STUDIO", vbNullString)))
 Else
 'winamp 2.x
  winamptitle.Caption = WM_GET(SONG_TITLE)
 End If
End If
If Check12.value = 1 Then Label16(2).Caption = "Value:" & Time
End Sub
Private Sub chng()
'update the preview labels
Dim fontt As Boolean
On Error Resume Next

If Check8.value = 1 Then
 fontt = 0
Else
 fontt = 1
End If

fan1.ForeColor = clrr(Combo2.Text)
fan1.BackColor = clrr(backcol.Text)
fan1.FontSize = Int(Combo3.Text / 3)
fan1.Font = fontnamec.Text
fan1.BackStyle = fontt

fan2.ForeColor = clrr(Combo2.Text)
fan2.BackColor = clrr(backcol.Text)
fan2.FontSize = Int(Combo3.Text / 3)
fan2.Font = fontnamec.Text
fan2.BackStyle = fontt

fan3.ForeColor = clrr(Combo2.Text)
fan3.BackColor = clrr(backcol.Text)
fan3.FontSize = Int(Combo3.Text / 3)
fan3.Font = fontnamec.Text
fan3.BackStyle = fontt

temp1.ForeColor = clrr(Combo2.Text)
temp1.BackColor = clrr(backcol.Text)
temp1.FontSize = Int(Combo3.Text / 3)
temp1.Font = fontnamec.Text
temp1.BackStyle = fontt

temp2.ForeColor = clrr(Combo2.Text)
temp2.BackColor = clrr(backcol.Text)
temp2.FontSize = Int(Combo3.Text / 3)
temp2.Font = fontnamec.Text
temp2.BackStyle = fontt

temp3.ForeColor = clrr(Combo2.Text)
temp3.BackColor = clrr(backcol.Text)
temp3.FontSize = Int(Combo3.Text / 3)
temp3.Font = fontnamec.Text
temp3.BackStyle = fontt

lcpu.ForeColor = clrr(Combo2.Text)
lcpu.BackColor = clrr(backcol.Text)
lcpu.FontSize = Int(Combo3.Text / 3)
lcpu.Font = fontnamec.Text
lcpu.BackStyle = fontt

ltime.ForeColor = clrr(Combo2.Text)
ltime.BackColor = clrr(backcol.Text)
ltime.FontSize = Int(Combo3.Text / 3)
ltime.Font = fontnamec.Text
ltime.BackStyle = fontt

memory.ForeColor = clrr(Combo2.Text)
memory.BackColor = clrr(backcol.Text)
memory.FontSize = Int(Combo3.Text / 3)
memory.Font = fontnamec.Text
memory.BackStyle = fontt

wintime.ForeColor = clrr(Combo2.Text)
wintime.BackColor = clrr(backcol.Text)
wintime.FontSize = Int(Combo3.Text / 3)
wintime.Font = fontnamec.Text
wintime.BackStyle = fontt
End Sub
Private Function CtoF(ByVal mTemp As String, mCelFah As Byte) As String
Dim f As Integer, c As Integer
   Select Case mCelFah
  
    Case 1
         CtoF = (mTemp * 1.8) + 32 & "°F"
      Case 2
      
   CtoF = mTemp & "°C"
   
End Select
End Function

Private Sub VScroll1_Change()
spaces.Caption = 10 - VScroll1.value

End Sub

Private Sub wintime_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
'take coordinate of x and y
If Button = 1 Then
cury = y
curx = x
End If
End Sub

Private Sub wintime_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
'move label to current x y
If Button = 1 Then
wintime.Left = wintime.Left + (x - curx)
wintime.Top = wintime.Top + (y - cury)
End If
End Sub

Private Sub winuptm_Click()
wintime.Visible = winuptm.value

End Sub
