Attribute VB_Name = "CommDlgs"
'// From planet-source-code
'// Common Dialogs Module
'//
'// Description:
'// Provides wrapper functions into the various Windows OS common dialog boxes
'//

Option Explicit

'//
'// Structures
'//

Private Type OPENFILENAME
    lStructSize As Long
    hwnd As Long
    hInstance As Long
    lpstrFilter As String
    lpstrCustomFilter As String
    nMaxCustFilter As Long
    nFilterIndex As Long
    lpstrFile As String
    nMaxFile As Long
    lpstrFileTitle As String
    nMaxFileTitle As Long
    lpstrInitialDir As String
    lpstrTitle As String
    Flags As Long
    nFileOffset As Integer
    nFileExtension As Integer
    lpstrDefExt As String
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
End Type

Private Type SHITEMID
    cb As Long
    abID As Byte
End Type

Private Type ITEMIDLIST
    mkid As SHITEMID
End Type

Private Type BROWSEINFO
    hOwner As Long
    pidlRoot As Long
    pszDisplayName As String
    lpszTitle As String
    ulFlags As Long
    lpfn As Long
    lParam As Long
    iImage As Long
End Type

'//
'// Win32s (Private Functions for Wrappers Below)
'//

Private Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long
Private Declare Function GetSaveFileName Lib "comdlg32.dll" Alias "GetSaveFileNameA" (pOpenfilename As OPENFILENAME) As Long
Private Declare Function SHGetPathFromIDList Lib "SHELL32.DLL" Alias "SHGetPathFromIDListA" (ByVal pidl As Long, ByVal pszPath As String) As Long
Private Declare Function SHBrowseForFolder Lib "SHELL32.DLL" Alias "SHBrowseForFolderA" (lpBrowseInfo As BROWSEINFO) As Long 'ITEMIDLIST

'//
'// Constants (Private)
'//

Private Const FW_BOLD = 700
Private Const GMEM_MOVEABLE = &H2
Private Const GMEM_ZEROINIT = &H40
Private Const GHND = (GMEM_MOVEABLE Or GMEM_ZEROINIT)
Private Const OFN_ALLOWMULTISELECT = &H200
Private Const OFN_CREATEPROMPT = &H2000
Private Const OFN_ENABLEHOOK = &H20
Private Const OFN_ENABLETEMPLATE = &H40
Private Const OFN_ENABLETEMPLATEHANDLE = &H80
Private Const OFN_EXPLORER = &H80000
Private Const OFN_EXTENSIONDIFFERENT = &H400
Private Const OFN_FILEMUSTEXIST = &H1000
Private Const OFN_HIDEREADONLY = &H4
Private Const OFN_LONGNAMES = &H200000
Private Const OFN_NOCHANGEDIR = &H8
Private Const OFN_NODEREFERENCELINKS = &H100000
Private Const OFN_NOLONGNAMES = &H40000
Private Const OFN_NONETWORKBUTTON = &H20000
Private Const OFN_NOREADONLYRETURN = &H8000
Private Const OFN_NOTESTFILECREATE = &H10000
Private Const OFN_NOVALIDATE = &H100
Private Const OFN_OVERWRITEPROMPT = &H2
Private Const OFN_PATHMUSTEXIST = &H800
Private Const OFN_READONLY = &H1
Private Const OFN_SHAREAWARE = &H4000
Private Const OFN_SHAREFALLTHROUGH = 2
Private Const OFN_SHARENOWARN = 1
Private Const OFN_SHAREWARN = 0
Private Const OFN_SHOWHELP = &H10
Private Const PD_ENABLEPRINTHOOK = &H1000
Private Const PD_ENABLEPRINTTEMPLATE = &H4000
Private Const PD_ENABLEPRINTTEMPLATEHANDLE = &H10000
Private Const PD_ENABLESETUPHOOK = &H2000
Private Const PD_ENABLESETUPTEMPLATE = &H8000
Private Const PD_ENABLESETUPTEMPLATEHANDLE = &H20000
Private Const PD_NONETWORKBUTTON = &H200000
Private Const PD_PRINTSETUP = &H40
Private Const PD_USEDEVMODECOPIES = &H40000
Private Const PD_USEDEVMODECOPIESANDCOLLATE = &H40000
Private Const PD_NOWARNING = &H80
Private Const CF_ANSIONLY = &H400&
Private Const CF_APPLY = &H200&
Private Const CF_BITMAP = 2
Private Const CF_PRINTERFONTS = &H2
Private Const CF_PRIVATEFIRST = &H200
Private Const CF_PRIVATELAST = &H2FF
Private Const CF_RIFF = 11
Private Const CF_SCALABLEONLY = &H20000
Private Const CF_SCREENFONTS = &H1
Private Const CF_BOTH = (CF_SCREENFONTS Or CF_PRINTERFONTS)
Private Const CF_DIB = 8
Private Const CF_DIF = 5
Private Const CF_DSPBITMAP = &H82
Private Const CF_DSPENHMETAFILE = &H8E
Private Const CF_DSPMETAFILEPICT = &H83
Private Const CF_DSPTEXT = &H81
Private Const CF_EFFECTS = &H100&
Private Const CF_ENABLEHOOK = &H8&
Private Const CF_ENABLETEMPLATE = &H10&
Private Const CF_ENABLETEMPLATEHANDLE = &H20&
Private Const CF_ENHMETAFILE = 14
Private Const CF_FIXEDPITCHONLY = &H4000&
Private Const CF_FORCEFONTEXIST = &H10000
Private Const CF_GDIOBJFIRST = &H300
Private Const CF_GDIOBJLAST = &H3FF
Private Const CF_INITTOLOGFONTSTRUCT = &H40&
Private Const CF_LIMITSIZE = &H2000&
Private Const CF_METAFILEPICT = 3
Private Const CF_NOFACESEL = &H80000
Private Const CF_NOVERTFONTS = &H1000000
Private Const CF_NOVECTORFONTS = &H800&
Private Const CF_NOOEMFONTS = CF_NOVECTORFONTS
Private Const CF_NOSCRIPTSEL = &H800000
Private Const CF_NOSIMULATIONS = &H1000&
Private Const CF_NOSIZESEL = &H200000
Private Const CF_NOSTYLESEL = &H100000
Private Const CF_OEMTEXT = 7
Private Const CF_OWNERDISPLAY = &H80
Private Const CF_PALETTE = 9
Private Const CF_PENDATA = 10
Private Const CF_SCRIPTSONLY = CF_ANSIONLY
Private Const CF_SELECTSCRIPT = &H400000
Private Const CF_SHOWHELP = &H4&
Private Const CF_SYLK = 4
Private Const CF_TEXT = 1
Private Const CF_TIFF = 6
Private Const CF_TTONLY = &H40000
Private Const CF_UNICODETEXT = 13
Private Const CF_USESTYLE = &H80&
Private Const CF_WAVE = 12
Private Const CF_WYSIWYG = &H8000
Private Const CFERR_CHOOSEFONTCODES = &H2000
Private Const CFERR_MAXLESSTHANMIN = &H2002
Private Const CFERR_NOFONTS = &H2001
Private Const CC_ANYCOLOR = &H100
Private Const CC_CHORD = 4
Private Const CC_CIRCLES = 1
Private Const CC_ELLIPSES = 8
Private Const CC_ENABLEHOOK = &H10
Private Const CC_ENABLETEMPLATE = &H20
Private Const CC_ENABLETEMPLATEHANDLE = &H40
Private Const CC_FULLOPEN = &H2
Private Const CC_INTERIORS = 128
Private Const CC_NONE = 0
Private Const CC_PIE = 2
Private Const CC_PREVENTFULLOPEN = &H4
Private Const CC_RGBINIT = &H1
Private Const CC_ROUNDRECT = 256 '
Private Const CC_SHOWHELP = &H8
Private Const CC_SOLIDCOLOR = &H80
Private Const CC_STYLED = 32
Private Const CC_WIDE = 16
Private Const CC_WIDESTYLED = 64
Private Const CCERR_CHOOSECOLORCODES = &H5000
Private Const LOGPIXELSY = 90
Private Const CCHDEVICENAME = 32
Private Const CCHFORMNAME = 32
Private Const SIMULATED_FONTTYPE = &H8000
Private Const PRINTER_FONTTYPE = &H4000
Private Const SCREEN_FONTTYPE = &H2000
Private Const BOLD_FONTTYPE = &H100
Private Const ITALIC_FONTTYPE = &H200
Private Const REGULAR_FONTTYPE = &H400
Private Const WM_CHOOSEFONT_GETLOGFONT = (&H400 + 1)
Private Const LBSELCHSTRING = "commdlg_LBSelChangedNotify"
Private Const SHAREVISTRING = "commdlg_ShareViolation"
Private Const FILEOKSTRING = "commdlg_FileNameOK"
Private Const COLOROKSTRING = "commdlg_ColorOK"
Private Const SETRGBSTRING = "commdlg_SetRGBColor"
Private Const FINDMSGSTRING = "commdlg_FindReplace"
Private Const HELPMSGSTRING = "commdlg_help"
Private Const CD_LBSELNOITEMS = -1
Private Const CD_LBSELCHANGE = 0
Private Const CD_LBSELSUB = 1
Private Const CD_LBSELADD = 2
Private Const NOERROR = 0
Private Const CSIDL_DESKTOP = &H0
Private Const CSIDL_PROGRAMS = &H2
Private Const CSIDL_CONTROLS = &H3
Private Const CSIDL_PRINTERS = &H4
Private Const CSIDL_PERSONAL = &H5
Private Const CSIDL_FAVORITES = &H6
Private Const CSIDL_STARTUP = &H7
Private Const CSIDL_RECENT = &H8
Private Const CSIDL_SENDTO = &H9
Private Const CSIDL_BITBUCKET = &HA
Private Const CSIDL_STARTMENU = &HB
Private Const CSIDL_DESKTOPDIRECTORY = &H10
Private Const CSIDL_DRIVES = &H11
Private Const CSIDL_NETWORK = &H12
Private Const CSIDL_NETHOOD = &H13
Private Const CSIDL_FONTS = &H14
Private Const CSIDL_TEMPLATES = &H15
Private Const BIF_RETURNONLYFSDIRS = &H1
Private Const BIF_DONTGOBELOWDOMAIN = &H2
Private Const BIF_STATUSTEXT = &H4
Private Const BIF_RETURNFSANCESTORS = &H8
Private Const BIF_BROWSEFORCOMPUTER = &H1000
Private Const BIF_BROWSEFORPRINTER = &H2000
Private Const HWND_BROADCAST = &HFFFF&
Private Const WM_WININICHANGE = &H1A
'// BrowseFolder Function
'//
'// Description:
'// Allows the user to interactively browse and select a folder found in the file system.
'//
'// Syntax:
'// StrVar = BrowseFolder(hWnd, StrVar)
'//
'// Example:
'// szFilename = BrowseFolder(Me.hWnd, "Browse for application folder:")
'//

Public Function BrowseFolder(hwnd As Long, szDialogTitle As String) As String

    Dim x As Long, BI As BROWSEINFO, dwIList As Long, szPath As String, wPos As Integer
    
    BI.hOwner = hwnd
    BI.lpszTitle = szDialogTitle
    BI.ulFlags = BIF_RETURNONLYFSDIRS
    dwIList = SHBrowseForFolder(BI)
    szPath = Space$(512)
    x = SHGetPathFromIDList(ByVal dwIList, ByVal szPath)
    If x Then
        wPos = InStr(szPath, Chr(0))
        BrowseFolder = Left$(szPath, wPos - 1)
    Else
        BrowseFolder = ""
    End If

End Function

'//
'// DialogFile Function
'//
'// Description:
'// Displays the File Open/Save As common dialog boxes.
'//
'// Syntax:
'// StrVar = DialogFile(hWnd, IntVar, StrVar, StrVar, StrVar, StrVar, StrVar)
'//
'// Example:
'// szFilename = DialogFile(Me.hWnd, 1, "Open", "", "Bitmaps" & Chr$(0) & "*.bmp" & Chr$(0) & "All Files" & Chr$(0) & "*.*", App.Path, "bmp")
'//

Public Function DialogFile(hwnd As Long, wMode As Integer, szDialogTitle As String, szFilename As String, szFilter As String, szDefDir As String, szDefExt As String) As String

    Dim x As Long, OFN As OPENFILENAME, szFile As String, szFileTitle As String
    
    OFN.lStructSize = Len(OFN)
    OFN.hwnd = hwnd
    OFN.lpstrTitle = szDialogTitle
    OFN.lpstrFile = szFilename & String$(250 - Len(szFilename), 0)
    OFN.nMaxFile = 255
    OFN.lpstrFileTitle = String$(255, 0)
    OFN.nMaxFileTitle = 255
    OFN.lpstrFilter = szFilter
    OFN.nFilterIndex = 1
    OFN.lpstrInitialDir = szDefDir
    OFN.lpstrDefExt = szDefExt

    If wMode = 1 Then
        OFN.Flags = OFN_HIDEREADONLY Or OFN_PATHMUSTEXIST Or OFN_FILEMUSTEXIST
        x = GetOpenFileName(OFN)
    Else
        OFN.Flags = OFN_HIDEREADONLY Or OFN_OVERWRITEPROMPT Or OFN_PATHMUSTEXIST
        x = GetSaveFileName(OFN)
    End If
    
    If x <> 0 Then
    
        '// If InStr(OFN.lpstrFileTitle, Chr$(0)) > 0 Then
        '//     szFileTitle = Left$(OFN.lpstrFileTitle, InStr(OFN.lpstrFileTitle, Chr$(0)) - 1)
        '// End If
        If InStr(OFN.lpstrFile, Chr$(0)) > 0 Then
            szFile = Left$(OFN.lpstrFile, InStr(OFN.lpstrFile, Chr$(0)) - 1)
        End If
        '// OFN.nFileOffset is the number of characters from the beginning of the
        '// full path to the start of the file name
        '// OFN.nFileExtension is the number of characters from the beginning of the
        '// full path to the file's extention, including the (.)
        '// MsgBox "File Name is " & szFileTitle & Chr$(13) & Chr$(10) & "Full path and file is " & szFile, , "Open"
        
        '// DialogFile = szFile & "|" & szFileTitle
        DialogFile = szFile
    
    Else
    
        DialogFile = sett.br.Text
        
        
    End If
    
End Function

