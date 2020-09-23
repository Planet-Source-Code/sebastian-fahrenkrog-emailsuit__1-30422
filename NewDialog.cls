VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cmDlg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"

'Common Dialog Class
'
'Calls Common Dialogs without using mscomdlg.ocx
'
'accesses COMDLG32.DLL directly
'
'also allows you to hook the dialog box to apply other controls, center the dialog, etc.
'
'(c) 2000 by Joe Hart (bghost@ti.cz)
'
'Thanks to Bruce McKinney and Steve McMahon for information about Common Dialogs
'Option Explicit

'note.  I prefer type libraries, which eliminate all these declares, enums, constants and some of the
'structures (UDTs) but because the type library I use that has these definitions in it is huge,
'I have put them here.  I may in the future strip out the common dialog commands and constants
'from my type library and post it on the internet.  I also plan on making a tutorial on type libraries soon.

Option Explicit

Private Declare Function SendMessageStr Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long


'Constants
' Messages:
Private Const WM_DESTROY = &H2
Private Const WM_NOTIFY = &H4E
Private Const WM_NCDESTROY = &H82
Private Const WM_GETDLGCODE = &H87
Private Const WM_INITDIALOG = &H110
Private Const WM_COMMAND = &H111

' Notification codes:
Private Const H_MAX As Long = &HFFFF + 1
Private Const CDN_FIRST = (H_MAX - 601)
Private Const CDN_LAST = (H_MAX - 699)

'Notifications when Open or Save dialog status changes
Private Const CDN_INITDONE = (CDN_FIRST - &H0)
Private Const CDN_SELCHANGE = (CDN_FIRST - &H1)
Private Const CDN_FOLDERCHANGE = (CDN_FIRST - &H2)
Private Const CDN_SHAREVIOLATION = (CDN_FIRST - &H3)
Private Const CDN_HELP = (CDN_FIRST - &H4)
Private Const CDN_FILEOK = (CDN_FIRST - &H5)
Private Const CDN_TYPECHANGE = (CDN_FIRST - &H6)
Private Const CDN_INCLUDEITEM = (CDN_FIRST - &H7)

Private Const LF_FACESIZE = 32
Private Const MAX_FILE = 260
Private Const SPI_GETWORKAREA = 48


'Enumerations
Public Enum EOpenFile
   OFN_READONLY = &H1
   OFN_OVERWRITEPROMPT = &H2
   OFN_HIDEREADONLY = &H4
   OFN_NOCHANGEDIR = &H8
   OFN_SHOWHELP = &H10
   OFN_ENABLEHOOK = &H20
   OFN_ENABLETEMPLATE = &H40
   OFN_ENABLETEMPLATEHANDLE = &H80
   OFN_NOVALIDATE = &H100
   OFN_ALLOWMULTISELECT = &H200
   OFN_EXTENSIONDIFFERENT = &H400
   OFN_PATHMUSTEXIST = &H800
   OFN_FILEMUSTEXIST = &H1000
   OFN_CREATEPROMPT = &H2000
   OFN_SHAREAWARE = &H4000
   OFN_NOREADONLYRETURN = &H8000
   OFN_NOTESTFILECREATE = &H10000
   OFN_NONETWORKBUTTON = &H20000
   OFN_NOLONGNAMES = &H40000
   OFN_EXPLORER = &H80000
   OFN_NODEREFERENCELINKS = &H100000
   OFN_LONGNAMES = &H200000
   OFN_ENABLEINCLUDENOTIFY = &H400000
   OFN_ENABLESIZING = &H800000
   OFN_NOREADONLYRETURN_C = &H8000&
End Enum

Public Enum EChooseColor
   CC_RGBINIT = &H1
   CC_FULLOPEN = &H2
   CC_PREVENTFULLOPEN = &H4
   CC_ColorShowHelp = &H8
   CC_SOLIDCOLOR = &H80
   CC_ANYCOLOR = &H100
   CC_ENABLEHOOK = &H10
   CC_ENABLETEMPLATE = &H20
   CC_ENABLETEMPLATEHANDLE = &H40
End Enum

Public Enum EChooseFont
   CF_SCREENFONTS = &H1
   CF_PRINTERFONTS = &H2
   CF_BOTH = &H3
   CF_FONTSHOWHELP = &H4
   CF_USESTYLE = &H80
   CF_EFFECTS = &H100
   CF_ANSIONLY = &H400
   CF_NOVECTORFONTS = &H800
   CF_NOOEMFONTS = &H800
   CF_NOSIMULATIONS = &H1000
   CF_LIMITSIZE = &H2000
   CF_FIXEDPITCHONLY = &H4000
   CF_WYSIWYG = &H8000
   CF_FORCEFONTEXIST = &H10000
   CF_SCALABLEONLY = &H20000
   CF_TTONLY = &H40000
   CF_NOFACESEL = &H80000
   CF_NOSTYLESEL = &H100000
   CF_NOSIZESEL = &H200000
   CF_SELECTSCRIPT = &H400000
   CF_NOSCRIPTSEL = &H800000
   CF_NOVERTFONTS = &H1000000
   CF_INITTOLOGFONTSTRUCT = &H40
   CF_APPLY = &H200
   CF_ENABLEHOOK = &H8
   CF_ENABLETEMPLATE = &H10
   CF_ENABLETEMPLATEHANDLE = &H20
End Enum

Public Enum EFontType
    SIMULATED_FONTTYPE = &H8000
    PRINTER_FONTTYPE = &H4000
    SCREEN_FONTTYPE = &H2000
    BOLD_FONTTYPE = &H100
    ITALIC_FONTTYPE = &H200
    REGULAR_FONTTYPE = &H400
End Enum

Public Enum EDialogError
    CDERR_DIALOGFAILURE = &HFFFF
    CDERR_GENERALCODES = &H0&
    CDERR_STRUCTSIZE = &H1&
    CDERR_INITIALIZATION = &H2&
    CDERR_NOTEMPLATE = &H3&
    CDERR_NOHINSTANCE = &H4&
    CDERR_LOADSTRFAILURE = &H5&
    CDERR_FINDRESFAILURE = &H6&
    CDERR_LOADRESFAILURE = &H7&
    CDERR_LOCKRESFAILURE = &H8&
    CDERR_MEMALLOCFAILURE = &H9&
    CDERR_MEMLOCKFAILURE = &HA&
    CDERR_NOHOOK = &HB&
    CDERR_REGISTERMSGFAIL = &HC&
    CFERR_CHOOSEFONTCODES = &H2000&
    CFERR_NOFONTS = &H2001&
    CFERR_MAXLESSTHANMIN = &H2002&
    FNERR_FILENAMECODES = &H3000&
    FNERR_SUBCLASSFAILURE = &H3001&
    FNERR_INVALIDFILENAME = &H3002&
    FNERR_BUFFERTOOSMALL = &H3003&
    CCERR_CHOOSECOLORCODES = &H5000&
End Enum

'Structures (User Defined Types)
Private Type TOPENFILENAME
  lStructSize As Long
  hwndOwner As Long
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
  flags As Long
  nFileOffset As Integer
  nFileExtension As Integer
  lpstrDefExt As String
  lCustData As Long
  lpfnHook As Long
  lpTemplateName As String
End Type

Private Type TCHOOSECOLOR
    lStructSize As Long
    hwndOwner As Long
    hInstance As Long
    rgbResult As Long
    lpCustColors As Long
    flags As Long
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As Long
End Type

Private Type TCHOOSEFONT
    lStructSize As Long
    hwndOwner As Long
    hdc As Long
    lpLogFont As Long
    iPointSize As Long
    flags As Long
    rgbColors As Long
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As Long
    hInstance As Long
    lpszStyle As String
    nFontType As Integer
    iAlign As Integer
    nSizeMin As Long
    nSizeMax As Long
End Type

Private Type LOGFONT
    lfHeight As Long
    lfWidth As Long
    lfEscapement As Long
    lfOrientation As Long
    lfWeight As Long
    lfItalic As Byte
    lfUnderline As Byte
    lfStrikeOut As Byte
    lfCharSet As Byte
    lfOutPrecision As Byte
    lfClipPrecision As Byte
    lfQuality As Byte
    lfPitchAndFamily As Byte
    lfFaceName(LF_FACESIZE) As Byte
End Type

Private Type NMHDR
    hWndFrom As Long
    idFrom As Long
    Code As Long
End Type

Private Type POINTL
    x As Long
    y As Long
End Type

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

'Declarations
Private Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long
Private Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOPENFILENAME As TOPENFILENAME) As Long
Private Declare Function GetSaveFileName Lib "comdlg32.dll" Alias "GetSaveFileNameA" (pOPENFILENAME As TOPENFILENAME) As Long
Private Declare Function ChooseColor Lib "comdlg32.dll" Alias "ChooseColorA" (pCHOOSECOLOR As TCHOOSECOLOR) As Long
Private Declare Function ChooseFont Lib "COMDLG32" Alias "ChooseFontA" (pCHOOSEFONT As TCHOOSEFONT) As Long
Private Declare Function CommDlgExtendedError Lib "comdlg32.dll" () As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Private Declare Function GetParent Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function MoveWindow Lib "user32" (ByVal hwnd As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
Private Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, ByRef lpvParam As Any, ByVal fuWinIni As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (lpvDest As Any, lpvSource As Any, ByVal cbCopy As Long)
Private Declare Sub CopyMemoryStr Lib "kernel32" Alias "RtlMoveMemory" (lpvDest As Any, ByVal lpvSource As String, ByVal cbCopy As Long)
Private Declare Function PrintDlg Lib "comdlg32.dll" Alias "PrintDlgA" (pPrintdlg As PrintDlg) As Long

'local variables to hold property value(s)
Private m_Font As New StdFont
Private m_CancelError As Boolean
Private m_DefaultExt As String
Private m_DialogTitle As String
Private m_FileName As String
Private m_FileTitle As String
Private m_FilterIndex As Integer
Private m_Filter As String
Private m_flags As Long
Private m_InitDir As String
Private m_MaxFileSize As Integer
Private m_hWnd As Long
Private m_FileExt As Integer
Private m_fHook As Boolean
Private m_FontMinSize As Long
Private m_FontMaxSize As Long
Private m_FontColor As Long
Private m_Color As Long
Private m_ExtendedErr As Long
Private alCustom(0 To 15) As Long

'events
Public Event InitDialog(ByVal hDlg As Long)
Public Event FileChange(ByVal hDlg As Long)
Public Event FolderChange(ByVal hDlg As Long)
Public Event DialogOK(ByRef bCancel As Boolean)
Public Event TypeChange(ByVal hDlg As Long)
Public Event DialogClose()

' Messages which can be sent to the standard dialog elements
Private Const WM_USER = &H400
Private Const CDM_FIRST = (WM_USER + 100)
Private Const CDM_LAST = (WM_USER + 200)
Private Const CDM_GETSPEC = (CDM_FIRST + &H0)
Private Const CDM_GETFILEPATH = (CDM_FIRST + &H1)
Private Const CDM_GETFOLDERPATH = (CDM_FIRST + &H2)
Private Const CDM_GETFOLDERIDLIST = (CDM_FIRST + &H3)
Private Const CDM_SETCONTROLTEXT = (CDM_FIRST + &H4)
Private Const CDM_HIDECONTROL = (CDM_FIRST + &H5)
Private Const CDM_SETDEFEXT = (CDM_FIRST + &H6)

' IDs for standard common dialog controls
Private Const ID_OPEN = &H1  'Open or Save button
Private Const ID_CANCEL = &H2 'Cancel Button
Private Const ID_HELP = &H40E 'Help Button
Private Const ID_READONLY = &H410 'Read-only check box
Private Const ID_FILETYPELABEL = &H441 'Files of type label
Private Const ID_FILELABEL = &H442 'File name label
Private Const ID_FOLDERLABEL = &H443 'Look in label
Private Const ID_LIST = &H461 'Parent of file list
Private Const ID_FORMAT = &H470 'File type combo box
Private Const ID_FOLDER = &H471 'Folder combo box
Private Const ID_FILETEXT = &H480 'File name text box

'used for page setup dialogs
Private Type POINTAPI
  x As Long
  y As Long
End Type


'type for page setup dialogs
Private Type PAGESETUPDLG
  lStructSize As Long
  hwndOwner As Long
  hDevMode As Long
  hDevNames As Long
  flags As Long
  ptPaperSize As POINTAPI
  rtMinMargin As RECT
  rtMargin As RECT
  hInstance As Long
  lCustData As Long
  lpfnPageSetupHook As Long
  lpfnPagePaintHook As Long
  lpPageSetupTemplateName As String
  hPageSetupTemplate As Long
End Type

'printer dialog
Private Type PrintDlg
  lStructSize As Long
  hwndOwner As Long
  hDevMode As Long
  hDevNames As Long
  hdc As Long
  flags As Long
  nFromPage As Integer
  nToPage As Integer
  nMinPage As Integer
  nMaxPage As Integer
  nCopies As Integer
  hInstance As Long
  lCustData As Long
  lpfnPrintHook As Long
  lpfnSetupHook As Long
  lpPrintTemplateName As String
  lpSetupTemplateName As String
  hPrintTemplate As Long
  hSetupTemplate As Long
End Type


Private m_cHookedDialog As Long

Property Let HookedDialog(ByRef cThis As cmDlg)
    'Set cHookedDialog = cThis
    m_cHookedDialog = ObjPtr(cThis)
End Property

Property Get HookedDialog() As cmDlg
   Dim oThis As cmDlg
   If (m_cHookedDialog <> 0) Then
      ' Turn the pointer into an illegal, uncounted interface
      CopyMemory oThis, m_cHookedDialog, 4
      ' Do NOT hit the End button here! You will crash!
      ' Assign to legal reference
      Set HookedDialog = oThis
      ' Still do NOT hit the End button here! You will still crash!
      ' Destroy the illegal reference
      CopyMemory oThis, 0&, 4
   End If
End Property
Public Sub ClearHookedDialog()
    m_cHookedDialog = 0
End Sub

Public Function DialogHookFunction(ByVal hDlg As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
   Dim ComDlg As cmDlg
   Set ComDlg = HookedDialog
   If Not (ComDlg Is Nothing) Then 'just to make sure the class was created properly...
      DialogHookFunction = ComDlg.DialogHook(hDlg, msg, wParam, lParam)
   End If
End Function

Public Property Get GetComDlgFileName(ByVal hDlg As Long) As String
   Dim sBuf As String
   Dim Pos As Long
   Dim hwnd As Long
   hwnd = GetParent(hDlg)
   sBuf = String$(260, 0)
   SendMessageStr hwnd, CDM_GETFILEPATH, 260, sBuf
   GetComDlgFileName = NullTrim(sBuf)
End Property

Public Function NullTrim(s) As String
'convert a null terminated string to standard vb string, deleting any leading or trailing spaces
     Dim i As Integer
     i = InStr(s, vbNullChar)
     If i > 0 Then s = Left$(s, i - 1)
     s = Trim$(s)
     NullTrim = s
End Function



Public Function DialogHook(ByVal hDlg As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long)
   Dim NotifyMessage As NMHDR
  
   Select Case msg
   Case WM_INITDIALOG
      RaiseEvent InitDialog(hDlg)
   Case WM_NOTIFY
      CopyMemory NotifyMessage, ByVal lParam, Len(NotifyMessage)
      Select Case NotifyMessage.Code
      Case CDN_SELCHANGE
         ' Changed selected file:
         RaiseEvent FileChange(hDlg)
      Case CDN_FOLDERCHANGE
         ' Changed folder:
         RaiseEvent FolderChange(hDlg)
      Case CDN_FILEOK
         ' Clicked OK:
         SetWindowLong hDlg, 0&, 0&
      Case CDN_HELP
         ' Help clicked
      Case CDN_TYPECHANGE
         RaiseEvent TypeChange(hDlg)
      Case CDN_INCLUDEITEM
         ' Hmmm
      End Select
   Case WM_DESTROY
      RaiseEvent DialogClose
   End Select

End Function

Public Sub CenterDialog(ByVal hDlg As Long, Optional ByRef oCenterTo As Object)
   Dim lhWnd As Long
   Dim WindRect As RECT
   Dim DialogRect As RECT
   Dim tp As POINTL
   Dim hWndCenterTo As Long
   Dim lL As Long
   Dim lT As Long
   Dim lR As Long

   lhWnd = GetParent(hDlg)
   GetWindowRect lhWnd, DialogRect
   On Error Resume Next
   hWndCenterTo = oCenterTo.hwnd
   If (Err.Number = 0) Then
      GetWindowRect hWndCenterTo, WindRect
   Else
      ' Assume the screen object:
      lR = SystemParametersInfo(SPI_GETWORKAREA, 0, WindRect, 0)
      If (lR = 0) Then
         ' Call failed - just use standard screen:
         WindRect.Left = 0
         WindRect.Top = 0
         WindRect.Right = Screen.Width \ Screen.TwipsPerPixelX
         WindRect.Bottom = Screen.Height \ Screen.TwipsPerPixelY
      End If
   End If
   On Error GoTo 0
   If (WindRect.Right > 0) And (WindRect.Bottom > 0) Then
        lL = WindRect.Left + (((WindRect.Right - WindRect.Left) - (DialogRect.Right - DialogRect.Left)) \ 2)
        lT = WindRect.Top + (((WindRect.Bottom - WindRect.Top) - (DialogRect.Bottom - DialogRect.Top)) \ 2)
        MoveWindow lhWnd, lL, lT, (DialogRect.Right - DialogRect.Left), (DialogRect.Bottom - DialogRect.Top), 1
    End If
End Sub

Public Property Let FileExt(ByVal vData As Integer)
   m_FileExt = vData
End Property

Public Property Get FileExt() As Integer
   FileExt = m_FileExt
End Property

Public Property Let hwnd(ByVal vData As Long)
   m_hWnd = vData
End Property

Public Property Get hwnd() As Long
   hwnd = m_hWnd
End Property

Public Sub ShowSave()
   'Shows the Save File Dialog
   Dim OpenFileName As TOPENFILENAME
   Dim L As Long
   With OpenFileName
      'set the data
      .flags = FileFlags
      .hwndOwner = hwnd
      .hInstance = 0
      .lCustData = 0
      .lpfnHook = 0
      .lpstrDefExt = StrPtr(DefaultExt)
      .lpstrFile = FileName & String$(MAX_FILE - Len(FileName) + 1, vbNullChar)
      .lpstrFileTitle = FileTitle & Space$(256)
      .lpstrFilter = m_Filter
      .lpstrInitialDir = InitDir
      .lpstrTitle = DialogTitle
      .lpTemplateName = 0
      .lStructSize = Len(OpenFileName)
      .nFileExtension = 0
      .nFileOffset = 0
      .nFilterIndex = FilterIndex
      .nMaxCustFilter = 0
      .nMaxFile = MAX_FILE
      .nMaxFileTitle = MAX_FILE
   End With
   'and call the dialog box
   L = GetSaveFileName(OpenFileName)
   Select Case L
   Case 1
      With OpenFileName
         'now fill the data with result from dialog
         FileFlags = .flags
         DefaultExt = .lpstrDefExt
         FileName = NullTrim(.lpstrFile)
         FileTitle = NullTrim(.lpstrFileTitle)
         FileExt = .nFileExtension
         m_Filter = NullTrim(.lpstrFilter)
         InitDir = NullTrim(.lpstrInitialDir)
         FilterIndex = NullTrim(.nFilterIndex)
      End With
   Case 0
      'if user pressed cancel then generate error if CancelError is true (default is false)
      If CancelError Then Err.Raise 1001, "Run-time error", "Cancel was selected"
   Case Else
     ' Extended error
        m_ExtendedErr = CommDlgExtendedError()
        Err.Raise m_ExtendedErr
   End Select
End Sub

Public Sub ShowOpen()
   'Shows the Open File Dialog
   Dim ofn As TOPENFILENAME
   Dim L As Long
   With ofn
      'fill the data
      .flags = m_flags
      .hwndOwner = m_hWnd
      .hInstance = 0
      .lpfnHook = 0
      .lCustData = 0
      .lpstrDefExt = m_DefaultExt
      .lpstrFile = m_FileName & String$(MAX_FILE - Len(m_FileName) + 1, 0)
      .lpstrFileTitle = m_FileTitle & Space$(256)
      .lpstrFilter = m_Filter
      .lpstrInitialDir = m_InitDir
      .lpstrTitle = m_DialogTitle
      .lpTemplateName = 0
      .lStructSize = Len(ofn)
      .nFileExtension = 0
      .nFileOffset = 0
      .nFilterIndex = m_FilterIndex
      .nMaxCustFilter = 0
      .nMaxFile = MAX_FILE
      .nMaxFileTitle = MAX_FILE
      'apply hook if needed.
      'If m_fHook Then
      '   HookedDialog = Me
      '   .lpfnHook = HookAddress(AddressOf DialogHookFunction)
      '   .Flags = .Flags Or OFN_ENABLEHOOK Or OFN_EXPLORER
      'End If
   End With
   'call the dialog
   L = GetOpenFileName(ofn)
   ClearHookedDialog
   Select Case L
   Case 1
      With ofn
         'and fill data with results from dialog
         m_flags = .flags
         m_DefaultExt = .lpstrDefExt
         m_FileName = NullTrim(.lpstrFile)
         m_FileTitle = NullTrim(.lpstrFileTitle)
         m_FileExt = .nFileExtension
         m_Filter = NullTrim(.lpstrFilter)
         m_InitDir = NullTrim(.lpstrInitialDir)
         m_FilterIndex = NullTrim(.nFilterIndex)
      End With
   Case 0
      'if user pressed cancel then generate error if CancelError is true (default is false)
      If m_CancelError Then Err.Raise 1002, "Run-time error", "Cancel was selected"
   Case Else
      m_ExtendedErr = CommDlgExtendedError()
      Err.Raise m_ExtendedErr
   End Select
End Sub

Public Sub ShowFont()
   Dim PrinterDC As Long
   Dim L As Long
    ' Unwanted m_flags bits as we don't support them
    Const CF_FontNotSupported = CF_APPLY Or CF_ENABLEHOOK Or CF_ENABLETEMPLATE
    ' m_flags can get reference variable or constant with bit m_flags
    
    'Set the hdc for the printer if printerfonts are being used
    If m_flags And CF_PRINTERFONTS Then PrinterDC = Printer.hdc
    ' Must have some fonts
    If (m_flags And CF_PRINTERFONTS) = 0 Then m_flags = m_flags Or CF_SCREENFONTS
    
    'check to see if there was a color selected
    If m_FontColor > 0 Then m_flags = m_flags Or CF_EFFECTS
    
    'check to see if there were minimum or maximum sizes
    If m_FontMinSize > 0 Or m_FontMaxSize > 0 Then m_flags = m_flags Or CF_LIMITSIZE
    
    ' Put in required internal m_flags and remove unsupported
    m_flags = (m_flags Or CF_INITTOLOGFONTSTRUCT) And Not CF_FontNotSupported
    
    ' Initialize LOGFONT variable
    Dim LogFnt As LOGFONT
    Const PointsPerTwip = 1440 / 72
    LogFnt.lfHeight = -(m_Font.Size * (PointsPerTwip / Screen.TwipsPerPixelY))
    LogFnt.lfWeight = m_Font.Weight
    LogFnt.lfItalic = m_Font.Italic
    LogFnt.lfUnderline = m_Font.Underline
    LogFnt.lfStrikeOut = m_Font.Strikethrough
    ' Other fields zero
    StrToBytes LogFnt.lfFaceName, m_Font.name

    ' Initialize TCHOOSEFONT variable
    Dim ChooseFnt As TCHOOSEFONT
    With ChooseFnt
      .lStructSize = Len(ChooseFnt)
      .hwndOwner = m_hWnd
      .hdc = PrinterDC
      .lpLogFont = VarPtr(LogFnt)
      .iPointSize = m_Font.Size * 10
      .flags = m_flags
      .rgbColors = Color
      .nSizeMin = m_FontMinSize
      .nSizeMax = m_FontMaxSize
    End With
    
    ' Call the dialog box
    L = ChooseFont(ChooseFnt)
    Select Case L
    Case 1
        ' Success
        m_flags = ChooseFnt.flags
        m_FontColor = ChooseFnt.rgbColors
        m_Font.Bold = ChooseFnt.nFontType And BOLD_FONTTYPE
        m_Font.Italic = LogFnt.lfItalic
        m_Font.Strikethrough = LogFnt.lfStrikeOut
        m_Font.Underline = LogFnt.lfUnderline
        m_Font.Weight = LogFnt.lfWeight
        m_Font.Size = ChooseFnt.iPointSize / 10
        m_Font.name = StrConv(LogFnt.lfFaceName, vbUnicode)
    Case 0
      'canceled
      If m_CancelError Then Err.Raise 1001, "Run-time error", "Cancel was selected"
    Case Else
         ' Extended error
        m_ExtendedErr = CommDlgExtendedError()
        Err.Raise m_ExtendedErr
   End Select
        
End Sub

Sub ShowColor()
    Dim ChooseClr As TCHOOSECOLOR
    Dim afMask As Long
    Dim L As Long
    
    With ChooseClr
      .lStructSize = Len(ChooseClr)
    
      .hwndOwner = m_hWnd
      .rgbResult = m_Color
      
      ' Mask out unwanted bits
      afMask = CLng(Not (CC_ENABLEHOOK Or _
                       CC_ENABLETEMPLATE))
      .flags = afMask And m_flags
      .lpCustColors = VarPtr(alCustom(0))
    End With
   L = ChooseColor(ChooseClr)
    
    Select Case L
    Case 1
        ' Success
        m_Color = ChooseClr.rgbResult
    Case 0
        ' Cancelled
        If m_CancelError = True Then Err.Raise 1004, , "Choose Color Dialog was canceled"
        m_Color = -1
    Case Else
        ' Extended error
        m_ExtendedErr = CommDlgExtendedError()
        Err.Raise m_ExtendedErr
    End Select

End Sub

Public Function ShowPrinter() As Boolean
  'returns true if the dialog was used to assign a printer,
  'and/or print properties
  
  Dim pdlg As PrintDlg
  Dim lngResult As Long
  
  'set initial properties
  
  'window handle of owner
  pdlg.hwndOwner = m_hWnd
    
  'structure size
  pdlg.lStructSize = Len(pdlg)
  
  'call the api function
  lngResult& = PrintDlg(pdlg)
    
  If lngResult& <> 0 Then
    ShowPrinter = True
  Else
    ShowPrinter = False
    
  End If
  
  
End Function

Private Sub InitColors()
    Dim i As Integer
    ' Initialize with first 16 system interface colors
    For i = 0 To 15
        alCustom(i) = GetSysColor(i)
    Next
End Sub

' Property to read or modify custom colors (use to save colors in registry)
Public Property Get CustomColor(i As Integer) As Long
    If i >= 0 And i <= 15 Then
        CustomColor = alCustom(i)
    Else
        CustomColor = -1
    End If
End Property
Public Property Let Color(NewColor As Long)
   m_Color = NewColor
End Property
 
Public Property Get Color() As Long
    Color = m_Color
End Property
 
Public Property Let FontColor(NewColor As Long)
   m_FontColor = NewColor
End Property

Public Property Get FontColor() As Long
   FontColor = m_FontColor
End Property
Public Property Let FontMinSize(MinSize As Long)
   m_FontMinSize = MinSize
End Property

Public Property Let FontMaxSize(MaxSize As Long)
   m_FontMaxSize = MaxSize
End Property

Private Function HookAddress(Pointer As Long) As Long
    HookAddress = Pointer
End Function

Public Property Let InitDir(ByVal vData As String)
   ' Directory to open window in
   ' Default: "C:\"
   m_InitDir = vData
End Property
Public Property Get InitDir() As String
   InitDir = m_InitDir
End Property

Public Property Let FileFlags(ByVal vData As EOpenFile)
   ' Flags for the file dialogs
   m_flags = vData
End Property

Public Property Let flags(NewFlags As Long)
   'used for compatibility with the standard dialog control.  It is recommended that you use the
   'FileFlags, FontFlags and ColorFlags instead as they provide you with a list of the flags available.
   m_flags = NewFlags
End Property

Public Property Get flags() As Long
   flags = m_flags
End Property

Public Property Get FileFlags() As EOpenFile
   FileFlags = m_flags
End Property

Public Property Let FontFlags(ByVal vData As EChooseFont)
   'flags for the font dialog
   m_flags = vData
End Property

Public Property Get FontFlags() As EChooseFont
   FontFlags = m_flags
End Property

Public Property Let ColorFlags(ByVal vData As EChooseColor)
   'flages for the color dialog
   m_flags = vData
End Property

Public Property Get ColorFlags() As EChooseColor
   ColorFlags = m_flags
End Property

Public Property Let Filter(ByVal vData As String)
   ' Filters that the user can select in drowpdown combo
   ' Usage: Friendlyname1|*.ex1|Freindlyname2|*.ex2 etc.
   ' Default: "All Files (*.*)|*.*"
   Dim pipepos As String
   Do While InStr(vData, "|") > 0
      pipepos = InStr(vData, "|")
      If pipepos > 0 Then
         vData = Left$(vData, pipepos - 1) & vbNullChar & Right$(vData, Len(vData) - pipepos)
      End If
   Loop
   If Right$(vData, 2) <> vbNullChar & vbNullChar Then vData = vData & vbNullChar
   If Right$(vData, 2) <> vbNullChar & vbNullChar Then vData = vData & vbNullChar
   m_Filter = vData
End Property

Public Property Get Filter() As String
   Dim nullpos As String
   Dim tempfilter As String
   tempfilter = m_Filter
   Do While InStr(tempfilter, vbNullChar) > 0
      nullpos = InStr(tempfilter, vbNullChar)
      If nullpos > 0 Then
         tempfilter = Left$(tempfilter, nullpos - 1) & vbNullChar & Right$(tempfilter, Len(tempfilter) - nullpos)
      End If
   Loop
   If Right$(tempfilter, 1) = "|" Then tempfilter = Left$(tempfilter, Len(tempfilter) - 1)
   If Right$(tempfilter, 1) = "|" Then tempfilter = Left$(tempfilter, Len(tempfilter) - 1)
   Filter = tempfilter
End Property

Public Property Let FilterIndex(ByVal vData As Integer)
   ' Index of filter to select as default
   ' The first item is 1, second 2, etc.
   ' Default: 1
   m_FilterIndex = vData
End Property

Public Property Get FilterIndex() As Integer
   FilterIndex = m_FilterIndex
End Property

Public Property Let FileTitle(ByVal vData As String)
   ' The name of the file without path
   m_FileTitle = vData
End Property

Public Property Get FileTitle() As String
   FileTitle = m_FileTitle
End Property

Public Property Let FileName(ByVal vData As String)
' Name of the file, including path
   m_FileName = vData
End Property

Public Property Get FileName() As String
   FileName = m_FileName
End Property

Public Property Let DialogTitle(ByVal vData As String)
' The name of the dialog box
   m_DialogTitle = vData
End Property

Public Property Get DialogTitle() As String
   DialogTitle = m_DialogTitle
End Property

Public Property Let DefaultExt(ByVal vData As String)
   ' The default extension added if one is not specified in the name
   m_DefaultExt = vData
End Property

Public Property Get DefaultExt() As String
   DefaultExt = m_DefaultExt
End Property

Public Property Let CancelError(ByVal vData As Boolean)
   ' Raise an error if user clicks cancel
   ' Default: False
   m_CancelError = vData
End Property

Public Property Get CancelError() As Boolean
   CancelError = m_CancelError
End Property

Private Sub StrToBytes(ab() As Byte, s As String)
    If IsArrayEmpty(ab) Then
        ' Assign to empty array
        ab = StrConv(s, vbFromUnicode)
    Else
        Dim cab As Long
        ' Copy to existing array, padding or truncating if necessary
        cab = UBound(ab) - LBound(ab) + 1
        If Len(s) < cab Then s = s & String$(cab - Len(s), 0)
        CopyMemoryStr ab(LBound(ab)), s, cab
    End If
End Sub

Public Property Get FontBold() As Boolean
    'return object's FontBold property
    FontBold = m_Font.Bold
End Property

Public Property Let FontBold(ByVal vNewValue As Boolean)
    'Assign object's FontBold property
    m_Font.Bold = vNewValue
End Property

Public Property Get FontItalic() As Boolean
    'Return object's FontItalic property
    FontItalic = m_Font.Italic
End Property

Public Property Let FontItalic(ByVal vNewValue As Boolean)
    'Assign object's FontItalic property
    m_Font.Italic = vNewValue
End Property

Public Property Get FontName() As String
    'Return object's Fontname property
    FontName = m_Font.name
End Property

Public Property Let FontName(ByVal vNewValue As String)
    'Assign object's FontName property
    m_Font.name = vNewValue
End Property

Public Property Get FontSize() As Long
    'Return object's FontSize property
    FontSize = m_Font.Size
End Property

Public Property Let FontSize(ByVal vNewValue As Long)
    'Assign object's FontSize property
    m_Font.Size = vNewValue
End Property

Public Property Get Font() As StdFont
    Set Font = m_Font
End Property

Public Property Let Font(sFont As StdFont)
    Set m_Font = sFont
End Property

Private Sub Class_Initialize()
   'set up defaults
   CancelError = False
   DefaultExt = ""
   DialogTitle = ""
   FileName = ""
   FileTitle = ""
   Filter = "All Files|*.*"
   FilterIndex = 1
   InitDir = App.path
   hwnd = 0
   InitColors
End Sub

Private Function IsArrayEmpty(va As Variant) As Boolean
    Dim v As Variant
    On Error Resume Next
    v = va(LBound(va))
    IsArrayEmpty = (Err <> 0)
End Function

Public Property Get Hook() As Boolean
       Hook = m_fHook
End Property

Public Property Let Hook(NewHook As Boolean)
       m_fHook = NewHook
End Property

