VERSION 5.00
Begin VB.UserControl ucComboBoxEx 
   BackColor       =   &H80000005&
   ClientHeight    =   540
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3330
   ScaleHeight     =   540
   ScaleWidth      =   3330
End
Attribute VB_Name = "ucComboBoxEx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Version V1.20 by Zhu JinYong,BaMao,HuiGong,CongYang,Anqing,Anhui,China
'Thanks give to:
'Dana Seaman www.cyberactivex.com   -Dana have rich knowledge about Unicode,his Unicode Activex are great
'Selftaught VBComctl
'Carles P.V. Style of programming
'Steven www.vbaccelerator.com
'---------------------------------------------------------------------------------------
' (*) Self-Subclassing UserControl template (IDE safe) by Paul Caton:
'
'     Self-subclassing Controls/Forms - NO dependencies
'     http://www.planet-source-code.com/vb/scripts/ShowCode.asp?txtCodeId=54117&lngWId=1
'
'----------------------------------------------------------------------------------------
'Requires:    OleGuids3.tlb (in IDE only)
'Dependence mIOIAComboBoxEx.bas

'Fixed 'Overflow error' if Font name is Unicode
'Fixed FindItem issue when item is Unicode
'Remove compilation condition 'Unicode=1'
'Solved XP Theme problem to cause CreateWindowExW failed (16/June/2007)
'Use Paul's Updated Self-Subclass and LaVolpe's more robust Version to solve UserControl Terminate error. (20/April/2008)

Option Explicit

'EditBox
Private Const EM_GETSEL = &HB0
Private Const EM_SETSEL = &HB1

' Combo box styles:
Private Const CBS_AUTOHSCROLL As Long = &H40&
Private Const CBS_DROPDOWN As Long = &H2&
Private Const CBS_DROPDOWNLIST As Long = &H3&
Private Const CBS_HASSTRINGS As Long = &H200&
Private Const CBS_DISABLENOSCROLL As Long = &H800&
Private Const CBS_NOINTEGRALHEIGHT As Long = &H400&
Private Const CBS_OWNERDRAWFIXED As Long = &H10&
Private Const CBS_OWNERDRAWVARIABLE As Long = &H20&
Private Const CBS_SIMPLE As Long = &H1&
Private Const CBS_SORT As Long = &H100&

' Combo box messages:
Private Const CB_ADDSTRING As Long = &H143
Private Const CB_DELETESTRING As Long = &H144
Private Const CB_DIR As Long = &H145
Private Const CB_ERR As Long = (-1)
Private Const CB_ERRSPACE As Long = (-2)
Private Const CB_FINDSTRING As Long = &H14C
Private Const CB_FINDSTRINGEXACT As Long = &H158
Private Const CB_GETCOUNT As Long = &H146
Private Const CB_GETCURSEL As Long = &H147
Private Const CB_GETDROPPEDCONTROLRECT As Long = &H152
Private Const CB_GETDROPPEDSTATE As Long = &H157
Private Const CB_GETEDITSEL As Long = &H140
Private Const CB_GETEXTENDEDUI As Long = &H156
Private Const CB_GETITEMDATA As Long = &H150
Private Const CB_GETITEMHEIGHT As Long = &H154
Private Const CB_GETLBTEXT As Long = &H148
Private Const CB_GETLBTEXTLEN As Long = &H149
Private Const CB_GETLOCALE As Long = &H15A
Private Const CB_INSERTSTRING As Long = &H14A
Private Const CB_LIMITTEXT As Long = &H141
Private Const CB_MSGMAX As Long = &H15B
Private Const CB_OKAY As Long = 0
Private Const CB_RESETCONTENT As Long = &H14B
Private Const CB_SELECTSTRING As Long = &H14D
Private Const CB_SETCURSEL As Long = &H14E
Private Const CB_SETEDITSEL As Long = &H142
Private Const CB_SETEXTENDEDUI As Long = &H155
Private Const CB_SETITEMDATA As Long = &H151
Private Const CB_SETITEMHEIGHT As Long = &H153
Private Const CB_SETLOCALE As Long = &H159
Private Const CB_SHOWDROPDOWN As Long = &H14F
Private Const CB_GETDROPPEDWIDTH As Long = &H15F
Private Const CB_SETDROPPEDWIDTH As Long = &H160

' Combo box notifications:
Private Const CBN_CLOSEUP As Long = 8
Private Const CBN_DBLCLK As Long = 2
Private Const CBN_DROPDOWN As Long = 7
Private Const CBN_EDITCHANGE As Long = 5
Private Const CBN_EDITUPDATE As Long = 6
Private Const CBN_KILLFOCUS As Long = 4
Private Const CBN_SELCHANGE As Long = 1
Private Const CBN_SELENDCANCEL As Long = 10
Private Const CBN_SELENDOK As Long = 9
Private Const CBN_SETFOCUS As Long = 3

' Owner draw style types:
Private Const ODS_CHECKED As Long = &H8
Private Const ODS_DISABLED As Long = &H4
Private Const ODS_FOCUS As Long = &H10
Private Const ODS_GRAYED As Long = &H2
Private Const ODS_SELECTED As Long = &H1
Private Const ODS_COMBOBOXEDIT As Long = &H1000

' Owner draw action types:
Private Const ODA_DRAWENTIRE As Long = &H1
Private Const ODA_FOCUS As Long = &H4
Private Const ODA_SELECT As Long = &H2

' Combo box extended styles:
Private Const CBES_EX_NOEDITIMAGE = &H1& ' no image to left of edit portion
Private Const CBES_EX_NOEDITIMAGEINDENT = &H2& ' edit box and dropdown box will not display images
Private Const CBES_EX_PATHWORDBREAKPROC = &H4& ' NT only. Edit box uses \ . and / as word delimiters
'#if (_WIN32_IE >= 0x0400)
Private Const CBES_EX_NOSIZELIMIT = &H8& ' Allow combo box ex vertical size < combo, clipped.
Private Const CBES_EX_CASESENSITIVE = &H10& ' case sensitive search

Private Const CBEIF_TEXT As Long = &H1
Private Const CBEIF_IMAGE As Long = &H2
Private Const CBEIF_SELECTEDIMAGE As Long = &H4
Private Const CBEIF_OVERLAY As Long = &H8
Private Const CBEIF_INDENT As Long = &H10
Private Const CBEIF_LPARAM As Long = &H20
Private Const CBEIF_DI_SETITEM As Long = &H10000000

Private Const WM_USER As Long = &H400
Private Const CBEM_SETIMAGELIST As Long = (WM_USER + 2)
Private Const CBEM_DELETEITEM As Long = CB_DELETESTRING
Private Const CBEMAXSTRLEN = 260
Private Const CBEM_GETCOMBOCONTROL As Long = (WM_USER + 6)
Private Const CBEM_GETEDITCONTROL As Long = (WM_USER + 7)
Private Const CBEM_SETEXSTYLE = (WM_USER + 8)

Private Const CBEM_INSERTITEMW As Long = (WM_USER + 11)
Private Const CBEM_SETITEMW As Long = (WM_USER + 12)
Private Const CBEM_GETITEMW As Long = (WM_USER + 13)

Private Const CBEM_INSERTITEMA As Long = (WM_USER + 1)
Private Const CBEM_SETITEMA As Long = (WM_USER + 5)
Private Const CBEM_GETITEMA As Long = (WM_USER + 4)

Private Const CBENF_KILLFOCUS As Long = 1
Private Const CBENF_RETURN As Long = 2
Private Const CBENF_ESCAPE As Long = 3
Private Const CBENF_DROPDOWN As Long = 4

Private Const CBEN_FIRST As Long = (-800)
Private Const CBEN_BEGINEDIT As Long = (CBEN_FIRST - 4)

Private Const CBEN_ENDEDITW As Long = (CBEN_FIRST - 6)

Private Const CBEN_ENDEDITA As Long = (CBEN_FIRST - 5)

Private Const MA_NOACTIVATE As Long = 3

Public Enum eComboBoxExStyle
    cboSimple = 0
    cboDropDownCombo = 1
    cboDropDownList = 2
End Enum

Public Enum ECCXExtendedStyle
    eccxNoEditImage = CBES_EX_NOEDITIMAGE
    eccxNoImages = CBES_EX_NOEDITIMAGEINDENT
    eccxCaseSensitiveSearch = CBES_EX_CASESENSITIVE
End Enum

Public Enum eComboBoxExEndEditReason
    cboEndEditKillFocus = CBENF_KILLFOCUS
    cboEndEditReturn = CBENF_RETURN
    cboEndEditEscape = CBENF_ESCAPE
    cboEndEditDropDown = CBENF_DROPDOWN
End Enum

Private Type OSVERSIONINFO
    dwVersionInfoSize                           As Long
    dwMajorVersion                              As Long
    dwMinorVersion                              As Long
    dwBuildNumber                               As Long
    dwPlatformId                                As Long
    szCSDVersion(0 To 127)                      As Byte
End Type

Private Type COMBOBOXEXITEM
    mask As Long    ' CBEIF..
    iItem As Long
    'pszText As String
    pszText As Long
    cchTextMax As Long
    iImage As Long
    iSelectedImage As Long
    iOverlay As Long
    iIndent As Long
    lParam As Long
End Type

Private Type NMHDR
    hwndFrom As Long
    idfrom As Long
    code As Long
End Type

Private Type NMCBEENDEDIT
    hdr As NMHDR
    fChanged As Long
    iNewSelection As Long
    szText(0 To CBEMAXSTRLEN - 1) As Byte '// CBEMAXSTRLEN is 260
    iWhy As Integer
End Type

Private Type NMCBEENDEDITW
    hdr As NMHDR
    fChanged As Long
    iNewSelection As Long
    szText(0 To 518) As Byte '// CBEMAXSTRLEN is 260
    iWhy As Long
End Type

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Public Event DropDown()
Public Event CloseUp()
Public Event EditChange()
Public Event ListIndexChange()
Public Event BeginEdit()
Public Event Click()
Public Event EndEdit(ByVal bEditChanged As Boolean, ByRef iNewIndex As Long, ByRef sText As String, ByVal iWhy As eComboBoxExEndEditReason)
Public Event KeyDown(KeyCode As Integer, Shift As Integer)
Public Event KeyUp(KeyCode As Integer, Shift As Integer)
Public Event RequestDropDownResize(ByRef lLeft As Long, ByRef lTop As Long, ByRef lRight As Long, ByRef lBottom As Long, ByRef bCancel As Boolean)

Private Const PROP_Font             As String = "Fnt"
Private Const PROP_Style            As String = "Sty"
Private Const PROP_ExStyle            As String = "ExSty"
Private Const PROP_DroppedHeight    As String = "DHt"
Private Const PROP_DroppedWidth     As String = "DWd"
Private Const PROP_Enabled          As String = "Enbld"
Private Const PROP_MaxLength        As String = "MaxLen"
Private Const PROP_ExtendedUI       As String = "ExtUI"
Private Const PROP_Themeable        As String = "Them"
Private Const PROP_ImageSize          As String = "ImgSize"

Private Const DEF_Style             As Long = cboDropDownCombo
Private Const DEF_ExStyle             As Long = CBES_EX_NOEDITIMAGE + CBES_EX_NOEDITIMAGEINDENT
Private Const DEF_DroppedHeight     As Long = 160
Private Const DEF_DroppedWidth      As Long = 0
Private Const DEF_Enabled           As Boolean = True
Private Const DEF_MaxLength         As Long = 0
Private Const DEF_ExtendedUI        As Boolean = False
Private Const DEF_Themeable         As Boolean = True
Private Const DEF_ImageSize    As Long = 16

Private WithEvents m_oFont As StdFont
Attribute m_oFont.VB_VarHelpID = -1

Private m_hFont As Long
Private m_hImageList As Long
Private m_hWnd                       As Long
Private m_hWndCombo                  As Long
Private m_hWndParent              As Long
Private m_hWndEdit                 As Long

Private m_eStyle                     As eComboBoxExStyle
Private m_lNewIndex                  As Long

Private m_lDroppedHeight             As Long
Private m_lDroppedWidth              As Long
Private m_bInFocus As Boolean
Private m_bSubclass As Boolean

Private m_lMaxLength                 As Long
Private m_bRedraw                    As Boolean
Private m_bExtendedUI                As Boolean
Private m_bThemeable                 As Boolean
Private m_eExStyle As ECCXExtendedStyle
Private m_hWndDropDown As Long

Private m_tItem                      As COMBOBOXEXITEM
Private m_uIPAO                As IPAOHookStructComboBoxEx

Private m_hIml As Long
Private m_lImageSize As Long
Private m_bIsNt                                 As Boolean

'= Window general =======================================================================

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (lpDst As Any, lpSrc As Any, ByVal Length As Long)
'Private Declare Function SendMessageA Lib "user32" (ByVal hWnd As Long, _
                                                     ByVal wMsg As Long, _
                                                     ByVal wParam As Long, _
                                                     lParam As Any) As Long
'Private Declare Function SendMessageW Lib "user32.dll" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Private Declare Function SendMessageLongA Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, _
                                                                             ByVal wMsg As Long, _
                                                                             ByVal wParam As Long, _
                                                                             ByVal lParam As Long) As Long

Private Declare Function SendMessageLongW Lib "user32" Alias "SendMessageW" (ByVal hWnd As Long, _
                                                                             ByVal wMsg As Long, _
                                                                             ByVal wParam As Long, _
                                                                             ByVal lParam As Long) As Long
Private Declare Function PostMessage Lib "user32.dll" Alias "PostMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInfo As OSVERSIONINFO) As Long

Private Const VER_PLATFORM_WIN32_NT             As Long = 2

Private Type LOGFONT
    lfHeight         As Long
    lfWidth          As Long
    lfEscapement     As Long
    lfOrientation    As Long
    lfWeight         As Long
    lfItalic         As Byte
    lfUnderline      As Byte
    lfStrikeOut      As Byte
    lfCharSet        As Byte
    lfOutPrecision   As Byte
    lfClipPrecision  As Byte
    lfQuality        As Byte
    lfPitchAndFamily As Byte
    lfFaceName(32)   As Byte
End Type

Private Const LOGPIXELSY             As Long = 90
Private Const FW_NORMAL              As Long = 400
Private Const FW_BOLD                As Long = 700

Private Const WC_COMBOBOXEX As String = "ComboBoxEx32"

Private Const WC_EDIT    As String = "Edit"
Private Const ucComboBoxEx          As String = "ucComboBoxEx"
Private Const WC_REBAR As String = "ReBarWindow32"

'== Window
Private Const GWL_STYLE        As Long = (-16)
Private Const GWL_EXSTYLE      As Long = (-20)
Private Const GWL_ID As Long = -12
Private Const WS_EX_CLIENTEDGE As Long = &H200&
Private Const WS_TABSTOP       As Long = &H10000
Private Const WS_DISABLED      As Long = &H8000000
Private Const WS_CHILD         As Long = &H40000000
Private Const WS_VISIBLE As Long = &H10000000
Private Const WM_SIZE          As Long = &H5
Private Const WM_ERASEBKGND    As Long = &H14
Private Const WM_SETFONT       As Long = &H30
Private Const WM_NOTIFY        As Long = &H4E
Private Const WM_COMMAND       As Long = &H111
Private Const WM_GETTEXTLENGTH As Long = &HE
Private Const WM_GETTEXT As Long = &HD
Private Const WM_SETREDRAW As Long = &HB
Private Const WM_SETFOCUS                  As Long = &H7
Private Const WM_MOUSEACTIVATE             As Long = &H21
Private Const WM_KILLFOCUS As Long = &H8
Private Const WS_EX_LEFT As Long = &H0&
Private Const WM_KEYDOWN = &H100
Private Const WM_KEYUP = &H101
Private Const WM_CTLCOLORLISTBOX = &H134
Private Const WM_DRAWITEM = &H2B
Private Const WM_CHAR = &H102
Private Const WM_CTLCOLOREDIT = &H133

Private Const SW_HIDE = 0
Private Const SW_SHOWNORMAL = 1
Private Const SW_SHOW  As Long = 5
Private Const SM_CXVSCROLL As Long = 2

Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hDc As Long, ByVal nIndex As Long) As Long
Private Declare Function CreateFontIndirect Lib "gdi32" Alias "CreateFontIndirectA" (lpLogFont As LOGFONT) As Long
Private Declare Function MulDiv Lib "kernel32" (ByVal nNumber As Long, ByVal nNumerator As Long, ByVal nDenominator As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long

Private Declare Function CreateWindowExA Lib "user32" (ByVal dwExStyle As Long, _
                                                       ByVal lpClassName As String, _
                                                       ByVal lpWindowName As String, _
                                                       ByVal dwStyle As Long, _
                                                       ByVal x As Long, _
                                                       ByVal y As Long, _
                                                       ByVal nWidth As Long, _
                                                       ByVal nHeight As Long, _
                                                       ByVal hWndParent As Long, _
                                                       ByVal hMenu As Long, _
                                                       ByVal hInstance As Long, _
                                                       lpParam As Any) As Long

Private Declare Function CreateWindowExW Lib "user32" (ByVal dwExStyle As Long, _
                                                       ByVal lpClassName As Long, _
                                                       ByVal lpWindowName As Long, _
                                                       ByVal dwStyle As Long, _
                                                       ByVal x As Long, _
                                                       ByVal y As Long, _
                                                       ByVal nWidth As Long, _
                                                       ByVal nHeight As Long, _
                                                       ByVal hWndParent As Long, _
                                                       ByVal hMenu As Long, _
                                                       ByVal hInstance As Long, _
                                                       lpParam As Any) As Long
Private Declare Function DestroyWindow Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare Function ShowWindow Lib "user32" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long
Private Declare Function MoveWindow Lib "user32" (ByVal hWnd As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
Private Declare Function EnableWindow Lib "user32" (ByVal hWnd As Long, ByVal fEnable As Long) As Long
Private Declare Function GetFocus Lib "user32" () As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function GetWindowTextLength Lib "user32.dll" Alias "GetWindowTextLengthA" (ByVal hWnd As Long) As Long
Private Declare Function GetWindowText Lib "user32.dll" Alias "GetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Private Declare Function GetParent Lib "user32.dll" (ByVal hWnd As Long) As Long
Private Declare Function FindWindowEx Lib "user32.dll" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Private Declare Function GetWindowRect Lib "user32.dll" (ByVal hWnd As Long, ByRef lpRect As RECT) As Long
Private Declare Function GetWindowTextA Lib "user32.dll" (ByVal hWnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Private Declare Function GetWindowTextLengthA Lib "user32.dll" (ByVal hWnd As Long) As Long
Private Declare Function SetWindowTextA Lib "user32.dll" (ByVal hWnd As Long, ByVal lpString As String) As Long
Private Declare Function GetWindowTextW Lib "user32.dll" (ByVal hWnd As Long, ByVal lpString As Long, ByVal cch As Long) As Long
Private Declare Function GetWindowTextLengthW Lib "user32.dll" (ByVal hWnd As Long) As Long
Private Declare Function SetWindowTextW Lib "user32.dll" (ByVal hWnd As Long, ByVal lpString As Long) As Long
Private Declare Function GetClientRect Lib "user32.dll" (ByVal hWnd As Long, ByRef lpRect As RECT) As Long
Private Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
'Private Declare Function IsWindow Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function IsWindowVisible Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function CopyRect Lib "user32" (lpDestRect As RECT, lpSourceRect As RECT) As Long

Private Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long
Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Private Declare Function SetFocus Lib "user32" (ByVal hWnd As Long) As Long

'== Image list
Private Const CLR_NONE        As Long = -1
Private Const ILC_MASK        As Long = &H1
Private Const ILC_COLORDDB    As Long = &HFE

Private Declare Function ImageList_Create Lib "comctl32" (ByVal MinCx As Long, ByVal MinCy As Long, ByVal flags As Long, ByVal cInitial As Long, ByVal cGrow As Long) As Long
Private Declare Function ImageList_Add Lib "comctl32" (ByVal hImageList As Long, ByVal hBitmap As Long, ByVal hBitmapMask As Long) As Long
Private Declare Function ImageList_AddMasked Lib "comctl32" (ByVal hImageList As Long, ByVal hbmImage As Long, ByVal crMask As Long) As Long
Private Declare Function ImageList_AddIcon Lib "comctl32" (ByVal hImageList As Long, ByVal hIcon As Long) As Long
Private Declare Function ImageList_Destroy Lib "comctl32" (ByVal hImageList As Long) As Long

Private Const MAX_PATH                   As Long = 260

' Common control shared messages
Private Const CCM_FIRST              As Long = &H2000
'Private Const CCM_SETBKCOLOR         As Long = (CCM_FIRST + 1)    ' lParam = bkColor
'Private Const CCM_SETCOLORSCHEME     As Long = (CCM_FIRST + 2)    ' lParam = COLORSCHEME struct ptr
'Private Const CCM_GETCOLORSCHEME     As Long = (CCM_FIRST + 3)    ' lParam = COLORSCHEME struct ptr
'Private Const CCM_GETDROPTARGET      As Long = (CCM_FIRST + 4)
Private Const CCM_SETUNICODEFORMAT   As Long = (CCM_FIRST + 5)
Private Const CCM_GETUNICODEFORMAT   As Long = (CCM_FIRST + 6)
Private Const CCM_SETVERSION         As Long = (CCM_FIRST + 7)
Private Const CCM_GETVERSION         As Long = (CCM_FIRST + 8)

Private Type tagINITCOMMONCONTROLSEX
    dwSize                                      As Long
    dwICC                                       As Long
End Type
Private Const ICC_USEREX_CLASSES = &H200

Private Declare Sub InitCommonControls Lib "comctl32" ()
Private Declare Function InitCommonControlsEx Lib "comctl32" (lpInitCtrls As tagINITCOMMONCONTROLSEX) As Boolean
Private Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
Private Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long

Private m_lhMod                                 As Long
'========================================================================================
' Subclasser declarations
'========================================================================================

Private z_IDEflag           As Long         'Flag indicating we are in IDE
Private z_ScMem             As Long         'Thunk base address
Private z_scFunk            As Collection   'hWnd/thunk-address collection
Private z_hkFunk            As Collection   'hook/thunk-address collection
Private z_cbFunk            As Collection   'callback/thunk-address collection
Private Const IDX_INDEX     As Long = 2     'index of the subclassed hWnd OR hook type
Private Const IDX_CALLBACKORDINAL As Long = 22 ' Ubound(callback thunkdata)+1, index of the callback

' Declarations:
Private Declare Sub RtlMoveMemory Lib "kernel32" (ByVal Destination As Long, ByVal Source As Long, ByVal Length As Long)
Private Declare Function IsBadCodePtr Lib "kernel32" (ByVal lpfn As Long) As Long
Private Declare Function VirtualAlloc Lib "kernel32" (ByVal lpAddress As Long, ByVal dwSize As Long, ByVal flAllocationType As Long, ByVal flProtect As Long) As Long
Private Declare Function VirtualFree Lib "kernel32" (ByVal lpAddress As Long, ByVal dwSize As Long, ByVal dwFreeType As Long) As Long
Private Declare Function GetModuleHandleA Lib "kernel32" (ByVal lpModuleName As String) As Long
Private Declare Function GetModuleHandleW Lib "kernel32" (ByVal lpModuleName As Long) As Long
Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Private Enum eThunkType
    SubclassThunk = 0
    HookThunk = 1
    CallbackThunk = 2
End Enum

'-Selfsub specific declarations----------------------------------------------------------------------------
Private Enum eMsgWhen                                                   'When to callback
    MSG_BEFORE = 1                                                        'Callback before the original WndProc
    MSG_AFTER = 2                                                         'Callback after the original WndProc
    MSG_BEFORE_AFTER = MSG_BEFORE Or MSG_AFTER                            'Callback before and after the original WndProc
End Enum

'Private Const IDX_SHUTDOWN  As Long = 1     'Thunk data index of the termination flag
'Private Const IDX_INDEX     As Long = 2     'Thunk data index of the subclassed hWnd
'Private Const IDX_EBMODE    As Long = 3     'Thunk data index of the EbMode function address
'Private Const IDX_CWP       As Long = 4     'Thunk data index of the CallWindowProc function address
'Private Const IDX_SWL       As Long = 5     'Thunk data index of the SetWindowsLong function address
'Private Const IDX_FREE      As Long = 6     'Thunk data index of the VirtualFree function address
'Private Const IDX_BADPTR    As Long = 7     'Thunk data index of the IsBadCodePtr function address
'Private Const IDX_OWNER     As Long = 8     'Thunk data index of the Owner object's vTable address
Private Const IDX_WNDPROC   As Long = 9     'Thunk data index of the original WndProc
'Private Const IDX_CALLBACK  As Long = 10    'Thunk data index of the callback method address
Private Const IDX_BTABLE    As Long = 11    'Thunk data index of the Before table
Private Const IDX_ATABLE    As Long = 12    'Thunk data index of the After table
Private Const IDX_PARM_USER As Long = 13    'Thunk data index of the User-defined callback parameter data index
'Private Const IDX_EBX       As Long = 16    'Thunk code patch index of the thunk data
Private Const IDX_UNICODE   As Long = 75    'Must be Ubound(subclass thunkdata)+1; index for unicode support
Private Const ALL_MESSAGES  As Long = -1    'All messages callback
Private Const MSG_ENTRIES   As Long = 32    'Number of msg table entries. Set to 1 if using ALL_MESSAGES for all subclassed windows

' \\LaVolpe - Added non-ANSI version API calls
Private Declare Function CallWindowProcA Lib "user32" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function CallWindowProcW Lib "user32" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function GetCurrentProcessId Lib "kernel32" () As Long
Private Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hWnd As Long, lpdwProcessId As Long) As Long
Private Declare Function IsWindow Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function IsWindowUnicode Lib "user32.dll" (ByVal hWnd As Long) As Long
Private Declare Function SendMessageA Lib "user32.dll" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByRef lParam As Any) As Long
Private Declare Function SendMessageW Lib "user32.dll" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByRef lParam As Any) As Long
Private Declare Function SetWindowLongA Lib "user32" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SetWindowLongW Lib "user32" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
'-------------------------------------------------------------------------------------------------

'-SelfHook specific declarations----------------------------------------------------------------------------
Private Declare Function SetWindowsHookExA Lib "user32" (ByVal idHook As Long, ByVal lpfn As Long, ByVal hmod As Long, ByVal dwThreadId As Long) As Long
Private Declare Function SetWindowsHookExW Lib "user32" (ByVal idHook As Long, ByVal lpfn As Long, ByVal hmod As Long, ByVal dwThreadId As Long) As Long
Private Declare Function UnhookWindowsHookEx Lib "user32" (ByVal hHook As Long) As Long
Private Declare Function GetDesktopWindow Lib "user32.dll" () As Long

Private Enum eHookType  ' http://msdn2.microsoft.com/en-us/library/ms644990.aspx
    WH_MSGFILTER = -1
    WH_JOURNALRECORD = 0
    WH_JOURNALPLAYBACK = 1
    WH_KEYBOARD = 2
    WH_GETMESSAGE = 3
    WH_CALLWNDPROC = 4
    WH_CBT = 5
    WH_SYSMSGFILTER = 6
    WH_MOUSE = 7
    WH_DEBUG = 9
    WH_SHELL = 10
    WH_FOREGROUNDIDLE = 11
    WH_CALLWNDPROCRET = 12
    WH_KEYBOARD_LL = 13       ' NT/2000/XP+ only, Global hook only
    WH_MOUSE_LL = 14          ' NT/2000/XP+ only, Global hook only
End Enum

Private Function pvStripNulls(ByVal sString As String) As String

Dim lPos As Long

    lPos = InStr(sString, vbNullChar)

    If (lPos = 1) Then
        pvStripNulls = vbNullString
    ElseIf (lPos > 1) Then
        pvStripNulls = Left$(sString, lPos - 1)
        Exit Function
    End If

    pvStripNulls = sString

End Function

Private Sub m_oFont_FontChanged(ByVal PropertyName As String)

    Set Font = m_oFont

End Sub

Private Sub UserControl_GotFocus()

    If Len(Text) Then
        pvSetSelStartEnd 0, Len(Text)
    End If

End Sub

Private Sub UserControl_Initialize()

'-- Initialize font object

    Set m_oFont = New StdFont
    m_lhMod = LoadLibrary("shell32.dll")
    InitCommonControls
    InitComctl32
    VersionCheck

End Sub

Private Function InitComctl32() As Boolean

'/* init comctl32 ComboboxEx class

Dim tIcc As tagINITCOMMONCONTROLSEX

    With tIcc
        .dwSize = Len(tIcc)
        .dwICC = ICC_USEREX_CLASSES
    End With
    InitComctl32 = InitCommonControlsEx(tIcc)

End Function

Private Sub pvDestroyImageList()

    If (m_hImageList) Then
        If (ImageList_Destroy(m_hImageList)) Then
            m_hImageList = 0
        End If
    End If

End Sub

Private Sub pvDestroyFont()

    If (m_hFont) Then
        If (DeleteObject(m_hFont)) Then
            m_hFont = 0
        End If
    End If

End Sub

Private Sub UserControl_InitProperties()

'---------------------------------------------------------------------------------------
' Date      : 9/09/05
' Purpose   : Initialize property values to the defaults.
'---------------------------------------------------------------------------------------

    Set Font = UserControl.Ambient.Font
    m_eStyle = DEF_Style
    m_eExStyle = DEF_ExStyle
    m_lImageSize = DEF_ImageSize
    m_lDroppedHeight = DEF_DroppedHeight
    m_lDroppedWidth = DEF_DroppedWidth
    UserControl.Enabled = DEF_Enabled
    m_lMaxLength = DEF_MaxLength
    m_bExtendedUI = DEF_ExtendedUI
    m_bRedraw = True

    Call pvCreate

End Sub

Private Function VersionCheck() As Boolean

'/* nt version chck

Dim tVer    As OSVERSIONINFO

    With tVer
        .dwVersionInfoSize = LenB(tVer)
        GetVersionEx tVer
        m_bIsNt = ((.dwPlatformId And VER_PLATFORM_WIN32_NT) = VER_PLATFORM_WIN32_NT)
    End With

    VersionCheck = m_bIsNt

End Function

Private Sub UserControl_LostFocus()

    m_bInFocus = False

End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

'---------------------------------------------------------------------------------------
' Date      : 9/09/05
' Purpose   : Read property values from a previously persisted instance.
'---------------------------------------------------------------------------------------

    Set Font = PropBag.ReadProperty("Font", m_oFont)

    m_eStyle = PropBag.ReadProperty(PROP_Style, DEF_Style)
    m_eExStyle = PropBag.ReadProperty(PROP_ExStyle, DEF_ExStyle)
    m_lDroppedHeight = PropBag.ReadProperty(PROP_DroppedHeight, DEF_DroppedHeight)
    m_lDroppedWidth = PropBag.ReadProperty(PROP_DroppedWidth, DEF_DroppedWidth)
    UserControl.Enabled = PropBag.ReadProperty(PROP_Enabled, DEF_Enabled)
    m_lMaxLength = PropBag.ReadProperty(PROP_MaxLength, DEF_MaxLength)
    m_bExtendedUI = PropBag.ReadProperty(PROP_ExtendedUI, DEF_ExtendedUI)
    m_bRedraw = True
    m_lImageSize = PropBag.ReadProperty(PROP_ImageSize, DEF_ImageSize)

    Call pvCreate

End Sub

Private Sub UserControl_Resize()

    Call pvResize

End Sub

Private Sub UserControl_Terminate()
    
    If (m_bSubclass) Then
        Call mIOIPAComboBoxEx.TerminateIPAO(m_uIPAO)
        Call ssc_Terminate
        Call pvDestroyImageList
        Call pvDestroyFont
    End If

    Call pvDestroyWindow(m_hWndCombo)
    Call pvDestroyWindow(m_hWndEdit)
    Call pvDestroyWindow(m_hWnd)

    If Not (m_lhMod = 0) Then
        FreeLibrary m_lhMod
        m_lhMod = 0
    End If
    Exit Sub
    
  If Err.Number <> 0 Then
     MsgBox "Error " & Err.Number & vbCrLf & Err.Description, vbCritical, "ucComboBoxEx Terminate Error"
     Exit Sub
 End If
 
End Sub

Private Function pvDestroyWindow(ByVal hWnd As Long) As Boolean

    If (hWnd) Then
        If (DestroyWindow(hWnd)) Then
            pvDestroyWindow = True
            hWnd = 0
        End If
    End If

End Function

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

'---------------------------------------------------------------------------------------
' Date      : 9/09/05
' Purpose   : Persist our property values.
'---------------------------------------------------------------------------------------

    PropBag.WriteProperty "Font", Font
    PropBag.WriteProperty PROP_Style, m_eStyle, DEF_Style
    PropBag.WriteProperty PROP_ExStyle, m_eExStyle, DEF_ExStyle
    PropBag.WriteProperty PROP_DroppedHeight, m_lDroppedHeight, DEF_DroppedHeight
    PropBag.WriteProperty PROP_DroppedWidth, m_lDroppedWidth, DEF_DroppedWidth
    PropBag.WriteProperty PROP_Enabled, UserControl.Enabled, DEF_Enabled
    PropBag.WriteProperty PROP_MaxLength, m_lMaxLength, DEF_MaxLength
    PropBag.WriteProperty PROP_ExtendedUI, m_bExtendedUI, DEF_ExtendedUI
    PropBag.WriteProperty PROP_ImageSize, m_lImageSize, DEF_ImageSize

End Sub

Private Function pvShiftState() As Integer

Dim lS As Integer

    If (GetAsyncKeyState(vbKeyShift) < 0) Then
        lS = lS Or vbShiftMask
    End If
    If (GetAsyncKeyState(vbKeyMenu) < 0) Then
        lS = lS Or vbAltMask
    End If
    If (GetAsyncKeyState(vbKeyControl) < 0) Then
        lS = lS Or vbCtrlMask
    End If
    pvShiftState = lS

End Function

'========================================================================================
' OLEInPlaceActiveObject interface
'========================================================================================

Private Sub pvSetIPAO()

Dim pOleObject          As IOleObject
Dim pOleInPlaceSite     As IOleInPlaceSite
Dim pOleInPlaceFrame    As IOleInPlaceFrame
Dim pOleInPlaceUIWindow As IOleInPlaceUIWindow
Dim rcPos               As RECT
Dim rcClip              As RECT
Dim uFrameInfo          As OLEINPLACEFRAMEINFO

    On Error Resume Next

        Set pOleObject = Me
        Set pOleInPlaceSite = pOleObject.GetClientSite

        If (Not pOleInPlaceSite Is Nothing) Then
            Call pOleInPlaceSite.GetWindowContext(pOleInPlaceFrame, pOleInPlaceUIWindow, VarPtr(rcPos), VarPtr(rcClip), VarPtr(uFrameInfo))
            If (Not pOleInPlaceFrame Is Nothing) Then
                Call pOleInPlaceFrame.SetActiveObject(m_uIPAO.ThisPointer, vbNullString)
            End If
            If (Not pOleInPlaceUIWindow Is Nothing) Then '-- And Not m_bMouseActivate
                Call pOleInPlaceUIWindow.SetActiveObject(m_uIPAO.ThisPointer, vbNullString)
            Else
                Call pOleObject.DoVerb(OLEIVERB_UIACTIVATE, 0, pOleInPlaceSite, 0, UserControl.hWnd, VarPtr(rcPos))
            End If
        End If

    On Error GoTo 0

End Sub

Friend Function frTranslateAccel(pMsg As Msg) As Boolean

Dim pOleObject      As IOleObject
Dim pOleControlSite As IOleControlSite
Dim lhWndFocus As Long
Dim bToEdit As Boolean
Dim iShift As Integer
Dim iSel As Long
Dim iLen As Long
'Purpose   : Intercept arrow keys, home/end/pageup/pagedown and return keys.

    On Error Resume Next

        Select Case pMsg.message

        Case WM_KEYDOWN, WM_KEYUP

            Select Case pMsg.wParam

            Case vbKeyTab

                If (pvShiftState() And vbCtrlMask) Then
                    Set pOleObject = Me
                    Set pOleControlSite = pOleObject.GetClientSite
                    If (Not pOleControlSite Is Nothing) Then
                        Call pOleControlSite.TranslateAccelerator(VarPtr(pMsg), pvShiftState() And vbShiftMask)
                    End If
                End If
                frTranslateAccel = False

            Case vbKeyUp, vbKeyDown, vbKeyLeft, vbKeyRight, vbKeyPageDown, vbKeyPageUp, vbKeyHome, vbKeyEnd, vbKeyReturn, vbKeyEscape

                If (pMsg.wParam = vbKeyReturn) Or (pMsg.wParam = vbKeyEscape) Then
                    'only eat the return/esc keys if the combo is dropped
                    If SendMessageLongA(m_hWnd, CB_GETDROPPEDSTATE, 0, 0) = 0 Then Exit Function
                    SendMessageLongA m_hWndCombo, pMsg.message, pMsg.wParam, pMsg.lParam
                    frTranslateAccel = True
                    Exit Function
                End If

                bToEdit = (GetFocus() = m_hWndEdit)
                If m_eStyle = cboDropDownCombo Then
                    If pMsg.wParam = vbKeyHome Or pMsg.wParam = vbKeyEnd Then 'Or pMsg.wParam = vbKeyReturn Then
                        If Dropped Then
                            If pMsg.wParam = vbKeyHome Then
                                iShift = pvShiftState()
                                If (iShift And vbShiftMask) = vbShiftMask Then
                                    iSel = SelStart
                                    SelStart = 0
                                    If iSel > 0 Then
                                        SelLength = iSel + 1
                                    End If
                                Else
                                    SelStart = 0
                                    SelLength = 0
                                End If
                                frTranslateAccel = True
                                Exit Function
                            ElseIf pMsg.wParam = vbKeyEnd Then
                                iShift = pvShiftState()
                                If (iShift And vbShiftMask) = vbShiftMask Then
                                    iSel = SelStart
                                    iLen = Len(Text)
                                    If iLen - iSel >= 0 Then
                                        SelLength = iLen - iSel
                                    End If
                                Else
                                    pvSetSelStartEnd Len(Text), Len(Text)
                                End If
                                frTranslateAccel = True
                                Exit Function
                            Else
                                bToEdit = True
                            End If
                        End If
                    End If
                End If
                If bToEdit Then
                    SendMessageLongA m_hWndEdit, pMsg.message, pMsg.wParam, pMsg.lParam
                    bToEdit = False
                Else
                    SendMessageLongA m_hWndCombo, pMsg.message, pMsg.wParam, pMsg.lParam
                End If

                frTranslateAccel = True
            End Select
        End Select

    On Error GoTo 0

End Function

Private Sub pvSetSelStartEnd(ByVal lStart As Long, ByVal lEnd As Long)

Dim lR As Long
' Set the start and end of the selection in the edit
' box portion of a drop down combo box:

    If Not (m_hWnd = 0) Then
        lStart = lStart And &H7FFF&
        lEnd = lEnd And &H7FFF&
        If m_bIsNt Then
            lR = SendMessageLongW(m_hWndEdit, EM_SETSEL, lStart, lEnd)
        Else
            lR = SendMessageLongA(m_hWndEdit, EM_SETSEL, lStart, lEnd)
        End If
    End If

End Sub

Private Sub pvCreate()

Dim lExStyle  As Long
Dim dwStyle As Long
Dim lWidth As Long
Dim lHeight As Long

    m_hWndParent = UserControl.hWnd

    Call pvDestroy

    Select Case m_eStyle
    Case cboSimple
        dwStyle = dwStyle Or CBS_SIMPLE
    Case cboDropDownList
        dwStyle = dwStyle Or CBS_DROPDOWNLIST
    Case cboDropDownCombo
        dwStyle = dwStyle Or CBS_DROPDOWN
    Case Else
        Debug.Assert False
        dwStyle = dwStyle Or CBS_DROPDOWN
    End Select

    lWidth = UserControl.ScaleWidth \ Screen.TwipsPerPixelX
    lHeight = (UserControl.ScaleHeight \ Screen.TwipsPerPixelX) * 8

    If m_bIsNt Then
        m_hWnd = CreateWindowExW(0, StrPtr(WC_COMBOBOXEX), StrPtr(vbNullString), dwStyle Or WS_CHILD Or CBS_AUTOHSCROLL, _
                 0, 0, lWidth, lHeight, m_hWndParent, _
                 0, App.hInstance, ByVal 0&)
    Else
        m_hWnd = CreateWindowExA(0, WC_COMBOBOXEX, vbNullString, dwStyle Or WS_CHILD Or CBS_AUTOHSCROLL, _
                 0, 0, lWidth, lHeight, m_hWndParent, _
                 0, App.hInstance, ByVal 0&)
    End If

    If m_hWnd Then

        If m_bIsNt Then
            Call SendMessageLongW(m_hWnd, CCM_SETUNICODEFORMAT, 1&, 0&)
            Debug.Assert SendMessageLongW(m_hWnd, CCM_GETUNICODEFORMAT, 0&, 0&) <> 0&
        End If

        SendMessageLongA m_hWnd, CB_SETDROPPEDWIDTH, m_lDroppedWidth, 0
        SendMessageLongA m_hWnd, CB_LIMITTEXT, m_lMaxLength, 0
        SendMessageLongA m_hWnd, CB_SETEXTENDEDUI, -m_bExtendedUI, 0

        Call pvSubclass

        '-- Initialize font
        Set Font = m_oFont

        '-- Initialize imagelist
        m_hImageList = ImageList_Create(m_lImageSize, m_lImageSize, ILC_MASK Or ILC_COLORDDB, 0, 0)

        SendMessageLongA m_hWnd, CBEM_SETEXSTYLE, 0&, ByVal m_eExStyle

        If (m_hImageList) Then
            If Not (m_eExStyle = eccxNoImages Or eccxNoEditImage) Then
                Call SendMessageLongA(m_hWnd, CBEM_SETIMAGELIST, 0, m_hImageList)
            End If
        End If

        '-- Ensure window over parent

        SetParent m_hWnd, m_hWndParent
        ShowWindow m_hWnd, SW_SHOWNORMAL
        EnableWindow m_hWnd, Abs(Enabled)
        SendMessageLongA m_hWnd, WM_SETREDRAW, -m_bRedraw, 0

        Call pvResize

    End If

End Sub

Private Sub pvSubclass()

    If (m_bSubclass = False) Then

        If (m_hWnd) Then
            '-- Only on run-time
            If UserControl.Ambient.UserMode Then
                '-- Initialize IOLEInPlaceActiveObject
                Call mIOIPAComboBoxEx.InitIPAO(m_uIPAO, Me)

                Call ssc_Subclass(m_hWndParent, , 1, , , True)
                Call ssc_AddMsg(m_hWndParent, WM_SETFOCUS)
                Call ssc_AddMsg(m_hWndParent, WM_SIZE)
                Call ssc_AddMsg(m_hWndParent, WM_NOTIFY, MSG_BEFORE)
                Call ssc_AddMsg(m_hWndParent, WM_COMMAND, MSG_BEFORE)

                Call ssc_Subclass(m_hWnd, , 1, , , True)
                Call ssc_AddMsg(m_hWnd, WM_CTLCOLORLISTBOX)

                m_hWndCombo = SendMessageLongA(m_hWnd, CBEM_GETCOMBOCONTROL, 0, 0)
                Call ssc_Subclass(m_hWndCombo, , 1, , , True)
                Call ssc_AddMsg(m_hWndCombo, WM_SETFOCUS)
                Call ssc_AddMsg(m_hWndCombo, WM_MOUSEACTIVATE, MSG_BEFORE)

                Select Case m_eStyle
                Case cboDropDownCombo
                    m_hWndEdit = SendMessageLongA(m_hWnd, CBEM_GETEDITCONTROL, 0, 0)
                    Call ssc_Subclass(m_hWndEdit, , 1, , , True)
                    Call ssc_AddMsg(m_hWndEdit, WM_MOUSEACTIVATE, MSG_BEFORE)
                    Call ssc_AddMsg(m_hWndEdit, WM_SETFOCUS)
                    Call ssc_AddMsg(m_hWndEdit, WM_KILLFOCUS)
                    Call ssc_AddMsg(m_hWndEdit, WM_CHAR, MSG_BEFORE)
                    Call ssc_AddMsg(m_hWndEdit, WM_KEYUP)
                    Call ssc_AddMsg(m_hWndCombo, WM_CTLCOLOREDIT)

                Case cboSimple
                    m_hWndEdit = FindWindowEx(m_hWndParent, ByVal 0&, "Edit", ByVal 0&)
                    If m_hWndEdit Then
                        Call ssc_Subclass(m_hWndEdit, , 1, , , True)
                        Call ssc_AddMsg(m_hWndEdit, WM_MOUSEACTIVATE, MSG_BEFORE)
                        Call ssc_AddMsg(m_hWndEdit, WM_SETFOCUS)
                        Call ssc_AddMsg(m_hWndEdit, WM_CHAR, MSG_BEFORE)
                        Call ssc_AddMsg(m_hWndEdit, WM_KEYUP)
                        Call ssc_AddMsg(m_hWndCombo, WM_CTLCOLOREDIT)
                    End If
                Case Else
                    Call ssc_AddMsg(m_hWndCombo, WM_KEYDOWN)
                    Call ssc_AddMsg(m_hWndCombo, WM_CHAR)
                    Call ssc_AddMsg(m_hWndCombo, WM_KEYUP)
                End Select
            End If
            m_bSubclass = True
        End If
    End If

End Sub

Private Sub pvSubclassStop()

    If (m_hWnd) Then

        Select Case m_eStyle
        Case cboDropDownCombo, cboSimple
            If m_hWndEdit Then
                Call ssc_UnSubclass(m_hWndEdit)
                If m_hWndCombo Then
                    Call ssc_UnSubclass(m_hWndCombo)
                End If
                Call ssc_UnSubclass(m_hWnd)
                Call ssc_UnSubclass(m_hWndParent)
            End If
        Case Else
            If m_hWndCombo Then
                Call ssc_UnSubclass(m_hWndCombo)
                Call ssc_UnSubclass(m_hWnd)
                Call ssc_UnSubclass(m_hWndParent)
            End If
        End Select

    End If

End Sub

Private Sub pvDestroy()

    If m_bSubclass Then
        Call ssc_Terminate
        m_bSubclass = False
        Call pvDestroyWindow(m_hWndCombo)
        Call pvDestroyWindow(m_hWndEdit)
        Call pvDestroyWindow(m_hWnd)
        Call pvDestroyFont
    End If

End Sub

Private Function MakeLong(ByVal iLower As Integer, ByVal iUpper As Integer) As Long

'---------------------------------------------------------------------------------------
' Date      : 1/17/05
' Purpose   : Combine two words into a dword.
'---------------------------------------------------------------------------------------

    MakeLong = iLower Or (iUpper * &H10000)

End Function

Private Sub pvResize()

Dim tR As RECT
Dim bDesignMode As Boolean
Dim lHeight As Long

    bDesignMode = Not UserControl.Ambient.UserMode

    If m_hWnd Then
        If Not (m_eStyle = cboSimple) Then
            If bDesignMode Then
                ' Make sure the User Control's height is correct:
                lHeight = SendMessageLongA(m_hWnd, CB_GETITEMHEIGHT, -1, 0)
                UserControl.Extender.Height = ScaleY((lHeight + 6), vbPixels, vbContainerSize)
            End If
        End If
        GetClientRect UserControl.hWnd, tR
        MoveWindow m_hWnd, 0, 0, tR.Right - tR.Left, tR.Bottom - tR.Top, 1
        If m_eStyle <> cboSimple Then
            lHeight = tR.Bottom - tR.Top + 2 + SendMessageLongA(m_hWnd, CB_GETITEMHEIGHT, 0, 0) * 8
        Else
            lHeight = tR.Bottom - tR.Top
        End If
        MoveWindow m_hWndCombo, 0, 0, tR.Right - tR.Left, lHeight, 1
    End If

End Sub

Private Sub pvPropChanged(ByRef s As String)

    If Ambient.UserMode = False Then PropertyChanged s

End Sub

Private Property Get pItem_Text(ByVal iIndex As Long) As String

Dim a(260) As Byte
Dim lLen   As Long

    If m_hWnd Then
        With m_tItem
            .mask = CBEIF_TEXT
            .pszText = VarPtr(a(0))
            .cchTextMax = UBound(a)
            .iItem = iIndex
        End With
        If m_bIsNt Then
            lLen = SendMessageW(m_hWnd, CBEM_GETITEMW, 0, m_tItem)
        Else
            lLen = SendMessageA(m_hWnd, CBEM_GETITEMA, 0, m_tItem)
        End If
        If lLen > 0 Then
            If m_bIsNt Then
                pItem_Text = a
            Else
                pItem_Text = Left$(StrConv(a(), vbUnicode), lLen)
            End If
        Else
            pItem_Text = ""
        End If

    End If

End Property

Private Property Let pItem_Text(ByVal iIndex As Long, ByVal sText As String)

'---------------------------------------------------------------------------------------
' Date      : 9/9/05
' Purpose   : Set the text of a combo list item.
'---------------------------------------------------------------------------------------

    If m_hWnd Then
        With m_tItem
            .mask = CBEIF_TEXT
            .iItem = iIndex
            If m_bIsNt Then
                .pszText = StrPtr(sText)
            Else
                .pszText = StrPtr(StrConv(sText & vbNullChar, vbFromUnicode))
            End If
            .cchTextMax = Len(sText)
        End With
        If m_bIsNt Then
            SendMessageW m_hWnd, CBEM_SETITEMW, 0, m_tItem
        Else
            SendMessageA m_hWnd, CBEM_SETITEMA, 0, m_tItem
        End If
    End If

End Property

Private Property Get pItem_Info(ByVal iIndex As Long, ByVal iMask As Long) As Long

'---------------------------------------------------------------------------------------
' Date      : 9/9/05
' Purpose   : Get a 32 bit value in the COMBOBOXEXITEM structure.
'---------------------------------------------------------------------------------------

    If m_hWnd Then
        With m_tItem
            .mask = iMask
            .iItem = iIndex
            If m_bIsNt Then
                If SendMessageW(m_hWnd, CBEM_GETITEMW, 0, m_tItem) Then
                    If iMask = CBEIF_LPARAM Then
                        pItem_Info = .lParam
                    ElseIf iMask = CBEIF_IMAGE Then
                        pItem_Info = .iImage
                    ElseIf iMask = CBEIF_SELECTEDIMAGE Then
                        pItem_Info = .iSelectedImage
                    ElseIf iMask = CBEIF_INDENT Then
                        pItem_Info = .iIndent
                    End If
                End If
            Else
                If SendMessageA(m_hWnd, CBEM_GETITEMA, 0, m_tItem) Then
                    If iMask = CBEIF_LPARAM Then
                        pItem_Info = .lParam
                    ElseIf iMask = CBEIF_IMAGE Then
                        pItem_Info = .iImage
                    ElseIf iMask = CBEIF_SELECTEDIMAGE Then
                        pItem_Info = .iSelectedImage
                    ElseIf iMask = CBEIF_INDENT Then
                        pItem_Info = .iIndent
                    End If
                End If
            End If
        End With
    End If

End Property

Private Property Let pItem_Info(ByVal iIndex As Long, ByVal iMask As Long, ByVal iNew As Long)

'---------------------------------------------------------------------------------------
' Date      : 9/9/05
' Purpose   : Set a 32 bit value in the COMBOBOXEXITEM structure.
'---------------------------------------------------------------------------------------

    If m_hWnd Then
        With m_tItem
            .mask = iMask
            .iItem = iIndex
            If iMask = CBEIF_LPARAM Then
                .lParam = iNew
            ElseIf iMask = CBEIF_IMAGE Then
                .iImage = iNew
            ElseIf iMask = CBEIF_SELECTEDIMAGE Then
                .iSelectedImage = iNew
            ElseIf iMask = CBEIF_INDENT Then
                .iIndent = iNew
            End If
            If m_bIsNt Then
                SendMessageW m_hWnd, CBEM_SETITEMW, 0, m_tItem
            Else
                SendMessageA m_hWnd, CBEM_SETITEMA, 0, m_tItem
            End If
        End With
    End If

End Property

Public Property Get SelLength() As Long

Dim lEnd    As Long
Dim lStart  As Long

    If m_hWndEdit Then
        SendMessageLongA m_hWndEdit, EM_GETSEL, VarPtr(lStart), VarPtr(lEnd)
        SelLength = lEnd - lStart
    End If

End Property

Public Property Let SelLength(ByVal lNew As Long)

Dim lStart As Long
Dim lEnd As Long

    If m_hWndEdit Then
        SendMessageLongA m_hWndEdit, EM_GETSEL, VarPtr(lStart), VarPtr(lEnd)
        SendMessageLongA m_hWndEdit, EM_SETSEL, lStart, lStart + lNew
    End If

End Property

Public Property Get SelEnd() As Long

'---------------------------------------------------------------------------------------
' Date      : 9/9/05
' Purpose   : Get the selend of the edit portion (if any) of the TextBox.
'---------------------------------------------------------------------------------------

Dim i As Long

    If m_hWndEdit Then
        SendMessageLongA m_hWndEdit, EM_GETSEL, VarPtr(i), VarPtr(SelEnd)
    End If

End Property

Public Property Let SelEnd(ByVal lNew As Long)

'---------------------------------------------------------------------------------------
' Date      : 9/9/05
' Purpose   : Set the selend of the edit portion (if any) of the TextBox.
'---------------------------------------------------------------------------------------

    If m_hWndEdit Then
Dim lStart As Long, lEnd As Long
        SendMessageLongA m_hWndEdit, EM_GETSEL, VarPtr(lStart), VarPtr(lEnd)
        If lStart < lNew _
           Then SendMessageLongA m_hWndEdit, EM_SETSEL, lStart, lNew _
           Else SendMessageLongA m_hWndEdit, EM_SETSEL, lNew, lNew
    End If

End Property

Public Property Get SelStart() As Long

Dim lEnd    As Long
Dim lStart  As Long

    If m_hWndEdit Then
        SendMessageLongA m_hWndEdit, EM_GETSEL, VarPtr(SelStart), VarPtr(lEnd)
    End If

End Property

Public Property Let SelStart(ByVal lNew As Long)

Dim lEnd    As Long
Dim lStart  As Long

    If m_hWndEdit Then
        SendMessageA m_hWndEdit, EM_GETSEL, VarPtr(lStart), VarPtr(lEnd)
        If lEnd > lNew _
           Then SendMessageLongA m_hWndEdit, EM_SETSEL, lNew, lEnd _
           Else SendMessageLongA m_hWndEdit, EM_SETSEL, lNew, lNew
    End If

End Property

Public Property Get SelText() As String

Dim lStart      As Long
Dim lEnd        As Long

    If m_hWndEdit Then
        SendMessageLongA m_hWndEdit, EM_GETSEL, VarPtr(lStart), VarPtr(lEnd)
        On Error Resume Next
            SelText = Mid$(Text, lStart, (lEnd - lStart))
        On Error GoTo 0
    Else
        SelText = Text
    End If

End Property

Public Property Get Style() As eComboBoxExStyle

'---------------------------------------------------------------------------------------
' Date      : 9/9/05
' Purpose   : Return the style of combo box.
'---------------------------------------------------------------------------------------

    Style = m_eStyle

End Property

Public Property Let Style(ByVal iNew As eComboBoxExStyle)

'---------------------------------------------------------------------------------------
' Date      : 9/9/05
' Purpose   : Set the style of combo box.
'---------------------------------------------------------------------------------------

    If Not (m_eStyle = iNew) Then
        m_eStyle = iNew
        If Not (m_hWnd = 0) Then
            Call pvCreate
        End If
        pvPropChanged PROP_Style
    End If

End Property

Public Property Get ImageSize() As Long

'---------------------------------------------------------------------------------------
' Date      : 9/9/05
' Purpose   : Return the Image size of combo box.
'---------------------------------------------------------------------------------------

    ImageSize = m_lImageSize

End Property

Public Property Let ImageSize(ByVal iNew As Long)

'---------------------------------------------------------------------------------------
' Date      : 9/9/05
' Purpose   : Set the Image size  of combo box.
'---------------------------------------------------------------------------------------

    m_lImageSize = iNew
    pvPropChanged PROP_ImageSize

End Property

Public Property Get ExtendedStyle() As ECCXExtendedStyle

    ExtendedStyle = m_eExStyle

End Property

Public Property Let ExtendedStyle(ByVal iNew As ECCXExtendedStyle)

    m_eExStyle = iNew And 3&
    pvPropChanged PROP_ExStyle

End Property

Public Function AddBitmap(ByVal hBitmap As Long, Optional ByVal MaskColor As Long = CLR_NONE) As Long

    If (m_hWnd And m_hImageList) Then

        '-- Add bitmap/s to ComboBoxEx's imagelist
        If (MaskColor <> CLR_NONE) Then
            AddBitmap = ImageList_AddMasked(m_hImageList, hBitmap, MaskColor)
        Else
            AddBitmap = ImageList_Add(m_hImageList, hBitmap, 0)
        End If
    End If

End Function

Public Function AddIcon(ByVal hIcon As Long) As Long

    If (m_hWnd And m_hImageList) Then

        '-- Add icons to toolbar's imagelist
        AddIcon = ImageList_AddIcon(m_hImageList, hIcon)
    End If

End Function

Public Function AddItem( _
                        ByRef sText As String, _
                        Optional ByVal iIndexInsertBefore As Long = -1, _
                        Optional ByVal iIconIndex As Long = -1, _
                        Optional ByVal iIconIndexSelected As Long = -1, _
                        Optional ByVal iItemData As Long, _
                        Optional ByVal iIndent As Long) _
                        As Boolean

'---------------------------------------------------------------------------------------
' Date      : 9/9/05
' Purpose   : Add a combobox item.
'---------------------------------------------------------------------------------------

    If iIconIndexSelected < 0 Then iIconIndexSelected = iIconIndex

    If m_hWnd Then
        With m_tItem
            .mask = CBEIF_TEXT Or CBEIF_LPARAM Or CBEIF_IMAGE Or CBEIF_SELECTEDIMAGE Or CBEIF_INDENT Or CBEIF_LPARAM
            .iItem = iIndexInsertBefore
            .iImage = iIconIndex
            .iSelectedImage = iIconIndexSelected
            .iOverlay = -1
            .iIndent = iIndent
            .lParam = iItemData

            If m_bIsNt Then
                .pszText = StrPtr(sText)
            Else
                .pszText = StrPtr(StrConv(sText, vbFromUnicode))
            End If
            .cchTextMax = Len(sText)

            If iIndexInsertBefore < -1 Then iIndexInsertBefore = -1
            If m_bIsNt Then
                m_lNewIndex = SendMessageW(m_hWnd, CBEM_INSERTITEMW, iIndexInsertBefore, m_tItem)
            Else
                m_lNewIndex = SendMessageA(m_hWnd, CBEM_INSERTITEMA, iIndexInsertBefore, m_tItem)
            End If
            AddItem = CBool(m_lNewIndex > -1)
        End With
    End If

End Function

Public Function RemoveItem(ByVal iIndex As Long) As Boolean

'---------------------------------------------------------------------------------------
' Date      : 9/9/05
' Purpose   : Remove a combobox item.
'---------------------------------------------------------------------------------------

    If m_hWnd Then
        RemoveItem = CBool(SendMessageA(m_hWnd, CBEM_DELETEITEM, iIndex, 0) > -1)
        If RemoveItem Then
            If iIndex = m_lNewIndex Then
                m_lNewIndex = -1
            ElseIf iIndex < m_lNewIndex Then
                m_lNewIndex = m_lNewIndex - 1
            End If
        End If
    End If

End Function

Public Property Get DroppedWidth() As Single

    DroppedWidth = ScaleX(m_lDroppedWidth, vbPixels, vbContainerSize)

End Property

Public Property Let DroppedWidth(ByVal fNew As Single)

    m_lDroppedWidth = ScaleX(fNew, vbContainerSize, vbPixels)
    If m_hWnd Then
        SendMessageA m_hWnd, CB_SETDROPPEDWIDTH, m_lDroppedWidth, 0
    End If
    pvPropChanged PROP_DroppedWidth

End Property

Public Property Get DroppedHeight() As Single

    DroppedHeight = ScaleY(m_lDroppedHeight, vbPixels, vbContainerSize)

End Property

Public Property Let DroppedHeight(ByRef fNew As Single)

    m_lDroppedHeight = ScaleY(fNew, vbContainerSize, vbPixels)
    pvPropChanged PROP_DroppedHeight

End Property

Public Property Get Font() As StdFont

    Set Font = m_oFont

End Property

Public Property Set Font(ByVal New_Font As StdFont)

Dim uLF   As LOGFONT
Dim lChar As Long
Dim b()   As Byte

    If New_Font Is Nothing Then Exit Property

    ' Set the control's default font:
    Set UserControl.Font = New_Font

    With New_Font
        b = StrConv(.Name, vbFromUnicode)
        ReDim Preserve b(0 To 31) As Byte
        CopyMemory uLF.lfFaceName(0), b(0), 32&

        uLF.lfHeight = -MulDiv(.Size, GetDeviceCaps(UserControl.hDc, LOGPIXELSY), 72)
        uLF.lfItalic = .Italic
        uLF.lfWeight = IIf(.Bold, FW_BOLD, FW_NORMAL)
        uLF.lfUnderline = .Underline
        uLF.lfStrikeOut = .Strikethrough
        uLF.lfCharSet = .Charset
    End With
    Call pvDestroyFont: m_hFont = CreateFontIndirect(uLF)

    Call SendMessageLongA(m_hWnd, WM_SETFONT, m_hFont, ByVal 1&)

    ' Make sure the User Control's height is correct:
    If m_eStyle <> cboSimple Then
        UserControl.Extender.Height = ScaleY((SendMessageLongA(m_hWnd, CB_GETITEMHEIGHT, -1, 0) + 6), vbPixels, vbContainerSize) '* Screen.TwipsPerPixelY
    End If
    Set m_oFont = New_Font
    PropertyChanged "Font"

End Property

Public Property Get NewIndex() As Long

    NewIndex = m_lNewIndex

End Property

Public Property Get Enabled() As Boolean
Attribute Enabled.VB_UserMemId = -514

    Enabled = UserControl.Enabled

End Property

Public Property Let Enabled(ByVal bNew As Boolean)

    UserControl.Enabled = bNew
    pvPropChanged PROP_Enabled

End Property

Public Property Get Dropped() As Boolean
Attribute Dropped.VB_MemberFlags = "400"

    If m_eStyle > cboSimple Then
        If m_hWnd Then Dropped = CBool(SendMessageA(m_hWnd, CB_GETDROPPEDSTATE, 0, 0))
    End If

End Property

Public Property Let Dropped(ByVal bNew As Boolean)

    If m_eStyle > cboSimple Then
        If m_hWnd Then SendMessageA m_hWnd, CB_SHOWDROPDOWN, -bNew, 0
    End If

End Property

Public Property Get ListCount() As Long

    If m_hWnd Then ListCount = SendMessageLongA(m_hWnd, CB_GETCOUNT, 0, 0)

End Property

Public Property Get ListIndex() As Long
Attribute ListIndex.VB_MemberFlags = "400"

    If m_hWnd Then ListIndex = SendMessageLongA(m_hWnd, CB_GETCURSEL, 0, 0)

End Property

Public Property Let ListIndex(ByVal iNew As Long)

Dim lR As Long

    If m_hWnd Then
        lR = SendMessageLongA(m_hWnd, CB_SETCURSEL, iNew, 0)
        If lR = CB_ERR And iNew <> -1 Then
            Err.Raise 381, App.EXEName & ".ucComboBoxEx"
        End If
    End If

End Property

Public Property Get Text() As String

    If m_hWndEdit Then
        Text = pvGetWindowText(m_hWndEdit)
    ElseIf m_hWnd Then
        Text = pItem_Text(SendMessageA(m_hWnd, CB_GETCURSEL, 0, 0))
    End If

End Property

Public Property Let Text(ByRef sNew As String)

'---------------------------------------------------------------------------------------
' Date      : 9/9/05
' Purpose   : If there is an edit box, set the text.  Otherwise, search the list for an
'             item that matches sNew.  If found, set the listindex.  If not, raise an error.
'---------------------------------------------------------------------------------------

Dim lsAnsi As String

    lsAnsi = StrConv(sNew & vbNullChar, vbFromUnicode)
    If m_hWnd Then
        If m_hWndEdit Then
            Call pvSetWindowText(sNew)
        Else
Dim liIndex As Long
            liIndex = SendMessageLongA(m_hWnd, CB_FINDSTRINGEXACT, 0, StrPtr(lsAnsi))
            If liIndex > -1 Then
                SendMessageLongA m_hWnd, CB_SETCURSEL, liIndex, 0
            End If
        End If
    End If

End Property

Private Sub pvSetWindowText(ByVal sText As String)

Dim lPtr As Long

    If IsWindowUnicode(m_hWndEdit) Then
        If Len(sText) = 0 Then
            SetWindowTextW m_hWndEdit, StrPtr(vbNullString)
            Exit Sub
        End If
        lPtr = StrPtr(sText)
        SetWindowTextW m_hWndEdit, lPtr
    Else
        If Len(sText) = 0 Then
            SetWindowTextA m_hWndEdit, vbNullString
            Exit Sub
        End If
        SetWindowTextA m_hWndEdit, sText
    End If

End Sub

Private Function pvGetWindowText(ByVal hWnd As Long) As String

Dim lLen             As Long
Dim sBuf             As String

    lLen = 1 + pvGetTextLen(hWnd)

    If (lLen > 1) Then
        sBuf = String$(lLen, 0)
        If IsWindowUnicode(hWnd) Then
            GetWindowTextW hWnd, StrPtr(sBuf), lLen
        Else
            GetWindowTextA hWnd, sBuf, lLen
        End If
        pvGetWindowText = pvStripNulls(sBuf)
    Else
        pvGetWindowText = vbNullString
    End If

End Function

Private Function pvGetTextLen(ByVal hWnd As Long) As Long

' Get length of the caption

    If IsWindowUnicode(hWnd) Then
        pvGetTextLen = GetWindowTextLengthW(hWnd)
    Else
        pvGetTextLen = GetWindowTextLengthA(hWnd)
    End If

End Function

Public Property Get FindItem(ByRef sText As String, Optional ByVal bExact As Boolean) As Long

'---------------------------------------------------------------------------------------
' Purpose   : Search the list for an item and return the index if found.  -1 otherwise.
'Problem with Unicode String (Removed) by Zhu J.Y.
'---------------------------------------------------------------------------------------

Dim lS As String

    If m_hWnd Then
        lS = StrConv(sText & vbNullChar, vbFromUnicode)
        FindItem = SendMessageLongA(m_hWnd, IIf(bExact, CB_FINDSTRINGEXACT, CB_FINDSTRING), 0, StrPtr(lS))
    End If

End Property

Public Property Get FindItem_Unicode(ByVal sToFind As String, _
                                     Optional ByVal bExactMatch As Boolean) As Long

Dim lR As Long
Dim count As Long
' Find the index of the item sToFind, optionally
' exact matching. Return -1 if the item is not
' found.

    If Not (m_hWnd = 0) Then
        FindItem_Unicode = -1 'Set to not found
        count = ListCount
        If ListCount Then
            For lR = 1 To ListCount
                If bExactMatch Then
                    If List(lR) = sToFind Then
                        FindItem_Unicode = lR
                        Exit Property
                    End If
                Else
                    If InStr(List(lR), sToFind) Then
                        FindItem_Unicode = lR
                        Exit Property
                    End If
                End If
            Next lR
        End If
    End If

End Property

Public Sub Clear()

    If m_hWnd Then
        SendMessageA m_hWnd, CB_RESETCONTENT, 0, 0
        m_lNewIndex = -1
    End If

End Sub

Public Property Get ItemIndent(ByVal iIndex As Long) As Long

    ItemIndent = pItem_Info(iIndex, CBEIF_INDENT)

End Property

Public Property Let ItemIndent(ByVal iIndex As Long, ByVal iNew As Long)

    pItem_Info(iIndex, CBEIF_INDENT) = iNew

End Property

Public Property Get ItemIconIndex(ByVal iIndex As Long) As Long

    ItemIconIndex = pItem_Info(iIndex, CBEIF_IMAGE)

End Property

Public Property Let ItemIconIndex(ByVal iIndex As Long, ByVal iNew As Long)

    pItem_Info(iIndex, CBEIF_IMAGE) = iNew

End Property

Public Property Get ItemIconIndexSelected(ByVal iIndex As Long) As Long

    ItemIconIndexSelected = pItem_Info(iIndex, CBEIF_SELECTEDIMAGE)

End Property

Public Property Let ItemIconIndexSelected(ByVal iIndex As Long, ByVal iNew As Long)

    pItem_Info(iIndex, CBEIF_SELECTEDIMAGE) = iNew

End Property

Public Property Get ItemData(ByVal iIndex As Long) As Long

    ItemData = pItem_Info(iIndex, CBEIF_LPARAM)

End Property

Public Property Let ItemData(ByVal iIndex As Long, ByVal iNew As Long)

    pItem_Info(iIndex, CBEIF_LPARAM) = iNew

End Property

Public Property Get ItemText(ByVal iIndex As Long) As String

    ItemText = pItem_Text(iIndex)

End Property

Public Property Let ItemText(ByVal iIndex As Long, ByVal sNew As String)

    pItem_Text(iIndex) = sNew

End Property

Public Property Get List(ByVal iIndex As Long) As String

    List = pItem_Text(iIndex)

End Property

Public Property Let List(ByVal iIndex As Long, ByRef sNew As String)

    pItem_Text(iIndex) = sNew

End Property

Public Property Get MaxLength() As Long

' Same as MaxLength property of a Text control.  Only
' valid for drop down combo boxes:

    If Not (m_eStyle = cboDropDownList) Then
        MaxLength = m_lMaxLength
    Else
        'Err.Raise 383, "ucComboboxEx." & App.EXEName
    End If

End Property

Public Property Let MaxLength(ByVal iNew As Long)

' Same as MaxLength property of a Text control.  Only
' valid for drop down combo boxes:
' Purpose   : Set the maximum allowable number of characters in the edit portion of the window.

    If Not (m_eStyle = cboDropDownCombo) Then
        If iNew > &H7FFFFFFF Then iNew = &H7FFFFFFF
        m_lMaxLength = iNew
        If m_hWnd Then SendMessageA m_hWnd, CB_LIMITTEXT, m_lMaxLength, 0
        pvPropChanged PROP_MaxLength
    Else
        '
    End If

End Property

Public Property Get ExtendedUI() As Boolean

'---------------------------------------------------------------------------------------
' Date      : 9/9/05
' Purpose   : Return whether we are using the extended ui features available through comctl32.dll.
'---------------------------------------------------------------------------------------

    ExtendedUI = m_bExtendedUI

End Property

Public Property Let ExtendedUI(ByVal bNew As Boolean)

'---------------------------------------------------------------------------------------
' Date      : 9/9/05
' Purpose   : Set whether we are using the extended ui features available through comctl32.dll.
'---------------------------------------------------------------------------------------

    m_bExtendedUI = bNew
    pvPropChanged PROP_ExtendedUI
    If m_hWnd Then SendMessageA m_hWnd, CB_SETEXTENDEDUI, -m_bExtendedUI, 0

End Property

Public Property Get Redraw() As Boolean
Attribute Redraw.VB_MemberFlags = "400"

'---------------------------------------------------------------------------------------
' Date      : 9/9/05
' Purpose   : Return whether redrawing items is enabled.
'---------------------------------------------------------------------------------------

    Redraw = m_bRedraw

End Property

Public Property Let Redraw(ByVal bNew As Boolean)

'---------------------------------------------------------------------------------------
' Date      : 9/9/05
' Purpose   : Set whether redrawing items is enabled.  This can increase performance
'             when adding multiple items to a simple style combobox.
'---------------------------------------------------------------------------------------

    m_bRedraw = bNew
    If m_hWnd Then SendMessageA m_hWnd, WM_SETREDRAW, -m_bRedraw, 0

End Property

Public Property Get hWnd() As Long

'---------------------------------------------------------------------------------------
' Date      : 9/9/05
' Purpose   : Return the hwnd of the usercontrol.
'---------------------------------------------------------------------------------------

    hWnd = UserControl.hWnd

End Property

Public Property Get hWndComboEx() As Long

'---------------------------------------------------------------------------------------
' Date      : 9/9/05
' Purpose   : Return the hwnd of the ComboBoxEx.
'---------------------------------------------------------------------------------------

    hWndComboEx = m_hWnd

End Property

Public Property Get hWndCombo() As Long

'---------------------------------------------------------------------------------------
' Date      : 9/9/05
' Purpose   : Return the hwnd of the ComboBox.
'---------------------------------------------------------------------------------------

    hWndCombo = m_hWndCombo

End Property

Public Property Let ImageList(ByRef vThis As Variant)

Dim himl             As Long
'Dim lX               As Long

' Set the ImageList handle property either from a VB
' image list or directly:

    If VarType(vThis) = vbObject Then
        ' Assume VB ImageList control.  Note that unless
        ' some call has been made to an object within a
        ' VB ImageList the image list itself is not
        ' created.  Therefore hImageList returns error. So
        ' ensure that the ImageList has been initialised by
        ' drawing into nowhere:
        On Error Resume Next
            ' Get the image list initialised..
            vThis.ListImages(1).Draw 0, 0, 0, 1
            himl = vThis.hImageList
            If (Err.Number <> 0) Then
                Err.Clear
                himl = vThis.himl
                If (Err.Number <> 0) Then
                    himl = 0
                End If
            End If
        On Error GoTo 0
    ElseIf VarType(vThis) = vbLong Then
        ' Assume ImageList handle:
        himl = vThis
    Else
        Err.Raise vbObjectError + 1049, "vbalDriveCboEx." & App.EXEName, "ImageList property expects ImageList object or long hImageList handle."
    End If

    ' If we have a valid image list, then associate it with the control:
    If (himl <> 0) Then
        m_hIml = himl
        'ImageList_GetIconSize m_hIml, m_lIconSizeX, m_lIconSizeY
        'Set the Imagelist for the ComboBox
        SendMessageLongA m_hWnd, CBEM_SETIMAGELIST, 0, m_hIml
        Set Font = m_oFont
    End If

End Property

'========================================================================================
' Subclass code - The programmer may call any of the following Subclass_??? routines
'========================================================================================

'-------------------------------------------------------------------------------------------------
'========================================================================================================
' TO USE IDE-SAFE CALLBACKS...
'==============================
' 1. Include all common-use items in above "common-use" section
' 2. Include the following routines:  scb_ = self-callback
'     scb_SetCallbackAddr - used the same as AddressOf for class/form/uc/ppg routines
'  -  scb_ReleaseCallback - used to release memory for a specific callback
'  -  scb_TerminateCallbacks - used to terminate all callback addresses (release allocated memory)
'(-) can be removed if not needed. See their comments
' all of above can be made public if so desired
' 3. In your form/uc/class unload/terminate event, include the statement: scb_TerminateCallbacks
'
' Special notes:
'  a. scb_SetCallbackAddr uses function ordinals. Last routine in the class/form/uc/ppg is Ordinal #1, second to last is Ordinal #2, etc
'         IMPORTANT: Routines used for callbacks MUST NOT be Public, can be Private or Friend
'  b. Within scb_SetCallbackAddr, set the number of simultaneous callback addresses you will need
'  c. The callback function in your class/form/uc/ppg must be a Function that returns Long.
'     For example. The callback routine for a timer has no return value per MSDN, but when that timer
'     procedure is coded, the callback function must return a long. The return value is not important in that case
'  d. VB5 users. Search for zFnAddr("vba6", "EbMode") and replace with zFnAddr("vba5", "EbMode")
'  e. Do not terminate application while stopped in the callback procedure
'  f. At a minimum, review scb_SetCallbackAddr, scb_ReleaseCAllback & scb_TerminateCallbacks
'
'-SelfCallback specific declarations----------------------------------------------------------
' none
'-------------------------------------------------------------------------------------------------

Private Sub Class_Terminate()   ' sample terminate/unload event

    ssc_Terminate      '(add this to Unload or Terminate event if you are subclassing)
    shk_TerminateHooks  '(add this to Unload or Terminate event if you are hooking)
    scb_TerminateCallbacks '(add this to Unload or Terminate event if you are using callbacks)

End Sub

'-The following routines are exclusively for the ssc_subclass routines----------------------------

Private Function ssc_Subclass(ByVal lng_hWnd As Long, _
                              Optional ByVal lParamUser As Long = 0, _
                              Optional ByVal nOrdinal As Long = 1, _
                              Optional ByVal oCallback As Object = Nothing, _
                              Optional ByVal bIdeSafety As Boolean = True, _
                              Optional ByVal bUnicode As Boolean = False) As Boolean 'Subclass the specified window handle

'*************************************************************************************************
'* lng_hWnd   - Handle of the window to subclass
'* lParamUser - Optional, user-defined callback parameter
'* nOrdinal   - Optional, ordinal index of the callback procedure. 1 = last private method, 2 = second last private method, etc.
'* oCallback  - Optional, the object that will receive the callback. If undefined, callbacks are sent to this object's instance
'* bIdeSafety - Optional, enable/disable IDE safety measures. NB: you should really only disable IDE safety in a UserControl for design-time subclassing
'* bUnicode - Optional, if True, Unicode API calls will be made to the window vs ANSI calls
'*************************************************************************************************
'* cSelfSub - self-subclassing class template
'* Paul_Caton@hotmail.com
'* Copyright free, use and abuse as you see fit.
'*
'* v1.0 Re-write of the SelfSub/WinSubHook-2 submission to Planet Source Code............ 20060322
'* v1.1 VirtualAlloc memory to prevent Data Execution Prevention faults on Win64......... 20060324
'* v1.2 Thunk redesigned to handle unsubclassing and memory release...................... 20060325
'* v1.3 Data array scrapped in favour of property accessors.............................. 20060405
'* v1.4 Optional IDE protection added
'*      User-defined callback parameter added
'*      All user routines that pass in a hWnd get additional validation
'*      End removed from zError.......................................................... 20060411
'* v1.5 Added nOrdinal parameter to ssc_Subclass
'*      Switched machine-code array from Currency to Long................................ 20060412
'* v1.6 Added an optional callback target object
'*      Added an IsBadCodePtr on the callback address in the thunk prior to callback..... 20060413
'*************************************************************************************************
' Subclassing procedure must be declared identical to the one at the end of this class (Sample at Ordinal #1)

' \\LaVolpe - reworked routine a bit, revised the ASM to allow auto-unsubclass on WM_DESTROY

Dim z_Sc(0 To IDX_UNICODE) As Long                 'Thunk machine-code initialised here
Const CODE_LEN      As Long = 4 * IDX_UNICODE      'Thunk length in bytes

Const MEM_LEN       As Long = CODE_LEN + (8 * (MSG_ENTRIES))  'Bytes to allocate per thunk, data + code + msg tables
Const PAGE_RWX      As Long = &H40&                'Allocate executable memory
Const MEM_COMMIT    As Long = &H1000&              'Commit allocated memory
Const MEM_RELEASE   As Long = &H8000&              'Release allocated memory flag
Const IDX_EBMODE    As Long = 3                    'Thunk data index of the EbMode function address
Const IDX_CWP       As Long = 4                    'Thunk data index of the CallWindowProc function address
Const IDX_SWL       As Long = 5                    'Thunk data index of the SetWindowsLong function address
Const IDX_FREE      As Long = 6                    'Thunk data index of the VirtualFree function address
Const IDX_BADPTR    As Long = 7                    'Thunk data index of the IsBadCodePtr function address
Const IDX_OWNER     As Long = 8                    'Thunk data index of the Owner object's vTable address
Const IDX_CALLBACK  As Long = 10                   'Thunk data index of the callback method address
Const IDX_EBX       As Long = 16                   'Thunk code patch index of the thunk data
Const GWL_WNDPROC   As Long = -4                   'SetWindowsLong WndProc index
Const WNDPROC_OFF   As Long = &H38                 'Thunk offset to the WndProc execution address
Const SUB_NAME      As String = "ssc_Subclass"     'This routine's name

Dim nAddr         As Long
Dim nID           As Long
Dim nMyID         As Long

    If IsWindow(lng_hWnd) = 0 Then                      'Ensure the window handle is valid
        zError SUB_NAME, "Invalid window handle"
        Exit Function
    End If

    nMyID = GetCurrentProcessId                         'Get this process's ID
    GetWindowThreadProcessId lng_hWnd, nID              'Get the process ID associated with the window handle
    If nID <> nMyID Then                                'Ensure that the window handle doesn't belong to another process
        zError SUB_NAME, "Window handle belongs to another process"
        Exit Function
    End If

    If oCallback Is Nothing Then Set oCallback = Me     'If the user hasn't specified the callback owner

    nAddr = zAddressOf(oCallback, nOrdinal)             'Get the address of the specified ordinal method
    If nAddr = 0 Then                                   'Ensure that we've found the ordinal method
        zError SUB_NAME, "Callback method not found"
        Exit Function
    End If

    z_ScMem = VirtualAlloc(0, MEM_LEN, MEM_COMMIT, PAGE_RWX) 'Allocate executable memory

    If z_ScMem <> 0 Then                                  'Ensure the allocation succeeded

        If z_scFunk Is Nothing Then Set z_scFunk = New Collection 'If this is the first time through, do the one-time initialization
        On Error GoTo CatchDoubleSub                              'Catch double subclassing
        z_scFunk.Add z_ScMem, "h" & lng_hWnd                    'Add the hWnd/thunk-address to the collection
        On Error GoTo 0

        ' \\Tai Chi Minh Ralph Eastwood - fixed bug where the MSG_AFTER was not being honored
        ' \\LaVolpe - modified thunks to allow auto-unsubclassing when WM_DESTROY received
        z_Sc(14) = &HD231C031: z_Sc(15) = &HBBE58960: z_Sc(16) = &H12345678: z_Sc(17) = &HF63103FF: z_Sc(18) = &H750C4339: z_Sc(19) = &H7B8B4A38: z_Sc(20) = &H95E82C: z_Sc(21) = &H7D810000: z_Sc(22) = &H228&: z_Sc(23) = &HC70C7500: z_Sc(24) = &H20443: z_Sc(25) = &H5E90000: z_Sc(26) = &H39000000: z_Sc(27) = &HF751475: z_Sc(28) = &H25E8&: z_Sc(29) = &H8BD23100: z_Sc(30) = &H6CE8307B: z_Sc(31) = &HFF000000: z_Sc(32) = &H10C2610B: z_Sc(33) = &HC53FF00: z_Sc(34) = &H13D&: z_Sc(35) = &H85BE7400: z_Sc(36) = &HE82A74C0: z_Sc(37) = &H2&: z_Sc(38) = &H75FFE5EB: z_Sc(39) = &H2C75FF30: z_Sc(40) = &HFF2875FF: z_Sc(41) = &H73FF2475: z_Sc(42) = &H1053FF24: z_Sc(43) = &H811C4589: z_Sc(44) = &H13B&: z_Sc(45) = &H39727500:
        z_Sc(46) = &H6D740473: z_Sc(47) = &H2473FF58: z_Sc(48) = &HFFFFFC68: z_Sc(49) = &H873FFFF: z_Sc(50) = &H891453FF: z_Sc(51) = &H7589285D: z_Sc(52) = &H3045C72C: z_Sc(53) = &H8000&: z_Sc(54) = &H8920458B: z_Sc(55) = &H4589145D: z_Sc(56) = &HC4816124: z_Sc(57) = &H4&: z_Sc(58) = &H8B1862FF: z_Sc(59) = &H853AE30F: z_Sc(60) = &H810D78C9: z_Sc(61) = &H4C7&: z_Sc(62) = &H28458B00: z_Sc(63) = &H2975AFF2: z_Sc(64) = &H2873FF52: z_Sc(65) = &H5A1C53FF: z_Sc(66) = &H438D1F75: z_Sc(67) = &H144D8D34: z_Sc(68) = &H1C458D50: z_Sc(69) = &HFF3075FF: z_Sc(70) = &H75FF2C75: z_Sc(71) = &H873FF28: z_Sc(72) = &HFF525150: z_Sc(73) = &H53FF2073: z_Sc(74) = &HC328C328

        z_Sc(IDX_EBX) = z_ScMem                                                 'Patch the thunk data address
        z_Sc(IDX_INDEX) = lng_hWnd                                               'Store the window handle in the thunk data
        z_Sc(IDX_BTABLE) = z_ScMem + CODE_LEN                                   'Store the address of the before table in the thunk data
        z_Sc(IDX_ATABLE) = z_ScMem + CODE_LEN + ((MSG_ENTRIES + 1) * 4)         'Store the address of the after table in the thunk data
        z_Sc(IDX_OWNER) = ObjPtr(oCallback)                                     'Store the callback owner's object address in the thunk data
        z_Sc(IDX_CALLBACK) = nAddr                                              'Store the callback address in the thunk data
        z_Sc(IDX_PARM_USER) = lParamUser                                        'Store the lParamUser callback parameter in the thunk data

        ' \\LaVolpe - validate unicode request & cache unicode usage
        If bUnicode Then bUnicode = (IsWindowUnicode(lng_hWnd) <> 0&)
        z_Sc(IDX_UNICODE) = bUnicode                                            'Store whether the window is using unicode calls or not

        ' \\LaVolpe - added extra parameter "bUnicode" to the zFnAddr calls
        z_Sc(IDX_FREE) = zFnAddr("kernel32", "VirtualFree", bUnicode)           'Store the VirtualFree function address in the thunk data
        z_Sc(IDX_BADPTR) = zFnAddr("kernel32", "IsBadCodePtr", bUnicode)        'Store the IsBadCodePtr function address in the thunk data

        Debug.Assert zInIDE
        If bIdeSafety = True And z_IDEflag = 1 Then                             'If the user wants IDE protection
            z_Sc(IDX_EBMODE) = zFnAddr("vba6", "EbMode", bUnicode)                'Store the EbMode function address in the thunk data
        End If

        ' \\LaVolpe - use ANSI for non-unicode usage, else use WideChar calls
        If bUnicode Then
            z_Sc(IDX_CWP) = zFnAddr("user32", "CallWindowProcW", bUnicode)          'Store CallWindowProc function address in the thunk data
            z_Sc(IDX_SWL) = zFnAddr("user32", "SetWindowLongW", bUnicode)           'Store the SetWindowLong function address in the thunk data
            z_Sc(IDX_UNICODE) = 1
            RtlMoveMemory z_ScMem, VarPtr(z_Sc(0)), CODE_LEN                        'Copy the thunk code/data to the allocated memory
            nAddr = SetWindowLongW(lng_hWnd, GWL_WNDPROC, z_ScMem + WNDPROC_OFF)    'Set the new WndProc, return the address of the original WndProc
        Else
            z_Sc(IDX_CWP) = zFnAddr("user32", "CallWindowProcA", bUnicode)          'Store CallWindowProc function address in the thunk data
            z_Sc(IDX_SWL) = zFnAddr("user32", "SetWindowLongA", bUnicode)           'Store the SetWindowLong function address in the thunk data
            RtlMoveMemory z_ScMem, VarPtr(z_Sc(0)), CODE_LEN                        'Copy the thunk code/data to the allocated memory
            nAddr = SetWindowLongA(lng_hWnd, GWL_WNDPROC, z_ScMem + WNDPROC_OFF)    'Set the new WndProc, return the address of the original WndProc
        End If
        If nAddr = 0 Then                                                           'Ensure the new WndProc was set correctly
            zError SUB_NAME, "SetWindowLong failed, error #" & Err.LastDllError
            GoTo ReleaseMemory
        End If
        'Store the original WndProc address in the thunk data
        RtlMoveMemory z_ScMem + IDX_WNDPROC * 4, VarPtr(nAddr), 4&              ' z_Sc(IDX_WNDPROC) = nAddr
        ssc_Subclass = True                                                     'Indicate success
    Else
        zError SUB_NAME, "VirtualAlloc failed, error: " & Err.LastDllError
    End If

Exit Function                                                             'Exit ssc_Subclass

CatchDoubleSub:
    zError SUB_NAME, "Window handle is already subclassed"

ReleaseMemory:
    VirtualFree z_ScMem, 0, MEM_RELEASE                                       'ssc_Subclass has failed after memory allocation, so release the memory

End Function

'Terminate all subclassing

Private Sub ssc_Terminate()

' can be made public. Releases all subclassing
' can be removed and zTerminateThunks can be called directly

    zTerminateThunks SubclassThunk

End Sub

'UnSubclass the specified window handle

Private Sub ssc_UnSubclass(ByVal lng_hWnd As Long)

' can be made public. Releases a specific subclass
' can be removed and zUnThunk can be called directly

    zUnThunk lng_hWnd, SubclassThunk

End Sub

'Add the message value to the window handle's specified callback table

Private Sub ssc_AddMsg(ByVal lng_hWnd As Long, ByVal uMsg As Long, Optional ByVal When As eMsgWhen = MSG_AFTER)

' Note: can be removed if not needed and zAddMsg can be called directly

    If IsBadCodePtr(zMap_VFunction(lng_hWnd, SubclassThunk)) = 0 Then                 'Ensure that the thunk hasn't already released its memory
        If When And MSG_BEFORE Then                                             'If the message is to be added to the before original WndProc table...
            zAddMsg uMsg, IDX_BTABLE                                              'Add the message to the before table
        End If
        If When And MSG_AFTER Then                                              'If message is to be added to the after original WndProc table...
            zAddMsg uMsg, IDX_ATABLE                                              'Add the message to the after table
        End If
    End If

End Sub

'Delete the message value from the window handle's specified callback table

Private Sub ssc_DelMsg(ByVal lng_hWnd As Long, ByVal uMsg As Long, Optional ByVal When As eMsgWhen = MSG_AFTER)

' Note: can be removed if not needed and zDelMsg can be called directly

    If IsBadCodePtr(zMap_VFunction(lng_hWnd, SubclassThunk)) = 0 Then                'Ensure that the thunk hasn't already released its memory
        If When And MSG_BEFORE Then                                             'If the message is to be deleted from the before original WndProc table...
            zDelMsg uMsg, IDX_BTABLE                                              'Delete the message from the before table
        End If
        If When And MSG_AFTER Then                                              'If the message is to be deleted from the after original WndProc table...
            zDelMsg uMsg, IDX_ATABLE                                              'Delete the message from the after table
        End If
    End If

End Sub

'Call the original WndProc

Private Function ssc_CallOrigWndProc(ByVal lng_hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

' \\LaVolpe - Use ANSI API calls for non-unicode usage else use WideChar calls
' Note: can be removed if you do not use this function inside of your window procedure

    If IsBadCodePtr(zMap_VFunction(lng_hWnd, SubclassThunk)) = 0 Then            'Ensure that the thunk hasn't already released its memory
        If zData(IDX_UNICODE) Then
            ssc_CallOrigWndProc = CallWindowProcW(zData(IDX_WNDPROC), lng_hWnd, uMsg, wParam, lParam) 'Call the original WndProc of the passed window handle parameter
        Else
            ssc_CallOrigWndProc = CallWindowProcA(zData(IDX_WNDPROC), lng_hWnd, uMsg, wParam, lParam) 'Call the original WndProc of the passed window handle parameter
        End If
    End If

End Function

'Get the subclasser lParamUser callback parameter

Private Function zGet_lParamUser(ByVal hWnd_Hook_ID As Long, vType As eThunkType) As Long

'Note: can be removed if you never need to retrieve/update your user-defined paramter. See ssc_Subclass

    If vType <> CallbackThunk Then
        If IsBadCodePtr(zMap_VFunction(hWnd_Hook_ID, vType)) = 0 Then        'Ensure that the thunk hasn't already released its memory
            zGet_lParamUser = zData(IDX_PARM_USER)                                'Get the lParamUser callback parameter
        End If
    End If

End Function

'Let the subclasser lParamUser callback parameter

Private Sub zSet_lParamUser(ByVal hWnd_Hook_ID As Long, vType As eThunkType, newValue As Long)

'Note: can be removed if you never need to retrieve/update your user-defined paramter. See ssc_Subclass

    If vType <> CallbackThunk Then
        If IsBadCodePtr(zMap_VFunction(hWnd_Hook_ID, vType)) = 0 Then          'Ensure that the thunk hasn't already released its memory
            zData(IDX_PARM_USER) = newValue                                         'Set the lParamUser callback parameter
        End If
    End If

End Sub

'Add the message to the specified table of the window handle

Private Sub zAddMsg(ByVal uMsg As Long, ByVal nTable As Long)

Dim nCount As Long                                                        'Table entry count
Dim nBase  As Long                                                        'Remember z_ScMem
Dim i      As Long                                                        'Loop index

    nBase = z_ScMem                                                           'Remember z_ScMem so that we can restore its value on exit
    z_ScMem = zData(nTable)                                                   'Map zData() to the specified table

    If uMsg = ALL_MESSAGES Then                                               'If ALL_MESSAGES are being added to the table...
        nCount = ALL_MESSAGES                                                   'Set the table entry count to ALL_MESSAGES
    Else
        nCount = zData(0)                                                       'Get the current table entry count
        If nCount >= MSG_ENTRIES Then                                           'Check for message table overflow
            zError "zAddMsg", "Message table overflow. Either increase the value of Const MSG_ENTRIES or use ALL_MESSAGES instead of specific message values"
            GoTo Bail
        End If

        For i = 1 To nCount                                                     'Loop through the table entries
            If zData(i) = 0 Then                                                  'If the element is free...
                zData(i) = uMsg                                                     'Use this element
                GoTo Bail                                                           'Bail
            ElseIf zData(i) = uMsg Then                                           'If the message is already in the table...
                GoTo Bail                                                           'Bail
            End If
        Next i                                                                  'Next message table entry

        nCount = i                                                              'On drop through: i = nCount + 1, the new table entry count
        zData(nCount) = uMsg                                                    'Store the message in the appended table entry
    End If

    zData(0) = nCount                                                         'Store the new table entry count
Bail:
    z_ScMem = nBase                                                           'Restore the value of z_ScMem

End Sub

'Delete the message from the specified table of the window handle

Private Sub zDelMsg(ByVal uMsg As Long, ByVal nTable As Long)

Dim nCount As Long                                                        'Table entry count
Dim nBase  As Long                                                        'Remember z_ScMem
Dim i      As Long                                                        'Loop index

    nBase = z_ScMem                                                           'Remember z_ScMem so that we can restore its value on exit
    z_ScMem = zData(nTable)                                                   'Map zData() to the specified table

    If uMsg = ALL_MESSAGES Then                                               'If ALL_MESSAGES are being deleted from the table...
        zData(0) = 0                                                            'Zero the table entry count
    Else
        nCount = zData(0)                                                       'Get the table entry count

        For i = 1 To nCount                                                     'Loop through the table entries
            If zData(i) = uMsg Then                                               'If the message is found...
                zData(i) = 0                                                        'Null the msg value -- also frees the element for re-use
                GoTo Bail                                                           'Bail
            End If
        Next i                                                                  'Next message table entry

        zError "zDelMsg", "Message &H" & Hex$(uMsg) & " not found in table"
    End If

Bail:
    z_ScMem = nBase                                                           'Restore the value of z_ScMem

End Sub

'-SelfHook code------------------------------------------------------------------------------------
'-The following routines are exclusively for the shk_SetHook routines----------------------------

Private Function shk_SetHook(ByVal HookType As eHookType, _
                             Optional ByVal bGlobal As Boolean, _
                             Optional ByVal When As eMsgWhen = MSG_BEFORE, _
                             Optional ByVal lParamUser As Long = 0, _
                             Optional ByVal nOrdinal As Long = 1, _
                             Optional ByVal oCallback As Object = Nothing, _
                             Optional ByVal bIdeSafety As Boolean = True, _
                             Optional ByVal bUnicode As Boolean = False) As Boolean 'Setting specified hook

'*************************************************************************************************
'* HookType - One of the eHookType enumerators
'* bGlobal - If False, then hook applies to app's thread else it applies Globally (only supported by WH_KEYBOARD_LL & WH_MOUSE_LL)
'* When - either MSG_AFTER, MSG_BEFORE or MSG_BEFORE_AFTER
'* lParamUser - Optional, user-defined callback parameter
'* nOrdinal   - Optional, ordinal index of the callback procedure. 1 = last private method, 2 = second last private method, etc.
'* oCallback  - Optional, the object that will receive the callback. If undefined, callbacks are sent to this object's instance
'* bIdeSafety - Optional, enable/disable IDE safety measures. NB: you should really only disable IDE safety in a UserControl for design-time subclassing
'* bUnicode - Optional, if True, Unicode API calls will be made to the window vs ANSI calls
'*************************************************************************************************
' Hook procedure must be declared identical to the one near the end of this class (Sample at Ordinal #2)
' Subclassing procedure must be declared identical to the one at the end of this class (Sample at Ordinal #1)

' \\LaVolpe - The ASM for this procedure rewritten to mirror Paul Caton's SelfSub ASM
'       Therefore, it appears to be crash proof and allows a choice of whether you want
'       hook messages before and/or after the VB gets the message

Dim z_Sc() As Long                          'Thunk machine-code initialised here
Const MEM_LEN      As Long = 4 + 4 * 62     'Thunk length in bytes (last # must be = max zSc() array item)

Const PAGE_RWX      As Long = &H40&         'Allocate executable memory
Const MEM_COMMIT    As Long = &H1000&       'Commit allocated memory
Const MEM_RELEASE   As Long = &H8000&       'Release allocated memory flag
Const IDX_HOOKPROC  As Long = 9             'Thunk data index of the previous hook proc
Const IDX_EBMODE    As Long = 3             'Thunk data index of the EbMode function address
Const IDX_CNH       As Long = 4             'Thunk data index of the CallNextHook function address
Const IDX_UNW       As Long = 5             'Thunk data index of the UnhookWindowsEx function address
Const IDX_FREE      As Long = 6             'Thunk data index of the VirtualFree function address
Const IDX_BADPTR    As Long = 7             'Thunk data index of the IsBadCodePtr function address
Const IDX_OWNER     As Long = 8             'Thunk data index of the Owner object's vTable address
Const IDX_CALLBACK  As Long = 10            'Thunk data index of the callback method address
Const IDX_BTABLE    As Long = 11            'Thunk data index of the Before flag
Const IDX_ATABLE    As Long = 12            'Thunk data index of the After flag
Const IDX_EBX       As Long = 16            'Thunk code patch index of the thunk data
Const PROC_OFF      As Long = &H38          'Thunk offset to the HookProc execution address
Const SUB_NAME      As String = "shk_SetHook" 'This routine's name
Dim nAddr         As Long
Dim nID           As Long
Dim nMyID         As Long

    If oCallback Is Nothing Then Set oCallback = Me 'If the user hasn't specified the callback owner

    nAddr = zAddressOf(oCallback, nOrdinal)         'Get the address of the specified ordinal method
    If nAddr = 0 Then                               'Ensure that we've found the ordinal method
        zError SUB_NAME, "Callback method not found"
        Exit Function
    End If

    If Not bGlobal Then nID = App.ThreadID           ' thread ID to be used if not global hook

    z_ScMem = VirtualAlloc(0, MEM_LEN, MEM_COMMIT, PAGE_RWX)    'Allocate executable memory

    If z_ScMem <> 0 Then                                        'Ensure the allocation succeeded

        If z_hkFunk Is Nothing Then Set z_hkFunk = New Collection   'If this is the first time through, do the one-time initialization
        On Error GoTo CatchDoubleSub                                'Catch double subclassing
        z_hkFunk.Add z_ScMem, "h" & HookType                      'Add the hWnd/thunk-address to the collection
        On Error GoTo 0

        ReDim z_Sc(0 To (MEM_LEN - 4) \ 4)

        '\\LaVolep = complete rewritten hooking procedure
        z_Sc(14) = &HD231C031: z_Sc(15) = &HBBE58960: z_Sc(16) = &H12345678: z_Sc(17) = &HF63103FF: z_Sc(18) = &H750C4339: z_Sc(19) = &H73394A2A: z_Sc(20) = &HE80A742C: z_Sc(21) = &H7A&: z_Sc(22) = &H75147539: z_Sc(23) = &H2AE814: z_Sc(24) = &H73390000: z_Sc(25) = &H310A7430: z_Sc(26) = &H307B8BD2: z_Sc(27) = &H61E8&: z_Sc(28) = &H610BFF00: z_Sc(29) = &HFF0010C2: z_Sc(30) = &H13D0C53: z_Sc(31) = &H74000000: z_Sc(32) = &H74C085CC: z_Sc(33) = &H2E827: z_Sc(34) = &HE5EB0000: z_Sc(35) = &HFF2C75FF: z_Sc(36) = &H75FF2875: z_Sc(37) = &H2473FF24: z_Sc(38) = &H891053FF: z_Sc(39) = &H3B811C45: z_Sc(40) = &H1&: z_Sc(41) = &H73395575: z_Sc(42) = &H58507404: z_Sc(43) = &HFF2473FF: z_Sc(44) = &H5D891453: z_Sc(45) = &H2C758930:
        z_Sc(46) = &H2845C7: z_Sc(47) = &H8B000080: z_Sc(48) = &H5D892045: z_Sc(49) = &H24458914: z_Sc(50) = &H4C48161: z_Sc(51) = &HFF000000: z_Sc(52) = &HFF521862: z_Sc(53) = &H53FF2873: z_Sc(54) = &H1F755A1C: z_Sc(55) = &H8D34438D: z_Sc(56) = &H8D50144D: z_Sc(57) = &H73FF1C45: z_Sc(58) = &H2C75FF08: z_Sc(59) = &HFF2875FF: z_Sc(60) = &H51502475: z_Sc(61) = &H2073FF52: z_Sc(62) = &HC32853FF

        z_Sc(IDX_EBX) = z_ScMem                         'Patch the thunk data address
        z_Sc(IDX_INDEX) = HookType                       'Store the hook type in the thunk data
        z_Sc(IDX_OWNER) = ObjPtr(oCallback)             'Store the callback owner's object address in the thunk data
        z_Sc(IDX_CALLBACK) = nAddr                      'Store the callback address in the thunk data
        z_Sc(IDX_PARM_USER) = lParamUser                'Store the lParamUser callback parameter in the thunk data
        ' \\LaVolpe - validate unicode request & cache unicode usage
        If bUnicode Then bUnicode = (IsWindowUnicode(GetDesktopWindow) <> 0&)

        ' \\LaVolpe - added extra parameter "bUnicode" to the zFnAddr calls
        z_Sc(IDX_FREE) = zFnAddr("kernel32", "VirtualFree", bUnicode)       'Store the VirtualFree function address in the thunk data
        z_Sc(IDX_BADPTR) = zFnAddr("kernel32", "IsBadCodePtr", bUnicode)    'Store the IsBadCodePtr function address in the thunk data
        z_Sc(IDX_CNH) = zFnAddr("user32", "CallNextHookEx", bUnicode)       'Store CallWindowProc function address in the thunk data
        z_Sc(IDX_UNW) = zFnAddr("user32", "UnhookWindowsHookEx", bUnicode)  'Store the SetWindowLong function address in the thunk data

        Debug.Assert zInIDE
        If bIdeSafety = True And z_IDEflag = 1 Then                             'If the user wants IDE protection
            z_Sc(IDX_EBMODE) = zFnAddr("vba6", "EbMode", bUnicode)            'Store the EbMode function address in the thunk data
        End If

        If (When And MSG_BEFORE) = MSG_BEFORE Then z_Sc(IDX_BTABLE) = 1     ' non-zero flag if Before messages desired
        If (When And MSG_AFTER) = MSG_AFTER Then z_Sc(IDX_ATABLE) = 1       ' non-zero flag if After messages desired

        RtlMoveMemory z_ScMem, VarPtr(z_Sc(0)), MEM_LEN                     'Copy the thunk code/data to the allocated memory
        If bUnicode Then
            nAddr = SetWindowsHookExW(HookType, z_ScMem + PROC_OFF, App.hInstance, nID) 'Set the new WndProc, return the address of the original WndProc
        Else
            nAddr = SetWindowsHookExA(HookType, z_ScMem + PROC_OFF, App.hInstance, nID) 'Set the new WndProc, return the address of the original WndProc
        End If
        If nAddr = 0 Then                                                   'Ensure the new WndProc was set correctly
            zError SUB_NAME, "SetWindowLong failed, error #" & Err.LastDllError
            GoTo ReleaseMemory
        End If
        RtlMoveMemory z_ScMem + IDX_HOOKPROC * 4, VarPtr(nAddr), 4&          'Store the callback address

        shk_SetHook = True                                                  'Indicate success
    Else
        zError SUB_NAME, "VirtualAlloc failed, error: " & Err.LastDllError
    End If

Exit Function                                                             'Exit ssc_Subclass

CatchDoubleSub:
    zError SUB_NAME, "Window handle is already subclassed"

ReleaseMemory:
    VirtualFree z_ScMem, 0, MEM_RELEASE                                       'ssc_Subclass has failed after memory allocation, so release the memory

End Function

Private Function shk_UnHook(HookType As eHookType) As Boolean

' can be made public. Releases a specific hook
' can be removed and zUnThunk can be called directly

    zUnThunk HookType, HookThunk

End Function

Private Sub shk_TerminateHooks()

' can be made public. Releases all hooks
' can be removed and zTerminateThunks can be called directly

    zTerminateThunks HookThunk

End Sub

'-SelfCallback code------------------------------------------------------------------------------------
'-The following routines are exclusively for the scb_SetCallbackAddr routines----------------------------

Private Function scb_SetCallbackAddr(ByVal nParamCount As Long, _
                                     Optional ByVal nOrdinal As Long = 1, _
                                     Optional ByVal oCallback As Object = Nothing, _
                                     Optional ByVal bIdeSafety As Boolean = True) As Long   'Return the address of the specified callback thunk

'*************************************************************************************************
'* nParamCount  - The number of parameters that will callback
'* nOrdinal     - Callback ordinal number, the final private method is ordinal 1, the second last is ordinal 2, etc...
'* oCallback    - Optional, the object that will receive the callback. If undefined, callbacks are sent to this object's instance
'* bIdeSafety   - Optional, set to false to disable IDE protection.
'*************************************************************************************************
' Callback procedure must return a Long even if, per MSDN, the callback procedure is a Sub vs Function
' The number of parameters are dependent on the individual callback procedures

Const MEM_LEN     As Long = IDX_CALLBACKORDINAL * 4 + 4     'Memory bytes required for the callback thunk
Const PAGE_RWX    As Long = &H40&                           'Allocate executable memory
Const MEM_COMMIT  As Long = &H1000&                         'Commit allocated memory
Const SUB_NAME      As String = "scb_SetCallbackAddr"       'This routine's name
Const INDX_OWNER    As Long = 0
Const INDX_CALLBACK As Long = 1
Const INDX_EBMODE   As Long = 2
Const INDX_BADPTR   As Long = 3
Const INDX_EBX      As Long = 5
Const INDX_PARAMS   As Long = 12
Const INDX_PARAMLEN As Long = 17

Dim z_Cb()    As Long    'Callback thunk array
Dim nCallback As Long

    If z_cbFunk Is Nothing Then
        Set z_cbFunk = New Collection           'If this is the first time through, do the one-time initialization
    Else
        On Error Resume Next                    'Catch already initialized?
            z_ScMem = z_cbFunk.Item("h" & nOrdinal) 'Test it
            If Err = 0 Then
                scb_SetCallbackAddr = z_ScMem + 16  'we had this one, just reference it
                Exit Function
            End If
        On Error GoTo 0
    End If

    If nParamCount < 0 Then                     ' validate parameters
        zError SUB_NAME, "Invalid Parameter count"
        Exit Function
    End If

    If oCallback Is Nothing Then Set oCallback = Me     'If the user hasn't specified the callback owner
    nCallback = zAddressOf(oCallback, nOrdinal)         'Get the callback address of the specified ordinal
    If nCallback = 0 Then
        zError SUB_NAME, "Callback address not found."
        Exit Function
    End If
    z_ScMem = VirtualAlloc(0, MEM_LEN, MEM_COMMIT, PAGE_RWX) 'Allocate executable memory

    If z_ScMem = 0& Then
        zError SUB_NAME, "VirtualAlloc failed, error: " & Err.LastDllError  ' oops
        Exit Function
    End If
    z_cbFunk.Add z_ScMem, "h" & nOrdinal                  'Add the callback/thunk-address to the collection

    ReDim z_Cb(0 To IDX_CALLBACKORDINAL) As Long          'Allocate for the machine-code array

    ' Create machine-code array
    z_Cb(4) = &HBB60E089: z_Cb(6) = &H73FFC589: z_Cb(7) = &HC53FF04: z_Cb(8) = &H7B831F75: z_Cb(9) = &H20750008: z_Cb(10) = &HE883E889: z_Cb(11) = &HB9905004: z_Cb(13) = &H74FF06E3: z_Cb(14) = &HFAE2008D: z_Cb(15) = &H53FF33FF: z_Cb(16) = &HC2906104: z_Cb(18) = &H830853FF: z_Cb(19) = &HD87401F8: z_Cb(20) = &H4589C031: z_Cb(21) = &HEAEBFC

    z_Cb(INDX_BADPTR) = zFnAddr("kernel32", "IsBadCodePtr", False)
    z_Cb(INDX_OWNER) = ObjPtr(oCallback)                  'Set the Owner
    z_Cb(INDX_CALLBACK) = nCallback                       'Set the callback address
    z_Cb(IDX_CALLBACKORDINAL) = nOrdinal                  'Cache ordinal used for zTerminateThunks

    Debug.Assert zInIDE
    If bIdeSafety = True And z_IDEflag = 1 Then             'If the user wants IDE protection
        z_Cb(INDX_EBMODE) = zFnAddr("vba6", "EbMode", False)  'EbMode Address
    End If

    z_Cb(INDX_PARAMS) = nParamCount                         'Set the parameter count
    z_Cb(INDX_PARAMLEN) = nParamCount * 4                   'Set the number of stck bytes to release on thunk return

    '\\LaVolpe - redirect address to proper location in virtual memory. Was: z_Cb(INDX_EBX) = VarPtr(z_Cb(INDX_OWNER))
    z_Cb(INDX_EBX) = z_ScMem                                'Set the data address relative to virtual memory pointer

    RtlMoveMemory z_ScMem, VarPtr(z_Cb(INDX_OWNER)), MEM_LEN 'Copy thunk code to executable memory
    scb_SetCallbackAddr = z_ScMem + 16                       'Thunk code start address

End Function

Private Sub scb_ReleaseCallback(ByVal nOrdinal As Long)

' can be made public. Releases a specific callback
' can be removed and zUnThunk can be called directly

    zUnThunk nOrdinal, CallbackThunk

End Sub

Private Sub scb_TerminateCallbacks()

' can be made public. Releases all callbacks
' can be removed and zTerminateThunks can be called directly

    zTerminateThunks CallbackThunk

End Sub

'========================================================================
' COMMON USE ROUTINES
'-The following routines are used for each of the three types of thunks
'========================================================================

'Map zData() to the thunk address for the specified window handle

Private Function zMap_VFunction(ByVal vFuncTarget As Long, vType As eThunkType) As Long

'\\LaVolpe - Redone to be shared/used by Hook and Subclass routines

' vFuncTarget is one of the following, depending on vType
'   - Subclassing:  the hWnd of the window subclassed
'   - Hooking:      the hook type created
'   - Callbacks:    the ordinal of the callback

Dim thunkCol As Collection

    If vType = CallbackThunk Then
        Set thunkCol = z_cbFunk
    ElseIf vType = HookThunk Then
        Set thunkCol = z_hkFunk
    ElseIf vType = SubclassThunk Then
        Set thunkCol = z_scFunk
    Else
        zError "zMap_Vfunction", "Invalid thunk type passed"
        Exit Function
    End If

    If thunkCol Is Nothing Then
        zError "zMap_VFunction", "Thunk hasn't been initialized"
    Else
        On Error GoTo Catch
        z_ScMem = thunkCol("h" & vFuncTarget)                    'Get the thunk address
        zMap_VFunction = z_ScMem
    End If

Exit Function                                               'Exit returning the thunk address

Catch:
    zError "zMap_VFunction", "Thunk type for ID of " & vFuncTarget & " does not exist"

End Function

'Error handler

Private Sub zError(ByVal sRoutine As String, ByVal sMsg As String)

' \\LaVolpe -  Note. These two lines can be rem'd out if you so desire. But don't remove the routine

    App.LogEvent TypeName(Me) & "." & sRoutine & ": " & sMsg, vbLogEventTypeError
    MsgBox sMsg & ".", vbExclamation + vbApplicationModal, "Error in " & TypeName(Me) & "." & sRoutine

End Sub

'Return the address of the specified DLL/procedure

Private Function zFnAddr(ByVal sDLL As String, ByVal sProc As String, ByVal asUnicode As Boolean) As Long

' \\LaVolpe - Use ANSI calls for non-unicode usage, else use WideChar calls

    If asUnicode Then
        zFnAddr = GetProcAddress(GetModuleHandleW(StrPtr(sDLL)), sProc)         'Get the specified procedure address
    Else
        zFnAddr = GetProcAddress(GetModuleHandleA(sDLL), sProc)                 'Get the specified procedure address
    End If
    Debug.Assert zFnAddr                                                      'In the IDE, validate that the procedure address was located
    ' ^^ FYI VB5 users. Search for zFnAddr("vba6", "EbMode") and replace with zFnAddr("vba5", "EbMode")

End Function

'Return the address of the specified ordinal method on the oCallback object, 1 = last private method, 2 = second last private method, etc

Private Function zAddressOf(ByVal oCallback As Object, ByVal nOrdinal As Long) As Long

' Note: used both in subclassing and hooking routines

Dim bSub  As Byte                                                         'Value we expect to find pointed at by a vTable method entry
Dim bVal  As Byte
Dim nAddr As Long                                                         'Address of the vTable
Dim i     As Long                                                         'Loop index
Dim J     As Long                                                         'Loop limit

    RtlMoveMemory VarPtr(nAddr), ObjPtr(oCallback), 4                         'Get the address of the callback object's instance
    If Not zProbe(nAddr + &H1C, i, bSub) Then                                 'Probe for a Class method
        If Not zProbe(nAddr + &H6F8, i, bSub) Then                              'Probe for a Form method
            ' \\LaVolpe - Added propertypage offset
            If Not zProbe(nAddr + &H710, i, bSub) Then                            'Probe for a PropertyPage method
                If Not zProbe(nAddr + &H7A4, i, bSub) Then                          'Probe for a UserControl method
                    Exit Function                                                   'Bail...
                End If
            End If
        End If
    End If

    i = i + 4                                                                 'Bump to the next entry
    J = i + 1024                                                              'Set a reasonable limit, scan 256 vTable entries
    Do While i < J
        RtlMoveMemory VarPtr(nAddr), i, 4                                       'Get the address stored in this vTable entry

        If IsBadCodePtr(nAddr) Then                                             'Is the entry an invalid code address?
            RtlMoveMemory VarPtr(zAddressOf), i - (nOrdinal * 4), 4               'Return the specified vTable entry address
            Exit Do                                                               'Bad method signature, quit loop
        End If

        RtlMoveMemory VarPtr(bVal), nAddr, 1                                    'Get the byte pointed to by the vTable entry
        If bVal <> bSub Then                                                    'If the byte doesn't match the expected value...
            RtlMoveMemory VarPtr(zAddressOf), i - (nOrdinal * 4), 4               'Return the specified vTable entry address
            Exit Do                                                               'Bad method signature, quit loop
        End If

        i = i + 4                                                               'Next vTable entry
    Loop

End Function

'Probe at the specified start address for a method signature

Private Function zProbe(ByVal nStart As Long, ByRef nMethod As Long, ByRef bSub As Byte) As Boolean

Dim bVal    As Byte
Dim nAddr   As Long
Dim nLimit  As Long
Dim nEntry  As Long

    nAddr = nStart                                                            'Start address
    nLimit = nAddr + 32                                                       'Probe eight entries
    Do While nAddr < nLimit                                                   'While we've not reached our probe depth
        RtlMoveMemory VarPtr(nEntry), nAddr, 4                                  'Get the vTable entry

        If nEntry <> 0 Then                                                     'If not an implemented interface
            RtlMoveMemory VarPtr(bVal), nEntry, 1                                 'Get the value pointed at by the vTable entry
            If bVal = &H33 Or bVal = &HE9 Then                                    'Check for a native or pcode method signature
                nMethod = nAddr                                                     'Store the vTable entry
                bSub = bVal                                                         'Store the found method signature
                zProbe = True                                                       'Indicate success
                Exit Do                                                             'Return
            End If
        End If

        nAddr = nAddr + 4                                                       'Next vTable entry
    Loop

End Function

Private Function zInIDE() As Long

    z_IDEflag = 1
    zInIDE = z_IDEflag

End Function

Private Property Get zData(ByVal nIndex As Long) As Long

    RtlMoveMemory VarPtr(zData), z_ScMem + (nIndex * 4), 4

End Property

Private Property Let zData(ByVal nIndex As Long, ByVal nValue As Long)

    RtlMoveMemory z_ScMem + (nIndex * 4), VarPtr(nValue), 4

End Property

Private Sub zUnThunk(ByVal thunkID As Long, ByVal vType As eThunkType)

' Releases a specific subclass, hook or callback
' thunkID depends on vType:
'   - Subclassing:  the hWnd of the window subclassed
'   - Hooking:      the hook type created
'   - Callbacks:    the ordinal of the callback

Const IDX_SHUTDOWN  As Long = 1
Const MEM_RELEASE As Long = &H8000&                                'Release allocated memory flag

    If zMap_VFunction(thunkID, vType) Then
        Select Case vType
        Case SubclassThunk
            If IsBadCodePtr(z_ScMem) = 0 Then       'Ensure that the thunk hasn't already released its memory
                zData(IDX_SHUTDOWN) = 1             'Set the shutdown indicator
                zDelMsg ALL_MESSAGES, IDX_BTABLE    'Delete all before messages
                zDelMsg ALL_MESSAGES, IDX_ATABLE    'Delete all after messages
                '\\LaVolpe - Force thunks to replace original window procedure handle. Without this, app can crash when a window is subclassed multiple times simultaneously
                If zData(IDX_UNICODE) Then          'Force window procedure handle to be replaced
                    SendMessageW thunkID, 0&, 0&, ByVal 0&
                Else
                    SendMessageA thunkID, 0&, 0&, ByVal 0&
                End If
            End If
            z_scFunk.Remove "h" & thunkID           'Remove the specified thunk from the collection
        Case HookThunk
            If IsBadCodePtr(z_ScMem) = 0 Then       'Ensure that the thunk hasn't already released its memory
                zData(IDX_SHUTDOWN) = 1             'Set the shutdown indicator
                zData(IDX_ATABLE) = 0               ' want no more After messages
                zData(IDX_BTABLE) = 0               ' want no more Before messages
            End If
            z_hkFunk.Remove "h" & thunkID           'Remove the specified thunk from the collection
        Case CallbackThunk
            If IsBadCodePtr(z_ScMem) = 0 Then       'Ensure that the thunk hasn't already released its memory
                VirtualFree z_ScMem, 0, MEM_RELEASE 'Release allocated memory
            End If
            z_cbFunk.Remove "h" & thunkID           'Remove the specified thunk from the collection
        End Select
    End If

End Sub

Private Sub zTerminateThunks(ByVal vType As eThunkType)

' Removes all thunks of a specific type: subclassing, hooking or callbacks

Dim i As Long
Dim thunkCol As Collection

    Select Case vType
    Case SubclassThunk
        Set thunkCol = z_scFunk
    Case HookThunk
        Set thunkCol = z_hkFunk
    Case CallbackThunk
        Set thunkCol = z_cbFunk
    Case Else
        Exit Sub
    End Select

    If Not (thunkCol Is Nothing) Then                 'Ensure that hooking has been started
        With thunkCol
            For i = .count To 1 Step -1                   'Loop through the collection of hook types in reverse order
                z_ScMem = .Item(i)                          'Get the thunk address
                If IsBadCodePtr(z_ScMem) = 0 Then           'Ensure that the thunk hasn't already released its memory
                    Select Case vType
                    Case SubclassThunk
                        zUnThunk zData(IDX_INDEX), SubclassThunk     'Unsubclass
                    Case HookThunk
                        zUnThunk zData(IDX_INDEX), HookThunk             'Unhook
                    Case CallbackThunk
                        zUnThunk zData(IDX_CALLBACKORDINAL), CallbackThunk ' release callback
                    End Select
                End If
            Next i                                        'Next member of the collection
        End With
        Set thunkCol = Nothing                         'Destroy the hook/thunk-address collection
    End If

End Sub

' \\LaVolpe - EXAMPLE FUNCTIONS/CALLBACKS. Examples pulled from various postings

' ordinal #5 - Example of a window enumeration callback used with scb_SetCallbackAddr
' http://msdn2.microsoft.com/en-us/library/aa922950.aspx

Private Function myEnumWindowsProc(ByVal hWnd As Long, ByVal lParam As Long) As Long

'    Debug.Print hWnd
'    myEnumWindowsProc = 1 'Continue

End Function

' ordinal #4 - Example of font callback procedure used with scb_SetCallbackAddr
' http://msdn2.microsoft.com/en-us/library/Aa911409.aspx

' the LOGFONT and NEWTEXTMETRIC UDTs not declared in this template. Add this routine to a new class and unrem it for testing
'Private Function myEnumFontFamProc(lpNLF As LOGFONT, lpNTM As NEWTEXTMETRIC, ByVal FontType As Long, ByVal lParam As Long) As Long
'
'    ' Very cool; note that the call backs can receive UDTs too, just ensure they are ByRef and not ByVal
'    Dim FaceName As String
'
'    FaceName = StrConv(lpNLF.lfFaceName, vbUnicode)
'    FaceName = Left$(FaceName, InStr(FaceName, vbNullChar) - 1)
'    Debug.Print FaceName
'
'    myEnumFontFamProc = 1 'Continue
'
'End Function

' ordinal #3 - Example of a timer procedure callback used with scb_SetCallbackAddr
' http://msdn2.microsoft.com/en-us/library/ms644907.aspx

Private Function myTimerProc(ByVal hWnd As Long, ByVal tMsg As Long, ByVal TimerID As Long, ByVal tickCount As Long) As Long

' note: although a time procedure, per MSDN, does not return a value,
' the function that is used for callbacks must return a value therefore
' all callback routines must be functions, even if the return value is not used
' YOUR CODE HERE

End Function

' ordinal #2 ' Example of a hook procedure used with shk_SetHook

Private Sub myHookProc(ByVal bBefore As Boolean, _
                       ByRef bHandled As Boolean, _
                       ByRef lReturn As Long, _
                       ByVal nCode As Long, _
                       ByVal wParam As Long, _
                       ByVal lParam As Long, _
                       ByRef lParamUser As Long)

'*************************************************************************************************
' http://msdn2.microsoft.com/en-us/library/ms644990.aspx
'* bBefore    - Indicates whether the callback is before or after the next HookProc. Usually
'*              you will know unless the callback for the Msg is specified as
'*              MSG_BEFORE_AFTER (both before and after the next HookProc).
'* bHandled   - In a before next hook in chain callback, setting bHandled to True will prevent the
'*              message being passed to the  next hook in chain. Has no effect for After messages
'* lReturn    - Return value. Set as per the MSDN documentation for the hook type used
'* nCode      - A code the hook procedure uses to determine how to process the message
'* wParam     - Message related data, hook type specific
'* lParam     - Message related data, hook type specific
'* lParamUser - User-defined callback parameter. Change vartype as needed (i.e., Object, UDT, etc)
'*************************************************************************************************

' note. All hook type procedures are identically declared like this one;
'   however, the meaning of nCode, wParam, lParam are specific to the type of hook

' YOUR CODE HERE

End Sub

'- ordinal #1, example of a subclassing procedure used with ssc_Subclass

Private Sub myWndProc(ByVal bBefore As Boolean, _
                      ByRef bHandled As Boolean, _
                      ByRef lReturn As Long, _
                      ByVal lng_hWnd As Long, _
                      ByVal uMsg As Long, _
                      ByVal wParam As Long, _
                      ByVal lParam As Long, _
                      ByRef lParamUser As Long)

'*************************************************************************************************
'* bBefore    - Indicates whether the callback is before or after the original WndProc. Usually
'*              you will know unless the callback for the uMsg value is specified as
'*              MSG_BEFORE_AFTER (both before and after the original WndProc).
'* bHandled   - In a before original WndProc callback, setting bHandled to True will prevent the
'*              message being passed to the original WndProc. Has no effect for After messages
'* lReturn    - WndProc return value. Set as per the MSDN documentation for the message value,
'*              and/or, in an after the original WndProc callback, act on the return value as set
'*              by the original WndProc.
'* lng_hWnd   - Window handle.
'* uMsg       - Message value.
'* wParam     - Message related data.
'* lParam     - Message related data.
'* lParamUser - User-defined callback parameter. Change vartype as needed (i.e., Object, UDT, etc)
'*************************************************************************************************

Dim tNMH As NMHDR
Dim tNMHE As NMCBEENDEDIT
Dim tNMHEW As NMCBEENDEDITW
Dim iKeyCode As Integer
Dim tR As RECT
Dim bCancel As Boolean

    If uMsg = WM_SETFOCUS Then
        If Not m_bInFocus Then
            If IsWindowVisible(lng_hWnd) Then
                If (m_hWndCombo = lng_hWnd) Or (m_hWndEdit = lng_hWnd) Or (m_hWnd = lng_hWnd) Then
                    Call pvSetIPAO
                    m_bInFocus = True
                Else
                    Call SetFocus(m_hWnd)
                End If
            End If
        End If
    End If

    If uMsg = WM_MOUSEACTIVATE Then
        If Not m_bInFocus Then
            If GetFocus() <> m_hWndCombo And GetFocus() <> m_hWndEdit Then
                ' Click mouse down but miss the contained control; eat
                ' activate and setfocus to the the user control, this in
                ' turn focuses the contained Comboex
                Call SetFocus(UserControl.hWnd)
                bHandled = True
                lReturn = MA_NOACTIVATE
                Exit Sub
            End If
        End If
    End If

    If uMsg = WM_CTLCOLORLISTBOX Or WM_CTLCOLOREDIT Then
        ' This is the only way to get the handle of the
        ' list box portion of a combo box:
        If (uMsg = WM_CTLCOLORLISTBOX) Then
            If m_eStyle <> cboSimple Then
                If (m_hWndDropDown = 0) Then
                    m_hWndDropDown = lParam
                    If (IsWindow(m_hWndDropDown)) Then
                        GetWindowRect m_hWndDropDown, tR
                        bCancel = False
                        RaiseEvent RequestDropDownResize(tR.Left, tR.Top, tR.Right, tR.Bottom, bCancel)
                        If Not bCancel Then
                            MoveWindow m_hWndDropDown, tR.Left, tR.Top, tR.Right - tR.Left, tR.Bottom - tR.Top, 1
                        End If
                    End If
                    If m_hWndEdit <> 0 Then
                        Call SetFocus(m_hWndEdit)
                    End If
                End If
            End If
        End If
    End If

    Select Case lng_hWnd

    Case m_hWndParent
        Select Case uMsg
        Case WM_COMMAND
            bHandled = True
            lReturn = 0
            If m_hWnd Then
                If lParam = m_hWnd Then
                    Select Case ((wParam And &HFFFF0000) \ &H10000)
                    Case CBN_DROPDOWN
                        RaiseEvent DropDown
                    Case CBN_CLOSEUP
                        m_hWndDropDown = 0
                        RaiseEvent CloseUp
                    Case CBN_SELCHANGE
                        RaiseEvent EditChange
                        RaiseEvent ListIndexChange
                        RaiseEvent Click
                    Case CBN_EDITCHANGE
                        If m_hWndCombo Then
                            If (wParam And &HFFFF&) = GetWindowLong(m_hWndCombo, GWL_ID) Then
                                RaiseEvent EditChange
                            End If
                        Else
                            RaiseEvent EditChange
                        End If
                    End Select
                End If
            End If

        Case WM_NOTIFY
            bHandled = True
            CopyMemory tNMH, ByVal lParam, Len(tNMH)
            If m_hWnd Then
                If tNMH.hwndFrom = m_hWnd Then
                    Select Case tNMH.code
                    Case CBEN_BEGINEDIT
                        RaiseEvent BeginEdit
                    Case CBEN_ENDEDITA, CBEN_ENDEDITW
                        If m_bIsNt Then
                            CopyMemory tNMHEW, ByVal lParam, LenB(tNMHEW)
                            RaiseEvent EndEdit((tNMHEW.fChanged <> 0), tNMHEW.iNewSelection, pvStripNulls(tNMHEW.szText), tNMHEW.iWhy)
                        Else
                            CopyMemory tNMHE, ByVal lParam, LenB(tNMHE)
                            RaiseEvent EndEdit((tNMHE.fChanged <> 0), tNMHE.iNewSelection, pvStripNulls(StrConv(tNMHE.szText, vbUnicode)), tNMHE.iWhy)
                        End If
                    End Select
                End If
            End If

        Case WM_SIZE
            pvResize

        End Select

    Case m_hWndEdit
        Select Case uMsg
        Case WM_KEYUP
            iKeyCode = (wParam And &HFF)
            RaiseEvent KeyUp(iKeyCode, pvShiftState())
        Case WM_KEYDOWN
            iKeyCode = (wParam And &HFF)
            RaiseEvent KeyDown(iKeyCode, pvShiftState())

        End Select

    End Select

    ' *************************************************************
    ' C A U T I O N   C A U T I O N   C A U T I O N   C A U T I O N
    ' -------------------------------------------------------------
    ' DO NOT ADD ANY OTHER CODE BELOW THE "END SUB" STATEMENT BELOW
    '   add this warning banner to the last routine in your class
    ' *************************************************************

End Sub
