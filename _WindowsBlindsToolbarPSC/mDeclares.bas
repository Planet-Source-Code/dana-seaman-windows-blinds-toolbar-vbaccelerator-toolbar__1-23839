Attribute VB_Name = "mDeclares"
Option Explicit
' ======================================================================================
' Name:     mGDI
' Author:   Steve McMahon (steve@vbaccelerator.com)
' Date:     22 December 1998
'
' Copyright © 1998-1999 Steve McMahon for vbAccelerator
' --------------------------------------------------------------------------------------
' Visit vbAccelerator - advanced free source code for VB programmers
' http://vbaccelerator.com
' --------------------------------------------------------------------------------------
'
' Various GDI declares and helper functions for the vbAcceleratorGrid
' control.
'
' FREE SOURCE CODE - ENJOY!
' ======================================================================================
#Const DEBUGMODE = 0
Public Type RECT
   left As Long
   top As Long
   right As Long
   bottom As Long
End Type
Public Type POINTAPI
   x As Long
   y As Long
End Type
Public Type NMHDR
   hwndFrom As Long
   idfrom As Long
   code As Long
End Type

Public Type TOOLINFO
   cbSize As Long
   uFlags As Long
   hWnd As Long
   uID As Long
   rct As RECT
   hInst As Long
   lpszText As Long
End Type

Public Type ToolTipText
   hdr As NMHDR
   lpszText As Long
   szText As String * 80
   hInst As Long
   uFlags As Long
End Type

Public Const H_MAX As Long = &HFFFF + 1
Public Const WM_USER = &H400
Public Const TTM_RELAYEVENT = (WM_USER + 7)
'Tool Tip messages
Public Const TTM_ACTIVATE = (WM_USER + 1)
'#If UNICODE Then
'   Public Const TTM_ADDTOOLW = (WM_USER + 50)
'   Public Const TTM_ADDTOOL = TTM_ADDTOOLW
'   Public Const TTM_DELTOOLW = (WM_USER + 51)
'   Public Const TTM_DELTOOL = TTM_DELTOOLW
'#Else
   Public Const TTM_ADDTOOLA = (WM_USER + 4)
   Public Const TTM_ADDTOOL = TTM_ADDTOOLA
   Public Const TTM_DELTOOLA = (WM_USER + 5)
   Public Const TTM_DELTOOL = TTM_DELTOOLA
'#End If


Private Const MAX_PATH = 260
Private Type OPENFILENAME
    lStructSize As Long          ' Filled with UDT size
    hwndOwner As Long            ' Tied to Owner
    hInstance As Long            ' Ignored (used only by templates)
    lpstrFilter As String        ' Tied to Filter
    lpstrCustomFilter As String  ' Ignored (exercise for reader)
    nMaxCustFilter As Long       ' Ignored (exercise for reader)
    nFilterIndex As Long         ' Tied to FilterIndex
    lpstrFile As String          ' Tied to FileName
    nMaxFile As Long             ' Handled internally
    lpstrFileTitle As String     ' Tied to FileTitle
    nMaxFileTitle As Long        ' Handled internally
    lpstrInitialDir As String    ' Tied to InitDir
    lpstrTitle As String         ' Tied to DlgTitle
    flags As Long                ' Tied to Flags
    nFileOffset As Integer       ' Ignored (exercise for reader)
    nFileExtension As Integer    ' Ignored (exercise for reader)
    lpstrDefExt As String        ' Tied to DefaultExt
    lCustData As Long            ' Ignored (needed for hooks)
    lpfnHook As Long             ' Ignored (good luck with hooks)
    lpTemplateName As Long       ' Ignored (good luck with templates)
End Type

Private Declare Function GetOpenFileName Lib "COMDLG32" _
    Alias "GetOpenFileNameA" (file As OPENFILENAME) As Long
Private Declare Function CommDlgExtendedError Lib "COMDLG32" () As Long
Private Declare Function lstrlen Lib "kernel32" Alias "lstrlenA" (ByVal lpString As String) As Long

Private m_lApiReturn As Long
Private m_lExtendedError As Long
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
End Enum

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

Private Declare Function ChooseColor Lib "COMDLG32.DLL" _
    Alias "ChooseColorA" (Color As TCHOOSECOLOR) As Long

Public Enum EChooseColor
    CC_RGBInit = &H1
    CC_FullOpen = &H2
    CC_PreventFullOpen = &H4
    CC_ColorShowHelp = &H8
' Win95 only
    CC_SolidColor = &H80
    CC_AnyColor = &H100
' End Win95 only
    CC_ENABLEHOOK = &H10
    CC_ENABLETEMPLATE = &H20
    CC_EnableTemplateHandle = &H40
End Enum

' Array of custom colors lasts for life of app
Private alCustom(0 To 15) As Long, fNotFirst As Boolean


'ToolTip Notification
Public Const TTN_FIRST = (H_MAX - 520&)
'#If UNICODE Then
'   Public Const TTN_NEEDTEXTW = (TTN_FIRST - 10&)
'   Public Const TTN_NEEDTEXT = TTN_NEEDTEXTW
'#Else
   Public Const TTN_NEEDTEXTA = (TTN_FIRST - 0&)
   Public Const TTN_NEEDTEXT = TTN_NEEDTEXTA
'#End If
Private Const TOOLTIPS_CLASS = "tooltips_class32"
Private Const TTF_IDISHWND = &H1
Private Const LPSTR_TEXTCALLBACK As Long = -1

Public Declare Sub InitCommonControls Lib "COMCTL32.DLL" ()
Public Type CommonControlsEx
    dwSize As Long
    dwICC As Long
End Type
Public Declare Function InitCommonControlsEx Lib "COMCTL32.DLL" (iccex As CommonControlsEx) As Boolean
Public Const ICC_BAR_CLASSES = &H4
Public Const ICC_COOL_CLASSES = &H400
Public Const ICC_USEREX_CLASSES = &H200& '// comboex
Public Const ICC_WIN95_CLASSES = &HFF&

'//Common Control Constants
Public Const CCS_TOP = &H1&
'Public Const CCS_NOMOVEY = &H2&
Public Const CCS_BOTTOM = &H3&
Public Const CCS_NORESIZE = &H4&
Public Const CCS_NOPARENTALIGN = &H8&
Public Const CCS_ADJUSTABLE = &H20&
Public Const CCS_NODIVIDER = &H40&
Public Const CCS_VERT = &H80&
Public Const CCS_LEFT = (CCS_VERT Or CCS_TOP)
Public Const CCS_RIGHT = (CCS_VERT Or CCS_BOTTOM)
'Public Const CCS_NOMOVEX = (CCS_VERT Or CCS_NOMOVEY)

Public Const CCM_FIRST = &H2000&                  '// Common control shared messages
Public Const CCM_SETCOLORSCHEME = (CCM_FIRST + 2)     '// lParam is color scheme
Public Const CCM_GETCOLORSCHEME = (CCM_FIRST + 3)     '// fills in COLORSCHEME pointed to by lParam
Type COLORSCHEME
   dwSize As Long
   clrBtnHighlight As Long       '// highlight color
   clrBtnShadow As Long          '// shadow color
End Type

Private Const NM_FIRST = H_MAX               '(0U-  0U)       '// generic to all controls
'Private Const NM_LAST = H_MAX - 99              '(0U- 99U)

'//====== Generic WM_NOTIFY notification codes =================================

'Public Const NM_OUTOFMEMORY = (NM_FIRST - 1)
'Public Const NM_CLICK = (NM_FIRST - 2)                ' // uses NMCLICK struct
'Public Const NM_DBLCLK = (NM_FIRST - 3)
Public Const NM_RETURN = (NM_FIRST - 4)
'Public Const NM_RCLICK = (NM_FIRST - 5)               ' // uses NMCLICK struct
'Public Const NM_RDBLCLK = (NM_FIRST - 6)
'Public Const NM_SETFOCUS = (NM_FIRST - 7)
'Public Const NM_KILLFOCUS = (NM_FIRST - 8)
'#if (_WIN32_IE >= 0x0300)
'Public Const NM_CUSTOMDRAW = (NM_FIRST - 12)
'Public Const NM_HOVER = (NM_FIRST - 13)
'#End If
'#if (_WIN32_IE >= 0x0400)
Public Const NM_NCHITTEST = (NM_FIRST - 14)           ' // uses NMMOUSE struct
Public Const NM_KEYDOWN = (NM_FIRST - 15)             ' // uses NMKEY struct
Public Const NM_RELEASEDCAPTURE = (NM_FIRST - 16)
 'Public Const NM_SETCURSOR = (NM_FIRST - 17)           ' // uses NMMOUSE struct
'Public Const NM_CHAR = (NM_FIRST - 18)                ' // uses NMCHAR struct

'//====== Generic WM_NOTIFY notification structures ============================
Public Type NMMOUSE
   hdr As NMHDR
   dwItemSpec As Long
   dwItemData As Long
   pt As POINTAPI
   dwHitInfo As Long '// any specifics about where on the item or control the mouse is
End Type
' NMCLICK = NMMOUSE

'// Generic structure for a key
Type NMKEY
   hdr As NMHDR
   nVKey As Long
   uFlags As Long
End Type

'// Generic structure for a character
Type NMCHAR
   hdr As NMHDR
   ch As Long
   dwItemPrev As Long     '// Item previously selected
   dwItemNext As Long     '// Item to be selected
End Type

Public Const HINST_COMMCTRL = -1&

Private Const S_OK = &H0
Private Type DLLVERSIONINFO
    cbSize As Long
    dwMajor As Long
    dwMinor As Long
    dwBuildNumber As Long
    dwPlatformId As Long
End Type
Private Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
Private Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long
Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Private Declare Function DllGetVersion Lib "COMCTL32" (pdvi As DLLVERSIONINFO) As Long

Public Type TBBUTTON
   iBitmap As Long
   idCommand As Long
   fsState As Byte
   fsStyle As Byte
   bReserved1 As Byte
   bReserved2 As Byte
   dwData As Long
   iString As Long
End Type

' Toolbar and button styles:
Public Const TBSTYLE_BUTTON = &H0
Public Const TBSTYLE_SEP = &H1
Public Const TBSTYLE_CHECK = &H2
Public Const TBSTYLE_GROUP = &H4
Public Const TBSTYLE_CHECKGROUP = (TBSTYLE_GROUP Or TBSTYLE_CHECK)
Public Const TBSTYLE_DROPDOWN = &H8
Public Const TBSTYLE_TOOLTIPS = &H100
Public Const TBSTYLE_WRAPABLE = &H200
Public Const TBSTYLE_ALTDRAG = &H400
Public Const TBSTYLE_FLAT = &H800
Public Const TBSTYLE_LIST = &H1000
Public Const TBSTYLE_AUTOSIZE = &H10         '// automatically calculate the cx of the button
Public Const TBSTYLE_NOPREFIX = &H20         '// if this button should not have accel prefix
Public Const BTNS_WHOLEDROPDOWN = &H80 '??? IE5 only
Public Const TBSTYLE_REGISTERDROP = &H4000&
Public Const TBSTYLE_TRANSPARENT = &H8000&


'/* Toolbar messages needed elsewhere */
Public Const TB_PRESSBUTTON = (WM_USER + 3)
Public Const TB_BUTTONCOUNT = (WM_USER + 24)
Public Const TB_GETHOTITEM = (WM_USER + 71)
Public Const TB_SETHOTITEM = (WM_USER + 72)           '// wParam == iHotItem
Public Const TB_GETBUTTON = (WM_USER + 23)
Public Const TB_GETRECT = (WM_USER + 51)             '// wParam is the Cmd instead of index

Public Const TB_ISBUTTONENABLED = (WM_USER + 9)
Public Const TB_ISBUTTONCHECKED = (WM_USER + 10)
Public Const TB_ISBUTTONPRESSED = (WM_USER + 11)
Public Const TB_ISBUTTONHIDDEN = (WM_USER + 12)
Public Const TB_ISBUTTONINDETERMINATE = (WM_USER + 13)
Public Const TB_ISBUTTONHIGHLIGHTED = (WM_USER + 14)

Private Declare Function SetWindowsHookEx Lib "user32" Alias "SetWindowsHookExA" (ByVal idHook As Long, ByVal lpfn As Long, ByVal hmod As Long, ByVal dwThreadId As Long) As Long
Private Declare Function UnhookWindowsHookEx Lib "user32" (ByVal hHook As Long) As Long
Private Declare Function CallNextHookEx Lib "user32" (ByVal hHook As Long, ByVal nCode As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function GetCurrentThreadId Lib "kernel32" () As Long
Private Const WH_MSGFILTER As Long = (-1)
Private Const WH_KEYBOARD As Long = 2
Private Const MSGF_MENU = 2
Private Const HC_ACTION = 0

' =========================================================================


' Tooltips:
Private m_hWndToolTip As Long
Private m_iRef As Long
Public msToolTipBuffer As String         'Tool tip text; This string must have
                                         'module or global level scope, because
                                         'a pointer to it is copied into a
                                         'ToolTipText structure

' Next Control ID:
Private m_iID As Long

' Rebar Resizing information
Private Type tRebarInter
   hWndRebar As Long
   hWndParent As Long
End Type
Private m_tRebarInter() As tRebarInter
Private m_iRebarCount As Long
' Padding between rebars & edges
Private m_lPad As Long

' Message filter hook:
Private m_hMsgHook As Long
Private m_lMsgHookPtr As Long

' Keyboard hook (for accelerators):
Private m_hKeyHook As Long
Private m_lKeyHookPtr() As Long
Private m_iKeyHookCount As Long

'Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Type ICONINFO
   fIcon As Long
   xHotspot As Long
   yHotspot As Long
   hBmMask As Long
   hbmColor As Long
End Type
Private Declare Function GetIconInfo Lib "user32" (ByVal hIcon As Long, piconinfo As ICONINFO) As Long
Public Declare Function GetObjectAPI Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long
Public Type BITMAP '14 bytes
    bmType As Long
    bmWidth As Long
    bmHeight As Long
    bmWidthBytes As Long
    bmPlanes As Integer
    bmBitsPixel As Integer
    bmBits As Long
End Type
Private Declare Function CreateIconIndirect Lib "user32" (piconinfo As ICONINFO) As Long
Private Declare Function CreateBitmapIndirect Lib "gdi32" (lpBitmap As BITMAP) As Long
Public Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Public Declare Function GetDeviceCaps Lib "gdi32" (ByVal hdc As Long, ByVal nIndex As Long) As Long
'Private Const BITSPIXEL = 12         '  Number of bits per pixel

'Private Declare Function GetBitmapBits Lib "gdi32" (ByVal hBitmap As Long, ByVal dwCount As Long, lpBits As Any) As Long
'Private Declare Function SetBitmapBits Lib "gdi32" (ByVal hBitmap As Long, ByVal dwCount As Long, lpBits As Any) As Long
Type RGBQUAD
   rgbBlue As Byte
   rgbGreen As Byte
   rgbRed As Byte
   rgbReserved As Byte
End Type
Private Type BITMAPINFOHEADER '40 bytes
   biSize As Long
   biWidth As Long
   biHeight As Long
   biPlanes As Integer
   biBitCount As Integer
   biCompression As Long
   biSizeImage As Long
   biXPelsPerMeter As Long
   biYPelsPerMeter As Long
   biClrUsed As Long
   biClrImportant As Long
End Type
Private Type BITMAPINFO_1BPP
   bmiHeader As BITMAPINFOHEADER
   bmiColors(0 To 1) As RGBQUAD
End Type
Private Type BITMAPINFO_4BPP
   bmiHeader As BITMAPINFOHEADER
   bmiColors(0 To 15) As RGBQUAD
End Type
Private Type BITMAPINFO_8BPP
   bmiHeader As BITMAPINFOHEADER
   bmiColors(0 To 255) As RGBQUAD
End Type
Private Type BITMAPINFO_ABOVE8
   bmiHeader As BITMAPINFOHEADER
End Type

'Private Const DIB_PAL_COLORS = 1 '  color table in palette indices
'Private Const DIB_PAL_INDICES = 2 '  No color table indices into surf palette
'Private Const DIB_PAL_LOGINDICES = 4 '  No color table indices into DC palette
'Private Const DIB_PAL_PHYSINDICES = 2 '  No color table indices into surf palette
Private Const DIB_RGB_COLORS = 0 '  color table in RGBs
Private Declare Function GetDIBits Lib "gdi32" (ByVal aHDC As Long, ByVal hBitmap As Long, ByVal nStartScan As Long, ByVal nNumScans As Long, lpBits As Any, lpBI As Any, ByVal wUsage As Long) As Long
Private Declare Function SetDIBits Lib "gdi32" (ByVal hdc As Long, ByVal hBitmap As Long, ByVal nStartScan As Long, ByVal nNumScans As Long, lpBits As Any, lpBI As Any, ByVal wUsage As Long) As Long
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" ( _
    lpvDest As Any, lpvSource As Any, ByVal cbCopy As Long)
Public Declare Function CreateDCAsNull Lib "gdi32" Alias "CreateDCA" (ByVal lpDriverName As String, lpDeviceName As Any, lpOutput As Any, lpInitData As Any) As Long
Public Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Const BI_RGB = 0&
'Private Const BI_RLE4 = 2&
'Private Const BI_RLE8 = 1&




Public Type Msg
    hWnd As Long
    message As Long
    wParam As Long
    lParam As Long
    Time As Long
    pt As POINTAPI
End Type
Public Type WINDOWPOS
   hWnd As Long
   hWndInsertAfter As Long
   x As Long
   y As Long
   cX As Long
   cY As Long
   flags As Long
End Type
Public Type NCCALCSIZE_PARAMS
   rgrc(0 To 2) As RECT
   lppos As Long 'WINDOWPOS
End Type
Public Type DRAWITEMSTRUCT
   CtlType As Long
   CtlID As Long
   itemID As Long
   itemAction As Long
   itemState As Long
   hwndItem As Long
   hdc As Long
   rcItem As RECT
   ItemData As Long
End Type

Public Declare Function CreateWindowEX Lib "user32" Alias "CreateWindowExA" (ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, lpParam As Any) As Long
Public Declare Function DestroyWindow Lib "user32" (ByVal hWnd As Long) As Long
'Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Any) As Long
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Public Declare Function GetClassLong Lib "user32" Alias "GetClassLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Public Declare Function ShowWindow Lib "user32" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long
Public Declare Function MoveWindow Lib "user32" (ByVal hWnd As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
Public Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cX As Long, ByVal cY As Long, ByVal wFlags As Long) As Long
Public Declare Function GetActiveWindow Lib "user32" () As Long

Public Declare Function UpdateWindow Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function InvalidateRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT, ByVal bErase As Long) As Long
Public Declare Function InvalidateRectAsNull Lib "user32" Alias "InvalidateRect" (ByVal hWnd As Long, lpRect As Any, ByVal bErase As Long) As Long
Public Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Long) As Long

Public Declare Function GetParent Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long

'Public Declare Function ScreenToClient Lib "user32" (ByVal hwnd As Long, lpPoint As POINTAPI) As Long
'Public Declare Function ClientToScreen Lib "user32" (ByVal hwnd As Long, lpPoint As POINTAPI) As Long
Public Declare Function MapWindowPoints Lib "user32" (ByVal hwndFrom As Long, ByVal hwndTo As Long, lppt As Any, ByVal cPoints As Long) As Long

Public Declare Function IsWindowVisible Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function IsZoomed Lib "user32" (ByVal hWnd As Long) As Long
'Public Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
'Public Declare Function GetClientRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
'Public Declare Function InflateRect Lib "user32" (lpRect As RECT, ByVal x As Long, ByVal y As Long) As Long
'Public Declare Function OffsetRect Lib "user32" (lpRect As RECT, ByVal x As Long, ByVal y As Long) As Long
Public Declare Function UnionRect Lib "user32" (lpDestRect As RECT, lpSrc1Rect As RECT, lpSrc2Rect As RECT) As Long
'Public Declare Function FillRect Lib "user32" (ByVal hdc As Long, lpRect As RECT, ByVal hBrush As Long) As Long

Public Declare Function GetProp Lib "user32" Alias "GetPropA" (ByVal hWnd As Long, ByVal lpString As String) As Long
Public Declare Function SetProp Lib "user32" Alias "SetPropA" (ByVal hWnd As Long, ByVal lpString As String, ByVal hData As Long) As Long
Public Declare Function RemoveProp Lib "user32" Alias "RemovePropA" (ByVal hWnd As Long, ByVal lpString As String) As Long

Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

Public Declare Function ReleaseCapture Lib "user32" () As Long

Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
Public Const SM_CXBORDER = 5
Public Const SM_CXDLGFRAME = 7
Public Const SM_CXFIXEDFRAME = SM_CXDLGFRAME
Public Const SM_CXFRAME = 32
Public Const SM_CXHSCROLL = 21
'Public Const SM_CXVSCROLL = 2
'Public Const SM_CYCAPTION = 4
Public Const SM_CYDLGFRAME = 8
Public Const SM_CYFIXEDFRAME = SM_CYDLGFRAME
Public Const SM_CYFRAME = 33
'Public Const SM_CYHSCROLL = 3
Public Const SM_CYMENU = 15
Public Const SM_CYSMSIZE = 31
Public Const SM_CXSMSIZE = 30

Public Type TPMPARAMS
    cbSize As Long
    rcExclude As RECT
End Type

Public Declare Function GetMenu Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function SetMenu Lib "user32" (ByVal hWnd As Long, ByVal hMenu As Long) As Long
Public Declare Function GetSubMenu Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Public Declare Function TrackPopupMenu Lib "user32" (ByVal hMenu As Long, ByVal wFlags As Long, ByVal x As Long, ByVal y As Long, ByVal nReserved As Long, ByVal hWnd As Long, lprc As Any) As Long
Public Declare Function TrackPopupMenuEx Lib "user32" (ByVal hMenu As Long, ByVal un As Long, ByVal n1 As Long, ByVal n2 As Long, ByVal hWnd As Long, lpTPMParams As TPMPARAMS) As Long
Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Public Declare Function PtInRect Lib "user32" (lpRect As RECT, ByVal x As Long, ByVal y As Long) As Long
Public Declare Function GetMenuItemCount Lib "user32" (ByVal hMenu As Long) As Long
Public Declare Function GetSystemMenu Lib "user32" (ByVal hWnd As Long, ByVal bRevert As Long) As Long

Public Const TPM_CENTERALIGN = &H4&
Public Const TPM_LEFTALIGN = &H0&
Public Const TPM_LEFTBUTTON = &H0&
Public Const TPM_RIGHTALIGN = &H8&
Public Const TPM_RIGHTBUTTON = &H2&
Public Const TPM_TOPALIGN = &H0
Public Const TPM_VCENTERALIGN = &H10
Public Const TPM_BOTTOMALIGN = &H20
Public Const TPM_HORIZONTAL = &H0             '/* Horz alignment matters more */
Public Const TPM_VERTICAL = &H40              '/* Vert alignment matters more */
Public Const TPM_NONOTIFY = &H80              '/* Don't send any notification msgs */
Public Const TPM_RETURNCMD = &H100

'Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDest As Any, pSrc As Any, ByVal ByteLen As Long)

'Public Declare Function CreateDCAsNull Lib "gdi32" Alias "CreateDCA" (ByVal lpDriverName As String, lpDeviceName As Any, lpOutput As Any, lpInitData As Any) As Long
Public Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
'Public Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Public Declare Function GetWindowDC Lib "user32" (ByVal hWnd As Long) As Long
'Public Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Public Declare Function ReleaseDC Lib "user32" (ByVal hWnd As Long, ByVal hdc As Long) As Long
Public Declare Function CreateRectRgn Lib "gdi32" (ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long) As Long
Public Declare Function LoadImage Lib "user32" Alias "LoadImageA" (ByVal hInst As Long, ByVal lpsz As String, ByVal un1 As Long, ByVal n1 As Long, ByVal n2 As Long, ByVal un2 As Long) As Long
Public Declare Function LoadImageLong Lib "user32" Alias "LoadImageA" (ByVal hInst As Long, ByVal lpsz As Long, ByVal un1 As Long, ByVal n1 As Long, ByVal n2 As Long, ByVal un2 As Long) As Long
'Public Const CF_BITMAP = 2
'Public Const LR_LOADMAP3DCOLORS = &H1000&
'Public Const LR_LOADFROMFILE = &H10
'Public Const LR_LOADTRANSPARENT = &H20
'Public Const IMAGE_BITMAP = 0
'Public Declare Function DrawIconEx Lib "user32" (ByVal hdc As Long, ByVal xLeft As Long, ByVal yTop As Long, ByVal hIcon As Long, ByVal cxWidth As Long, ByVal cyWidth As Long, ByVal istepIfAniCur As Long, ByVal hbrFlickerFreeDraw As Long, ByVal diFlags As Long) As Long
'Public Const DI_NORMAL = &H3
Public Declare Function DrawFrameControl Lib "user32" (ByVal lHDC As Long, tR As RECT, ByVal eFlag As Long, ByVal eStyle As Long) As Long
Public Declare Function DrawEdge Lib "user32" (ByVal hdc As Long, qrc As RECT, ByVal edge As Long, ByVal grfFlags As Long) As Long
Public Declare Function DrawCaption Lib "user32" (ByVal hWnd As Long, ByVal hdc As Long, pcRect As RECT, ByVal un As Long) As Long
Public Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Public Declare Function GetSysColorBrush Lib "user32" (ByVal nIndex As Long) As Long
Public Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Public Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Public Declare Function OleTranslateColor Lib "oleaut32.dll" (ByVal lOleColor As Long, ByVal lHPalette As Long, lColorRef As Long) As Long
Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal XSrc As Long, ByVal YSrc As Long, ByVal dwRop As Long) As Long
' General Win declares:
Public Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long
' Sys colours:
'Public Const COLOR_WINDOWFRAME = 6
'Public Const COLOR_BTNFACE = 15
'Public Const COLOR_BTNTEXT = 18
Public Const COLOR_INACTIVECAPTION = 3
Public Const COLOR_ACTIVEBORDER = 10
Public Const COLOR_ACTIVECAPTION = 2
Public Const COLOR_INACTIVEBORDER = 11

Public Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer

Public Declare Function GetWindow Lib "user32" (ByVal hWnd As Long, ByVal wCmd As Long) As Long
Public Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, lpsz2 As Any) As Long
Public Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hWnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long

Public Const ICON_SMALL = 0
Public Const ICON_BIG = 1
Public Const WM_GETICON = &H7F&

Public Const ODT_BUTTON = 4

' Window relationship:
Public Const GW_CHILD = 5
Public Const GW_HWNDNEXT = 2

' Syscommand values:
Public Const SC_MOVE = &HF012&
Public Const SC_MINIMIZE = &HF020
Public Const SC_CLOSE = &HF060
Public Const SC_KEYMENU = &HF100

'Window Styles:
Public Const WS_CHILD = &H40000000
Public Const WS_VISIBLE = &H10000000
Public Const WS_CLIPCHILDREN = &H2000000
Public Const WS_CLIPSIBLINGS = &H4000000
Public Const WS_BORDER = &H800000
Public Const WS_EX_TOPMOST = &H8&
Public Const WS_EX_TOOLWINDOW = &H80&
Public Const CW_USEDEFAULT = &H80000000

' Class long values:
Public Const GCL_HICON = (-14)
Public Const GCL_HICONSM = (-34)

' Messages:
Public Const WM_MEASUREITEM = &H2C
Public Const WM_DESTROY = &H2
Public Const WM_SIZE = &H5
Public Const WM_ACTIVATE = &H6
Public Const WM_SETFOCUS = &H7
Public Const WM_KILLFOCUS = &H8
Public Const WM_PAINT = &HF
Public Const WM_ERASEBKGND = &H14
Public Const WM_ACTIVATEAPP = &H1C
Public Const WM_CANCELMODE = &H1F
Public Const WM_GETMINMAXINFO = &H24
Public Const WM_DRAWITEM = &H2B
Public Const WM_WINDOWPOSCHANGING = &H46
Public Const WM_NOTIFY = &H4E
Public Const WM_NCHITTEST = &H84
Public Const WM_NCCALCSIZE = &H83
Public Const WM_NCPAINT = &H85
Public Const WM_NCACTIVATE = &H86
'Public Const WM_KEYDOWN = &H100
'Public Const WM_KEYUP = &H101
Public Const WM_COMMAND = &H111
Public Const WM_SYSCOMMAND = &H112
Public Const WM_INITMENU = &H116
Public Const WM_INITMENUPOPUP = &H117
Public Const WM_MENUSELECT = &H11F
Public Const WM_MENUCHAR = &H120
Public Const WM_EXITSIZEMOVE = &H232
Public Const WM_MOUSEMOVE = &H200
Public Const WM_LBUTTONDOWN = &H201
Public Const WM_LBUTTONUP = &H202
Public Const WM_RBUTTONDOWN = &H204
Public Const WM_RBUTTONUP = &H205
Public Const WM_MBUTTONDOWN = &H207
Public Const WM_MBUTTONUP = &H208
Public Const WM_PARENTNOTIFY = &H210
Public Const WM_ENTERMENULOOP = &H211
Public Const WM_EXITMENULOOP = &H212
Public Const WM_MDIACTIVATE = &H222
Public Const WM_MDIRESTORE = &H223
Public Const WM_MDIMAXIMIZE = &H225
Public Const WM_MDIGETACTIVE = &H229
Public Const WM_MDISETMENU = &H230


' WM_NCHITTEST return values:
Public Const HTBORDER = 18
Public Const HTBOTTOM = 15
Public Const HTBOTTOMLEFT = 16
Public Const HTBOTTOMRIGHT = 17
Public Const HTCAPTION = 2
Public Const HTCLIENT = 1
Public Const HTERROR = (-2)
Public Const HTGROWBOX = 4
Public Const HTHSCROLL = 6
Public Const HTLEFT = 10
Public Const HTMAXBUTTON = 9
Public Const HTMENU = 5
Public Const HTMINBUTTON = 8
Public Const HTNOWHERE = 0
Public Const HTRIGHT = 11
Public Const HTSYSMENU = 3
Public Const HTTOP = 12
Public Const HTTOPLEFT = 13
Public Const HTTOPRIGHT = 14
Public Const HTTRANSPARENT = (-1)
Public Const HTVSCROLL = 7
Public Const HTREDUCE = HTMINBUTTON
Public Const HTSIZE = HTGROWBOX
Public Const HTSIZEFIRST = HTLEFT
Public Const HTSIZELAST = HTBOTTOMRIGHT
Public Const HTZOOM = HTMAXBUTTON

' WM_NCCALCSIZE return values;
Public Const WVR_ALIGNBOTTOM = &H40
Public Const WVR_ALIGNLEFT = &H20
Public Const WVR_ALIGNRIGHT = &H80
Public Const WVR_ALIGNTOP = &H10
Public Const WVR_HREDRAW = &H100
Public Const WVR_VALIDRECTS = &H400
Public Const WVR_VREDRAW = &H200
Public Const WVR_REDRAW = (WVR_HREDRAW Or WVR_VREDRAW)

' Window Long:
'Public Const GWL_STYLE = (-16)
'Public Const GWL_EXSTYLE = (-20)
Public Const GWL_USERDATA = (-21)
Public Const GWL_WNDPROC = -4
Public Const GWL_HWNDPARENT = (-8)

' WM_ACTIVATE wParam LoWords:
Public Const WA_INACTIVE = 0
Public Const WA_CLICKACTIVE = 2
Public Const WA_ACTIVE = 1

' Show window
Public Const SW_SHOW = 5
Public Const SW_HIDE = 0
Public Const SW_SHOWNORMAL = 1

' SetWIndowPos
Public Const HWND_TOPMOST = -1
'Public Const SWP_NOSIZE = &H1
'Public Const SWP_NOMOVE = &H2
'Public Const SWP_NOREDRAW = &H8
'Public Const SWP_SHOWWINDOW = &H40
'Public Const SWP_FRAMECHANGED = &H20        '  The frame changed: send WM_NCCALCSIZE
'Public Const SWP_DRAWFRAME = SWP_FRAMECHANGED
'Public Const SWP_HIDEWINDOW = &H80
'Public Const SWP_NOACTIVATE = &H10
'Public Const SWP_NOCOPYBITS = &H100
'Public Const SWP_NOOWNERZORDER = &H200      '  Don't do owner Z ordering
'Public Const SWP_NOREPOSITION = SWP_NOOWNERZORDER
'Public Const SWP_NOZORDER = &H4

' DrawFrameControl:
Public Const DFC_CAPTION = 1
'Public Const DFC_MENU = 2
'Public Const DFC_SCROLL = 3
'Public Const DFC_BUTTON = 4
'#if(WINVER >= =&H0500)
'Public Const DFC_POPUPMENU = 5
'#endif /* WINVER >= =&H0500 */

Public Const DFCS_CAPTIONCLOSE = &H0
Public Const DFCS_CAPTIONMIN = &H1
Public Const DFCS_CAPTIONMAX = &H2
Public Const DFCS_CAPTIONRESTORE = &H3
'Public Const DFCS_CAPTIONHELP = &H4

Public Const DFCS_INACTIVE = &H100
Public Const DFCS_PUSHED = &H200
'Public Const DFCS_CHECKED = &H400

' DrawEdge:
Public Const BDR_RAISEDOUTER = &H1
Public Const BDR_SUNKENOUTER = &H2
Public Const BDR_RAISEDINNER = &H4
Public Const BDR_SUNKENINNER = &H8

'Public Const BDR_OUTER = &H3
'Public Const BDR_INNER = &HC
'Public Const BDR_RAISED = &H5
'Public Const BDR_SUNKEN = &HA

Public Const EDGE_RAISED = (BDR_RAISEDOUTER Or BDR_RAISEDINNER)
Public Const EDGE_SUNKEN = (BDR_SUNKENOUTER Or BDR_SUNKENINNER)
Public Const EDGE_ETCHED = (BDR_SUNKENOUTER Or BDR_RAISEDINNER)
Public Const EDGE_BUMP = (BDR_RAISEDOUTER Or BDR_SUNKENINNER)

Public Const BF_LEFT = &H1
Public Const BF_TOP = &H2
Public Const BF_RIGHT = &H4
Public Const BF_BOTTOM = &H8

'Public Const BF_TOPLEFT = (BF_TOP Or BF_LEFT)
'Public Const BF_TOPRIGHT = (BF_TOP Or BF_RIGHT)
'Public Const BF_BOTTOMLEFT = (BF_BOTTOM Or BF_LEFT)
'Public Const BF_BOTTOMRIGHT = (BF_BOTTOM Or BF_RIGHT)
Public Const BF_RECT = (BF_LEFT Or BF_TOP Or BF_RIGHT Or BF_BOTTOM)

' Button control:
Public Const BM_GETCHECK = &HF0&
Public Const BM_GETSTATE = &HF2&
'Public Const BST_UNCHECKED = &H0&
Public Const BST_CHECKED = &H1&
'Public Const BST_INDETERMINATE = &H2&
Public Const BST_PUSHED = &H4&
'Public Const BST_FOCUS = &H8&

' flags for DrawCaption
Public Const DC_ACTIVE = &H1
'Public Const DC_SMALLCAP = &H2
Public Const DC_ICON = &H4
Public Const DC_TEXT = &H8
'Public Const DC_INBUTTON = &H10
'#if(WINVER >= 0x0500)
'Public Const DC_GRADIENT = &H20
'#endif /* WINVER >= 0x0500 */

'#Const DEBUGMSGBOX = 0
 
'Public Sub debugmsg(ByVal sMsg As String)
'#If DEBUGMSGBOX = 1 Then
'   MsgBox sMsg, vbInformation
'#Else
'   Debug.Print sMsg
'#End If
'End Sub





Public Type IMAGEINFO
    hBitmapImage As Long
    hBitmapMask As Long
    cPlanes As Long
    cBitsPerPixel As Long
    rcImage As RECT
End Type
    
' General:
Declare Function GetWindowWord Lib "user32" (ByVal hWnd As Long, ByVal nIndex As Long) As Integer
    Public Const GWW_HINSTANCE = (-6)
    
' GDI object functions:
'Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
'Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
'Declare Function GetObjectAPI Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long
Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
'Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
'Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hdc As Long) As Long
'Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
'Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
'Declare Function GetDeviceCaps Lib "gdi32" (ByVal hdc As Long, ByVal nIndex As Long) As Long
    Public Const BITSPIXEL = 12
    Public Const LOGPIXELSX = 88    '  Logical pixels/inch in X
    Public Const LOGPIXELSY = 90    '  Logical pixels/inch in Y
' System metrics:
'Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
    Public Const SM_CXICON = 11
    Public Const SM_CYICON = 12
   ' Public Const SM_CXFRAME = 32
   ' Public Const SM_CYCAPTION = 4
    'Public Const SM_CYFRAME = 33
    Public Const SM_CYBORDER = 6
   ' Public Const SM_CXBORDER = 5

' Region paint and fill functions:
Declare Function PaintRgn Lib "gdi32" (ByVal hdc As Long, ByVal hRgn As Long) As Long
Declare Function ExtFloodFill Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal crColor As Long, ByVal wFillType As Long) As Long
    Public Const FLOODFILLBORDER = 0
    Public Const FLOODFILLSURFACE = 1
Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long) As Long

' Pen functions:
Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
    Public Const PS_DASH = 1
    Public Const PS_DASHDOT = 3
    Public Const PS_DASHDOTDOT = 4
    Public Const PS_DOT = 2
    Public Const PS_SOLID = 0
    Public Const PS_NULL = 5

' Brush functions:
'Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Declare Function CreatePatternBrush Lib "gdi32" (ByVal hBitmap As Long) As Long

' Line functions:
Declare Function LineTo Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long) As Long
'Type POINTAPI
'        x As Long
'        y As Long
'End Type
Declare Function MoveToEx Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, lpPoint As POINTAPI) As Long

' Colour functions:
Declare Function SetTextColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
Declare Function SetBkColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
Declare Function SetBkMode Lib "gdi32" (ByVal hdc As Long, ByVal nBkMode As Long) As Long
    'Public Const OPAQUE = 2
   ' Public Const TRANSPARENT = 1
'Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long
   ' Public Const COLOR_ACTIVEBORDER = 10
   ' Public Const COLOR_ACTIVECAPTION = 2
   ' Public Const COLOR_ADJ_MAX = 100
   ' Public Const COLOR_ADJ_MIN = -100
   ' Public Const COLOR_APPWORKSPACE = 12
   ' Public Const COLOR_BACKGROUND = 1
    Public Const COLOR_BTNFACE = 15
   ' Public Const COLOR_BTNHIGHLIGHT = 20
   ' Public Const COLOR_BTNSHADOW = 16
    Public Const COLOR_BTNTEXT = 18
   ' Public Const COLOR_CAPTIONTEXT = 9
   ' Public Const COLOR_GRAYTEXT = 17
   ' Public Const COLOR_HIGHLIGHT = 13
   ' Public Const COLOR_HIGHLIGHTTEXT = 14
   ' Public Const COLOR_INACTIVEBORDER = 11
   ' Public Const COLOR_INACTIVECAPTION = 3
   ' Public Const COLOR_INACTIVECAPTIONTEXT = 19
   ' Public Const COLOR_MENU = 4
   ' Public Const COLOR_MENUTEXT = 7
   ' Public Const COLOR_SCROLLBAR = 0
    Public Const COLOR_WINDOW = 5
   ' Public Const COLOR_WINDOWFRAME = 6
   ' Public Const COLOR_WINDOWTEXT = 8
   ' Public Const COLORONCOLOR = 3

' Shell Extract icon functions:
Declare Function FindExecutable Lib "shell32.dll" Alias "FindExecutableA" (ByVal lpFile As String, ByVal lpDirectory As String, ByVal lpResult As String) As Long
Declare Function ExtractIcon Lib "shell32.dll" Alias "ExtractIconA" (ByVal hInst As Long, ByVal lpszExeFileName As String, ByVal nIconIndex As Long) As Long

' Icon functions:
Declare Function DrawIcon Lib "user32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal hIcon As Long) As Long
Declare Function DestroyIcon Lib "user32" (ByVal hIcon As Long) As Long
'Declare Function DrawIconEx Lib "user32" (ByVal hdc As Long, ByVal xLeft As Long, ByVal yTop As Long, ByVal hIcon As Long, ByVal cxWidth As Long, ByVal cyWidth As Long, ByVal istepIfAniCur As Long, ByVal hbrFlickerFreeDraw As Long, ByVal diFlags As Long) As Boolean
'Public Const DI_MASK = &H1&
'Public Const DI_IMAGE = &H2&
Public Const DI_NORMAL = &H3&
'Public Const DI_COMPAT = &H4&
'Public Const DI_DEFAULTSIZE = &H8&

'Declare Function LoadImage Lib "user32" Alias "LoadImageA" (ByVal hInst As Long, ByVal lpsz As String, ByVal un1 As Long, ByVal n1 As Long, ByVal n2 As Long, ByVal un2 As Long) As Long
'Declare Function LoadImageLong Lib "user32" Alias "LoadImageA" (ByVal hInst As Long, ByVal lpsz As Long, ByVal un1 As Long, ByVal n1 As Long, ByVal n2 As Long, ByVal un2 As Long) As Long
    'Public Const LR_LOADMAP3DCOLORS = &H1000
    'Public Const LR_LOADFROMFILE = &H10
    'Public Const LR_LOADTRANSPARENT = &H20
Declare Function DestroyCursor Lib "user32" (ByVal hCursor As Long) As Long

' Blitting functions
'Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
    Public Const SRCAND = &H8800C6
    Public Const SRCCOPY = &HCC0020
    Public Const SRCERASE = &H440328
    Public Const SRCINVERT = &H660046
    Public Const SRCPAINT = &HEE0086
   ' Public Const BLACKNESS = &H42
    Public Const WHITENESS = &HFF0062
Declare Function PatBlt Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal dwRop As Long) As Long
Declare Function LoadBitmapBynum Lib "user32" Alias "LoadBitmapA" (ByVal hInstance As Long, ByVal lpBitmapName As Long) As Long
'Type BITMAP '14 bytes
'    bmType As Long
'    bmWidth As Long
'    bmHeight As Long
'    bmWidthBytes As Long
'    bmPlanes As Integer
'    bmBitsPixel As Integer
'    bmBits As Long
'End Type

'Declare Function CreateBitmapIndirect Lib "gdi32" (lpBitmap As BITMAP) As Long

' Text functions:
'Type RECT
'    Left As Long
'    Top As Long
'    Right As Long
'    Bottom As Long
'End Type
'Declare Function PtInRect Lib "user32" (lpRect As RECT, ByVal ptx As Long, ByVal pty As Long) As Long
'Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hdc As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
   ' Public Const DT_BOTTOM = &H8&
    'Public Const DT_CENTER = &H1&
    'Public Const DT_LEFT = &H0&
    'Public Const DT_CALCRECT = &H400&
    Public Const DT_WORDBREAK = &H10&
   ' Public Const DT_VCENTER = &H4&
    Public Const DT_TOP = &H0&
    Public Const DT_TABSTOP = &H80&
   ' Public Const DT_SINGLELINE = &H20&
    Public Const DT_RIGHT = &H2&
    Public Const DT_NOCLIP = &H100&
    Public Const DT_INTERNAL = &H1000&
    Public Const DT_EXTERNALLEADING = &H200&
    Public Const DT_EXPANDTABS = &H40&
    Public Const DT_CHARSTREAM = 4&
    Public Const DT_NOPREFIX = &H800&
Type DRAWTEXTPARAMS
    cbSize As Long
    iTabLength As Long
    iLeftMargin As Long
    iRightMargin As Long
    uiLengthDrawn As Long
End Type
Declare Function DrawTextEx Lib "user32" Alias "DrawTextExA" (ByVal hdc As Long, ByVal lpsz As String, ByVal n As Long, lpRect As RECT, ByVal un As Long, lpDrawTextParams As DRAWTEXTPARAMS) As Long
Declare Function DrawTextExAsNull Lib "user32" Alias "DrawTextExA" (ByVal hdc As Long, ByVal lpsz As String, ByVal n As Long, lpRect As RECT, ByVal un As Long, ByVal lpDrawTextParams As Long) As Long
  '  Public Const DT_EDITCONTROL = &H2000&
  '  Public Const DT_PATH_ELLIPSIS = &H4000&
  '  Public Const DT_END_ELLIPSIS = &H8000&
  '  Public Const DT_MODIFYSTRING = &H10000
  '  Public Const DT_RTLREADING = &H20000
  '  Public Const DT_WORD_ELLIPSIS = &H40000

Type SIZEAPI
    cX As Long
    cY As Long
End Type
'Public Declare Function GetTextExtentPoint32 Lib "gdi32" Alias "GetTextExtentPoint32A" (ByVal hdc As Long, ByVal lpsz As String, ByVal cbString As Long, lpSize As SIZEAPI) As Long
'Public Declare Function GetStockObject Lib "gdi32" (ByVal nIndex As Long) As Long
   ' Public Const ANSI_FIXED_FONT = 11
   ' Public Const ANSI_VAR_FONT = 12
    Public Const SYSTEM_FONT = 13
   ' Public Const DEFAULT_GUI_FONT = 17 'win95 only
'Declare Function FillRect Lib "user32" (ByVal hdc As Long, lpRect As RECT, ByVal hBrush As Long) As Long
'Declare Function Rectangle Lib "gdi32" (ByVal hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
'Declare Function DrawEdge Lib "user32" (ByVal hdc As Long, qrc As RECT, ByVal edge As Long, ByVal grfFlags As Long) As Long
'    Public Const BF_LEFT = 1
'    Public Const BF_TOP = 2
'    Public Const BF_RIGHT = 4
'    Public Const BF_BOTTOM = 8
   ' Public Const BF_RECT = BF_LEFT Or BF_TOP Or BF_RIGHT Or BF_BOTTOM
'     Public Const BF_MIDDLE = 2048
   ' Public Const BDR_SUNKENINNER = 8
   ' Public Const BDR_SUNKENOUTER = 2
   ' Public Const BDR_RAISEDOUTER = 1
   ' Public Const BDR_RAISEDINNER = 4

'Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
    Public Const SW_SHOWNOACTIVATE = 4

' Scrolling and region functions:
'Declare Function ScrollDC Lib "user32" (ByVal hdc As Long, ByVal dx As Long, ByVal dy As Long, lprcScroll As RECT, lprcClip As RECT, ByVal hrgnUpdate As Long, lprcUpdate As RECT) As Long
'Declare Function SetWindowRgn Lib "user32" (ByVal hwnd As Long, ByVal hRgn As Long, ByVal bRedraw As Long) As Long
Declare Function SelectClipRgn Lib "gdi32" (ByVal hdc As Long, ByVal hRgn As Long) As Long
'Declare Function CreateRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
'Declare Function CreateRectRgnIndirect Lib "gdi32" (lpRect As RECT) As Long
'Declare Function CreatePolyPolygonRgn Lib "gdi32" (lpPoint As POINTAPI, lpPolyCounts As Long, ByVal nCount As Long, ByVal nPolyFillMode As Long) As Long
'Declare Function CreatePolygonRgn Lib "gdi32" (lpPoint As POINTAPI, ByVal nCount As Long, ByVal nPolyFillMode As Long)
'Declare Function SaveDC Lib "gdi32" (ByVal hdc As Long) As Long
'Declare Function RestoreDC Lib "gdi32" (ByVal hdc As Long, ByVal hSavedDC As Long) As Long
'Declare Function CreateDCAsNull Lib "gdi32" Alias "CreateDCA" (ByVal lpDriverName As String, lpDeviceName As Any, lpOutput As Any, lpInitData As Any) As Long

Public Const LF_FACESIZE = 32
'Type LOGFONT
'    lfHeight As Long
'    lfWidth As Long
'    lfEscapement As Long
'    lfOrientation As Long
'    lfWeight As Long
'    lfItalic As Byte
'    lfUnderline As Byte
'    lfStrikeOut As Byte
'    lfCharSet As Byte
'    lfOutPrecision As Byte
'    lfClipPrecision As Byte
'    lfQuality As Byte
'    lfPitchAndFamily As Byte
'    lfFaceName(LF_FACESIZE) As Byte
'End Type
Public Const FW_NORMAL = 400
Public Const FW_BOLD = 700
'Public Const FF_DONTCARE = 0
'Public Const DEFAULT_QUALITY = 0
'Public Const DEFAULT_PITCH = 0
'Public Const DEFAULT_CHARSET = 1
'Declare Function CreateFontIndirect& Lib "gdi32" Alias "CreateFontIndirectA" (lpLogFont As LOGFONT)
Declare Function MulDiv Lib "kernel32" (ByVal nNumber As Long, ByVal nNumerator As Long, ByVal nDenominator As Long) As Long
'Declare Function DrawFocusRect Lib "user32" (ByVal hdc As Long, lpRect As RECT) As Long

Declare Function DrawState Lib "user32" Alias "DrawStateA" _
    (ByVal hdc As Long, _
    ByVal hBrush As Long, _
    ByVal lpDrawStateProc As Long, _
    ByVal lParam As Long, _
    ByVal wParam As Long, _
    ByVal x As Long, _
    ByVal y As Long, _
    ByVal cX As Long, _
    ByVal cY As Long, _
    ByVal fuFlags As Long) As Long

'/* Image type */
'Public Const DST_COMPLEX = &H0&
Public Const DST_TEXT = &H1&
'Public Const DST_PREFIXTEXT = &H2&
Public Const DST_ICON = &H3&
'Public Const DST_BITMAP = &H4&

' /* State type */
'Public Const DSS_NORMAL = &H0&
'Public Const DSS_UNION = &H10& ' Dither
Public Const DSS_DISABLED = &H20&
'Public Const DSS_MONO = &H80& ' Draw in colour of brush specified in hBrush
'Public Const DSS_RIGHT = &H8000&

'Public Declare Function OleTranslateColor Lib "OLEPRO32.DLL" (ByVal OLE_COLOR As Long, ByVal HPALETTE As Long, pccolorref As Long) As Long
Public Const CLR_INVALID = -1

' Image list functions:
'Public Declare Function ImageList_GetBkColor Lib "comctl32" (ByVal hImageList As Long) As Long
'Public Declare Function ImageList_ReplaceIcon Lib "comctl32" (ByVal hImageList As Long, ByVal i As Long, ByVal hIcon As Long) As Long
'Public Declare Function ImageList_Convert Lib "comctl32" Alias "ImageList_Draw" (ByVal hImageList As Long, ByVal ImgIndex As Long, ByVal hDCDest As Long, ByVal x As Long, ByVal y As Long, ByVal flags As Long) As Long
Public Declare Function ImageList_Create Lib "COMCTL32" (ByVal MinCx As Long, ByVal MinCy As Long, ByVal flags As Long, ByVal cInitial As Long, ByVal cGrow As Long) As Long
Public Declare Function ImageList_AddMasked Lib "COMCTL32" (ByVal hImageList As Long, ByVal hbmImage As Long, ByVal crMask As Long) As Long
'Public Declare Function ImageList_Replace Lib "comctl32" (ByVal hImageList As Long, ByVal ImgIndex As Long, ByVal hbmImage As Long, ByVal hBmMask As Long) As Long
'Public Declare Function ImageList_Add Lib "comctl32" (ByVal hImageList As Long, ByVal hbmImage As Long, hBmMask As Long) As Long
Public Declare Function ImageList_Remove Lib "COMCTL32" (ByVal hImageList As Long, ByVal ImgIndex As Long) As Long

'Public Declare Function ImageList_GetImageInfo Lib "Comctl32.dll" ( _
'        ByVal hIml As Long, _
'        ByVal i As Long, _
'        pImageInfo As IMAGEINFO _
'    ) As Long
Public Declare Function ImageList_AddIcon Lib "COMCTL32" (ByVal himl As Long, ByVal hIcon As Long) As Long
Public Declare Function ImageList_GetIcon Lib "COMCTL32" (ByVal hImageList As Long, ByVal ImgIndex As Long, ByVal fuFlags As Long) As Long
'Public Declare Function ImageList_SetImageCount Lib "comctl32" (ByVal hImageList As Long, uNewCount As Long)
Public Declare Function ImageList_GetImageCount Lib "COMCTL32" (ByVal hImageList As Long) As Long
Public Declare Function ImageList_Destroy Lib "COMCTL32" (ByVal hImageList As Long) As Long
Public Declare Function ImageList_GetIconSize Lib "COMCTL32" (ByVal hImageList As Long, cX As Long, cY As Long) As Long
'Public Declare Function ImageList_SetIconSize Lib "comctl32" (ByVal hImageList As Long, cx As Long, cy As Long) As Long

' ImageList functions:
' Draw:
Public Declare Function ImageList_Draw Lib "COMCTL32.DLL" ( _
        ByVal himl As Long, _
        ByVal i As Long, _
        ByVal hdcDst As Long, _
        ByVal x As Long, _
        ByVal y As Long, _
        ByVal fStyle As Long _
    ) As Long
'Public Const ILD_NORMAL = 0&
Public Const ILD_TRANSPARENT = 1&
'Public Const ILD_BLEND25 = 2&
Public Const ILD_SELECTED = 4&
'Public Const ILD_FOCUS = 4&
'Public Const ILD_MASK = &H10&
'Public Const ILD_IMAGE = &H20&
'Public Const ILD_ROP = &H40&
'Public Const ILD_OVERLAYMASK = 3840&
'Public Declare Function ImageList_GetImageRect Lib "Comctl32.dll" ( _
'        ByVal hIml As Long, _
'        ByVal i As Long, _
'        prcImage As RECT _
'    ) As Long
' Messages:
Public Declare Function ImageList_DrawEx Lib "COMCTL32" (ByVal himl As Long, ByVal i As Long, ByVal hdcDst As Long, ByVal x As Long, ByVal y As Long, ByVal dx As Long, ByVal dy As Long, ByVal rgbBk As Long, ByVal rgbFg As Long, ByVal fStyle As Long) As Long
'Public Declare Function ImageList_LoadImage Lib "comctl32" Alias "ImageList_LoadImageA" (ByVal hInst As Long, ByVal lpbmp As String, ByVal cx As Long, ByVal cGrow As Long, ByVal crMask As Long, ByVal uType As Long, ByVal uFlags As Long)
'Public Declare Function ImageList_SetBkColor Lib "comctl32" (ByVal hImageList As Long, ByVal clrBk As Long) As Long

Public Const ILC_MASK = &H1&
 
'Public Const CLR_DEFAULT = -16777216
'Public Const CLR_HILIGHT = -16777216
Public Const CLR_NONE = -1

Public Const ILCF_MOVE = &H0&
Public Const ILCF_SWAP = &H1&
Public Declare Function ImageList_Copy Lib "COMCTL32" (ByVal himlDst As Long, ByVal iDst As Long, ByVal himlSrc As Long, ByVal iSrc As Long, ByVal uFlags As Long) As Long

Private Declare Function GetTempFileName Lib "kernel32" Alias "GetTempFileNameA" (ByVal lpszPath As String, ByVal lpPrefixString As String, ByVal wUnique As Long, ByVal lpTempFileName As String) As Long
Private Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
'Private Const MAX_PATH = 260
'Private Declare Function lstrlen Lib "kernel32" Alias "lstrlenA" (ByVal lpString As String) As Long

Private Type PictDesc
    cbSizeofStruct As Long
    picType As Long
    hImage As Long
    xExt As Long
    yExt As Long
End Type
Private Type Guid
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(0 To 7) As Byte
End Type
Private Declare Function OleCreatePictureIndirect Lib "OLEPRO32.DLL" (lpPictDesc As PictDesc, riid As Guid, ByVal fPictureOwnsHandle As Long, ipic As IPicture) As Long



' the grid
Public Type tGridCell
   oBackColor As OLE_COLOR
   oForeColor As OLE_COLOR
   iFntIndex As Long
   sText As Variant
   eTextFlags As Long 'ECGTextAlignFlags
   iIconIndex As Long
   bSelected As Boolean
   bDirtyFlag As Boolean
   lIndent As Long
   lExtraIconIndex As Long
   lItemData As Long
   ' 19/10/1999: More options
   bOwnerDraw As Boolean
   lCellBorderStyle As Long
End Type
Public Type tRowPosition
   lHeight As Long
   lStartY As Long
   bVisible As Boolean
   bFixed As Boolean
   sKey As String
   bGroupRow As Boolean
   lGroupStartColIndex As Long
End Type


Public Type LOGBRUSH
   lbStyle As Long
   lbColor As Long
   lbHatch As Long
End Type


Public Type SYSTEMTIME
  wYear             As Integer
  wMonth            As Integer
  wDayOfWeek        As Integer
  wDay              As Integer
  wHour             As Integer
  wMinute           As Integer
  wSecond           As Integer
  wMilliseconds     As Long
End Type
Declare Function GetSystemDefaultLCID Lib "kernel32" () As Long
Declare Function GetLocaleInfo Lib "kernel32" Alias "GetLocaleInfoA" (ByVal Locale As Long, ByVal LCType As Long, ByVal lpLCData As String, ByVal cchData As Long) As Long
Declare Function GetTimeFormat Lib "kernel32" Alias "GetTimeFormatA" (ByVal Locale As Long, ByVal dwFlags As Long, lpTime As SYSTEMTIME, ByVal lpFormat As String, ByVal lpTimeStr As String, ByVal cchTime As Long) As Long
Declare Function GetDateFormat Lib "kernel32" Alias "GetDateFormatA" (ByVal Locale As Long, ByVal dwFlags As Long, lpDate As SYSTEMTIME, ByVal lpFormat As String, ByVal lpDateStr As String, ByVal cchDate As Long) As Long
Declare Function FileTimeToSystemTime Lib "kernel32" (lpFileTime As Currency, lpSystemTime As SYSTEMTIME) As Long
Declare Function FileTimeToLocalFileTime Lib "kernel32" (lpFileTime As Currency, lpLocalFileTime As Currency) As Long
'Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Long) As Long
'Declare Function UpdateWindow Lib "user32" (ByVal hwnd As Long) As Long

'Private Declare Function CreateDCAsNull Lib "gdi32" Alias "CreateDCA" ( _
'   ByVal lpDriverName As String, lpDeviceName As Any, lpOutput As Any, lpInitData As Any) As Long
Private Declare Function SetROP2 Lib "gdi32" (ByVal hdc As Long, ByVal nDrawMode As Long) As Long
  '   Private Const R2_BLACK = 1 ' 0
  '   Private Const R2_COPYPEN = 13 ' P
  '   Private Const R2_LAST = 16
  '   Private Const R2_MASKNOTPEN = 3 ' DPna
  '   Private Const R2_MASKPEN = 9 ' DPa
  '   Private Const R2_MASKPENNOT = 5 ' PDna
  '   Private Const R2_MERGENOTPEN = 12    ' DPno
  '   Private Const R2_MERGEPEN = 15 ' DPo
  '   Private Const R2_MERGEPENNOT = 14    ' PDno
  '   Private Const R2_NOP = 11    ' D
  '   Private Const R2_NOT = 6 ' Dn
  '   Private Const R2_NOTCOPYPEN = 4 ' PN
  '   Private Const R2_NOTMASKPEN = 8 ' DPan
  '   Private Const R2_NOTMERGEPEN = 2 ' DPon
     Private Const R2_NOTXORPEN = 10 ' DPxn
  '   Private Const R2_WHITE = 16 ' 1
  '   Private Const R2_XORPEN = 7 ' DPx
'Public Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Public Declare Function ScrollDC Lib "user32" (ByVal hdc As Long, ByVal dx As Long, ByVal dy As Long, lprcScroll As RECT, lprcClip As RECT, ByVal hrgnUpdate As Long, lprcUpdate As RECT) As Long
Public Declare Function Rectangle Lib "gdi32" (ByVal hdc As Long, ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long) As Long
Public Declare Function FrameRect Lib "user32" (ByVal hdc As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Public Declare Function DrawFocusRect Lib "user32" (ByVal hdc As Long, lpRect As RECT) As Long
'Public Declare Function ExtCreatePen Lib "gdi32" (ByVal dwPenStyle As Long, ByVal dwWidth As Long, lplb As LOGBRUSH, ByVal dwStyleCount As Long, lpStyle As Long) As Long
'Public Const PS_COSMETIC = &H0
'Public Const PS_SOLID = 0
Public Declare Function InflateRect Lib "user32" (lpRect As RECT, ByVal x As Long, ByVal y As Long) As Long
Public Declare Function OffsetRect Lib "user32" (lpRect As RECT, ByVal x As Long, ByVal y As Long) As Long
'Public Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
Public Const SM_CXVSCROLL = 2
Public Const SM_CYHSCROLL = 3
Public Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Public Declare Function GetClientRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Public Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hdc As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
'#if(WINVER >= =&H0500)
'Public Const DT_NOFULLWIDTHCHARBREAK = &H80000
'#if(_WIN32_WINNT >= =&H0500)
'Public Const DT_HIDEPREFIX = &H100000
'Public Const DT_PREFIXONLY = &H200000
'#endif /* _WIN32_WINNT >= =&H0500 */
'#endif /* WINVER >= =&H0500 */
'Public Declare Function GetSysColorBrush Lib "user32" (ByVal nIndex As Long) As Long
'Public Const COLOR_HIGHLIGHT = 13
'Public Const COLOR_HIGHLIGHTTEXT = 14
'Public Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Public Declare Function FillRect Lib "user32" (ByVal hdc As Long, lpRect As RECT, ByVal hBrush As Long) As Long
'Public Declare Function SetTextColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
'Public Declare Function SetBkColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
'Public Declare Function SetBkMode Lib "gdi32" (ByVal hdc As Long, ByVal nBkMode As Long) As Long
Public Const OPAQUE = 2
Public Const TRANSPARENT = 1

'Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
'Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hdc As Long, ByVal nIndex As Long) As Long
'Private Const LOGPIXELSX = 88    '  Logical pixels/inch in X
'Private Const LOGPIXELSY = 90    '  Logical pixels/inch in Y
'Private Declare Function MulDiv Lib "kernel32" (ByVal nNumber As Long, ByVal nNumerator As Long, ByVal nDenominator As Long) As Long
'Private Const LF_FACESIZE = 32
Public Type LOGFONT
   lfHeight As Long ' The font size (see below)
   lfWidth As Long ' Normally you don't set this, just let Windows create the Default
   lfEscapement As Long ' The angle, in 0.1 degrees, of the font
   lfOrientation As Long ' Leave as default
   lfWeight As Long ' Bold, Extra Bold, Normal etc
   lfItalic As Byte ' As it says
   lfUnderline As Byte ' As it says
   lfStrikeOut As Byte ' As it says
   lfCharSet As Byte ' As it says
   lfOutPrecision As Byte ' Leave for default
   lfClipPrecision As Byte ' Leave for default
   lfQuality As Byte ' Leave for default
   lfPitchAndFamily As Byte ' Leave for default
   lfFaceName(LF_FACESIZE) As Byte ' The font name converted to a byte array
End Type
Public Declare Function CreateFontIndirect Lib "gdi32" Alias "CreateFontIndirectA" (lpLogFont As LOGFONT) As Long
'Public Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
'Public Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
'Public Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
'Public Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
'Private Const FW_NORMAL = 400
'Private Const FW_BOLD = 700
'Private Const FF_DONTCARE = 0
'Private Const DEFAULT_QUALITY = 0
'Private Const DEFAULT_PITCH = 0
'Private Const DEFAULT_CHARSET = 1
'Private Declare Function OleTranslateColor Lib "OLEPRO32.DLL" (ByVal OLE_COLOR As Long, ByVal HPALETTE As Long, pccolorref As Long) As Long
'Private Const CLR_INVALID = -1
' Corrected Draw State function declarations:
'Private Declare Function DrawState Lib "user32" Alias "DrawStateA" _
'   (ByVal hdc As Long, _
'   ByVal hBrush As Long, _
'   ByVal lpDrawStateProc As Long, _
'   ByVal lParam As Long, _
'   ByVal wParam As Long, _
'   ByVal x As Long, _
'   ByVal y As Long, _
'   ByVal cx As Long, _
'   ByVal cy As Long, _
'   ByVal fuFlags As Long) As Long
'Private Declare Function DrawStateString Lib "user32" Alias "DrawStateA" _
   (ByVal hdc As Long, _
   ByVal hBrush As Long, _
   ByVal lpDrawStateProc As Long, _
   ByVal lpString As String, _
   ByVal cbStringLen As Long, _
   ByVal x As Long, _
   ByVal y As Long, _
   ByVal cx As Long, _
   ByVal cy As Long, _
   ByVal fuFlags As Long) As Long

' Missing Draw State constants declarations:
'/* Image type */
'Private Const DST_COMPLEX = &H0
'Private Const DST_TEXT = &H1
'Private Const DST_PREFIXTEXT = &H2
'Private Const DST_ICON = &H3
'Private Const DST_BITMAP = &H4

' /* State type */
'Private Const DSS_NORMAL = &H0
'Private Const DSS_UNION = &H10
'Private Const DSS_DISABLED = &H20
'Private Const DSS_MONO = &H80
'Private Const DSS_RIGHT = &H8000

' Create a new icon based on an image list icon:
'Private Declare Function ImageList_GetIcon Lib "Comctl32.dll" ( _
'        ByVal hIml As Long, _
'        ByVal i As Long, _
'        ByVal diIgnore As Long _
'    ) As Long
' Draw an item in an ImageList:
'Private Declare Function ImageList_Draw Lib "Comctl32.dll" ( _
'        ByVal hIml As Long, _
'        ByVal i As Long, _
'        ByVal hdcDst As Long, _
'        ByVal x As Long, _
'        ByVal y As Long, _
'        ByVal fStyle As Long _
'    ) As Long
' Draw an item in an ImageList with more control over positioning
' and colour:
'Private Declare Function ImageList_DrawEx Lib "Comctl32.dll" ( _
'      ByVal hIml As Long, ByVal i As Long, ByVal hdcDst As Long, _
'      ByVal x As Long, ByVal y As Long, _
'      ByVal dx As Long, ByVal dy As Long, _
'      ByVal rgbBk As Long, _
'      ByVal rgbFg As Long, ByVal fStyle As Long) As Long
' Built in ImageList drawing methods:
'Private Const ILD_NORMAL = 0
'Private Const ILD_TRANSPARENT = 1
'Private Const ILD_BLEND25 = 2
'Private Const ILD_SELECTED = 4
'Private Const ILD_FOCUS = 4
'Private Const ILD_OVERLAYMASK = 3840
' Use default rgb colour:
'Public Const CLR_NONE = -1
'Private Declare Function DestroyIcon Lib "user32" (ByVal hIcon As Long) As Long
'Public Declare Function ImageList_GetIconSize Lib "comctl32" (ByVal hImageList As Long, cx As Long, cy As Long) As Long
'Public Declare Function ImageList_GetImageCount Lib "comctl32" (ByVal hImageList As Long) As Long

' Standard GDI draw icon function:
Public Declare Function DrawIconEx Lib "user32" (ByVal hdc As Long, ByVal xLeft As Long, ByVal yTop As Long, ByVal hIcon As Long, ByVal cxWidth As Long, ByVal cyWidth As Long, ByVal istepIfAniCur As Long, ByVal hbrFlickerFreeDraw As Long, ByVal diFlags As Long) As Long
'Private Const DI_MASK = &H1
'Private Const DI_IMAGE = &H2
'Public Const DI_NORMAL = &H3
'Private Const DI_COMPAT = &H4
'Private Const DI_DEFAULTSIZE = &H8

'Public Declare Function LoadImageByNum Lib "user32" Alias "LoadImageA" (ByVal hInst As Long, ByVal lpsz As Long, ByVal un1 As Long, ByVal n1 As Long, ByVal n2 As Long, ByVal un2 As Long) As Long
    Public Const LR_LOADMAP3DCOLORS = &H1000
    Public Const LR_LOADFROMFILE = &H10
    Public Const LR_LOADTRANSPARENT = &H20
   ' Public Const IMAGE_BITMAP = 0

'Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" ( _
'    lpvDest As Any, lpvSource As Any, ByVal cbCopy As Long)
Public Declare Function ClientToScreen Lib "user32" (ByVal hWnd As Long, lpPoint As POINTAPI) As Long
Public Declare Function ScreenToClient Lib "user32" (ByVal hWnd As Long, lpPoint As POINTAPI) As Long
'Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Const GWL_EXSTYLE = (-20)
Public Const GWL_STYLE = (-16)
'Public Const WS_EX_WINDOWEDGE = &H100
Public Const WS_EX_CLIENTEDGE = &H200
Public Const WS_EX_STATICEDGE = &H20000
Public Const WS_HSCROLL = &H100000
Public Const WS_VSCROLL = &H200000

'Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
'Public Enum ESetWindowPosStyles
'    SWP_SHOWWINDOW = &H40
'    SWP_HIDEWINDOW = &H80
'    SWP_FRAMECHANGED = &H20 ' The frame changed: send WM_NCCALCSIZE
'    SWP_NOACTIVATE = &H10
'    SWP_NOCOPYBITS = &H100
'    SWP_NOMOVE = &H2
'    SWP_NOOWNERZORDER = &H200 ' Don't do owner Z ordering
'    SWP_NOREDRAW = &H8
'    SWP_NOREPOSITION = SWP_NOOWNERZORDER
'    SWP_NOSIZE = &H1
'    SWP_NOZORDER = &H4
'    SWP_DRAWFRAME = SWP_FRAMECHANGED
'    HWND_NOTOPMOST = -2
'End Enum
' Window relationship functions:
'Public Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
'Public Declare Function GetParent Lib "user32" (ByVal hwnd As Long) As Long
' Message functions:
'Public Declare Function SendMessageByString Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long
Public Declare Function SendMessageByLong Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
'Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
'Public Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

Public Const WM_KEYDOWN = &H100
Public Const WM_KEYUP = &H101
'------
' Difference between day zero for VB dates and Win32 dates
' (or #12-30-1899# - #01-01-1601#)
Private Const rDayZeroBias As Double = 109205#    ' Abs(CDbl(#01-01-1601#))
' 10000000 nanoseconds * 60 seconds * 60 minutes * 24 hours / 10000
' comes to 86400000 (the 10000 adjusts for fixed point in Currency)
Private Const rMillisecondPerDay As Double = 10000000# * 60# * 60# * 24# / 10000#

Public Sub DrawDragImage( _
      ByRef rcNew As RECT, _
      ByVal bFirst As Boolean, _
      ByVal bLast As Boolean _
   )
Static rcCurrent As RECT
Dim hdc As Long
   
   ' First get the Desktop DC:
   hdc = CreateDCAsNull("DISPLAY", ByVal 0&, ByVal 0&, ByVal 0&)
   ' Set the draw mode to XOR:
   SetROP2 hdc, R2_NOTXORPEN
   
   '// Draw over and erase the old rectangle
   If Not (bFirst) Then
      Rectangle hdc, rcCurrent.left, rcCurrent.top, rcCurrent.right, rcCurrent.bottom
   End If
   
   If Not (bLast) Then
      '// Draw the new rectangle
      Rectangle hdc, rcNew.left, rcNew.top, rcNew.right, rcNew.bottom
   End If
   
   ' Store this position so we can erase it next time:
   LSet rcCurrent = rcNew
   
   ' Free the reference to the Desktop DC we got (make sure you do this!)
   DeleteDC hdc
    
End Sub

Public Sub DrawImage( _
      ByVal himl As Long, _
      ByVal iIndex As Long, _
      ByVal hdc As Long, _
      ByVal xPixels As Integer, _
      ByVal yPixels As Integer, _
      ByVal lIconSizeX As Long, ByVal lIconSizeY As Long, _
      Optional ByVal bSelected = False, _
      Optional ByVal bCut = False, _
      Optional ByVal bDisabled = False, _
      Optional ByVal oCutDitherColour As OLE_COLOR = vbWindowBackground, _
      Optional ByVal hExternalIml As Long = 0 _
    )
Dim hIcon As Long
Dim lFlags As Long
Dim lhIml As Long
Dim lColor As Long
Dim iImgIndex As Long

   ' Draw the image at 1 based index or key supplied in vKey.
   ' on the hDC at xPixels,yPixels with the supplied options.
   ' You can even draw an ImageList from another ImageList control
   ' if you supply the handle to hExternalIml with this function.
   
   iImgIndex = iIndex
   If (iImgIndex > -1) Then
      If (hExternalIml <> 0) Then
          lhIml = hExternalIml
      Else
          lhIml = himl
      End If
      
      lFlags = ILD_TRANSPARENT
      If (bSelected) Or (bCut) Then
          lFlags = lFlags Or ILD_SELECTED
      End If
      
      If (bCut) Then
        ' Draw dithered:
        lColor = TranslateColor(oCutDitherColour)
        If (lColor = -1) Then lColor = TranslateColor(vbWindowBackground)
        ImageList_DrawEx _
              lhIml, _
              iImgIndex, _
              hdc, _
              xPixels, yPixels, 0, 0, _
              CLR_NONE, lColor, _
              lFlags
      ElseIf (bDisabled) Then
        ' extract a copy of the icon:
        hIcon = ImageList_GetIcon(himl, iImgIndex, 0)
        ' Draw it disabled at x,y:
        DrawState hdc, 0, 0, hIcon, 0, xPixels, yPixels, lIconSizeX, lIconSizeY, DST_ICON Or DSS_DISABLED
        ' Clear up the icon:
        DestroyIcon hIcon
              
      Else
        ' Standard draw:
        ImageList_Draw _
            lhIml, _
            iImgIndex, _
            hdc, _
            xPixels, _
            yPixels, _
            lFlags
      End If
   End If
End Sub


'Public Function TranslateColor(ByVal oClr As OLE_COLOR, _
'                        Optional hPal As Long = 0) As Long
'    ' Convert Automation color to Windows color
'    If OleTranslateColor(oClr, hPal, TranslateColor) Then
'        TranslateColor = CLR_INVALID
'    End If
'End Function

Public Sub pOLEFontToLogFont(fntTHis As StdFont, hdc As Long, tLF As LOGFONT)
Dim sFont As String
Dim iChar As Integer

   ' Convert an OLE StdFont to a LOGFONT structure:
   With tLF
       sFont = fntTHis.Name
       ' There is a quicker way involving StrConv and CopyMemory, but
       ' this is simpler!:
       For iChar = 1 To Len(sFont)
           .lfFaceName(iChar - 1) = CByte(Asc(Mid$(sFont, iChar, 1)))
       Next iChar
       ' Based on the Win32SDK documentation:
       .lfHeight = -MulDiv((fntTHis.Size), (GetDeviceCaps(hdc, LOGPIXELSY)), 72)
       .lfItalic = fntTHis.Italic
       If (fntTHis.Bold) Then
           .lfWeight = FW_BOLD
       Else
           .lfWeight = FW_NORMAL
       End If
       .lfUnderline = fntTHis.Underline
       .lfStrikeOut = fntTHis.Strikethrough
       .lfCharSet = fntTHis.Charset
   End With

End Sub
Public Sub TileArea( _
        ByVal hdc As Long, _
        ByVal x As Long, _
        ByVal y As Long, _
        ByVal Width As Long, _
        ByVal Height As Long, _
        ByVal lSrcDC As Long, _
        ByVal lBitmapW As Long, _
        ByVal lBitmapH As Long, _
        ByVal lSrcOffsetX As Long, _
        ByVal lSrcOffsetY As Long _
    )
Dim lSrcX As Long
Dim lSrcY As Long
Dim lSrcStartX As Long
Dim lSrcStartY As Long
Dim lSrcStartWidth As Long
Dim lSrcStartHeight As Long
Dim lDstX As Long
Dim lDstY As Long
Dim lDstWidth As Long
Dim lDstHeight As Long

    lSrcStartX = ((x + lSrcOffsetX) Mod lBitmapW)
    lSrcStartY = ((y + lSrcOffsetY) Mod lBitmapH)
    lSrcStartWidth = (lBitmapW - lSrcStartX)
    lSrcStartHeight = (lBitmapH - lSrcStartY)
    lSrcX = lSrcStartX
    lSrcY = lSrcStartY
    
    lDstY = y
    lDstHeight = lSrcStartHeight
    
    Do While lDstY < (y + Height)
        If (lDstY + lDstHeight) > (y + Height) Then
            lDstHeight = y + Height - lDstY
        End If
        lDstWidth = lSrcStartWidth
        lDstX = x
        lSrcX = lSrcStartX
        Do While lDstX < (x + Width)
            If (lDstX + lDstWidth) > (x + Width) Then
                lDstWidth = x + Width - lDstX
                If (lDstWidth = 0) Then
                    lDstWidth = 4
                End If
            End If
            'If (lDstWidth > Width) Then lDstWidth = Width
            'If (lDstHeight > Height) Then lDstHeight = Height
            BitBlt hdc, lDstX, lDstY, lDstWidth, lDstHeight, lSrcDC, lSrcX, lSrcY, vbSrcCopy
            lDstX = lDstX + lDstWidth
            lSrcX = 0
            lDstWidth = lBitmapW
        Loop
        lDstY = lDstY + lDstHeight
        lSrcY = 0
        lDstHeight = lBitmapH
    Loop
End Sub



Public Property Get ObjectFromPtr(ByVal lPtr As Long) As Object
Dim oTemp As Object
   ' Turn the pointer into an illegal, uncounted interface
   CopyMemory oTemp, lPtr, 4
   ' Do NOT hit the End button here! You will crash!
   ' Assign to legal reference
   Set ObjectFromPtr = oTemp
   ' Still do NOT hit the End button here! You will still crash!
   ' Destroy the illegal reference
   CopyMemory oTemp, 0&, 4
   ' OK, hit the End button if you must--you'll probably still crash,
   ' but it will be because of the subclass, not the uncounted reference
End Property
Public Function IconToPicture(ByVal hIcon As Long) As IPicture
    
    If hIcon = 0 Then Exit Function
        
    ' This is all magic if you ask me:
    Dim NewPic As Picture, PicConv As PictDesc, IGuid As Guid
    
    PicConv.cbSizeofStruct = Len(PicConv)
    PicConv.picType = vbPicTypeIcon
    PicConv.hImage = hIcon
    
    ' Fill in magic IPicture GUID {7BF80980-BF32-101A-8BBB-00AA00300CAB}
    With IGuid
        .Data1 = &H7BF80980
        .Data2 = &HBF32
        .Data3 = &H101A
        .Data4(0) = &H8B
        .Data4(1) = &HBB
        .Data4(2) = &H0
        .Data4(3) = &HAA
        .Data4(4) = &H0
        .Data4(5) = &H30
        .Data4(6) = &HC
        .Data4(7) = &HAB
    End With
    OleCreatePictureIndirect PicConv, IGuid, True, NewPic
    
    Set IconToPicture = NewPic
    
End Function

Public Function BitmapToPicture(ByVal hBmp As Long) As IPicture

   If (hBmp = 0) Then Exit Function
   
   Dim NewPic As Picture, tPicConv As PictDesc, IGuid As Guid
   
   ' Fill PictDesc structure with necessary parts:
   With tPicConv
      .cbSizeofStruct = Len(tPicConv)
      .picType = vbPicTypeBitmap
      .hImage = hBmp
   End With
   
   ' Fill in IDispatch Interface ID
   With IGuid
      .Data1 = &H20400
      .Data4(0) = &HC0
      .Data4(7) = &H46
   End With
   
   ' Create a picture object:
   OleCreatePictureIndirect tPicConv, IGuid, True, NewPic
   
   ' Return it:
   Set BitmapToPicture = NewPic
      

End Function

Public Function TranslateColor(ByVal clr As OLE_COLOR, _
                        Optional hPal As Long = 0) As Long
   If OleTranslateColor(clr, hPal, TranslateColor) Then
      TranslateColor = CLR_INVALID
   End If
End Function
Public Function SerialiseIcon( _
      ByVal lHDC As Long, _
      ByVal hIcon As Long, _
      ByRef b() As Byte, _
      ByVal lByteStart As Long, _
      ByRef lArraySize As Long _
   ) As Boolean
Dim tII As ICONINFO
Dim lR As Long
Dim lMonoSize As Long
Dim lColourSize As Long
   
   ' decompose icon:
   lR = GetIconInfo(hIcon, tII)
   If (lR <> 0) Then
      ' store fIcon, xHotspot, yHotspot:
      CopyMemory b(lByteStart), tII, 12
      ' store the colour bitmap:
      lByteStart = lByteStart + 12
      If (SerialiseBitmap(lHDC, tII.hbmColor, False, b(), lByteStart, lColourSize)) Then
         lByteStart = lByteStart + lColourSize
         If (SerialiseBitmap(lHDC, tII.hBmMask, True, b(), lByteStart, lMonoSize)) Then
            lByteStart = lByteStart + lMonoSize
            lArraySize = lColourSize + lMonoSize + 12
            SerialiseIcon = True
         End If
      End If
      DeleteObject tII.hbmColor
      DeleteObject tII.hBmMask
   End If
End Function
Private Function SerialiseBitmap( _
      ByVal lHDC As Long, _
      ByVal hBm As Long, _
      ByVal bMono As Boolean, _
      ByRef b() As Byte, _
      ByVal lByteStart As Long, _
      ByRef lByteSize As Long _
   ) As Boolean
Dim tbm As BITMAP
Dim tBI1 As BITMAPINFO_1BPP
'Dim tBI4 As BITMAPINFO_4BPP
'Dim tBI8 As BITMAPINFO_8BPP
Dim tBI As BITMAPINFO_ABOVE8
Dim lSize As Long
Dim lR As Long
   
   ' Get the BITMAP structure:
   lR = GetObjectAPI(hBm, Len(tbm), tbm)
   If (lR <> 0) Then
      ' Store the BITMAP structure:
      CopyMemory b(lByteStart), tbm, Len(tbm)
      ' Create a bitmap info structure:
      If (bMono) Then
         With tBI1.bmiHeader
            .biSize = Len(tBI1.bmiHeader)
            .biWidth = tbm.bmWidth
            .biHeight = tbm.bmHeight
            .biPlanes = 1
            .biBitCount = 1
            .biCompression = BI_RGB
         End With
         lSize = (tBI1.bmiHeader.biWidth + 7) / 8
         lSize = ((lSize + 3) \ 4) * 4
         lSize = lSize * tBI1.bmiHeader.biHeight
         lR = GetDIBits(lHDC, hBm, 0, tbm.bmHeight, b(lByteStart + Len(tbm)), tBI1, DIB_RGB_COLORS)
      Else
         With tBI.bmiHeader
            .biSize = Len(tBI.bmiHeader)
            .biWidth = tbm.bmWidth
            .biHeight = tbm.bmHeight
            .biPlanes = 1
            .biBitCount = 24
            .biCompression = BI_RGB
         End With
         ' Get the Bitmap bits into the byte array:
         lSize = tBI.bmiHeader.biWidth
         lSize = lSize * 3
         lSize = ((lSize + 3) / 4) * 4
         lSize = lSize * tBI.bmiHeader.biHeight
         'lR = GetBitmapBits(hBm, lSize, b(lByteStart + Len(tBM)))
         lR = GetDIBits(lHDC, hBm, 0, tbm.bmHeight, b(lByteStart + Len(tbm)), tBI, DIB_RGB_COLORS)
      End If
      
      If (lR <> 0) Then
         ' Success.  Return size:
         lByteSize = lSize + Len(tbm)
         SerialiseBitmap = True
      End If
   End If

End Function
Public Function DeSerialiseIcon( _
      ByVal lHDC As Long, _
      ByRef hIcon As Long, _
      ByRef b() As Byte, _
      ByVal lByteStart As Long, _
      ByRef lArraySize As Long _
   )
Dim tII As ICONINFO
Dim lColourSize As Long
Dim lMonoSize As Long

   hIcon = 0
   ' get fIcon, xHotspot, yHotspot:
   CopyMemory tII, b(lByteStart), 12
   tII.fIcon = 1
   lByteStart = lByteStart + 12
   ' get the colour bitmap:
   If (DeSerialiseBitmap(lHDC, tII.hbmColor, False, b(), lByteStart, lColourSize)) Then
      lByteStart = lByteStart + lColourSize
      ' get the mono bitmap:
      If (DeSerialiseBitmap(lHDC, tII.hBmMask, True, b(), lByteStart, lMonoSize)) Then
         ' Set the size:
         lArraySize = lColourSize + lMonoSize + 12
         
         ' Create the icon from the structure:
         hIcon = CreateIconIndirect(tII)
         DeSerialiseIcon = (hIcon <> 0)
        
         DeleteObject tII.hbmColor
         DeleteObject tII.hBmMask
        
      Else
         DeleteObject tII.hbmColor
      End If
   End If
   
End Function
Private Function DeSerialiseBitmap( _
      ByVal lHDC As Long, _
      ByRef hBm As Long, _
      ByVal bMono As Boolean, _
      ByRef b() As Byte, _
      ByVal lByteStart As Long, _
      ByRef lByteSize As Long _
   ) As Boolean
Dim tbm As BITMAP
Dim tBI1 As BITMAPINFO_1BPP
'Dim tBI4 As BITMAPINFO_4BPP
'Dim tBI8 As BITMAPINFO_8BPP
Dim tBI As BITMAPINFO_ABOVE8
Dim lSize As Long
Dim lR As Long
   
   'Debug.Print lByteStart, lByteSize
   ' Get the BITMAP structure:
   CopyMemory tbm, b(lByteStart), Len(tbm)
   ' Create the bitmap:
   If Not (bMono) Then
      hBm = CreateCompatibleBitmap(lHDC, tbm.bmWidth, tbm.bmHeight)
   Else
      hBm = CreateBitmapIndirect(tbm)
   End If
   If (hBm <> 0) Then
      ' Get the Bitmap bits from the byte array:
      'lSize = tBM.bmWidthBytes * tBM.bmHeight
      'lR = SetBitmapBits(hBm, lSize, b(lByteStart + Len(tBM)))
      If (bMono) Then
         With tBI1.bmiHeader
            .biSize = Len(tBI1.bmiHeader)
            .biWidth = tbm.bmWidth
            .biHeight = tbm.bmHeight
            .biPlanes = 1
            .biBitCount = 1
            .biCompression = BI_RGB
         End With
         lSize = (tBI1.bmiHeader.biWidth + 7) / 8
         lSize = ((lSize + 3) \ 4) * 4
         lSize = lSize * tBI1.bmiHeader.biHeight

         tBI1.bmiColors(1).rgbBlue = 255
         tBI1.bmiColors(1).rgbGreen = 255
         tBI1.bmiColors(1).rgbRed = 255
         lR = SetDIBits(lHDC, hBm, 0, tbm.bmHeight, b(lByteStart + Len(tbm)), tBI1, DIB_RGB_COLORS)
      Else
         With tBI.bmiHeader
            .biSize = Len(tBI.bmiHeader)
            .biWidth = tbm.bmWidth
            .biHeight = tbm.bmHeight
            .biPlanes = 1
            .biBitCount = 24
            .biCompression = BI_RGB
         End With
         
         lSize = tBI.bmiHeader.biWidth
         lSize = lSize * 3
         lSize = ((lSize + 3) / 4) * 4
         lSize = lSize * tBI.bmiHeader.biHeight
         
         lR = SetDIBits(lHDC, hBm, 0, tbm.bmHeight, b(lByteStart + Len(tbm)), tBI, DIB_RGB_COLORS)
      End If
      
      lByteSize = lSize + Len(tbm)
      If (lR <> 0) Then
         DeSerialiseBitmap = True
      Else
         DeleteObject hBm
      End If
   End If
   
End Function





Public Function GetTempFile(Optional Prefix As String) As String
Dim PathName As String
Dim sRet As String

    If Prefix = "" Then Prefix = ""
    PathName = GetTempDir
    
    sRet = String(MAX_PATH, 0)
    GetTempFileName PathName, Prefix, 0, sRet
    GetTempFile = StrZToStr(sRet)
    
End Function

Private Function GetTempDir() As String
Dim sRet As String, C As Long
    sRet = String(MAX_PATH, 0)
    C = GetTempPath(MAX_PATH, sRet)
    If C = 0 Then
        GetTempDir = App.Path
    Else
        GetTempDir = left$(sRet, C)
    End If
End Function
'Private Function StrZToStr(s As String) As String
'    StrZToStr = Left$(s, lstrlen(s))
'End Function

Private Property Get TbarMenuFromPtr(ByVal lPtr As Long) As cTbarMenu
Dim oTemp As Object
   ' Turn the pointer into an illegal, uncounted interface
   CopyMemory oTemp, lPtr, 4
   ' Do NOT hit the End button here! You will crash!
   ' Assign to legal reference
   Set TbarMenuFromPtr = oTemp
   ' Still do NOT hit the End button here! You will still crash!
   ' Destroy the illegal reference
   CopyMemory oTemp, 0&, 4
   ' OK, hit the End button if you must--you'll probably still crash,
   ' but it will be because of the subclass, not the uncounted reference
End Property
Private Property Get TbarFromPtr(ByVal lPtr As Long) As cToolbar
Dim oTemp As Object
   ' Turn the pointer into an illegal, uncounted interface
   CopyMemory oTemp, lPtr, 4
   ' Do NOT hit the End button here! You will crash!
   ' Assign to legal reference
   Set TbarFromPtr = oTemp
   ' Still do NOT hit the End button here! You will still crash!
   ' Destroy the illegal reference
   CopyMemory oTemp, 0&, 4
   ' OK, hit the End button if you must--you'll probably still crash,
   ' but it will be because of the subclass, not the uncounted reference
End Property

'////////////////
'// Menu filter hook just passes to virtual CMenuBar function
'//
Private Function MenuInputFilter(ByVal nCode As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Dim cM As cTbarMenu
Dim lpMsg As Msg
   If nCode = MSGF_MENU Then
      Set cM = TbarMenuFromPtr(m_lMsgHookPtr)
      CopyMemory lpMsg, ByVal lParam, Len(lpMsg)
      If (cM.MenuInput(lpMsg)) Then
         MenuInputFilter = 1
         Exit Function
      End If
   End If
   MenuInputFilter = CallNextHookEx(m_hMsgHook, nCode, wParam, lParam)
End Function
Public Sub AttachMsgHook(cThis As cTbarMenu)
Dim lpfn As Long
   DetachMsgHook
   m_lMsgHookPtr = ObjPtr(cThis)
   lpfn = HookAddress(AddressOf MenuInputFilter)
   m_hMsgHook = SetWindowsHookEx(WH_MSGFILTER, lpfn, 0&, GetCurrentThreadId())
   Debug.Assert (m_hMsgHook <> 0)
End Sub
Public Sub DetachMsgHook()
   If (m_hMsgHook <> 0) Then
      UnhookWindowsHookEx m_hMsgHook
      m_hMsgHook = 0
   End If
End Sub
Private Function KeyboardFilter(ByVal nCode As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Dim bKeyUp As Boolean
Dim bAlt As Boolean, bCtrl As Boolean, bShift As Boolean
Dim cT As cToolbar
Dim i As Long

On Error GoTo ErrorHandler

   If nCode = HC_ACTION And m_iKeyHookCount > 0 Then
      ' Key up or down:
      bKeyUp = ((lParam And &H80000000) = &H80000000)
      ' Alt pressed?
      bAlt = ((lParam And &H20000000) = &H20000000)
      bCtrl = (GetAsyncKeyState(vbKeyControl) <> 0)
      bShift = (GetAsyncKeyState(vbKeyShift) <> 0)
      If Not (bKeyUp Or bCtrl Or bShift) And bAlt Then
         ' Alt- key pressed:
         For i = 1 To m_iKeyHookCount
            If m_lKeyHookPtr(i) <> 0 Then
               Set cT = TbarFromPtr(m_lKeyHookPtr(i))
               If Not cT Is Nothing Then
                  Debug.Print "KeyboardFilter: AltKeyPress"
                  If cT.AltKeyPress(wParam) Then
                     ' Eat message
                     KeyboardFilter = 1
                     Exit Function
                  End If
               End If
            End If
         Next i
      End If
   End If
   KeyboardFilter = CallNextHookEx(m_hKeyHook, nCode, wParam, lParam)

   Exit Function
   
ErrorHandler:
   Exit Function

End Function
Public Sub AttachKeyboardHook(cThis As cToolbar)
Dim lpfn As Long
Dim lPtr As Long
Dim i As Long
   
   If m_iKeyHookCount = 0 Then
      lpfn = HookAddress(AddressOf KeyboardFilter)
      m_hKeyHook = SetWindowsHookEx(WH_KEYBOARD, lpfn, 0&, GetCurrentThreadId())
      Debug.Assert (m_hKeyHook <> 0)
   End If
   lPtr = ObjPtr(cThis)
   For i = 1 To m_iKeyHookCount
      If lPtr = m_lKeyHookPtr(i) Then
         ' we already have it:
         Debug.Assert False
         Exit Sub
      End If
   Next i
   ReDim Preserve m_lKeyHookPtr(1 To m_iKeyHookCount + 1) As Long
   m_iKeyHookCount = m_iKeyHookCount + 1
   m_lKeyHookPtr(m_iKeyHookCount) = lPtr
   
End Sub
Public Sub DetachKeyboardHook(cThis As cToolbar)
Dim i As Long
Dim lPtr As Long
Dim iThis As Long
   
   lPtr = ObjPtr(cThis)
   For i = 1 To m_iKeyHookCount
      If m_lKeyHookPtr(i) = lPtr Then
         iThis = i
         Exit For
      End If
   Next i
   If iThis <> 0 Then
      If m_iKeyHookCount > 1 Then
         For i = iThis To m_iKeyHookCount - 1
            m_lKeyHookPtr(i) = m_lKeyHookPtr(i + 1)
         Next i
      End If
      m_iKeyHookCount = m_iKeyHookCount - 1
      If m_iKeyHookCount >= 1 Then
         ReDim Preserve m_lKeyHookPtr(1 To m_iKeyHookCount) As Long
      Else
         Erase m_lKeyHookPtr
      End If
   Else
      ' Trying to detach a toolbar which was never attached...
      ' This will happen at design time
   End If
   
   If m_iKeyHookCount <= 0 Then
      If (m_hKeyHook <> 0) Then
         UnhookWindowsHookEx m_hKeyHook
         m_hKeyHook = 0
      End If
   End If
   
End Sub
Private Function HookAddress(ByVal lPtr As Long) As Long
   HookAddress = lPtr
End Function

Public Sub AddRebar( _
      ByVal hWnd As Long, _
      ByVal hWndParent As Long _
   )
   m_iRebarCount = m_iRebarCount + 1
   ReDim Preserve m_tRebarInter(1 To m_iRebarCount) As tRebarInter
   With m_tRebarInter(m_iRebarCount)
      .hWndParent = hWndParent
      .hWndRebar = hWnd
   End With
End Sub
Public Sub RemoveRebar( _
      ByVal hWnd As Long _
   )
Dim i As Long
Dim iT As Long
   For i = 1 To m_iRebarCount
      If m_tRebarInter(i).hWndRebar = hWnd Then
      Else
         iT = iT + 1
         If (iT <> i) Then
            LSet m_tRebarInter(iT) = m_tRebarInter(i)
         End If
      End If
   Next i
   
   If iT <> m_iRebarCount Then
      m_iRebarCount = iT
      If iT = 0 Then
         Erase m_tRebarInter
      Else
         ReDim Preserve m_tRebarInter(1 To m_iRebarCount) As tRebarInter
      End If
   End If
End Sub
Public Sub AdjustForOtherRebars( _
      ByVal hWnd As Long, _
      ByRef lLeft As Long, ByRef lTop As Long, _
      ByRef lWidth As Long, ByRef lHeight As Long _
   )
Dim i As Long
Dim iIndex As Long
Dim hWndP As Long
Dim lThisP As Long
Dim lP As Long
Dim rc As RECT, rcP As RECT

   m_lPad = 2
   
   For i = 1 To m_iRebarCount
      If m_tRebarInter(i).hWndRebar = hWnd Then
         iIndex = i
         hWndP = m_tRebarInter(i).hWndParent
         lThisP = GetProp(hWnd, "vbal:cRebarPosition")
         Exit For
      End If
   Next i
   
   If iIndex >= 1 Then
      GetWindowRect hWndP, rcP
      For i = 1 To iIndex - 1
         If m_tRebarInter(i).hWndParent = hWndP Then
            If IsWindowVisible(m_tRebarInter(i).hWndRebar) Then
               GetWindowRect m_tRebarInter(i).hWndRebar, rc
               lP = GetProp(m_tRebarInter(i).hWndRebar, "vbal:cRebarPosition")
               Select Case lThisP
               Case 0 'top
                  Select Case lP
                  Case 0
                     lTop = lTop + rc.bottom - rc.top + m_lPad
                  Case 1
                     lLeft = lLeft + rc.right - rc.left + m_lPad
                     lWidth = lWidth - (rc.right - rc.left + m_lPad)
                  Case 2
                     lWidth = lWidth - (rc.right - rc.left + m_lPad)
                  End Select
               Case 1 'left
                  Select Case lP
                  Case 0
                     lTop = lTop + rc.bottom - rc.top + m_lPad
                     lHeight = lHeight - (rc.bottom - rc.top + m_lPad)
                  Case 1
                     lLeft = lLeft + rc.right - rc.left + m_lPad
                  Case 3
                     lHeight = lHeight - (rc.bottom - rc.top + m_lPad)
                  End Select
               Case 2 'right
                  Select Case lP
                  Case 0
                     lTop = lTop + rc.bottom - rc.top + m_lPad
                     lHeight = lHeight - (rc.bottom - rc.top + m_lPad)
                  Case 2
                     lLeft = lLeft - (rc.right - rc.left + m_lPad)
                  Case 3
                     lHeight = lHeight - (rc.bottom - rc.top + m_lPad)
                  End Select
               Case 3 'bottom
                  Select Case lP
                  Case 1
                     lLeft = lLeft + (rc.right - rc.left + m_lPad)
                     lWidth = lWidth - (rc.right - rc.left + m_lPad)
                  Case 2
                     lWidth = lWidth - (rc.right - rc.left + m_lPad)
                  Case 3
                     lTop = lTop - (rc.bottom - rc.top + m_lPad)
                  End Select
               End Select
            End If
         End If
      Next i
   End If
   
End Sub

Public Function ComCtlVersion( _
        ByRef lMajor As Long, _
        ByRef lMinor As Long, _
        Optional ByRef lBuild As Long _
    ) As Boolean
Dim hmod As Long
Dim lR As Long
Dim lptrDLLVersion As Long
Dim tDVI As DLLVERSIONINFO

   lMajor = 0: lMinor = 0: lBuild = 0
   
   hmod = LoadLibrary("comctl32.dll")
   If (hmod <> 0) Then
      lR = S_OK
      '/*
      ' You must get this function explicitly because earlier versions of the DLL
      ' don't implement this function. That makes the lack of implementation of the
      ' function a version marker in itself. */
      lptrDLLVersion = GetProcAddress(hmod, "DllGetVersion")
      If (lptrDLLVersion <> 0) Then
         tDVI.cbSize = Len(tDVI)
         lR = DllGetVersion(tDVI)
         If (lR = S_OK) Then
            lMajor = tDVI.dwMajor
            lMinor = tDVI.dwMinor
            lBuild = tDVI.dwBuildNumber
         End If
      Else
         'If GetProcAddress failed, then the DLL is a version previous to the one
         'shipped with IE 3.x.
         lMajor = 4
      End If
      FreeLibrary hmod
      ComCtlVersion = True
   End If

End Function

Public Property Get NewButtonID() As Long
   m_iID = m_iID + 1
   NewButtonID = m_iID
End Property

Public Property Get hwndToolTip() As Long
   If m_hWndToolTip = 0 Then
      Create
   End If
   hwndToolTip = m_hWndToolTip
End Property
Public Sub AddToToolTip(ByVal hWnd As Long)
Dim tTi As TOOLINFO

   If m_hWndToolTip = 0 Then
      Create
   End If
    
   With tTi
      .cbSize = Len(tTi)
      .uID = hWnd
      .hWnd = hWnd
      .hInst = App.hInstance
      .uFlags = TTF_IDISHWND
      .lpszText = LPSTR_TEXTCALLBACK
   End With
   
   SendMessage m_hWndToolTip, TTM_ADDTOOL, 0, tTi
   SendMessageLong m_hWndToolTip, TTM_ACTIVATE, 1, 0
   m_iRef = m_iRef + 1

End Sub
Public Sub RemoveFromToolTip(ByVal hWnd As Long)
Dim tTi As TOOLINFO
   If m_hWndToolTip <> 0 Then
      With tTi
         .cbSize = Len(tTi)
         .uID = hWnd
         .hWnd = hWnd
      End With
      SendMessage m_hWndToolTip, TTM_DELTOOL, 0, tTi
      
      m_iRef = m_iRef - 1
      If m_iRef <= 0 Then
         DestroyWindow m_hWndToolTip
         m_hWndToolTip = 0
         m_iRef = 0
      End If
   End If
End Sub
 
Public Sub Create()
   ' Create the tooltip:
   InitCommonControls
   m_hWndToolTip = CreateWindowEX(WS_EX_TOPMOST, TOOLTIPS_CLASS, vbNullString, 0, _
             CW_USEDEFAULT, CW_USEDEFAULT, _
             CW_USEDEFAULT, CW_USEDEFAULT, _
             0, 0, _
             App.hInstance, _
             ByVal 0)
   SendMessage m_hWndToolTip, TTM_ACTIVATE, 1, ByVal 0
End Sub

Public Function VBGetOpenFileName(Filename As String, _
                           Optional FileTitle As String, _
                           Optional FileMustExist As Boolean = True, _
                           Optional MultiSelect As Boolean = False, _
                           Optional ReadOnly As Boolean = False, _
                           Optional HideReadOnly As Boolean = False, _
                           Optional Filter As String = "All (*.*)| *.*", _
                           Optional FilterIndex As Long = 1, _
                           Optional InitDir As String, _
                           Optional DlgTitle As String, _
                           Optional DefaultExt As String, _
                           Optional Owner As Long = -1, _
                           Optional flags As Long = 0) As Boolean

    Dim opfile As OPENFILENAME, s As String, afFlags As Long
    Dim lMax As Long
    
    m_lApiReturn = 0
    m_lExtendedError = 0

With opfile
    .lStructSize = Len(opfile)
    
    ' Add in specific flags and strip out non-VB flags
    
    .flags = (-FileMustExist * OFN_FILEMUSTEXIST) Or _
            (-MultiSelect * OFN_ALLOWMULTISELECT) Or _
             (-ReadOnly * OFN_READONLY) Or _
             (-HideReadOnly * OFN_HIDEREADONLY) Or _
             (flags And CLng(Not (OFN_ENABLEHOOK Or _
                                  OFN_ENABLETEMPLATE)))
    ' Owner can take handle of owning window
    If Owner <> -1 Then .hwndOwner = Owner
    ' InitDir can take initial directory string
    .lpstrInitialDir = InitDir
    ' DefaultExt can take default extension
    .lpstrDefExt = DefaultExt
    ' DlgTitle can take dialog box title
    .lpstrTitle = DlgTitle
    
    ' To make Windows-style filter, replace | and : with nulls
    Dim ch As String, i As Integer
    For i = 1 To Len(Filter)
        ch = Mid$(Filter, i, 1)
        If ch = "|" Or ch = ":" Then
            s = s & vbNullChar
        Else
            s = s & ch
        End If
    Next
    ' Put double null at end
    s = s & vbNullChar & vbNullChar
    .lpstrFilter = s
    .nFilterIndex = FilterIndex

    ' Pad file and file title buffers to maximum path
    lMax = MAX_PATH
    If (.flags And OFN_ALLOWMULTISELECT) = OFN_ALLOWMULTISELECT Then
      lMax = 8192
    End If
    s = Filename & String$(lMax - Len(Filename), 0)
    .lpstrFile = s
    .nMaxFile = lMax
    s = FileTitle & String$(lMax - Len(FileTitle), 0)
    .lpstrFileTitle = s
    .nMaxFileTitle = lMax
    ' All other fields set to zero
    
    m_lApiReturn = GetOpenFileName(opfile)
    Select Case m_lApiReturn
    Case 1
        ' Success
        VBGetOpenFileName = True
        If (.flags And OFN_ALLOWMULTISELECT) = OFN_ALLOWMULTISELECT Then
            FileTitle = ""
            lMax = InStr(.lpstrFile, Chr$(0) & Chr$(0))
            If (lMax = 0) Then
               Filename = StrZToStr(.lpstrFile)
            Else
               Filename = left$(.lpstrFile, lMax - 1)
            End If
        Else
            Filename = StrZToStr(.lpstrFile)
            FileTitle = StrZToStr(.lpstrFileTitle)
        End If
        flags = .flags
        ' Return the filter index
        FilterIndex = .nFilterIndex
        ' Look up the filter the user selected and return that
        Filter = FilterLookup(.lpstrFilter, FilterIndex)
        If (.flags And OFN_READONLY) Then ReadOnly = True
    Case 0
        ' Cancelled
        VBGetOpenFileName = False
        Filename = ""
        FileTitle = ""
        flags = 0
        FilterIndex = -1
        Filter = ""
    Case Else
        ' Extended error
        m_lExtendedError = CommDlgExtendedError()
        VBGetOpenFileName = False
        Filename = ""
        FileTitle = ""
        flags = 0
        FilterIndex = -1
        Filter = ""
    End Select
End With
End Function

Private Function StrZToStr(s As String) As String
    StrZToStr = left$(s, lstrlen(s))
End Function

Private Function FilterLookup(ByVal sFilters As String, ByVal iCur As Long) As String
    Dim iStart As Long, iEnd As Long, s As String
    iStart = 1
    If sFilters = "" Then Exit Function
    Do
        ' Cut out both parts marked by null character
        iEnd = InStr(iStart, sFilters, vbNullChar)
        If iEnd = 0 Then Exit Function
        iEnd = InStr(iEnd + 1, sFilters, vbNullChar)
        If iEnd Then
            s = Mid$(sFilters, iStart, iEnd - iStart)
        Else
            s = Mid$(sFilters, iStart)
        End If
        iStart = iEnd + 1
        If iCur = 1 Then
            FilterLookup = s
            Exit Function
        End If
        iCur = iCur - 1
    Loop While iCur
End Function

Function VBChooseColor(Color As Long, _
                       Optional AnyColor As Boolean = True, _
                       Optional FullOpen As Boolean = False, _
                       Optional DisableFullOpen As Boolean = False, _
                       Optional Owner As Long = -1, _
                       Optional flags As Long) As Boolean
Dim chclr As TCHOOSECOLOR

    chclr.lStructSize = Len(chclr)
    
    ' Color must get reference variable to receive result
    ' Flags can get reference variable or constant with bit flags
    ' Owner can take handle of owning window
    If Owner <> -1 Then chclr.hwndOwner = Owner

    ' Assign color (default uninitialized value of zero is good default)
    chclr.rgbResult = Color

    ' Mask out unwanted bits
    Dim afMask As Long
    afMask = CLng(Not (CC_ENABLEHOOK Or _
                       CC_ENABLETEMPLATE))
    ' Pass in flags
    chclr.flags = afMask And (CC_RGBInit Or _
                  IIf(AnyColor, CC_AnyColor, CC_SolidColor) Or _
                  (-FullOpen * CC_FullOpen) Or _
                  (-DisableFullOpen * CC_PreventFullOpen))

   ' If first time, initialize to white
   If fNotFirst = False Then
      InitColors
   End If

    chclr.lpCustColors = VarPtr(alCustom(0))
    ' All other fields zero
    
    m_lApiReturn = ChooseColor(chclr)
    Select Case m_lApiReturn
    Case 1
        ' Success
        VBChooseColor = True
        Color = chclr.rgbResult
    Case 0
        ' Cancelled
        VBChooseColor = False
        Color = -1
    Case Else
        ' Extended error
        m_lExtendedError = CommDlgExtendedError()
        VBChooseColor = False
        Color = -1
    End Select

End Function
Private Sub InitColors()
    Dim i As Integer
    ' Initialize with first 16 system interface colors
    For i = 0 To 15
        alCustom(i) = GetSysColor(i)
    Next
    fNotFirst = True
End Sub



Public Sub debugmsg(ByVal smsg As String)
#If DEBUGMODE = 1 Then
   MsgBox smsg
#Else
   Debug.Print smsg
#End If
End Sub




