VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cIconMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit



Private Type POINTAPI
   x As Long
   y As Long
End Type
Private Type RECT
   left As Long
   tOp As Long
   Right As Long
   Bottom As Long
End Type

' =======================================================================
' MENU private declares:
' =======================================================================

' Menu flag constants:
Private Const MF_APPEND = &H100&
Private Const MF_BITMAP = &H4&
Private Const MF_BYCOMMAND = &H0&
Private Const MF_BYPOSITION = &H400&
Private Const MF_CALLBACKS = &H8000000
Private Const MF_CHANGE = &H80&
Private Const MF_CHECKED = &H8&
Private Const MF_CONV = &H40000000
Private Const MF_DELETE = &H200&
Private Const MF_DISABLED = &H2&
Private Const MF_ENABLED = &H0&
Private Const MF_END = &H80
Private Const MF_ERRORS = &H10000000
Private Const MF_GRAYED = &H1&
Private Const MF_HELP = &H4000&
Private Const MF_HILITE = &H80&
Private Const MF_HSZ_INFO = &H1000000
Private Const MF_INSERT = &H0&
Private Const MF_LINKS = &H20000000
Private Const MF_MASK = &HFF000000
Private Const MF_MENUBARBREAK = &H20&
Private Const MF_MENUBREAK = &H40&
Private Const MF_MOUSESELECT = &H8000&
Private Const MF_OWNERDRAW = &H100&
Private Const MF_POPUP = &H10&
Private Const MF_POSTMSGS = &H4000000
Private Const MF_REMOVE = &H1000&
Private Const MF_SENDMSGS = &H2000000
Private Const MF_SEPARATOR = &H800&
Private Const MF_STRING = &H0&
Private Const MF_SYSMENU = &H2000&
Private Const MF_UNCHECKED = &H0&
Private Const MF_UNHILITE = &H0&
Private Const MF_USECHECKBITMAPS = &H200&
Private Const MF_DEFAULT = &H1000&

Private Const MFT_STRING = MF_STRING
Private Const MFT_BITMAP = MF_BITMAP
Private Const MFT_MENUBARBREAK = MF_MENUBARBREAK
Private Const MFT_MENUBREAK = MF_MENUBREAK
Private Const MFT_OWNERDRAW = MF_OWNERDRAW
Private Const MFT_RADIOCHECK = &H200&
Private Const MFT_SEPARATOR = MF_SEPARATOR
Private Const MFT_RIGHTORDER = &H2000&

' New versions of the names...
Private Const MFS_GRAYED = &H3&
Private Const MFS_DISABLED = MFS_GRAYED
Private Const MFS_CHECKED = MF_CHECKED
Private Const MFS_HILITE = MF_HILITE
Private Const MFS_ENABLED = MF_ENABLED
Private Const MFS_UNCHECKED = MF_UNCHECKED
Private Const MFS_UNHILITE = MF_UNHILITE
Private Const MFS_DEFAULT = MF_DEFAULT

' MenuItemInfo Mask constants
Private Const MIIM_STATE = &H1&
Private Const MIIM_ID = &H2&
Private Const MIIM_SUBMENU = &H4&
Private Const MIIM_CHECKMARKS = &H8&
Private Const MIIM_TYPE = &H10&
Private Const MIIM_DATA = &H20&

Private Const SC_RESTORE = &HF120&
Private Const SC_MOVE = &HF010&
Private Const SC_SIZE = &HF000&
Private Const SC_MAXIMIZE = &HF030&
Private Const SC_MINIMIZE = &HF020&
Private Const SC_CLOSE = &HF060&
     
Private Const SC_ARRANGE = &HF110&
Private Const SC_HOTKEY = &HF150&
Private Const SC_HSCROLL = &HF080&
Private Const SC_KEYMENU = &HF100&
Private Const SC_MOUSEMENU = &HF090&
Private Const SC_NEXTWINDOW = &HF040&
Private Const SC_PREVWINDOW = &HF050&
Private Const SC_SCREENSAVE = &HF140&
Private Const SC_TASKLIST = &HF130&
Private Const SC_VSCROLL = &HF070&
Private Const SC_ZOOM = SC_MAXIMIZE
Private Const SC_ICON = SC_MINIMIZE

' Owner draw information:
Private Const ODS_CHECKED = &H8
Private Const ODS_DISABLED = &H4
Private Const ODS_FOCUS = &H10
Private Const ODS_GRAYED = &H2
Private Const ODS_SELECTED = &H1
Private Const ODT_BUTTON = 4
Private Const ODT_COMBOBOX = 3
Private Const ODT_LISTBOX = 2
Private Const ODT_MENU = 1

Private Type MEASUREITEMSTRUCT
   CtlType As Long
   CtlID As Long
   itemID As Long
   itemWidth As Long
   itemHeight As Long
   ItemData As Long
End Type

Private Type DRAWITEMSTRUCT
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

Private Type MENUITEMINFO
   cbSize As Long
   fMask As Long
   fType As Long
   fState As Long
   wID As Long
   hSubMenu As Long
   hbmpChecked As Long
   hbmpUnchecked As Long
   dwItemData As Long
   dwTypeData As Long
   cch As Long
End Type
Private Type MENUITEMINFO_STRINGDATA
   cbSize As Long
   fMask As Long
   fType As Long
   fState As Long
   wID As Long
   hSubMenu As Long
   hbmpChecked As Long
   hbmpUnchecked As Long
   dwItemData As Long
   dwTypeData As String
   cch As Long
End Type

Private Type MENUITEMTEMPLATE
   mtOption As Integer
   mtID As Integer
   mtString As Byte
End Type
Private Type MENUITEMTEMPLATEHEADER
   versionNumber As Integer
   Offset As Integer
End Type

Private Declare Function GetMenu Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function GetSystemMenu Lib "user32" (ByVal hWnd As Long, ByVal bRevert As Long) As Long
Private Declare Function SetMenu Lib "user32" (ByVal hWnd As Long, ByVal hMenu As Long) As Long

Private Declare Function GetMenuItemID Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Private Declare Function GetMenuCheckMarkDimensions Lib "user32" () As Long
Private Declare Function GetMenuContextHelpId Lib "user32" (ByVal hMenu As Long) As Long
Private Declare Function GetMenuDefaultItem Lib "user32" (ByVal hMenu As Long, ByVal fByPos As Long, ByVal gmdiFlags As Long) As Long
Private Declare Function GetMenuItemCount Lib "user32" (ByVal hMenu As Long) As Long
Private Declare Function GetMenuItemInfo Lib "user32" Alias "GetMenuItemInfoA" (ByVal hMenu As Long, ByVal uItem As Long, ByVal fByPosition As Boolean, lpMenuItemInfo As MENUITEMINFO) As Long
Private Declare Function GetMenuItemInfoStr Lib "user32" Alias "GetMenuItemInfoA" (ByVal hMenu As Long, ByVal uItem As Long, ByVal fByPosition As Boolean, lpMenuItemInfo As MENUITEMINFO_STRINGDATA) As Long
Private Declare Function SetMenuItemInfo Lib "user32" Alias "SetMenuItemInfoA" (ByVal hMenu As Long, ByVal un As Long, ByVal bool As Boolean, lpcMenuItemInfo As MENUITEMINFO) As Long
Private Declare Function SetMenuItemInfoStr Lib "user32" Alias "SetMenuItemInfoA" (ByVal hMenu As Long, ByVal un As Long, ByVal bool As Boolean, lpcMenuItemInfo As MENUITEMINFO_STRINGDATA) As Long
Private Declare Function GetMenuItemRect Lib "user32" (ByVal hWnd As Long, ByVal hMenu As Long, ByVal uItem As Long, lprcItem As RECT) As Long
Private Declare Function GetMenuState Lib "user32" (ByVal hMenu As Long, ByVal wID As Long, ByVal wFlags As Long) As Long

Private Declare Function CreateMenu Lib "user32" () As Long
Private Declare Function DestroyMenu Lib "user32" (ByVal hMenu As Long) As Long
Private Declare Function CreatePopupMenu Lib "user32" () As Long

Private Declare Function AppendMenuBylong Lib "user32" Alias "AppendMenuA" (ByVal hMenu As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, ByVal lpNewItem As Long) As Long
Private Declare Function AppendMenuByString Lib "user32" Alias "AppendMenuA" (ByVal hMenu As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, ByVal lpNewItem As String) As Long
Private Declare Function DeleteMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
Private Declare Function ModifyMenu Lib "user32" Alias "ModifyMenuA" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, ByVal lpString As Any) As Long
Private Declare Function ModifyMenuByLong Lib "user32" Alias "ModifyMenuA" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, ByVal lpString As Long) As Long
Private Declare Function RemoveMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
Private Declare Function InsertMenuByLong Lib "user32" Alias "InsertMenuA" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, ByVal lpNewItem As Long) As Long
Private Declare Function InsertMenuByString Lib "user32" Alias "InsertMenuA" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, ByVal lpNewItem As String) As Long
Private Declare Function InsertMenuItem Lib "user32" Alias "InsertMenuItemA" (ByVal hMenu As Long, ByVal un As Long, ByVal bool As Boolean, lpcMenuItemInfo As MENUITEMINFO) As Long

Private Declare Function CheckMenuItem Lib "user32" (ByVal hMenu As Long, ByVal wIDCheckItem As Long, ByVal wCheck As Long) As Long
Private Declare Function CheckMenuRadioItem Lib "user32" (ByVal hMenu As Long, ByVal un1 As Long, ByVal un2 As Long, ByVal un3 As Long, ByVal un4 As Long) As Long
Private Declare Function EnableMenuItem Lib "user32" (ByVal hMenu As Long, ByVal wIDEnableItem As Long, ByVal wEnable As Long) As Long
Private Declare Function HiliteMenuItem Lib "user32" (ByVal hWnd As Long, ByVal hMenu As Long, ByVal wIDHiliteItem As Long, ByVal wHilite As Long) As Long

Private Declare Function MenuItemFromPoint Lib "user32" (ByVal hWnd As Long, ByVal hMenu As Long, ByVal ptScreen As POINTAPI) As Long
Private Declare Function DrawMenuBar Lib "user32" (ByVal hWnd As Long) As Long

Private Declare Function GetSubMenu Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long


' =======================================================================
' GDI private declares:
' =======================================================================

Private Declare Function CreateDCAsNull Lib "gdi32" Alias "CreateDCA" (ByVal lpDriverName As String, lpDeviceName As Any, lpOutput As Any, lpInitData As Any) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal XSrc As Long, ByVal YSrc As Long, ByVal dwRop As Long) As Long
Private Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hdc As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
Private Declare Function SetBkColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
Private Declare Function SetTextColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
Private Declare Function SetBkMode Lib "gdi32" (ByVal hdc As Long, ByVal nBkMode As Long) As Long
Private Declare Function OffsetRect Lib "user32" (lpRect As RECT, ByVal x As Long, ByVal y As Long) As Long
Private Declare Function DrawEdge Lib "user32" (ByVal hdc As Long, qrc As RECT, ByVal edge As Long, ByVal grfFlags As Long) As Long
Private Declare Function OleTranslateColor Lib "OLEPRO32.DLL" (ByVal OLE_COLOR As Long, ByVal HPALETTE As Long, pccolorref As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function DestroyIcon Lib "user32" (ByVal hIcon As Long) As Long
Private Declare Function FillRect Lib "user32" (ByVal hdc As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function DrawState Lib "user32" Alias "DrawStateA" _
   (ByVal hdc As Long, ByVal hBrush As Long, _
   ByVal lpDrawStateProc As Long, _
   ByVal lParam As Long, ByVal wParam As Long, _
   ByVal x As Long, ByVal y As Long, _
   ByVal cX As Long, ByVal cY As Long, _
   ByVal fuFlags As Long) As Long

'/* Image type */
Private Const DST_COMPLEX = &H0
Private Const DST_TEXT = &H1
Private Const DST_PREFIXTEXT = &H2
Private Const DST_ICON = &H3
Private Const DST_BITMAP = &H4

' /* State type */
Private Const DSS_NORMAL = &H0
Private Const DSS_UNION = &H10         ' /* Gray string appearance */
Private Const DSS_DISABLED = &H20
Private Const DSS_MONO = &H80
Private Const DSS_RIGHT = &H8000

Private Const DT_BOTTOM = &H8
Private Const DT_CENTER = &H1
Private Const DT_LEFT = &H0
Private Const DT_CALCRECT = &H400
Private Const DT_WORDBREAK = &H10
Private Const DT_VCENTER = &H4
Private Const DT_TOP = &H0
Private Const DT_TABSTOP = &H80
Private Const DT_SINGLELINE = &H20
Private Const DT_RIGHT = &H2
Private Const DT_NOCLIP = &H100
Private Const DT_INTERNAL = &H1000
Private Const DT_EXTERNALLEADING = &H200
Private Const DT_EXPANDTABS = &H40
Private Const DT_CHARSTREAM = 4
Private Const DT_NOPREFIX = &H800
Private Const DT_EDITCONTROL = &H2000&
Private Const DT_PATH_ELLIPSIS = &H4000&
Private Const DT_END_ELLIPSIS = &H8000&
Private Const DT_MODIFYSTRING = &H10000
Private Const DT_RTLREADING = &H20000
Private Const DT_WORD_ELLIPSIS = &H40000

Private Const OPAQUE = 2
Private Const TRANSPARENT = 1

' DrawEdge:
Private Const BDR_RAISEDOUTER = &H1
Private Const BDR_SUNKENOUTER = &H2
Private Const BDR_RAISEDINNER = &H4
Private Const BDR_SUNKENINNER = &H8

Private Const BDR_OUTER = &H3
Private Const BDR_INNER = &HC
Private Const BDR_RAISED = &H5
Private Const BDR_SUNKEN = &HA

Private Const EDGE_RAISED = (BDR_RAISEDOUTER Or BDR_RAISEDINNER)
Private Const EDGE_SUNKEN = (BDR_SUNKENOUTER Or BDR_SUNKENINNER)
Private Const EDGE_ETCHED = (BDR_SUNKENOUTER Or BDR_RAISEDINNER)
Private Const EDGE_BUMP = (BDR_RAISEDOUTER Or BDR_SUNKENINNER)

Private Const BF_LEFT = &H1
Private Const BF_TOP = &H2
Private Const BF_RIGHT = &H4
Private Const BF_BOTTOM = &H8

Private Const BF_TOPLEFT = (BF_TOP Or BF_LEFT)
Private Const BF_TOPRIGHT = (BF_TOP Or BF_RIGHT)
Private Const BF_BOTTOMLEFT = (BF_BOTTOM Or BF_LEFT)
Private Const BF_BOTTOMRIGHT = (BF_BOTTOM Or BF_RIGHT)
Private Const BF_RECT = (BF_LEFT Or BF_TOP Or BF_RIGHT Or BF_BOTTOM)

Private Const CLR_INVALID = -1

Private Declare Function ImageList_GetIconSize Lib "COMCTL32.DLL" ( _
        ByVal himl As Long, _
        cX As Long, cY As Long _
    ) As Long
Private Declare Function ImageList_Draw Lib "COMCTL32.DLL" ( _
        ByVal himl As Long, ByVal i As Long, _
        ByVal hdcDst As Long, ByVal x As Long, ByVal y As Long, _
        ByVal fStyle As Long _
    ) As Long
' Create a new icon based on an image list icon:
Private Declare Function ImageList_GetIcon Lib "COMCTL32.DLL" ( _
        ByVal himl As Long, _
        ByVal i As Long, _
        ByVal diIgnore As Long _
    ) As Long
Private Const ILD_NORMAL = 0
Private Const ILD_TRANSPARENT = 1
Private Const ILD_BLEND25 = 2
Private Const ILD_SELECTED = 4
Private Const ILD_FOCUS = 4
Private Const ILD_MASK = &H10&
Private Const ILD_IMAGE = &H20&
Private Const ILD_ROP = &H40&
Private Const ILD_OVERLAYMASK = 3840

' =======================================================================
' General Win private declares:
' =======================================================================
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDest As Any, pSrc As Any, ByVal ByteLen As Long)

Private Declare Function PtInRect Lib "user32" (lpRect As RECT, ByVal x As Long, ByVal y As Long) As Long
Private Declare Function MapWindowPoints Lib "user32" (ByVal hwndFrom As Long, ByVal hwndTo As Long, lppt As Any, ByVal cPoints As Long) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function ScreenToClient Lib "user32" (ByVal hWnd As Long, lpPoint As POINTAPI) As Long
Private Declare Function InflateRect Lib "user32" (lpRect As RECT, ByVal x As Long, ByVal y As Long) As Long
Private Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Private Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long

Private Const HWND_DESKTOP = 0


' Window Messages
Private Const WM_DESTROY = &H2
Private Const WM_SETTEXT = &HC
Private Const WM_WININICHANGE = &H1A
Private Const WM_ACTIVATEAPP = &H1C
Private Const WM_CANCELMODE = &H1F
Private Const WM_SETCURSOR = &H20
Private Const WM_MEASUREITEM = &H2C
Private Const WM_DRAWITEM = &H2B
Private Const WM_STYLECHANGING = &H7C
Private Const WM_STYLECHANGED = &H7D
Private Const WM_NCCALCSIZE = &H83
Private Const WM_NCHITTEST = &H84
Private Const WM_NCPAINT = &H85
Private Const WM_NCACTIVATE = &H86
Private Const WM_NCLBUTTONDOWN = &HA1
Private Const WM_NCLBUTTONUP = &HA2
Private Const WM_NCLBUTTONDBLCLK = &HA3
Private Const WM_KEYDOWN = &H100
Private Const WM_COMMAND = &H111
Private Const WM_SYSCOMMAND = &H112
Private Const WM_INITMENUPOPUP = &H117
Private Const WM_MENUSELECT = &H11F
Private Const WM_MENUCHAR = &H120
Private Const WM_MDIGETACTIVE = &H229
Private Const WM_ENTERMENULOOP = &H211
Private Const WM_EXITMENULOOP = &H212


' =======================================================================
' IMPLEMENTATION
' =======================================================================

Public Enum ECPHighlightStyleConstants
   ECPHighlightStyleStandard = 0
   ECPHighlightStyleGradient = 1
   ECPHighlightStyleButton = 2
End Enum

Private m_cMemDC As cMemDC
Private m_cBitmap As cMemDC
Private m_cNCM As cNCMetrics
Private m_cBrush As cDottedBrush

Private m_hWnd As Long

Private m_oActiveMenuColor As OLE_COLOR
Private m_oInActiveMenuColor As OLE_COLOR
Private m_oMenuBackgroundColor As OLE_COLOR
'added dgs 09 May 2001
Private m_oGradientStartColor As OLE_COLOR
Private m_oGradientStopColor As OLE_COLOR

Private m_fnt As StdFont
Private m_fntSymbol As StdFont
Private m_pic As IPicture

Private m_eHighlightStyle As ECPHighlightStyleConstants

Private m_iRestore As Long
Private m_hMenuRestore() As Long
Private m_iMenuPosition() As Long
Private m_tMIIS() As MENUITEMINFO_STRINGDATA
Private m_sCaption() As String
Private m_sShortCut() As String
Private m_sAccelerator() As String
Private m_lMenuTextSize() As Long
Private m_lMenuShortCutSize() As Long

Private m_iHaveSeenCount As Long
Private m_hMenuSeen() As Long

Private m_lMenuItemHeight As Long

Private m_colIcons As New Collection
Private m_hIml As Long
Private m_lIconSize As Long

Implements ISubclass

Public Property Let ImageList( _
        ByRef vImageList As Variant _
    )
    m_hIml = 0
    If (VarType(vImageList) = vbLong) Then
        ' Assume a handle to an image list:
        m_hIml = vImageList
    ElseIf (VarType(vImageList) = vbObject) Then
        ' Assume a VB image list:
        On Error Resume Next
        ' Get the image list initialised..
        vImageList.ListImages(1).Draw 0, 0, 0, 1
        m_hIml = vImageList.hImageList
        If (Err.Number = 0) Then
            ' OK
        Else
            Debug.Print "Failed to Get Image list Handle", "cVGrid.ImageList"
        End If
        On Error GoTo 0
    End If
    If (m_hIml <> 0) Then
        Dim cX As Long, cY As Long
        If (ImageList_GetIconSize(m_hIml, cX, cY) <> 0) Then
            m_lIconSize = cY
        End If
    End If
End Property

Public Property Get IconIndex(ByVal sCaption As String) As Long
   IconIndex = m_colIcons.Item(sCaption)
End Property
Public Property Get IconIndexForIndex(ByVal lItem As Long) As Long
   IconIndexForIndex = m_colIcons.Item(lItem)
End Property
Public Property Let IconIndex(ByVal sCaption As String, ByVal nIndex As Long)
   If nIndex = 0 Then
      m_colIcons.Remove sCaption
   Else
      On Error Resume Next
      m_colIcons.Item sCaption
      If Err.Number = 0 Then
         m_colIcons.Remove sCaption
      End If
      On Error GoTo 0
      m_colIcons.Add nIndex, sCaption
   End If
End Property
Public Sub IconItemCaptionChanged(ByVal sOldCaption As String, ByVal sNewCaption As String)
Dim lIdx As Long
   lIdx = -1
   On Error Resume Next
   lIdx = m_colIcons.Item(sOldCaption)
   If lIdx > -1 Then
      m_colIcons.Remove sOldCaption
      m_colIcons.Add lIdx, sNewCaption
   End If
End Sub
Public Property Get IconItemCount() As Long
   IconItemCount = m_colIcons.Count
End Property
Public Sub ClearIcons()
   Set m_colIcons = New Collection
End Sub
Public Property Get HighlightStyle() As ECPHighlightStyleConstants
   HighlightStyle = m_eHighlightStyle
End Property
Public Property Let HighlightStyle(ByVal eStyle As ECPHighlightStyleConstants)
   m_eHighlightStyle = eStyle
End Property
Private Property Get hFontSymbol() As Long
Dim iFn As IFont
   Set iFn = m_fntSymbol
   hFontSymbol = iFn.hFont
End Property
Friend Property Let Font( _
      fntTHis As StdFont _
   )
   pSetFont fntTHis
End Property
Public Property Set Font( _
      fntTHis As StdFont _
   )
   pSetFont fntTHis
End Property
Public Property Get Font() As StdFont
Dim lHDC As Long
   If m_fnt Is Nothing Then
      lHDC = CreateDCAsNull("DISPLAY", ByVal 0&, ByVal 0&, ByVal 0&)
      Set Font = m_cNCM.Font(lHDC, MenuFont)
      DeleteDC lHDC
   Else
      Set Font = m_fnt
   End If
End Property
Private Sub pSetFont(fntTHis As StdFont)
   Set m_fnt = fntTHis
   m_fntSymbol.Size = Font.Size * 1.2
End Sub
Public Property Get BackgroundPicture() As IPicture
   Set BackgroundPicture = m_pic
End Property
Public Property Let BackgroundPicture(ByRef ipic As IPicture)
   pSetPicture ipic
End Property
Public Property Set BackgroundPicture(ByRef ipic As IPicture)
   pSetPicture ipic
End Property
Private Sub pSetPicture(ByRef ipic As IPicture)
   If Not ipic Is Nothing Then
      Set m_pic = ipic
      Set m_cBitmap = New cMemDC
      m_cBitmap.CreateFromPicture ipic
   Else
      Set m_cBitmap = Nothing
   End If
End Sub
Public Property Let ActiveMenuForeColor(ByVal oColor As OLE_COLOR)
   m_oActiveMenuColor = oColor
End Property
Public Property Let GradientStartColor(ByVal oColor As OLE_COLOR)
   m_oGradientStartColor = oColor
End Property
Public Property Let GradientStopColor(ByVal oColor As OLE_COLOR)
   m_oGradientStopColor = oColor
End Property

Public Property Get ActiveMenuForeColor() As OLE_COLOR
   If m_oActiveMenuColor = CLR_INVALID Then
      ActiveMenuForeColor = vbHighlightText
   Else
      ActiveMenuForeColor = m_oActiveMenuColor
   End If
End Property
Public Property Let InActiveMenuForeColor(ByVal oColor As OLE_COLOR)
   m_oInActiveMenuColor = oColor
End Property
Public Property Get InActiveMenuForeColor() As OLE_COLOR
   If m_oInActiveMenuColor = CLR_INVALID Then
      InActiveMenuForeColor = vbMenuText
   Else
      InActiveMenuForeColor = m_oInActiveMenuColor
   End If
End Property
Public Property Let MenuBackgroundColor(ByVal oColor As OLE_COLOR)
   m_oMenuBackgroundColor = oColor
End Property
Public Property Get MenuBackgroundColor() As OLE_COLOR
   If m_oMenuBackgroundColor = CLR_INVALID Then
      MenuBackgroundColor = vbMenuBar
   Else
      MenuBackgroundColor = m_oMenuBackgroundColor
   End If
End Property
Private Property Get hFont() As Long
Dim iFn As IFont
   Set iFn = Font
   hFont = iFn.hFont
End Property
Public Sub Attach(ByVal lHwnd As Long)
   Detach
   m_hWnd = lHwnd
   AttachMessage Me, m_hWnd, WM_WININICHANGE
   AttachMessage Me, m_hWnd, WM_DRAWITEM
   AttachMessage Me, m_hWnd, WM_MEASUREITEM
   AttachMessage Me, m_hWnd, WM_MENUCHAR
   AttachMessage Me, m_hWnd, WM_INITMENUPOPUP
   AttachMessage Me, m_hWnd, WM_EXITMENULOOP
   AttachMessage Me, m_hWnd, WM_DESTROY
End Sub
Public Sub Detach()
   If Not m_hWnd = 0 Then
      DetachMessage Me, m_hWnd, WM_WININICHANGE
      DetachMessage Me, m_hWnd, WM_DRAWITEM
      DetachMessage Me, m_hWnd, WM_MEASUREITEM
      DetachMessage Me, m_hWnd, WM_MENUCHAR
      DetachMessage Me, m_hWnd, WM_INITMENUPOPUP
      DetachMessage Me, m_hWnd, WM_EXITMENULOOP
      DetachMessage Me, m_hWnd, WM_DESTROY
   End If
End Sub

Private Sub OwnerDrawMenu(ByVal hMenu As Long)
Dim lC As Long
Dim tMIIS As MENUITEMINFO_STRINGDATA
Dim tMII As MENUITEMINFO
Dim iMenu As Long
Dim sCap As String
Dim sShortCut As String
Dim tR As RECT
Dim iPos As Long
Dim lID As Long
Dim bHaveSeen As Boolean
Dim hFntOld As Long
Dim lMenuTextSize As Long
Dim lMenuShortCutSize As Long
Dim i As Long
                  
   ' Set OD flag on the fly...
   bHaveSeen = pbHaveSeen(hMenu)

   hFntOld = SelectObject(m_cMemDC.hdc, hFont)
   lC = GetMenuItemCount(hMenu)
   For iMenu = 0 To lC - 1
      
      If Not bHaveSeen Then

         tMIIS.fMask = MIIM_TYPE Or MIIM_DATA Or MIIM_ID
         tMIIS.cch = 127
         tMIIS.dwTypeData = String$(128, 0)
         tMIIS.cbSize = LenB(tMIIS)
         GetMenuItemInfoStr hMenu, iMenu, True, tMIIS
         'Debug.Print "New Item", tMIIS.dwTypeData

         lID = plAddToRestoreList(hMenu, iMenu, tMIIS)

         If Not (tMIIS.fType And MFT_OWNERDRAW) = MFT_OWNERDRAW Then
            ' Setting this flag causes tMIIS.dwTypeData to be
            ' overwritten with our own app-defined value:
            tMII.fType = tMIIS.fType Or MFT_OWNERDRAW And Not MFT_STRING
            tMII.dwItemData = lID
            tMII.cbSize = LenB(tMII)
            tMII.fMask = MIIM_TYPE Or MIIM_DATA
            SetMenuItemInfo hMenu, iMenu, True, tMII
         End If

      Else

         tMII.fMask = MIIM_TYPE Or MIIM_DATA Or MIIM_ID
         tMII.cbSize = Len(tMIIS)
         ReDim b(0 To 128) As Byte
         tMII.dwTypeData = VarPtr(b(0))
         GetMenuItemInfo hMenu, iMenu, True, tMII
         lID = tMII.dwItemData

         If Not ((tMII.fType And MFT_OWNERDRAW) = MFT_OWNERDRAW) Then

            lID = plReplaceIndex(hMenu, iMenu)

            'Debug.Print "VB has done something to it!", lID
            tMIIS.fMask = MIIM_TYPE Or MIIM_DATA Or MIIM_ID
            tMIIS.cch = 127
            tMIIS.dwTypeData = String$(128, 0)
            tMIIS.cbSize = LenB(tMIIS)
            GetMenuItemInfoStr hMenu, iMenu, True, tMIIS
'
            If lID = 0 Then
               ' New item,just added:
               lID = plAddToRestoreList(hMenu, iMenu, tMIIS)
            Else
               ' replacing existing:
               pReplaceRestoreList lID, hMenu, iMenu, tMIIS
            End If

            ' Setting this flag causes tMIIS.dwTypeData to be
            ' overwritten with our own app-defined value:
            tMII.fType = tMIIS.fType Or MFT_OWNERDRAW And Not MFT_STRING
            tMII.dwItemData = lID
            tMII.cbSize = LenB(tMII)
            tMII.fMask = MIIM_TYPE Or MIIM_DATA
            SetMenuItemInfo hMenu, iMenu, True, tMII

         End If

      End If

      If lID > 0 And lID <= m_iRestore Then
         sCap = m_sCaption(lID)
         sShortCut = m_sShortCut(lID)

         'Debug.Print m_sCaption(lID), m_sShortCut(lID)

         DrawText m_cMemDC.hdc, sCap, -1, tR, DT_LEFT Or DT_SINGLELINE Or DT_CALCRECT
         If tR.Right - tR.left + 1 > lMenuTextSize Then
            lMenuTextSize = tR.Right - tR.left + 1
         End If
         If Len(sShortCut) > 0 Then
            DrawText m_cMemDC.hdc, sShortCut, -1, tR, DT_LEFT Or DT_SINGLELINE Or DT_CALCRECT
            If tR.Right - tR.left + 1 > lMenuShortCutSize Then
               lMenuShortCutSize = tR.Right - tR.left + 1
            End If
         End If
         m_lMenuItemHeight = tR.Bottom - tR.tOp + 2
         If m_lMenuItemHeight < m_lIconSize - 1 Then
            m_lMenuItemHeight = m_lIconSize - 1
         End If

      Else
         'Debug.Print "ERROR! ERROR! ERROR!"
      End If
      
   Next iMenu
   
   For i = 1 To m_iRestore
      If m_hMenuRestore(i) = hMenu Then
         m_lMenuTextSize(i) = lMenuTextSize
         m_lMenuShortCutSize(i) = lMenuShortCutSize
      End If
   Next i
   
   SelectObject m_cMemDC.hdc, hFntOld
   
End Sub
Private Function pbHaveSeen(ByVal hMenu As Long) As Boolean
   
   ' When WM_INITMENUPOPUP fires, this may or not be
   ' a new menu.  We use an array to store which menus
   ' we've already worked on:

Dim i As Long
   
   For i = 1 To m_iHaveSeenCount
      If hMenu = m_hMenuSeen(i) Then
         pbHaveSeen = True
         Exit Function
      End If
   Next i
   m_iHaveSeenCount = m_iHaveSeenCount + 1
   ReDim Preserve m_hMenuSeen(1 To m_iHaveSeenCount) As Long
   m_hMenuSeen(m_iHaveSeenCount) = hMenu

End Function
Private Function plReplaceIndex(ByVal hMenu As Long, ByVal iMenu As Long)
Dim i As Long
   For i = 1 To m_iRestore
      If m_hMenuRestore(i) = hMenu Then
         If m_iMenuPosition(i) = iMenu Then
            plReplaceIndex = i
            Exit Function
         End If
      End If
   Next i
End Function
Private Function plAddToRestoreList(ByVal hMenu As Long, ByVal iMenu As Long, tMIIS As MENUITEMINFO_STRINGDATA) As Long
   
   ' Here we store information about a menu item.  When the
   ' menus are closed again we can reset things back to the
   ' way they were using this struct.

   m_iRestore = m_iRestore + 1
   ReDim Preserve m_hMenuRestore(1 To m_iRestore) As Long
   ReDim Preserve m_iMenuPosition(1 To m_iRestore) As Long
   ReDim Preserve m_tMIIS(1 To m_iRestore) As MENUITEMINFO_STRINGDATA
   ReDim Preserve m_sCaption(1 To m_iRestore) As String
   ReDim Preserve m_sShortCut(1 To m_iRestore) As String
   ReDim Preserve m_sAccelerator(1 To m_iRestore) As String
   ReDim Preserve m_lMenuTextSize(1 To m_iRestore) As Long
   ReDim Preserve m_lMenuShortCutSize(1 To m_iRestore) As Long
   pReplaceRestoreList m_iRestore, hMenu, iMenu, tMIIS
   plAddToRestoreList = m_iRestore

End Function
Private Sub pReplaceRestoreList(ByVal lIdx As Long, hMenu As Long, iMenu As Long, tMIIS As MENUITEMINFO_STRINGDATA)
Dim sCap As String
Dim sShortCut As String
Dim iPos As Long

   m_hMenuRestore(lIdx) = hMenu
   m_iMenuPosition(lIdx) = iMenu
   LSet m_tMIIS(lIdx) = tMIIS
   If tMIIS.cch > 0 Then
      sCap = left$(tMIIS.dwTypeData, tMIIS.cch)
   Else
      sCap = ""
   End If
   iPos = InStr(sCap, vbTab)
   If iPos > 0 Then
      m_sShortCut(lIdx) = Mid$(sCap, iPos + 1)
      m_sCaption(lIdx) = left$(sCap, iPos - 1)
   Else
      m_sCaption(lIdx) = sCap
      m_sShortCut(lIdx) = ""
   End If
   iPos = InStr(m_sCaption(lIdx), "&")
   If iPos > 0 And iPos < Len(m_sCaption(lIdx)) Then
      m_sAccelerator(lIdx) = UCase$(Mid$(m_sCaption(lIdx), iPos + 1, 1))
   End If
End Sub
Private Function InternalIDForWindowsID(ByVal wID As Long) As Long
Dim i As Long
   ' linear search I'm afraid, but it is only called once
   ' per menu item shown (when WM_MEASUREITEM is fired)
   For i = 1 To m_iRestore
      If m_tMIIS(i).wID = wID Then
         InternalIDForWindowsID = i
         Exit Function
      End If
   Next i
End Function
Private Sub pRestoreList()
Dim i As Long
   'Debug.Print "RESTORELIST"
   ' erase the lot:
   For i = 1 To m_iRestore
      SetMenuItemInfoStr m_hMenuRestore(i), m_iMenuPosition(i), True, m_tMIIS(i)
   Next i
   m_iRestore = 0
   Erase m_hMenuRestore
   Erase m_iMenuPosition
   Erase m_tMIIS
   Erase m_sCaption()
   Erase m_sShortCut()
   Erase m_sAccelerator()
   m_iHaveSeenCount = 0
   Erase m_hMenuSeen()
End Sub

Private Sub Class_Initialize()
   Set m_cNCM = New cNCMetrics
   m_cNCM.GetMetrics
   Set m_cMemDC = New cMemDC
   m_cMemDC.Width = Screen.Width \ Screen.TwipsPerPixelY
   m_cMemDC.Height = 24
   m_oActiveMenuColor = CLR_INVALID
   m_oInActiveMenuColor = CLR_INVALID
   m_oMenuBackgroundColor = CLR_INVALID
   m_oGradientStartColor = CLR_INVALID
   m_oGradientStopColor = CLR_INVALID
   Set m_fntSymbol = New StdFont
   m_fntSymbol.Name = "Marlett"
   m_fntSymbol.Size = Font.Size * 1.2
   Set m_cBrush = New cDottedBrush
   m_cBrush.Create
End Sub

Private Sub Class_Terminate()
   Set m_cMemDC = Nothing
End Sub

Private Property Let ISubclass_MsgResponse(ByVal RHS As EMsgResponse)
   '
End Property

Private Property Get ISubclass_MsgResponse() As EMsgResponse
   If _
      CurrentMessage = WM_DESTROY Or _
      CurrentMessage = WM_INITMENUPOPUP Or _
      CurrentMessage = WM_WININICHANGE Or _
      CurrentMessage = WM_EXITMENULOOP _
   Then
      ISubclass_MsgResponse = emrPreprocess
   Else
      ISubclass_MsgResponse = emrConsume
   End If
End Property

Private Function ISubclass_WindowProc(ByVal hWnd As Long, ByVal iMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Dim iMenu As Long
Dim iLastDownOn As Long
Dim iLastOver As Long
Dim lR As Long
Dim lFlag As Long
Dim hMenu As Long
Dim iChar As Long

   Select Case iMsg
   Case WM_INITMENUPOPUP
      If (lParam And &HFFFF0000) = 0 Then
         OwnerDrawMenu wParam
      End If
      
   Case WM_MEASUREITEM
      ISubclass_WindowProc = MeasureItem(wParam, lParam)
   
   Case WM_DRAWITEM
      DrawItem wParam, lParam
      
   Case WM_MENUCHAR
      ' Check that this is my menu:
      lFlag = wParam \ &H10000
      If ((lFlag And MF_SYSMENU) <> MF_SYSMENU) Then
         hMenu = lParam
         iChar = (wParam And &HFFFF&)
         ' See if this corresponds to an accelerator on the menu:
         lR = ParseMenuChar(hMenu, iChar)
         If lR > 0 Then
            ISubclass_WindowProc = lR
            Exit Function
         End If
      End If
      ISubclass_WindowProc = CallOldWindowProc(m_hWnd, WM_MENUCHAR, wParam, lParam)
   
   Case WM_WININICHANGE
      m_cNCM.GetMetrics
   
   Case WM_EXITMENULOOP
      pRestoreList
   
   Case WM_DESTROY
      Detach
   
   End Select
   
End Function
Private Function ParseMenuChar( _
        ByVal hMenu As Long, _
        ByVal iChar As Integer _
    ) As Long
Dim sChar As String
Dim l As Long
Dim lH() As Long
Dim sItems() As String
    
    sChar = UCase$(Chr$(iChar))
    For l = 1 To m_iRestore
        If (m_hMenuRestore(l) = hMenu) Then
            If (m_sAccelerator(l) = sChar) Then
               ParseMenuChar = &H20000 Or m_iMenuPosition(l)
               ' Debug.Print "Found Menu Char"
               Exit Function
            End If
        End If
    Next l
    
End Function

Private Function MeasureItem(ByVal wParam As Long, ByVal lParam As Long) As Long
Dim tMIS As MEASUREITEMSTRUCT
Dim lID As Long
   
   CopyMemory tMIS, ByVal lParam, LenB(tMIS)
   If tMIS.CtlType = ODT_MENU Then
                  
      ' because we don't get the popup menu handle
      ' in the tMIS structure, we have to do an internal
      ' lookup to find info about this menu item.
      ' poor implementation of MEASUREITEMSTRUCT - it
      ' should have a .hWndItem field like DRAWITEMSTRUCT
      ' - spm
      lID = InternalIDForWindowsID(tMIS.itemID)
            
      ' Width:
      tMIS.itemWidth = 4 + 22 + m_lMenuTextSize(lID) + 4
      If m_lMenuShortCutSize(lID) > 0 Then
         tMIS.itemWidth = tMIS.itemWidth + 4 + m_lMenuShortCutSize(lID) + 4
      End If
      
      ' Height:
      If lID > 0 And lID <= m_iRestore Then
         If (m_tMIIS(lID).fType And MFT_SEPARATOR) = MFT_SEPARATOR Then
            tMIS.itemHeight = 8
         Else
            ' menu item height is always the same
            tMIS.itemHeight = m_lMenuItemHeight + 6
         End If
      Else
         ' problem.
      End If
      
      CopyMemory ByVal lParam, tMIS, LenB(tMIS)
      
   Else
      MeasureItem = CallOldWindowProc(m_hWnd, WM_MEASUREITEM, wParam, lParam)
   End If
End Function
Private Function DrawItem(ByVal wParam As Long, ByVal lParam As Long) As Long
Dim tDIS As DRAWITEMSTRUCT
Dim hBr As Long
Dim tR As RECT, tTR As RECT, tWR As RECT
Dim lHDC As Long
Dim hFntOld As Long
Dim hFntsOld As Long
Dim tMII As MENUITEMINFO
Dim bRadioCheck As Boolean, bDisabled As Boolean, bChecked As Boolean, bHighlighted As Boolean
Dim lID As Long
Dim lSelLeft As Long
Dim sCC As String
Dim lIconIndex As Long
Dim lX As Long, lY As Long

   CopyMemory tDIS, ByVal lParam, LenB(tDIS)
   If tDIS.CtlType = ODT_MENU Then
      ' tDIS.hWndItem is the menu containing the item, tDIS.itemID is the wID
      
      m_cMemDC.Width = tDIS.rcItem.Right - tDIS.rcItem.left + 1
      m_cMemDC.Height = tDIS.rcItem.Bottom - tDIS.rcItem.tOp + 1
      lHDC = m_cMemDC.hdc
      hFntOld = SelectObject(lHDC, hFont)
      
      LSet tR = tDIS.rcItem
      OffsetRect tR, -tR.left, -tR.tOp
            
      ' Fill background:
      tTR.Right = m_cMemDC.Width
      tTR.Bottom = m_cMemDC.Height
      
      If m_cBitmap Is Nothing Then
         hBr = CreateSolidBrush(TranslateColor(MenuBackgroundColor))
         FillRect lHDC, tTR, hBr
         DeleteObject hBr
      Else
         TileArea lHDC, tR.left, tR.tOp, tR.Right - tR.left + 1, tR.Bottom - tR.tOp + 1, m_cBitmap.hdc, m_cBitmap.Width, m_cBitmap.Height, tDIS.rcItem.tOp
      End If
      
      tR.tOp = tR.tOp + 1
            
      SetBkMode lHDC, TRANSPARENT
      
      ' Draw the text:
      tMII.cbSize = LenB(tMII)
      tMII.fMask = MIIM_TYPE Or MIIM_STATE Or MIIM_DATA
      ReDim b(0 To 128) As Byte
      tMII.dwTypeData = VarPtr(b(0))
      GetMenuItemInfo tDIS.hwndItem, tDIS.itemID, False, tMII
      
      If (tMII.fType And MFT_SEPARATOR) = MFT_SEPARATOR Then
         ' Separator:
         LSet tWR = tR
         tWR.tOp = (tWR.Bottom - tWR.tOp - 2) \ 2 + tWR.tOp
         tWR.Bottom = tWR.tOp + 2
         InflateRect tWR, -12, 0
         DrawEdge lHDC, tWR, BDR_SUNKENOUTER, BF_TOP Or BF_BOTTOM
      Else
         ' Text item:
         bRadioCheck = ((tMII.fType And MFT_RADIOCHECK) = MFT_RADIOCHECK)
         bDisabled = ((tMII.fState And MFS_DISABLED) = MFS_DISABLED)
         bChecked = ((tMII.fState And MFS_CHECKED) = MFS_CHECKED)
         bHighlighted = ((tMII.fState And MFS_HILITE) = MFS_HILITE)
         
         lID = tMII.dwItemData

         ' Icon?
         lIconIndex = -1
         On Error Resume Next
         lIconIndex = m_colIcons.Item(m_sCaption(lID))
         On Error GoTo 0
         
         If bChecked Or lIconIndex > -1 Then
            lSelLeft = 4 + (tR.Bottom - tR.tOp + 1 - 4)
         End If
                           
         If bHighlighted And Not bDisabled Then
            If m_eHighlightStyle = ECPHighlightStyleGradient Then
               ' Draw a gradient:
               '-----------added dgs 09 May 2001
               If m_oGradientStartColor = CLR_INVALID Then
                  m_oGradientStartColor = vbHighlight
               End If
               If m_oGradientStopColor = CLR_INVALID Then
                  m_oGradientStopColor = MenuBackgroundColor
               End If
               '-----------
               LSet tWR = tR
               tWR.left = lSelLeft
               tWR.Right = tWR.left + 4 + (tR.Bottom - tR.tOp + 1 - 4)
               hBr = CreateSolidBrush(TranslateColor(m_oGradientStartColor))
               FillRect lHDC, tWR, hBr
               DeleteObject hBr
               LSet tWR = tR
               tWR.left = tWR.left + 4 + (tR.Bottom - tR.tOp + 1 - 4)
               DrawGradient lHDC, tWR, TranslateColor(m_oGradientStartColor), TranslateColor(m_oGradientStopColor), False
            ElseIf m_eHighlightStyle = ECPHighlightStyleButton Then
               ' do nothing now
            Else
               ' standard:
               LSet tWR = tR
               tWR.left = lSelLeft
               hBr = CreateSolidBrush(TranslateColor(vbHighlight))
               FillRect lHDC, tWR, hBr
               DeleteObject hBr
            End If
         End If
         
         If bDisabled Then
            SetTextColor lHDC, TranslateColor(vb3DHighlight)
         Else
            If bHighlighted Then
               SetTextColor lHDC, TranslateColor(ActiveMenuForeColor)
            Else
               SetTextColor lHDC, TranslateColor(InActiveMenuForeColor)
            End If
         End If
         
         ' Get the check/icon space:
         LSet tWR = tR
         If m_eHighlightStyle = ECPHighlightStyleButton Then
            InflateRect tWR, -2, -2
         Else
            tWR.left = tWR.left + 1
         End If
         tWR.Right = tWR.left + (tWR.Bottom - tWR.tOp + 1 - 2)
         
         ' Check:
         If bChecked Then
                        
            ' Colour in:
            If Not bHighlighted Then
               SetBkMode lHDC, OPAQUE
               m_cBrush.Rectangle lHDC, tWR.left, tWR.tOp, tWR.Right - tWR.left, tWR.Bottom - tWR.tOp, 1, PATCOPY, True, MenuBackgroundColor, vb3DHighlight
               SetBkMode lHDC, TRANSPARENT
                  
               If bDisabled Then
                  SetTextColor lHDC, TranslateColor(vb3DHighlight)
               Else
                  If bHighlighted Then
                     SetTextColor lHDC, TranslateColor(ActiveMenuForeColor)
                  Else
                     SetTextColor lHDC, TranslateColor(InActiveMenuForeColor)
                  End If
               End If
            End If
            
            DrawEdge lHDC, tWR, BDR_SUNKENOUTER, BF_RECT
                        
            If lIconIndex = -1 Then
               ' Draw the appropriate symbol:
               SelectObject lHDC, hFntOld
               hFntsOld = SelectObject(lHDC, hFontSymbol)
               If bHighlighted Then
                  SetTextColor lHDC, TranslateColor(InActiveMenuForeColor)
               End If
               If bRadioCheck Then
                  pDrawText lHDC, "h", tWR, DT_VCENTER Or DT_CENTER Or DT_SINGLELINE, bDisabled
               Else
                  pDrawText lHDC, "b", tWR, DT_VCENTER Or DT_CENTER Or DT_SINGLELINE, bDisabled
               End If
               SelectObject lHDC, hFntsOld
               hFntOld = SelectObject(lHDC, hFont)
            Else
               lX = tWR.left + (tWR.Right - tWR.left + 1 - m_lIconSize) \ 2
               lY = tWR.tOp + (tWR.Bottom - tWR.tOp + 1 - m_lIconSize) \ 2
               If bDisabled Then
                  ImageListDrawIconDisabled lHDC, m_hIml, lIconIndex, lX, lY, m_lIconSize
               Else
                  ImageListDrawIcon lHDC, m_hIml, lIconIndex, lX, lY
               End If
            End If
         Else
            If lIconIndex > -1 Then
               If bHighlighted And Not bDisabled Then
                  If m_eHighlightStyle <> ECPHighlightStyleButton Then
                     DrawEdge lHDC, tWR, BDR_RAISEDINNER, BF_RECT
                  End If
               End If
               lX = tWR.left + (tWR.Right - tWR.left + 1 - m_lIconSize) \ 2
               lY = tWR.tOp + (tWR.Bottom - tWR.tOp + 1 - m_lIconSize) \ 2
               lX = lX + 2 * Abs(m_eHighlightStyle = ECPHighlightStyleButton)
               If bDisabled Then
                  ImageListDrawIconDisabled lHDC, m_hIml, lIconIndex, lX, lY, m_lIconSize
               Else
                  ImageListDrawIcon lHDC, m_hIml, lIconIndex, lX, lY
               End If
            End If

         End If
         
                  
         ' Draw text:
         
         LSet tWR = tR
         tWR.left = 4 + (tR.Bottom - tR.tOp + 1 - 4) + 2
         If lID > 0 And lID <= m_iRestore Then
            pDrawText lHDC, m_sCaption(lID), tWR, DT_LEFT Or DT_SINGLELINE Or DT_VCENTER, bDisabled
            If Len(m_sShortCut(lID)) > 0 Then
               tWR.left = tWR.left + m_lMenuTextSize(lID) + 4 + 4
               pDrawText lHDC, m_sShortCut(lID), tWR, DT_LEFT Or DT_SINGLELINE Or DT_VCENTER, bDisabled
            End If
         End If
         
         ' Highlighted:
         If bHighlighted And m_eHighlightStyle = ECPHighlightStyleButton And Not (bDisabled) Then
            LSet tWR = tR
            InflateRect tWR, -2, 0
            DrawEdge lHDC, tWR, BDR_RAISEDINNER, BF_RECT
         End If
         
      End If
      
      SelectObject lHDC, hFntOld
      
      BitBlt tDIS.hdc, tDIS.rcItem.left, tDIS.rcItem.tOp, tDIS.rcItem.Right - tDIS.rcItem.left + 1, tDIS.rcItem.Bottom - tDIS.rcItem.tOp + 1, lHDC, 0, 0, vbSrcCopy
      
   Else
      DrawItem = CallOldWindowProc(m_hWnd, WM_DRAWITEM, wParam, lParam)
   End If
End Function
Private Function pDrawText(ByVal lHDC As Long, ByVal sText As String, tR As RECT, ByVal dtFlags As Long, ByVal bDisabled As Boolean)
Dim tWR As RECT
   LSet tWR = tR
   If bDisabled Then
      SetTextColor lHDC, TranslateColor(vb3DHighlight)
      OffsetRect tWR, 1, 1
   End If
   DrawText lHDC, sText, -1, tWR, dtFlags
   If bDisabled Then
      SetTextColor lHDC, TranslateColor(vbButtonShadow)
      OffsetRect tWR, -1, -1
      DrawText lHDC, sText, -1, tWR, dtFlags
   End If
End Function
Private Sub DrawGradient( _
      ByVal hdc As Long, _
      ByRef rct As RECT, _
      ByVal lEndColour As Long, _
      ByVal lStartColour As Long, _
      ByVal bVertical As Boolean _
   )
Dim lStep As Long
Dim lPos As Long, lSize As Long
Dim bRGB(1 To 3) As Integer
Dim bRGBStart(1 To 3) As Integer
Dim dR(1 To 3) As Double
Dim dPos As Double, d As Double
Dim hBr As Long
Dim tR As RECT
   
   LSet tR = rct
   If bVertical Then
      lSize = (tR.Bottom - tR.tOp)
   Else
      lSize = (tR.Right - tR.left)
   End If
   lStep = lSize \ 255
   If (lStep < 3) Then
       lStep = 3
   End If
       
   bRGB(1) = lStartColour And &HFF&
   bRGB(2) = (lStartColour And &HFF00&) \ &H100&
   bRGB(3) = (lStartColour And &HFF0000) \ &H10000
   bRGBStart(1) = bRGB(1): bRGBStart(2) = bRGB(2): bRGBStart(3) = bRGB(3)
   dR(1) = (lEndColour And &HFF&) - bRGB(1)
   dR(2) = ((lEndColour And &HFF00&) \ &H100&) - bRGB(2)
   dR(3) = ((lEndColour And &HFF0000) \ &H10000) - bRGB(3)
        
   For lPos = lSize To 0 Step -lStep
      ' Draw bar:
      If bVertical Then
         tR.tOp = tR.Bottom - lStep
      Else
         tR.left = tR.Right - lStep
      End If
      If tR.tOp < rct.tOp Then
         tR.tOp = rct.tOp
      End If
      If tR.left < rct.left Then
         tR.left = rct.left
      End If
      
      'Debug.Print tR.Right, tR.left, (bRGB(3) * &H10000 + bRGB(2) * &H100& + bRGB(1))
      hBr = CreateSolidBrush((bRGB(3) * &H10000 + bRGB(2) * &H100& + bRGB(1)))
      FillRect hdc, tR, hBr
      DeleteObject hBr
            
      ' Adjust colour:
      dPos = ((lSize - lPos) / lSize)
      If bVertical Then
         tR.Bottom = tR.tOp
         bRGB(1) = bRGBStart(1) + dR(1) * dPos
         bRGB(2) = bRGBStart(2) + dR(2) * dPos
         bRGB(3) = bRGBStart(3) + dR(3) * dPos
      Else
         tR.Right = tR.left
         bRGB(1) = bRGBStart(1) + dR(1) * dPos
         bRGB(2) = bRGBStart(2) + dR(2) * dPos
         bRGB(3) = bRGBStart(3) + dR(3) * dPos
      End If
      
   Next lPos

End Sub

Private Sub TileArea( _
        ByVal hdcTo As Long, _
        ByVal x As Long, _
        ByVal y As Long, _
        ByVal Width As Long, _
        ByVal Height As Long, _
        ByVal hDcSrc As Long, _
        ByVal SrcWidth As Long, _
        ByVal SrcHeight As Long, _
        ByVal lOffsetY As Long _
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

    lSrcStartX = (x Mod SrcWidth)
    lSrcStartY = ((y + lOffsetY) Mod SrcHeight)
    lSrcStartWidth = (SrcWidth - lSrcStartX)
    lSrcStartHeight = (SrcHeight - lSrcStartY)
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
            BitBlt hdcTo, lDstX, lDstY, lDstWidth, lDstHeight, hDcSrc, lSrcX, lSrcY, vbSrcCopy
            lDstX = lDstX + lDstWidth
            lSrcX = 0
            lDstWidth = SrcWidth
        Loop
        lDstY = lDstY + lDstHeight
        lSrcY = 0
        lDstHeight = SrcHeight
    Loop
End Sub

' Convert Automation color to Windows color
Private Function TranslateColor(ByVal clr As OLE_COLOR, _
                        Optional hPal As Long = 0) As Long
    If OleTranslateColor(clr, hPal, TranslateColor) Then
        TranslateColor = CLR_INVALID
    End If
End Function

Private Sub ImageListDrawIcon( _
        ByVal hdc As Long, _
        ByVal himl As Long, _
        ByVal iIconIndex As Long, _
        ByVal lX As Long, _
        ByVal lY As Long, _
        Optional ByVal bSelected As Boolean = False, _
        Optional ByVal bBlend25 As Boolean = False _
    )
Dim lFlags As Long
Dim lR As Long

    lFlags = ILD_TRANSPARENT
    If (bSelected) Then
        lFlags = lFlags Or ILD_SELECTED
    End If
    If (bBlend25) Then
        lFlags = lFlags Or ILD_BLEND25
    End If
    lR = ImageList_Draw( _
            himl, _
            iIconIndex, _
            hdc, _
            lX, _
            lY, _
            lFlags)
    If (lR = 0) Then
        Debug.Print "Failed to draw Image: " & iIconIndex & " onto hDC " & hdc, "ImageListDrawIcon"
    End If
End Sub
Private Sub ImageListDrawIconDisabled( _
        ByVal hdc As Long, _
        ByVal himl As Long, _
        ByVal iIconIndex As Long, _
        ByVal lX As Long, _
        ByVal lY As Long, _
        ByVal lSize As Long _
    )
Dim lR As Long
Dim hIcon As Long

   hIcon = ImageList_GetIcon(himl, iIconIndex, 0)
   lR = DrawState(hdc, 0, 0, hIcon, 0, lX, lY, lSize, lSize, DST_ICON Or DSS_DISABLED)
   DestroyIcon hIcon
   
End Sub





