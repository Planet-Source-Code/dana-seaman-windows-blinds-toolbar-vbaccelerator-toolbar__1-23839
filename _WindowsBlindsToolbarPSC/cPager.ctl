VERSION 5.00
Begin VB.UserControl cPager 
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ControlContainer=   -1  'True
   ScaleHeight     =   3600
   ScaleWidth      =   4800
   ToolboxBitmap   =   "cPager.ctx":0000
End
Attribute VB_Name = "cPager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ======================================================================
' cPager Control
' Steve McMahon
' Date: 28 April 1998
'
' A complete implementation of the Pager control as supplied with
' COMCTL32 v4.72 and higher.
' Requires SSUBTMR.DLL to run.
' ======================================================================


' ======================================================================
' Declares and types:
' ======================================================================
' Windows general:
Private Const WM_USER = &H400
Private Const WM_NOTIFY = &H4E
Private Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function SendMessageStr Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function CreateWindowEX Lib "user32" Alias "CreateWindowExA" (ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, lpParam As Any) As Long
Private Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Private Declare Function MoveWindow Lib "user32" (ByVal hwnd As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
Private Declare Function DestroyWindow Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Private Declare Function GetParent Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDest As Any, pSrc As Any, ByVal ByteLen As Long)
Private Declare Function IsWindowVisible Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function IsWindowEnabled Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function SetFocus Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function GetFocus Lib "user32" () As Long
Private Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Private Declare Function GetWindow Lib "user32" (ByVal hwnd As Long, ByVal wCmd As Long) As Long
' Window style bit functions:
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" _
    (ByVal hwnd As Long, ByVal nIndex As Long, _
    ByVal dwNewLong As Long _
    ) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" _
    (ByVal hwnd As Long, ByVal nIndex As Long _
    ) As Long
' Window Long indexes:
Private Enum EWindowLongIndexes
    GWL_EXSTYLE = (-20)
    GWL_HINSTANCE = (-6)
    GWL_HWNDPARENT = (-8)
    GWL_ID = (-12)
    GWL_STYLE = (-16)
    GWL_USERDATA = (-21)
    GWL_WNDPROC = (-4)
End Enum
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cX As Long, ByVal cY As Long, ByVal wFlags As Long) As Long
Private Enum ESetWindowPosStyles
    SWP_SHOWWINDOW = &H40
    SWP_HIDEWINDOW = &H80
    SWP_FRAMECHANGED = &H20        '  The frame changed: send WM_NCCALCSIZE
    SWP_NOACTIVATE = &H10
    SWP_NOCOPYBITS = &H100
    SWP_NOMOVE = &H2
    SWP_NOOWNERZORDER = &H200      '  Don't do owner Z ordering
    SWP_NOREDRAW = &H8
    SWP_NOREPOSITION = SWP_NOOWNERZORDER
    SWP_NOSIZE = &H1
    SWP_NOZORDER = &H4
    SWP_DRAWFRAME = SWP_FRAMECHANGED
    HWND_NOTOPMOST = -2
End Enum
Private Declare Function OleTranslateColor Lib "OLEPRO32.DLL" (ByVal OLE_COLOR As Long, ByVal HPALETTE As Long, pccolorref As Long) As Long
Private Const CLR_INVALID = -1
Private Const SW_HIDE = 0
Private Const GW_CHILD = 5
Private Const WS_CHILD = &H40000000
Private Const WS_VISIBLE = &H10000000
'Private Const WS_CLIPCHILDREN = &H2000000
'Private Const WS_CLIPSIBLINGS = &H4000000
Private Const WS_EX_CLIENTEDGE = &H200
'Private Const WS_BORDER = &H800000
'Private Const WM_HSCROLL = &H114
'Private Const WM_VSCROLL = &H115
'Private Const WM_GETTEXT = &HD
Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

' Common controls general:
Private Type tagInitCommonControlsEx
    lngSize As Long
    lngICC As Long
End Type
Private Declare Function InitCommonControlsEx Lib "COMCTL32.DLL" (iccex As tagInitCommonControlsEx) As Boolean

Private Const H_MAX As Long = &HFFFF + 1
Private Const PGN_FIRST = (H_MAX - 900)                  '// Pager Control
Private Const PGN_LAST = (H_MAX - 950)
Private Const PGM_FIRST = &H1400                  '// Pager control messages
Private Const CCM_FIRST = &H2000                   '// Common control shared messages
Private Const CCM_GETDROPTARGET = (CCM_FIRST + 4)
Private Type NMHDR
    hwndFrom As Long
    idfrom As Long
    code As Long
End Type
Private Const ICC_PAGESCROLLER_CLASS = &H1000&      '// page scroller

' Some toolbar stuff:
Private Const TB_BUTTONCOUNT = (WM_USER + 24)
Private Const TB_SETEXTENDEDSTYLE = (WM_USER + 84)    ' // For TBSTYLE_EX_*
Private Const TBSTYLE_FLAT = &H800
Private Const TB_GETITEMRECT = (WM_USER + 29)

' ///  ====================== Pager Control =============================
' //---------------------------------------------------------------------------------------
' //---------------------------------------------------------------------------------------

' //Pager Class Name
Private Const WC_PAGESCROLLERW = "SysPager"
Private Const WC_PAGESCROLLERA = "SysPager"

#If UNICODE Then
Private Const WC_PAGESCROLLER = WC_PAGESCROLLERW
#Else
Private Const WC_PAGESCROLLER = WC_PAGESCROLLERA
#End If


' //---------------------------------------------------------------------------------------
' // Pager Control Styles
' //---------------------------------------------------------------------------------------
Private Const PGS_VERT = &H0
Private Const PGS_HORZ = &H1
Private Const PGS_AUTOSCROLL = &H2
Private Const PGS_DRAGNDROP = &H4


' //---------------------------------------------------------------------------------------
' // Pager Button State
' //---------------------------------------------------------------------------------------
' //The scroll can be in one of the following control State
Private Const PGF_INVISIBLE = 0        ' // Scroll button is not visible
Private Const PGF_NORMAL = 1           ' // Scroll button is in normal state
Private Const PGF_GRAYED = 2           ' // Scroll button is in grayed state
Private Const PGF_DEPRESSED = 4        ' // Scroll button is in depressed state
Private Const PGF_HOT = 8              ' // Scroll button is in hot state


' // The following identifiers specifies the button control
Private Const PGB_TOPORLEFT = 0
Private Const PGB_BOTTOMORRIGHT = 1

' //---------------------------------------------------------------------------------------
' // Pager Control  Messages
' //---------------------------------------------------------------------------------------
Private Const PGM_SETCHILD = (PGM_FIRST + 1)            ' // lParam == hwnd
'private const Pager_SetChild(hwnd, hwndChild) \
'        (void)SNDMSG((hwnd), PGM_SETCHILD, 0, (LPARAM)(hwndChild))

Private Const PGM_RECALCSIZE = (PGM_FIRST + 2)
'private const Pager_RecalcSize(hwnd) \
'        (void)SNDMSG((hwnd), PGM_RECALCSIZE, 0, 0)

'Private Const PGM_FORWARDMOUSE = (PGM_FIRST + 3)
'private const Pager_ForwardMouse(hwnd, bForward) \
'        (void)SNDMSG((hwnd), PGM_FORWARDMOUSE, (WPARAM)(bForward), 0)

Private Const PGM_SETBKCOLOR = (PGM_FIRST + 4)
'private const Pager_SetBkColor(hwnd, clr) \
'        (COLORREF)SNDMSG((hwnd), PGM_SETBKCOLOR, 0, (LPARAM)clr)

'Private Const PGM_GETBKCOLOR = (PGM_FIRST + 5)
'private const Pager_GetBkColor(hwnd) \
'        (COLORREF)SNDMSG((hwnd), PGM_GETBKCOLOR, 0, 0)

Private Const PGM_SETBORDER = (PGM_FIRST + 6)
'private const Pager_SetBorder(hwnd, iBorder) \
'        (int)SNDMSG((hwnd), PGM_SETBORDER, 0, (LPARAM)iBorder)

Private Const PGM_GETBORDER = (PGM_FIRST + 7)
'private const Pager_GetBorder(hwnd) \
'        (int)SNDMSG((hwnd), PGM_GETBORDER, 0, 0)

Private Const PGM_SETPOS = (PGM_FIRST + 8)
'private const Pager_SetPos(hwnd, iPos) \
'        (int)SNDMSG((hwnd), PGM_SETPOS, 0, (LPARAM)iPos)

Private Const PGM_GETPOS = (PGM_FIRST + 9)
'private const Pager_GetPos(hwnd) \
'        (int)SNDMSG((hwnd), PGM_GETPOS, 0, 0)

Private Const PGM_SETBUTTONSIZE = (PGM_FIRST + 10)
'private const Pager_SetButtonSize(hwnd, iSize) \
'        (int)SNDMSG((hwnd), PGM_SETBUTTONSIZE, 0, (LPARAM)iSize)

Private Const PGM_GETBUTTONSIZE = (PGM_FIRST + 11)
'private const Pager_GetButtonSize(hwnd) \
'        (int)SNDMSG((hwnd), PGM_GETBUTTONSIZE, 0,0)

Private Const PGM_GETBUTTONSTATE = (PGM_FIRST + 12)
'private const Pager_GetButtonState(hwnd, iButton) \
'        (DWORD)SNDMSG((hwnd), PGM_GETBUTTONSTATE, 0, (LPARAM)iButton)

'Private Const PGM_GETDROPTARGET = CCM_GETDROPTARGET
'private const Pager_GetDropTarget(hwnd, ppdt) \
'        (void)SNDMSG((hwnd), PGM_GETDROPTARGET, 0, (LPARAM)ppdt)

' //---------------------------------------------------------------------------------------
' //Pager Control Notification Messages
' //---------------------------------------------------------------------------------------


' // PGN_SCROLL Notification Message

Private Const PGN_SCROLL = (PGN_FIRST - 1)

Private Const PGF_SCROLLUP = 1
Private Const PGF_SCROLLDOWN = 2
Private Const PGF_SCROLLLEFT = 4
Private Const PGF_SCROLLRIGHT = 8


' //Keys down
'Private Const PGK_SHIFT = 1
'Private Const PGK_CONTROL = 2
'Private Const PGK_MENU = 4


' // This structure is sent along with PGN_SCROLL notifications
Private Type NMPGSCROLL
    hdr As NMHDR
    fwKeys As Integer     ' // Specifies which keys are down when this notification is send
    rcParent As RECT          ' // Contains Parent Window Rect
    iDir As Long              ' // Scrolling Direction
    iXpos As Long             ' // Horizontal scroll position
    iYpos As Long             ' // Vertical scroll position
    iScroll As Long           ' // [in/out] Amount to scroll
End Type ' NMHDR + 2 + 16 + 16
' The NMPGSCROLL structure is a complete pain in the arse because the
' fwKeys member is a WORD, i.e. 2 bytes.  VB dword aligns rcParent
' if we use an integer, so we have to do silly stuff with bytes:
Private Type NMPGSCROLLVB
    hdr As NMHDR
    bTheRest(0 To 33) As Byte
End Type

' // PGN_CALCSIZE Notification Message

Private Const PGN_CALCSIZE = (PGN_FIRST - 2)

Private Const PGF_CALCWIDTH = 1
Private Const PGF_CALCHEIGHT = 2

Private Type NMPGCALCSIZE
    hdr As NMHDR
    dwFlag As Long
    iWidth As Long
    iHeight As Long
End Type

' //' //======================  End Pager Control ==========================================



' ======================================================================
' Interface:
' ======================================================================
Public Enum ECPGOrientation
    PGHorizontal = PGS_HORZ
    PGVertical = PGS_VERT
End Enum
Public Enum ECPGScrollDir
    PGScrollLeft = PGF_SCROLLLEFT
    PGScrollRight = PGF_SCROLLRIGHT
    PGScrollUp = PGF_SCROLLUP
    PGScrollDown = PGF_SCROLLDOWN
End Enum
Public Enum ECPGBorderStyle
    PGNone = 0
    PGFixedSingle = 1
End Enum
Public Enum ECPGButtonType
    PGTopOrLeft = PGB_TOPORLEFT
    PGBottomOrRight = PGB_BOTTOMORRIGHT
End Enum
Public Enum ECPGButtonState
    PGInvisible = PGF_INVISIBLE  'The button is not visible.
    PGNormal = PGF_NORMAL  'The button is in normal state.
    PGGrayed = PGF_GRAYED  'The button is in grayed state.
    PGDepressed = PGF_DEPRESSED  'The button is in pressed state.
    PGHot = PGF_HOT  'The button is in hot state.
End Enum
Public Event RequestSize(ByRef lWidth As Long, ByRef lHeight As Long)
Public Event Scroll(ByVal eDir As ECPGScrollDir, ByRef lDelta As Long)

' ======================================================================
' Private Implementation:
' ======================================================================
Implements ISubclass
Private m_emr As EMsgResponse
Private m_bSubClassing As Boolean
Private m_hWnd As Long
Private m_bDragDrop As Boolean
Private m_bAutoHScroll As Boolean
Private m_eOrientation As ECPGOrientation
Private m_lPosition As Long
Private m_lButtonSize As Long

Public Property Get hwnd() As Long
   hwnd = UserControl.hwnd
End Property
Public Property Get PagerhWnd() As Long
   PagerhWnd = m_hWnd
End Property

Public Property Get ButtonState(ByVal eButton As ECPGButtonType) As ECPGButtonState
    If (m_hWnd <> 0) Then
        ButtonState = SendMessageLong(m_hWnd, PGM_GETBUTTONSTATE, 0, eButton)
    End If
End Property

Public Property Get ButtonSize() As Long
    If (m_hWnd <> 0) Then
        ButtonSize = SendMessageLong(m_hWnd, PGM_GETBUTTONSIZE, 0, 0)
    End If
End Property
Public Property Let ButtonSize(ByVal lSize As Long)
    m_lButtonSize = lSize
    If (m_hWnd <> 0) Then
        SendMessageLong m_hWnd, PGM_SETBUTTONSIZE, 0, lSize
        PropertyChanged "ButtonSize"
    End If
End Property
Public Property Get BackColor() As OLE_COLOR
    BackColor = UserControl.BackColor
End Property
Public Property Let BackColor(ByVal oColor As OLE_COLOR)
    If (oColor <> UserControl.BackColor) Then
        UserControl.BackColor = oColor
        If (m_hWnd <> 0) Then
            SendMessageLong m_hWnd, PGM_SETBKCOLOR, 0, TranslateColor(oColor)
        End If
        PropertyChanged "BackColor"
    End If
End Property
Public Property Get Position() As Long
    If (m_hWnd <> 0) Then
        Position = SendMessageLong(m_hWnd, PGM_GETPOS, 0, 0)
    Else
        Position = m_lPosition
    End If
End Property
Public Property Let Position(ByVal lPos As Long)
    If (lPos <> m_lPosition) Then
        m_lPosition = lPos
        If (m_hWnd <> 0) Then
            SendMessageLong m_hWnd, PGM_SETPOS, 0, lPos
        End If
        PropertyChanged "Position"
    End If
End Property

Public Property Get BorderStyle() As ECPGBorderStyle
Dim lStyle As Long
    ' Determine if Client Edge extended style is set:
    lStyle = GetWindowLong(UserControl.hwnd, GWL_EXSTYLE)
    If ((lStyle And WS_EX_CLIENTEDGE) = WS_EX_CLIENTEDGE) Then
        BorderStyle = PGFixedSingle
    End If
End Property
Public Property Let BorderStyle(ByVal eStyle As ECPGBorderStyle)
Dim lStyle As Long
Dim lNStyle As Long
    ' Get window extended style:
    lStyle = GetWindowLong(UserControl.hwnd, GWL_EXSTYLE)
    ' Ensure the ClientEdge bit is set correctly:
    If (eStyle = PGFixedSingle) Then
        lNStyle = lStyle Or WS_EX_CLIENTEDGE
    Else
        lNStyle = lStyle And Not WS_EX_CLIENTEDGE
    End If
    ' If this results in a change:
    If (lNStyle <> lStyle) Then
        ' Change the window style:
        SetWindowLong UserControl.hwnd, GWL_EXSTYLE, lNStyle
        ' Ensure the style 'takes':
        SetWindowPos UserControl.hwnd, 0, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE Or SWP_NOREPOSITION Or SWP_NOZORDER Or SWP_FRAMECHANGED
        ' Refresh the UserControl:
        UserControl.Refresh
        
        PropertyChanged "BorderStyle"
    End If
End Property


Public Property Get InternalBorderWidth() As Long
    If (m_hWnd <> 0) Then
        InternalBorderWidth = SendMessageLong(m_hWnd, PGM_GETBORDER, 0, 0)
    End If
End Property
Public Property Let InternalBorderWidth(ByVal lWidth As Long)
    If (lWidth <> InternalBorderWidth) Then
        If (m_hWnd <> 0) Then
            SendMessageLong m_hWnd, PGM_SETBORDER, 0, lWidth
        End If
        PropertyChanged "InternalBorderWidth"
    End If
End Property

Public Property Get DragDrop() As Boolean
    DragDrop = m_bDragDrop
End Property
Public Property Let DragDrop(ByVal bDragDrop As Boolean)
    If (m_bDragDrop <> bDragDrop) Then
        m_bDragDrop = bDragDrop
        If (m_hWnd <> 0) Then
            pInitialise
        End If
        PropertyChanged "DragDrop"
    End If
End Property
Public Property Get AutoHScroll() As Boolean
    AutoHScroll = m_bAutoHScroll
End Property
Public Property Let AutoHScroll(ByVal bAutoHScroll As Boolean)
    If (m_bAutoHScroll <> bAutoHScroll) Then
        m_bAutoHScroll = bAutoHScroll
        If (m_hWnd <> 0) Then
            pInitialise
        End If
        PropertyChanged "AutoHScroll"
    End If
End Property
Public Property Get Orientation() As ECPGOrientation
    Orientation = m_eOrientation
End Property
Public Property Let Orientation(ByVal eOrientation As ECPGOrientation)
    If (m_eOrientation <> eOrientation) Then
        m_eOrientation = eOrientation
        If (m_hWnd <> 0) Then
            pInitialise
        End If
        PropertyChanged "Orientation"
    End If
End Property

Public Sub AddChildWindow(ByVal hWndA As Long)
Dim hWndTb As Long
Dim lButtons As Long
Dim rc As RECT
    ' Check for a VB toolbar:
    If (ClassName(hWndA) = "ToolbarWndClass") Then
        ' Make the toolbar flat:
        hWndTb = GetWindow(hWndA, GW_CHILD)
        pSetStyle hWndTb, TBSTYLE_FLAT, True
        ' Set the toolbar size:
        lButtons = SendMessageLong(hWndTb, TB_BUTTONCOUNT, 0, 0)
        If (lButtons > 0) Then
            SendMessage hWndTb, TB_GETITEMRECT, lButtons - 1, rc
            MoveWindow hWndTb, 0, 0, rc.Right, rc.Bottom, 1
        End If
    End If

    SetParent hWndA, m_hWnd
    SendMessageLong m_hWnd, PGM_SETCHILD, 0, hWndA
End Sub
Private Sub pSetStyle(ByVal lhWnd As Long, ByVal lStyleBit As Long, ByVal bState As Boolean)
Dim lStyle As Long
    lStyle = GetWindowLong(lhWnd, GWL_STYLE)
    If (bState) Then
        lStyle = lStyle Or lStyleBit
    Else
        lStyle = lStyle And Not lStyleBit
    End If
    SetWindowLong lhWnd, GWL_STYLE, lStyle
End Sub

Public Sub RecalcSize()
    SendMessageLong m_hWnd, PGM_RECALCSIZE, 0, 0
End Sub

Private Sub pInitialise()
Dim tICCEX As tagInitCommonControlsEx
Dim dwStyle As Long
    
    ' Ensure we don't already have UpDown control:
    pTerminate
    
    ' ENsure common controls are initialised for pagers:
    tICCEX.lngICC = ICC_PAGESCROLLER_CLASS
    tICCEX.lngSize = Len(tICCEX)
    InitCommonControlsEx tICCEX
    
    ' Create Pager Control:
    dwStyle = WS_VISIBLE Or WS_CHILD     'Or WS_BORDER
    dwStyle = dwStyle Or m_eOrientation
    If (m_bAutoHScroll) Then
        dwStyle = dwStyle Or PGS_AUTOSCROLL
    End If
    If (m_bDragDrop) Then
        dwStyle = dwStyle Or PGS_DRAGNDROP
    End If
    
    m_hWnd = CreateWindowEX(0, WC_PAGESCROLLER, "cVBPager", dwStyle, _
        0, 0, UserControl.ScaleWidth \ Screen.TwipsPerPixelX, UserControl.ScaleHeight \ Screen.TwipsPerPixelY, _
        UserControl.hwnd, 0, App.hInstance, UserControl.hwnd)
    Debug.Assert (m_hWnd <> 0)
    If (m_hWnd <> 0) Then
        If (UserControl.Ambient.UserMode) Then
            ' Attach Messages
            pAttachMessages
        End If
    End If
    
End Sub

Private Sub pTerminate()
    
    If (m_hWnd <> 0) Then
        ' Stop subclassing:
        pDetachMessages
        ' Destroy the window:
        ShowWindow m_hWnd, SW_HIDE
        SetParent m_hWnd, 0
        Debug.Print DestroyWindow(m_hWnd)
        m_hWnd = 0
    End If
    
End Sub
Private Sub pAttachMessages()
    AttachMessage Me, UserControl.hwnd, WM_NOTIFY
    m_emr = emrPreprocess
    m_bSubClassing = True
End Sub
Private Sub pDetachMessages()
    If (m_bSubClassing) Then
        DetachMessage Me, UserControl.hwnd, WM_NOTIFY
        m_bSubClassing = False
    End If
End Sub
' Convert Automation color to Windows color
Private Function TranslateColor(ByVal clr As OLE_COLOR, _
                        Optional hPal As Long = 0) As Long
    If OleTranslateColor(clr, hPal, TranslateColor) Then
       TranslateColor = CLR_INVALID
    End If
End Function
Private Function ClassName(ByVal lhWnd As Long) As String
Dim lLen As Long
Dim sBuf As String
    lLen = 260
    sBuf = String$(lLen, 0)
    lLen = GetClassName(lhWnd, sBuf, lLen)
    If (lLen <> 0) Then
        ClassName = Left$(sBuf, lLen)
    End If
End Function

Private Property Let ISubClass_MsgResponse(ByVal RHS As EMsgResponse)
    m_emr = RHS
End Property

Private Property Get ISubClass_MsgResponse() As EMsgResponse
    ISubClass_MsgResponse = m_emr
End Property

Private Function ISubClass_WindowProc(ByVal hwnd As Long, ByVal iMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Dim tNMH As NMHDR
Dim tPGCS As NMPGCALCSIZE
Dim tPGS As NMPGSCROLLVB
Dim iDir As Long, iScroll As Long

    ' Process Messages:
    Select Case iMsg
    Case WM_NOTIFY
        CopyMemory tNMH, ByVal lParam, Len(tNMH)
        If (tNMH.hwndFrom = m_hWnd) Then
            Select Case tNMH.code
            Case PGN_CALCSIZE
                CopyMemory tPGCS, ByVal lParam, Len(tPGCS)
                Select Case tPGCS.dwFlag
                Case PGF_CALCWIDTH
                    RaiseEvent RequestSize(tPGCS.iWidth, 0)
                Case PGF_CALCHEIGHT
                    RaiseEvent RequestSize(0, tPGCS.iHeight)
                End Select
                CopyMemory ByVal lParam, tPGCS, Len(tPGCS)
            Case PGN_SCROLL
                ' Silly stuff with bytes - see declaration of NMPGSCROLL:
                CopyMemory tPGS, ByVal lParam, Len(tPGS)
                CopyMemory iDir, tPGS.bTheRest(18), 4
                CopyMemory iScroll, tPGS.bTheRest(30), 4
                Debug.Print iDir, iScroll
                RaiseEvent Scroll(iDir, iScroll)
                CopyMemory tPGS.bTheRest(30), iScroll, 4
                CopyMemory ByVal lParam, tPGS, Len(tPGS)
            End Select
        End If
    End Select
End Function

Private Sub UserControl_Initialize()
    Debug.Print "cPager:Initialize"
    m_eOrientation = PGHorizontal
    m_bDragDrop = False
    m_bAutoHScroll = False
End Sub

Private Sub UserControl_InitProperties()
    BorderStyle = PGFixedSingle
    pInitialise
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    Orientation = PropBag.ReadProperty("Orientation", PGHorizontal)
    DragDrop = PropBag.ReadProperty("DragDrop", False)
    AutoHScroll = PropBag.ReadProperty("AutoHScroll", False)
    BorderStyle = PropBag.ReadProperty("BorderStyle", PGFixedSingle)
    pInitialise
    InternalBorderWidth = PropBag.ReadProperty("InternalBorderWidth", 0)
    Position = PropBag.ReadProperty("Position", 0)
    BackColor = PropBag.ReadProperty("BackColor", vbButtonFace)
End Sub

Private Sub UserControl_Resize()
    ' Resize:
    If (m_hWnd <> 0) Then
        MoveWindow m_hWnd, 0, 0, UserControl.ScaleWidth \ Screen.TwipsPerPixelX, UserControl.ScaleHeight \ Screen.TwipsPerPixelY, 1
    End If
End Sub

Private Sub UserControl_Terminate()
    pTerminate
    Debug.Print "cUpDown:Terminate"
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    PropBag.WriteProperty "Orientation", Orientation, PGHorizontal
    PropBag.WriteProperty "DragDrop", DragDrop, False
    PropBag.WriteProperty "AutoHScroll", AutoHScroll, False
    PropBag.WriteProperty "BorderStyle", BorderStyle, PGFixedSingle
    PropBag.WriteProperty "InternalBorderWidth", InternalBorderWidth, 0
    PropBag.WriteProperty "Position", Position, 0
    PropBag.WriteProperty "BackColor", BackColor, vbButtonFace
    pTerminate
End Sub
