VERSION 5.00
Begin VB.UserControl cToolbar 
   BorderStyle     =   1  'Fixed Single
   CanGetFocus     =   0   'False
   ClientHeight    =   540
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3855
   ClipControls    =   0   'False
   InvisibleAtRuntime=   -1  'True
   ScaleHeight     =   540
   ScaleWidth      =   3855
   ToolboxBitmap   =   "cToolbar.ctx":0000
   Begin VB.Label lblInfo 
      Caption         =   "'Toolbar control'"
      Height          =   255
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   4275
   End
End
Attribute VB_Name = "cToolbar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' =========================================================================
' vbAccelerator Toolbar control v2.0
' Copyright © 1998-1999 Steve McMahon (steve@dogma.demon.co.uk)
'
' This is a complete form toolbar implementation designed
' for hosting in a vbAccelerator ReBar control.
'
' -------------------------------------------------------------------------
' Visit vbAccelerator at http://vbaccelerator.com
' =========================================================================

' ==============================================================================
' Declares, constants and types required for toolbar:
' ==============================================================================

Private Type TBADDBITMAP
    hInst As Long
    nID As Long
End Type

Private Type NMTOOLBAR_SHORT
    hdr As NMHDR
    iItem As Long
End Type

Private Type TBBUTTONINFO
   cbSize As Long
   dwMask As Long
   idCommand As Long
   iImage As Long
   fsState As Byte
   fsStyle As Byte
   cX As Integer
   lParam As Long
   pszText As Long
   cchText As Long
End Type

Private Type NMTBHOTITEM
   hdr As NMHDR
   idOld As Long
   idNew As Long
   dwFlags As Long           '// HICF_*
End Type

Private Const TBIF_IMAGE = &H1&
Private Const TBIF_TEXT = &H2&
'Private Const TBIF_STATE = &H4&
'Private Const TBIF_STYLE = &H8&
'Private Const TBIF_LPARAM = &H10&
'Private Const TBIF_COMMAND = &H20&
'Private Const TBIF_SIZE = &H40&

' Toolbar button states:
Private Enum ectbButtonStates
   TBSTATE_CHECKED = &H1
   TBSTATE_PRESSED = &H2
   TBSTATE_ENABLED = &H4
   TBSTATE_WRAP = &H20
   TBSTATE_ELLIPSES = &H40
   TBSTATE_INDETERMINATE = &H10
   TBSTATE_HIDDEN = &H8
End Enum

' Toolbar notification messages:
'Private Const TBN_LAST = &H720
Private Const TBN_FIRST = -700&
'Private Const TBN_GETBUTTONINFOA = (TBN_FIRST - 0)
'Private Const TBN_GETBUTTONINFOW = (TBN_FIRST - 20)
'Private Const TBN_BEGINDRAG = (TBN_FIRST - 1)
'Private Const TBN_ENDDRAG = (TBN_FIRST - 2)
'Private Const TBN_BEGINADJUST = (TBN_FIRST - 3)
'Private Const TBN_ENDADJUST = (TBN_FIRST - 4)
'Private Const TBN_RESET = (TBN_FIRST - 5)
'Private Const TBN_QUERYINSERT = (TBN_FIRST - 6)
'Private Const TBN_QUERYDELETE = (TBN_FIRST - 7)
'Private Const TBN_TOOLBARCHANGE = (TBN_FIRST - 8)
'Private Const TBN_CUSTHELP = (TBN_FIRST - 9)
'Private Const TBN_CLOSEUP = (TBN_FIRST - 11)
Private Const TBN_DROPDOWN = (TBN_FIRST - 10)
Private Const TBN_HOTITEMCHANGE = (TBN_FIRST - 13)

' Toolbar messages:
'Private Const TB_ENABLEBUTTON = (WM_USER + 1)
Private Const TB_CHECKBUTTON = (WM_USER + 2)
'Private Const TB_HIDEBUTTON = (WM_USER + 4)
'Private Const TB_INDETERMINATE = (WM_USER + 5)
'Private Const TB_MARKBUTTON = (WM_USER + 6)

Private Const TB_SETSTATE = (WM_USER + 17)
Private Const TB_GETSTATE = (WM_USER + 18)

Private Const TB_ADDBITMAP = (WM_USER + 19)
Private Const TB_ADDBUTTONS = (WM_USER + 20)
Private Const TB_INSERTBUTTON = (WM_USER + 21)
Private Const TB_DELETEBUTTON = (WM_USER + 22)
Private Const TB_GETBUTTON = (WM_USER + 23)
Private Const TB_COMMANDTOINDEX = (WM_USER + 25)

'Private Const TB_SAVERESTOREA = (WM_USER + 26)
'Private Const TB_SAVERESTOREW = (WM_USER + 76)
'Private Const TB_CUSTOMIZE = (WM_USER + 27)
Private Const TB_ADDSTRING = (WM_USER + 28)
Private Const TB_GETITEMRECT = (WM_USER + 29)

Private Const TB_BUTTONSTRUCTSIZE = (WM_USER + 30)
'Private Const TB_SETBUTTONSIZE = (WM_USER + 31)
Private Const TB_SETBITMAPSIZE = (WM_USER + 32)
Private Const TB_AUTOSIZE = (WM_USER + 33)

'Private Const TB_GETTOOLTIPS = (WM_USER + 35)
'Private Const TB_SETTOOLTIPS = (WM_USER + 36)
Private Const TB_SETPARENT = (WM_USER + 37)
'Private Const TB_SETROWS = (WM_USER + 39)
'Private Const TB_GETROWS = (WM_USER + 40)
'Private Const TB_SETCMDID = (WM_USER + 42)
'Private Const TB_CHANGEBITMAP = (WM_USER + 43)
'Private Const TB_GETBITMAP = (WM_USER + 44)
'Private Const TB_GETBUTTONTEXTA = (WM_USER + 45)
'Private Const TB_GETBUTTONTEXTW = (WM_USER + 75)

'#if (_WIN32_IE >= 0x0300)
'Private Const TB_SETINDENT = (WM_USER + 47)
Private Const TB_SETIMAGELIST = (WM_USER + 48)
'Private Const TB_GETIMAGELIST = (WM_USER + 49)
Private Const TB_LOADIMAGES = (WM_USER + 50)
Private Const TB_GETRECT = (WM_USER + 51)             '// wParam is the Cmd instead of index
Private Const TB_SETHOTIMAGELIST = (WM_USER + 52)
'Private Const TB_GETHOTIMAGELIST = (WM_USER + 53)
Private Const TB_SETDISABLEDIMAGELIST = (WM_USER + 54)
'Private Const TB_GETDISABLEDIMAGELIST = (WM_USER + 55)
'Private Const TB_SETSTYLE = (WM_USER + 56)
'Private Const TB_GETSTYLE = (WM_USER + 57)
'Private Const TB_GETBUTTONSIZE = (WM_USER + 58)
'Private Const TB_SETBUTTONWIDTH = (WM_USER + 59)
Private Const TB_SETMAXTEXTROWS = (WM_USER + 60)
Private Const TB_GETTEXTROWS = (WM_USER + 61)
'#endif

'Private Const TB_GETBUTTONINFO = (WM_USER + 65)
Private Const TB_SETBUTTONINFO = (WM_USER + 66)

'#if (_WIN32_IE >= 0x0400)
'Private Const TB_GETOBJECT = (WM_USER + 62)            '// wParam == IID, lParam void **ppv
'Private Const TB_SETANCHORHIGHLIGHT = (WM_USER + 73)   '// wParam == TRUE/FALSE
'Private Const TB_GETANCHORHIGHLIGHT = (WM_USER + 74)
Private Const TB_MAPACCELERATORA = (WM_USER + 78)      '// wParam == ch, lParam int * pidBtn
'Private Const TB_MAPACCELERATORW = (WM_USER + 90)      '// wParam == ch,
'Private Const TB_MAPACCELERATOR = TB_MAPACCELERATORA

'Private Type TBINSERTMARK
'    iButton As Long
'    dwFlags As Long
'End Type
'Private Const TBIMHT_AFTER = &H1      '// TRUE = insert After iButton, otherwise before
'Private Const TBIMHT_BACKGROUND = &H2 '// TRUE iff missed buttons completely

'Private Const TB_GETINSERTMARK = (WM_USER + 79)        '// lParam == LPTBINSERTMARK
'Private Const TB_SETINSERTMARK = (WM_USER + 80)        '// lParam == LPTBINSERTMARK
'Private Const TB_INSERTMARKHITTEST = (WM_USER + 81)    '// wParam == LPPOINT lParam == LPTBINSERTMARK
'Private Const TB_MOVEBUTTON = (WM_USER + 82)

'Private Const TB_GETMAXSIZE = (WM_USER + 83)           '// lParam == LPSIZE

' Extended style:
Private Const TB_SETEXTENDEDSTYLE = (WM_USER + 84)    ' // For TBSTYLE_EX_*
'Private Const TB_GETEXTENDEDSTYLE = (WM_USER + 85)     '// For TBSTYLE_EX_*
Private Const TB_GETPADDING = (WM_USER + 86)
Private Const TB_SETPADDING = (WM_USER + 87)
'Private Const TB_SETINSERTMARKCOLOR = (WM_USER + 88)
'Private Const TB_GETINSERTMARKCOLOR = (WM_USER + 89)

'Private Const TB_SETCOLORSCHEME = CCM_SETCOLORSCHEME       '// lParam is color scheme
'Private Const TB_GETCOLORSCHEME = CCM_GETCOLORSCHEME       '// fills in COLORSCHEME pointed to by lParam
'#endif  // _WIN32_IE >= 0x0400

Private Const TBSTYLE_EX_DRAWDDARROWS = &H1

'//Standard image types:
Private Const IDB_STD_SMALL_COLOR = 0
Private Const IDB_STD_LARGE_COLOR = 1
Private Const IDB_VIEW_SMALL_COLOR = 4
Private Const IDB_VIEW_LARGE_COLOR = 5
Private Const IDB_HIST_SMALL_COLOR = 8
Private Const IDB_HIST_LARGE_COLOR = 9

'// icon indexes for standard bitmap

Private Const STD_CUT = 0
Private Const STD_COPY = 1
Private Const STD_PASTE = 2
Private Const STD_UNDO = 3
Private Const STD_REDOW = 4
Private Const STD_DELETE = 5
Private Const STD_FILENEW = 6
Private Const STD_FILEOPEN = 7
Private Const STD_FILESAVE = 8
Private Const STD_PRINTPRE = 9
Private Const STD_PROPERTIES = 10
Private Const STD_HELP = 11
Private Const STD_FIND = 12
Private Const STD_REPLACE = 13
Private Const STD_PRINT = 14

'// icon indexes for standard view bitmap

Private Const VIEW_LARGEICONS = 0
Private Const VIEW_SMALLICONS = 1
Private Const VIEW_LIST = 2
Private Const VIEW_DETAILS = 3
Private Const VIEW_SORTNAME = 4
Private Const VIEW_SORTSIZE = 5
Private Const VIEW_SORTDATE = 6
Private Const VIEW_SORTTYPE = 7
Private Const VIEW_PARENTFOLDER = 8
Private Const VIEW_NETCONNECT = 9
Private Const VIEW_NETDISCONNECT = 10
Private Const VIEW_NEWFOLDER = 11
'#if (_WIN32_IE >= 0x0400)
'Private Const VIEW_VIEWMENU = 12
'#End If

'#if (_WIN32_IE >= 0x0300)
Private Const HIST_BACK = 0
Private Const HIST_FORWARD = 1
Private Const HIST_FAVORITES = 2
Private Const HIST_ADDTOFAVORITES = 3
Private Const HIST_VIEWTREE = 4
'#End If

Private Declare Function CreateToolbarEx Lib "COMCTL32" (ByVal hwnd As Long, ByVal ws As Long, ByVal wID As Long, ByVal nBitmaps As Long, ByVal hBMInst As Long, ByVal wBMID As Long, ByRef lpButtons As TBBUTTON, ByVal iNumButtons As Long, ByVal dxButton As Long, ByVal dyButton As Long, ByVal dxBitmap As Long, ByVal dyBitmap As Long, ByVal uStructSize As Long) As Long

' ==============================================================================
' INTERFACE
' ==============================================================================
' Enumerations:
Public Enum ECTBToolButtonSyle
    CTBNormal = TBSTYLE_BUTTON
    CTBSeparator = TBSTYLE_SEP
    CTBCheck = TBSTYLE_CHECK
    CTBCheckGroup = TBSTYLE_CHECKGROUP
    CTBDropDown = TBSTYLE_DROPDOWN
    CTBAutoSize = TBSTYLE_AUTOSIZE
    CTBDropDownArrow = BTNS_WHOLEDROPDOWN
End Enum
Public Enum ECTBImageListTypes
   CTBImageListNormal = TB_SETIMAGELIST
   CTBImageListHot = TB_SETHOTIMAGELIST
   CTBImageListDisabled = TB_SETDISABLEDIMAGELIST
End Enum
Public Enum ECTBToolbarStyle
    CTBFlat = TBSTYLE_FLAT
    CTBList = TBSTYLE_LIST
    CTBTransparent = -1 ' special - here we remove Toolbar from owner window
End Enum
Public Enum ECTBImageSourceTypes
    CTBResourceBitmap
    CTBLoadFromFile
    CTBExternalImageList
    CTBPicture
    CTBStandardImageSources
End Enum
Public Enum ECTBStandardImageSourceTypes
   CTBHistoryLargeColor = IDB_HIST_LARGE_COLOR
   CTBHistorySmallColor = IDB_HIST_SMALL_COLOR
   CTBStandardLargeColor = IDB_STD_LARGE_COLOR
   CTBStandardSmallColor = IDB_STD_SMALL_COLOR
   CTBViewLargeColor = IDB_VIEW_LARGE_COLOR
   CTBViewSmallColor = IDB_VIEW_SMALL_COLOR
End Enum
Public Enum ECTBStandardImageIndexConstants
   ' History:
   CTBHistAddToFavourites = HIST_ADDTOFAVORITES ' 'Add 'to 'favorites.
   CTBHistBack = HIST_BACK ' 'Move 'back.
   CTBHistFavourites = HIST_FAVORITES ' 'Open 'favorites 'folder.
   CTBHistForward = HIST_FORWARD ' 'Move 'forward.
   CTBHistViewTree = HIST_VIEWTREE ' 'View 'tree.
   'Standard:
   CTBStdCopy = STD_COPY ' 'Copy 'operation.
   CTBStdCut = STD_CUT ' 'Cut 'operation.
   CTBStdDelete = STD_DELETE ' 'Delete 'operation.
   CTBStdFileNew = STD_FILENEW ' 'New 'file 'operation.
   CTBStdFileOpen = STD_FILEOPEN ' 'Open 'file 'operation.
   CTBStdFIleSave = STD_FILESAVE ' 'Save 'file 'operation.
   CTBStdFind = STD_FIND ' 'Find 'operation.
   CTBStdHelp = STD_HELP ' 'Help 'operation.
   CTBStdPaste = STD_PASTE ' 'Paste 'operation.
   CTBStdPrint = STD_PRINT ' 'Print 'operation.
   CTBStdPrintPreview = STD_PRINTPRE ' 'Print 'preview 'operation.
   CTBStdProperties = STD_PROPERTIES ' 'Properties 'operation.
   CTBStdRedo = STD_REDOW ' 'Redo 'operation.
   CTBStdReplace = STD_REPLACE ' 'Replace 'operation.
   CTBStdUndo = STD_UNDO ' 'Undo 'operation.
   'View
   CTBViewDetails = VIEW_DETAILS ' 'Details 'view.
   CTBViewLargeIcons = VIEW_LARGEICONS ' 'Large 'icons 'view.
   CTBViewList = VIEW_LIST ' 'List 'view.
   CTBViewNetConnect = VIEW_NETCONNECT ' 'Connect 'to 'network 'drive.
   CTBViewNetDisconnect = VIEW_NETDISCONNECT ' 'Disconnect 'from 'network 'drive.
   CTBViewNewFolder = VIEW_NEWFOLDER ' 'New 'folder.
   CTBViewParentFolder = VIEW_PARENTFOLDER ' 'Go 'to 'parent 'folder.
   CTBViewSmallIcons = VIEW_SMALLICONS ' 'Small 'icon 'view.
   CTBViewSortDate = VIEW_SORTDATE ' 'Sort 'by 'date.
   CTBViewSortName = VIEW_SORTNAME ' 'Sort 'by 'name.
   CTBViewSortSize = VIEW_SORTSIZE ' 'Sort 'by 'size.
   CTBViewSortType = VIEW_SORTTYPE ' 'Sort 'by 'type.
End Enum
Public Enum ECTBHotItemChangeReasonConstants
   HICF_OTHER = 0
   HICF_MOUSE = 1 '// Triggered by mouse
   HICF_ARROWKEYS = 2 ' // Triggered by arrow keys
   HICF_ACCELERATOR = 4  '// Triggered by accelerator
   HICF_DUPACCEL = 8               '// This accelerator is not unique
   HICF_ENTERING = 10               '// idOld is invalid
   HICF_LEAVING = 20                '// idNew is invalid
   HICF_RESELECT = 40               '// hot item reselected
End Enum

' Events:
Public Event ButtonClick(ByVal lButton As Long)
Attribute ButtonClick.VB_Description = "Raised when a toolbar button is clicked."
Public Event DropDownPress(ByVal lButton As Long)
Attribute DropDownPress.VB_Description = "Raised when a drop-down arrow on a drop-down button is pressed (Note: COMCTL32.DLL versions below 4.71 do not display drop-down buttons)"
Public Event HotItemChange(ByVal iNew As Long, ByVal iOld As Long, ByVal eReason As ECTBHotItemChangeReasonConstants)
Attribute HotItemChange.VB_Description = "Raised whenever the hot button changes in a flat toolbar."

' ==============================================================================
' INTERNAL INFORMATION
' ==============================================================================
' Subclassing
Implements ISubclass
Private m_bInSubClass As Boolean

' Classes to turn toolbar into a menu:
Private m_cMenu As New cTbarMenu

Private m_bIsMenu As Boolean
Private m_hMenu As Long
Private m_lPtrMenu As Long

' Hwnd of tool bar itself:
Private m_hWndToolBar As Long
Private m_hWndParentForm As Long
' Where the button images are coming from
Private m_eImageSourceType As ECTBImageSourceTypes
Private m_pic As StdPicture
Private m_hBmp As Long
Private m_sFileName As String
Private m_lResourceID As Long
Private m_hInstance As Long
Private m_hIml As Long
Private m_hImlHot As Long
Private m_hImlDis As Long
Private m_eStandardType As ECTBStandardImageSourceTypes

' Button size:
Private m_iButtonWidth As Integer
Private m_iButtonHeight As Integer

' Style information:
Private m_bWithText As Boolean
Private m_bWrappable As Boolean

' Button information:
' Types:
Private Type ButtonInfoStore
    wID As Integer
    iImage As Integer
    sTipText As String
    iTextIndexNum As Integer
    sCaption As String
    bShowText As Boolean
    idString As Long
    iLarge As Integer
    xWidth As Integer
    xHeight As Integer
    sKey As String
    eStyle As ECTBToolButtonSyle
    hSubMenu As Long
End Type
Private m_tBInfo() As ButtonInfoStore
' Last return code from toolbar API or sendmessage call
Private m_lR As Long

' Strings in the toolbar:
Private m_lStringIDCount As Long
Private m_sString() As String
Private m_lStringID() As Long

' Common Controls Version:
Private m_lMajorVer As Long
Private m_lMinorVer As Long
Private m_lBuild As Long

Friend Function AltKeyPress(ByVal eKeyCode As KeyCodeConstants) As Boolean
Dim wID As Long
Dim iKey As Long
Dim iB As Long
Dim i As Long
Dim sAccel As String

   If m_hWndToolBar <> 0 Then
      iB = -1
      sAccel = UCase$(Chr$(eKeyCode))
      For i = 0 To ButtonCount - 1
         If psGetAccelerator(m_tBInfo(i).sCaption) = sAccel Then
            iB = i
            wID = m_tBInfo(i).wID
            Exit For
         End If
      Next i
      If iB > -1 Then
         ' Am i a member of an active form?
         If m_hWndParentForm = GetActiveWindow() Then
            ButtonPressed(iB) = True
            SendMessageLong m_hWndToolBar, WM_COMMAND, wID, m_hWndToolBar
            ButtonPressed(iB) = False
            AltKeyPress = True
         End If
      Else
         'Debug.Assert iB > -1
      End If
   End If
   
End Function

Friend Sub pMenuClick(ByVal iButton As Long)
Dim lR As Long
   
   'Debug.Print iButton
   If Not m_lPtrMenu = 0 Then
      PopupObject.CreateSubClass m_hWndParentForm
   End If

   m_cMenu.CoolMenuAttach m_hWndParentForm, m_hWndToolBar, m_hMenu
   lR = m_cMenu.TrackPopup(iButton)
   m_cMenu.CoolMenuDetach
   
   If Not m_lPtrMenu = 0 Then
      If lR <> 0 Then
         Debug.Print "THAT WAS MENU ITEM: ", lR
         PopupObject.EmulateMenuClick lR
      End If
      PopupObject.DestroySubClass
   End If
   
End Sub

Private Property Get PopupObject() As Object
Dim oTemp As Object
   CopyMemory oTemp, m_lPtrMenu, 4
   Set PopupObject = oTemp
   CopyMemory oTemp, 0&, 4
End Property

Public Property Get AutosizeButtonPadding() As Long
Attribute AutosizeButtonPadding.VB_Description = "Gets/sets the number of pixels by which to pad out buttons with the CTBAutosize property set."
   ' NB Only applies to autosize buttons
   If m_hWndToolBar <> 0 Then
      AutosizeButtonPadding = (SendMessageLong(m_hWndToolBar, TB_GETPADDING, 0, 0) And &H7FFF&)
   End If
End Property
Public Property Let AutosizeButtonPadding(ByVal lPadding As Long)
Dim lxy As Long
   If m_hWndToolBar <> 0 Then
      lxy = (lPadding And &H7FFF&) Or (lPadding And &H7FFF& * &H10000)
      SendMessageLong m_hWndToolBar, TB_SETPADDING, 0, lxy
   End If
End Property

Public Sub GetComCtrlVersionInfo( _
      ByRef lMajor As Long, _
      ByRef lMinor As Long, _
      Optional ByRef lBuild As Long _
   )
Attribute GetComCtrlVersionInfo.VB_Description = "Returns the system's COMCTL32.DLL version."
   lMajor = m_lMajorVer
   lMinor = m_lMinorVer
   lBuild = m_lBuild
   End Sub
      

Public Property Get ButtonCount() As Long
Attribute ButtonCount.VB_Description = "Returns the number of buttons in a toolbar."
   If m_hWndToolBar <> 0 Then
      ButtonCount = SendMessageLong(m_hWndToolBar, TB_BUTTONCOUNT, 0, 0)
   End If
End Property

Public Property Get ButtonToolTip(ByVal vButton As Variant) As String
Attribute ButtonToolTip.VB_Description = "Gets/sets the tool tip shown for a button."
Dim iB As Long
    iB = ButtonIndex(vButton)
    If (iB > -1) Then
        ButtonToolTip = m_tBInfo(iB).sTipText
    End If
End Property
Public Property Let ButtonToolTip(ByVal vButton As Variant, ByVal sToolTip As String)
Dim iB As Long
    iB = ButtonIndex(vButton)
    If (iB > -1) Then
        m_tBInfo(iB).sTipText = sToolTip
    End If
End Property
Private Function pbGetIndexForID(ByVal iBtnId As Long) As Long
Dim iB As Long
    pbGetIndexForID = -1
    For iB = 0 To UBound(m_tBInfo)
        If (m_tBInfo(iB).wID = iBtnId) Then
            pbGetIndexForID = iB
            Exit For
        End If
    Next iB
End Function

Public Property Get ButtonImage(ByVal vButton As Variant) As Long
Attribute ButtonImage.VB_Description = "Gets/sets the zero based index of a button's image."
Dim iB As Long
   iB = ButtonIndex(vButton)
   If (iB <> -1) Then
      ButtonImage = m_tBInfo(iB).iImage
   End If
End Property
Public Property Let ButtonImage(ByVal vButton As Variant, ByVal iImage As Long)
Dim iB As Long

   ' If we are running pre 4.71 we must remove the button and add it again.
   ' 4.71+ we can use the TB_SETBUTTONINFO method to change it on the fly:
   If (m_lMajorVer > 4) Or ((m_lMajorVer = 4) And (m_lMinorVer > 70)) Then
      Dim tBI As TBBUTTONINFO
      Dim iID As Long
      
      iB = ButtonIndex(vButton)
      If (iB <> -1) Then
         iID = m_tBInfo(iB).wID
         tBI.cbSize = Len(tBI)
         tBI.dwMask = TBIF_IMAGE
         tBI.iImage = iImage
         If (SendMessage(m_hWndToolBar, TB_SETBUTTONINFO, iID, tBI) <> 0) Then
            m_tBInfo(iB).iImage = iImage
         End If
      End If
   Else
      iB = ButtonIndex(vButton)
      If (iB <> -1) Then
         ' Delete this button...
         'RemoveButton iB
         '
      End If
      
   End If
End Property

Public Property Get ButtonCaption(ByVal vButton As Variant) As String
Attribute ButtonCaption.VB_Description = "Gets/sets the caption of a button."
Dim iB As Long
    iB = ButtonIndex(vButton)
    If (iB <> -1) Then
        ButtonCaption = m_tBInfo(iB).sCaption
    End If
End Property
Public Property Let ButtonCaption(ByVal vButton As Variant, ByVal sCaption As String)
Dim iB As Integer
Dim bEnd As Boolean

   iB = ButtonIndex(vButton)
   If (iB > -1) Then
      
   
      If ((m_lMajorVer > 4) Or ((m_lMajorVer = 4) And (m_lMinorVer > 70))) And sCaption <> "" Then
         Dim tBI As TBBUTTONINFO
         Dim sBuf As String
         Dim iID As Long
         
         If iB <> -1 Then
            ' Remove any existing accelerator associated with caption:
            plRemoveString m_tBInfo(iB).sCaption
         
            ' don't add too many strings...
            plAddStringIfRequired sCaption
            If m_tBInfo(iB).bShowText Then
               sBuf = sCaption
               sBuf = sBuf & String$(80 - Len(sBuf), 0)
            Else
               sBuf = String$(80, 0)
            End If
            sBuf = StrConv(sBuf, vbFromUnicode)
            
            iID = m_tBInfo(iB).wID
            tBI.cbSize = Len(tBI)
            tBI.pszText = StrPtr(sBuf)
            tBI.dwMask = TBIF_TEXT
            If (SendMessage(m_hWndToolBar, TB_SETBUTTONINFO, iID, tBI) <> 0) Then
               m_tBInfo(iB).sCaption = sCaption
            End If
            
         End If
      Else
      
         ' Hmmm.  YOu can't remove any of the captions that have
         ' been added to the toolbar control, so if we keep on
         ' adding the damn things...  Don't change button captions
         ' to too many different things!
         Dim tBInfo As ButtonInfoStore
         LSet tBInfo = m_tBInfo(iB)
         If iB = ButtonCount - 1 Then
            bEnd = True
         End If
         RemoveButton iB
         If bEnd Then
            AddButton tBInfo.sTipText, tBInfo.iImage, , tBInfo.iLarge, sCaption, tBInfo.eStyle, tBInfo.sKey
         Else
            AddButton tBInfo.sTipText, tBInfo.iImage, iB, tBInfo.iLarge, sCaption, tBInfo.eStyle, tBInfo.sKey
         End If
      End If
   End If

End Property
Public Property Get ButtonTextVisible(ByVal vButton As Variant) As Boolean
Attribute ButtonTextVisible.VB_Description = "Gets/sets whether the caption for a button is visible or not."
Dim iB As Integer
   iB = ButtonIndex(vButton)
   If iB > -1 Then
      ButtonTextVisible = m_tBInfo(iB).bShowText
   End If
End Property
Public Property Let ButtonTextVisible(ByVal vButton As Variant, ByVal bState As Boolean)
Dim iB As Integer
Dim tBI As ButtonInfoStore
Dim bEnd As Boolean
Dim bChecked As Boolean
Dim bEnabled As Boolean
Dim bVisible As Boolean, bSet As Boolean
Dim lStyle As Long, lR As Long

   lStyle = GetWindowLong(m_hWndToolBar, GWL_STYLE)
   If (lStyle And TBSTYLE_LIST) <> TBSTYLE_LIST Then
      lR = SendMessageLong(m_hWndToolBar, TB_GETTEXTROWS, 0, 0)
      If bState Then
         If lR < 1 Then
            SendMessageLong m_hWndToolBar, TB_SETMAXTEXTROWS, 1, 0
            bSet = True
         End If
      Else
         If lR > 0 Then
            SendMessageLong m_hWndToolBar, TB_SETMAXTEXTROWS, 0, 0
            bSet = True
         End If
      End If
      If bSet Then
         For iB = 0 To ButtonCount - 1
            m_tBInfo(iB).bShowText = bState
         Next iB
      End If
   Else
      iB = ButtonIndex(vButton)
      If iB > -1 Then
         If bState <> m_tBInfo(iB).bShowText Then
            ' Hide/show text for this button:
            bChecked = ButtonChecked(iB)
            bEnabled = ButtonEnabled(iB)
            bVisible = ButtonVisible(iB)
   
            LSet tBI = m_tBInfo(iB)
            bEnd = (iB = ButtonCount - 1)
            
            RemoveButton iB
            If bEnd Then
               If bState Then
                  AddButton tBI.sTipText, tBI.iImage, , tBI.iLarge, tBI.sCaption, tBI.eStyle, tBI.sKey
               Else
                  AddButton tBI.sTipText, tBI.iImage, , tBI.iLarge, , tBI.eStyle, tBI.sKey
                  With m_tBInfo(iB)
                     .sCaption = tBI.sCaption
                  End With
               End If
            Else
               If bState Then
                  AddButton tBI.sTipText, tBI.iImage, iB, tBI.iLarge, tBI.sCaption, tBI.eStyle, tBI.sKey
               Else
                  AddButton tBI.sTipText, tBI.iImage, iB, tBI.iLarge, , tBI.eStyle, tBI.sKey
                  With m_tBInfo(iB)
                     .sCaption = tBI.sCaption
                  End With
               End If
            End If
            ButtonEnabled(iB) = bEnabled
            ButtonChecked(iB) = bChecked
            ButtonVisible(iB) = bVisible
            m_tBInfo(iB).bShowText = bState
            
         End If
      End If
   End If
End Property

Public Property Get ButtonIndex(ByVal vButton As Variant) As Integer
Attribute ButtonIndex.VB_Description = "Returns the zero based index of a button given its key or position."
Dim iB As Integer
Dim iIndex As Integer
    iIndex = -1
    If (IsNumeric(vButton)) Then
        iIndex = CInt(vButton)
    Else
        For iB = 0 To UBound(m_tBInfo)
            If (m_tBInfo(iB).sKey = vButton) Then
                iIndex = iB
                Exit For
            End If
        Next iB
    End If
    If (iIndex > -1) And (iIndex <= UBound(m_tBInfo)) Then
        ButtonIndex = iIndex
    Else
        ' error
        debugmsg "Button index failed"
        ButtonIndex = -1
    End If
    
End Property
Public Property Get ButtonKey(ByVal iButton As Long) As String
Attribute ButtonKey.VB_Description = "Returns the key of a button given its position."
   If (iButton > -1) And (iButton < ButtonCount) Then
      ButtonKey = m_tBInfo(iButton).sKey
   End If
End Property

Public Property Get ButtonEnabled(ByVal vButton As Variant) As Boolean
Attribute ButtonEnabled.VB_Description = "Gets/sets whether a button is enabled."
Dim iButton As Long
Dim iID As Long
    iButton = ButtonIndex(vButton)
    If (iButton <> -1) Then
        iID = m_tBInfo(iButton).wID
        ButtonEnabled = pbGetState(iID, TBSTATE_ENABLED)
    End If
End Property
Public Property Let ButtonEnabled(ByVal vButton As Variant, ByVal bState As Boolean)
Dim iButton As Long
Dim iID As Long
Dim lEnable As Long
    iButton = ButtonIndex(vButton)
    If (iButton <> -1) Then
        iID = m_tBInfo(iButton).wID
        pbSetState iID, TBSTATE_ENABLED, bState
    End If
End Property
Public Property Get ButtonVisible(ByVal vButton As Variant) As Boolean
Attribute ButtonVisible.VB_Description = "Gets/sets whether a button is visible in the toolbar."
Dim iButton As Long
Dim iID As Long
    iButton = ButtonIndex(vButton)
    If (iButton <> -1) Then
        iID = m_tBInfo(iButton).wID
        ButtonVisible = Not (pbGetState(iID, TBSTATE_HIDDEN))
    End If
End Property
Public Property Let ButtonVisible(ByVal vButton As Variant, ByVal bState As Boolean)
Dim iButton As Long
Dim iID As Long
Dim lEnable As Long
    iButton = ButtonIndex(vButton)
    If (iButton <> -1) Then
        iID = m_tBInfo(iButton).wID
        pbSetState iID, TBSTATE_HIDDEN, Not (bState)
        ResizeToolbar
    End If
End Property
Public Property Get ButtonWidth(ByVal vButton As Variant)
Dim iButton As Long
Dim tR As RECT
   iButton = ButtonIndex(vButton)
   If (iButton <> -1) Then
      SendMessage m_hWndToolBar, TB_GETRECT, m_tBInfo(iButton).wID, tR
      ButtonWidth = tR.Right - tR.Left
   End If
End Property
Public Property Get ButtonHeight(ByVal vButton As Variant) As Long
Dim iButton As Long
Dim tR As RECT
   iButton = ButtonIndex(vButton)
   If (iButton <> -1) Then
      SendMessage m_hWndToolBar, TB_GETRECT, m_tBInfo(iButton).wID, tR
      ButtonHeight = tR.Bottom - tR.Top
   End If
End Property
Public Property Get ButtonHot(ByVal vButton As Variant) As Boolean
Attribute ButtonHot.VB_Description = "Gets/sets whether a button in a flat toolbar appears in the ""hot"" state (i.e. looks like the mouse is over it)"
Dim iB As Integer
   iB = ButtonIndex(vButton)
   If iB > -1 Then
      ButtonHot = (SendMessageLong(m_hWndToolBar, TB_GETHOTITEM, 0, 0) = iB)
   End If
End Property
Public Property Let ButtonHot(ByVal vButton As Variant, ByVal bHot As Boolean)
Dim iB As Integer
   iB = ButtonIndex(vButton)
   If iB > -1 Then
      If ButtonHot(iB) Then
         If Not bHot Then
            SendMessageLong m_hWndToolBar, TB_SETHOTITEM, -1, 0
         End If
      Else
         If bHot Then
            SendMessageLong m_hWndToolBar, TB_SETHOTITEM, iB, 0
         End If
      End If
   End If
End Property
Public Property Get MaxButtonWidth() As Long
Attribute MaxButtonWidth.VB_Description = "Gets/sets the maximum allowable button width."
Dim i As Long
Dim lW As Long
Dim lMaxW As Long
   For i = 0 To ButtonCount - 1
      lW = ButtonWidth(i)
      If lW > lMaxW Then
         lMaxW = lW
      End If
   Next i
   MaxButtonWidth = lMaxW
End Property
Public Property Get MaxButtonHeight() As Long
Attribute MaxButtonHeight.VB_Description = "Gets/sets the maximum allowable button height."
Dim i As Long
Dim lH As Long
Dim lMaxH As Long
   For i = 0 To ButtonCount - 1
      lH = ButtonHeight(i)
      If lH > lMaxH Then
         lMaxH = lH
      End If
   Next i
   MaxButtonHeight = lMaxH
End Property
Public Property Get ButtonChecked(ByVal vButton As Variant) As Boolean
Attribute ButtonChecked.VB_Description = "Gets/sets whether a button is checked or not (if the button has the checked or check group style)"
Dim iButton As Long
Dim iID As Long
    iButton = ButtonIndex(vButton)
    If (iButton <> -1) Then
        iID = m_tBInfo(iButton).wID
        ButtonChecked = pbGetState(iID, TBSTATE_CHECKED)
    End If
End Property
Public Property Let ButtonChecked(ByVal vButton As Variant, ByVal bState As Boolean)
Dim iButton As Long
Dim iID As Long
   iButton = ButtonIndex(vButton)
   If (iButton <> -1) Then
      iID = m_tBInfo(iButton).wID
      'Check the button
      SendMessageLong m_hWndToolBar, TB_CHECKBUTTON, iID, Abs(bState)
      If (ButtonPressed(iButton) <> bState) Then
         SendMessageLong m_hWndToolBar, TB_CHECKBUTTON, iID, Abs(bState)
      End If
   End If
End Property
Public Property Get ButtonPressed(ByVal vButton As Variant) As Boolean
Attribute ButtonPressed.VB_Description = "Gets/sets whether a button is pressed."
Dim iButton As Long
Dim iID As Long
    iButton = ButtonIndex(vButton)
    If (iButton <> -1) Then
        iID = m_tBInfo(iButton).wID
        ButtonPressed = pbGetState(iID, TBSTATE_PRESSED)
    End If
End Property
Public Property Let ButtonPressed(ByVal vButton As Variant, ByVal bState As Boolean)
Dim iButton As Long
Dim iID As Long
    iButton = ButtonIndex(vButton)
    If (iButton <> -1) Then
        iID = m_tBInfo(iButton).wID
        pbSetState iID, TBSTATE_PRESSED, bState
    End If
End Property
Public Property Get ButtonTextWrap(ByVal vButton As Variant) As Boolean
Attribute ButtonTextWrap.VB_Description = "Gets/sets whether button text will wrap onto a newline if it is too long."
Dim iButton As Long
Dim iID As Long
    iButton = ButtonIndex(vButton)
    If (iButton <> -1) Then
        iID = m_tBInfo(iButton).wID
        ButtonTextWrap = pbGetState(iID, TBSTATE_WRAP)
    End If
End Property
Public Property Let ButtonTextWrap(ByVal vButton As Variant, ByVal bState As Boolean)
Dim iButton As Long
Dim iID As Long
    iButton = ButtonIndex(vButton)
    If (iButton <> -1) Then
        iID = m_tBInfo(iButton).wID
        pbSetState iID, TBSTATE_WRAP, bState
    End If
End Property
Public Property Get ButtonTextEllipses(ByVal vButton As Variant) As Boolean
Attribute ButtonTextEllipses.VB_Description = "Gets/sets whether button text will be truncated if the button text is too long."
Dim iButton As Long
Dim iID As Long
    iButton = ButtonIndex(vButton)
    If (iButton <> -1) Then
        iID = m_tBInfo(iButton).wID
        ButtonTextEllipses = pbGetState(iID, TBSTATE_ELLIPSES)
    End If
End Property
Public Property Let ButtonTextEllipses(ByVal vButton As Variant, ByVal bState As Boolean)
Dim iButton As Long
Dim iID As Long
    iButton = ButtonIndex(vButton)
    If (iButton <> -1) Then
        iID = m_tBInfo(iButton).wID
        pbSetState iID, TBSTATE_ELLIPSES, bState
    End If
End Property
Private Function pbGetState(ByVal iIDBtn As Long, ByVal fStateFlag As ectbButtonStates) As Boolean
Dim fState As Long
    fState = SendMessageLong(m_hWndToolBar, TB_GETSTATE, iIDBtn, 0)
    pbGetState = ((fState And fStateFlag) = fStateFlag)
End Function
Private Function pbSetState(ByVal iIDBtn As Long, ByVal fStateFlag As ectbButtonStates, ByVal bState As Boolean)
Dim fState As Long
    fState = SendMessageLong(m_hWndToolBar, TB_GETSTATE, iIDBtn, 0)
    If (bState) Then
        fState = fState Or fStateFlag
    Else
        fState = fState And Not fStateFlag
    End If
    If (SendMessageLong(m_hWndToolBar, TB_SETSTATE, iIDBtn, fState) = 0) Then
        debugmsg "Button state failed"
    Else
        pbSetState = True
    End If
End Function
 
Public Property Get hwnd() As Long
Attribute hwnd.VB_Description = "Returns the window handle of the control."
    hwnd = m_hWndToolBar
End Property

Public Sub DestroyToolBar()
Attribute DestroyToolBar.VB_Description = "Destroys the toolbar and all resources associated with it."
On Error Resume Next
'We need to clean up our windows
    pSubclass False
    If (m_hWndToolBar <> 0) Then
        ShowWindow m_hWndToolBar, SW_HIDE
        SetParent m_hWndToolBar, 0
        DestroyWindow (m_hWndToolBar)
        m_hWndToolBar = 0
    End If
   m_hWndParentForm = 0
End Sub
Public Sub CreateFromMenu( _
      ByRef cMenu As Object _
   )
Attribute CreateFromMenu.VB_Description = "Sets up a toolbar based on a cPopupMenu object so the toolbar can act as the form's menu."
Dim i As Long
Dim hSubMenu As Long
Dim sCaption As String
Dim iPos As Long
Dim bEnabled As Boolean
Dim sKey As String
Dim iB As Long, iFB As Long
   
   iFB = -1
   
   If m_hWndToolBar = 0 Then
      CreateToolbar , True, True, True, 0
   Else
      ' remove all buttons:
      For i = 0 To ButtonCount - 1
         RemoveButton i
      Next i
   End If
   
   ' Now add buttons according to menu:
   With cMenu
      If .Count > 0 Then
         m_hMenu = .hMenu(1)
         For i = 1 To .Count
            If .hMenu(i) = m_hMenu Then
               sCaption = .Caption(i)
               bEnabled = .Enabled(i)
               sKey = .ItemKey(i)
               ' assume here that next menu along
               ' is the sub menu.  Fix required!!!
               hSubMenu = GetSubMenu(m_hMenu, iPos)
               ' Add the button:
               'Debug.Print sCaption, bEnabled, hSubMenu
               AddButton , , , , sCaption, CTBAutoSize, sKey
               iB = ButtonCount - 1
               If iB > -1 Then
                  If iFB = -1 Then
                     iFB = iB
                  End If
                  m_tBInfo(iB).hSubMenu = hSubMenu
                  If Not bEnabled Then
                     ButtonEnabled(iB) = False
                  End If
                  iPos = iPos + 1
               End If
            End If
         Next i
      End If
   End With
   m_lPtrMenu = ObjPtr(cMenu)
   
End Sub

Public Sub CreateToolbar( _
      Optional ButtonSize As Integer = 16, _
      Optional StyleList As Boolean, _
      Optional WithText As Boolean, _
      Optional Wrappable As Boolean, _
      Optional PicSize As Integer)
Attribute CreateToolbar.VB_Description = "Initialises a toolbar for use."
On Error Resume Next
Dim Button As TBBUTTON
Dim lParam As Long
Dim ListButtons As Boolean
Dim dwStyle As Long

   DestroyToolBar

   m_bWrappable = Wrappable
   m_bWithText = WithText

   dwStyle = WS_CHILD Or WS_VISIBLE Or WS_CLIPCHILDREN
   dwStyle = dwStyle Or CCS_NOPARENTALIGN Or CCS_NORESIZE Or CCS_ADJUSTABLE Or CCS_NODIVIDER
   dwStyle = dwStyle Or TBSTYLE_TOOLTIPS Or TBSTYLE_FLAT
   If (StyleList) Then
      dwStyle = dwStyle Or TBSTYLE_LIST
   End If
   If (Wrappable) Then
      dwStyle = dwStyle Or TBSTYLE_WRAPABLE
   End If

    m_hWndToolBar = CreateWindowEX(0, "ToolbarWindow32", "", _
         dwStyle, _
         0, 0, 0, 0, UserControl.Parent.hwnd, 0&, App.hInstance, 0&)
  
    SendMessageLong m_hWndToolBar, TB_SETPARENT, UserControl.Parent.hwnd, 0
  
    m_lR = SendMessageLong(m_hWndToolBar, TB_BUTTONSTRUCTSIZE, LenB(Button), 0)
     
   AddBitmapIfRequired
   If m_eImageSourceType <> -1 Then
      lParam = ButtonSize + (ButtonSize * &H10000)
   Else
      lParam = 0
   End If
   m_lR = SendMessageLong(m_hWndToolBar, TB_SETBITMAPSIZE, 0, lParam)
     
   Set m_pic = Nothing

   SetProp m_hWndToolBar, "vbalTbar:ControlPtr", ObjPtr(Me)
   m_hWndParentForm = UserControl.Parent.hwnd
   If TypeOf UserControl.Parent Is MDIForm Then
      SetProp m_hWndToolBar, "vbalTbar:MDIClient", GetWindow(m_hWndParentForm, GW_CHILD)
   End If
   pSubclass True, m_hWndParentForm
   AddToToolTip m_hWndToolBar
    
End Sub
Public Property Get ListStyle() As Boolean
   ListStyle = pbIsStyle(TBSTYLE_LIST)
End Property
Public Property Let ListStyle(ByVal bState As Boolean)
   pbSetStyle TBSTYLE_LIST, bState
End Property
Public Property Get Wrappable() As Boolean
   Wrappable = pbIsStyle(TBSTYLE_WRAPABLE)
End Property
Public Property Let Wrappable(ByVal bState As Boolean)
   pbSetStyle TBSTYLE_WRAPABLE, bState
End Property
Private Function pbSetStyle(ByVal lStyleBit As Long, ByVal bState As Boolean) As Boolean
Dim lS As Long
Dim iB As Long
   If Not pbIsStyle(lStyleBit) = bState Then
      lS = GetWindowLong(m_hWndToolBar, GWL_STYLE)
      If bState Then
         lS = lS Or lStyleBit
      Else
         lS = lS And Not lStyleBit
      End If
      SetWindowLong m_hWndToolBar, GWL_STYLE, lS
      If bState Then
         For iB = 0 To ButtonCount - 1
            ButtonTextVisible(iB) = Not (ButtonTextVisible(iB))
            ButtonTextVisible(iB) = Not (ButtonTextVisible(iB))
         Next iB
      Else
         If ButtonCount > 0 Then
            ButtonTextVisible(0) = Not (ButtonTextVisible(0))
            ButtonTextVisible(0) = Not (ButtonTextVisible(0))
         End If
      End If
      ResizeToolbar
   End If
End Function
Private Function pbIsStyle(ByVal lStyleBit As Long) As Boolean
Dim lS As Long
   If m_hWndToolBar <> 0 Then
      lS = GetWindowLong(m_hWndToolBar, GWL_STYLE)
      If (lS And lStyleBit) = lStyleBit Then
         pbIsStyle = True
      End If
   End If
End Function
Public Property Let ImageSource( _
        ByVal eType As ECTBImageSourceTypes _
    )
Attribute ImageSource.VB_Description = "Sets the type of image (file, picture, resource, image list or standard image list) to be used as the source of the button's images."
    m_eImageSourceType = eType
End Property
Public Property Let ImageResourceID(ByVal lResourceId As Long)
Attribute ImageResourceID.VB_Description = "Sets a resource id to be used as the source of the button's images."
    m_lResourceID = lResourceId
End Property
Public Property Let ImageResourcehInstance(ByVal hInstance As Long)
Attribute ImageResourcehInstance.VB_Description = "Sets the hInstance of the binary containing the resource specified in ImageResourceID."
   m_hInstance = hInstance
End Property
Public Property Let ImageFile(ByVal sFile As String)
Attribute ImageFile.VB_Description = "Sets a bitmap file to be used as the source of the buttons images."
    m_sFileName = sFile
End Property
Public Sub SetImageList( _
      ByVal vThis As Variant, _
      Optional ByVal eType As ECTBImageListTypes = CTBImageListNormal _
   )
Attribute SetImageList.VB_Description = "Sets the image list to be used for standard, hot or disabled button images."
Dim hIml As Long
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
      hIml = vThis.hImageList
      If (Err.Number <> 0) Then
         Err.Clear
         hIml = vThis.hIml
         If Err.Number <> 0 Then
             hIml = 0
         End If
       End If
       On Error GoTo 0
   ElseIf VarType(vThis) = vbLong Then
       ' Assume ImageList handle:
       hIml = vThis
   Else
       Err.Raise vbObjectError + 1049, "cToolbar." & App.EXEName, "ImageList property expects ImageList object or long hImageList handle."
   End If
    
   ' If we have a valid image list, then associate it with the control:
   Select Case eType
   Case CTBImageListDisabled
      m_hImlDis = hIml
   Case CTBImageListHot
      m_hImlHot = hIml
   Case CTBImageListNormal
      m_hIml = hIml
   End Select
   
   If m_hWndToolBar <> 0 Then
      If (hIml <> 0) Then
         m_lR = SendMessageLong(m_hWndToolBar, eType, 0, hIml)
      End If
   End If
      
End Sub
Public Property Let ImagePicture(ByVal picThis As Long)
Attribute ImagePicture.VB_Description = "Sets a picture object to be used as the source of the button's images."
   'was StdPicture
   'Set m_pic = picThis
   m_hBmp = picThis
End Property
Public Property Let ImageStandardBitmapType(ByVal eType As ECTBStandardImageSourceTypes)
Attribute ImageStandardBitmapType.VB_Description = "Sets the standard image list bitmap to be used to generate the button images."
   m_eStandardType = eType
End Property


Private Sub AddBitmapIfRequired()
Dim tbab As TBADDBITMAP
    
   Select Case m_eImageSourceType
   Case CTBStandardImageSources
      SendMessageLong m_hWndToolBar, TB_LOADIMAGES, m_eStandardType, HINST_COMMCTRL
   Case CTBPicture
      tbab.hInst = 0
      tbab.nID = m_hBmp  'm_pic.Handle
      ' Add the bitmap containing button images to the toolbar.
      m_lR = SendMessage(m_hWndToolBar, TB_ADDBITMAP, 54, tbab)
   Case CTBLoadFromFile
      tbab.hInst = 0
     tbab.nID = LoadImage(0, m_sFileName, IMAGE_BITMAP, 0, 0, _
                LR_LOADFROMFILE Or LR_LOADMAP3DCOLORS Or LR_LOADTRANSPARENT)
     ' tbab.nID = LoadImage(0, m_sFileName, IMAGE_BITMAP, 0, 0, _
     '            LR_LOADFROMFILE Or LR_LOADTRANSPARENT)
      m_lR = SendMessage(m_hWndToolBar, TB_ADDBITMAP, 54, tbab)
   Case CTBResourceBitmap
      tbab.hInst = 0
      tbab.nID = LoadImageLong(m_hInstance, m_lResourceID, IMAGE_BITMAP, 0, 0, _
                   LR_LOADMAP3DCOLORS Or LR_LOADTRANSPARENT)
     ' tbab.nID = LoadImageLong(m_hInstance, m_lResourceID, IMAGE_BITMAP, 0, 0, _
     '               LR_LOADTRANSPARENT)
      m_lR = SendMessage(m_hWndToolBar, TB_ADDBITMAP, 54, tbab)
   Case CTBExternalImageList
      If m_hIml <> 0 Then
         SendMessageLong m_hWndToolBar, CTBImageListNormal, 0, m_hIml
      End If
      If m_hImlHot <> 0 Then
         SendMessageLong m_hWndToolBar, CTBImageListHot, 0, m_hImlHot
      End If
      If m_hImlDis <> 0 Then
         SendMessageLong m_hWndToolBar, CTBImageListDisabled, 0, m_hImlDis
      End If
   End Select
    
End Sub

Public Sub RemoveButton(ByVal vButton As Variant)
Attribute RemoveButton.VB_Description = "Removes a button from the toolbar."
Dim iB As Integer
Dim iCount As Long
Dim iNewCount As Long
Dim i As Long
Dim iT As Long
Dim sCaption As String
   
   iB = ButtonIndex(vButton)
   If (iB > -1) Then
      iCount = ButtonCount
      
      If iCount <= 0 Then
         Debug.Assert iCount > 0
      Else
         sCaption = m_tBInfo(iB).sCaption
         m_lR = SendMessageLong(m_hWndToolBar, TB_DELETEBUTTON, iB, 0)
         iNewCount = ButtonCount
         If iNewCount = 0 Then
            Erase m_tBInfo
         Else
            For i = 0 To iNewCount - 1
               If i >= iB Then
                  LSet m_tBInfo(i) = m_tBInfo(i + 1)
               End If
            Next i
            ReDim Preserve m_tBInfo(0 To iNewCount - 1) As ButtonInfoStore
         End If
         plRemoveString sCaption
      End If
   End If
   
End Sub

Public Sub AddButton( _
        Optional ByVal sTip As String = "", _
        Optional ByVal iImage As Integer = -1, _
        Optional ByVal vButtonBefore As Variant, _
        Optional ByVal xLarge As Integer = 0, _
        Optional ByVal sButtonText As String, _
        Optional ByVal eButtonStyle As ECTBToolButtonSyle, _
        Optional ByVal sKey As String = "" _
    )
Attribute AddButton.VB_Description = "Adds or inserts a button to the toolbar."
Dim tB As TBBUTTON
Dim lParam As Long
Dim iB As Integer, i As Integer
Dim bInsert As Boolean
Dim iCount As Long
Dim idString As Long

   iCount = ButtonCount
   If iCount = 0 Then
      ' Make sure we can have drop-down buttons:
      SendMessageLong m_hWndToolBar, TB_SETEXTENDEDSTYLE, 0, TBSTYLE_EX_DRAWDDARROWS
   End If

   ' Are we adding or inserting?
   If Not (IsMissing(vButtonBefore)) Then
      iB = ButtonIndex(vButtonBefore)
      If (iB > -1) Then
         bInsert = True
      End If
   End If
     
   ' Do we need to add a new string for this button?
   idString = -1
   If Len(sButtonText) > 0 Then
      idString = plAddStringIfRequired(sButtonText)
   End If
 
   With tB
      .iBitmap = iImage
      .idCommand = NewButtonID
      .fsState = TBSTATE_ENABLED
      .fsStyle = eButtonStyle
      .dwData = 0
      .iString = idString
   End With
   
   If (bInsert) Then
      m_lR = SendMessage(m_hWndToolBar, TB_INSERTBUTTON, iB, tB)
      If (m_lR <> 0) Then
         ' We need to insert into the structure:
         ReDim Preserve m_tBInfo(0 To iCount) As ButtonInfoStore
         For i = iCount To iB + 1 Step -1
            LSet m_tBInfo(i) = m_tBInfo(i - 1)
         Next i
         With m_tBInfo(iB)
            .wID = tB.idCommand
            .iImage = iImage
            .sTipText = sTip
            .iLarge = xLarge
            .sKey = sKey
            .bShowText = m_bWithText
            .sCaption = sButtonText
            .eStyle = eButtonStyle
         End With
      End If
   Else
      m_lR = SendMessage(m_hWndToolBar, TB_ADDBUTTONS, 1, tB)
      If (m_lR <> 0) Then
         ' Add this button to the list:
         ReDim Preserve m_tBInfo(0 To iCount) As ButtonInfoStore
         With m_tBInfo(iCount)
            .wID = tB.idCommand
            .iImage = iImage
            .sTipText = sTip
            .iLarge = xLarge
            .sKey = sKey
            .bShowText = m_bWithText
            .sCaption = sButtonText
            .eStyle = eButtonStyle
         End With
      End If
   End If
   
   ' Size window:
   ResizeToolbar
    
End Sub
Private Function plAddStringIfRequired(ByVal sString As String) As Long
Dim id As Long
Dim i As Long
Dim b() As Byte
Dim sAccel As String

   ' Signal default:
   id = -1
   
   ' Check if we already have the string - if we do, then use that
   For i = 1 To m_lStringIDCount
      If (m_sString(i) = sString) Then
         id = m_lStringID(i)
         Exit For
      End If
   Next i
   
   ' If string not found, then add one:
   If (id = -1) Then
      b = StrConv(sString, vbFromUnicode)
      i = UBound(b) + 2
      ReDim Preserve b(0 To i) As Byte
      b(i - 1) = 0
      b(i) = 0
      
      id = SendMessage(m_hWndToolBar, TB_ADDSTRING, 0, b(0))
      
      m_lStringIDCount = m_lStringIDCount + 1
      ReDim Preserve m_sString(1 To m_lStringIDCount) As String
      ReDim Preserve m_lStringID(1 To m_lStringIDCount) As Long
      m_sString(m_lStringIDCount) = sString
      m_lStringID(m_lStringIDCount) = id

   End If
   
   ' Return the Id:
   plAddStringIfRequired = id
   
End Function
Private Function psGetAccelerator(ByVal sString As String) As String
Dim iPos As Long
   iPos = InStr(sString, "&")
   If iPos <> 0 And iPos <> InStr(sString, "&&") Then
      If iPos < Len(sString) Then
         psGetAccelerator = UCase$(Mid$(sString, iPos + 1, 1))
      End If
   End If
End Function
Private Function plRemoveString(ByVal sCaption As String)
   ' unfortunately you cannot remove a string
   ' from the toolbar itself (because, as MSJ puts it,
   ' ".. the toolbar is braindead ..")
   
End Function
Public Sub ResizeToolbar()
Attribute ResizeToolbar.VB_Description = "Resizes the toolbar."
Dim tR As RECT, tPR As RECT, tCR As RECT
Dim tp As POINTAPI
Dim lCount As Long
Dim i As Long
Dim Button As TBBUTTON
Dim lW As Long, lH As Long

   ' Get number of buttons:
   lCount = SendMessageLong(m_hWndToolBar, TB_BUTTONCOUNT, 0, 0)
   If (lCount > 0) Then
      ' Get the total length:
      lW = ToolbarWidth
      lH = ToolbarHeight
      
      ' Get rectangle for toolbar.  Unfortunately the rebar doesn't
      ' seem to like ClientToScreen and gives the wrong answer!  So
      ' do it manually:
      GetWindowRect m_hWndToolBar, tR
      GetWindowRect GetParent(m_hWndToolBar), tPR
      GetClientRect GetParent(m_hWndToolBar), tCR
      
      'Debug.Print tR.Top, tPR.Top, tCR.Top
      tp.x = tR.Left - tPR.Left - 2
      tp.y = tR.Top - tPR.Top - 2
      
      ' Make window correct size:
      If (m_bWrappable) Then
         SetWindowPos m_hWndToolBar, 0, tp.x, tp.y, lW, lH, SWP_FRAMECHANGED Or SWP_NOMOVE Or SWP_NOOWNERZORDER
      Else
         SetWindowPos m_hWndToolBar, 0, tp.x, tp.y, lW, lH, SWP_FRAMECHANGED Or SWP_NOMOVE Or SWP_NOOWNERZORDER
      End If
    End If
End Sub
Public Property Get ToolbarWidth() As Long
Attribute ToolbarWidth.VB_Description = "Gets the width of the toolbar."
Dim lSize As Long
Dim lCount As Long
Dim lWidth As Long
Dim i As Long
Dim rc As RECT

   ' Get number of buttons:
   lCount = SendMessageLong(m_hWndToolBar, TB_BUTTONCOUNT, 0, 0)
   If (lCount > 0) Then
      ' Get the total length:
      For i = 0 To lCount - 1
         If (ButtonVisible(i)) Then
            SendMessage m_hWndToolBar, TB_GETITEMRECT, i, rc
            lSize = lSize + rc.Right - rc.Left
         End If
      Next i
      ToolbarWidth = lSize
   End If
End Property
Public Property Get ToolbarHeight() As Long
Attribute ToolbarHeight.VB_Description = "Gets the height of the toolbar."
Dim lSize As Long
Dim lCount As Long
Dim i As Long
Dim rc As RECT
   ' Get number of buttons:
   lCount = SendMessageLong(m_hWndToolBar, TB_BUTTONCOUNT, 0, 0)
   If (lCount > 0) Then
      ' Get the height:
      i = 0
      Do While ButtonVisible(i) = False
         i = i + 1
         If i >= lCount Then
            Exit Do
         End If
      Loop
      SendMessage m_hWndToolBar, TB_GETITEMRECT, i, rc
      ToolbarHeight = rc.Bottom
   End If
End Property


Public Sub ButtonSize(xWidth As Integer, xHeight As Integer)
Attribute ButtonSize.VB_Description = "Gets the rectangle of a button."
   m_iButtonWidth = xWidth
   m_iButtonHeight = xHeight
   SendMessageLong m_hWndToolBar, TB_AUTOSIZE, 0, 0
End Sub
Public Sub GetDropDownPosition( _
        ByVal id As Integer, _
        ByRef x As Long, _
        ByRef y As Long _
    )
Attribute GetDropDownPosition.VB_Description = "Returns the position to show a drop-down menu for a button in response to the DropDownPress event."
Dim rc As RECT
Dim tp As POINTAPI
    
    SendMessage m_hWndToolBar, TB_GETITEMRECT, id, rc
    tp.x = rc.Left
    tp.y = rc.Bottom
    MapWindowPoints m_hWndToolBar, m_hWndParentForm, tp, 1
    x = tp.x * Screen.TwipsPerPixelX
    y = tp.y * Screen.TwipsPerPixelY
    
End Sub

Private Sub pInitialise()
Dim tICCEX As CommonControlsEx

   If Not (UserControl.Ambient.UserMode) Then
     ' We are in design mode:
     lblInfo.Caption = "Toolbar Control: " & UserControl.Extender.Name
   Else
      UserControl.BorderStyle() = 0
      lblInfo.Visible = False
      UserControl.Extender.Left = -UserControl.Width * 2
      ' We are in run
      With tICCEX
          .dwSize = LenB(tICCEX)
          .dwICC = ICC_BAR_CLASSES
      End With
      'We need to make this call to make sure the common controls are loaded
      InitCommonControlsEx tICCEX
      m_hWndToolBar = 0
      ' Start checking for accelerator key presses here:
      AttachKeyboardHook Me
   End If
   
End Sub
Private Sub pSubclass(ByVal bState As Boolean, Optional ByVal lhWnd As Long = 0)
Static s_lhWndSave As Long

    If (m_bInSubClass <> bState) Then
        If (bState) Then
            'Debug.Print "Subclassing:Start"
            Debug.Assert (lhWnd <> 0)
            If (s_lhWndSave <> 0) Then
                pSubclass False
            End If
            s_lhWndSave = lhWnd
            pAttMsg lhWnd, WM_COMMAND
            pAttMsg lhWnd, WM_MOUSEMOVE
            pAttMsg lhWnd, WM_LBUTTONDOWN
            pAttMsg lhWnd, WM_LBUTTONUP
            pAttMsg lhWnd, WM_RBUTTONDOWN
            pAttMsg lhWnd, WM_RBUTTONUP
            pAttMsg lhWnd, WM_MBUTTONDOWN
            pAttMsg lhWnd, WM_MBUTTONUP
            pAttMsg lhWnd, WM_NOTIFY
            s_lhWndSave = lhWnd
            m_bInSubClass = True
        Else
            'Debug.Print "Subclassing:End"
            pDelMsg s_lhWndSave, WM_COMMAND
            pDelMsg s_lhWndSave, WM_MOUSEMOVE
            pDelMsg s_lhWndSave, WM_LBUTTONDOWN
            pDelMsg s_lhWndSave, WM_LBUTTONUP
            pDelMsg s_lhWndSave, WM_RBUTTONDOWN
            pDelMsg s_lhWndSave, WM_RBUTTONUP
            pDelMsg s_lhWndSave, WM_MBUTTONDOWN
            pDelMsg s_lhWndSave, WM_MBUTTONUP
            pDelMsg s_lhWndSave, WM_NOTIFY
            s_lhWndSave = 0
            m_bInSubClass = False
        End If
    End If
End Sub
Private Sub pTerminate()
   ' Remove from tooltip:
   RemoveFromToolTip m_hWndToolBar
    ' Clear up hook:
    DetachKeyboardHook Me
    ' Clear toolbar window:
   DestroyToolBar
   ' Background picture -> nothing if any:
   Set m_pic = Nothing
End Sub
Private Sub pAttMsg(ByVal lhWnd As Long, ByVal lMsg As Long)
    AttachMessage Me, lhWnd, lMsg
End Sub
Private Sub pDelMsg(ByVal lhWnd As Long, ByVal lMsg As Long)
    DetachMessage Me, lhWnd, lMsg
End Sub

Public Function RaiseButtonClick(ByVal iIDButton As Long)
Attribute RaiseButtonClick.VB_Description = "Causes a button click to occur."
   ' Required as part of the WM_COMMAND handler:
   RaiseEvent ButtonClick(iIDButton)
End Function

Private Property Let ISubClass_MsgResponse(ByVal RHS As EMsgResponse)
   '
End Property

Private Property Get ISubClass_MsgResponse() As EMsgResponse
   Select Case CurrentMessage
   Case WM_COMMAND, WM_MOUSEMOVE, WM_LBUTTONDOWN, WM_LBUTTONUP, WM_RBUTTONDOWN, WM_RBUTTONUP, WM_MBUTTONDOWN, WM_MBUTTONUP, WM_NOTIFY
      ISubClass_MsgResponse = emrPostProcess
   End Select
End Property

Private Function ISubClass_WindowProc(ByVal hwnd As Long, ByVal iMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Dim msgStruct As Msg
Dim hdr As NMHDR
Dim ttt As ToolTipText
Dim pt32 As POINTAPI
Dim ptX As Long
Dim ptY As Long
Dim hWndOver As Long
Dim b() As Byte
Dim iB As Long
Dim lPtr As Long
Dim iOld As Long, iNew As Long
Dim eReason As ECTBHotItemChangeReasonConstants
Dim bS As Boolean
  
On Error Resume Next

   Select Case iMsg
   Case WM_COMMAND
      If (lParam = m_hWndToolBar) Then
         'Debug.Print wParam, lParam
         iB = SendMessageLong(m_hWndToolBar, TB_COMMANDTOINDEX, wParam, 0)
         If iB > -1 Then
            If m_tBInfo(iB).hSubMenu <> 0 Then
               bS = ButtonPressed(iB)
               ButtonPressed(iB) = True
               ' First tell the client we're about to show the menu
               RaiseEvent ButtonClick(iB)
               ' Now show the menu:
               pMenuClick iB
               ButtonPressed(iB) = False
               ISubClass_WindowProc = 0
               SendMessageLong m_hWndParentForm, WM_EXITMENULOOP, 0, 0
            Else
               bS = ButtonPressed(iB)
               ButtonPressed(iB) = True
               RaiseEvent ButtonClick(iB)
               ButtonPressed(iB) = bS
               ISubClass_WindowProc = 0
            End If
         End If
      End If
   
   Case WM_MOUSEMOVE, WM_LBUTTONDOWN, WM_LBUTTONUP, WM_RBUTTONDOWN, WM_RBUTTONUP, WM_MBUTTONDOWN, WM_MBUTTONUP
      With msgStruct
         .lParam = lParam
         .wParam = wParam
         .message = iMsg
         .hwnd = hwnd
      End With
      
      'Pass the structure
      SendMessage hwndToolTip, TTM_RELAYEVENT, 0, msgStruct
      
   
   Case WM_NOTIFY
      CopyMemory hdr, ByVal lParam, Len(hdr)
         
      Select Case hdr.code
      Case TTN_NEEDTEXT
         Dim idNum As Integer
         idNum = hdr.idfrom
         On Error Resume Next
         
         iB = pbGetIndexForID(idNum)
         If (iB > -1) Then
            msToolTipBuffer = StrConv(ButtonToolTip(iB), vbFromUnicode)
            If Err.Number = 0 Then
               If (Len(msToolTipBuffer) > 0) Then
                  msToolTipBuffer = msToolTipBuffer & vbNullChar
                  ' Debug.Print "Show tool tip", ButtonToolTip(iB)
                  CopyMemory ttt, ByVal lParam, Len(ttt)
                  ttt.lpszText = StrPtr(msToolTipBuffer)
                  CopyMemory ByVal lParam, ttt, Len(ttt)
               End If
            Else
               Err.Clear
            End If
         End If
         
      Case TBN_DROPDOWN
         If (hdr.hwndFrom = m_hWndToolBar) Then
            Dim nmTB As NMTOOLBAR_SHORT
            CopyMemory nmTB, ByVal lParam, Len(nmTB)
            iB = SendMessageLong(m_hWndToolBar, TB_COMMANDTOINDEX, nmTB.iItem, 0)
            RaiseEvent DropDownPress(iB)
         End If
         
      Case TBN_HOTITEMCHANGE
         If (hdr.hwndFrom = m_hWndToolBar) Then
            If m_lMajorVer > 4 Or (m_lMajorVer = 4 And m_lMinorVer >= 70) Then
               Dim nmTBHI As NMTBHOTITEM
               CopyMemory nmTBHI, ByVal lParam, Len(nmTBHI)
               eReason = nmTBHI.dwFlags
               iOld = -1: iNew = -1
               If (eReason And HICF_ENTERING) <> HICF_ENTERING Then
                  iOld = SendMessageLong(m_hWndToolBar, TB_COMMANDTOINDEX, nmTBHI.idOld, 0)
               End If
               If (eReason And HICF_LEAVING) <> HICF_LEAVING Then
                  iNew = SendMessageLong(m_hWndToolBar, TB_COMMANDTOINDEX, nmTBHI.idNew, 0)
               End If
               RaiseEvent HotItemChange(iNew, iOld, eReason)
            End If
         End If
         
      End Select
      
   End Select
    
End Function


Private Sub UserControl_Initialize()
   debugmsg "cToolbar:Initialize"
   If Not (ComCtlVersion(m_lMajorVer, m_lMinorVer, m_lBuild)) Then
      m_lMajorVer = 4
      m_lMinorVer = 0
      m_lBuild = 0
   End If
   m_eImageSourceType = -1
End Sub

Private Sub UserControl_InitProperties()
    ' Initialise the control
    pInitialise
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    ' Read properties:
    
    ' Initialise the control
    pInitialise
End Sub

Private Sub UserControl_Terminate()
    pTerminate
    Set m_cMenu = Nothing
    debugmsg "cToolbar:Terminate"
    'MsgBox "cToolbar:Terminate"
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    ' Write properties:
End Sub
