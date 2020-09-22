VERSION 5.00
Begin VB.UserControl cReBar 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   525
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4905
   InvisibleAtRuntime=   -1  'True
   ScaleHeight     =   525
   ScaleWidth      =   4905
   ToolboxBitmap   =   "cReBar.ctx":0000
   Begin VB.Label lblRebar 
      Caption         =   "'Rebar Control'"
      Height          =   315
      Left            =   60
      TabIndex        =   0
      Top             =   120
      Width           =   3255
   End
End
Attribute VB_Name = "cReBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' =========================================================================
' vbAccelerator Rebar control v2.0
' Copyright © 1998-1999 Steve McMahon (steve@dogma.demon.co.uk)
'
' This is a complete rebar implementation.
'
' -------------------------------------------------------------------------
' Visit vbAccelerator at http://vbaccelerator.com
' =========================================================================

' ==============================================================================
' Declares, constants and types required for toolbar:
' ==============================================================================
Private Type NMREBAR
    hdr As NMHDR
    dwMask As Long
    uBand As Long
    fStyle As Long
    wID As Long
    lParam As Long
End Type
Private Type NMRBAUTOSIZE
    hdr As NMHDR
    fChanged As Long
    rcTarget As RECT
    rcActual As RECT
End Type
Private Type NMREBARCHILDSIZE
    hdr As NMHDR
    uBand As Long
    wID As Long
    rcChild As RECT
    rcBand As RECT
End Type
Private Type NMREBARCHEVRON
   hdr As NMHDR
   uBand As Long
   wID As Long
   lParam As Long
   rcChevron As RECT
End Type
'Private Type REBARINFO
'    cbSize As Integer
'    fMask As Integer
'    hIml As Long
'End Type
Private Type REBARBANDINFO
    cbSize As Long
    fMask As Long
    fStyle As Long
    clrFore As Long
    clrBack As Long
    lpText As String
    cch As Long
    iImage As Long
    hWndChild As Long
    cxMinChild As Long
    cyMinChild As Long
    cx As Long
    hbmBack As Long
    wID As Long
End Type
Private Type REBARBANDINFO_471
    cbSize As Long
    fMask As Long
    fStyle As Long
    clrFore As Long
    clrBack As Long
    lpText As String
    cch As Long
    iImage As Integer 'Image
    hWndChild As Long
    cxMinChild As Long
    cyMinChild As Long
    cx As Long
    hbmBack As Long 'hBitmap
    wID As Long
    cyChild As Long
    cyMaxChild As Long
    cyIntegral As Long
    cxIdeal As Long
    lParam As Long
    cxHeader As Long
End Type
Private Type REBARBANDINFO_NOTEXT
    cbSize As Long
    fMask As Long
    fStyle As Long
    clrFore As Long
    clrBack As Long
    lpText As Long
    cch As Long
    iImage As Integer 'Image
    hWndChild As Long
    cxMinChild As Long
    cyMinChild As Long
    cx As Long
    hbmBack As Long 'hBitmap
    wID As Long
End Type
Private Type REBARBANDINFO_NOTEXT_471
    cbSize As Long
    fMask As Long
    fStyle As Long
    clrFore As Long
    clrBack As Long
    lpText As Long
    cch As Long
    iImage As Integer 'Image
    hWndChild As Long
    cxMinChild As Long
    cyMinChild As Long
    cx As Long
    hbmBack As Long 'hBitmap
    wID As Long
    cyChild As Long
    cyMaxChild As Long
    cyIntegral As Long
    cxIdeal As Long
    lParam As Long
    cxHeader As Long
End Type

Private Const REBARCLASSNAME = "ReBarWindow32"

'Rebar Styles
'Private Const RBS_TOOLTIPS = &H100&
Private Const RBS_VARHEIGHT = &H200&
Private Const RBS_BANDBORDERS = &H400&
'Private Const RBS_FIXEDORDER = &H800&
Private Const RBS_AUTOSIZE = &H2000&
'Private Const RBS_VERTICALGRIPPER = &H4000& '  // this always has the vertical gripper (default for horizontal mode)
Private Const RBS_DBLCLKTOGGLE = &H8000&

Private Const RBBS_BREAK = &H1               ' break to new line
Private Const RBBS_FIXEDSIZE = &H2           ' band can't be sized
Private Const RBBS_CHILDEDGE = &H4           ' edge around top & bottom of child window
Private Const RBBS_HIDDEN = &H8              ' don't show
'Private Const RBBS_NOVERT = &H10             ' don't show when vertical
Private Const RBBS_FIXEDBMP = &H20           ' bitmap doesn't move during band resize
'Private Const RBBS_VARIABLEHEIGHT = &H40
'Private Const RBBS_GRIPPERALWAYS = &H80      ' always show the gripper
Private Const RBBS_NOGRIPPER = &H100 '// never show the gripper
Private Const RBBS_CHEVRON = &H200& ' // If you set cxIdeal, version 5.00 only...

Private Const RBS_EX_OFFICE9 = &H1&     '// new gripper, chevron, focus handling

Private Const RBBIM_STYLE = &H1
Private Const RBBIM_COLORS = &H2
Private Const RBBIM_TEXT = &H4
'Private Const RBBIM_IMAGE = &H8
Private Const RBBIM_CHILD = &H10
Private Const RBBIM_CHILDSIZE = &H20
Private Const RBBIM_SIZE = &H40
Private Const RBBIM_BACKGROUND = &H80
Private Const RBBIM_ID = &H100
' 4.72 +
Private Const RBBIM_IDEALSIZE = &H200
'Private Const RBBIM_LPARAM = &H400
'Private Const RBBIM_HEADERSIZE = &H800

Private Const RB_INSERTBANDA = (WM_USER + 1)
Private Const RB_DELETEBAND = (WM_USER + 2)
'Private Const RB_GETBARINFO = (WM_USER + 3)
'Private Const RB_SETBARINFO = (WM_USER + 4)
Private Const RB_GETBANDINFO = (WM_USER + 5)
Private Const RB_SETBANDINFOA = (WM_USER + 6)
Private Const RB_SETPARENT = (WM_USER + 7)
'Private Const RB_HITTEST = (WM_USER + 8)
Private Const RB_GETRECT = (WM_USER + 9)
'Private Const RB_INSERTBANDW = (WM_USER + 10)
Private Const RB_SETBANDINFOW = (WM_USER + 11)
Private Const RB_GETBANDCOUNT = (WM_USER + 12)
'Private Const RB_GETROWCOUNT = (WM_USER + 13)
'Private Const RB_GETROWHEIGHT = (WM_USER + 14)

Private Const RB_IDTOINDEX = (WM_USER + 16)    '// wParam == id
'Private Const RB_GETTOOLTIPS = (WM_USER + 17)
'Private Const RB_SETTOOLTIPS = (WM_USER + 18)
'Private Const RB_SETBKCOLOR = (WM_USER + 19)
'Private Const RB_GETBKCOLOR = (WM_USER + 20)
'Private Const RB_SETTEXTCOLOR = (WM_USER + 21)
'Private Const RB_GETTEXTCOLOR = (WM_USER + 22)
Private Const RB_SIZETORECT = (WM_USER + 23)   '// resize the rebar/break bands and such to this rect (lparam)

'Private Const RB_BEGINDRAG = (WM_USER + 24)
'Private Const RB_ENDDRAG = (WM_USER + 25)
'Private Const RB_DRAGMOVE = (WM_USER + 26)
Private Const RB_GETBARHEIGHT = (WM_USER + 27)

Private Const RB_GETBANDINFOA = (WM_USER + 29)

Private Const RB_MINIMIZEBAND = (WM_USER + 30)
Private Const RB_MAXIMIZEBAND = (WM_USER + 31)
'Private Const RB_GETBANDBORDERS = (WM_USER + 34) '// returns in lparam = lprc the amount of edges added to band wparam

Private Const RB_SHOWBAND = (WM_USER + 35)         '// show/hide band
'Private Const RB_SETPALETTE = (WM_USER + 37)
'Private Const RB_GETPALETTE = (WM_USER + 38)
Private Const RB_MOVEBAND = (WM_USER + 39)         ' // move band

'Private Const RB_SETBANDFOCUS = (WM_USER + 40) '// (UINT) wParam == band index      lParam == TRUE/FALSE
                                        '// returns TRUE if gave band focus, else FALSE
'Private Const RB_GETBANDFOCUS = (WM_USER + 41) '// returns index of band with focus (-1 if none)
'Private Const RB_CYCLEFOCUS = (WM_USER + 42)    '// (UINT) wParam == band index      (BOOL) lParam == back/forward
                                                '// returns index of band that got focus (-1 if none)
Private Const RB_SETEXTENDEDSTYLE = (WM_USER + 43)


'Private Const RBHT_NOWHERE = &H1
'Private Const RBHT_CAPTION = &H2
'Private Const RBHT_CLIENT = &H3
'Private Const RBHT_GRABBER = &H4
'Private Const RBHT_CHEVRON = &H8

Private Const RB_INSERTBAND = RB_INSERTBANDA
Private Const RB_SETBANDINFO = RB_SETBANDINFOA
Private Const RB_GETBANDINFO471 = RB_GETBANDINFOA

Private Const RBN_FIRST = H_MAX - 831                  '// rebar
Private Const RBN_LAST = H_MAX - 859
Private Const RBN_HEIGHTCHANGE = (RBN_FIRST - 0)
Private Const RBN_GETOBJECT = (RBN_FIRST - 1)
Private Const RBN_LAYOUTCHANGED = (RBN_FIRST - 2)
Private Const RBN_AUTOSIZE = (RBN_FIRST - 3)
Private Const RBN_BEGINDRAG = (RBN_FIRST - 4)
Private Const RBN_ENDDRAG = (RBN_FIRST - 5)
Private Const RBN_DELETINGBAND = (RBN_FIRST - 6)       '// Uses NMREBAR
Private Const RBN_DELETEDBAND = (RBN_FIRST - 7)        '// Uses NMREBAR
Private Const RBN_CHILDSIZE = (RBN_FIRST - 8)
Private Const RBN_SETFOCUS = (RBN_FIRST - 9)            '// Uses NMREBAR
Private Const RBN_CHEVRONPUSHED = (RBN_FIRST - 10)
' ==============================================================================
' INTERFACE
' ==============================================================================
' Enumerations:
Public Enum ERBPositionConstants
   erbPositionTop
   erbPositionLeft
   erbPositionRight
   erbPositionBottom
End Enum
Public Enum ECRBImageSourceTypes
    CRBResourceBitmap
    CRBLoadFromFile
    CRBPicture
End Enum

' Internal Implementation:
Private m_hWnd As Long ' Rebar
Private m_hWndCtlParent As Long ' Rebar window parent
Private m_hWndMsgParent As Long ' Where messages are sent
Private m_bSubClassing As Boolean
Private m_bInTerminate As Boolean
Private m_lMajor As Long, m_lMinor As Long

' Position:
Private m_ePosition As ERBPositionConstants

' Background imaage:
Private m_sPicture As String
Private m_lResourceID As Long
Private m_hInstance As Long
'Private m_pic As StdPicture
Private m_pic As StdPicture
Private m_hBmp As Long
Private m_eImageSourceType As ECRBImageSourceTypes

' Band original location information:
Private Type tRebarWndStore
   hwndItem As Long
   hWndItemParent As Long
   tR As RECT
End Type
Private m_tWndStore() As tRebarWndStore
Private m_iWndItemCount As Integer

' Band keys:
Private Type tRebarDataStore
   wID As Long
   vData As Variant
End Type
Private m_tDataStore() As tRebarDataStore
Private m_lIDCount As Long

Private m_bVisible As Boolean

Implements ISubclass

' Events:
Public Event HeightChanged(lNewHeight As Long)
Attribute HeightChanged.VB_Description = "Raised whenever the height of the rebar changes, for example when the user moves the bands around. "
Public Event BeginBandDrag(ByVal wID As Long, ByRef bCancel As Boolean)
Attribute BeginBandDrag.VB_Description = "Raised when the user is about to start dragging a band."
Public Event EndBandDrag(ByVal wID As Long)
Attribute EndBandDrag.VB_Description = "Raised when the user has completed dragging a band within the rebar."
Public Event BandChildResize(ByVal wID As Long, ByVal lBandLeft As Long, ByVal lBandTop As Long, ByVal lBandRight As Long, ByVal lBandBottom As Long, ByRef lChildLeft As Long, ByRef lChildTop As Long, ByRef lChildRight As Long, ByRef lChildBottom As Long)
Attribute BandChildResize.VB_Description = "Raised whenever a child is resized because of a change in size of a band."
Public Event LayoutChanged()
Attribute LayoutChanged.VB_Description = "Raised whenever the layout of the rebar bands changes, due to either the rebar being resized or the user dragging the bands."
Public Event ChevronPushed(ByVal wID As Long, ByVal lLeft As Long, ByVal lTop As Long, ByVal lRight As Long, ByVal lBottom As Long)
Attribute ChevronPushed.VB_Description = "Raised when a band chevron is pressed."


Public Sub Autosize()
Attribute Autosize.VB_Description = "Attempts to automatically move the Rebar bands so they best fit the specified rectangle (in pixels relative to the rebar's container).   Not available for COMCTL32.DLL version below 4.71."
Dim lWidth As Long
Dim lHeight As Long
Dim RC As RECT, rcP As RECT
   If (m_ePosition = erbPositionBottom) Or (m_ePosition = erbPositionTop) Then
      GetWindowRect m_hWndCtlParent, rcP
      lWidth = rcP.Right - rcP.Left
      lHeight = RebarHeight
   Else
      GetWindowRect m_hWndCtlParent, rcP
      lHeight = rcP.Bottom - rcP.TOp
      lWidth = RebarWidth
   End If
   RC.Right = lWidth
   RC.Bottom = lHeight
   SendMessage m_hWnd, RB_SIZETORECT, 0, RC
End Sub

Public Property Get Position() As ERBPositionConstants
Attribute Position.VB_Description = "Gets/sets the orientation of the rebar on its container."
Attribute Position.VB_MemberFlags = "400"
   Position = m_ePosition
End Property
Public Property Let Position(ByVal ePosition As ERBPositionConstants)
Dim dwStyle As Long
Dim dwNewStyle As Long
Dim hWndP As Long
Dim RC As RECT
   If (m_ePosition <> ePosition) Then
      m_ePosition = ePosition
      
      If (m_hWnd <> 0) Then
         SetProp m_hWnd, "vbal:cRebarPosition", m_ePosition
         
         ' Move...
         dwStyle = GetWindowLong(m_hWnd, GWL_STYLE)
         dwNewStyle = dwStyle
         dwNewStyle = dwNewStyle And Not (CCS_LEFT Or CCS_TOP Or CCS_RIGHT Or CCS_BOTTOM)
         Select Case m_ePosition
         Case erbPositionTop
            dwNewStyle = dwNewStyle Or CCS_TOP
         Case erbPositionRight
            dwNewStyle = dwNewStyle Or CCS_RIGHT
         Case erbPositionLeft
            dwNewStyle = dwNewStyle Or CCS_LEFT
         Case erbPositionBottom
            dwNewStyle = dwNewStyle Or CCS_BOTTOM
         End Select
         If dwNewStyle <> dwStyle Then
            SetWindowLong m_hWnd, GWL_STYLE, dwNewStyle
         End If
         
         RebarSize
         RaiseEvent HeightChanged(RebarHeight)
         RebarSize
         
      End If
      
   End If
End Property

Private Sub pCreateSubClass()
   If Not (m_bSubClassing) Then
      If m_hWnd <> 0 Then
         m_hWndMsgParent = UserControl.Parent.hWnd
         If (m_hWndMsgParent > 0) Then
            ' Debug.Print "Subclassing window: " & m_hWndMsgParent
            AttachMessage Me, m_hWndMsgParent, WM_NOTIFY
            AttachMessage Me, m_hWndMsgParent, WM_DESTROY
            m_bSubClassing = True
         End If
         SendMessageLong m_hWnd, RB_SETPARENT, m_hWndMsgParent, 0
      End If
   End If
End Sub

Private Sub pDestroySubClass()
   If (m_bSubClassing) Then
      DetachMessage Me, m_hWndMsgParent, WM_NOTIFY
      DetachMessage Me, m_hWndMsgParent, WM_DESTROY
      m_hWndMsgParent = 0
      m_bSubClassing = False
   End If
End Sub

' Interface properties
Private Property Get ISubclass_MsgResponse() As EMsgResponse
   Select Case CurrentMessage
   Case WM_NOTIFY
      ISubclass_MsgResponse = emrPostProcess
   End Select
End Property
Private Property Let ISubclass_MsgResponse(ByVal emrA As EMsgResponse)
   '
End Property

Private Function ISubclass_WindowProc(ByVal hWnd As Long, _
                                      ByVal iMsg As Long, _
                                      ByVal wParam As Long, _
                                      ByVal lParam As Long) As Long
Dim lHeight As Long
Dim tNMH As NMHDR
Dim tNMR As NMREBAR
Dim tNMRBA As NMRBAUTOSIZE
Dim tNMRCS As NMREBARCHILDSIZE
Dim tNMRC As NMREBARCHEVRON
Dim tNMMouse As NMMOUSE
Dim tR As RECT
Dim bCancel As Boolean
Dim rcChild As RECT
Dim i As Long
Dim lHwnd As Long
Dim wID As Long
   
   ' Don't try to raise events when the control is terminating -
   ' you will crash!
   If Not (m_bInTerminate) And Not (m_hWnd = 0 Or m_hWndMsgParent = 0) Then
   
      If iMsg = WM_NOTIFY Then
         CopyMemory tNMH, ByVal lParam, Len(tNMH)
         If tNMH.hwndFrom = m_hWnd Then
         
            Select Case tNMH.code
            Case NM_NCHITTEST
               ' NC hittest.  Apparently we can return alternative HT_ values
               ' here but I cannot get it to do anything
               CopyMemory tNMMouse, ByVal lParam, Len(tNMMouse)
               ' ...
               
            Case RBN_HEIGHTCHANGE
               ' Height change notification:
               RebarSize
               lHeight = RebarHeight
               RaiseEvent HeightChanged(lHeight)
            
            Case RBN_AUTOSIZE
               ' Autosize notification, 4.71+
               If m_lMajor > 4 Or (m_lMajor = 4 And m_lMinor >= 71) Then
                  CopyMemory tNMRBA, ByVal lParam, Len(tNMRBA)
                  ' This event isn't of any use because the CCS_NORESIZE style
                  ' is set.  I do not recommend turning CCS_NORESIZE off as it
                  ' is very easy to get infinite loops during resize code without
                  ' it...
               End If
               
            Case RBN_BEGINDRAG, RBN_ENDDRAG
               ' Band dragging notifications, 4.71+
               If m_lMajor > 4 Or (m_lMajor = 4 And m_lMinor >= 71) Then
                  ' user began dragging a band:
                  CopyMemory tNMR, ByVal lParam, Len(tNMR)
                  If tNMR.uBand > -1 Then
                     If tNMH.code = RBN_BEGINDRAG Then
                        bCancel = False
                        RaiseEvent BeginBandDrag(tNMR.wID, bCancel)
                        If bCancel Then
                           ISubclass_WindowProc = 1
                        Else
                           ISubclass_WindowProc = 0
                        End If
                     Else
                        RaiseEvent EndBandDrag(tNMR.wID)
                     End If
                  Else
                     ' no band affected.
                     RaiseEvent EndBandDrag(-1)
                  End If
               End If
            
            Case RBN_CHILDSIZE
               ' Child size change notifications, 4.71+
               If m_lMajor > 4 Or (m_lMajor = 4 And m_lMinor >= 71) Then
                  ' user began dragging a band:
                  CopyMemory tNMRCS, ByVal lParam, Len(tNMRCS)
                  LSet rcChild = tNMRCS.rcChild
                  RaiseEvent BandChildResize(tNMRCS.wID, tNMRCS.rcBand.Left, tNMRCS.rcBand.TOp, tNMRCS.rcBand.Right, tNMRCS.rcBand.Bottom, rcChild.Left, rcChild.TOp, rcChild.Right, rcChild.Bottom)
                  If rcChild.Left <> tNMRCS.rcChild.Left Or rcChild.TOp <> tNMRCS.rcChild.TOp Or rcChild.Right <> tNMRCS.rcChild.Right Or rcChild.Bottom <> tNMRCS.rcChild.Bottom Then
                     LSet tNMRCS.rcChild = rcChild
                     CopyMemory ByVal lParam, tNMRCS, Len(tNMRCS)
                  End If
                  ISubclass_WindowProc = 1
               End If
            
            Case RBN_DELETEDBAND, RBN_DELETINGBAND
               ' band deletion notifications, 4.71+
               If m_lMajor > 4 Or (m_lMajor = 4 And m_lMinor >= 71) Then
                  ' A band has just been deleted:
                  CopyMemory tNMR, ByVal lParam, Len(tNMR)
                  If tNMH.code = RBN_DELETEDBAND Then
                     pRemoveID tNMR.wID
                  Else
                     lHwnd = plGetHwndOfBandChild(m_hWnd, tNMR.uBand, wID)
                     If lHwnd <> 0 Then
                        pResetParent lHwnd
                     End If
                  End If
               End If
                     
            Case RBN_LAYOUTCHANGED
               ' layout changed notification, 4.71+
               If m_lMajor > 4 Or (m_lMajor = 4 And m_lMinor >= 71) Then
                  RaiseEvent LayoutChanged
               End If
            
            Case RBN_CHEVRONPUSHED
               Debug.Print "Chevron Pushed"
               If m_lMajor >= 5 Then
                  CopyMemory tNMRC, ByVal lParam, Len(tNMRC)
                  LSet tR = tNMRC.rcChevron
                  MapWindowPoints m_hWnd, m_hWndMsgParent, tR, 2
                  tR.Left = tR.Left * Screen.TwipsPerPixelX
                  tR.TOp = tR.TOp * Screen.TwipsPerPixelY
                  tR.Right = tR.Right * Screen.TwipsPerPixelX
                  tR.Bottom = tR.Bottom * Screen.TwipsPerPixelY
                  RaiseEvent ChevronPushed(tNMRC.wID, tR.Left, tR.TOp, tR.Right, tR.Bottom)
               End If
            
            'Case Else
            '   Debug.Print tNMH.code
               
            End Select
         
         End If
      ElseIf iMsg = WM_DESTROY Then
         ' Debug.Print "GOT WM_DESTROY!"
      End If
   End If

End Function

Public Property Get BandVisible(ByVal lBand As Long) As Boolean
Attribute BandVisible.VB_Description = "Gets/sets whether a rebar band is visible or not.  Not available for COMCTL32.DLL version below 4.71."
Dim lStyle As Long
    If (lBand >= 0) And (lBand < BandCount) Then
        If (pbGetBandInfo(m_hWnd, lBand, fMask:=RBBIM_STYLE, fStyle:=lStyle)) Then
            BandVisible = ((lStyle And RBBS_HIDDEN) <> RBBS_HIDDEN)
        End If
    Else
        BandVisible = False
    End If
   
End Property
Public Property Let BandVisible(ByVal lBand As Long, ByVal bState As Boolean)
Dim lS As Long
   If (lBand >= 0) And (lBand < BandCount) Then
      lS = Abs(bState)
      SendMessageLong m_hWnd, RB_SHOWBAND, lBand, lS
   End If
End Property
Public Property Get BandChildEdge(ByVal lBand As Long) As Boolean
Attribute BandChildEdge.VB_Description = "Gets/sets whether a band draws a narrow  internal border around the child control."
Dim lStyle As Long
   If m_lMajor > 4 Or (m_lMajor = 4 And m_lMinor >= 71) Then
      If (lBand >= 0) And (lBand < BandCount) Then
          If (pbGetBandInfo(m_hWnd, lBand, fMask:=RBBIM_STYLE, fStyle:=lStyle)) Then
              BandChildEdge = ((lStyle And RBBS_CHILDEDGE) = RBBS_CHILDEDGE)
          End If
      Else
          BandChildEdge = False
      End If
   Else
      'Unsupported
   End If
   
End Property
Public Property Let BandChildEdge(ByVal lBand As Long, ByVal bState As Boolean)
Dim lStyle As Long
Dim bCurrent As Boolean
Dim tRbbi471 As REBARBANDINFO_NOTEXT_471

   If m_lMajor > 4 Or (m_lMajor = 4 And m_lMinor >= 71) Then
      If (lBand >= 0) And (lBand < BandCount) Then
         If pbGetBandInfo(m_hWnd, lBand, fMask:=RBBIM_STYLE, fStyle:=lStyle) Then
            bCurrent = ((lStyle And RBBS_CHILDEDGE) = RBBS_CHILDEDGE)
            If bState <> bCurrent Then
               If bCurrent Then
                  lStyle = lStyle And Not RBBS_CHILDEDGE
               Else
                  lStyle = lStyle Or RBBS_CHILDEDGE
               End If
               With tRbbi471
                  .cbSize = LenB(tRbbi471)
                  .fMask = RBBIM_STYLE
                  .fStyle = lStyle
               End With
               SendMessage m_hWnd, RB_SETBANDINFO, lBand, tRbbi471
            End If
         End If
      End If
   Else
      'Unsupported
   End If
End Property
Public Property Get BandGripper(ByVal lBand As Long) As Boolean
Attribute BandGripper.VB_Description = "Gets/sets whether a rebar band has a gripper or not.  (COMCTL32.DLL v5 or higher only)"
Dim lStyle As Long
   If m_lMajor > 4 Or (m_lMajor = 4 And m_lMinor >= 71) Then
      If (lBand >= 0) And (lBand < BandCount) Then
          If (pbGetBandInfo(m_hWnd, lBand, fMask:=RBBIM_STYLE, fStyle:=lStyle)) Then
              BandGripper = ((lStyle And RBBS_NOGRIPPER) <> RBBS_NOGRIPPER)
          End If
      Else
         ' IncorrectBand
      End If
   Else
      'Unsupported
   End If
End Property
Public Property Let BandGripper(ByVal lBand As Long, ByVal bState As Boolean)
Dim lStyle As Long
Dim bCurrent As Boolean
Dim tRbbi471 As REBARBANDINFO_NOTEXT_471

   If m_lMajor > 4 Or (m_lMajor = 4 And m_lMinor >= 71) Then
      If (lBand >= 0) And (lBand < BandCount) Then
         If pbGetBandInfo(m_hWnd, lBand, fMask:=RBBIM_STYLE, fStyle:=lStyle) Then
            bCurrent = ((lStyle And RBBS_NOGRIPPER) <> RBBS_NOGRIPPER)
            If bState <> bCurrent Then
               If bCurrent Then
                  lStyle = lStyle Or RBBS_NOGRIPPER
               Else
                  lStyle = lStyle And Not RBBS_NOGRIPPER
               End If
               With tRbbi471
                  .cbSize = LenB(tRbbi471)
                  .fMask = RBBIM_STYLE
                  .fStyle = lStyle
               End With
               SendMessage m_hWnd, RB_SETBANDINFO, lBand, tRbbi471
            End If
         End If
      Else
         ' IncorrectBand
      End If
   Else
      'Unsupported
   End If
End Property
Public Property Get BandChevron(ByVal lBand As Long) As Boolean
Attribute BandChevron.VB_Description = "Gets/sets whether a band will show  a chevron if it is sized too small for the contents to fit. (COMCTL32.DLL v5 or higher only)"
Dim lStyle As Long
   If m_lMajor >= 5 Then
      If (lBand >= 0) And (lBand < BandCount) Then
          If (pbGetBandInfo(m_hWnd, lBand, fMask:=RBBIM_STYLE, fStyle:=lStyle)) Then
              BandChevron = ((lStyle And RBBS_CHEVRON) = RBBS_CHEVRON)
          End If
      Else
         ' IncorrectBand
      End If
   Else
      'Unsupported
   End If
End Property
Public Property Let BandChevron(ByVal lBand As Long, ByVal bState As Boolean)
Dim lStyle As Long
Dim lCX As Long
Dim bCurrent As Boolean
Dim tRbbi471 As REBARBANDINFO_NOTEXT_471

   If m_lMajor >= 5 Then
      If (lBand >= 0) And (lBand < BandCount) Then
         If pbGetBandInfo(m_hWnd, lBand, fMask:=RBBIM_STYLE Or RBBIM_CHILDSIZE, cxMinChild:=lCX, fStyle:=lStyle) Then
            bCurrent = ((lStyle And RBBS_CHEVRON) = RBBS_CHEVRON)
            If bState <> bCurrent Then
               If bCurrent Then
                  lStyle = lStyle And Not RBBS_CHEVRON
               Else
                  lStyle = lStyle Or RBBS_CHEVRON
               End If
               With tRbbi471
                  .cbSize = LenB(tRbbi471)
                  .fMask = RBBIM_STYLE Or RBBIM_IDEALSIZE
                  .fStyle = lStyle
                  .cxIdeal = lCX
               End With
               SendMessage m_hWnd, RB_SETBANDINFO, lBand, tRbbi471
            End If
         End If
      Else
         ' IncorrectBand
      End If
   Else
      'Unsupported
   End If
End Property

Property Get BandChildMinHeight(ByVal lBand As Long) As Long
Attribute BandChildMinHeight.VB_Description = "Gets/sets the minimum height of a rebar band."
Dim cy As Long
   If (lBand >= 0) And (lBand < BandCount) Then
      If (pbGetBandInfo(m_hWnd, lBand, fMask:=RBBIM_CHILDSIZE, cyMinChild:=cy)) Then
         BandChildMinHeight = cy
      End If
   Else
      BandChildMinHeight = -1
      ' IncorrectBand
   End If
End Property
Property Let BandChildMinHeight(ByVal lBand As Long, lHeight As Long)
   If (lBand >= 0) And (lBand < BandCount) Then
      Dim tRbbi As REBARBANDINFO_NOTEXT
      Dim lR As Long
      tRbbi.fMask = RBBIM_CHILDSIZE Or RBBIM_CHILD
      tRbbi.cbSize = Len(tRbbi)
      lR = SendMessage(m_hWnd, RB_GETBANDINFO, lBand, tRbbi)
      If (lR <> 0) Then
         If (tRbbi.hWndChild <> 0) Then
            tRbbi.fMask = RBBIM_CHILDSIZE
            tRbbi.cyMinChild = lHeight
            lR = SendMessage(m_hWnd, RB_SETBANDINFOA, lBand, tRbbi)
         End If
      End If
   Else
      ' IncorrectBand
   End If
End Property
Property Get BandChildMaxHeight(ByVal lBand As Long) As Long
Attribute BandChildMaxHeight.VB_Description = "Gets/sets the maximum height a band can size to (COMCTL32.DLL v5 or higher only)"
Dim cy As Long
   If m_lMajor > 4 Or (m_lMajor = 4 And m_lMinor >= 71) Then
      If (lBand >= 0) And (lBand < BandCount) Then
         If (pbGetBandInfo(m_hWnd, lBand, fMask:=RBBIM_CHILDSIZE, cyMaxChild:=cy)) Then
            BandChildMaxHeight = cy
         End If
      Else
         BandChildMaxHeight = -1
         ' IncorrectBand
      End If
   Else
      ' Unsupported
   End If
End Property
Property Let BandChildMaxHeight(ByVal lBand As Long, lHeight As Long)
   If m_lMajor > 4 Or (m_lMajor = 4 And m_lMinor >= 71) Then
      If (lBand >= 0) And (lBand < BandCount) Then
         Dim tRbbi As REBARBANDINFO_NOTEXT_471
         Dim lR As Long
         tRbbi.fMask = RBBIM_CHILDSIZE Or RBBIM_CHILD
         tRbbi.cbSize = Len(tRbbi)
         lR = SendMessage(m_hWnd, RB_GETBANDINFO, lBand, tRbbi)
         If (lR <> 0) Then
            If (tRbbi.hWndChild <> 0) Then
               tRbbi.fMask = RBBIM_CHILDSIZE
               tRbbi.cyMaxChild = lHeight
               lR = SendMessage(m_hWnd, RB_SETBANDINFOA, lBand, tRbbi)
            End If
         End If
      Else
         ' IncorrectBand
      End If
   Else
      ' Unsupported
   End If
End Property
Property Get BandChildMinWidth(ByVal lBand As Long) As Long
Attribute BandChildMinWidth.VB_Description = "Gets/sets the minimum width of rebar band."
Dim cx As Long
   If (lBand >= 0) And (lBand < BandCount) Then
      If (pbGetBandInfo(m_hWnd, lBand, fMask:=RBBIM_CHILDSIZE, cxMinChild:=cx)) Then
         BandChildMinWidth = cx
      End If
   Else
      BandChildMinWidth = -1
      ' IncorrectBand
   End If

End Property
Property Let BandChildMinWidth(ByVal lBand As Long, lWidth As Long)
   If (lBand >= 0) And (lBand < BandCount) Then
      Dim tRbbi As REBARBANDINFO_NOTEXT
      Dim lR As Long
      Dim tR As RECT
      
      tRbbi.fMask = RBBIM_CHILDSIZE Or RBBIM_CHILD
      tRbbi.cbSize = Len(tRbbi)
      lR = SendMessage(m_hWnd, RB_GETBANDINFO, lBand, tRbbi)
      If (lR <> 0) Then
         If (tRbbi.hWndChild <> 0) Then
            tRbbi.fMask = RBBIM_CHILDSIZE
            tRbbi.cxMinChild = lWidth
            lR = SendMessage(m_hWnd, RB_SETBANDINFOA, lBand, tRbbi)
            SendMessageLong m_hWnd, RB_MINIMIZEBAND, lBand, 0
         End If
      End If
   Else
      ' IncorrectBand
   End If
End Property
Public Sub BandChildResized(ByVal lBand As Long, ByVal lWidth As Long, ByVal lHeight As Long)
   If (lBand >= 0) And (lBand < BandCount) Then
      Dim tRbbi As REBARBANDINFO_NOTEXT
      Dim lR As Long
      Dim tR As RECT
      tRbbi.fMask = RBBIM_CHILDSIZE Or RBBIM_CHILD Or RBBIM_STYLE
      tRbbi.cbSize = Len(tRbbi)
      lR = SendMessage(m_hWnd, RB_GETBANDINFO, lBand, tRbbi)
      If (lR <> 0) And (tRbbi.hWndChild <> 0) Then
         tRbbi.fMask = RBBIM_CHILDSIZE
         tRbbi.cxMinChild = lWidth
         tRbbi.cyMinChild = lHeight
         If m_lMajor >= 5 Then
            Dim tRbbi471 As REBARBANDINFO_NOTEXT_471
            CopyMemory tRbbi471, tRbbi, LenB(tRbbi)
            If (tRbbi.fStyle And RBBS_CHEVRON) = RBBS_CHEVRON Then
               tRbbi471.fMask = tRbbi471.fMask Or RBBIM_IDEALSIZE
               tRbbi471.cxIdeal = lWidth
            End If
            lR = SendMessage(m_hWnd, RB_SETBANDINFOA, lBand, tRbbi)
            SendMessageLong m_hWnd, RB_MINIMIZEBAND, lBand, 0
         Else
            lR = SendMessage(m_hWnd, RB_SETBANDINFOA, lBand, tRbbi)
            SendMessageLong m_hWnd, RB_MINIMIZEBAND, lBand, 0
         End If
      End If
   End If
End Sub

Public Sub BandMove(ByVal lBand As Long, ByVal lIndexTo As Long)
Attribute BandMove.VB_Description = "Moves a band from one position to another.  All bands in lower positions are moved up.   Not available for COMCTL32.DLL version below 4.71."
    If (lBand >= 0) And (lBand < BandCount) Then
      If (lIndexTo >= 0) And (lIndexTo < BandCount) Then
         SendMessageLong m_hWnd, RB_MOVEBAND, lBand, lIndexTo
      Else
         ' Incorrectband
      End If
   Else
      ' Incorrectband
   End If
End Sub
Public Sub BandMinimise(ByVal lBand As Long)
Attribute BandMinimise.VB_Description = "Minimises a rebar band in the current layout."
    If (lBand >= 0) And (lBand < BandCount) Then
        SendMessageLong m_hWnd, RB_MINIMIZEBAND, lBand, 0
    Else
      ' IncorrectBand
    End If
End Sub
Public Sub BandMaximise(ByVal lBand As Long)
Attribute BandMaximise.VB_Description = "Maximises a rebar band in the current layout."
    If (lBand >= 0) And (lBand < BandCount) Then
        SendMessageLong m_hWnd, RB_MAXIMIZEBAND, lBand, 0
    Else
      ' IncorrectBand
    End If
End Sub
Public Sub GetBandRectangle( _
      ByVal lBand As Long, _
      Optional ByRef lLeft As Long, _
      Optional ByRef lTop As Long, _
      Optional ByRef lRight As Long, _
      Optional ByRef lBottom As Long _
   )
Attribute GetBandRectangle.VB_Description = "Returns the internal bounding rectangle for a rebar band. Not available for COMCTL32.DLL version below 4.71."
Dim tR As RECT
   If (lBand >= 0) And (lBand <= BandCount) Then
      SendMessage m_hWnd, RB_GETRECT, lBand, tR
      lLeft = tR.Left
      lTop = tR.TOp
      lRight = tR.Right
      lBottom = tR.Bottom
   Else
      ' IncorrectBand
   End If
End Sub
Property Get BandCount() As Long
Attribute BandCount.VB_Description = "Returns the number of bands in the rebar."
    BandCount = SendMessage(m_hWnd, RB_GETBANDCOUNT, 0&, ByVal 0&)
End Property

Private Function pbGetBandInfo( _
        ByVal lHwnd As Long, _
        ByVal lBand As Long, _
        Optional ByRef fMask As Long, _
        Optional ByRef fStyle As Long, _
        Optional ByRef clrFore As Long, _
        Optional ByRef clrBack As Long, _
        Optional ByRef cch As Long, _
        Optional ByRef iImage As Integer, _
        Optional ByRef hWndChild As Long, _
        Optional ByRef cxMinChild As Long, _
        Optional ByRef cyMinChild As Long, _
        Optional ByRef cx As Long, _
        Optional ByRef hbmpBack As Long, _
        Optional ByRef wID As Long, _
        Optional ByRef cyIntegral As Long, _
        Optional ByRef cyChild As Long, _
        Optional ByRef cyMaxChild As Long, _
        Optional ByRef lParam As Long, _
        Optional ByRef cxHeader As Long _
    ) As Boolean
Dim tRbbi As REBARBANDINFO_NOTEXT
Dim tRbbi471 As REBARBANDINFO_NOTEXT_471
Dim lR As Long

   If m_lMajor < 4 Or (m_lMajor = 4 And m_lMinor < 71) Then
      ' Use old version
      tRbbi.cbSize = LenB(tRbbi)
      tRbbi.fMask = fMask
      lR = SendMessage(lHwnd, RB_GETBANDINFO, lBand, tRbbi)
      If (lR <> 0) Then
         With tRbbi
            fMask = .fMask
            fStyle = .fStyle
            clrFore = .clrFore
            clrBack = .clrBack
            cch = .cch
            iImage = .iImage
            hWndChild = .hWndChild
            cxMinChild = .cxMinChild
            cyMinChild = .cyMinChild
            cx = .cx
            hbmpBack = .hbmBack
            wID = .wID
         End With
         pbGetBandInfo = True
      End If
   Else
      tRbbi471.cbSize = LenB(tRbbi471)
      tRbbi471.fMask = fMask
      lR = SendMessage(lHwnd, RB_GETBANDINFO471, lBand, tRbbi471)
      If (lR <> 0) Then
         With tRbbi471
            fMask = .fMask
            fStyle = .fStyle
            clrFore = .clrFore
            clrBack = .clrBack
            cch = .cch
            iImage = .iImage
            hWndChild = .hWndChild
            cxMinChild = .cxMinChild
            cyMinChild = .cyMinChild
            cx = .cx
            hbmpBack = .hbmBack
            cyIntegral = .cyIntegral
            cyChild = .cyChild
            cyMaxChild = .cyMaxChild
            cyMinChild = .cyMinChild
            cxHeader = .cxHeader
            lParam = .lParam
            wID = .wID
         End With
         pbGetBandInfo = True
       End If
   End If
End Function
Public Property Get HasBitmap() As Boolean
Attribute HasBitmap.VB_Description = "Returns whether a background bitmap is loaded into the rebar or not."
   HasBitmap = (BackgroundBitmapHandle <> 0)
End Property

Public Property Let ImageSource( _
        ByVal eType As ECRBImageSourceTypes _
    )
Attribute ImageSource.VB_Description = "Specifies which type of bitmap source (file, picture or resource) should be used as the source of the rebar's background bitmap."
    m_eImageSourceType = eType
End Property
Public Property Let ImageResourceID(ByVal lResourceId As Long)
Attribute ImageResourceID.VB_Description = "Sets a resource id to be used  to be used as the source of the rebar's background bitmap."
   ClearPicture
   m_lResourceID = lResourceId
End Property
Public Property Let ImageResourcehInstance(ByVal hInstance As Long)
Attribute ImageResourcehInstance.VB_Description = "Specifies the hInstance from which to load the resource set by the ImageResourceID property."
   m_hInstance = hInstance
End Property
Public Property Let ImageFile(ByVal sFile As String)
Attribute ImageFile.VB_Description = "Sets a bitmap file to be used as the source of the rebar's background bitmap."
   ClearPicture
   m_sPicture = sFile
End Property
Public Property Let ImagePicture(ByVal picThis As Long)
'was Public Property Let ImagePicture(ByVal picThis As StdPicture)
   ClearPicture
   m_hBmp = picThis
   'Set m_pic = picThis
End Property
Public Property Get BackgroundBitmap() As String
Attribute BackgroundBitmap.VB_Description = "Gets/sets the background bitmap file.  Has no effect unless it is called before the rebar is created.  Note: you can't recreate a rebar at run-time if you have COMCTL32.DLL version lower than 4.71."
   BackgroundBitmap = m_sPicture
End Property
Public Property Let BackgroundBitmap(ByVal sFile As String)
   ImageSource = CRBLoadFromFile
   ImageFile = sFile
End Property
Private Property Get BackgroundBitmapHandle() As Long

   ' Set up the picture if we don't already have one:
   If (m_hBmp = 0) Then
      Select Case m_eImageSourceType
      Case CRBPicture
         If Not (m_pic Is Nothing) Then
            m_hBmp = m_pic.Handle
         End If
      Case CTBLoadFromFile
         If (m_sPicture <> "") Then
           ' m_hBmp = LoadImage(0, m_sPicture, IMAGE_BITMAP, 0, 0, _
           '          LR_LOADFROMFILE Or LR_LOADMAP3DCOLORS Or LR_LOADTRANSPARENT)
            m_hBmp = LoadImage(0, m_sPicture, IMAGE_BITMAP, 0, 0, _
                     LR_LOADFROMFILE Or LR_LOADTRANSPARENT)
         End If
      Case CTBResourceBitmap
         m_hBmp = LoadImageLong(m_hInstance, m_lResourceID, IMAGE_BITMAP, 0, 0, _
                     LR_LOADMAP3DCOLORS Or LR_LOADTRANSPARENT)
      End Select
   End If

   BackgroundBitmapHandle = m_hBmp
End Property

Public Function AddBandByHwnd( _
        ByVal hWnd As Long, _
        Optional ByVal sBandText As String = "", _
        Optional ByVal bBreakLine As Boolean = True, _
        Optional ByVal bFixedSize As Boolean = False, _
        Optional ByVal vData As Variant _
    ) As Long
Attribute AddBandByHwnd.VB_Description = "Adds a band to the rebar given the hWnd of the control to place in the band.  The minimum width of the band will be set to the control's size."
Dim hBmp As Long
Dim lX As Long
Dim lBand As Long
Dim hWndP As Long
Dim wID As Long
    
   If (m_hWnd = 0) Then
      debugmsg "Call To AddBandByHWnd before rebar created."
   End If
   
   If (m_hWnd <> 0) Then
      hBmp = BackgroundBitmapHandle()
      
      hWndP = GetParent(hWnd)
      If (hWndP <> 0) Then
         pAddWnds hWnd, hWndP
      End If
      wID = plAddId(vData)
      If (Not (pbRBAddBandByhWnd(m_hWnd, wID, hWnd, sBandText, hBmp, bBreakLine, bFixedSize, lBand))) Then
         debugmsg "Failed to add Band"
         pRemoveID wID
      Else
         AddBandByHwnd = wID
         If Not (m_bSubClassing) Then
             ' Start subclassing:
             'Debug.Print "Start subclassing"
             pCreateSubClass
         End If
         RebarSize
      End If
   End If
End Function
Private Function pbRBAddBandByhWnd( _
        ByVal hWndRebar As Long, _
        ByVal wID As Long, _
        Optional ByVal hWndChild As Long = 0, _
        Optional ByVal sBandText As String = "", _
        Optional ByVal hBmp As Long = 0, _
        Optional ByVal bBreakLine As Boolean = True, _
        Optional ByVal bFixedSize As Boolean = False, _
        Optional ByRef ltRBand As Long _
    ) As Boolean

If hWndRebar = 0 Then
    MsgBox "No hWndRebar!"
    Exit Function
End If

Dim sClassName As String
Dim hWndReal As Long
Dim tRBand As REBARBANDINFO
Dim tRBand471 As REBARBANDINFO_471
Dim tRBandNT As REBARBANDINFO_NOTEXT
Dim tRBandNT471 As REBARBANDINFO_NOTEXT_471
Dim bNoText As Boolean
Dim rct As RECT
Dim fMask As Long
Dim fStyle As Long
Dim dwStyle As Long
Dim bListStyle As Boolean

   hWndReal = hWndChild
   
   If Not (hWndChild = 0) Then
      'Check to see if it's a toolbar (so we can
      'make if flat)
      fMask = RBBIM_CHILD Or RBBIM_CHILDSIZE
      sClassName = Space$(255)
      GetClassName hWndChild, sClassName, 255
      'see if it's a real Windows toolbar
      If InStr(UCase$(sClassName), "TOOLBARWINDOW32") Then
         dwStyle = GetWindowLong(hWndChild, GWL_STYLE)
         dwStyle = dwStyle Or TBSTYLE_FLAT Or TBSTYLE_TRANSPARENT
         SetWindowLong hWndChild, GWL_STYLE, dwStyle
      End If
      'Could be a VB Toolbar -- make it flat anyway.
      If InStr(UCase$(sClassName), "TOOLBARWNDCLASS") Then
          hWndReal = GetWindow(hWndChild, GW_CHILD)
         dwStyle = GetWindowLong(hWndReal, GWL_STYLE)
         dwStyle = dwStyle Or TBSTYLE_FLAT Or TBSTYLE_TRANSPARENT
         SetWindowLong hWndReal, GWL_STYLE, dwStyle
      End If
   End If
   
   GetWindowRect hWndReal, rct
   
   If hBmp <> 0 Then
       fMask = fMask Or RBBIM_BACKGROUND
   End If
   fMask = fMask Or RBBIM_STYLE Or RBBIM_ID Or RBBIM_COLORS Or RBBIM_SIZE
   If sBandText <> "" Then
      fMask = fMask Or RBBIM_TEXT
      tRBand.lpText = sBandText
      tRBand.cch = Len(sBandText)
   Else
      bNoText = True
   End If
   
   fStyle = RBBS_FIXEDBMP ' or RBBS_CHILDEDGE
   If bBreakLine = True Then
       fStyle = fStyle Or RBBS_BREAK
   End If
   If bFixedSize = True Then
       fStyle = fStyle Or RBBS_FIXEDSIZE
   Else
       fStyle = fStyle And Not RBBS_FIXEDSIZE
   End If
      
   If (bNoText) Then
      With tRBandNT
         .fMask = fMask
         .fStyle = fStyle
         'Only set if there's a child window
         If hWndReal <> 0 Then
            .hWndChild = hWndReal
            If m_ePosition = erbPositionLeft Or m_ePosition = erbPositionRight Then
               .cxMinChild = rct.Bottom - rct.TOp
               .cyMinChild = rct.Right - rct.Left
            Else
               .cxMinChild = rct.Right - rct.Left
               .cyMinChild = rct.Bottom - rct.TOp
            End If
         End If
         'Set the rest OK
         .wID = wID
         .clrBack = GetSysColor(COLOR_BTNFACE)
         .clrFore = GetSysColor(COLOR_BTNTEXT)
         .cx = 200
         .hbmBack = hBmp
         'The length of the type
         .cbSize = LenB(tRBandNT)
      End With
      If m_lMajor > 4 Or (m_lMajor = 4 And m_lMinor >= 71) Then
         CopyMemory tRBandNT471, tRBandNT, LenB(tRBandNT)
         tRBandNT471.cbSize = LenB(tRBandNT471)
         tRBandNT471.fMask = tRBandNT471.fMask Or RBBIM_IDEALSIZE
         tRBandNT471.cxIdeal = tRBandNT471.cxMinChild
         tRBandNT471.fStyle = tRBandNT471.fStyle Or RBBS_CHEVRON
         pbRBAddBandByhWnd = (SendMessage(hWndRebar, RB_INSERTBAND, -1, tRBandNT471) <> 0)
      Else
         pbRBAddBandByhWnd = (SendMessage(hWndRebar, RB_INSERTBAND, -1, tRBandNT) <> 0)
      End If
   Else
      With tRBand
         .fMask = fMask
         .fStyle = fStyle
         'Only set if there's a child window
         If hWndReal <> 0 Then
            .hWndChild = hWndReal
            If m_ePosition = erbPositionLeft Or m_ePosition = erbPositionRight Then
               .cxMinChild = rct.Bottom - rct.TOp
               .cyMinChild = rct.Right - rct.Left
            Else
               .cxMinChild = rct.Right - rct.Left
               .cyMinChild = rct.Bottom - rct.TOp
            End If
         End If
         'Set the rest OK
         .wID = wID
         .clrBack = GetSysColor(COLOR_BTNFACE)
         .clrFore = GetSysColor(COLOR_BTNTEXT)
         .cx = 200
         .hbmBack = hBmp
         'The length of the type
         .cbSize = LenB(tRBand)
      End With
      If m_lMajor > 4 Or (m_lMajor = 4 And m_lMinor >= 71) Then
         CopyMemory tRBand471, tRBand, LenB(tRBandNT)
         tRBand471.cbSize = LenB(tRBand471)
         pbRBAddBandByhWnd = (SendMessage(hWndRebar, RB_INSERTBAND, -1, tRBand471) <> 0)
      Else
         pbRBAddBandByhWnd = (SendMessage(hWndRebar, RB_INSERTBAND, -1, tRBand) <> 0)
      End If
   End If
   
   ltRBand = BandCount

End Function

Private Sub pRemoveID( _
        ByVal wID As Long _
    )
Dim lItem As Long
Dim lTarget As Long
    
   For lItem = 1 To m_lIDCount
      If (m_tDataStore(lItem).wID = wID) Then
      Else
         lTarget = lTarget + 1
         If (lTarget <> lItem) Then
            LSet m_tDataStore(lTarget) = m_tDataStore(lItem)
         End If
      End If
   Next lItem
   If lTarget = 0 Then
      debugmsg "Removed all IDs and data"
      m_lIDCount = 0
      Erase m_tDataStore
   Else
      If (lTarget <> m_lIDCount) Then
         debugmsg "Reduced ID Count to : " & lTarget
         m_lIDCount = lTarget
         ReDim Preserve m_tDataStore(1 To m_lIDCount) As tRebarDataStore
      End If
   End If
    
End Sub
Property Get BandIndexForId( _
        ByVal wID As Long _
    ) As Long
Attribute BandIndexForId.VB_Description = "Returns the internal index of a band given the band's id."
Dim lItem As Long
Dim tRbbi As REBARBANDINFO_NOTEXT
Dim lIndex As Long
Dim lR As Long

   If m_lMajor < 4 Or (m_lMajor = 4 And m_lMinor < 71) Then
      lIndex = -1
      tRbbi.cbSize = Len(tRbbi)
      tRbbi.fMask = RBBIM_ID
      For lItem = 0 To BandCount - 1
          lR = SendMessage(m_hWnd, RB_GETBANDINFO, lItem, tRbbi)
          If (lR <> 0) Then
              If (wID = tRbbi.wID) Then
                  lIndex = lItem
                  Exit For
              End If
          End If
      Next lItem
      BandIndexForId = lIndex
   Else
      BandIndexForId = SendMessageLong(m_hWnd, RB_IDTOINDEX, wID, 0)
   End If
End Property
Property Get BandIDForIndex( _
      ByVal lIndex As Long _
   ) As Long
Attribute BandIDForIndex.VB_Description = "Gets the ID of band given its 0-based index in the rebar."
Dim lR As Long
Dim tRbbi As REBARBANDINFO_NOTEXT

   tRbbi.cbSize = Len(tRbbi)
   tRbbi.fMask = RBBIM_ID
   lR = SendMessage(m_hWnd, RB_GETBANDINFO, lIndex, tRbbi)
   BandIDForIndex = tRbbi.wID
   
End Property
Public Property Get BandData( _
      ByVal wID As Long _
   ) As Variant
Attribute BandData.VB_Description = "Gets/sets a variant value associated with a band in the rebar."
Dim lItem As Long
   For lItem = 1 To m_lIDCount
      If m_tDataStore(lItem).wID = wID Then
         BandData = m_tDataStore(lItem).vData
         Exit For
      End If
   Next lItem
End Property

Property Get BandIndexForData( _
        ByVal vData As Variant _
    ) As Long
Attribute BandIndexForData.VB_Description = "Returns the index of a band given the band's key."
Dim lItem As Long
Dim lAt As Long
Dim vItem As Variant
On Error Resume Next
    lAt = -1
    For lItem = 1 To m_lIDCount
      If IsMissing(m_tDataStore(lItem).vData) Then
         vItem = ""
      ElseIf IsObject(m_tDataStore(lItem).vData) Then
         If (vData Is m_tDataStore(lItem).vData) Then
            lAt = lItem
            Exit For
         End If
      Else
         If vData = m_tDataStore(lItem).vData Then
            lAt = lItem
            Exit For
         End If
      End If
      
    Next lItem
    If (lAt > 0) Then
        lAt = BandIndexForId(m_tDataStore(lAt).wID)
    End If
    BandIndexForData = lAt
End Property
Private Function plAddId( _
        ByVal vData As Variant _
    ) As Long
    m_lIDCount = m_lIDCount + 1
    ReDim Preserve m_tDataStore(1 To m_lIDCount) As tRebarDataStore
    m_tDataStore(m_lIDCount).wID = m_lIDCount
    m_tDataStore(m_lIDCount).vData = vData
    plAddId = m_lIDCount
End Function
Private Sub pAddWnds( _
        ByVal hwndItem As Long, _
        ByVal hWndParent As Long _
    )
   m_iWndItemCount = m_iWndItemCount + 1
   ReDim Preserve m_tWndStore(1 To m_iWndItemCount) As tRebarWndStore
   With m_tWndStore(m_iWndItemCount)
      .hwndItem = hwndItem
      .hWndItemParent = hWndParent
      GetWindowRect hwndItem, .tR
   End With
End Sub
Private Sub pResetParent( _
        ByVal hwndItem As Long _
    )
Dim iItem As Long
Dim iTarget As Long
Dim bSuccess As Boolean
    
   For iItem = 1 To m_iWndItemCount
      If (m_tWndStore(iItem).hwndItem = hwndItem) Then
         SetParent m_tWndStore(iItem).hwndItem, m_tWndStore(iItem).hWndItemParent
         ' Reset the size to original:
         SetWindowPos m_tWndStore(iItem).hwndItem, 0, m_tWndStore(iItem).tR.Left, m_tWndStore(iItem).tR.TOp, m_tWndStore(iItem).tR.Right - m_tWndStore(iItem).tR.Left, m_tWndStore(iItem).tR.Bottom - m_tWndStore(iItem).tR.TOp, SWP_NOREDRAW Or SWP_NOZORDER Or SWP_NOOWNERZORDER
         'MoveWindow m_tWndStore(iItem).hWndItem, m_tWndStore(iItem).tR.Left, m_tWndStore(iItem).tR.Top, m_tWndStore(iItem).tR.Right - m_tWndStore(iItem).tR.Left, m_tWndStore(iItem).tR.Bottom - m_tWndStore(iItem).tR.Top, 1
         bSuccess = True
      Else
         iTarget = iTarget + 1
         If iTarget <> iItem Then
            LSet m_tWndStore(iTarget) = m_tWndStore(iItem)
         End If
      End If
   Next iItem
   
   If (iTarget = 0) Then
      debugmsg "Successfully reset all parents"
      m_iWndItemCount = 0
      Erase m_tWndStore
   Else
      If iTarget <> m_iWndItemCount Then
         debugmsg "Decrease wnd count to " & iTarget
         m_iWndItemCount = iTarget
         ReDim Preserve m_tWndStore(1 To m_iWndItemCount) As tRebarWndStore
      End If
   End If
   
   
   If Not bSuccess Then
      debugmsg "Failed to reset parent.."
      ' At least ensure it won't stop the rebar terminating:
      ShowWindow hwndItem, SW_HIDE
      SetParent hwndItem, 0
   End If
End Sub
Public Sub RebarSize()
Attribute RebarSize.VB_Description = "Sizes the rebar to the parent object."
Dim lLeft As Long, lTop As Long
Dim cx As Long, cy As Long
Dim RC As RECT, rcB As RECT, rcI As RECT, rcP As RECT
   
   If (m_hWnd <> 0) Then
      GetWindowRect m_hWnd, rcB
      OffsetRect rcB, -rcB.Left, -rcB.TOp
      GetClientRect m_hWndCtlParent, rcP
      If (m_ePosition = erbPositionBottom) Or (m_ePosition = erbPositionTop) Then
         cx = rcP.Right - rcP.Left
         cy = RebarHeight
         If m_ePosition = erbPositionBottom Then
            lTop = rcP.Bottom - RC.TOp - cy
         End If
         AdjustForOtherRebars m_hWnd, lLeft, lTop, cx, cy
         SetWindowPos m_hWnd, 0, lLeft, lTop, cx, cy, SWP_NOZORDER Or SWP_NOACTIVATE
      Else
         cy = rcP.Bottom - rcP.TOp
         cx = RebarHeight
         If m_ePosition = erbPositionRight Then
            lLeft = rcP.Right - rcP.Left - cx
         End If
         AdjustForOtherRebars m_hWnd, lLeft, lTop, cx, cy
         SetWindowPos m_hWnd, 0, lLeft, lTop, cx, cy, SWP_NOZORDER Or SWP_NOACTIVATE
      End If
      GetWindowRect m_hWnd, RC
      OffsetRect RC, -RC.Left, -RC.TOp
      UnionRect rcI, RC, rcB
      InvalidateRect m_hWnd, rcI, True
      UpdateWindow m_hWnd
   End If
   
End Sub
Property Get hWnd() As Long
Attribute hWnd.VB_Description = "Returns the window handle of the control.  Use RebarhWnd to get the handle of the Rebar itself."
    hWnd = UserControl.hWnd
End Property
Property Get RebarHwnd() As Long
Attribute RebarHwnd.VB_Description = "Returns the windows handle of the Rebar window."
    RebarHwnd = m_hWnd
End Property
Public Property Get RebarHeight() As Long
Attribute RebarHeight.VB_Description = "Gets the current height of the rebar."
Dim tc As RECT
    'If (m_hWnd <> 0) Then
    '  GetWindowRect m_hWnd, tc
    '  RebarHeight = (tc.Bottom - tc.Top)
    'End If
    ' Get the height that would be good for the rebar:
   If m_bVisible Then
      RebarHeight = SendMessageLong(m_hWnd, RB_GETBARHEIGHT, 0, 0) + 4
   Else
      RebarHeight = 0
   End If
End Property
Public Property Get RebarWidth() As Long
Dim tc As RECT
   If (m_hWnd <> 0) Then
      If m_bVisible Then
         GetWindowRect m_hWnd, tc
         RebarWidth = (tc.Right - tc.Left)
      Else
         RebarWidth = 0
      End If
   End If
End Property
Private Function pbLoadCommCtls() As Boolean
Dim ctEx As CommonControlsEx

    ctEx.dwSize = Len(ctEx)
    ctEx.dwICC = ICC_COOL_CLASSES Or _
        ICC_USEREX_CLASSES Or ICC_WIN95_CLASSES
    
    pbLoadCommCtls = (InitCommonControlsEx(ctEx) <> 0)

End Function

Public Function CreateRebar(ByVal hWndParent As Long) As Boolean
Attribute CreateRebar.VB_Description = "Initialises a rebar for use and allows you to specify the host window for the rebar.  For a standard form, this should be the form.  For an MDI form, this should be a PictureBox aligned to the top of the MDI form."
   If (UserControl.Ambient.UserMode) Then
      DestroyRebar
      ' Set up the rebar:
      If (pbCreateRebar(hWndParent)) Then
         SetProp m_hWnd, "vbal:cRebarPosition", m_ePosition
         m_hWndCtlParent = hWndParent
         AddRebar m_hWnd, m_hWndCtlParent
      End If
   End If
End Function
Public Function AddResizeObject(ByVal hWndParent As Long, ByVal hWnd As Long, ByVal ePosition As ERBPositionConstants)
Attribute AddResizeObject.VB_Description = "Adds a control to the list of objects to be considered when resizing a rebar on screen.  Other rebars are automatically taken into account."
   AddRebar hWnd, hWndParent
   SetProp hWnd, "vbal:cRebarPosition", ePosition
End Function
Private Function pbCreateRebar(ByVal hWndParent As Long) As Boolean
Dim lWidth As Long
Dim lHeight As Long
Dim bVertical As Boolean
Dim hwndCoolBar As Long
Dim lResult As Long
Dim cStyle As Long
Dim RC As RECT

    If (UserControl.Ambient.UserMode) Then
    
      ' Try to load the Common Controls support for the
      ' rebar control:
      If (pbLoadCommCtls()) Then
         'Debug.Print "Loaded Coolbar support"
         ' If we have done this, then build a rebar:
         'lWidth = UserControl.Parent.ScaleWidth \ Screen.TwipsPerPixelX
         'lHeight = UserControl.Height \ Screen.TwipsPerPixelY
         GetWindowRect hWndParent, RC
         lWidth = RC.Right - RC.Left
         lHeight = RC.Bottom - RC.TOp

         ComCtlVersion m_lMajor, m_lMinor
         cStyle = WS_CHILD Or WS_BORDER Or _
             WS_CLIPCHILDREN Or WS_CLIPSIBLINGS Or _
             WS_VISIBLE
         Select Case m_ePosition
         Case erbPositionTop
            cStyle = cStyle Or CCS_TOP
         Case erbPositionRight
            cStyle = cStyle Or CCS_RIGHT
         Case erbPositionLeft
            cStyle = cStyle Or CCS_LEFT
         Case erbPositionBottom
            cStyle = cStyle Or CCS_BOTTOM
         End Select
         cStyle = cStyle Or CCS_NORESIZE
         cStyle = cStyle Or CCS_NODIVIDER
         
         cStyle = cStyle Or RBS_DBLCLKTOGGLE
         cStyle = cStyle Or RBS_VARHEIGHT Or RBS_BANDBORDERS
         cStyle = cStyle Or RBS_AUTOSIZE
   
         m_hWnd = CreateWindowEX(WS_EX_TOOLWINDOW, _
                              REBARCLASSNAME, "", _
                              cStyle, 0, 0, lWidth, lHeight, _
                              hWndParent, ICC_COOL_CLASSES, App.hInstance, ByVal 0&)
         If (m_hWnd <> 0) Then
            ' Debug.Print "Created Rebar Window"
            AddToToolTip m_hWnd
            If m_lMajor >= 5 Then
               SendMessageLong m_hWnd, RB_SETEXTENDEDSTYLE, 0, RBS_EX_OFFICE9
            End If
            pbCreateRebar = True
         End If
      End If
    End If
    
End Function
Public Sub DestroyRebar()
Attribute DestroyRebar.VB_Description = "Removes all bands from a rebar and clears all resources associated with it."
   If (m_hWnd <> 0) Then
      ' Debug.Print "Destroying rebar window"
      RemoveRebar m_hWnd
      
      RemoveFromToolTip m_hWnd
      RemoveAllRebarBands
      
      pDestroySubClass
      ShowWindow m_hWnd, SW_HIDE
      SetParent m_hWnd, 0
      DestroyWindow m_hWnd
      m_hWnd = 0
      m_hWndCtlParent = 0
                
   End If
   
End Sub

Public Sub RemoveAllRebarBands()
Attribute RemoveAllRebarBands.VB_Description = "Removes all bands from the rebar.  To prevent controls not terminating when a form unloads because they are contained by a different parent, call this method."
Dim lBands As Long
Dim lBand As Long
    If (m_hWnd <> 0) Then
        lBands = BandCount
        For lBand = 0 To lBands - 1
            RemoveBand 0
        Next lBand
        pDestroySubClass
    End If
End Sub
Public Sub RemoveBand( _
        ByVal lBand As Long _
    )
Attribute RemoveBand.VB_Description = "Removes a specified band from the rebar control."
Dim lHwnd As Long
Dim wID As Long

    If (m_hWnd <> 0) Then
        ' If a valid band:
        If (lBand >= 0) And (lBand < BandCount) Then
            If m_lMajor < 4 Or (m_lMajor = 4 And m_lMinor < 71) Then
               ' Remove the child from this band:
               lHwnd = plGetHwndOfBandChild(m_hWnd, lBand, wID)
               If (lHwnd <> 0) Then
                   pResetParent lHwnd
               End If
               ' Remove the band:
               SendMessageLong m_hWnd, RB_DELETEBAND, lBand, 0&
               ' Remove the id for this band:
               pRemoveID wID
               ' No bands left? Stop subclassing:
               If (BandCount = 0) Then
                  debugmsg "All bands destroyed"
                  pDestroySubClass
               End If
            Else
               SendMessageLong m_hWnd, RB_DELETEBAND, lBand, 0&
               If BandCount = 0 Then
                  debugmsg "All bands destroyed"
                  pDestroySubClass
               End If
            End If
        End If
    End If
End Sub
Private Function plGetHwndOfBandChild( _
        ByVal lHwnd As Long, _
        ByVal lBand As Long, _
        ByVal wID As Long _
    ) As Long
Dim lParam As Long
Dim tRbbi As REBARBANDINFO_NOTEXT
Dim lR As Long

    tRbbi.cbSize = Len(tRbbi)
    tRbbi.fMask = RBBIM_CHILD Or RBBIM_ID
    lR = SendMessage(lHwnd, RB_GETBANDINFO, lBand, tRbbi)
    If (lR <> 0) Then
        plGetHwndOfBandChild = tRbbi.hWndChild
        wID = tRbbi.wID
    End If
End Function

Public Property Get Visible() As Boolean
Attribute Visible.VB_Description = "Gets/sets whether the entire rebar will be visible or not."
   Visible = m_bVisible
End Property
Public Property Let Visible(ByVal bState As Boolean)
   m_bVisible = bState
   If m_hWnd <> 0 Then
      If Not bState Then
         ShowWindow m_hWnd, SW_HIDE
         RaiseEvent HeightChanged(0)
      Else
         ShowWindow m_hWnd, SW_SHOW
         RaiseEvent HeightChanged(RebarHeight)
      End If
   End If
   PropertyChanged "Visible"
End Property

Private Sub ClearPicture()
   If (m_hBmp <> 0) Then
      If (m_pic Is Nothing) Then
         DeleteObject m_hBmp
         m_hBmp = 0
      End If
   End If
   m_sPicture = ""
   m_lResourceID = 0
   Set m_pic = Nothing
End Sub

'Private Sub ISubclass_WndMessage(Caller As SubClassLib.Subclass, hwnd As Long, Msg As Long, wParam As Long, lParam As Long, ReturnValue As Long, Consume As Boolean)

'End Sub

Private Sub UserControl_Initialize()
    debugmsg "cRebar:Initialise"
    m_lMajor = 4
    m_lMinor = 0
    m_bVisible = True
End Sub

Private Sub UserControl_InitProperties()
   ' If init properties we must be in design mode.
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    
    ' Read in properties here:
    ' ...
    
End Sub

Private Sub UserControl_Resize()
   If (UserControl.Ambient.UserMode) Then
      UserControl.Width = 0
      UserControl.Height = 0
   End If
End Sub

Private Sub UserControl_Terminate()
    m_bInTerminate = True
    DestroyRebar
    ClearPicture
    debugmsg "cRebar:Terminate"
    'MsgBox "cRebar:Terminate"
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    
    ' Write properties here:
    ' ...
    
End Sub


