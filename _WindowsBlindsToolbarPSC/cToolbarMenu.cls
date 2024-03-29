VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cToolbarMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' =======================================================================
' FileName:    cToolbarMenu
' Author:      Steve McMahon
' Date:        8 Feb 2000
'
' Allows menus to pop up and cancel as the user hovers
' over toolbar buttons.
'
'
' Copyright © 2000 Steve McMahon
' =======================================================================

Private Enum TRACKINGSTATE   '{ // menubar has three states:
   TRACK_NONE = 0 ',   // * normal, not tracking anything
   TRACK_BUTTON ',     // * tracking buttons (F10/Alt mode)
   TRACK_POPUP       '// * tracking popups
End Enum

' Track popup menu constants:

Private m_iTrackingState As TRACKINGSTATE
Private m_bProcessRightArrow As Boolean
Private m_bProcessLeftArrow  As Boolean
Private m_hMenuTracking As Long
Private m_iPopupTracking As Long
Private m_bEscapeWasPressed As Boolean
Private m_tPMouse As POINTAPI
Private m_iNewPopup As Long
Private m_bIn As Boolean

Private m_hWnd As Long
Private m_lPtr As Long

Private m_iExit As Integer

Implements ISubclass


Friend Sub CoolMenuAttach(ByRef hWndA As Long, ByVal cBar As cMenuBar)
Dim lPtr As Long

   m_iExit = 0
   CoolMenuDetach
   m_hWnd = hWndA
   SendMessageLong m_hWnd, WM_ENTERMENULOOP, 0, 0
   AttachMessage Me, m_hWnd, WM_MENUSELECT
   m_lPtr = ObjPtr(cBar)
   
End Sub
Friend Sub CoolMenuDetach()
   If (m_hWnd <> 0) Then
      SendMessageLong m_hWnd, WM_EXITMENULOOP, 0, 0
      DetachMessage Me, m_hWnd, WM_MENUSELECT
      m_hWnd = 0
   End If
   m_hWnd = 0
   m_lPtr = 0
End Sub

'/////////////////
'// When user selects a new menu item, note whether it has a submenu
'// and/or parent menu, so I know whether right/left arrow should
'// move to the next popup.
'//
Private Sub MenuSelect(ByVal hMenu As Long, ByVal iItem As Long)
   If (m_iTrackingState > 0) Then
      '// process right-arrow if item is NOT a submenu
      m_bProcessRightArrow = (GetSubMenu(hMenu, iItem) = 0)
      '// process left-arrow if curent menu is one I'm tracking
      m_bProcessLeftArrow = (hMenu = m_hMenuTracking)
   End If
End Sub


'//////////////////
'// Handle menu input event: Look for left/right to change popup menu,
'// mouse movement over over a different menu button for "hot" popup effect.
'// Returns TRUE if message handled (to eat it).
'//
Friend Function MenuInput(m As Msg) As Boolean
Dim iMsg As Long
Dim vKey As Long
Dim tP As POINTAPI
Dim iButton As Long

   'ASSERT_VALID(this);
   Debug.Assert m_iTrackingState = TRACK_POPUP  '; // sanity check
   iMsg = m.message

   If (iMsg = WM_KEYDOWN) Then
      
      '// handle left/right-arow.
      vKey = m.wParam
      If ((vKey = vbKeyLeft And m_bProcessLeftArrow) Or _
         (vKey = vbKeyRight And m_bProcessRightArrow)) Then

         'MBTRACE(_T("CMenuBar::OnMenuInput: handle VK_LEFT/RIGHT\n"));
         CancelMenuAndTrackNewOne _
            GetNextOrPrevButton(m_iPopupTracking, vKey = vbKeyLeft)
         MenuInput = True ' // eat it
      
      ' // escape:
      ElseIf (vKey = vbKeyEscape) Then
         m_bEscapeWasPressed = True ';    // (menu will abort itself)
      End If
      
   ElseIf (iMsg = WM_MOUSEMOVE Or iMsg = WM_LBUTTONDOWN) Then
      '// handle mouse move or click
      LSet tP = m.pt
      'ScreenToClient m_hWndBand, tP

      If (iMsg = WM_MOUSEMOVE) Then
         'If (tP.X <> m_tPMouse.X) And (tP.Y <> m_tPMouse.Y) Then
            iButton = HitTest(tP)
            If IsValidButton(iButton) Then
               If iButton <> m_iPopupTracking Then
                  '// user moved mouse over a different button: track its popup
                  CancelMenuAndTrackNewOne iButton
               End If
            End If
            LSet m_tPMouse = tP
         'End If
      ElseIf iMsg = WM_LBUTTONDOWN Then
         If (HitTest(tP) = m_iPopupTracking) Then
            '// user clicked on same button I am tracking: cancel menu
            'MBTRACE(_T("CMenuBar:OnMenuInput: handle mouse click to exit popup\n"));
            CancelMenuAndTrackNewOne -1
            MenuInput = True ' // eat it
         End If
      End If
      
   ElseIf iMsg = WM_LBUTTONUP Or iMsg = WM_RBUTTONUP Then
   
   End If

End Function

Private Function HitTest(pt As POINTAPI) As Long
Dim cBar As cMenuBar
   If GetBar(cBar) Then
      HitTest = cBar.HitTest(pt)
   End If

End Function
Private Property Get IsValidButton(ByVal iButton As Long) As Boolean
   If (iButton > 0) Then
      IsValidButton = True
   End If
End Property

'//////////////////
'// Cancel the current popup menu by posting WM_CANCELMODE, and track a new
'// menu. iNewPopup is which new popup to track (-1 to quit).
'//
Private Sub CancelMenuAndTrackNewOne(ByVal iNewPopup As Long)
Dim cBar As cMenuBar
Dim hMenuPopup As Long
   'MBTRACE(_T("CMenuBar::CancelMenuAndTrackNewOne: %d\n"), iNewPopup);
   'ASSERT_VALID(this);
   If iNewPopup > 0 Then
      If (iNewPopup <> m_iPopupTracking) Then
         If GetBar(cBar) Then
            hMenuPopup = cBar.GetMenuHandle(iNewPopup)
            If hMenuPopup <> 0 Then
               'PostMessage m_hWndOwner, WM_CANCELMODE, 0, 0 ' // quit menu loop
               PostMessage m_hWnd, WM_CANCELMODE, 0, 0
               m_iNewPopup = iNewPopup                '// go to this popup (-1 = quit)
            End If
         End If
      End If
   End If
End Sub

'//////////////////
'// Track the popup submenu associated with the i'th button in the menu bar.
'// This fn actually goes into a loop, tracking different menus until the user
'// selects a command or exits the menu.
'//
Friend Function TrackPopup(ByVal iButton As Long) As Long
Dim nMenuItems As Long
Dim tPM As TPMPARAMS
Dim rcButton As RECT
Dim pt As POINTAPI
Dim hMenuPopup As Long
Dim lR As Long
Dim hwnd As Long
Dim lRtnID As Long
Dim cBar As cMenuBar

   If Not m_bIn Then
      m_bIn = True
      m_iNewPopup = iButton
      'Debug.Assert m_hMenu <> 0
      If GetBar(cBar) Then
         
         nMenuItems = cBar.Count 'GetMenuItemCount(m_hMenu)
      
         Do While (m_iNewPopup > -1)               '// while user selects another menu
            
            lRtnID = 0
      
            m_iNewPopup = -1                '// assume quit after this
            PressButton iButton, True       '// press the button
            'UpdateWindow ToolbarhWnd(m_hWnd)             '// and force repaint now
      
            SetTrackingState TRACK_POPUP, iButton '// enter tracking state
      
            '// Need to install a hook to trap menu input in order to make
            '// left/right-arrow keys and "hot" mouse tracking work.
            '//
            AttachMsgHook Me
      
            '// get submenu and display it beneath button
            GetRect iButton, rcButton
            'ClientRectToScreen m_hWndBand, rcButton
            tPM.cbSize = Len(tPM)
            ComputeMenuTrackPoint rcButton, tPM, pt
            
            'hMenuPopup = GetSubMenu(m_hMenu, iButton)
            hMenuPopup = cBar.GetMenuHandle(iButton)
            
            If hMenuPopup <> 0 Then
               ' Show the menu:
               m_hMenuTracking = hMenuPopup
               lR = TrackPopupMenuEx(hMenuPopup, _
                  TPM_LEFTALIGN Or TPM_LEFTBUTTON Or TPM_VERTICAL, _
                  pt.x, pt.y, m_hWnd, tPM)
               'lR is the ID of the menu
               lRtnID = lR
            End If
                  
            '// uninstall hook.
            DetachMsgHook
      
            PressButton iButton, False    ';   // un-press button
            'UpdateWindow ToolbarhWNd(m_hWnd)                '// and force repaint now
            
            '// If the user exited the menu loop by pressing Escape,
            '// return to track-button state; otherwise normal non-tracking state.
            If (m_bEscapeWasPressed) Then
               SetTrackingState TRACK_NONE, iButton
            Else
               SetTrackingState TRACK_NONE, iButton
            End If
            
            '// If the user moved mouse to a new top-level popup (eg from File to
            '// Edit button), I will have posted a WM_CANCELMODE to quit
            '// the first popup, and set m_iNewPopup to the new menu to show.
            '// Otherwise, m_iNewPopup will be -1 as set above.
            '// So just set iButton to the next popup menu and keep looping...
            iButton = m_iNewPopup
            
         Loop
      
         ' Set hot button if mouse is over, otherwise not:
         
         ' The ID of the selected menu
         TrackPopup = lRtnID
      End If
      m_bIn = False
   End If
End Function
Private Sub ComputeMenuTrackPoint(ByRef rc As RECT, tPM As TPMPARAMS, tP As POINTAPI)
   tP.x = rc.Left
   tP.y = rc.Bottom
   LSet tPM.rcExclude = rc
End Sub

Private Function GetBar(ByRef cBar As cMenuBar) As Boolean
   If Not m_lPtr = 0 Then
      Set cBar = ObjectFromPtr(m_lPtr)
      'Debug.Print "GetBar:OK"
      GetBar = True
   End If
End Function

Private Sub PressButton(ByVal iButton As Long, ByVal bState As Boolean)
Dim fState As Long
Dim cBar As cMenuBar

   If GetBar(cBar) Then
      If iButton > 0 And iButton <= cBar.Count Then
         cBar.PressButton iButton, bState
      End If
   End If

End Sub

Private Sub GetRect(ByVal iButton As Long, ByRef tR As RECT)
Dim cBar As cMenuBar
   tR.Left = 0: tR.TOp = 0: tR.Bottom = 0: tR.Right = 0
   If GetBar(cBar) Then
      If iButton > 0 And iButton <= cBar.Count Then
         cBar.GetRect iButton, tR
      End If
   End If
End Sub
Private Function GetHotItem() As Long
Dim cBar As cMenuBar
   If GetBar(cBar) Then
      GetHotItem = cBar.HotItem
   End If
End Function
Private Function SetHotItem(ByVal iButton As Long) As Long
Dim cBar As cMenuBar
   If GetBar(cBar) Then
      'Debug.Print "Setting hot item: " & iButton
      cBar.HotItem = iButton
   End If
End Function
Private Function GetButtonVisible(ByVal iButton As Long) As Boolean
   GetButtonVisible = True
End Function
Private Function GetButtonCount() As Long
Dim cBar As cMenuBar
   If GetBar(cBar) Then
      GetButtonCount = cBar.Count
   End If
End Function

Private Sub SetTrackingState(ByVal iState As TRACKINGSTATE, ByVal iButton As Long)
   If (iState <> m_iTrackingState) Then
      If (iState = TRACK_NONE) Then
         iButton = -1
      End If
'#ifdef _DEBUG
'      static LPCTSTR StateName[] = { _T("NONE"), _T("BUTTON"), _T("POPUP") };
'      MBTRACE(_T("CMenuBar::SetTrackingState to %s, button=%d\n"),
'         StateName[iState], iButton);
'#End If

      SetHotItem iButton              '// could be none (-1)

      If (iState = TRACK_POPUP) Then
         '// set related state stuff
         m_bEscapeWasPressed = False 'FALSE;   // assume Esc key not pressed
         m_bProcessRightArrow = True        '// assume left/right arrow..
         m_bProcessLeftArrow = True         '; // ..will move to prev/next popup
         m_iPopupTracking = iButton          '// which popup I'm tracking
      End If
      m_iTrackingState = iState
   End If
End Sub


Private Function GetNextOrPrevButton(ByVal iButton As Long, ByVal bPrev As Boolean) As Long
Dim iSB As Long
Dim bfound As Boolean

   If (bPrev) Then
      iSB = iButton
      Do While Not bfound
         
         iButton = iButton - 1
         If iButton < 1 Then
            iButton = GetButtonCount()
         End If
         
         If Not (GetButtonVisible(iButton)) Then
            If iButton = iSB Then
               iButton = -1
               Exit Do
            End If
         Else
            bfound = True
         End If
         
      Loop
      
   Else
      iSB = iButton
      Do While Not bfound
         iButton = iButton + 1
         If (iButton > GetButtonCount()) Then
            iButton = 1
         End If
         
         If Not GetButtonVisible(iButton) Then
            If iButton = iSB Then
               iButton = -1
               Exit Do
            End If
         Else
            bfound = True
         End If
         
      Loop
      
   End If
   GetNextOrPrevButton = iButton
End Function
'//////////////////
'// Toggle state from home state to button-tracking and back
'//
Private Sub ToggleTrackButtonMode()
   If (m_iTrackingState = TRACK_NONE Or m_iTrackingState = TRACK_BUTTON) Then
      If m_iTrackingState = TRACK_NONE Then
         SetTrackingState TRACK_BUTTON, 1
      Else
         SetTrackingState TRACK_NONE, 1
     End If
   End If
End Sub



Private Property Get ISubclass_MsgResponse() As EMsgResponse
   If CurrentMessage = WM_MENUSELECT Then
      ISubclass_MsgResponse = emrPreprocess
   End If
End Property

Private Property Let ISubclass_MsgResponse(ByVal RHS As EMsgResponse)
   '
End Property


Private Function ISubclass_WindowProc(ByVal hwnd As Long, ByVal iMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
   Select Case iMsg
   Case WM_MENUSELECT
      MenuSelect lParam, (wParam And &HFFFF&)
   End Select
End Function


