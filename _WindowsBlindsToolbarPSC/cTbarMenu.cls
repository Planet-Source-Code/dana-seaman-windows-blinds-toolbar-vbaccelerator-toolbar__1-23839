VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cTbarMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' =========================================================================
' cTbarMenu.cls
'
' vbAccelerator Toolbar control
' Copyright © 1998-1999 Steve McMahon (steve@dogma.demon.co.uk)
'
' Code to implement menu tracking against a toolbar
' rather than a menu bar.
'
' Based on code written by Paul DiLascia for Microsoft Systems
' Journal, C++ Q&A, January 1998.  Converted to VB and
' MDI form support added.
'
' -------------------------------------------------------------------------
' Visit vbAccelerator at http://vbaccelerator.com
' =========================================================================

Private Enum TRACKINGSTATE   '{ // menubar has three states:
   TRACK_NONE = 0 ',   // * normal, not tracking anything
   TRACK_BUTTON ',     // * tracking buttons (F10/Alt mode)
   TRACK_POPUP       '// * tracking popups
End Enum

Private Const CONTROLWIN_BUTTON = &H7FFF&

Private m_iTrackingState As TRACKINGSTATE
Private m_bProcessRightArrow As Boolean
Private m_bProcessLeftArrow  As Boolean
Private m_hMenuTracking As Long
Private m_iPopupTracking As Long
Private m_bEscapeWasPressed As Boolean
Private m_tPMouse As POINTAPI
Private m_iNewPopup As Long

Private m_hWnd As Long
Private m_hWndOwner As Long
Private m_hMenu As Long

Implements ISubclass

Friend Sub CoolMenuAttach(ByRef hWndA As Long, ByRef hWndToolbar As Long, ByRef hMenu As Long)
   CoolMenuDetach
   m_hWndOwner = hWndA
   m_hWnd = hWndToolbar
   m_hMenu = hMenu
   SendMessage m_hWndOwner, WM_ENTERMENULOOP, 0, 0
   AttachMessage Me, m_hWndOwner, WM_MENUSELECT
End Sub
Friend Sub CoolMenuDetach()
   If (m_hWndOwner <> 0) Then
      SendMessage m_hWndOwner, WM_EXITMENULOOP, 0, 0
      DetachMessage Me, m_hWndOwner, WM_MENUSELECT
      m_hWndOwner = 0
   End If
   m_hWnd = 0
End Sub
Private Property Get GetMDIChildIsMaximized() As Boolean
Dim hwnd As Long
   hwnd = FindWindowEx(m_hWndOwner, 0, "MDIClient", ByVal 0&) 'GetWindow(m_hWndOwner, GW_CHILD)
   If hwnd <> 0 Then
      hwnd = SendMessageLong(hwnd, WM_MDIGETACTIVE, 0, 0)
      If IsZoomed(hwnd) Then
         GetMDIChildIsMaximized = True
      End If
   End If
End Property

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
Dim tp As POINTAPI
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
      LSet tp = m.pt
      ScreenToClient m_hWnd, tp

      If (iMsg = WM_MOUSEMOVE) Then
         'If (tP.X <> m_tPMouse.X) And (tP.Y <> m_tPMouse.Y) Then
            iButton = HitTest(tp)
            If IsValidButton(iButton) Then
               If iButton <> m_iPopupTracking Then
                  '// user moved mouse over a different button: track its popup
                  CancelMenuAndTrackNewOne iButton
               End If
            End If
            LSet m_tPMouse = tp
         'End If
      ElseIf iMsg = WM_LBUTTONDOWN Then
         If (HitTest(tp) = m_iPopupTracking) Then
            '// user clicked on same button I am tracking: cancel menu
            'MBTRACE(_T("CMenuBar:OnMenuInput: handle mouse click to exit popup\n"));
            CancelMenuAndTrackNewOne -1
            MenuInput = True ' // eat it
         End If
      End If
   
   End If

End Function

Private Function HitTest(pt As POINTAPI) As Long
Dim iHit As Long
Dim rc As RECT

   'int iHit = CFlatToolBar::HitTest(p);
   'if (iHit>0) {
   '   GetClientRect(&rc);
   '   if (!rc.PtInRect(p)) // if point is outside window
   '      iHit = -1;        // can't be a hit!
   '}
   
   GetClientRect m_hWnd, rc
   If (PtInRect(rc, pt.x, pt.y) = 0) Then
      iHit = -1
      If GetMDIChildIsMaximized Then
         GetRect CONTROLWIN_BUTTON, rc
         If pt.x >= rc.Left And pt.x <= rc.Right Then
            ' In sys menu...
            'Debug.Print "HitTest in SysMenu"
            iHit = CONTROLWIN_BUTTON
         End If
      End If
   Else
      Dim iT As Long
      iHit = -1
      For iT = 0 To GetButtonCount - 1
         GetRect iT, rc
         If pt.x >= rc.Left And pt.x <= rc.Right Then
            iHit = iT
            Exit For
         End If
      Next iT
   End If
   HitTest = iHit
   
End Function
Private Property Get IsValidButton(ByVal iButton As Long) As Boolean
   If (iButton > -1) Then
      IsValidButton = True
   End If
End Property

'//////////////////
'// Cancel the current popup menu by posting WM_CANCELMODE, and track a new
'// menu. iNewPopup is which new popup to track (-1 to quit).
'//
Private Sub CancelMenuAndTrackNewOne(ByVal iNewPopup As Long)

   'MBTRACE(_T("CMenuBar::CancelMenuAndTrackNewOne: %d\n"), iNewPopup);
   'ASSERT_VALID(this);
   If (iNewPopup <> m_iPopupTracking) Then
      PostMessage m_hWndOwner, WM_CANCELMODE, 0, 0 ' // quit menu loop
      m_iNewPopup = iNewPopup                '// go to this popup (-1 = quit)
   End If
End Sub

Private Sub PressButton(ByVal iButton As Long, ByVal bState As Boolean)
Dim tBB As TBBUTTON
Dim fState As Long

   If iButton >= 0 And iButton < GetButtonCount() Then
      SendMessage m_hWnd, TB_GETBUTTON, iButton, tBB
      fState = Abs(bState)
      SendMessageLong m_hWnd, TB_PRESSBUTTON, tBB.idCommand, fState
      
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

   Debug.Assert m_hMenu <> 0
   
   nMenuItems = GetMenuItemCount(m_hMenu)

   Do While (iButton >= 0)                '// while user selects another menu
      
      lRtnID = 0

      m_iNewPopup = -1                '// assume quit after this
      PressButton iButton, True       '// press the button
      UpdateWindow m_hWnd             '// and force repaint now

      '// post a simulated arrow-down into the message stream
      '// so TrackPopupMenu will read it and move to the first item
         ' SPM: note this isn't done in the IE/Office menu bar.
      PostMessage m_hWndOwner, WM_KEYDOWN, vbKeyDown, 1
      PostMessage m_hWndOwner, WM_KEYUP, vbKeyDown, 1

      SetTrackingState TRACK_POPUP, iButton '// enter tracking state

      '// Need to install a hook to trap menu input in order to make
      '// left/right-arrow keys and "hot" mouse tracking work.
      '//
      AttachMsgHook Me

      '// get submenu and display it beneath button
      GetRect iButton, rcButton
      ClientRectToScreen m_hWnd, rcButton
      tPM.cbSize = Len(tPM)
      ComputeMenuTrackPoint rcButton, tPM, pt
      
      ' MDI menu:
      If iButton = CONTROLWIN_BUTTON& Then
         ' Get the system menu for this MDI client...
         hwnd = GetProp(m_hWnd, "vbalTbar:MDIClient")
         If hwnd <> 0 Then
            hwnd = SendMessageLong(hwnd, WM_MDIGETACTIVE, 0, 0)
            hMenuPopup = GetSystemMenu(hwnd, 0)
         End If
         'Debug.Print hMenuPopup
      Else
         hMenuPopup = GetSubMenu(m_hMenu, iButton)
      End If
      If hMenuPopup <> 0 Then
         ' Show the menu:
         m_hMenuTracking = hMenuPopup
         lR = TrackPopupMenuEx(hMenuPopup, _
            TPM_LEFTALIGN Or TPM_LEFTBUTTON Or TPM_VERTICAL Or TPM_RETURNCMD, _
            pt.x, pt.y, m_hWndOwner, tPM)
         If iButton = CONTROLWIN_BUTTON Then
            SendMessageLong hwnd, WM_SYSCOMMAND, lR, 0
         Else
            'lR is the ID of the menu
            lRtnID = lR
         End If
      End If
      
      '// uninstall hook.
      DetachMsgHook

      PressButton iButton, False    ';   // un-press button
      UpdateWindow m_hWnd                '// and force repaint now
      
      '// If the user exited the menu loop by pressing Escape,
      '// return to track-button state; otherwise normal non-tracking state.
      If (m_bEscapeWasPressed) Then
         SetTrackingState TRACK_BUTTON, iButton
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

End Function
Private Sub ComputeMenuTrackPoint(ByRef rc As RECT, tPM As TPMPARAMS, tp As POINTAPI)
   tp.x = rc.Left
   tp.y = rc.Bottom
   LSet tPM.rcExclude = rc
End Sub


Private Sub GetRect(ByVal iButton As Long, ByRef tR As RECT)
Dim tBB As TBBUTTON
Dim iB
   If iButton = CONTROLWIN_BUTTON Then
      ' The system menu?
      iB = 0
   Else
      iB = iButton
   End If
   SendMessage m_hWnd, TB_GETBUTTON, iB, tBB
   SendMessage m_hWnd, TB_GETRECT, tBB.idCommand, tR
   If iButton = CONTROLWIN_BUTTON Then
      OffsetRect tR, -(GetSystemMetrics(SM_CYCAPTION) + 4), 0
      tR.Bottom = tR.Bottom - 2
   End If
End Sub
Private Function GetHotItem() As Long
   GetHotItem = SendMessageLong(m_hWnd, TB_GETHOTITEM, 0, 0)
End Function
Private Function SetHotItem(ByVal iButton As Long) As Long
   SendMessageLong m_hWnd, TB_SETHOTITEM, iButton, 0
End Function
Private Function GetButtonVisible(ByVal iButton As Long) As Boolean
Dim tBB As TBBUTTON
   If iButton = CONTROLWIN_BUTTON Then
      GetButtonVisible = GetMDIChildIsMaximized
   Else
      SendMessage m_hWnd, TB_GETBUTTON, iButton, tBB
      GetButtonVisible = (SendMessageLong(m_hWnd, TB_ISBUTTONHIDDEN, tBB.idCommand, 0) = 0)
   End If
End Function
Private Function GetButtonCount() As Long
   GetButtonCount = SendMessageLong(m_hWnd, TB_BUTTONCOUNT, 0, 0)
End Function

Private Sub ClientRectToScreen(ByVal hwnd As Long, ByRef tR As RECT)
Dim tp As POINTAPI
   tp.x = tR.Left: tp.y = tR.Top
   ClientToScreen m_hWnd, tp
   tR.Left = tp.x: tR.Top = tp.y
   tp.x = tR.Right: tp.y = tR.Bottom
   ClientToScreen m_hWnd, tp
   tR.Right = tp.x: tR.Bottom = tp.y
End Sub
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
Dim bFound As Boolean

   If (bPrev) Then
      iSB = iButton
      Do While Not bFound
         
         If GetMDIChildIsMaximized Then
            
            If iButton = 0 Then
               iButton = CONTROLWIN_BUTTON
            ElseIf iButton >= CONTROLWIN_BUTTON Then
               iButton = GetButtonCount() - 1
            Else
               iButton = iButton - 1
            End If
            
         Else
            iButton = iButton - 1
            If iButton < 0 Then
               iButton = GetButtonCount() - 1
            End If
         End If
         
         If Not (GetButtonVisible(iButton)) Then
            If iButton = iSB Then
               iButton = -1
               Exit Do
            End If
         Else
            bFound = True
         End If
         
      Loop
      
   Else
      iSB = iButton
      Do While Not bFound
         If GetMDIChildIsMaximized Then
            If iButton >= GetButtonCount() And iButton < CONTROLWIN_BUTTON Then
               iButton = CONTROLWIN_BUTTON
            ElseIf iButton = CONTROLWIN_BUTTON Then
               iButton = 0
            Else
               iButton = iButton + 1
            End If
         Else
            iButton = iButton + 1
            If (iButton >= GetButtonCount()) Then
               iButton = 0
            End If
         End If
         
         If Not GetButtonVisible(iButton) Then
            If iButton = iSB Then
               iButton = -1
               Exit Do
            End If
         Else
            bFound = True
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
         SetTrackingState TRACK_BUTTON, 0
      Else
         SetTrackingState TRACK_NONE, 0
     End If
   End If
End Sub

Private Property Let ISubClass_MsgResponse(ByVal RHS As EMsgResponse)
   '
End Property

Private Property Get ISubClass_MsgResponse() As EMsgResponse
   If CurrentMessage = WM_MENUSELECT Then
      ISubClass_MsgResponse = emrPreprocess
   End If
End Property

Private Function ISubClass_WindowProc(ByVal hwnd As Long, ByVal iMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
   Select Case iMsg
   Case WM_MENUSELECT
      MenuSelect lParam, (wParam And &HFFFF&)
   End Select
End Function

