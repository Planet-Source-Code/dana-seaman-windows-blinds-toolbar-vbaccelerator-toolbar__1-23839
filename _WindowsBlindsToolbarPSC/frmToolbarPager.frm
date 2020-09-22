VERSION 5.00
Begin VB.Form frmToolbarPager 
   Caption         =   "VbAccelerator Windows Blinds Toolbar"
   ClientHeight    =   4410
   ClientLeft      =   60
   ClientTop       =   285
   ClientWidth     =   8880
   LinkTopic       =   "Form1"
   ScaleHeight     =   4410
   ScaleWidth      =   8880
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   480
      Top             =   120
   End
   Begin VB.TextBox Text1 
      Height          =   2595
      Left            =   0
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   2
      Text            =   "frmToolbarPager.frx":0000
      Top             =   1380
      Width           =   8415
   End
   Begin VbAccelToolbarPager.cToolbar cToolbar1 
      Left            =   2580
      Top             =   0
      _ExtentX        =   3731
      _ExtentY        =   1402
   End
   Begin VbAccelToolbarPager.cReBar cReBar1 
      Left            =   1320
      Top             =   0
      _ExtentX        =   2249
      _ExtentY        =   1402
   End
   Begin VbAccelToolbarPager.cPager cPager1 
      Height          =   810
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   1429
      AutoHScroll     =   -1  'True
   End
   Begin VB.PictureBox Preview 
      BackColor       =   &H80000005&
      Height          =   810
      Left            =   4680
      ScaleHeight     =   750
      ScaleWidth      =   1500
      TabIndex        =   0
      ToolTipText     =   "Image Preview"
      Top             =   0
      Width           =   1560
   End
   Begin VB.Image Image1 
      Height          =   540
      Index           =   3
      Left            =   7740
      Picture         =   "frmToolbarPager.frx":0006
      Top             =   240
      Visible         =   0   'False
      Width           =   6480
   End
   Begin VB.Image Image1 
      Height          =   540
      Index           =   2
      Left            =   0
      Picture         =   "frmToolbarPager.frx":4108
      Top             =   840
      Visible         =   0   'False
      Width           =   6480
   End
   Begin VB.Image Image1 
      Height          =   750
      Index           =   5
      Left            =   6240
      Picture         =   "frmToolbarPager.frx":820A
      Top             =   0
      Visible         =   0   'False
      Width           =   1500
   End
   Begin VB.Image Image1 
      Height          =   480
      Index           =   4
      Left            =   6480
      Picture         =   "frmToolbarPager.frx":9234
      Top             =   840
      Visible         =   0   'False
      Width           =   5760
   End
   Begin VB.Image Image1 
      Height          =   240
      Index           =   1
      Left            =   7740
      Picture         =   "frmToolbarPager.frx":A1BF
      Top             =   0
      Visible         =   0   'False
      Width           =   20160
   End
   Begin VB.Image Image1 
      Height          =   1920
      Index           =   0
      Left            =   -60
      Picture         =   "frmToolbarPager.frx":BD4B
      Top             =   2460
      Visible         =   0   'False
      Width           =   1920
   End
   Begin VB.Menu mnuFileTOP 
      Caption         =   "File"
      Tag             =   "1605"
      Begin VB.Menu mnuFile 
         Caption         =   ""
         Index           =   0
         Tag             =   "1421"
      End
      Begin VB.Menu mnuFile 
         Caption         =   ""
         Index           =   1
         Tag             =   "1422"
      End
      Begin VB.Menu mnuFile 
         Caption         =   ""
         Index           =   2
         Tag             =   "1135"
      End
      Begin VB.Menu mnuFile 
         Caption         =   ""
         Index           =   3
         Tag             =   "1142"
      End
      Begin VB.Menu mnuFile 
         Caption         =   ""
         Index           =   4
         Tag             =   "1144"
      End
      Begin VB.Menu mnuFile 
         Caption         =   "-"
         Index           =   5
      End
      Begin VB.Menu mnuFile 
         Caption         =   ""
         Index           =   6
         Tag             =   "1640"
      End
      Begin VB.Menu mnuFile 
         Caption         =   ""
         Index           =   7
         Tag             =   "1641"
      End
      Begin VB.Menu mnuFile 
         Caption         =   "-"
         Index           =   8
      End
      Begin VB.Menu mnuFile 
         Caption         =   "&UUEncode"
         Index           =   9
      End
      Begin VB.Menu mnuFile 
         Caption         =   "UUDec&ode"
         Index           =   10
      End
      Begin VB.Menu mnuFile 
         Caption         =   "-"
         Index           =   11
      End
      Begin VB.Menu mnuFile 
         Caption         =   "&NotePad ... WordPad"
         Index           =   12
      End
      Begin VB.Menu mnuFile 
         Caption         =   "Shortcut"
         Index           =   13
      End
      Begin VB.Menu mnuFile 
         Caption         =   "E&xit"
         Index           =   14
      End
   End
   Begin VB.Menu mnuEditTOP 
      Caption         =   "&Clipboard"
      Begin VB.Menu mnuEdit 
         Caption         =   "N"
         Index           =   0
         Tag             =   "1400"
      End
      Begin VB.Menu mnuEdit 
         Caption         =   "P"
         Index           =   1
         Tag             =   "1401"
      End
      Begin VB.Menu mnuEdit 
         Caption         =   "N + P"
         Index           =   2
         Tag             =   "1402"
      End
      Begin VB.Menu mnuEdit 
         Caption         =   "-"
         Index           =   3
      End
      Begin VB.Menu mnuEdit 
         Caption         =   "dos N"
         Index           =   4
         Tag             =   "1403"
      End
      Begin VB.Menu mnuEdit 
         Caption         =   "dos P"
         Index           =   5
         Tag             =   "1404"
      End
      Begin VB.Menu mnuEdit 
         Caption         =   "dos N + P"
         Index           =   6
         Tag             =   "1405"
      End
      Begin VB.Menu mnuEdit 
         Caption         =   "-"
         Index           =   7
      End
      Begin VB.Menu mnuEdit 
         Caption         =   "SelTab"
         Index           =   8
         Tag             =   "1318"
      End
      Begin VB.Menu mnuEdit 
         Caption         =   "SelCom"
         Index           =   9
         Tag             =   "1319"
      End
      Begin VB.Menu mnuEdit 
         Caption         =   "-"
         Index           =   10
      End
      Begin VB.Menu mnuEdit 
         Caption         =   "AllTab"
         Index           =   11
         Tag             =   "1320"
      End
      Begin VB.Menu mnuEdit 
         Caption         =   "AllCom"
         Index           =   12
         Tag             =   "1321"
      End
   End
   Begin VB.Menu mnuViewTOP 
      Caption         =   "View"
      Tag             =   "1014"
      Begin VB.Menu mnuView 
         Caption         =   ""
         Index           =   0
         Tag             =   "1419"
      End
      Begin VB.Menu mnuView 
         Caption         =   ""
         Index           =   1
         Tag             =   "1420"
      End
      Begin VB.Menu mnuView 
         Caption         =   ""
         Index           =   2
         Tag             =   "1424"
      End
      Begin VB.Menu mnuView 
         Caption         =   "&Kilobytes"
         Index           =   3
      End
      Begin VB.Menu mnuView 
         Caption         =   "Tips"
         Index           =   4
         Tag             =   "1327"
      End
      Begin VB.Menu mnuView 
         Caption         =   "Sounds"
         Index           =   5
         Tag             =   "1426"
      End
   End
   Begin VB.Menu mnuToolsTOP 
      Caption         =   "Tools"
      Tag             =   "1606"
      Begin VB.Menu mnuTools 
         Caption         =   "Dos"
         Index           =   0
         Tag             =   "1005"
      End
      Begin VB.Menu mnuTools 
         Caption         =   "Fmt Disk"
         Index           =   1
         Tag             =   "1322"
      End
      Begin VB.Menu mnuTools 
         Caption         =   "Sync Atom"
         Index           =   2
         Tag             =   "1323 "
      End
      Begin VB.Menu mnuTools 
         Caption         =   "Build FS Db"
         Index           =   3
         Tag             =   "1324"
      End
      Begin VB.Menu mnuTools 
         Caption         =   "FF"
         Index           =   4
         Tag             =   "1325 "
      End
      Begin VB.Menu mnuTools 
         Caption         =   "DC"
         Index           =   5
         Tag             =   "1326"
      End
   End
   Begin VB.Menu mnuCtrlTOP 
      Caption         =   "Ctrl"
      Tag             =   "1607"
      Begin VB.Menu mnuCtrl 
         Caption         =   ""
         Index           =   0
      End
      Begin VB.Menu mnuCtrl 
         Caption         =   ""
         Index           =   1
      End
      Begin VB.Menu mnuCtrl 
         Caption         =   ""
         Index           =   2
      End
      Begin VB.Menu mnuCtrl 
         Caption         =   ""
         Index           =   3
      End
      Begin VB.Menu mnuCtrl 
         Caption         =   ""
         Index           =   4
      End
      Begin VB.Menu mnuCtrl 
         Caption         =   ""
         Index           =   5
      End
      Begin VB.Menu mnuCtrl 
         Caption         =   ""
         Index           =   6
      End
      Begin VB.Menu mnuCtrl 
         Caption         =   ""
         Index           =   7
      End
      Begin VB.Menu mnuCtrl 
         Caption         =   "&Modem"
         Index           =   8
      End
      Begin VB.Menu mnuCtrl 
         Caption         =   ""
         Index           =   9
      End
      Begin VB.Menu mnuCtrl 
         Caption         =   ""
         Index           =   10
      End
      Begin VB.Menu mnuCtrl 
         Caption         =   ""
         Index           =   11
      End
      Begin VB.Menu mnuCtrl 
         Caption         =   ""
         Index           =   12
      End
      Begin VB.Menu mnuCtrl 
         Caption         =   ""
         Index           =   13
      End
      Begin VB.Menu mnuCtrl 
         Caption         =   ""
         Index           =   14
      End
      Begin VB.Menu mnuCtrl 
         Caption         =   ""
         Index           =   15
      End
   End
   Begin VB.Menu mnuDriveTOP 
      Caption         =   "Drv"
      Tag             =   "1608"
      Begin VB.Menu mnuDrive 
         Caption         =   "VolLbl"
         Index           =   0
         Tag             =   "1328"
      End
      Begin VB.Menu mnuDrive 
         Caption         =   "-"
         Index           =   1
      End
      Begin VB.Menu mnuDrive 
         Caption         =   "MapNetDrv"
         Index           =   2
         Tag             =   "1329"
      End
      Begin VB.Menu mnuDrive 
         Caption         =   "UnMapNetDrv"
         Index           =   3
         Tag             =   "1330"
      End
   End
   Begin VB.Menu mnuLangTOP 
      Caption         =   "Lang"
      Tag             =   "1146"
      Begin VB.Menu mnuLang 
         Caption         =   "&Deutsch"
         Index           =   0
      End
      Begin VB.Menu mnuLang 
         Caption         =   "&English"
         Index           =   1
      End
      Begin VB.Menu mnuLang 
         Caption         =   "E&spanhõl"
         Index           =   2
      End
      Begin VB.Menu mnuLang 
         Caption         =   "&Français"
         Index           =   3
      End
      Begin VB.Menu mnuLang 
         Caption         =   "&Italiano"
         Index           =   4
      End
      Begin VB.Menu mnuLang 
         Caption         =   "&Português (brasileiro)"
         Index           =   5
      End
   End
   Begin VB.Menu mnuHelpTOP 
      Caption         =   "Hlp"
      Tag             =   "1100"
      Begin VB.Menu mnuHelp 
         Caption         =   "&Con"
         Index           =   0
         Shortcut        =   {F1}
         Tag             =   "1250"
      End
      Begin VB.Menu mnuHelp 
         Caption         =   "&Srch"
         Index           =   1
         Tag             =   "1251"
      End
      Begin VB.Menu mnuHelp 
         Caption         =   "&Idx"
         Index           =   2
         Tag             =   "1252"
      End
      Begin VB.Menu mnuHelp 
         Caption         =   "-"
         Index           =   3
      End
      Begin VB.Menu mnuHelp 
         Caption         =   "&Abt"
         Index           =   4
         Tag             =   "1002"
      End
   End
   Begin VB.Menu mnuCopyTOP 
      Caption         =   "Copy"
      Visible         =   0   'False
      Begin VB.Menu mnuCopy 
         Caption         =   "Cop"
         Index           =   0
         Shortcut        =   {F7}
         Tag             =   "1041"
      End
      Begin VB.Menu mnuCopy 
         Caption         =   "Smart"
         Index           =   1
         Shortcut        =   {F8}
         Tag             =   "1042"
      End
      Begin VB.Menu mnuCopy 
         Caption         =   "Mov"
         Index           =   2
         Shortcut        =   {F9}
         Tag             =   "1108"
      End
      Begin VB.Menu mnuCopy 
         Caption         =   "-"
         Index           =   3
      End
      Begin VB.Menu mnuCopy 
         Caption         =   "FTP &Upload"
         Index           =   4
      End
      Begin VB.Menu mnuCopy 
         Caption         =   "FTP &Download"
         Index           =   5
      End
   End
   Begin VB.Menu mnuSortTOP 
      Caption         =   "Sort"
      Visible         =   0   'False
      Begin VB.Menu mnuSort 
         Caption         =   ""
         Index           =   0
      End
      Begin VB.Menu mnuSort 
         Caption         =   ""
         Index           =   1
         Tag             =   "1521"
      End
      Begin VB.Menu mnuSort 
         Caption         =   ""
         Index           =   2
         Tag             =   "1522"
      End
      Begin VB.Menu mnuSort 
         Caption         =   ""
         Index           =   3
         Tag             =   "1523"
      End
      Begin VB.Menu mnuSort 
         Caption         =   ""
         Index           =   4
         Tag             =   "1525"
      End
      Begin VB.Menu mnuSort 
         Caption         =   ""
         Index           =   5
         Tag             =   "1524"
      End
      Begin VB.Menu mnuSort 
         Caption         =   ""
         Index           =   6
         Tag             =   "1526"
      End
      Begin VB.Menu mnuSort 
         Caption         =   ""
         Index           =   7
         Tag             =   "1527"
      End
      Begin VB.Menu mnuSort 
         Caption         =   ""
         Index           =   8
      End
      Begin VB.Menu mnuSort 
         Caption         =   ""
         Index           =   9
      End
      Begin VB.Menu mnuSort 
         Caption         =   ""
         Index           =   10
         Tag             =   "1528"
      End
   End
   Begin VB.Menu mnuSortZipTOP 
      Caption         =   "SortZip"
      Visible         =   0   'False
      Begin VB.Menu mnuSortZip 
         Caption         =   ""
         Index           =   0
         Tag             =   "1520"
      End
      Begin VB.Menu mnuSortZip 
         Caption         =   ""
         Index           =   1
         Tag             =   "1521"
      End
      Begin VB.Menu mnuSortZip 
         Caption         =   ""
         Index           =   2
         Tag             =   "1522"
      End
      Begin VB.Menu mnuSortZip 
         Caption         =   ""
         Index           =   3
         Tag             =   "1523"
      End
      Begin VB.Menu mnuSortZip 
         Caption         =   ""
         Index           =   4
         Tag             =   "1525"
      End
      Begin VB.Menu mnuSortZip 
         Caption         =   ""
         Index           =   5
         Tag             =   "1524"
      End
      Begin VB.Menu mnuSortZip 
         Caption         =   ""
         Index           =   6
         Tag             =   "1712"
      End
      Begin VB.Menu mnuSortZip 
         Caption         =   ""
         Index           =   7
         Tag             =   "1713"
      End
      Begin VB.Menu mnuSortZip 
         Caption         =   ""
         Index           =   8
         Tag             =   "1526"
      End
      Begin VB.Menu mnuSortZip 
         Caption         =   ""
         Index           =   9
         Tag             =   "1527"
      End
      Begin VB.Menu mnuSortZip 
         Caption         =   ""
         Index           =   10
         Tag             =   "1714"
      End
      Begin VB.Menu mnuSortZip 
         Caption         =   ""
         Index           =   11
         Tag             =   "1715"
      End
      Begin VB.Menu mnuSortZip 
         Caption         =   "CRC"
         Index           =   12
      End
      Begin VB.Menu mnuSortZip 
         Caption         =   ""
         Index           =   13
         Tag             =   "1716"
      End
   End
   Begin VB.Menu mnuZipTOP 
      Caption         =   "Zip"
      Visible         =   0   'False
      Begin VB.Menu mnuZip 
         Caption         =   ""
         Index           =   0
         Tag             =   "1550"
         Begin VB.Menu mnuZipAdd 
            Caption         =   "Ovr"
            Index           =   0
            Tag             =   "1239"
         End
         Begin VB.Menu mnuZipAdd 
            Caption         =   "Adv"
            Index           =   1
            Tag             =   "1240"
         End
      End
      Begin VB.Menu mnuZip 
         Caption         =   ""
         Index           =   1
         Tag             =   "1551"
      End
      Begin VB.Menu mnuZip 
         Caption         =   ""
         Index           =   2
         Tag             =   "1552"
      End
   End
   Begin VB.Menu mnuSelectTOP 
      Caption         =   "Select"
      Visible         =   0   'False
      Begin VB.Menu mnuSelect 
         Caption         =   ""
         Index           =   0
         Tag             =   "1031"
      End
      Begin VB.Menu mnuSelect 
         Caption         =   ""
         Index           =   1
         Tag             =   "1032"
      End
      Begin VB.Menu mnuSelect 
         Caption         =   ""
         Index           =   2
         Tag             =   "1033"
      End
      Begin VB.Menu mnuSelect 
         Caption         =   "-"
         Index           =   3
      End
      Begin VB.Menu mnuSelect 
         Caption         =   ""
         Index           =   4
         Tag             =   "1035"
      End
      Begin VB.Menu mnuSelect 
         Caption         =   ""
         Index           =   5
         Tag             =   "1036"
      End
   End
   Begin VB.Menu mnuPopTOP 
      Caption         =   "Pop"
      Visible         =   0   'False
      Begin VB.Menu mnuPop 
         Caption         =   "C"
         Index           =   0
         Tag             =   "1041"
      End
      Begin VB.Menu mnuPop 
         Caption         =   "CI"
         Index           =   1
         Tag             =   "1042"
      End
      Begin VB.Menu mnuPop 
         Caption         =   "M"
         Index           =   2
         Tag             =   "1108"
      End
      Begin VB.Menu mnuPop 
         Caption         =   "R"
         Index           =   3
         Tag             =   "1111"
      End
      Begin VB.Menu mnuPop 
         Caption         =   "D"
         Index           =   4
         Tag             =   "1107"
      End
      Begin VB.Menu mnuPop 
         Caption         =   "S"
         Index           =   5
         Tag             =   "1118"
      End
      Begin VB.Menu mnuPop 
         Caption         =   "N"
         Index           =   6
         Tag             =   "1109"
      End
      Begin VB.Menu mnuPop 
         Caption         =   "P"
         Index           =   7
         Tag             =   "1012"
      End
      Begin VB.Menu mnuPop 
         Caption         =   "D"
         Index           =   8
         Tag             =   "1142"
      End
      Begin VB.Menu mnuPop 
         Caption         =   "A"
         Index           =   9
         Tag             =   "1144"
      End
      Begin VB.Menu mnuPop 
         Caption         =   "Z"
         Index           =   10
         Tag             =   "1110"
      End
      Begin VB.Menu mnuPop 
         Caption         =   "CS"
         Index           =   11
         Tag             =   "1028"
      End
   End
End
Attribute VB_Name = "frmToolbarPager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Private ControlArray    As Variant
Private sFolder         As String
Private sFile           As String
Private sName           As String
Private sExtension      As String
Private sSize           As String
Private sType           As String
Private sModified       As String
Private sTime           As String
Private sCreated        As String
Private sAccessed       As String
Private sAttribute      As String
Private sMsDos          As String
Private sNone           As String
Private HTMLPage(4)     As String
'--------------------------------------------
Private cIM             As New cIconMenu
Public m_cIL16          As New cVBALImageList
Private m_cIL36         As New cVBALImageList
Private m_cIL36HOT      As New cVBALImageList
Private m_cIL32DIS      As New cVBALImageList
Private Sub PrepareImageLists()

   Dim pic As StdPicture
  
   Set pic = Image1(1) 'MenuIcons 16x16
   With m_cIL16
      .ColourDepth = &H8
      .IconSizeX = 16
      .IconSizeY = 16
      .Create
      .AddFromHandle pic.Handle, IMAGE_BITMAP, , &H4080C0
   End With
   
   Set pic = Image1(2) 'Toolbar 36x36 Normal
   With m_cIL36
      .ColourDepth = &H18
      .IconSizeX = 36
      .IconSizeY = 36
      .Create
      .AddFromHandle pic.Handle, IMAGE_BITMAP, , &HFF00FF
   End With

   Set pic = Image1(3) 'Toolbar 36x32 HOT
   With m_cIL36HOT
      .ColourDepth = &H18
      .IconSizeX = 36
      .IconSizeY = 36
      .Create
      .AddFromHandle pic.Handle, IMAGE_BITMAP, , &HFF00FF
   End With

   Set pic = Image1(4) 'Toolbar 32x32 DISabled
   ' 32x32 is roughly the size of the image
   ' painted on each button
   With m_cIL32DIS
      .ColourDepth = &H18
      .IconSizeX = 32
      .IconSizeY = 32
      .Create
      .AddFromHandle pic.Handle, IMAGE_BITMAP, , -1
   End With
   
End Sub

Private Sub cPager1_RequestSize(lWidth As Long, lHeight As Long)
    ' We only need to return the width because the
    ' pager is horizontal:
    lWidth = cToolbar1.ToolbarWidth
End Sub

Private Sub cPager1_Scroll(ByVal eDir As ECPGScrollDir, lDelta As Long)
   lDelta = 8
End Sub

Private Sub cReBar1_HeightChanged(lNewHeight As Long)
  SizeControls
End Sub

Private Sub cToolbar1_DropDownPress(ByVal lButton As Long)
  On Error GoTo ProcedureError
   Dim x As Long, y As Long

   cToolbar1.GetDropDownPosition lButton, x, y

   Select Case Mid$(cToolbar1.ButtonKey(lButton), 5)
      Case "copy"
         Me.PopupMenu mnuCopyTOP, , x, y
      Case "zip"
         Me.PopupMenu mnuZipTOP, , x, y
      Case "select"
         Me.PopupMenu mnuSelectTOP, , x, y
      Case "lang"
         Me.PopupMenu mnuLangTOP, , x, y
      Case "sort"
         If InZip Then
            Me.PopupMenu mnuSortZipTOP, , x, y
         Else
            Me.PopupMenu mnuSortTOP, , x, y
         End If

   End Select

ProcedureExit:
  Exit Sub
ProcedureError:
  If ErrMsgBox(Me.Name & ".cToolbar1_DropDownPress") = vbRetry Then Resume Next

End Sub

Private Sub Form_Load()

   ControlArray = Split("sysdm.cpl @1|appwiz.cpl @1|timedate.cpl|desk.cpl|main.cpl @3|inetcpl.cpl|joy.cpl|main.cpl @1|modem.cpl|main.cpl|mmsys.cpl|netcpl.cpl|password.cpl|main.cpl @2|intl.cpl|sysdm.cpl", "|")

'---- Rip resource strings from Windows Dll's ----
  
   sFolder = "(" & GetResourceString(4131) & ")"
   sFile = GetResourceString(4130) & " "
   sName = GetResourceString(8976)
   sExtension = "Ext" 'can't find this, use resource file
   sSize = GetResourceString(8978)
   sType = GetResourceString(8979)
   sModified = GetResourceString(8980)
   sTime = GetResourceStringFromFile("Intl.Cpl", 25)
   sCreated = GetResourceString(8996)
   sAccessed = GetResourceString(8997)
   sAttribute = GetResourceString(8987)
   sMsDos = "MsDos"
   sNone = GetResourceString(9808)

   PrepareImageLists 'Load picture strips into imagelist classes
   With cIM
      .Attach Me.hwnd
      .HighlightStyle = ECPHighlightStyleGradient
      'Following 2 properties are NOT in original control
      .GradientStartColor = vbYellow
      .GradientStopColor = &H80FF& 'Orange
      Set .BackgroundPicture = Image1(0).Picture
      .ImageList = m_cIL16.himl 'was vbalImageList1
   End With

   Set Preview.Picture = Image1(5).Picture
   
   '** Defaults to English (1000) if no registry entry
   Lang = GetSetting(App.Title, "Settings", "Lang", 1000)

   CreateToolbar
   mnuLang(Lang \ c1000).Checked = True
   UpdateLanguage
   
   Text1.Text = "This Demo implements VbAccelerator:" & vbCrLf & _
                "1. Toolbar, Rebar, Pager, Imagelist class, Subclassing, Pop Icon menus." & vbCrLf & _
                "2. Icon menus with background and customizable gradient highlight." & vbCrLf & _
                "3. Image strips made from button control and Windows Blinds style buttons." & vbCrLf & _
                "4. Strips are then attached to instances of Imagelist class." & vbCrLf & _
                "5. Swap 6 languages on the fly." & vbCrLf & _
                "6. 16 Control panel applets available from menu" & vbCrLf & _
                "7. VbAccelerator => Minor bugs fixed, properties added, code enhanced." & vbCrLf & _
                "8. When not using analog meter as progress indicator it's container can double as a thumbnail display !"
                
End Sub
Private Sub UpdateLanguage()
   On Error GoTo ProcedureError
   Dim j As Long, L4 As Long

   '-- first update toolbar and misc controls

SetControlCaptionStrings Me    ' indexed by language
   
'With Grid
'   If .Columns Then
'      .Redraw = False
'      .ColumnHeader("ext") = sExtension
'      .ColumnHeader("siz") = sSize
'      If InDoy Then
'         .ColumnHeader("dat") = GetResourceString(1067) 'ddd
'      Else
'         .ColumnHeader("dat") = sModified
'      End If
'      .ColumnHeader("tim") = sTime
'      .ColumnHeader("typ") = sType
'      If InZip Then
'         .ColumnHeader("nam") = sName
'         .ColumnHeader("cmp") = GetResourceString(1068)
'         .ColumnHeader("rat") = GetResourceString(1069)
'         .ColumnHeader("mtd") = GetResourceString(1070)
'         .ColumnHeader("enc") = GetResourceString(1308)
'         .ColumnHeader("pth") = GetResourceString(1030)
'         For L4 = 1 To .Rows
'            'ItemData"mtd" hold method, ItemData"enc" holds BitFlags
'            .CellText(L4, .ColumnIndex("mtd")) = MethodVerbose(.CellItemData(L4, .ColumnIndex("mtd")), .CellItemData(L4, .ColumnIndex("enc")))
'            If .CellIcon(L4, .ColumnIndex("enc")) = 2 Then
'               .CellText(L4, .ColumnIndex("enc")) = GetResourceString(1309)
'            Else
'               .CellText(L4, .ColumnIndex("enc")) = GetResourceString(1310)
'            End If
'         Next
'      Else
'         .ColumnHeader("nam") = sName '& " " & Filter
'         .ColumnHeader("atr") = sAttribute
'         .ColumnHeader("cre") = sCreated
'         .ColumnHeader("acc") = sAccessed
'         .ColumnHeader("dos") = sMsDos
'      End If
'      .Redraw = True
'   End If
'End With
'------------------------------
With cIM
      
   For L4 = 0 To 1
      .IconIndex(mnuZipAdd(L4).Caption) = 27
   Next
   
   For L4 = 0 To 2
      .IconIndex(mnuZip(L4).Caption) = 27 + L4
    Next
    
   For L4 = 0 To 3
      .IconIndex(mnuDrive(L4).Caption) = Choose(L4 + 1, 75, -1, 57, 58)
   Next

   For L4 = 0 To 4
      .IconIndex(mnuHelp(L4).Caption) = Choose(L4 + 1, 49, 55, 56, -1, 42)
   Next

   For L4 = 0 To 5
      .IconIndex(mnuView(L4).Caption) = Choose(L4 + 1, 32, 26, 46, 63, 42, 23)
      .IconIndex(mnuLang(L4).Caption) = 36 + L4
      .IconIndex(mnuCopy(L4).Caption) = Choose(L4 + 1, 25, 42, 25, -1, 66, 65)
      .IconIndex(mnuSelect(L4).Caption) = Choose(L4 + 1, 31, 50, 52, -1, 31, 50)
      .IconIndex(mnuTools(L4).Caption) = Choose(L4 + 1, 47, 53, 10, 63, 55, 53)
   Next

   mnuSort(0).Caption = sNone
   mnuSort(8).Caption = sCreated
   mnuSort(9).Caption = sAccessed
   For L4 = 0 To 10
      .IconIndex(mnuSort(L4).Caption) = Choose(L4 + 1, 47, 31, 78, 76, 63, 77, 32, 33, 10, 32, 73)
   Next

   For L4 = 0 To 11
      j = Choose(L4 + 1, 25, 42, 25, 50, 28, 64, 75, 75, 10, 73, 35, -1)
      If j > 0 Then .IconIndex(mnuPop(L4).Caption) = j
   Next

   For L4 = 0 To 12
      .IconIndex(mnuEdit(L4).Caption) = Choose(L4 + 1, 44, 44, 44, -1, 47, 47, 47, -1, 31, 31, -1, 79, 79)
      .IconIndex(mnuFile(L4).Caption) = Choose(L4 + 1, 45, 43, 25, 10, 73, -1, 61, 62, -1, 61, 62, -1, 60)
   Next

   For L4 = 0 To 13
      .IconIndex(mnuSortZip(L4).Caption) = Choose(L4 + 1, 47, 31, 78, 76, 63, 77, 35, 34, 32, 33, 35, 20, -1, 67)
   Next
   'Some *.cpl files are 16-bit (even in Win ME)
   'so we can't rip those resources without Thunking
   mnuCtrl(0).Caption = GetResourceString(1610)   'Add new hardware
   ExtractMenuCaption 1, 2001                'Add/Remove Programs
   ExtractMenuCaption 2, 300                 'Time/Date
   ExtractMenuCaption 3, 100                 'Display
   ExtractMenuCaption 4, 106                 'Fonts
   ExtractMenuCaption 5, 4312                'Internet
   ExtractMenuCaption 6, 1076                'Game
   ExtractMenuCaption 7, 102                 'Kybd
   mnuCtrl(8).Caption = "Modems"             'Modems
   ExtractMenuCaption 9, 100                 'Mouse
   ExtractMenuCaption 10, 4867               'Sounds & MM
   mnuCtrl(11).Caption = GetResourceString(1621)           'Network
   ExtractMenuCaption 12, 2002               'Password
   ExtractMenuCaption 13, 104                'Printer
   ExtractMenuCaption 14, 1                  'Regional Settings
   mnuCtrl(15).Caption = GetResourceString(1625)            'System

   For L4 = 0 To 15
      j = Choose(L4 + 1, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 23, 19, 20, 21, 22, 24)
      .IconIndex(mnuCtrl(L4).Caption) = j
   Next

End With

   SaveSetting App.Title, "Settings", "Language", Lang

ProcedureExit:
  Exit Sub
ProcedureError:
  If ErrMsgBox(Me.Name & ".UpdateLanguagee") = vbRetry Then Resume Next

End Sub
Private Function sPct(lNum As Long) As String
   sPct = "%" & Format$(lNum, "0#")
End Function
Private Function PadLeft(ByVal Item As String, ByVal Num As Integer) As String
   PadLeft = right$(Space$(Num) & Item, Num) & ": "
End Function
Private Sub ExtractMenuCaption(Index As Long, Key As Long)
   Dim CPL As String, L4 As Long
   CPL = ControlArray(Index)
   L4 = InStr(CPL, ".")
   If L4 Then
      CPL = left$(CPL, L4 + 3)
      mnuCtrl(Index).Caption = "&" & GetResourceStringFromFile(CPL, Key)
   End If
End Sub
Private Sub CreateToolbar()
   
   On Error GoTo ProcedureError
   Dim pic As StdPicture
   
   With cToolbar1
      .ImageSource = CTBExternalImageList
      .SetImageList m_cIL36.himl, CTBImageListNormal 'Normal
      .SetImageList m_cIL36HOT.himl, CTBImageListHot 'Rollover
      .SetImageList m_cIL32DIS.himl, CTBImageListDisabled 'Greyscale
         
      .CreateToolbar 36, , True, True
      
      .AddButton GetResourceString(1136), 0, , , GetResourceString(1106), CTBDropDownArrow Or CTBAutoSize, "1106copy"
     ' .AddButton GetResourceString(1138), 1, , , GetResourceString(1108), CTBNormal Or CTBAutoSize, "1108move"
      .AddButton GetResourceString(1141), 2, , , GetResourceString(1111), CTBNormal Or CTBAutoSize, "1111recycle"
      .AddButton GetResourceString(1137), 3, , , GetResourceString(1107), CTBNormal Or CTBAutoSize, "1107delete"
      .AddButton GetResourceString(1148), 4, , , GetResourceString(1118), CTBNormal Or CTBAutoSize, "1118shred"
      .AddButton GetResourceString(1139), 5, , , GetResourceString(1109), CTBNormal Or CTBAutoSize, "1109rename"
      .AddButton GetResourceString(1143), 11, , , GetResourceString(1113), CTBNormal Or CTBAutoSize, "1113vault"
      .AddButton GetResourceString(1140), 6, , , GetResourceString(1110), CTBDropDownArrow Or CTBAutoSize, "1110zip"
      .AddButton VbZlStr, -1, , , , CTBSeparator
      .AddButton GetResourceString(1424), 7, , , GetResourceString(1394), CTBCheck Or CTBAutoSize, "1394assoc"
      .AddButton GetResourceString(1147), 8, , , GetResourceString(1117), CTBDropDownArrow Or CTBAutoSize, "1117sort"
      .AddButton GetResourceString(1134), 9, , , GetResourceString(1104), CTBDropDownArrow Or CTBAutoSize, "1104select"
      .AddButton GetResourceString(1133), 10, , , GetResourceString(1103), CTBNormal Or CTBAutoSize, "1103prop"
  
  '   .AddButton GetResourceString(1142), 1, , , "Sort", CTBCheck Or CTBAutoSize, "1105clone"
  
  '   .AddButton GetResourceString(1135), 5,  , , GetResourceString(1105), CTBNormal Or CTBAutoSize, "1105clone"
  '   .AddButton GetResourceString(1142), 1, , , GetResourceString(1112), CTBNormal Or CTBAutoSize, "1112date"
  '   .AddButton GetResourceString(1425), 11, , , GetResourceString(1395), CTBCheck Or CTBAutoSize, "1395detail"
  '   .AddButton vbzlstr, -1, , , , CTBSeparator
  '   .AddButton GetResourceString(1142), 1, , , GetResourceString(1112), CTBNormal Or CTBAutoSize, "1112date"
  '   .AddButton GetResourceString(1130), 10, , , GetResourceString(1100), CTBDropDownArrow Or CTBAutoSize, "1100help"
  '   .AddButton vbzlstr, -1, , , , CTBSeparator
   End With
   
   With cPager1 'Add toolbar to Pager
      .AddChildWindow cToolbar1.hwnd
      .Height = (cToolbar1.ToolbarHeight + 4) * Screen.TwipsPerPixelY
      .top = -2 * .Height
      .TabStop = False
   End With
   
   LockWindowUpdate Me.hwnd
   
   With cReBar1 'Attach Pager and Preview PicBox to Rebar
      .ImageSource = CRBPicture
      Set pic = Image1(0)  'LoadResPicture(lColourID, vbResBitmap)
      .ImagePicture = pic.Handle
      .CreateRebar Me.hwnd
      .AddBandByHwnd cPager1.PagerhWnd, , , , "ToolbarBand"
      .AddBandByHwnd Preview.hwnd, , False, True, "View"
      '.BandChildMinWidth(.BandIndexForData("ToolbarBand")) = 64
   End With
   
   LockWindowUpdate 0

ProcedureExit:
  Exit Sub
ProcedureError:
  If ErrMsgBox(Me.Name & ".CreateToolbar") = vbRetry Then Resume Next

End Sub

Private Sub SizeControls()

   Dim RH As Integer
   On Error GoTo ProcedureError
   cReBar1.RebarSize
  ' RH = cReBar1.RebarHeight * Screen.TwipsPerPixelY
  ' Splitter1.Move 0, RH, Me.ScaleWidth, Me.ScaleHeight - (RH + vbalSB.Height)

ProcedureExit:
  Exit Sub
ProcedureError:
  If ErrMsgBox(Me.Name & ".SizeControls") = vbRetry Then Resume Next

End Sub
Private Sub Form_Resize()

   On Error GoTo ProcedureError
 
   If Me.WindowState <> vbMinimized Then
      If Me.Width < 3000 Then
         Me.Width = 3000
      End If
      SizeControls
   End If
 
ProcedureExit:
  Exit Sub
ProcedureError:
  If ErrMsgBox(Me.Name & ".Form_Resize") = vbRetry Then Resume Next
End Sub

Private Sub mnuCtrl_Click(Index As Integer)
   Dim CPL As String
   On Error GoTo ProcedureError
   
   CPL = ControlArray(Index)
   Shell "rundll32.exe shell32.dll,Control_RunDLL " & CPL, vbNormalFocus
ProcedureExit:
  Exit Sub
ProcedureError:
  If ErrMsgBox(Me.Name & ".mnuCtrl_Click") = vbRetry Then Resume Next
End Sub

Private Sub mnuFile_Click(Index As Integer)
   If Index = 14 Then
      Unload Me
   End If
End Sub

Private Sub mnuLang_Click(Index As Integer)
   On Error GoTo ProcedureError
   Dim temp As Long, LangIndex As Long, L4 As Long

      'uncheck old Lang
      mnuLang(Lang \ c1000).Checked = False
      'check new Lang
      mnuLang(Index).Checked = True

      Lang = Index * c1000

      'update toolbar Image (Flag)
      LangIndex = cToolbar1.ButtonIndex("1116lang")
      cToolbar1.ButtonImage(LangIndex) = Lang \ c1000 + 14

      'update the Toolbar Caption/Tips
      For L4 = 0 To cToolbar1.ButtonCount - 1
         temp = Val(left$(cToolbar1.ButtonKey(L4), 4))
         If temp >= c1000 Then 'Must be 1000 or above
            cToolbar1.ButtonCaption(L4) = GetResourceString(temp)
            cToolbar1.ButtonToolTip(L4) = GetResourceString(temp + 30)
         End If
      Next

      UpdateLanguage ' fix remaining stuff

ProcedureExit:
  Exit Sub
ProcedureError:
  If ErrMsgBox(Me.Name & ".mnuLang_Click") = vbRetry Then Resume Next

End Sub
Public Sub ProgressPanel(iItem As Variant, iTot As Variant)
   On Error GoTo ProcedureError

   If (iTot = 0) Then
      Exit Sub
   End If

   AnalogMeter 2, 0, 100, 0, 30, 2, vbWhite, 0.1, (iItem / iTot) * 100


ProcedureExit:
  Exit Sub
ProcedureError:
  If ErrMsgBox(Me.Name & ".ProgressPanel") = vbRetry Then Resume Next

End Sub


Private Sub AnalogMeter(ByVal Mtype As Integer, ByVal Emin As Double, ByVal Emax As Double, _
ByVal Mmin As Double, ByVal Mmax As Double, ByVal Handw As Integer, ByVal Color As Long, ByVal Handl As Double, _
ByVal value As Double)
   On Error GoTo ProcedureError
Dim sl As Integer, mDeg As Integer
Dim r1 As Double
Dim r2 As Double
Dim cX As Integer, cY As Integer, r As Single
Dim s As Double
Dim sin_ As Double, cos_ As Double
'
' The meter movement is based off of a clock (0-60 minutes)
'
' The meter is assigned to a picture array (using Picture1).
' The Picture1 attributes are important.  Set Scalemode = Pixel (3)
'                                             Autosize = true
'                                             Autoredraw = true
'                                             Font Transparent = true
'                                             Border Style = your preference
'                                             Backcolor = your preference
'
' The Meter Face Image:
' I used MSpaint (the version that allows transparent .gif type format).
' The image is drawn so that the meter needle is centered, or centered at the
' bottom of the image.  Saved as a transparent .gif so that the background will
' take-on whatever Picture1.Backcolor you choose (see attribute above).
' When you load in the image into Picture1, the Picture1 Box will autosize to
' the .gif image (see attribute above).
'
' Meter Type 1 (center dial), places the zero at the bottom, 6 o'clock position.
' Meter Type 2 (half dial), places the zero at the left, 9 o'clock position.
'
' Because not all meters start the needle at the bottom, you need to specify the
' Meter Minimum (MMin) and Maximum (MMax) to tell it where the zero and range
' offsets are at.  Example is the Air Pressure Meter.
'
' The Engineering units are the zero and range of the value you wish to display.
' If a meter has no numbers on it, you can simply pick 0 to 100.
'
' Index = Which Meter to Adjust
' Mtype = Which Type of Meter it is (1=center dial, 2=half dial)
' EMin  = Engineering Units Zero (Minimum)
' EMax  = Engineering Units Span (Range)
' MMin  = Minimum point on the meter face (0-60)
' MMax  = Maximum point on the meter face (0-60)
' Handw = Thickness of dial hand (1,2,3)
' Color = Color of dial hand (vbRed, vbBlack)
' Handl = Length of the dial hand (0.1 is a good length)
' Value = The value to set the dial

' Determine whether dial hand is in center or edge.
' Also scale the dial hand to picture dimensions
If Mtype = 1 Then
   mDeg = 0 ' degrees (0-360)
   cX = (Preview.Width / 2) - 30
   cY = (Preview.Height / 2) - 30
ElseIf Mtype = 2 Then
   mDeg = 270 ' degrees (0-360)
   'cx = (Preview.Width / Screen.TwipsPerPixelX) / 2
   'cy = (Preview.Height / Screen.TwipsPerPixelY) - 2
   cX = (Preview.Width / 2) - 30
   cY = Preview.Height - 105
End If
 
' Scale the dial hand length
r = IIf(cX > cY, cY, cX) - 75
sl = r * Handl 'length of meter hand

' Scale the Engineering Units
r1 = Emax - Emin
r2 = Mmax - Mmin
s = ((r2 / r1) * value) + Mmin
  
' Draw the dial hand
  sin_ = Sin((mDeg - s * 6) * Deg2Rad) * (r - sl) + cX
  cos_ = Cos((mDeg - s * 6) * Deg2Rad) * (r - sl) + cY
  'oldcolor = vbBlack
  Preview.ForeColor = Color
  Preview.DrawWidth = Handw
  Preview.Cls
  Preview.Line (cX, cY)-(sin_, cos_)
  Preview.ForeColor = vbBlack

ProcedureExit:
  Exit Sub
ProcedureError:
  If ErrMsgBox(Me.Name & ".AnalogMeter") = vbRetry Then Resume Next

End Sub

Private Sub Timer1_Timer()
   Static Pos As Integer
   Pos = (Pos + 2) Mod 100
   ProgressPanel Pos, 100
End Sub
