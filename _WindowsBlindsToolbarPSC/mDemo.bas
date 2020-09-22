Attribute VB_Name = "mDemo"
Option Explicit



'--------Constants--------
'Private Const TV_FIRST = &H1100
'Public Const TVM_SETBKCOLOR As Long = (TV_FIRST + 29)
Public Const HKEY_CLASSES_ROOT = &H80000000
'Public Const KEY_ALL_ACCESS = &H2003F
Private Const FO_MOVE = &H1
Private Const FO_COPY = &H2
Private Const FO_DELETE = &H3
Private Const FO_RENAME = &H4
Private Const FOF_NOCONFIRMATION = &H10
Private Const FOF_ALLOWUNDO = &H40
Private Const SEE_MASK_INVOKEIDLIST = &HC
Private Const SEE_MASK_NOCLOSEPROCESS = &H40
Private Const SEE_MASK_FLAG_NO_UI = &H400
'----------------------------------------------
Public Const rDayZeroBias As Double = 109205# 'Abs(CDbl(#01-01-1601#))
Public Const rMillisecondPerDay As Double = 10000000# * 60# * 60# * 24# / 10000#
'Public Const DATE_LONGDATE = &H2
'Public Const DATE_SHORTDATE = &H1
Public Const LOCALE_SSHORTDATE = &H1F
'Public Const LOCALE_SLONGDATE = &H20
Public Const LOCALE_STIMEFORMAT = &H1003
Public Const FILE_ATTRIBUTE_DIRECTORY = &H10
'Public Const WM_TIMECHANGE = &H1E
Public Const BIF_RETURNONLYFSDIRS = &H1
Public Const INVALID_HANDLE_VALUE = -1
Public Const LocalFileHeaderSig = &H4034B50
Public Const CentralFileHeaderSig = &H2014B50
Public Const EndCentralDirSig = &H6054B50
Public Const MAX_PATH = 260
'Public Const SHGFI_DISPLAYNAME = &H200
'Public Const SHGFI_TYPENAME = &H400
' Public Const MAXDWORD = (2 ^ 32) - 1   ' 0xFFFFFFFF
' W32FD.nFileSizeHigh = high-order DWORD
' W32FD.nFileSizeLow  = low-order DWORD
' BigSize = (W32FD.nFileSizeHigh * MAXDWORD) + W32FD.nFileSizeLow
Public Const KEY_ALL_ACCESS = &H2003F
Public Const VbZlStr = ""
Public Const Deg2Rad = 3.141592654 / 180 'Degrees to Radians
Public Const c1000 = 1000
Public Const sFill As String = "..............."
Public Const sAttr As String = "rhsvdalnt?lco?e"
'Complete list
'00 r   0001 0001 "Read Only"
'01 h   0002 0002 "Hidden"
'02 s   0004 0004 "System"
'03 v   0008 0008 "Volume"
'04 d   0016 0010 "Directory", "Folder"
'05 a   0032 0020 "Archive"
'06 l   0064 0040 "Alias", ".LNK"
'07 n   0128 0080 "Normal"
'08 t   0256 0100 "Temporary"
'09 ?   0512 0200  ??
'10 l   1024 0400 "Alias"
'11 c   2048 0800 "Compressed"
'12 o   4096 1000  "Offline"
'13 ?   8192 2000  ??
'14 e  16384 4000 "Encrypted
Public Const CourierWhite As String = "<font face='courier new'   color='white' size='2'>"
Public Const SMx As String = "1AEE666C5797CE4536917EE245445F057B811C3DC71667395479C7A862FEA95870833EE50012EC3DD3108FE9AFFD3AD795A4C0BD9F3C118912169ADF68E104A92884A9CED123FA249C4DFD90097CFDF995D5CB4F32896C699F248B47004DEAFD4836FB4A2E0E4C6F966C1B2CA915CEBFBE483E06F2429E3F0DB5BF467F757E6733CD599D1AE5261233C02DEFC11DDF4E34435F36E4AAE92FA32E691F09ADE65C2D7C160BEB2FBC90A99F580EA5B786DA6C0972337B05927DB2A3EEE6F20A94F0"
'----------------------------------------------
Public Buffer       As String * MAX_PATH
Public f_Type       As String * 80
Public HowManyTags  As Integer
Public WhichDates   As Integer
Public Lang         As Long
Public Ret          As Long
'Public FilterAttr   As Long
Public SourcePath   As String
Public DestinPath   As String
Public RegDateStr   As String
Public RegTimeStr   As String
Public OldRegDTStr  As String
Public SoundOn      As Boolean
Public InDoy        As Boolean
Public InZip        As Boolean
Public NewDateTime  As Date
Public OsVersion    As OSVERSIONINFO
Public Tile         As New cTile
'----------------------------------------------
Type WIN32_FIND_DATA
   dwFileAttributes  As Long
   ftCreationTime    As Currency   'As FILETIME
   ftLastAccessTime  As Currency   'As FILETIME
   ftLastWriteTime   As Currency   'As FILETIME
   nFileSizeHigh     As String * 4 'Long
   nFileSizeLow      As String * 4 'Long
   dwReserved0       As Long
   dwReserved1       As Long
   cFileName         As String * MAX_PATH
   cAlternate        As String * 14
End Type
Type BROWSEINFO
  hOwner          As Long
  pidlRoot        As Long
  pszDisplayName  As String
  lpszTitle       As String
  ulFlags         As Long
  lpfn            As Long
  lParam          As Long
  iImage          As Long
End Type
Type OSVERSIONINFO
    dwOSVersionInfoSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatformId As Long
    szCSDVersion As String * 128
End Type
'Type SHFILEINFO
'   hicon          As Long
'   iIcon          As Long
'   dwAttributes   As Long
'   szDisplayName  As String * MAX_PATH
'   szTypeName     As String * 80
'End Type
Public Type FTs
   Ext As String
   Type As String
End Type
Public Type PicBmp
   Size As Long
   tType As Long
   hBmp As Long
   hPal As Long
   Reserved As Long
End Type
Public Type Guid
   Data1 As Long
   Data2 As Integer
   Data3 As Integer
   Data4(7) As Byte
End Type
Public Type ZipFile
  Version            As Integer
  Flag               As Integer
  CompressionMethod  As Integer
  Time               As Integer
  Date               As Integer
  CRC32              As Long
  CompressedSize     As Long
  UncompressedSize   As Long
  FileNameLength     As Integer
  ExtraFieldLength   As Integer
  Filename           As String
  ExtraField         As String
End Type
Public Enum CompareMethod
    BinaryCompare
    TextCompare
End Enum
Private Enum OSVersionEnum
    VER_PLATFORM_WIN32s = 0
    VER_PLATFORM_WIN32_WINDOWS = 1
    VER_PLATFORM_WIN32_NT = 2
End Enum
Public Enum ESetWindowPosStyles
    SWP_SHOWWINDOW = &H40
    SWP_HIDEWINDOW = &H80
    SWP_FRAMECHANGED = &H20 ' The frame changed: send WM_NCCALCSIZE
    SWP_NOACTIVATE = &H10
    SWP_NOCOPYBITS = &H100
    SWP_NOMOVE = &H2
    SWP_NOOWNERZORDER = &H200 ' Don't do owner Z ordering
    SWP_NOREDRAW = &H8
    SWP_NOREPOSITION = SWP_NOOWNERZORDER
    SWP_NOSIZE = &H1
    SWP_NOZORDER = &H4
    SWP_DRAWFRAME = SWP_FRAMECHANGED
    HWND_NOTOPMOST = -2
    HWND_TOPMOST = -1
    TOPMOST_FLAGS = SWP_NOMOVE Or SWP_NOSIZE
End Enum
Public Type ITEMIDLIST
    mkid As Long
End Type
Public Enum SHFolders
    CSIDL_DESKTOP = &H0
    CSIDL_INTERNET = &H1
    CSIDL_PROGRAMS = &H2
    CSIDL_CONTROLS = &H3
    CSIDL_PRINTERS = &H4
    CSIDL_PERSONAL = &H5
    CSIDL_FAVORITES = &H6
    CSIDL_STARTUP = &H7
    CSIDL_RECENT = &H8
    CSIDL_SENDTO = &H9
    CSIDL_BITBUCKET = &HA
    CSIDL_STARTMENU = &HB
    CSIDL_DESKTOPDIRECTORY = &H10
    CSIDL_DRIVES = &H11
    CSIDL_NETWORK = &H12
    CSIDL_NETHOOD = &H13
    CSIDL_FONTS = &H14
    CSIDL_TEMPLATES = &H15
    CSIDL_COMMON_STARTMENU = &H16
    CSIDL_COMMON_PROGRAMS = &H17
    CSIDL_COMMON_STARTUP = &H18
    CSIDL_COMMON_DESKTOPDIRECTORY = &H19
    CSIDL_APPDATA = &H1A
    CSIDL_PRINTHOOD = &H1B
    CSIDL_ALTSTARTUP = &H1D '// DBCS
    CSIDL_COMMON_ALTSTARTUP = &H1E '// DBCS
    CSIDL_COMMON_FAVORITES = &H1F
    CSIDL_INTERNET_CACHE = &H20
    CSIDL_COOKIES = &H21
    CSIDL_HISTORY = &H22
End Enum
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
Private Type SHELLEXECUTEINFO
   cbSize         As Long
   fMask          As Long
   hwnd           As Long
   lpVerb         As String
   lpFile         As String
   lpParameters   As String
   lpDirectory    As String
   nShow          As Long
   hInstApp       As Long
   lpIDList       As Long      ' Optional parameter
   lpClass        As String    ' Optional parameter
   hkeyClass      As Long      ' Optional parameter
   dwHotKey       As Long      ' Optional parameter
   hIcon          As Long      ' Optional parameter
   hProcess       As Long      ' Optional parameter
End Type
Private Type SHFILEOPSTRUCT
     hwnd As Long
     wFunc As Long
     pFrom As String
     pTo As String
     fFlags As Integer
     fAnyOperationsAborted As Boolean
     hNameMappings As Long
     lpszProgressTitle As String
End Type


'----------------------------------------------
'IMPORTANT NOTE *****
'some declares changed to "As Any" or "Currency"
'in lieu of Type FILETIME
'
'Declare Function GetDesktopWindow Lib "user32" () As Long
'Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
'Declare Function GetLocaleInfo Lib "kernel32" Alias "GetLocaleInfoA" (ByVal Locale As Long, ByVal LCType As Long, ByVal lpLCData As String, ByVal cchData As Long) As Long
Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Long
Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
'Declare Function GetDeviceCaps Lib "gdi32" (ByVal hdc As Long, ByVal nIndex As Long) As Long
'Declare Function CreateDC Lib "gdi32" Alias "CreateDCA" (ByVal lpDriverName As String, ByVal lpDeviceName As String, ByVal lpOutput As String, lpInitData As DEVMODE) As Long
'Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
'Declare Function SetSystemTime Lib "kernel32" (lpSystemTime As SYSTEMTIME) As Long
Declare Function sndPlaySound Lib "winmm" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
Declare Function FindClose Lib "kernel32" (ByVal hFindFile As Long) As Long
Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileA" (ByVal lpFileName As String, lpFindFileData As WIN32_FIND_DATA) As Long
Declare Function FindNextFile Lib "kernel32" Alias "FindNextFileA" (ByVal hFindFile As Long, lpFindFileData As WIN32_FIND_DATA) As Long
Declare Function DosDateTimeToFileTime Lib "kernel32" (ByVal wFatDate As Long, ByVal wFatTime As Long, lpFileTime As Currency) As Long
Declare Function FileTimeToLocalFileTime Lib "kernel32" (lpFileTime As Any, lpLocalFileTime As Any) As Long
'Declare Function GetLogicalDrives Lib "kernel32" () As Long
Declare Function ReleaseCapture Lib "user32" () As Long
Declare Function CloseHandle Lib "kernel32" (ByVal hFile As Long) As Long
Declare Function GetSystemDefaultLCID Lib "kernel32" () As Long
Declare Function CreateFile Lib "kernel32" Alias "CreateFileA" (ByVal lpFileName As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, lpSecurityAttributes As Any, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
Declare Function GetTickCount Lib "kernel32" () As Long
Declare Function GetFileTime Lib "kernel32" (ByVal hFile As Long, lpCreationTime As Currency, lpLastAccessTime As Currency, lpLastWriteTime As Currency) As Long
Declare Function SetFileTime Lib "kernel32" (ByVal hFile As Long, lpCreationTime As Currency, lpLastAccessTime As Currency, lpLastWriteTime As Currency) As Long
Declare Function GetLocaleInfo Lib "kernel32" Alias "GetLocaleInfoA" (ByVal Locale As Long, ByVal LCType As Long, ByVal lpLCData As String, ByVal cchData As Long) As Long
'Declare Function GetDateFormat Lib "kernel32" Alias "GetDateFormatA" (ByVal Locale As Long, ByVal dwFlags As Long, lpDate As SYSTEMTIME, ByVal lpFormat As String, ByVal lpDateStr As String, ByVal cchDate As Long) As Long
'Declare Function GetTimeFormat Lib "kernel32" Alias "GetTimeFormatA" (ByVal Locale As Long, ByVal dwFlags As Long, lpTime As SYSTEMTIME, ByVal lpFormat As String, ByVal lpTimeStr As String, ByVal cchTime As Long) As Long
'Declare Sub GetLocalTime Lib "kernel32" (lpSystemTime As SYSTEMTIME)
Declare Function LocalFileTimeToFileTime Lib "kernel32" (lpLocalFileTime As Any, lpFileTime As Any) As Long
Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long
Declare Function SHBrowseForFolder Lib "shell32.dll" Alias "SHBrowseForFolderA" (lpBrowseInfo As BROWSEINFO) As Long
Declare Function SHGetPathFromIDList Lib "shell32.dll" Alias "SHGetPathFromIDListA" (ByVal pidl As Long, ByVal pszPath As String) As Long
Declare Function SHFileOperation Lib "shell32.dll" Alias "SHFileOperationA" (lpFileOp As SHFILEOPSTRUCT) As Long
Declare Function ShellExecuteEX Lib "shell32.dll" Alias "ShellExecuteEx" (SEI As SHELLEXECUTEINFO) As Long
Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
'Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cX As Long, ByVal cY As Long, ByVal wFlags As Long) As Long
Private Declare Function SetWindowRgn Lib "user32" (ByVal hwnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
Private Declare Function CreateRoundRectRgn Lib "gdi32" (ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long
'Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Long) As Long
Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)
Declare Function LoadString Lib "user32" Alias "LoadStringA" (ByVal hInstance As Long, ByVal uID As Long, ByVal lpBuffer As String, ByVal nBufferMax As Long) As Long
Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Declare Function GetSystemDirectory Lib "kernel32" Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
'Declare Function RegOpenKeyEx Lib "advapi32" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, ByRef phkResult As Long) As Long
Declare Function RegQueryValueEx Lib "advapi32" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, ByRef lpType As Long, ByVal lpData As String, ByRef lpcbData As Long) As Long
Declare Function RegCloseKey Lib "advapi32" (ByVal hKey As Long) As Long
'Declare Function SHGetFileInfo Lib "shell32.dll" Alias "SHGetFileInfoA" _
'           (ByVal pszPath As String, _
'            ByVal dwFileAttributes As Long, _
'            psfi As SHFILEINFO, _
'            ByVal cbSizeFileInfo As Long, _
'            ByVal uFlags As Long) As Long
Declare Function SHGetSpecialFolderLocation Lib "shell32.dll" (ByVal hwndOwner As Long, ByVal nFolder As SHFolders, ppidl As ITEMIDLIST) As Long
'Declare Function SHGetPathFromIDList Lib "shell32.dll" Alias "SHGetPathFromIDListA" (ByVal pidl As Long, ByVal pszPath As String) As Long
Declare Sub CoTaskMemFree Lib "ole32.dll" (ByVal pv As Long)
Public Declare Function OleCreatePictureIndirect Lib "OLEPRO32.DLL" _
  (PicDesc As PicBmp, _
   RefIID As Guid, _
   ByVal fPictureOwnsHandle As Long, _
   ipic As IPicture) As Long
Public Declare Function ExtractIconEx Lib "shell32.dll" Alias "ExtractIconExA" _
   (ByVal lpszFile As String, _
   ByVal nIconIndex As Long, _
   phiconLarge As Long, _
   phiconSmall As Long, _
   ByVal nIcons As Long) As Long
Public Declare Function DestroyIcon Lib "user32" _
   (ByVal hIcon As Long) As Long
Public Function GenDate(MyDate As Date, Optional JustDate As Boolean = False) As String
   
   If InDoy Then
      GenDate = Format$(MyDate, "yyyy/y")
      GenDate = left$(GenDate, 5) & right$("00" & Mid$(GenDate, 6), 3)
      If JustDate = False Then
         GenDate = GenDate & " " & FormatDateTime(MyDate, vbLongTime)
      End If
   Else
      If JustDate Then
         GenDate = FormatDateTime(MyDate, vbShortDate)
      Else
         GenDate = FormatDateTime(MyDate, vbGeneralDate)
      End If
   End If

End Function

Public Sub HoverSound()
    Const SYNC = 1
    If SoundOn Then
      sndPlaySound ByVal App.Path & "\linkhover.wav", SYNC
    End If
End Sub
Public Function RegDateTimeStr(LOCALE_SSHORTDATE_SLONGDATE) As String
   On Error GoTo ProcedureError
   Dim sLen As Long
   Dim sDate As String * 32
'------------------------------
   sLen = GetLocaleInfo(GetSystemDefaultLCID(), LOCALE_SSHORTDATE_SLONGDATE, sDate, 64)
   RegDateStr = SetLen(sDate, sLen)
'------------------------------
   sLen = GetLocaleInfo(GetSystemDefaultLCID(), LOCALE_STIMEFORMAT, sDate, 64)
   RegTimeStr = SetLen(sDate, sLen)
'------------------------------
   RegDateTimeStr = RegDateStr & RegTimeStr

ProcedureExit:
  Exit Function
ProcedureError:
  If ErrMsgBox("mDeclare.RegDateTimeStr") = vbRetry Then Resume Next

End Function
Public Function SetLen(s As String, lLen As Long) As String
   If lLen > 1 Then
      SetLen = left$(s, lLen - 1)
   End If
End Function
Public Function IsWinNt() As Boolean
'===============================================================================
'   IsWinNT - Returns true if we're running Windows NT.
'===============================================================================
    
    If OsVersion.dwOSVersionInfoSize = 0 Then               ' this is our first time making this call
        OsVersion.dwOSVersionInfoSize = Len(OsVersion)      ' initialize so API knows which version being used
        GetVersionEx OsVersion                              ' make the call once & then save/re-use it
    End If

    IsWinNt = OsVersion.dwPlatformId = VER_PLATFORM_WIN32_NT ' return the result
    
End Function
Public Function FileExistsW32FD(sSource As String, W32Fd As WIN32_FIND_DATA) As Boolean

   Dim hFile As Long
   'Returns True if file exists as well as raw data
   'in WIN32_FIND_DATA structure
   hFile = FindFirstFile(sSource, W32Fd)
   FileExistsW32FD = hFile <> INVALID_HANDLE_VALUE
   FindClose hFile

End Function
Public Sub SkinButtons(F1 As Form, ButtCount As Integer)
   On Error GoTo ProcedureError
   Dim L4 As Long
   For L4 = 0 To ButtCount
'      Set F1.Button(L4).SkinUp = frmSplash.Button.SkinUp
'      Set F1.Button(L4).SkinDown = frmSplash.Button.SkinDown
'      Set F1.Button(L4).SkinOver = frmSplash.Button.SkinOver
   Next
ProcedureExit:
  Exit Sub
ProcedureError:
  If ErrMsgBox("Public Sub SkinButtons") = vbRetry Then Resume Next

End Sub
Public Sub SetControlCaptionStrings(frm As Form)

   On Error GoTo ProcedureError

   Dim ctl  As Control

   '-- set the form's caption
   If frm.Tag <> "" Then
      frm.Caption = GetResourceString(CInt(frm.Tag))
   End If
   '-- set the font
   '-- Set fnt = frm.Font
   '-- fnt.Name = GetResourceString(20)
   '-- fnt.Size = CInt(GetResourceString(21))

   '-- set the controls' captions using the Tag property
   For Each ctl In frm.Controls
      If ctl.Tag <> "" Then
         Select Case TypeName(ctl)
            Case "Menu", "Label", "CheckBox", "OptionButton", "ButtonEx"
               ctl.Caption = GetResourceString(Int(ctl.Tag))
         End Select
      End If
   Next

ProcedureExit:
  Exit Sub
ProcedureError:
  If ErrMsgBox("mDeclare.SetControlCaptionStrings") = vbRetry Then Resume Next

End Sub

Public Function GetResourceStringFromFile(sModule As String, idString As Variant) As String

   Dim hModule As Long
   Dim nChars As Long

   hModule = LoadLibrary(sModule)
   If hModule Then
      nChars = LoadString(hModule, idString, Buffer, MAX_PATH)
      If nChars Then
         GetResourceStringFromFile = left$(Buffer, nChars)
      End If
      FreeLibrary hModule
   End If
End Function
Public Function GetResourceString(Num As Variant) As String
   On Error Resume Next
   Select Case Num
      Case 1000 To 1999 'Get from resource file (.Res)
         GetResourceString = LoadResString(Lang + Num)
      Case Else
         GetResourceString = GetResourceStringFromFile("Shell32.Dll", Num)
   End Select
End Function



Public Sub FormDrag(TheForm As Form)
   ReleaseCapture
   SendMessage TheForm.hwnd, &HA1, 2, 0&
End Sub

Public Sub MakeFormRounded(Obj As Object, Radius As Long)
    Dim hRgn As Long
    hRgn = CreateRoundRectRgn(0&, 0&, Obj.ScaleWidth, Obj.ScaleHeight, Obj.ScaleWidth / Radius, Obj.ScaleHeight / Radius)
    SetWindowRgn Obj.hwnd, hRgn, True
    Call DeleteObject(hRgn)
End Sub

Public Function ErrMsgBox(Msg As String) As Integer
   ErrMsgBox = MsgBox("Error: " & Err.Number & ". " & Err.Description, vbRetryCancel + vbCritical, Msg)
End Function


