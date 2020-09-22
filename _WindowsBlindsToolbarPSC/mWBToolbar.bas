Attribute VB_Name = "mWBToolbar"
Option Explicit

Public InSound As Boolean
Public SoundVal As Integer
Public Message As Integer

Private Declare Function sndPlaySound Lib "winmm" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
Private Declare Function CreateRoundRectRgn Lib "gdi32" (ByVal x1 As Long, ByVal y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function SetWindowRgn Lib "user32" (ByVal hWnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
Public Declare Function MessageBeep Lib "user32" (ByVal wType As Long) As Long
Public Declare Function LoadStandardIcon Lib "user32" Alias _
    "LoadIconA" (ByVal hInstance As Long, ByVal lpIconNum As _
    SystemIconConstants) As Long
Public Declare Function DrawIcon Lib "user32" (ByVal hDC _
    As Long, ByVal X As Long, ByVal Y As Long, _
    ByVal hIcon As Long) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Sub ReleaseCapture Lib "user32" ()
    
Enum SystemIconConstants
    IDI_Application = 32512
    IDI_Error = 32513       'vbCritical (a.k.a. IDI_Hand)
    IDI_Question = 32514    'vbQuestion
    IDI_Warning = 32515     'vbExlamation (a.k.a. IDI_Exclamation)
    IDI_Information = 32516 'vbInformation (a.k.a. IDI_Asterisk)
    IDI_WinLogo = 32517
End Enum

Public Const MB_IconAsterisk = &H10&
Public Const MB_IconQuestion = &H20&
Public Const MB_IconExclamation = &H30&
Public Const MB_IconInformation = &H40&

Public Sub FormDrag(TheForm As Form)
   ReleaseCapture
   SendMessage TheForm.hWnd, &HA1, 2, 0&
End Sub
Public Function QualifyPath(Path) As String
   If Right(Path, 1) = "\" Then
      QualifyPath = Path
   Else
      QualifyPath = Path & "\"
   End If
End Function

Public Sub PlaySound(Flavor As Integer)
   Const SYNC = 1
   Dim temp As String
   If InSound Then
      If Flavor = 60 Then
         temp = "Hover.wav"
      ElseIf Flavor = 61 Then
         temp = "Clique.wav"
      End If
      sndPlaySound ByVal App.Path & "\" & temp, SYNC
   End If
End Sub

Public Sub MesgBox(mbText As String, mbSound As Integer, _
                        mbTitle As String, mbCBut0 As String, _
                        Optional mbCBut1 As String, _
                        Optional mbCBut2 As String, _
                        Optional tInterval As Long)
                        
   Dim iScaleX As Integer
   
   SoundVal = mbSound  'Stores public value to play sound
   If frmMessgBox.ScaleMode = 3 Then 'Twip mode
      iScaleX = Screen.TwipsPerPixelX
   Else 'Pixel mode
      iScaleX = 1
   End If
   With frmMessgBox
      If mbCBut2 <> "" Then
         .Button(2).Left = 320 \ iScaleX
         .Button(1).Left = 1765 \ iScaleX
         .Button(0).Left = 3210 \ iScaleX
      ElseIf mbCBut1 <> "" Then
         .Button(1).Left = 800 \ iScaleX
         .Button(0).Left = 2725 \ iScaleX
      Else
         .Button(0).Left = 1762 \ iScaleX
      End If
      'Message
      .lblMsg.Caption = mbText
      'Message box title
      If mbTitle <> "" Then
         .lblTitle.Caption = mbTitle
      Else
         .lblTitle.Caption = App.Title
      End If
      'Set button captions
      .Button(0).Caption = mbCBut0
      .Button(1).Caption = mbCBut1
      .Button(2).Caption = mbCBut2
      If mbCBut1 <> "" Then
         .Button(1).Visible = True
         If mbCBut2 <> "" Then
            .Button(2).Visible = True
         End If
      End If
      .Timer1.Interval = tInterval
   End With
   
   frmMessgBox.Show 1

End Sub

Public Sub MakeFormRounded(Obj As Object, Radius As Long)
    Dim hRgn As Long
    hRgn = CreateRoundRectRgn(0, 0, Obj.ScaleWidth, Obj.ScaleHeight, Obj.ScaleWidth / Radius, Obj.ScaleHeight / Radius)
    SetWindowRgn Obj.hWnd, hRgn, True
    Call DeleteObject(hRgn)
End Sub
