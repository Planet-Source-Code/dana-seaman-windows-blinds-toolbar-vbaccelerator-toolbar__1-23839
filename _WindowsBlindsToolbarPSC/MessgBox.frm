VERSION 5.00
Begin VB.Form frmMessgBox 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   Caption         =   "Title"
   ClientHeight    =   2220
   ClientLeft      =   105
   ClientTop       =   -240
   ClientWidth     =   4650
   ControlBox      =   0   'False
   Icon            =   "MessgBox.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "MessgBox.frx":000C
   ScaleHeight     =   148
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   310
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin WBToolbarDemo.ButtonEx Button 
      Height          =   300
      Index           =   0
      Left            =   300
      TabIndex        =   1
      Top             =   1620
      Width           =   1080
      _ExtentX        =   1905
      _ExtentY        =   529
      Appearance      =   2
      Caption         =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TransparentColor=   16711935
      SkinOver        =   "MessgBox.frx":7BEE
      SkinUp          =   "MessgBox.frx":85E0
      TransparentColor=   16711935
   End
   Begin VB.Timer Timer1 
      Left            =   3960
      Top             =   240
   End
   Begin WBToolbarDemo.ButtonEx Button 
      Height          =   300
      Index           =   1
      Left            =   1770
      TabIndex        =   2
      Top             =   1620
      Visible         =   0   'False
      Width           =   1080
      _ExtentX        =   1905
      _ExtentY        =   529
      Appearance      =   2
      Caption         =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TransparentColor=   16711935
      SkinOver        =   "MessgBox.frx":8FD2
      SkinUp          =   "MessgBox.frx":99C4
      TransparentColor=   16711935
   End
   Begin WBToolbarDemo.ButtonEx Button 
      Height          =   300
      Index           =   2
      Left            =   3210
      TabIndex        =   3
      Top             =   1620
      Visible         =   0   'False
      Width           =   1080
      _ExtentX        =   1905
      _ExtentY        =   529
      Appearance      =   2
      Caption         =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TransparentColor=   16711935
      SkinOver        =   "MessgBox.frx":A3B6
      SkinUp          =   "MessgBox.frx":ADA8
      TransparentColor=   16711935
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Title"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   2100
      TabIndex        =   4
      Top             =   120
      Width           =   495
   End
   Begin VB.Shape Shape1 
      BorderWidth     =   4
      Height          =   2160
      Left            =   30
      Shape           =   4  'Rounded Rectangle
      Top             =   30
      Width           =   4590
   End
   Begin VB.Label lblMsg 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Message box text"
      ForeColor       =   &H80000008&
      Height          =   975
      Left            =   900
      TabIndex        =   0
      Top             =   540
      Width           =   3555
   End
End
Attribute VB_Name = "frmMessgBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private StandardIcon As Long

Private Sub Button_Click(Index As Integer)

   PlaySound 61
   Message = Index 'Stores value of pressed button
   Unload Me
End Sub

Private Sub Button_MouseEnter(Index As Integer)
   PlaySound 60
End Sub

Private Sub Form_Load()

   MakeFormRounded Me, 6

   Select Case SoundVal 'Plays sound, sets StandardIcon ID
      Case 0
         StandardIcon = 0
      Case 1
         MessageBeep MB_IconAsterisk
         StandardIcon = IDI_Error
      Case 2
         MessageBeep MB_IconQuestion
         StandardIcon = IDI_Question
      Case 3
         MessageBeep MB_IconExclamation
         StandardIcon = IDI_Warning
      Case 4
         MessageBeep MB_IconInformation
         StandardIcon = IDI_Information
      Case Else
   End Select
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   FormDrag Me
End Sub

Private Sub Form_Paint()
    Dim hIcon As Long
    hIcon = LoadStandardIcon(0&, StandardIcon)
    Call DrawIcon(Me.hDC, 30&, 40&, hIcon)
End Sub

Private Sub Timer1_Timer()
   SendKeys "~"  'a.k.a "{Enter}"
End Sub
