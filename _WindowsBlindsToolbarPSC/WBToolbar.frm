VERSION 5.00
Begin VB.Form frmWBToolbar 
   BackColor       =   &H0084846B&
   Caption         =   "Lightweight Windows Blinds Toolbar Clone / Deluxe MessageBox"
   ClientHeight    =   6570
   ClientLeft      =   4590
   ClientTop       =   1650
   ClientWidth     =   10320
   LinkTopic       =   "Form1"
   ScaleHeight     =   6570
   ScaleWidth      =   10320
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   5175
      Left            =   60
      Picture         =   "WBToolbar.frx":0000
      ScaleHeight     =   5175
      ScaleWidth      =   2550
      TabIndex        =   45
      Top             =   1320
      Width           =   2550
      Begin WBToolbarDemo.ButtonEx ButtonEx1 
         Height          =   480
         Index           =   0
         Left            =   555
         TabIndex        =   46
         Top             =   1200
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   847
         Appearance      =   2
         Caption         =   "Show Msg"
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
         TransparentColor=   16711935
      End
      Begin WBToolbarDemo.ButtonEx ButtonEx1 
         Height          =   480
         Index           =   1
         Left            =   555
         TabIndex        =   47
         Top             =   1800
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   847
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
         TransparentColor=   16711935
      End
      Begin WBToolbarDemo.ButtonEx ButtonEx1 
         Height          =   480
         Index           =   2
         Left            =   555
         TabIndex        =   48
         Top             =   2400
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   847
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
         TransparentColor=   16711935
      End
      Begin WBToolbarDemo.ButtonEx ButtonEx1 
         Height          =   480
         Index           =   3
         Left            =   555
         TabIndex        =   49
         Top             =   3000
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   847
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
         TransparentColor=   16711935
      End
      Begin WBToolbarDemo.ButtonEx ButtonEx1 
         Height          =   480
         Index           =   4
         Left            =   555
         TabIndex        =   50
         Top             =   3600
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   847
         Appearance      =   2
         Caption         =   "EXIT"
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
         TransparentColor=   16711935
      End
   End
   Begin VB.TextBox Text1 
      Height          =   2175
      Left            =   6240
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   44
      Top             =   1440
      Width           =   3795
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H0084846B&
      Caption         =   "Buttons"
      Height          =   1695
      Left            =   2820
      TabIndex        =   37
      Top             =   1440
      Width           =   2355
      Begin VB.TextBox txtCaption 
         Height          =   285
         Index           =   0
         Left            =   780
         TabIndex        =   40
         Text            =   "OK"
         Top             =   300
         Width           =   1290
      End
      Begin VB.TextBox txtCaption 
         Height          =   285
         Index           =   1
         Left            =   780
         TabIndex        =   39
         Text            =   "Retry"
         Top             =   720
         Width           =   1290
      End
      Begin VB.TextBox txtCaption 
         Height          =   285
         Index           =   2
         Left            =   780
         TabIndex        =   38
         Text            =   "Cancel"
         Top             =   1140
         Width           =   1290
      End
      Begin VB.Label lblButton 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "One"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   0
         Left            =   300
         TabIndex        =   43
         Tag             =   "1304"
         Top             =   360
         Width           =   315
      End
      Begin VB.Label lblButton 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Two"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   1
         Left            =   300
         TabIndex        =   42
         Tag             =   "1304"
         Top             =   780
         Width           =   315
      End
      Begin VB.Label lblButton 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Three"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   2
         Left            =   180
         TabIndex        =   41
         Tag             =   "1304"
         Top             =   1200
         Width           =   420
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H0084846B&
      Caption         =   "Timer Enabled"
      Height          =   1815
      Left            =   4560
      TabIndex        =   31
      Top             =   3420
      Width           =   1395
      Begin VB.OptionButton optTimer 
         BackColor       =   &H0084846B&
         Caption         =   "Off"
         Height          =   255
         Index           =   0
         Left            =   180
         TabIndex        =   35
         Top             =   300
         Width           =   1095
      End
      Begin VB.OptionButton optTimer 
         BackColor       =   &H0084846B&
         Caption         =   "On"
         Height          =   255
         Index           =   1
         Left            =   180
         TabIndex        =   34
         Top             =   660
         Value           =   -1  'True
         Width           =   1095
      End
      Begin VB.TextBox txtSeconds 
         Height          =   285
         Left            =   180
         Locked          =   -1  'True
         TabIndex        =   33
         Text            =   "5"
         Top             =   1080
         Width           =   450
      End
      Begin VB.VScrollBar VScroll 
         Height          =   285
         Left            =   600
         Max             =   59
         TabIndex        =   32
         Top             =   1080
         Width           =   180
      End
      Begin VB.Label lblMinutes 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Seconds"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   150
         TabIndex        =   36
         Tag             =   "1304"
         Top             =   1380
         Width           =   645
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H0084846B&
      Caption         =   "Icon/Sound"
      Height          =   2415
      Left            =   2820
      TabIndex        =   25
      Top             =   3420
      Width           =   1515
      Begin VB.OptionButton optIcon 
         BackColor       =   &H0084846B&
         Caption         =   "None"
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   30
         Top             =   360
         Width           =   915
      End
      Begin VB.OptionButton optIcon 
         BackColor       =   &H0084846B&
         Caption         =   "Error"
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   29
         Top             =   780
         Width           =   915
      End
      Begin VB.OptionButton optIcon 
         BackColor       =   &H0084846B&
         Caption         =   "Question"
         Height          =   195
         Index           =   2
         Left            =   120
         TabIndex        =   28
         Top             =   1200
         Width           =   1215
      End
      Begin VB.OptionButton optIcon 
         BackColor       =   &H0084846B&
         Caption         =   "Warning"
         Height          =   195
         Index           =   3
         Left            =   120
         TabIndex        =   27
         Top             =   1620
         Width           =   915
      End
      Begin VB.OptionButton optIcon 
         BackColor       =   &H0084846B&
         Caption         =   "Information"
         Height          =   195
         Index           =   4
         Left            =   120
         TabIndex        =   26
         Top             =   1980
         Value           =   -1  'True
         Width           =   1275
      End
   End
   Begin VB.PictureBox picTB1 
      BackColor       =   &H00FF00FF&
      Height          =   1215
      Left            =   0
      Picture         =   "WBToolbar.frx":6938
      ScaleHeight     =   1155
      ScaleWidth      =   10275
      TabIndex        =   0
      Top             =   0
      Width           =   10335
      Begin WBToolbarDemo.ButtonEx btnEx 
         Height          =   720
         Index           =   0
         Left            =   120
         TabIndex        =   1
         Top             =   120
         Width           =   720
         _ExtentX        =   1270
         _ExtentY        =   1270
         Appearance      =   2
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "WBToolbar.frx":B495
         PictureOffsetX  =   8
         TransparentColor=   16711935
         SkinOver        =   "WBToolbar.frx":B8D7
         SkinUp          =   "WBToolbar.frx":C229
         TransparentColor=   16711935
      End
      Begin WBToolbarDemo.ButtonEx btnEx 
         Height          =   720
         Index           =   1
         Left            =   960
         TabIndex        =   3
         Top             =   120
         Width           =   720
         _ExtentX        =   1270
         _ExtentY        =   1270
         Appearance      =   2
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "WBToolbar.frx":CBB7
         PictureOffsetX  =   8
         TransparentColor=   16711935
         TransparentColor=   16711935
      End
      Begin WBToolbarDemo.ButtonEx btnEx 
         Height          =   720
         Index           =   2
         Left            =   1800
         TabIndex        =   5
         Top             =   120
         Width           =   720
         _ExtentX        =   1270
         _ExtentY        =   1270
         Appearance      =   2
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "WBToolbar.frx":CFFF
         PictureOffsetX  =   10
         TransparentColor=   16711935
         TransparentColor=   16711935
      End
      Begin WBToolbarDemo.ButtonEx btnEx 
         Height          =   720
         Index           =   3
         Left            =   2640
         TabIndex        =   6
         Top             =   120
         Width           =   720
         _ExtentX        =   1270
         _ExtentY        =   1270
         Appearance      =   2
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "WBToolbar.frx":D44F
         PictureOffsetX  =   8
         TransparentColor=   16711935
         SkinOver        =   "WBToolbar.frx":DCA1
         SkinUp          =   "WBToolbar.frx":E6B9
         TransparentColor=   16711935
      End
      Begin WBToolbarDemo.ButtonEx btnEx 
         Height          =   720
         Index           =   4
         Left            =   3480
         TabIndex        =   7
         Top             =   120
         Width           =   720
         _ExtentX        =   1270
         _ExtentY        =   1270
         Appearance      =   2
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "WBToolbar.frx":F0D6
         PictureOffsetX  =   8
         TransparentColor=   16711935
         TransparentColor=   16711935
      End
      Begin WBToolbarDemo.ButtonEx btnEx 
         Height          =   720
         Index           =   5
         Left            =   4320
         TabIndex        =   8
         Top             =   120
         Width           =   720
         _ExtentX        =   1270
         _ExtentY        =   1270
         Appearance      =   2
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "WBToolbar.frx":F565
         PictureOffsetX  =   8
         TransparentColor=   16711935
         TransparentColor=   16711935
      End
      Begin WBToolbarDemo.ButtonEx btnEx 
         Height          =   720
         Index           =   6
         Left            =   5160
         TabIndex        =   9
         Top             =   120
         Width           =   720
         _ExtentX        =   1270
         _ExtentY        =   1270
         Appearance      =   2
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "WBToolbar.frx":100B7
         PictureOffsetX  =   8
         TransparentColor=   16711935
         SkinOver        =   "WBToolbar.frx":10D09
         SkinUp          =   "WBToolbar.frx":112EE
         TransparentColor=   16711935
      End
      Begin WBToolbarDemo.ButtonEx btnEx 
         Height          =   720
         Index           =   7
         Left            =   6000
         TabIndex        =   10
         Top             =   120
         Width           =   720
         _ExtentX        =   1270
         _ExtentY        =   1270
         Appearance      =   2
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "WBToolbar.frx":118D7
         PictureOffsetX  =   8
         TransparentColor=   16711935
         TransparentColor=   16711935
      End
      Begin WBToolbarDemo.ButtonEx btnEx 
         Height          =   720
         Index           =   8
         Left            =   6840
         TabIndex        =   11
         Top             =   120
         Width           =   720
         _ExtentX        =   1270
         _ExtentY        =   1270
         Appearance      =   2
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "WBToolbar.frx":12469
         PictureOffsetX  =   8
         TransparentColor=   16711935
         TransparentColor=   16711935
      End
      Begin WBToolbarDemo.ButtonEx btnEx 
         Height          =   720
         Index           =   9
         Left            =   7680
         TabIndex        =   12
         Top             =   120
         Width           =   720
         _ExtentX        =   1270
         _ExtentY        =   1270
         Appearance      =   2
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "WBToolbar.frx":128BA
         PictureOffsetX  =   8
         TransparentColor=   16711935
         SkinOver        =   "WBToolbar.frx":1350C
         SkinUp          =   "WBToolbar.frx":13C20
         TransparentColor=   16711935
      End
      Begin WBToolbarDemo.ButtonEx btnEx 
         Height          =   720
         Index           =   10
         Left            =   8520
         TabIndex        =   13
         Top             =   120
         Width           =   720
         _ExtentX        =   1270
         _ExtentY        =   1270
         Appearance      =   2
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "WBToolbar.frx":14335
         PictureOffsetX  =   10
         TransparentColor=   16711935
         TransparentColor=   16711935
      End
      Begin WBToolbarDemo.ButtonEx btnEx 
         Height          =   720
         Index           =   11
         Left            =   9360
         TabIndex        =   14
         Top             =   120
         Width           =   720
         _ExtentX        =   1270
         _ExtentY        =   1270
         Appearance      =   2
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "WBToolbar.frx":14F87
         PictureOffsetX  =   10
         TransparentColor=   16711935
         TransparentColor=   16711935
      End
      Begin VB.Shape shpRR 
         BorderColor     =   &H00000000&
         BorderWidth     =   2
         Height          =   1115
         Left            =   -2000
         Shape           =   4  'Rounded Rectangle
         Top             =   30
         Width           =   860
      End
      Begin VB.Label lblTB 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Vault"
         Height          =   195
         Index           =   11
         Left            =   9540
         TabIndex        =   24
         Top             =   900
         Width           =   375
      End
      Begin VB.Label lblTB 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Properties"
         Height          =   195
         Index           =   10
         Left            =   8520
         TabIndex        =   23
         Top             =   900
         Width           =   735
      End
      Begin VB.Label lblTB 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Select"
         Height          =   195
         Index           =   9
         Left            =   7815
         TabIndex        =   22
         Top             =   900
         Width           =   465
      End
      Begin VB.Label lblTB 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Sort"
         Height          =   195
         Index           =   8
         Left            =   7050
         TabIndex        =   21
         Top             =   900
         Width           =   315
      End
      Begin VB.Label lblTB 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Colors"
         Height          =   195
         Index           =   7
         Left            =   6135
         TabIndex        =   20
         Top             =   900
         Width           =   465
      End
      Begin VB.Label lblTB 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Zip"
         Height          =   195
         Index           =   6
         Left            =   5400
         TabIndex        =   19
         Top             =   900
         Width           =   255
      End
      Begin VB.Label lblTB 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Rename"
         Height          =   195
         Index           =   5
         Left            =   4380
         TabIndex        =   18
         Top             =   900
         Width           =   615
      End
      Begin VB.Label lblTB 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Shred"
         Height          =   195
         Index           =   4
         Left            =   3660
         TabIndex        =   17
         Top             =   900
         Width           =   435
      End
      Begin VB.Label lblTB 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Delete"
         Height          =   195
         Index           =   3
         Left            =   2745
         TabIndex        =   16
         Top             =   900
         Width           =   495
      End
      Begin VB.Label lblTB 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Recycle"
         Height          =   195
         Index           =   2
         Left            =   1860
         TabIndex        =   15
         Top             =   900
         Width           =   615
      End
      Begin VB.Label lblTB 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Move"
         Height          =   195
         Index           =   1
         Left            =   1095
         TabIndex        =   4
         Top             =   900
         Width           =   435
      End
      Begin VB.Label lblTB 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Copy"
         Height          =   195
         Index           =   0
         Left            =   300
         TabIndex        =   2
         Top             =   900
         Width           =   375
      End
   End
End
Attribute VB_Name = "frmWBToolbar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub btnEx_Click(Index As Integer)
   PlaySound 61
   MesgBox "You pressed button " & Index & vbCrLf & _
           Chr$(34) & lblTB(Index) & Chr$(34) & vbCr & _
           "AutoUnload activated (6.5 seconds)", 4, "Windows Blinds MsgBox", _
           "OK", "Retry", "Cancel", _
           6500
End Sub

Private Sub btnEx_MouseEnter(Index As Integer)
   PlaySound 60
   shpRR.Left = btnEx(Index).Left - 60
   lblTB(Index).FontBold = True
End Sub

Private Sub btnEx_MouseExit(Index As Integer)
   lblTB(Index).FontBold = False
End Sub

Private Sub btnEx_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   PlaySound 61
End Sub

Private Sub ButtonEx1_Click(Index As Integer)
   Dim i As Integer
   Dim TimerInterval As Long
   
   PlaySound 61
   
   Select Case Index
      Case 0
         For i = 0 To optIcon().UBound
            If optIcon(i).Value = True Then
               SoundVal = i
               Exit For
            End If
         Next
         If optTimer(1) Then
            TimerInterval = Val(txtSeconds) * 1000
         End If
        
         MesgBox "This is a Message Box" + vbCrLf + _
                  "with three lines" + vbCr + _
                  "of text", SoundVal, "Windows Blinds MsgBox", _
                  txtCaption(0), txtCaption(1), txtCaption(2), _
                  TimerInterval
         'Unrem next line to see return value
         'MsgBox "Button '" & Message & "' pressed"
      Case 4
         Unload Me
   End Select
End Sub

Private Sub ButtonEx1_MouseEnter(Index As Integer)
   PlaySound 60
End Sub

Private Sub Form_Load()

   Dim i As Integer, j As Integer
   InSound = True 'enable sound
   '--------
   For i = 0 To 4
      Set ButtonEx1(i).SkinUp = LoadPicture(QualifyPath(App.Path) & "\bmpgifjpg\nv1002.gif")
      Set ButtonEx1(i).SkinOver = LoadPicture(QualifyPath(App.Path) & "\bmpgifjpg\nv2002.gif")
   Next
   '--------
   For i = 0 To 3
      For j = 1 To 2
         Set btnEx(i * 3 + j).SkinUp = btnEx(i * 3).SkinUp
         Set btnEx(i * 3 + j).SkinOver = btnEx(i * 3).SkinOver
      Next
   Next
   
   txtSeconds = 5
   Text1.Text = "This Lightweight Toolbar can be used standalone where you don't need the capabilities of a full-featured toolbar. Messagebox features Windows Blinds style buttons, Icons/Sounds from system w/o overhead, optional timer for auto-close."
 
End Sub

Private Sub VScroll_Change()
   txtSeconds = VScroll.Max - VScroll.Value
End Sub
