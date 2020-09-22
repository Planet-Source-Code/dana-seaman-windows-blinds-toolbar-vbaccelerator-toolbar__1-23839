VERSION 5.00
Begin VB.Form frmTbStrip 
   BackColor       =   &H00FF00FF&
   Caption         =   "WB Toolbar Strip Builder"
   ClientHeight    =   2955
   ClientLeft      =   4590
   ClientTop       =   1650
   ClientWidth     =   10635
   LinkTopic       =   "Form1"
   ScaleHeight     =   2955
   ScaleWidth      =   10635
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      Height          =   1095
      Left            =   60
      MultiLine       =   -1  'True
      TabIndex        =   26
      Text            =   "TbStrip.frx":0000
      Top             =   1740
      Width           =   8055
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00FF00FF&
      BorderStyle     =   0  'None
      Height          =   540
      Left            =   60
      ScaleHeight     =   540
      ScaleWidth      =   6480
      TabIndex        =   13
      Top             =   960
      Width           =   6480
      Begin ToolbarStripBuilder.ButtonEx btnEx 
         Height          =   540
         Index           =   12
         Left            =   0
         TabIndex        =   14
         Top             =   0
         Width           =   540
         _ExtentX        =   953
         _ExtentY        =   953
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
         Picture         =   "TbStrip.frx":01B8
         PictureOffsetX  =   1
         PictureOffsetY  =   1
         TransparentColor=   16711935
         SkinOver        =   "TbStrip.frx":05FA
         SkinUp          =   "TbStrip.frx":0F5C
         TransparentColor=   16711935
      End
      Begin ToolbarStripBuilder.ButtonEx btnEx 
         Height          =   540
         Index           =   13
         Left            =   540
         TabIndex        =   15
         Top             =   0
         Width           =   540
         _ExtentX        =   953
         _ExtentY        =   953
         Appearance      =   2
         BackColor       =   14737632
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
         Picture         =   "TbStrip.frx":1EDE
         PictureOffsetX  =   1
         PictureOffsetY  =   1
         TransparentColor=   16711935
         TransparentColor=   16711935
      End
      Begin ToolbarStripBuilder.ButtonEx btnEx 
         Height          =   540
         Index           =   14
         Left            =   1080
         TabIndex        =   16
         Top             =   0
         Width           =   540
         _ExtentX        =   953
         _ExtentY        =   953
         Appearance      =   2
         BackColor       =   16777215
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
         Picture         =   "TbStrip.frx":2326
         PictureOffsetX  =   4
         TransparentColor=   16711935
         TransparentColor=   16711935
      End
      Begin ToolbarStripBuilder.ButtonEx btnEx 
         Height          =   540
         Index           =   15
         Left            =   1620
         TabIndex        =   17
         Top             =   0
         Width           =   540
         _ExtentX        =   953
         _ExtentY        =   953
         Appearance      =   2
         BackColor       =   14737632
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
         Picture         =   "TbStrip.frx":2776
         PictureOffsetX  =   4
         PictureOffsetY  =   -1
         TransparentColor=   16711935
         TransparentColor=   16711935
      End
      Begin ToolbarStripBuilder.ButtonEx btnEx 
         Height          =   540
         Index           =   16
         Left            =   2160
         TabIndex        =   18
         Top             =   0
         Width           =   540
         _ExtentX        =   953
         _ExtentY        =   953
         Appearance      =   2
         BackColor       =   16777215
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
         Picture         =   "TbStrip.frx":2FC8
         PictureOffsetX  =   2
         TransparentColor=   16711935
         TransparentColor=   16711935
      End
      Begin ToolbarStripBuilder.ButtonEx btnEx 
         Height          =   540
         Index           =   17
         Left            =   2700
         TabIndex        =   19
         Top             =   0
         Width           =   540
         _ExtentX        =   953
         _ExtentY        =   953
         Appearance      =   2
         BackColor       =   14737632
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
         Picture         =   "TbStrip.frx":3457
         PictureOffsetX  =   4
         TransparentColor=   16711935
         TransparentColor=   16711935
      End
      Begin ToolbarStripBuilder.ButtonEx btnEx 
         Height          =   540
         Index           =   18
         Left            =   3240
         TabIndex        =   20
         Top             =   0
         Width           =   540
         _ExtentX        =   953
         _ExtentY        =   953
         Appearance      =   2
         BackColor       =   16777215
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
         Picture         =   "TbStrip.frx":3FA9
         PictureOffsetX  =   4
         TransparentColor=   16711935
         TransparentColor=   16711935
      End
      Begin ToolbarStripBuilder.ButtonEx btnEx 
         Height          =   540
         Index           =   19
         Left            =   3780
         TabIndex        =   21
         Top             =   0
         Width           =   540
         _ExtentX        =   953
         _ExtentY        =   953
         Appearance      =   2
         BackColor       =   14737632
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
         Picture         =   "TbStrip.frx":4BFB
         PictureOffsetX  =   2
         TransparentColor=   16711935
         TransparentColor=   16711935
      End
      Begin ToolbarStripBuilder.ButtonEx btnEx 
         Height          =   540
         Index           =   20
         Left            =   4320
         TabIndex        =   22
         Top             =   0
         Width           =   540
         _ExtentX        =   953
         _ExtentY        =   953
         Appearance      =   2
         BackColor       =   16777215
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
         Picture         =   "TbStrip.frx":578D
         PictureOffsetX  =   4
         TransparentColor=   16711935
         TransparentColor=   16711935
      End
      Begin ToolbarStripBuilder.ButtonEx btnEx 
         Height          =   540
         Index           =   21
         Left            =   4860
         TabIndex        =   23
         Top             =   0
         Width           =   540
         _ExtentX        =   953
         _ExtentY        =   953
         Appearance      =   2
         BackColor       =   14737632
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
         Picture         =   "TbStrip.frx":5BDE
         PictureOffsetX  =   3
         TransparentColor=   16711935
         TransparentColor=   16711935
      End
      Begin ToolbarStripBuilder.ButtonEx btnEx 
         Height          =   540
         Index           =   22
         Left            =   5400
         TabIndex        =   24
         Top             =   0
         Width           =   540
         _ExtentX        =   953
         _ExtentY        =   953
         Appearance      =   2
         BackColor       =   16777215
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
         Picture         =   "TbStrip.frx":6830
         PictureOffsetX  =   4
         TransparentColor=   16711935
         TransparentColor=   16711935
      End
      Begin ToolbarStripBuilder.ButtonEx btnEx 
         Height          =   540
         Index           =   23
         Left            =   5940
         TabIndex        =   25
         Top             =   0
         Width           =   540
         _ExtentX        =   953
         _ExtentY        =   953
         Appearance      =   2
         BackColor       =   14737632
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
         Picture         =   "TbStrip.frx":7482
         PictureOffsetX  =   5
         TransparentColor=   16711935
         TransparentColor=   16711935
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FF00FF&
      BorderStyle     =   0  'None
      Height          =   720
      Left            =   60
      ScaleHeight     =   720
      ScaleWidth      =   8640
      TabIndex        =   0
      Top             =   60
      Width           =   8640
      Begin ToolbarStripBuilder.ButtonEx btnEx 
         Height          =   720
         Index           =   0
         Left            =   0
         TabIndex        =   1
         Top             =   0
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
         Picture         =   "TbStrip.frx":7CF4
         PictureOffsetX  =   8
         TransparentColor=   16711935
         SkinOver        =   "TbStrip.frx":8136
         SkinUp          =   "TbStrip.frx":884C
         TransparentColor=   16711935
      End
      Begin ToolbarStripBuilder.ButtonEx btnEx 
         Height          =   720
         Index           =   1
         Left            =   720
         TabIndex        =   2
         Top             =   0
         Width           =   720
         _ExtentX        =   1270
         _ExtentY        =   1270
         Appearance      =   2
         BackColor       =   14737632
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
         Picture         =   "TbStrip.frx":8EF7
         PictureOffsetX  =   8
         TransparentColor=   16711935
         TransparentColor=   16711935
      End
      Begin ToolbarStripBuilder.ButtonEx btnEx 
         Height          =   720
         Index           =   2
         Left            =   1440
         TabIndex        =   3
         Top             =   0
         Width           =   720
         _ExtentX        =   1270
         _ExtentY        =   1270
         Appearance      =   2
         BackColor       =   16777215
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
         Picture         =   "TbStrip.frx":933F
         PictureOffsetX  =   10
         TransparentColor=   16711935
         TransparentColor=   16711935
      End
      Begin ToolbarStripBuilder.ButtonEx btnEx 
         Height          =   720
         Index           =   3
         Left            =   2160
         TabIndex        =   4
         Top             =   0
         Width           =   720
         _ExtentX        =   1270
         _ExtentY        =   1270
         Appearance      =   2
         BackColor       =   14737632
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
         Picture         =   "TbStrip.frx":978F
         PictureOffsetX  =   8
         TransparentColor=   16711935
         TransparentColor=   16711935
      End
      Begin ToolbarStripBuilder.ButtonEx btnEx 
         Height          =   720
         Index           =   4
         Left            =   2880
         TabIndex        =   5
         Top             =   0
         Width           =   720
         _ExtentX        =   1270
         _ExtentY        =   1270
         Appearance      =   2
         BackColor       =   16777215
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
         Picture         =   "TbStrip.frx":9FE1
         PictureOffsetX  =   8
         TransparentColor=   16711935
         TransparentColor=   16711935
      End
      Begin ToolbarStripBuilder.ButtonEx btnEx 
         Height          =   720
         Index           =   5
         Left            =   3600
         TabIndex        =   6
         Top             =   0
         Width           =   720
         _ExtentX        =   1270
         _ExtentY        =   1270
         Appearance      =   2
         BackColor       =   14737632
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
         Picture         =   "TbStrip.frx":A470
         PictureOffsetX  =   8
         TransparentColor=   16711935
         TransparentColor=   16711935
      End
      Begin ToolbarStripBuilder.ButtonEx btnEx 
         Height          =   720
         Index           =   6
         Left            =   4320
         TabIndex        =   7
         Top             =   0
         Width           =   720
         _ExtentX        =   1270
         _ExtentY        =   1270
         Appearance      =   2
         BackColor       =   16777215
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
         Picture         =   "TbStrip.frx":AFC2
         PictureOffsetX  =   8
         TransparentColor=   16711935
         TransparentColor=   16711935
      End
      Begin ToolbarStripBuilder.ButtonEx btnEx 
         Height          =   720
         Index           =   7
         Left            =   5040
         TabIndex        =   8
         Top             =   0
         Width           =   720
         _ExtentX        =   1270
         _ExtentY        =   1270
         Appearance      =   2
         BackColor       =   14737632
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
         Picture         =   "TbStrip.frx":BC14
         PictureOffsetX  =   8
         TransparentColor=   16711935
         TransparentColor=   16711935
      End
      Begin ToolbarStripBuilder.ButtonEx btnEx 
         Height          =   720
         Index           =   8
         Left            =   5760
         TabIndex        =   9
         Top             =   0
         Width           =   720
         _ExtentX        =   1270
         _ExtentY        =   1270
         Appearance      =   2
         BackColor       =   16777215
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
         Picture         =   "TbStrip.frx":C7A6
         PictureOffsetX  =   8
         TransparentColor=   16711935
         TransparentColor=   16711935
      End
      Begin ToolbarStripBuilder.ButtonEx btnEx 
         Height          =   720
         Index           =   9
         Left            =   6480
         TabIndex        =   10
         Top             =   0
         Width           =   720
         _ExtentX        =   1270
         _ExtentY        =   1270
         Appearance      =   2
         BackColor       =   14737632
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
         Picture         =   "TbStrip.frx":CBF7
         PictureOffsetX  =   8
         TransparentColor=   16711935
         TransparentColor=   16711935
      End
      Begin ToolbarStripBuilder.ButtonEx btnEx 
         Height          =   720
         Index           =   10
         Left            =   7200
         TabIndex        =   11
         Top             =   0
         Width           =   720
         _ExtentX        =   1270
         _ExtentY        =   1270
         Appearance      =   2
         BackColor       =   16777215
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
         Picture         =   "TbStrip.frx":D849
         PictureOffsetX  =   10
         TransparentColor=   16711935
         TransparentColor=   16711935
      End
      Begin ToolbarStripBuilder.ButtonEx btnEx 
         Height          =   720
         Index           =   11
         Left            =   7920
         TabIndex        =   12
         Top             =   0
         Width           =   720
         _ExtentX        =   1270
         _ExtentY        =   1270
         Appearance      =   2
         BackColor       =   14737632
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
         Picture         =   "TbStrip.frx":E49B
         PictureOffsetX  =   10
         TransparentColor=   16711935
         TransparentColor=   16711935
      End
   End
End
Attribute VB_Name = "frmTbStrip"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()

   Dim i As Integer
   
   For i = 1 To 11
      Set btnEx(i).SkinUp = btnEx(0).SkinUp
      Set btnEx(i).SkinOver = btnEx(0).SkinOver
      Set btnEx(i + 12).SkinUp = btnEx(12).SkinUp
      Set btnEx(i + 12).SkinOver = btnEx(12).SkinOver
   Next
End Sub


