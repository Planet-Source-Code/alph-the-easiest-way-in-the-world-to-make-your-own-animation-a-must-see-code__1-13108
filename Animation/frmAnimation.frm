VERSION 5.00
Begin VB.Form frmAnimation 
   BackColor       =   &H0080FF80&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "How to create you own animation!"
   ClientHeight    =   8175
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4335
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   545
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   289
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer timAnimate 
      Interval        =   50
      Left            =   2160
      Top             =   1560
   End
   Begin VB.PictureBox picLock 
      Height          =   630
      Left            =   1200
      ScaleHeight     =   38
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   38
      TabIndex        =   2
      Top             =   1800
      Width           =   630
      Begin VB.Image imgLong 
         Height          =   6840
         Left            =   0
         MousePointer    =   2  'Cross
         Picture         =   "frmAnimation.frx":0000
         Top             =   0
         Width           =   570
      End
   End
   Begin VB.Image Pointer 
      Height          =   570
      Left            =   3720
      MousePointer    =   10  'Up Arrow
      Picture         =   "frmAnimation.frx":CEE2
      Top             =   1200
      Width           =   570
   End
   Begin VB.Line Line7 
      X1              =   8
      X2              =   72
      Y1              =   136
      Y2              =   136
   End
   Begin VB.Shape Shape4 
      Height          =   375
      Left            =   1440
      Top             =   2520
      Width           =   135
   End
   Begin VB.Shape Shape3 
      Height          =   375
      Left            =   1440
      Top             =   1320
      Width           =   135
   End
   Begin VB.Line Line6 
      X1              =   136
      X2              =   192
      Y1              =   136
      Y2              =   136
   End
   Begin VB.Label lblDescribe 
      BackColor       =   &H00C0FFC0&
      Caption         =   $"frmAnimation.frx":D914
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3855
      Left            =   120
      TabIndex        =   4
      Top             =   4200
      Width           =   2655
   End
   Begin VB.Label lblSource 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Source Picture -------------------------------->"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   3840
      Width           =   2655
   End
   Begin VB.Image imgSource 
      Height          =   6840
      Left            =   3120
      MousePointer    =   7  'Size N S
      Picture         =   "frmAnimation.frx":DA61
      Top             =   1200
      Width           =   570
   End
   Begin VB.Line Line5 
      X1              =   176
      X2              =   176
      Y1              =   104
      Y2              =   176
   End
   Begin VB.Line Line4 
      X1              =   56
      X2              =   144
      Y1              =   184
      Y2              =   184
   End
   Begin VB.Line Line3 
      X1              =   56
      X2              =   144
      Y1              =   96
      Y2              =   96
   End
   Begin VB.Line Line2 
      X1              =   24
      X2              =   24
      Y1              =   104
      Y2              =   176
   End
   Begin VB.Label lblTrick 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "What a Trick!"
      BeginProperty Font 
         Name            =   "Rugged Stencil"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   315
      Left            =   0
      TabIndex        =   1
      Top             =   3120
      Width           =   3045
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   288
      Y1              =   72
      Y2              =   72
   End
   Begin VB.Shape Shape1 
      Height          =   855
      Left            =   120
      Top             =   120
      Width           =   4125
   End
   Begin VB.Label lblIntroduct 
      BackColor       =   &H00FF8080&
      Caption         =   $"frmAnimation.frx":1A943
      Height          =   855
      Left            =   135
      TabIndex        =   0
      Top             =   120
      Width           =   4095
   End
   Begin VB.Shape Shape2 
      Height          =   4935
      Left            =   -360
      Shape           =   4  'Rounded Rectangle
      Top             =   3600
      Width           =   3255
   End
End
Attribute VB_Name = "frmAnimation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Unload(Cancel As Integer)
    MsgBox "Please vote for my code on PSC. Thank you.", vbOKOnly, "Please vote!"
End Sub

Private Sub timAnimate_Timer()
    Static Frame As Integer
    imgLong.Top = 0 - (Frame * 38)
    Pointer.Top = 80 + (Frame * 38)
    
    Frame = Frame + 1
    If Frame = 12 Then Frame = 0
End Sub
