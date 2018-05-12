VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "           "
   ClientHeight    =   2970
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5055
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2970
   ScaleWidth      =   5055
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   3960
      TabIndex        =   6
      Top             =   2460
      Width           =   975
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderStyle     =   6  'Inside Solid
      Index           =   3
      X1              =   240
      X2              =   4740
      Y1              =   1860
      Y2              =   1860
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      Index           =   2
      X1              =   255
      X2              =   4740
      Y1              =   1875
      Y2              =   1875
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderStyle     =   6  'Inside Solid
      Index           =   1
      X1              =   240
      X2              =   4740
      Y1              =   900
      Y2              =   900
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      Index           =   0
      X1              =   255
      X2              =   4740
      Y1              =   915
      Y2              =   915
   End
   Begin VB.Label lblDisclaimer 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmAbout.frx":000C
      ForeColor       =   &H00000000&
      Height          =   885
      Left            =   300
      TabIndex        =   5
      Top             =   1020
      WhatsThisHelpID =   30
      Width           =   4410
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   180
      Picture         =   "frmAbout.frx":00E2
      Top             =   180
      WhatsThisHelpID =   20
      Width           =   480
   End
   Begin VB.Label lblAppTitle 
      AutoSize        =   -1  'True
      Caption         =   "lblAppTitle"
      Height          =   195
      Left            =   900
      TabIndex        =   4
      Top             =   180
      WhatsThisHelpID =   20
      Width           =   735
   End
   Begin VB.Label lblCopyright 
      AutoSize        =   -1  'True
      Caption         =   "lblCopyright"
      Height          =   195
      Left            =   900
      TabIndex        =   3
      Top             =   600
      WhatsThisHelpID =   20
      Width           =   810
   End
   Begin VB.Label lblAddress 
      Caption         =   "Priore Software"
      Height          =   255
      Index           =   0
      Left            =   360
      TabIndex        =   2
      Top             =   1980
      WhatsThisHelpID =   20
      Width           =   3795
   End
   Begin VB.Label lblAddress 
      Caption         =   "Internet: support@prioregroup.com"
      Height          =   255
      Index           =   1
      Left            =   360
      TabIndex        =   1
      Top             =   2280
      WhatsThisHelpID =   20
      Width           =   3795
   End
   Begin VB.Label lblAddress 
      Caption         =   "Web: http://www.prioregroup.com"
      Height          =   255
      Index           =   2
      Left            =   360
      TabIndex        =   0
      Top             =   2580
      WhatsThisHelpID =   20
      Width           =   3795
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdOK_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Const STR_OCX$ = ".OCX"
    Dim OCX As String
    
    OCX = UCase$(App.EXEName)
    If Right$(OCX, 4) <> STR_OCX Then OCX = OCX & STR_OCX
    Me.Caption = "About " & App.Title
    lblAppTitle.Caption = App.Title & " - " & OCX
    lblCopyright.Caption = "Copyright © 2001/" & Format$(Now, "yyyy") & ", Danilo Priore"
End Sub

