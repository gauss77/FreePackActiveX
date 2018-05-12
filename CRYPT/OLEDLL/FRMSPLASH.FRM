VERSION 5.00
Begin VB.Form frmSplash 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   4530
   ClientLeft      =   2805
   ClientTop       =   3540
   ClientWidth     =   5265
   ClipControls    =   0   'False
   HelpContextID   =   20
   Icon            =   "FRMSPLASH.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4530
   ScaleWidth      =   5265
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdOk 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   3900
      TabIndex        =   8
      Top             =   3840
      Width           =   1215
   End
   Begin VB.Label lblAddress 
      Caption         =   "Web: http://www.prioregroup.com"
      Height          =   255
      Index           =   2
      Left            =   180
      TabIndex        =   7
      Top             =   3900
      WhatsThisHelpID =   20
      Width           =   3795
   End
   Begin VB.Label lblAddress 
      Caption         =   "Internet: support@prioregroup.com"
      Height          =   255
      Index           =   1
      Left            =   180
      TabIndex        =   6
      Top             =   3600
      WhatsThisHelpID =   20
      Width           =   3795
   End
   Begin VB.Label lblAddress 
      Caption         =   "Priore Software"
      Height          =   255
      Index           =   0
      Left            =   180
      TabIndex        =   5
      Top             =   3300
      WhatsThisHelpID =   20
      Width           =   3795
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      Index           =   1
      X1              =   180
      X2              =   5160
      Y1              =   3120
      Y2              =   3120
   End
   Begin VB.Label lblInfo 
      Caption         =   $"FRMSPLASH.frx":000C
      Height          =   435
      Index           =   2
      Left            =   180
      TabIndex        =   4
      Top             =   2460
      WhatsThisHelpID =   20
      Width           =   4935
   End
   Begin VB.Label lblInfo 
      Caption         =   "Complete pricing, shopping and licensing information can be found in this control's help file."
      Height          =   435
      Index           =   1
      Left            =   180
      TabIndex        =   3
      Top             =   1860
      WhatsThisHelpID =   20
      Width           =   4935
   End
   Begin VB.Label lblInfo 
      Caption         =   $"FRMSPLASH.frx":00A2
      Height          =   615
      Index           =   0
      Left            =   180
      TabIndex        =   2
      Top             =   1140
      WhatsThisHelpID =   20
      Width           =   4935
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      Index           =   0
      X1              =   180
      X2              =   5160
      Y1              =   960
      Y2              =   960
   End
   Begin VB.Label lblCopyright 
      AutoSize        =   -1  'True
      Caption         =   "lblCopyright"
      Height          =   195
      Left            =   180
      TabIndex        =   1
      Top             =   600
      WhatsThisHelpID =   20
      Width           =   810
   End
   Begin VB.Label lblAppTitle 
      AutoSize        =   -1  'True
      Caption         =   "lblAppTitle"
      Height          =   195
      Left            =   180
      TabIndex        =   0
      Top             =   180
      WhatsThisHelpID =   20
      Width           =   735
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdOk_Click()
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

