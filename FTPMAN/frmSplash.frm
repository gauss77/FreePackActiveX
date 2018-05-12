VERSION 5.00
Begin VB.Form frmSplash 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   4530
   ClientLeft      =   2805
   ClientTop       =   3540
   ClientWidth     =   6060
   ClipControls    =   0   'False
   Icon            =   "frmSplash.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4530
   ScaleWidth      =   6060
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tmrUnload 
      Interval        =   10000
      Left            =   5040
      Top             =   3480
   End
   Begin VB.Label lblAddress 
      Caption         =   "Web: http://priore.w3.to"
      Height          =   255
      Index           =   2
      Left            =   900
      TabIndex        =   7
      Top             =   3960
      Width           =   3795
   End
   Begin VB.Label lblAddress 
      Caption         =   "Internet: priore@w3.to"
      Height          =   255
      Index           =   1
      Left            =   900
      TabIndex        =   6
      Top             =   3660
      Width           =   3795
   End
   Begin VB.Label lblAddress 
      Caption         =   "Priore Software"
      Height          =   255
      Index           =   0
      Left            =   900
      TabIndex        =   5
      Top             =   3360
      Width           =   3795
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      Index           =   1
      X1              =   900
      X2              =   5880
      Y1              =   3180
      Y2              =   3180
   End
   Begin VB.Label lblInfo 
      Caption         =   $"frmSplash.frx":000C
      Height          =   435
      Index           =   2
      Left            =   900
      TabIndex        =   4
      Top             =   2520
      Width           =   4935
   End
   Begin VB.Label lblInfo 
      Caption         =   "Complete pricing, shopping and licensing information can be found in this control's help file."
      Height          =   435
      Index           =   1
      Left            =   900
      TabIndex        =   3
      Top             =   1920
      Width           =   4935
   End
   Begin VB.Label lblInfo 
      Caption         =   $"frmSplash.frx":00A2
      Height          =   495
      Index           =   0
      Left            =   900
      TabIndex        =   2
      Top             =   1260
      Width           =   4935
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      Index           =   0
      X1              =   900
      X2              =   5880
      Y1              =   1020
      Y2              =   1020
   End
   Begin VB.Label lblCopyright 
      AutoSize        =   -1  'True
      Caption         =   "lblCopyright"
      Height          =   195
      Left            =   900
      TabIndex        =   1
      Top             =   660
      Width           =   810
   End
   Begin VB.Label lblAppTitle 
      AutoSize        =   -1  'True
      Caption         =   "lblAppTitle"
      Height          =   195
      Left            =   900
      TabIndex        =   0
      Top             =   240
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   180
      Picture         =   "frmSplash.frx":012D
      Top             =   240
      Width           =   480
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    Const STR_OCX$ = ".OCX"
    Dim OCX As String
    
    OCX = UCase$(App.EXEName)
    If Right$(OCX, 4) <> STR_OCX Then OCX = OCX & STR_OCX
    Me.Caption = "About " & App.Title
    lblAppTitle.Caption = App.Title & " - " & OCX
    lblCopyright.Caption = "Copyright © 1999/" & Format$(Now, "yyyy") & ", Danilo Priore"
End Sub

Private Sub tmrUnload_Timer()
    Unload Me
End Sub
