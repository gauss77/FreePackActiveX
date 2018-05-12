VERSION 5.00
Object = "{B10809B6-F27D-11D2-9939-000000000000}#1.0#0"; "MOUSE.OCX"
Begin VB.Form frmMouse 
   Caption         =   "Mouse Info"
   ClientHeight    =   1605
   ClientLeft      =   4935
   ClientTop       =   3135
   ClientWidth     =   2280
   Icon            =   "frmMouse.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   1605
   ScaleWidth      =   2280
   Begin VB.CommandButton cmdNoLimit 
      Caption         =   "&No limit"
      Height          =   375
      Left            =   60
      TabIndex        =   2
      Top             =   1140
      Width           =   2115
   End
   Begin VB.CommandButton cmdLimit 
      Caption         =   "&Limit in this form"
      Height          =   375
      Left            =   60
      TabIndex        =   1
      Top             =   660
      Width           =   2115
   End
   Begin VB.Timer Timer1 
      Left            =   1740
      Top             =   120
   End
   Begin MouseInfoCtl.MouseInfo MouseInfo1 
      Left            =   900
      Top             =   60
      _ExtentX        =   847
      _ExtentY        =   847
      X               =   404
      Y               =   216
      DblClickTime    =   500
   End
   Begin VB.Label lblMouse 
      BackColor       =   &H80000018&
      BorderStyle     =   1  'Fixed Single
      Height          =   495
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   2115
   End
End
Attribute VB_Name = "frmMouse"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdLimit_Click()
    Set MouseInfo1.Control = frmMouse
    MouseInfo1.MouseLimit = True
End Sub

Private Sub cmdNoLimit_Click()
    MouseInfo1.MouseLimit = False
End Sub

Private Sub Form_Load()
    Timer1.Interval = 100
End Sub

Private Sub Timer1_Timer()
    lblMouse.Caption = "Mouse X Pos :" & MouseInfo1.x & _
    vbCrLf & "Mouse Y Pos :" & MouseInfo1.y
End Sub
