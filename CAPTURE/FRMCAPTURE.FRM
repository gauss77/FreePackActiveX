VERSION 5.00
Object = "*\ACaptureCtl.vbp"
Begin VB.Form frmCapture 
   Caption         =   "Capture"
   ClientHeight    =   4245
   ClientLeft      =   1755
   ClientTop       =   1590
   ClientWidth     =   5385
   Icon            =   "frmCapture.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4245
   ScaleWidth      =   5385
   Begin VB.CommandButton cmd 
      Caption         =   "&Screen"
      Height          =   375
      Index           =   3
      Left            =   4020
      TabIndex        =   4
      Top             =   3780
      Width           =   1215
   End
   Begin VB.CommandButton cmd 
      Caption         =   "&Window"
      Height          =   375
      Index           =   2
      Left            =   2700
      TabIndex        =   3
      Top             =   3780
      Width           =   1215
   End
   Begin VB.CommandButton cmd 
      Caption         =   "&Client"
      Height          =   375
      Index           =   1
      Left            =   1380
      TabIndex        =   2
      Top             =   3780
      Width           =   1215
   End
   Begin VB.CommandButton cmd 
      Caption         =   "&Active"
      Height          =   375
      Index           =   0
      Left            =   60
      TabIndex        =   1
      Top             =   3780
      Width           =   1215
   End
   Begin CaptureCtl.Capture Capture1 
      Height          =   3615
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   4815
      _ExtentX        =   8493
      _ExtentY        =   6376
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FontName        =   "MS Sans Serif"
      FontSize        =   8,25
      ScaleHeight     =   3615
      ScaleMode       =   0
      ScaleWidth      =   4815
      AutoSize        =   -1  'True
   End
End
Attribute VB_Name = "frmCapture"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Capture1_Resize()
    Me.Width = Capture1.Width + 60
    Me.Height = Capture1.Height + 1035
End Sub

Private Sub cmd_Click(Index As Integer)
    Capture1.CaptureType = Index
    Capture1.Capture
End Sub

Private Sub Form_Resize()
    Dim i As Integer
    
    For i = 0 To 3
        cmd(i).Top = Me.Height - 870
    Next
End Sub
