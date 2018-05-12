VERSION 5.00
Begin VB.Form frmSplash 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   3870
   ClientLeft      =   2910
   ClientTop       =   5460
   ClientWidth     =   6060
   ClipControls    =   0   'False
   Icon            =   "frmSplash.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3870
   ScaleWidth      =   6060
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tmrUnload 
      Interval        =   10000
      Left            =   5040
      Top             =   3180
   End
   Begin VB.Label lblWeb 
      AutoSize        =   -1  'True
      Caption         =   "http://priore.w3.to"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   195
      Index           =   1
      Left            =   1380
      TabIndex        =   8
      Top             =   3540
      Width           =   1290
   End
   Begin VB.Label lblWeb 
      AutoSize        =   -1  'True
      Caption         =   "URL:"
      Height          =   195
      Index           =   0
      Left            =   900
      TabIndex        =   7
      Top             =   3540
      Width           =   375
   End
   Begin VB.Label lblAddress 
      AutoSize        =   -1  'True
      Caption         =   "E-mail: priore@w3.to"
      Height          =   195
      Index           =   1
      Left            =   900
      TabIndex        =   6
      Top             =   3300
      Width           =   1455
   End
   Begin VB.Label lblAddress 
      AutoSize        =   -1  'True
      Caption         =   "Priore Software"
      Height          =   195
      Index           =   0
      Left            =   900
      TabIndex        =   5
      Top             =   3060
      Width           =   1080
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      Index           =   1
      X1              =   900
      X2              =   5880
      Y1              =   2880
      Y2              =   2880
   End
   Begin VB.Label lblInfo 
      Caption         =   "For a complete list of all of our controls, packages and other products, plase see the Internet WEB site of Priore."
      Height          =   435
      Index           =   2
      Left            =   900
      TabIndex        =   4
      Top             =   2220
      Width           =   4935
   End
   Begin VB.Label lblInfo 
      Caption         =   "Complete pricing, shopping and licensing information of other controls can be found in the Internet WEB site of Priore."
      Height          =   435
      Index           =   1
      Left            =   900
      TabIndex        =   3
      Top             =   1680
      Width           =   4935
   End
   Begin VB.Label lblInfo 
      Caption         =   "This is an freeware version of this control. it is not crippled in any way."
      Height          =   315
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
      Picture         =   "frmSplash.frx":000C
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

Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Private Sub Form_Load()
    Const STR_OCX$ = ".OCX"
    Dim OCX As String
    
    OCX = UCase$(App.EXEName)
    If Right$(OCX, 4) <> STR_OCX Then OCX = OCX & STR_OCX
    Me.Caption = "About " & App.Title
    lblAppTitle.Caption = App.Title & " - " & OCX
    lblCopyright.Caption = "Copyright © 1999/" & Format$(Now, "yyyy") & ", Danilo Priore"
End Sub

Private Sub lblWeb_Click(Index As Integer)
    Call ShellExecute(Me.hwnd, "open", "http://priore.w3.to", "", "", 0)
End Sub

Private Sub tmrUnload_Timer()
    Unload Me
End Sub
