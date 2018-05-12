VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "           "
   ClientHeight    =   2490
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5055
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2490
   ScaleWidth      =   5055
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   3840
      TabIndex        =   4
      Top             =   1980
      Width           =   975
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderStyle     =   6  'Inside Solid
      Index           =   3
      X1              =   240
      X2              =   4740
      Y1              =   1800
      Y2              =   1800
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      Index           =   2
      X1              =   255
      X2              =   4740
      Y1              =   1815
      Y2              =   1815
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderStyle     =   6  'Inside Solid
      Index           =   1
      X1              =   240
      X2              =   4740
      Y1              =   840
      Y2              =   840
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      Index           =   0
      X1              =   255
      X2              =   4740
      Y1              =   855
      Y2              =   855
   End
   Begin VB.Label lblDisclaimer 
      Caption         =   $"frmAbout.frx":000C
      ForeColor       =   &H00000000&
      Height          =   885
      Left            =   300
      TabIndex        =   3
      Top             =   960
      WhatsThisHelpID =   30
      Width           =   4410
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   180
      Picture         =   "frmAbout.frx":00EA
      Top             =   180
      WhatsThisHelpID =   20
      Width           =   480
   End
   Begin VB.Label lblAppTitle 
      AutoSize        =   -1  'True
      Caption         =   "lblAppTitle"
      Height          =   195
      Left            =   900
      TabIndex        =   2
      Top             =   180
      WhatsThisHelpID =   20
      Width           =   735
   End
   Begin VB.Label lblCopyright 
      AutoSize        =   -1  'True
      Caption         =   "lblCopyright"
      Height          =   195
      Left            =   900
      TabIndex        =   1
      Top             =   420
      WhatsThisHelpID =   20
      Width           =   810
   End
   Begin VB.Label lblAddress 
      AutoSize        =   -1  'True
      Caption         =   "Web: http://www.prioregroup.com"
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
      Index           =   2
      Left            =   300
      MouseIcon       =   "frmAbout.frx":03F4
      MousePointer    =   99  'Custom
      TabIndex        =   0
      Top             =   2040
      WhatsThisHelpID =   20
      Width           =   2445
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Private Sub cmdOk_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Me.Caption = "About " & App.Title
    lblAppTitle.Caption = App.Title & " - " & UCase$(App.EXEName)
    lblCopyright.Caption = "Copyright © 1998/" & Format$(Now, "yyyy") & ", Danilo Priore"
End Sub

Private Sub lblAddress_Click(Index As Integer)
    ShellExecute Me.hwnd, "open", "http://www.prioregroup.com", vbNullString, vbNullString, 5
End Sub
