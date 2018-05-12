VERSION 5.00
Begin VB.Form frmSplash 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   3945
   ClientLeft      =   2595
   ClientTop       =   2790
   ClientWidth     =   5265
   ClipControls    =   0   'False
   Icon            =   "frmSplashFree.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3945
   ScaleWidth      =   5265
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tmrUnload 
      Interval        =   10000
      Left            =   4560
      Top             =   3360
   End
   Begin VB.Label lblAddress 
      BackStyle       =   0  'Transparent
      Caption         =   "Web: http://priore.w3.to"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Index           =   2
      Left            =   900
      MouseIcon       =   "frmSplashFree.frx":000C
      MousePointer    =   99  'Custom
      TabIndex        =   7
      Top             =   3480
      Width           =   3795
   End
   Begin VB.Label lblAddress 
      BackStyle       =   0  'Transparent
      Caption         =   "Internet: priore@w3.to"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Index           =   1
      Left            =   900
      MouseIcon       =   "frmSplashFree.frx":0316
      MousePointer    =   99  'Custom
      TabIndex        =   6
      Top             =   3180
      Width           =   3795
   End
   Begin VB.Label lblAddress 
      BackStyle       =   0  'Transparent
      Caption         =   "Danilo Priore"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   0
      Left            =   900
      TabIndex        =   5
      Top             =   2940
      Width           =   3795
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000006&
      Index           =   1
      X1              =   900
      X2              =   5160
      Y1              =   2820
      Y2              =   2820
   End
   Begin VB.Label lblInfo 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmSplashFree.frx":0620
      ForeColor       =   &H80000008&
      Height          =   735
      Index           =   2
      Left            =   900
      TabIndex        =   4
      Top             =   2100
      Width           =   4155
   End
   Begin VB.Label lblInfo 
      BackStyle       =   0  'Transparent
      Caption         =   "Complete licensing information can be found in this control's help file."
      ForeColor       =   &H80000008&
      Height          =   555
      Index           =   1
      Left            =   900
      TabIndex        =   3
      Top             =   1560
      Width           =   4155
   End
   Begin VB.Label lblInfo 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmSplashFree.frx":06AF
      ForeColor       =   &H80000008&
      Height          =   735
      Index           =   0
      Left            =   900
      TabIndex        =   2
      Top             =   840
      Width           =   4215
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000006&
      Index           =   0
      X1              =   900
      X2              =   5160
      Y1              =   780
      Y2              =   780
   End
   Begin VB.Label lblCopyright 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "lblCopyright"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   900
      TabIndex        =   1
      Top             =   480
      Width           =   1005
   End
   Begin VB.Label lblAppTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "lblAppTitle"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   900
      TabIndex        =   0
      Top             =   240
      Width           =   915
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   180
      Picture         =   "frmSplashFree.frx":073E
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
    Me.Caption = "About"
    lblAppTitle.Caption = App.Title
    lblCopyright.Caption = "Copyright © 1999/" & Format$(Now, "yyyy") & ", Danilo Priore"
End Sub

Private Sub lblAddress_Click(Index As Integer)
    Dim lnk As String
    Select Case Index
        Case 1: lnk = "mailto:priore.w3.to"
        Case 2: lnk = "http://priore.w3.to"
    End Select
    ShellExecute Me.hwnd, "open", lnk, "", "", vbNormal
End Sub

Private Sub tmrUnload_Timer()
    Unload Me
End Sub
