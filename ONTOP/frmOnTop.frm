VERSION 5.00
Object = "{B1080B87-F27D-11D2-9939-000000000000}#1.0#0"; "ONTOP.OCX"
Begin VB.Form frmOnTop 
   Caption         =   "Window Manager"
   ClientHeight    =   3465
   ClientLeft      =   6015
   ClientTop       =   3855
   ClientWidth     =   5160
   Icon            =   "frmOnTop.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3465
   ScaleWidth      =   5160
   Begin WinManagerCtl.WindowSet WindowSet1 
      Left            =   3600
      Top             =   1980
      _ExtentX        =   847
      _ExtentY        =   847
   End
   Begin VB.Label Label1 
      Caption         =   "Example of Windows OnTop, flash the title bar, removing close item from system menu and show with special effects."
      Height          =   915
      Left            =   360
      TabIndex        =   0
      Top             =   660
      Width           =   3900
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
      End
   End
End
Attribute VB_Name = "frmOnTop"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    WindowSet1.Center = True
    WindowSet1.Value = True
    WindowSet1.RemoveClose = True
    WindowSet1.Effect = wsExplode
    WindowSet1.Flash = True
    WindowSet1.Show
End Sub

Private Sub mnuExit_Click()
    End
End Sub
