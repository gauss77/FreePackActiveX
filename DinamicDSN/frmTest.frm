VERSION 5.00
Begin VB.Form frmTest 
   Caption         =   "Form1"
   ClientHeight    =   7605
   ClientLeft      =   1650
   ClientTop       =   2145
   ClientWidth     =   6975
   LinkTopic       =   "Form1"
   ScaleHeight     =   7605
   ScaleWidth      =   6975
   Begin VB.CommandButton Command3 
      Caption         =   "Command3"
      Height          =   495
      Left            =   3300
      TabIndex        =   2
      Top             =   720
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   495
      Left            =   2040
      TabIndex        =   1
      Top             =   720
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   720
      TabIndex        =   0
      Top             =   720
      Width           =   1215
   End
End
Attribute VB_Name = "frmTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim d As New DINAMICDSN.DSN

Private Sub Command1_Click()
    MsgBox (d.AddDSN("Test", "Microsoft Access Driver (*.mdb)", "", "\\10.0.0.3\database\ASSIDAI2005.mdb", "Admin", ""))
End Sub

Private Sub Command2_Click()
    MsgBox (d.ConfigDSN("Test", "Microsoft Access Driver (*.mdb)", "", "c:\temp\db2_x.mdb", "Admin", ""))
End Sub

Private Sub Command3_Click()
    MsgBox (d.RemoveDSN("Test", "Microsoft Access Driver (*.mdb)"))
End Sub
