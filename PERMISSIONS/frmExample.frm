VERSION 5.00
Begin VB.Form frmExample 
   Caption         =   "NT Permissions Example"
   ClientHeight    =   4515
   ClientLeft      =   1920
   ClientTop       =   1635
   ClientWidth     =   6030
   Icon            =   "frmExample.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4515
   ScaleWidth      =   6030
   Begin VB.TextBox txtACL 
      Height          =   915
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   6
      Top             =   3540
      Width           =   5715
   End
   Begin VB.FileListBox File1 
      Height          =   3015
      Left            =   3060
      TabIndex        =   5
      Top             =   420
      Width           =   2835
   End
   Begin VB.DirListBox Dir1 
      Height          =   2565
      Left            =   120
      TabIndex        =   3
      Top             =   780
      Width           =   2775
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   600
      TabIndex        =   1
      Top             =   120
      Width           =   2295
   End
   Begin VB.Label lblFiles 
      AutoSize        =   -1  'True
      Caption         =   "Files"
      Height          =   195
      Left            =   3060
      TabIndex        =   4
      Top             =   180
      Width           =   315
   End
   Begin VB.Label lplDir 
      AutoSize        =   -1  'True
      Caption         =   "Directory"
      Height          =   195
      Left            =   120
      TabIndex        =   2
      Top             =   540
      Width           =   630
   End
   Begin VB.Label lblDrive 
      AutoSize        =   -1  'True
      Caption         =   "Drive:"
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   180
      Width           =   420
   End
End
Attribute VB_Name = "frmExample"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Dir1_Change()
    File1.Path = Dir1.Path
End Sub

Private Sub Drive1_Change()
    Dir1.Path = Drive1.Drive
End Sub

Private Sub File1_Click()
    Dim usr As UserGroup
    Dim acl As UsersGroups
    
    Set acl = ACLEdit1.GetACL(File1.Path & "\" & File1.FileName)
    If Not acl Is Nothing Then
        txtACL.Text = vbNullString
        For Each usr In acl
            txtACL.Text = txtACL.Text & usr.Name & " (" & usr.Access & ")" & vbCrLf
        Next
    End If
End Sub
