VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{918AABEF-81E0-11D6-9068-0080C88A8003}#2.0#0"; "WMInterfaceXPFree.ocx"
Begin VB.Form frmSposta 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Move Folder"
   ClientHeight    =   3975
   ClientLeft      =   1650
   ClientTop       =   1545
   ClientWidth     =   5040
   Icon            =   "frmSposta.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3975
   ScaleWidth      =   5040
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin WMInterfaceXPFree.CommandXPFree cmdOK 
      Default         =   -1  'True
      Height          =   315
      Left            =   3780
      TabIndex        =   0
      Top             =   420
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   556
      Caption         =   "&Ok"
   End
   Begin WMInterfaceXPFree.CommandXPFree cmdCancel 
      Cancel          =   -1  'True
      Height          =   315
      Left            =   3780
      TabIndex        =   1
      Top             =   840
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   556
      Caption         =   "&Cancel"
   End
   Begin MSComctlLib.TreeView tvwCat 
      DragIcon        =   "frmSposta.frx":2372
      Height          =   3435
      Left            =   120
      TabIndex        =   3
      Top             =   420
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   6059
      _Version        =   393217
      HideSelection   =   0   'False
      Sorted          =   -1  'True
      Style           =   7
      ImageList       =   "imlCat"
      Appearance      =   1
   End
   Begin MSComctlLib.ImageList imlCat 
      Left            =   4020
      Top             =   2040
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSposta.frx":2C3C
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSposta.frx":4FBE
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSposta.frx":7340
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label lblPrompt 
      AutoSize        =   -1  'True
      Caption         =   "Move the folder selected in the folder:"
      Height          =   195
      Left            =   120
      TabIndex        =   2
      Top             =   180
      Width           =   2670
   End
End
Attribute VB_Name = "frmSposta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Node As Node
Public Prompt As String
Public Title As String

Public Property Set Nodes(newNodes As Nodes)
    Dim itm As Node
    
    With tvwCat
        .Visible = False
        For Each itm In newNodes
            If Not itm.Parent Is Nothing Then
                .Nodes.Add itm.Parent.Key, tvwChild, itm.Key, itm.Text, itm.Image, itm.SelectedImage
            Else
                .Nodes.Add , , itm.Key, itm.Text, itm.Image, itm.SelectedImage
            End If
        Next
        For Each itm In .Nodes
            If itm.Children > 0 Then itm.Sorted = True
        Next
        .Nodes(1).Expanded = True
        .Visible = True
    End With
End Property

Private Sub cmdCancel_Click()
    Set Node = Nothing
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Set Node = tvwCat.SelectedItem
    Unload Me
End Sub

Private Sub Form_Load()
    Caption = Title
    lblPrompt.Caption = Prompt
    Set Node = Nothing
    cmdOK.Enabled = False
End Sub

Private Sub tvwCat_Collapse(ByVal Node As MSComctlLib.Node)
    If Node.Parent Is Nothing Then Node.Expanded = True
End Sub

Private Sub tvwCat_NodeClick(ByVal Node As MSComctlLib.Node)
    cmdOK.Enabled = True
End Sub

