VERSION 5.00
Object = "*\APicTile.vbp"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmPicture 
   Caption         =   "Picture Tile"
   ClientHeight    =   4800
   ClientLeft      =   1560
   ClientTop       =   1830
   ClientWidth     =   7305
   Icon            =   "frmPicture.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4800
   ScaleWidth      =   7305
   Begin ComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7305
      _ExtentX        =   12885
      _ExtentY        =   741
      Appearance      =   1
      ImageList       =   "ImageList1"
      _Version        =   327682
      BeginProperty Buttons {0713E452-850A-101B-AFC0-4210102A8DA7} 
         NumButtons      =   2
         BeginProperty Button1 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button2 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.ToolTipText     =   "Exit"
            Object.Tag             =   ""
            ImageIndex      =   1
         EndProperty
      EndProperty
   End
   Begin PictureTileCtl.PictureTile PictureTile1 
      Left            =   360
      Top             =   1140
      _ExtentX        =   2672
      _ExtentY        =   2672
      Picture         =   "frmPicture.frx":08CA
      ScaleHeight     =   1515
      ScaleMode       =   0
      ScaleWidth      =   1515
   End
   Begin ComctlLib.ImageList ImageList1 
      Left            =   3360
      Top             =   2100
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   1
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmPicture.frx":1264
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuExit 
         Caption         =   "&Exit"
      End
   End
End
Attribute VB_Name = "frmPicture"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    PictureTile1.StartY = Toolbar1.Height
    'PictureTile1.Picture = LoadPicture("newback.gif")
    PictureTile1.Refresh
End Sub

Private Sub mnuExit_Click()
    End
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As ComctlLib.Button)
    End
End Sub
