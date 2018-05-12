VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{B10809A3-F27D-11D2-9939-000000000000}#1.0#0"; "MNUPIC.OCX"
Begin VB.Form frmMenuPic 
   Caption         =   "Menu Picture"
   ClientHeight    =   2970
   ClientLeft      =   1650
   ClientTop       =   1830
   ClientWidth     =   5205
   Icon            =   "frmMenuPic.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   2970
   ScaleWidth      =   5205
   Begin MnuImageCtl.MenuImage MenuImage1 
      Left            =   2520
      Top             =   1440
      _ExtentX        =   847
      _ExtentY        =   847
   End
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   900
      Left            =   660
      Picture         =   "frmMenuPic.frx":030A
      ScaleHeight     =   900
      ScaleWidth      =   900
      TabIndex        =   0
      Top             =   1200
      Visible         =   0   'False
      Width           =   900
   End
   Begin ComctlLib.ImageList imlPic 
      Left            =   2220
      Top             =   780
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   7
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMenuPic.frx":0E95
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMenuPic.frx":11E7
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMenuPic.frx":1539
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMenuPic.frx":188B
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMenuPic.frx":1BDD
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMenuPic.frx":1F2F
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMenuPic.frx":2281
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileNew 
         Caption         =   "&New"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuFileOpen 
         Caption         =   "&Open"
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuFileBar1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileSave 
         Caption         =   "&Save"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuFileBar3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFilePrint 
         Caption         =   "&Print..."
         Shortcut        =   ^P
      End
      Begin VB.Menu mnuFileBar6 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "&Edit"
      Begin VB.Menu mnuEditCut 
         Caption         =   "C&ut"
         Shortcut        =   ^X
      End
      Begin VB.Menu mnuEditCopy 
         Caption         =   "&Copy"
         Shortcut        =   ^C
      End
      Begin VB.Menu mnuEditPaste 
         Caption         =   "&Paste"
         Shortcut        =   ^V
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuInfo 
         Caption         =   "&Info"
      End
   End
End
Attribute VB_Name = "frmMenuPic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    Me.Show
    MenuImage1.MenuCaption = "File"
    MenuImage1.ItemCaption = "New"
    Set MenuImage1.ImageChecked = imlPic.ListImages(1).Picture
    Set MenuImage1.ImageUnchecked = imlPic.ListImages(1).Picture
    MenuImage1.Refresh

    MenuImage1.MenuCaption = "File"
    MenuImage1.ItemCaption = "Open"
    Set MenuImage1.ImageChecked = imlPic.ListImages(2).Picture
    Set MenuImage1.ImageUnchecked = imlPic.ListImages(2).Picture
    MenuImage1.Refresh

    MenuImage1.MenuCaption = "File"
    MenuImage1.ItemCaption = "Save"
    Set MenuImage1.ImageChecked = imlPic.ListImages(3).Picture
    Set MenuImage1.ImageUnchecked = imlPic.ListImages(3).Picture
    MenuImage1.Refresh

    MenuImage1.MenuCaption = "Help"
    MenuImage1.ItemCaption = "Info"
    Set MenuImage1.Image = Picture1.Picture
    MenuImage1.Modify
End Sub

Private Sub mnuFileExit_Click()
    End
End Sub
