VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{B75B64BB-E252-4D5F-B1D9-7D06814BCA39}#3.0#0"; "NetInfo.ocx"
Begin VB.Form frmNet 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Net Info"
   ClientHeight    =   4005
   ClientLeft      =   3105
   ClientTop       =   1905
   ClientWidth     =   6375
   Icon            =   "frmNet.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4005
   ScaleWidth      =   6375
   ShowInTaskbar   =   0   'False
   Begin ComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   6375
      _ExtentX        =   11245
      _ExtentY        =   741
      Appearance      =   1
      ImageList       =   "ImageList1"
      _Version        =   327682
      BeginProperty Buttons {0713E452-850A-101B-AFC0-4210102A8DA7} 
         NumButtons      =   3
         BeginProperty Button1 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button2 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "connect"
            Object.ToolTipText     =   "Connect"
            Object.Tag             =   ""
            ImageIndex      =   1
         EndProperty
         BeginProperty Button3 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "disconnect"
            Object.ToolTipText     =   "Disconnect"
            Object.Tag             =   ""
            ImageIndex      =   2
         EndProperty
      EndProperty
   End
   Begin VB.ComboBox cboService 
      Height          =   315
      ItemData        =   "frmNet.frx":08CA
      Left            =   2700
      List            =   "frmNet.frx":08CC
      TabIndex        =   2
      Top             =   960
      Width           =   2055
   End
   Begin NetInfoCtl.NetInfo NetInfo1 
      Left            =   5100
      Top             =   2400
      _ExtentX        =   847
      _ExtentY        =   847
   End
   Begin VB.ListBox List1 
      Height          =   2205
      Left            =   120
      TabIndex        =   4
      Top             =   1620
      Width           =   4275
   End
   Begin VB.CommandButton cmdList 
      Caption         =   "&List"
      Default         =   -1  'True
      Height          =   375
      Left            =   4920
      TabIndex        =   3
      Top             =   960
      Width           =   1215
   End
   Begin VB.TextBox txtDomain 
      Height          =   315
      Left            =   120
      TabIndex        =   1
      Top             =   960
      Width           =   2475
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Service"
      Height          =   195
      Index           =   1
      Left            =   2700
      TabIndex        =   6
      Top             =   660
      Width           =   540
   End
   Begin ComctlLib.ImageList ImageList1 
      Left            =   5340
      Top             =   1680
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   2
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmNet.frx":08CE
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmNet.frx":0BE8
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      X1              =   120
      X2              =   6240
      Y1              =   1500
      Y2              =   1500
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Domain Name"
      Height          =   195
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   660
      Width           =   1005
   End
End
Attribute VB_Name = "frmNet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Author:  Danilo Priore
'Email:   support@ prioregroup.com
'URL    : http://www.prioregroup.com
'
'This code is written and distributed under
'the GNU General Public License which means
'that its source code is freely-distributed
'and available to the general public.

Option Explicit

Private Sub cmdList_Click()
    Dim i As Long
    Dim itm As NetInfoCtl.Item
    Dim itms As NetInfoCtl.Items
    
    MousePointer = vbHourglass
    If Len(txtDomain.Text) > 0 And Len(cboService.Text) > 0 Then
        NetInfo1.Domain = txtDomain.Text
        NetInfo1.InfoType = cboService.ItemData(cboService.ListIndex)
        Set itms = NetInfo1.GetInfo
        
        List1.Clear
        If Not itms Is Nothing Then
            For Each itm In itms
                List1.AddItem itm.ItemName & IIf(Len(itm.Description), " (" & itm.Description & ")", vbNullString)
            Next
        Else
            List1.AddItem "(not item found)"
        End If
    End If
    MousePointer = vbDefault
End Sub

Private Sub Form_Load()
    With cboService
        .AddItem "Computers"
        .ItemData(.NewIndex) = 0
        
        .AddItem "Groups"
        .ItemData(.NewIndex) = 1
    
        .AddItem "Users"
        .ItemData(.NewIndex) = 2
    
        .AddItem "Services"
        .ItemData(.NewIndex) = 3
    
        .AddItem "Printers"
        .ItemData(.NewIndex) = 4
        
        .AddItem "All"
        .ItemData(.NewIndex) = 255
    End With
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As ComctlLib.Button)
    Select Case Button.Key
        Case "connect"
            NetInfo1.NetConnection
        Case "disconnect"
            NetInfo1.NetDisconnection
    End Select
End Sub
