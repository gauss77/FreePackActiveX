VERSION 5.00
Object = "*\A..\..\..\..\DOCUME~1\Sorgenti\OCX\NCARDINF\NCardInfl.vbp"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmNetSimple 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "NetCard Info Simple Example"
   ClientHeight    =   2670
   ClientLeft      =   1635
   ClientTop       =   1545
   ClientWidth     =   6300
   Icon            =   "frmNetSimple.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2670
   ScaleWidth      =   6300
   Begin VB.CommandButton Command1 
      Caption         =   "&Full"
      Height          =   375
      Left            =   5040
      TabIndex        =   0
      Top             =   60
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Caption         =   "&Select Interface"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   4935
      Begin VB.ComboBox Combo1 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         ItemData        =   "frmNetSimple.frx":030A
         Left            =   120
         List            =   "frmNetSimple.frx":030C
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   240
         Width           =   4695
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Interface Information"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1875
      Left            =   0
      TabIndex        =   2
      Top             =   720
      Width           =   4935
      Begin MSComctlLib.ListView lvInterfaceInfo 
         Height          =   1515
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   4695
         _ExtentX        =   8281
         _ExtentY        =   2672
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   0
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   200
      Left            =   5040
      Top             =   1020
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit"
      Height          =   375
      Left            =   5040
      TabIndex        =   1
      Top             =   540
      Width           =   1215
   End
   Begin NetCardInfoCtl.NCardInfo NCardInfo1 
      Left            =   5700
      Top             =   1080
      _ExtentX        =   847
      _ExtentY        =   847
   End
End
Attribute VB_Name = "frmNetSimple"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub Combo1_Click()
    Call UpdateInterfaceInfo(Combo1.ListIndex + 1)
End Sub

Private Sub Command1_Click()
    frmNetEx.Show
    Unload Me
End Sub

Private Sub Form_Load()
    '
    Dim i As Integer
    Dim objInterface As CInterface
    '
    'Configure the listview control
    '
    'Add column headers
    '
    lvInterfaceInfo.ColumnHeaders.Add , , "Parameter", 3000
    lvInterfaceInfo.ColumnHeaders.Add , , "Value", 1600
    '
    'Add listview items - interface parameters
    '
    With lvInterfaceInfo.ListItems
        .Add , , "Type of interface"
        .Add , , "Physical address of adapter"
        .Add , , "Operational status"
        .Add , , "Bytes received"
        .Add , , "Bytes send"
    End With
    '
    'Add descriptions of the network interfaces into the listbox control
    For i = 1 To NCardInfo1.Count
        Set objInterface = NCardInfo1.GetInfo(i)
        Combo1.AddItem objInterface.InterfaceDescription
    Next
    '
    'Define selected item in the listbox control
    Combo1.ListIndex = 0
End Sub

Private Sub UpdateInterfaceInfo(intIndex As Integer)
    '
    Dim objInterface As CInterface
    '
    If intIndex > NCardInfo1.Count Then
        Exit Sub
    End If
    '
    Set objInterface = NCardInfo1.GetInfo(intIndex)
    '
    With lvInterfaceInfo.ListItems
        Select Case objInterface.InterfaceType
            Case MIB_IF_TYPE_ETHERNET: .Item(1).SubItems(1) = "Ethernet"
            Case MIB_IF_TYPE_FDDI: .Item(1).SubItems(1) = "FDDI"
            Case MIB_IF_TYPE_LOOPBACK: .Item(1).SubItems(1) = "Loopback"
            Case MIB_IF_TYPE_OTHER: .Item(1).SubItems(1) = "Other"
            Case MIB_IF_TYPE_PPP: .Item(1).SubItems(1) = "PPP"
            Case MIB_IF_TYPE_SLIP: .Item(1).SubItems(1) = "SLIP"
            Case MIB_IF_TYPE_TOKENRING: .Item(1).SubItems(1) = "TokenRing"
        End Select
        .Item(2).SubItems(1) = objInterface.AdapterAddress
        Select Case objInterface.OperationalStatus
            Case MIB_IF_OPER_STATUS_CONNECTED: .Item(3).SubItems(1) = "Connected"
            Case MIB_IF_OPER_STATUS_CONNECTING: .Item(3).SubItems(1) = "Connecting"
            Case MIB_IF_OPER_STATUS_DISCONNECTED: .Item(3).SubItems(1) = "Disconnected"
            Case MIB_IF_OPER_STATUS_NON_OPERATIONAL: .Item(3).SubItems(1) = "Non operational"
            Case MIB_IF_OPER_STATUS_OPERATIONAL: .Item(3).SubItems(1) = "Operational"
            Case MIB_IF_OPER_STATUS_UNREACHABLE: .Item(3).SubItems(1) = "Unreachable"
        End Select
        .Item(4).SubItems(1) = Trim(Format(objInterface.BytesReceived, "### ### ### ###"))
        .Item(5).SubItems(1) = Trim(Format(objInterface.BytesSent, "### ### ### ###"))
    End With
End Sub

Private Sub Timer1_Timer()
    Call UpdateInterfaceInfo(Combo1.ListIndex + 1)
End Sub


