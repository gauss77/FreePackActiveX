VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{B1080A3D-F27D-11D2-9939-000000000000}#2.0#0"; "wininfo.ocx"
Begin VB.Form frmWinInfo 
   Caption         =   "Windows Info"
   ClientHeight    =   1815
   ClientLeft      =   2055
   ClientTop       =   3585
   ClientWidth     =   6645
   Icon            =   "frmWinInfo.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   1815
   ScaleWidth      =   6645
   Begin WinInfoCtl.WindowsInfo WindowsInfo1 
      Left            =   5760
      Top             =   180
      _ExtentX        =   847
      _ExtentY        =   847
      Caption         =   "WindowsInfo"
      Count           =   69
      WinIndex        =   1
      WinTop          =   338
      WinWidth        =   67
      WinHeight       =   17
   End
   Begin VB.Timer Timer1 
      Left            =   6060
      Top             =   720
   End
   Begin ComctlLib.ListView ListView1 
      Height          =   1695
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   5475
      _ExtentX        =   9657
      _ExtentY        =   2990
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      _Version        =   327682
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
End
Attribute VB_Name = "frmWinInfo"
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

Dim WCount As Long

Private Sub Form_Load()
    With ListView1
        .ColumnHeaders.Add , , "Title"
        .ColumnHeaders.Add , , "Left"
        .ColumnHeaders.Add , , "Width"
        .ColumnHeaders.Add , , "Top"
        .ColumnHeaders.Add , , "Height"
        .ColumnHeaders.Add , , "Active"
        .ColumnHeaders.Add , , "Enabled"
        .ColumnHeaders.Add , , "Iconic"
        .ColumnHeaders.Add , , "Visible"
        .ColumnHeaders.Add , , "Zoomed"
        .ColumnHeaders.Add , , "Handle"
        .ColumnHeaders.Add , , "Class"
        .View = lvwReport
    End With
    
    Timer1.Interval = 1000
End Sub

Private Sub Form_Resize()
    On Local Error Resume Next
    With ListView1
        .Height = Me.Height - 525
        .Width = Me.Width - 285
    End With
End Sub

Private Sub ListView1_ColumnClick(ByVal ColumnHeader As ComctlLib.ColumnHeader)
    With ListView1
        .SortKey = ColumnHeader.Index - 1
        .Sorted = True
    End With
End Sub

Private Sub Timer1_Timer()
    Dim i As Long
    Dim lv As ListItem
    
    MousePointer = vbHourglass
    With WindowsInfo1
        ' aggiorna il controllo e quindi recupera
        ' i vari dati di tutte le finestre
        .Refresh
        ' controlla se ci sono finestre
        If .Count > 0 Then
            ' si ci sono delle finestre
            ' controlla se il numero delle finestre
            ' è variato
            If WCount <> .Count Then
                ' si il numero delle finestre è variato
                WCount = .Count
                ' aggiorna la listview
                ListView1.ListItems.Clear
                For i = 1 To .Count
                    .WinIndex = i + 1
                    Set lv = ListView1.ListItems.Add(, , .Caption)
                    lv.SubItems(1) = .WinLeft
                    lv.SubItems(2) = .WinWidth
                    lv.SubItems(3) = .WinTop
                    lv.SubItems(4) = .WinHeight
                    lv.SubItems(5) = .IsActive
                    lv.SubItems(6) = .IsEnabled
                    lv.SubItems(7) = .IsIconic
                    lv.SubItems(8) = .IsVisible
                    lv.SubItems(9) = .IsZoomed
                    lv.SubItems(10) = .Handle
                    lv.SubItems(11) = .ClassName
                Next
            End If
        End If
    End With
    MousePointer = vbDefault
End Sub
