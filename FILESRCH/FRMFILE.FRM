VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "*\AFILESRCH.VBP"
Begin VB.Form frmFile 
   Caption         =   "File Search"
   ClientHeight    =   3870
   ClientLeft      =   4515
   ClientTop       =   3045
   ClientWidth     =   4575
   Icon            =   "frmFile.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3870
   ScaleWidth      =   4575
   Begin ComctlLib.ListView ListView1 
      Height          =   1335
      Left            =   120
      TabIndex        =   5
      Top             =   960
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   2355
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   327682
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "Search"
      Default         =   -1  'True
      Height          =   375
      Left            =   3240
      TabIndex        =   4
      Top             =   60
      Width           =   1215
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   900
      TabIndex        =   3
      Top             =   480
      Width           =   2175
   End
   Begin VB.TextBox txtNome 
      Height          =   315
      Left            =   720
      TabIndex        =   1
      Top             =   60
      Width           =   2355
   End
   Begin FileSearchCtl.FileSearch FileSearch1 
      Left            =   3360
      Top             =   600
      _ExtentX        =   847
      _ExtentY        =   847
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      X1              =   120
      X2              =   4440
      Y1              =   2460
      Y2              =   2460
   End
   Begin VB.Label Label2 
      Height          =   495
      Left            =   180
      TabIndex        =   7
      Top             =   3240
      Width           =   4155
   End
   Begin VB.Label Label1 
      Height          =   495
      Left            =   180
      TabIndex        =   6
      Top             =   2640
      Width           =   4155
   End
   Begin VB.Label lblDrive 
      AutoSize        =   -1  'True
      Caption         =   "Drive:"
      Height          =   195
      Left            =   120
      TabIndex        =   2
      Top             =   540
      Width           =   420
   End
   Begin VB.Label lblName 
      AutoSize        =   -1  'True
      Caption         =   "Name:"
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   465
   End
End
Attribute VB_Name = "frmFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub cmdSearch_Click()
    Dim i As Long
    Dim Drv As String
    Dim li As ListItem
    
    Dim f As File
    Dim ff As Files
    
    MousePointer = vbHourglass
    If Len(txtNome.Text) > 0 Then
        Drv = Drive1.Drive
        If InStr(Drv, ":") > 0 Then Drv = Left$(Drv, InStr(Drv, ":"))
        If Right$(Drv, 1) <> "\" Then Drv = Drv & "\"
        With FileSearch1
            .Path = Drv
            .FileName = txtNome.Text
            Set ff = .Search
            If ff.Count > 0 Then
                For Each f In ff
                    Set li = ListView1.ListItems.Add(, , f.FileName)
                    li.SubItems(1) = f.Path
                    li.SubItems(2) = FileDateTime(f.Path & f.FileName)
                    li.SubItems(3) = FileLen(f.Path & f.FileName)
                Next
            Else
                MsgBox "File not found.", vbInformation
            End If
        End With
    End If
    MousePointer = vbDefault
End Sub

Private Sub FileSearch1_Progress(PathName As String, NameFile As String)
    With Label1
        .Caption = PathName
        .Refresh
    End With
    With Label2
        .Caption = NameFile
        .Refresh
    End With
End Sub

Private Sub Form_Load()
    
    With ListView1
        .ColumnHeaders.Add , , "File", ListView1.Width / 4
        .ColumnHeaders.Add , , "Path", ListView1.Width / 4
        .ColumnHeaders.Add , , "Date", ListView1.Width / 4
        .ColumnHeaders.Add , , "Size", ListView1.Width / 4
        .View = lvwReport
    End With
End Sub

