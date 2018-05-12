VERSION 5.00
Object = "{7FCFB161-49D5-4D74-B0DD-8D3523BC16E9}#3.0#0"; "PFingerCtl.ocx"
Begin VB.Form frmMain 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Priore FingerPrint Sample Database Version"
   ClientHeight    =   4260
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   5910
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4260
   ScaleWidth      =   5910
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   2535
      Left            =   3120
      TabIndex        =   2
      Top             =   720
      Width           =   2595
      Begin PFingerPrintCtl.FingerPrint FingerPrint1 
         Height          =   2115
         Left            =   360
         Top             =   240
         Width           =   1875
         _ExtentX        =   3307
         _ExtentY        =   3731
      End
   End
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      Height          =   4020
      Left            =   120
      Picture         =   "frmMain.frx":0000
      ScaleHeight     =   3960
      ScaleWidth      =   2625
      TabIndex        =   0
      Top             =   120
      Width           =   2685
   End
   Begin VB.Label lblMember 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   3060
      TabIndex        =   3
      Top             =   3360
      Width           =   2655
   End
   Begin VB.Label lblMsg 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Please, put fingerprint into reader to verify your identity"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   3120
      TabIndex        =   1
      Top             =   180
      Width           =   2655
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu itmNew 
         Caption         =   "&New member..."
         Shortcut        =   ^N
      End
      Begin VB.Menu itmSep 
         Caption         =   "-"
      End
      Begin VB.Menu itmExternalFile 
         Caption         =   "Compare with external file..."
      End
      Begin VB.Menu itmSep1 
         Caption         =   "-"
      End
      Begin VB.Menu itmExit 
         Caption         =   "E&xit"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim db As ADODB.Connection
Dim rs As ADODB.Recordset

Private Sub FingerPrint1_FingerIn()
     Dim aTemplate() As Byte
     Dim lsize As Long
     
    ' this function compare current fingerprint
    ' with fingerprints in to database
     
     ' inizialize message
     lblMember.Caption = vbNullString
     
     ' stop
     FingerPrint1.Interval = 0
     
     ' select all records
     Set rs = New ADODB.Recordset
     rs.Open "Members", db, adOpenKeyset, adLockOptimistic, adCmdTable
     Do While Not rs.EOF
        ' retrieve fingerprint
        lsize = rs.Fields("Template").ActualSize
        aTemplate = rs.Fields("Template").GetChunk(lsize)
        ' verify current user with current fingerprint
        ' NOTE: not need the keep finger in to reader during this operation
        If FingerPrint1.VerifyFingerB(aTemplate()) Then
            lblMember.ForeColor = vbBlack
            lblMember.Caption = "Welcome " & rs.Fields("Name") & vbCrLf & "the system have recognized to you."
            lblMember.Refresh
            Exit Do
        End If
        rs.MoveNext
     Loop
     rs.Close
     Set rs = Nothing
     
     ' if not have recognized
    If Len(lblMember.Caption) = 0 Then
        lblMember.ForeColor = vbRed
        lblMember.Caption = "Sorry, the system not have recognized to you!"
    End If
     
     ' restart
     FingerPrint1.Interval = 1000
End Sub

Private Sub FingerPrint1_FingerOut()
    lblMember.Caption = vbNullString
End Sub

Private Sub Form_Load()
    ' open the database
    Set db = New ADODB.Connection
    db.Open "Provider=Microsoft.Jet.OLEDB.4.0" _
             & ";Data Source=" & App.Path & "\fingers.Mdb" _
             & ";Jet OLEDB:Engine Type=5", "Admin", ""
             
    ' start
    FingerPrint1.Interval = 1000
End Sub

Private Sub itmExit_Click()
    ' close database
    db.Close
    Set db = Nothing
    End
End Sub

Private Sub itmExternalFile_Click()
    Dim aTemplate() As Byte
    Dim lsize As Long
    Dim rawdata() As Byte
    Dim pic As StdPicture
    
    ' this function compare a external fingerprint image
    ' with fingerprints in to database
    With FingerPrint1
       Set pic = .ShowLoadPicture("Windows Bitmap|*.bmp", App.Title)
       If pic Is Nothing Then Exit Sub
    
        ' convert picture to rawdata
        ' NOTE: before ConvertToBW if need
        rawdata = .PictureToRaw(pic)
        
        ' inizialize message
        lblMember.Caption = vbNullString
        
        ' stop
        .Interval = 0
    
         ' select all records
         Set rs = New ADODB.Recordset
         rs.Open "Members", db, adOpenKeyset, adLockOptimistic, adCmdTable
         Do While Not rs.EOF
            ' retrieve fingerprint from database
            lsize = rs.Fields("Template").ActualSize
            aTemplate = rs.Fields("Template").GetChunk(lsize)
            ' verify picture raw data with database tamplate data
            If .VerifyFingerEx(aTemplate, rawdata) Then
                lblMember.ForeColor = vbBlack
                lblMember.Caption = "Welcome " & rs.Fields("Name") & vbCrLf & "the system have recognized to you."
                lblMember.Refresh
            End If
            rs.MoveNext
         Loop
         rs.Close
         Set rs = Nothing
         
         ' if not have recognized
        If Len(lblMember.Caption) = 0 Then
            lblMember.ForeColor = vbRed
            lblMember.Caption = "Sorry, the system not have recognized to you!"
        End If
         
         ' restart
         .Interval = 1000
    End With
End Sub

Private Sub itmNew_Click()
    FingerPrint1.Interval = 0       ' stop
    Set frmNew.DBConnection = db    ' set the current DB connection
    frmNew.Show vbModal             ' show form
    FingerPrint1.Interval = 100     ' restart
End Sub
