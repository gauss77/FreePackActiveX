VERSION 5.00
Object = "{7FCFB161-49D5-4D74-B0DD-8D3523BC16E9}#3.0#0"; "pfingerctl.ocx"
Begin VB.Form frmNew 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "New member"
   ClientHeight    =   3675
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4635
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3675
   ScaleWidth      =   4635
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save To..."
      Height          =   375
      Left            =   1680
      TabIndex        =   5
      Top             =   3120
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3120
      TabIndex        =   6
      Top             =   3120
      Width           =   1215
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "Ok"
      Default         =   -1  'True
      Height          =   375
      Left            =   240
      TabIndex        =   4
      Top             =   3120
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Height          =   2295
      Left            =   240
      TabIndex        =   2
      Top             =   660
      Width           =   4155
      Begin PFingerPrintCtl.FingerPrint FingerPrint1 
         Height          =   1635
         Left            =   2220
         Top             =   300
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   2884
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "Please, put fingerprint into reader to save your identity"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   915
         Left            =   180
         TabIndex        =   3
         Top             =   240
         Width           =   1755
      End
   End
   Begin VB.TextBox txtName 
      Height          =   315
      Left            =   1380
      MaxLength       =   100
      TabIndex        =   1
      Top             =   180
      Width           =   2955
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Member name :"
      Height          =   195
      Left            =   180
      TabIndex        =   0
      Top             =   240
      Width           =   1095
   End
End
Attribute VB_Name = "frmNew"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim aTemplate() As Byte

' used for database connection
Public DBConnection As ADODB.Connection

Private Sub cmdCancel_Click()
    FingerPrint1.Interval = 0   ' stop
    Unload Me                   ' exit
End Sub

Private Sub cmdOk_Click()
    Dim rs As ADODB.Recordset
    
    ' open table
    Set rs = New ADODB.Recordset
    rs.Open "Members", DBConnection, adOpenDynamic, adLockOptimistic
    
    ' new record and save data in database
    rs.AddNew
    rs.Fields("Name") = txtName.Text
    rs.Fields("Template").AppendChunk aTemplate
    rs.Update
    rs.Close
    Set rs = Nothing
    
    ' message
    MsgBox "Fingerprint save correctly.", vbInformation
    Unload Me
End Sub

Private Sub cmdSave_Click()
    Dim pic As StdPicture
    
    With FingerPrint1
        ' retrieve picture from current rawdata
        Set pic = .RawToPicture(.RawDataB)
        .ShowSavePicture pic, "Windows Bitmap|*.bmp", App.Title
    End With
End Sub

Private Sub FingerPrint1_FingerIn()
    FingerPrint1.Interval = 0               ' stop (freeze)
    aTemplate = FingerPrint1.TemplateDataB  ' save in to memory
    cmdOk.Enabled = True                    ' enable button
    cmdSave.Enabled = True                  ' enable button
End Sub

Private Sub Form_Load()
    cmdOk.Enabled = False           ' disable button
    cmdSave.Enabled = False         ' disable button
    FingerPrint1.Interval = 1000    ' start
End Sub
