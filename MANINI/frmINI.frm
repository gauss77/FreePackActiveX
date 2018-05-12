VERSION 5.00
Object = "{B108097D-F27D-11D2-9939-000000000000}#1.0#0"; "MANINI.OCX"
Begin VB.Form frmINI 
   Caption         =   "INI Manager"
   ClientHeight    =   1815
   ClientLeft      =   3690
   ClientTop       =   2685
   ClientWidth     =   3945
   Icon            =   "frmINI.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   1815
   ScaleWidth      =   3945
   Begin VB.TextBox txtValue 
      Height          =   315
      Left            =   1200
      TabIndex        =   9
      Top             =   1380
      Width           =   1215
   End
   Begin VB.CommandButton cmdWrite 
      Caption         =   "&Write"
      Height          =   375
      Left            =   2640
      TabIndex        =   7
      Top             =   600
      Width           =   1215
   End
   Begin VB.CommandButton cmdRead 
      Caption         =   "&Read"
      Height          =   375
      Left            =   2640
      TabIndex        =   6
      Top             =   120
      Width           =   1215
   End
   Begin VB.TextBox txtKey 
      Height          =   315
      Left            =   1200
      TabIndex        =   5
      Top             =   960
      Width           =   1215
   End
   Begin VB.TextBox txtSection 
      Height          =   315
      Left            =   1200
      TabIndex        =   3
      Top             =   540
      Width           =   1215
   End
   Begin VB.TextBox txtFilename 
      Height          =   315
      Left            =   1200
      TabIndex        =   1
      Top             =   120
      Width           =   1215
   End
   Begin ManagerINICtl.FileINI FileINI1 
      Left            =   2700
      Top             =   1080
      _ExtentX        =   847
      _ExtentY        =   847
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Value"
      Height          =   195
      Left            =   660
      TabIndex        =   8
      Top             =   1440
      Width           =   405
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Key"
      Height          =   195
      Left            =   780
      TabIndex        =   4
      Top             =   1020
      Width           =   270
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Section"
      Height          =   195
      Left            =   540
      TabIndex        =   2
      Top             =   600
      Width           =   540
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "INI File Name"
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   180
      Width           =   960
   End
End
Attribute VB_Name = "frmINI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdRead_Click()
    MousePointer = vbHourglass
    With FileINI1
        ' imposta il nome del file ini
        .FileName = txtFilename.Text
        ' imposta il valore di default
        .Default = vbNullString
        ' imposta la sezione
        .Section = txtSection.Text
        ' imposta la chiave
        .Key = txtKey.Text
        ' legge il valore (la lettura diretta
        ' della proprietà attiva la lettura
        ' automatica dal file ini)
        txtValue.Text = .Value
        ' in alternativa è possibile usare anche
        ' il metodo ReadINI
    End With
    MousePointer = vbDefault
End Sub

Private Sub cmdWrite_Click()
    MousePointer = vbHourglass
    With FileINI1
        ' imposta il nome del file ini
        .FileName = txtFilename.Text
        ' imposta la sezione
        .Section = txtSection.Text
        ' imposta la chiave
        .Key = txtKey.Text
        ' scrive il valore (l'impostazione
        ' della proprietà attiva la scrittura
        ' automatica nel file ini)
        .Value = txtValue.Text
        ' in alternativa è possibile usare anche
        ' il metodo WriteINI
    End With
    MousePointer = vbDefault
End Sub
