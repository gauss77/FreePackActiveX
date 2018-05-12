VERSION 5.00
Object = "{B1080930-F27D-11D2-9939-000000000000}#1.0#0"; "DRVINFO.OCX"
Begin VB.Form frmDrive 
   Caption         =   "Drive Info"
   ClientHeight    =   6015
   ClientLeft      =   4995
   ClientTop       =   2175
   ClientWidth     =   2865
   Icon            =   "frmDrive.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6015
   ScaleWidth      =   2865
   Begin DriveInfoCtl.DriveInfo DriveInfo1 
      Left            =   1920
      Top             =   1320
      _ExtentX        =   847
      _ExtentY        =   847
      DriveType       =   "Hard-Disk"
      Size            =   2043
      Free            =   944
      Serial          =   1879048192
      Label           =   "WINDOWS"
   End
   Begin VB.CommandButton cmdStart 
      Caption         =   "&Start"
      Default         =   -1  'True
      Height          =   375
      Left            =   1500
      TabIndex        =   14
      Top             =   5520
      Width           =   1215
   End
   Begin VB.ComboBox cbFormatType 
      Height          =   315
      Left            =   120
      TabIndex        =   13
      Top             =   4920
      Width           =   2595
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   120
      TabIndex        =   1
      Top             =   420
      Width           =   1215
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      Index           =   1
      X1              =   60
      X2              =   2760
      Y1              =   4440
      Y2              =   4440
   End
   Begin VB.Label lblFormatType 
      AutoSize        =   -1  'True
      Caption         =   "Format Type"
      Height          =   195
      Left            =   120
      TabIndex        =   12
      Top             =   4620
      Width           =   885
   End
   Begin VB.Label lblDriveLabel 
      AutoSize        =   -1  'True
      Caption         =   "Drive Label"
      Height          =   195
      Left            =   120
      TabIndex        =   11
      Top             =   3600
      Width           =   810
   End
   Begin VB.Label lblLabel 
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Left            =   120
      TabIndex        =   10
      Top             =   3900
      Width           =   2595
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      Index           =   0
      X1              =   60
      X2              =   2760
      Y1              =   960
      Y2              =   960
   End
   Begin VB.Label lblSerial 
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Left            =   120
      TabIndex        =   9
      Top             =   3120
      Width           =   1215
   End
   Begin VB.Label lblSerialNumber 
      AutoSize        =   -1  'True
      Caption         =   "Serial Number"
      Height          =   195
      Left            =   120
      TabIndex        =   8
      Top             =   2820
      Width           =   990
   End
   Begin VB.Label lblFree 
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Left            =   1500
      TabIndex        =   7
      Top             =   2340
      Width           =   1215
   End
   Begin VB.Label lblFreeSpace 
      AutoSize        =   -1  'True
      Caption         =   "Free Space"
      Height          =   195
      Left            =   1500
      TabIndex        =   6
      Top             =   2040
      Width           =   825
   End
   Begin VB.Label lblSize 
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Left            =   120
      TabIndex        =   5
      Top             =   2340
      Width           =   1215
   End
   Begin VB.Label lblDriveSize 
      AutoSize        =   -1  'True
      Caption         =   "Drive Size"
      Height          =   195
      Left            =   120
      TabIndex        =   4
      Top             =   2040
      Width           =   720
   End
   Begin VB.Label lblType 
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Left            =   120
      TabIndex        =   3
      Top             =   1560
      Width           =   1215
   End
   Begin VB.Label lblDriveType 
      AutoSize        =   -1  'True
      Caption         =   "Drive Type"
      Height          =   195
      Left            =   120
      TabIndex        =   2
      Top             =   1260
      Width           =   780
   End
   Begin VB.Label lblDrive 
      AutoSize        =   -1  'True
      Caption         =   "Drive"
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   375
   End
End
Attribute VB_Name = "frmDrive"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cbFormatType_Change()
    DriveInfo1.FormatType = cbFormatType.ListIndex
End Sub

Private Sub cmdStart_Click()
    DriveInfo1.Start
End Sub

Private Sub Drive1_Change()
    DriveInfo1.Drive = Drive1.Drive & "\"
    lblType.Caption = DriveInfo1.DriveType
    lblSize.Caption = DriveInfo1.Size
    lblFree.Caption = DriveInfo1.Free
    lblSerial.Caption = DriveInfo1.Serial
    lblLabel.Caption = DriveInfo1.Label
End Sub

Private Sub Form_Load()
    cbFormatType.AddItem "Quick format"
    cbFormatType.AddItem "Complete format"
    cbFormatType.AddItem "Boot disk"
    cbFormatType.ListIndex = 0
    
    Call Drive1_Change
End Sub
