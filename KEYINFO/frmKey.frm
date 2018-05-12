VERSION 5.00
Object = "{B1080969-F27D-11D2-9939-000000000000}#1.0#0"; "KEYINFO.OCX"
Begin VB.Form frmKey 
   Caption         =   "Key Info"
   ClientHeight    =   1230
   ClientLeft      =   3075
   ClientTop       =   2265
   ClientWidth     =   4095
   Icon            =   "frmKey.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   1230
   ScaleWidth      =   4095
   Begin VB.Timer Timer1 
      Left            =   3420
      Top             =   240
   End
   Begin KeyInfoCtl.KeyboardInfo KeyboardInfo1 
      Left            =   2700
      Top             =   60
      _ExtentX        =   847
      _ExtentY        =   847
      KeyboardFunctions=   12
      KeyboardType    =   "Enhanced 101 or 102 key"
   End
   Begin VB.Label lblFun 
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Left            =   3180
      TabIndex        =   9
      Top             =   840
      Width           =   375
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Keyboard Functions"
      Height          =   195
      Left            =   1620
      TabIndex        =   8
      Top             =   900
      Width           =   1410
   End
   Begin VB.Label lblKeyType 
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Left            =   1620
      TabIndex        =   7
      Top             =   420
      Width           =   2355
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Keyboard Type"
      Height          =   195
      Left            =   1620
      TabIndex        =   6
      Top             =   180
      Width           =   1080
   End
   Begin VB.Label lblScroll 
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Left            =   900
      TabIndex        =   5
      Top             =   840
      Width           =   555
   End
   Begin VB.Label lblNum 
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Left            =   900
      TabIndex        =   4
      Top             =   480
      Width           =   555
   End
   Begin VB.Label lblCaps 
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Left            =   900
      TabIndex        =   3
      Top             =   120
      Width           =   555
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "SCROLL"
      Height          =   195
      Left            =   120
      TabIndex        =   2
      Top             =   900
      Width           =   630
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "NUM"
      Height          =   195
      Left            =   360
      TabIndex        =   1
      Top             =   540
      Width           =   375
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "CAPS"
      Height          =   195
      Left            =   360
      TabIndex        =   0
      Top             =   180
      Width           =   420
   End
End
Attribute VB_Name = "frmKey"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub Form_Load()
    ' attiva il timer per in controllo continuo
    ' dello stato della tastiera
    Timer1.Interval = 500
    ' visualizza il tipo di tastiera (non è necessario
    ' includere questo codice nel timer perche il tipo
    ' di tastiera rimane invariata)
    With KeyboardInfo1
        lblKeyType.Caption = .KeyboardType
        lblFun.Caption = .KeyboardFunctions
    End With
End Sub

Private Sub Timer1_Timer()
    ' visualizza lo stato dei tasti
    With KeyboardInfo1
        ' caps lock
        lblCaps.Caption = .CapsState
        ' num lock
        lblNum.Caption = .NumState
        ' scroll lock
        lblScroll.Caption = .ScrollState
    End With
End Sub
