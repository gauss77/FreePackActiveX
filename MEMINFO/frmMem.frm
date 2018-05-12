VERSION 5.00
Object = "{C224A181-F335-11D2-993D-000000000000}#1.0#0"; "MEMINFO.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmMem 
   Caption         =   "Memory Info"
   ClientHeight    =   1290
   ClientLeft      =   2160
   ClientTop       =   3330
   ClientWidth     =   9480
   Icon            =   "frmMem.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   1290
   ScaleWidth      =   9480
   Begin VB.Timer Timer1 
      Left            =   3120
      Top             =   720
   End
   Begin ComctlLib.ProgressBar pbTotal 
      Height          =   255
      Left            =   900
      TabIndex        =   1
      Top             =   60
      Width           =   7035
      _ExtentX        =   12409
      _ExtentY        =   450
      _Version        =   327682
      Appearance      =   1
   End
   Begin ComctlLib.ProgressBar pbFree 
      Height          =   255
      Left            =   900
      TabIndex        =   4
      Top             =   360
      Width           =   7035
      _ExtentX        =   12409
      _ExtentY        =   450
      _Version        =   327682
      Appearance      =   1
   End
   Begin ComctlLib.ProgressBar pbVirtual 
      Height          =   255
      Left            =   900
      TabIndex        =   7
      Top             =   660
      Width           =   7035
      _ExtentX        =   12409
      _ExtentY        =   450
      _Version        =   327682
      Appearance      =   1
   End
   Begin ComctlLib.ProgressBar pbTotVirtual 
      Height          =   255
      Left            =   900
      TabIndex        =   10
      Top             =   960
      Width           =   7035
      _ExtentX        =   12409
      _ExtentY        =   450
      _Version        =   327682
      Appearance      =   1
   End
   Begin MemInfoCtl.MemoryInfo MemoryInfo1 
      Left            =   8460
      Top             =   300
      _ExtentX        =   847
      _ExtentY        =   847
      Free            =   3160
      Total           =   63504
      Virtual         =   1952562
      TotalVirtual    =   2044000
   End
   Begin VB.Label lblTotVirtual 
      AutoSize        =   -1  'True
      Caption         =   "0"
      Height          =   195
      Left            =   8040
      TabIndex        =   11
      Top             =   1020
      Width           =   90
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Tot Virtual"
      Height          =   195
      Left            =   60
      TabIndex        =   9
      Top             =   960
      Width           =   720
   End
   Begin VB.Label lblVirtual 
      AutoSize        =   -1  'True
      Caption         =   "0"
      Height          =   195
      Left            =   8040
      TabIndex        =   8
      Top             =   720
      Width           =   90
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Virtual"
      Height          =   195
      Left            =   360
      TabIndex        =   6
      Top             =   720
      Width           =   435
   End
   Begin VB.Label lblFree 
      AutoSize        =   -1  'True
      Caption         =   "0"
      Height          =   195
      Left            =   8040
      TabIndex        =   5
      Top             =   420
      Width           =   90
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Free"
      Height          =   195
      Left            =   480
      TabIndex        =   3
      Top             =   420
      Width           =   315
   End
   Begin VB.Label lblTotal 
      AutoSize        =   -1  'True
      Caption         =   "0"
      Height          =   195
      Left            =   8040
      TabIndex        =   2
      Top             =   120
      Width           =   90
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Total"
      Height          =   195
      Left            =   420
      TabIndex        =   0
      Top             =   120
      Width           =   360
   End
End
Attribute VB_Name = "frmMem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    Dim tot As Long

    tot = MemoryInfo1.TotalVirtual + 102400
    pbTotal.Max = tot
    pbFree.Max = tot
    pbVirtual.Max = tot
    pbTotVirtual.Max = tot
    
    Timer1.Interval = 500
End Sub

Private Sub Timer1_Timer()
    Const strKB$ = " Kbyte"
    
    MousePointer = vbHourglass
    With MemoryInfo1
        pbTotal.Value = .Total
        lblTotal.Caption = .Total & strKB
        pbFree.Value = .Free
        lblFree.Caption = .Free & strKB
        pbVirtual.Value = .Virtual
        lblVirtual.Caption = .Virtual & strKB
        pbTotVirtual.Value = .TotalVirtual
        lblTotVirtual.Caption = .TotalVirtual & strKB
    End With
    MousePointer = vbDefault
End Sub
