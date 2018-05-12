VERSION 5.00
Object = "*\ATaskMan.vbp"
Begin VB.Form frmTask 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Task manager"
   ClientHeight    =   1695
   ClientLeft      =   1770
   ClientTop       =   1560
   ClientWidth     =   4455
   Icon            =   "frmTask.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1695
   ScaleWidth      =   4455
   ShowInTaskbar   =   0   'False
   Begin VB.CheckBox chkService 
      Caption         =   "Service"
      Height          =   495
      Left            =   180
      TabIndex        =   3
      Top             =   1080
      Width           =   3795
   End
   Begin VB.CommandButton cmdApply 
      Caption         =   "&Apply"
      Height          =   375
      Left            =   2880
      TabIndex        =   2
      Top             =   420
      Width           =   1215
   End
   Begin VB.ComboBox cbPriority 
      Height          =   315
      Left            =   180
      TabIndex        =   1
      Top             =   480
      Width           =   2415
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      X1              =   60
      X2              =   4320
      Y1              =   960
      Y2              =   960
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Priority"
      Height          =   195
      Left            =   180
      TabIndex        =   0
      Top             =   180
      Width           =   465
   End
   Begin TaskManCtl.TaskManager TaskManager1 
      Left            =   2760
      Top             =   60
      _ExtentX        =   847
      _ExtentY        =   847
   End
End
Attribute VB_Name = "frmTask"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cbPriority_Change()
    cmdApply.Enabled = True
End Sub

Private Sub chkService_Click()
    If chkService.Value = vbChecked Then
        TaskManager1.Service = True
    Else
        TaskManager1.Service = False
    End If
End Sub

Private Sub cmdApply_Click()
    Select Case cbPriority.ListIndex
        Case 0
            TaskManager1.Priority = Normal
        Case 1
            TaskManager1.Priority = Idle
        Case 2
            TaskManager1.Priority = High
        Case 3
            TaskManager1.Priority = RealTime
    End Select
End Sub

Private Sub Form_Load()
    cbPriority.AddItem "Normal"
    cbPriority.AddItem "Idle"
    cbPriority.AddItem "Hight"
    cbPriority.AddItem "Realtime"
    cbPriority.ListIndex = 0
    
    cmdApply.Enabled = False
End Sub
