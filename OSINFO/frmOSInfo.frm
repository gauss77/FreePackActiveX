VERSION 5.00
Object = "{B1080A16-F27D-11D2-9939-000000000000}#1.0#0"; "OSInfo.ocx"
Begin VB.Form frmOSInfo 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "OS Info"
   ClientHeight    =   6435
   ClientLeft      =   1935
   ClientTop       =   1860
   ClientWidth     =   5925
   Icon            =   "frmOSInfo.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6435
   ScaleWidth      =   5925
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdRun 
      Caption         =   "&Run"
      Height          =   375
      Left            =   4560
      TabIndex        =   9
      Top             =   5820
      Width           =   1215
   End
   Begin VB.TextBox txtRun 
      Height          =   315
      Left            =   120
      TabIndex        =   8
      Text            =   "notepad.exe"
      Top             =   5880
      Width           =   3135
   End
   Begin VB.CommandButton Command1 
      Caption         =   "E&xit now"
      Height          =   375
      Left            =   4560
      TabIndex        =   6
      Top             =   4860
      Width           =   1215
   End
   Begin VB.ComboBox cbExit 
      Height          =   315
      Left            =   120
      TabIndex        =   5
      Top             =   4860
      Width           =   2475
   End
   Begin VB.CheckBox chkAppbar 
      Caption         =   "Disabled Start Bar (App Bar)"
      Height          =   495
      Left            =   120
      TabIndex        =   3
      Top             =   3720
      Width           =   5655
   End
   Begin VB.CheckBox chkDesktop 
      Caption         =   "Disabled Desktop Icons"
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   3180
      Width           =   5655
   End
   Begin VB.CheckBox chkSysKeys 
      Caption         =   "Disabled Sys Keys"
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   2640
      Width           =   5655
   End
   Begin OSInfoCtl.OSInfo OSInfo1 
      Left            =   3240
      Top             =   4680
      _ExtentX        =   847
      _ExtentY        =   847
      BuildNumber     =   2600
      DesktopHandle   =   65556
      MajorVersion    =   5
      MinorVersion    =   1
      PlatformID      =   "Windows NT"
      ProcessorType   =   586
      SoundCard       =   -1  'True
      TempFile        =   "C:\DOCUME~1\danilo\IMPOST~1\Temp\tmp30B.tmp"
      TempPath        =   "C:\DOCUME~1\danilo\IMPOST~1\Temp\"
      UserName        =   "danilo"
      WinPath         =   "C:\WINDOWS\"
      WinSysPath      =   "C:\WINDOWS\System32\"
      WinStart        =   22860982
      ComputerName    =   "XTOP"
      TitleColor      =   12632256
      Display         =   7
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Name, and path, of application to run ?"
      Height          =   195
      Left            =   120
      TabIndex        =   7
      Top             =   5580
      Width           =   2760
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      Index           =   2
      X1              =   60
      X2              =   5820
      Y1              =   5460
      Y2              =   5460
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Exit Window Type"
      Height          =   195
      Left            =   120
      TabIndex        =   4
      Top             =   4560
      Width           =   1290
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      Index           =   1
      X1              =   60
      X2              =   5820
      Y1              =   4440
      Y2              =   4440
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      Index           =   0
      X1              =   60
      X2              =   5820
      Y1              =   2460
      Y2              =   2460
   End
   Begin VB.Label lblInfo 
      BackColor       =   &H80000018&
      BorderStyle     =   1  'Fixed Single
      Height          =   2175
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   5775
   End
End
Attribute VB_Name = "frmOSInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cbExit_Change()
    OSInfo1.ExitType = cbExit.ListIndex
End Sub

Private Sub chkAppbar_Click()
    If chkAppbar.Value = vbChecked Then OSInfo1.StartBar = False Else OSInfo1.StartBar = True
End Sub

Private Sub chkDesktop_Click()
    If chkDesktop.Value = vbChecked Then OSInfo1.DesktopIcons = False Else OSInfo1.DesktopIcons = True
End Sub

Private Sub chkSysKeys_Click()
    If chkSysKeys.Value = vbChecked Then OSInfo1.SysKeysDisabled = True Else OSInfo1.SysKeysDisabled = False
End Sub

Private Sub cmdRun_Click()
    MousePointer = vbHourglass
    If Len(txtRun.Text) > 0 Then OSInfo1.ShellPlus txtRun.Text
    MousePointer = vbDefault
End Sub

Private Sub Command1_Click()
    MousePointer = vbHourglass
    If MsgBox("You sure exit to Windows now ?", vbQuestion + vbYesNo) = vbYes Then
        OSInfo1.ExitWindow
    End If
    MousePointer = vbDefault
End Sub

Private Sub Form_Load()
    Dim tmp As String
    
    With OSInfo1
        tmp = vbCrLf
        tmp = tmp & "Computer Name : " & .ComputerName & vbCrLf
        tmp = tmp & "User Logged : " & .UserName & vbCrLf
        tmp = tmp & "Processor Type : " & .ProcessorType & vbCrLf
        tmp = tmp & "Windows Type : " & .PlatformID & vbCrLf
        tmp = tmp & "Windows Version : " & .MajorVersion & "." & .MinorVersion & "." & .BuildNumber & vbCrLf
        tmp = tmp & "Sound Card : " & IIf(.SoundCard, "Yes", "No") & vbCrLf
        tmp = tmp & "Windows Path : " & .WinPath & vbCrLf
        tmp = tmp & "Windows System Path : " & .WinSysPath & vbCrLf
        tmp = tmp & "Temp Path : " & .TempPath & vbCrLf
    End With
    lblInfo.Caption = tmp
    
    cbExit.AddItem "LogOff"
    cbExit.AddItem "Shutdown"
    cbExit.AddItem "Reboot"
    cbExit.AddItem "Force"
    cbExit.ListIndex = 2
End Sub
