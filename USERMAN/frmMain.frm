VERSION 5.00
Object = "{2B180F4B-EBCE-4807-9484-B5E2FA02EE92}#5.0#0"; "Ntusrman.ocx"
Begin VB.Form frmMain 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Example"
   ClientHeight    =   4470
   ClientLeft      =   1680
   ClientTop       =   1545
   ClientWidth     =   4125
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4470
   ScaleWidth      =   4125
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtComputer 
      Height          =   315
      Left            =   1440
      TabIndex        =   15
      Top             =   2280
      Width           =   2475
   End
   Begin VB.TextBox txtPassword 
      Height          =   315
      IMEMode         =   3  'DISABLE
      Left            =   1440
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   480
      Width           =   1215
   End
   Begin VB.TextBox txtFullname 
      Height          =   315
      Left            =   1440
      TabIndex        =   5
      Top             =   900
      Width           =   1215
   End
   Begin VB.CommandButton cmdRefresh 
      Caption         =   "Refresh"
      Height          =   375
      Left            =   2760
      TabIndex        =   13
      Top             =   3060
      Width           =   1215
   End
   Begin VB.TextBox txtGroupname 
      Height          =   315
      Left            =   1440
      TabIndex        =   9
      Top             =   1740
      Width           =   1215
   End
   Begin VB.CommandButton cmdRemove 
      Caption         =   "Remove"
      Height          =   375
      Left            =   2760
      TabIndex        =   14
      Top             =   3900
      Width           =   1215
   End
   Begin VB.ListBox lstUsers 
      Height          =   1230
      Left            =   120
      Sorted          =   -1  'True
      TabIndex        =   12
      Top             =   3060
      Width           =   2475
   End
   Begin VB.CommandButton cmdNew 
      Caption         =   "Add User"
      Height          =   375
      Left            =   2760
      TabIndex        =   10
      Top             =   60
      Width           =   1215
   End
   Begin VB.TextBox txtDescription 
      Height          =   315
      Left            =   1440
      TabIndex        =   7
      Top             =   1320
      Width           =   1215
   End
   Begin VB.TextBox txtUsername 
      Height          =   315
      Left            =   1440
      TabIndex        =   1
      Top             =   60
      Width           =   1215
   End
   Begin VB.Label lblComputer 
      Caption         =   "Computer Name:"
      Height          =   315
      Left            =   120
      TabIndex        =   16
      Top             =   2340
      Width           =   1215
   End
   Begin NTUserManager.UserManager UserManager1 
      Left            =   3000
      Top             =   1380
      _ExtentX        =   847
      _ExtentY        =   847
   End
   Begin VB.Label lblPassword 
      Caption         =   "Password:"
      Height          =   315
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   1215
   End
   Begin VB.Label lblFullname 
      Caption         =   "Full Name:"
      Height          =   315
      Left            =   120
      TabIndex        =   4
      Top             =   900
      Width           =   1215
   End
   Begin VB.Label lblList 
      Caption         =   "Users List:"
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   2760
      Width           =   1215
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      X1              =   60
      X2              =   3840
      Y1              =   2160
      Y2              =   2160
   End
   Begin VB.Label lblGroupname 
      Caption         =   "Group:"
      Height          =   315
      Left            =   120
      TabIndex        =   8
      Top             =   1740
      Width           =   1215
   End
   Begin VB.Label lblDescription 
      Caption         =   "Description:"
      Height          =   315
      Left            =   120
      TabIndex        =   6
      Top             =   1320
      Width           =   1215
   End
   Begin VB.Label lblUsername 
      Caption         =   "Username:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long

Private Sub RefreshList(ComputerName As String, Optional Groupname As String = "Users")
    Dim Usr As User
    Dim Lst As Users
    
    MousePointer = vbHourglass
    lstUsers.Clear
    Set Lst = UserManager1.GetUsers(ComputerName, Groupname)
    For Each Usr In Lst
        lstUsers.AddItem Usr.Username & IIf(Len(Usr.Fullname) > 0, " [" & Usr.Fullname & "]", vbNullString) & IIf(Len(Usr.Description) > 0, " (" & Usr.Description & ")", vbNullString)
    Next
    MousePointer = vbDefault
End Sub

Private Sub cmdNew_Click()
    If Len(txtUsername.Text) > 0 Then
        UserManager1.Add txtComputer.Text, txtUsername.Text, txtPassword.Text, txtFullname.Text, txtDescription.Text, txtGroupname.Text
    End If
End Sub

Private Sub cmdRefresh_Click()
    If Len(txtGroupname.Text) > 0 Then Call RefreshList(txtComputer.Text, txtGroupname.Text) Else Call RefreshList(txtComputer.Text)
End Sub

Private Sub cmdRemove_Click()
    MousePointer = vbHourglass
    UserManager1.Remove txtComputer.Text, lstUsers.Text
    MousePointer = vbDefault
End Sub

Private Sub Form_Load()
    Dim ret As Long
    Dim buffer As String
    
    MousePointer = vbHourglass
    cmdRemove.Enabled = False
    Show
    
    buffer = String$(255, vbNullChar)
    ret = GetComputerName(buffer, Len(buffer))
    buffer = Replace$(buffer, vbNullChar, vbNullString)
    txtComputer.Text = buffer
    
    Call RefreshList(buffer)
    MousePointer = vbDefault
End Sub

Private Sub lstUsers_Click()
    cmdRemove.Enabled = (Len(lstUsers.List(lstUsers.ListIndex)) > 0)
End Sub
