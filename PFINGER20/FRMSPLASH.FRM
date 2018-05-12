VERSION 5.00
Begin VB.Form frmSplash 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   " "
   ClientHeight    =   4170
   ClientLeft      =   2805
   ClientTop       =   3540
   ClientWidth     =   6405
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   HelpContextID   =   20
   Icon            =   "FRMSPLASH.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4170
   ScaleWidth      =   6405
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   3885
      Left            =   120
      Picture         =   "FRMSPLASH.frx":000C
      ScaleHeight     =   3855
      ScaleWidth      =   1035
      TabIndex        =   8
      Top             =   180
      Width           =   1065
   End
   Begin VB.Timer tmrUnload 
      Interval        =   1000
      Left            =   2640
      Top             =   420
   End
   Begin VB.Label lblTime 
      Alignment       =   1  'Right Justify
      Caption         =   "Please, wait... 7 sec."
      Height          =   255
      Left            =   3780
      TabIndex        =   9
      Top             =   3780
      Width           =   2355
   End
   Begin VB.Label lblURL 
      AutoSize        =   -1  'True
      Caption         =   "http://www.prioregroup.com"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   195
      Left            =   1320
      MouseIcon       =   "FRMSPLASH.frx":1F8F
      MousePointer    =   99  'Custom
      TabIndex        =   7
      Top             =   3780
      WhatsThisHelpID =   20
      Width           =   2010
   End
   Begin VB.Label lblEMail 
      AutoSize        =   -1  'True
      Caption         =   "support@prioregroup.com"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   195
      Left            =   1320
      MouseIcon       =   "FRMSPLASH.frx":2299
      MousePointer    =   99  'Custom
      TabIndex        =   6
      Top             =   3540
      WhatsThisHelpID =   20
      Width           =   1830
   End
   Begin VB.Label lblAddress 
      AutoSize        =   -1  'True
      Caption         =   "PrioreGroup Software"
      Height          =   195
      Index           =   0
      Left            =   1320
      TabIndex        =   5
      Top             =   3300
      WhatsThisHelpID =   20
      Width           =   1515
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      Index           =   1
      X1              =   1320
      X2              =   6300
      Y1              =   3180
      Y2              =   3180
   End
   Begin VB.Label lblInfo 
      Caption         =   $"FRMSPLASH.frx":25A3
      Height          =   435
      Index           =   2
      Left            =   1320
      TabIndex        =   4
      Top             =   2580
      WhatsThisHelpID =   20
      Width           =   4935
   End
   Begin VB.Label lblInfo 
      Caption         =   "Complete pricing, shopping and licensing information can be found in this control's help file or in PrioreGroup Home Page."
      Height          =   435
      Index           =   1
      Left            =   1320
      TabIndex        =   3
      Top             =   2100
      WhatsThisHelpID =   20
      Width           =   4935
   End
   Begin VB.Label lblInfo 
      Caption         =   $"FRMSPLASH.frx":2639
      Height          =   915
      Index           =   0
      Left            =   1320
      TabIndex        =   2
      Top             =   1260
      WhatsThisHelpID =   20
      Width           =   4935
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      Index           =   0
      X1              =   1320
      X2              =   6300
      Y1              =   1080
      Y2              =   1080
   End
   Begin VB.Label lblCopyright 
      AutoSize        =   -1  'True
      Caption         =   "lblCopyright"
      Height          =   195
      Left            =   1320
      TabIndex        =   1
      Top             =   720
      WhatsThisHelpID =   20
      Width           =   810
   End
   Begin VB.Label lblAppTitle 
      AutoSize        =   -1  'True
      Caption         =   "lblAppTitle"
      Height          =   195
      Left            =   1320
      TabIndex        =   0
      Top             =   300
      WhatsThisHelpID =   20
      Width           =   735
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const STR_OCX$ = ".OCX"
Private Const SW_SHOW = 5

Dim iCount As Integer
Dim bClose As Boolean

Private Declare Function GetWindowWord Lib "user32" (ByVal hwnd As Long, ByVal nIndex As Long) As Integer
Private Declare Function GetModuleFileName Lib "kernel32" Alias "GetModuleFileNameA" (ByVal hModule As Long, ByVal lpFileName As String, ByVal nSize As Long) As Long
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Public Property Get IDEAmbient() As Boolean
    On Local Error Resume Next
    Debug.Print 1 / 0
    IDEAmbient = (Err.Number <> 0)
End Property

Public Property Get EXEName() As String
    Dim hParent As Long
    Dim hInstance As Long
    Dim sExe As String
    
    sExe = String$(255, vbNullChar)
    hInstance = GetWindowWord(Me.hwnd, -6)
    GetModuleFileName hInstance, sExe, 255
    sExe = Replace$(sExe, vbNullChar, vbNullString)
    EXEName = sExe
End Property

Private Sub Form_Load()
    Dim OCX As String
    
    iCount = 7
    bClose = False
    
    OCX = UCase$(App.EXEName)
    If Right$(OCX, 4) <> STR_OCX Then OCX = OCX & STR_OCX
    Me.Caption = "About " & App.Title
    lblAppTitle.Caption = App.Title & " - " & OCX
    lblCopyright.Caption = "Copyright © 1998/" & Format$(Now, "yyyy") & ", Danilo Priore"
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If Not bClose Then Cancel = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Not bClose Then Cancel = True
End Sub

Private Sub lblEMail_Click()
    ShellExecute Me.hwnd, "open", "mailto:" & lblEMail.Caption, "", "", SW_SHOW
End Sub

Private Sub lblURL_Click()
    ShellExecute Me.hwnd, "open", lblURL.Caption, "", "", SW_SHOW
End Sub

Private Sub tmrUnload_Timer()
    iCount = iCount - 1
    lblTime.Caption = "Please, wait... " & iCount & " sec."
    lblTime.Refresh
    If iCount > 0 Then Exit Sub
    bClose = True
    Unload Me
End Sub
