VERSION 5.00
Object = "{7FCFB161-49D5-4D74-B0DD-8D3523BC16E9}#3.0#0"; "PFingerCtl.ocx"
Begin VB.Form frmSample 
   Caption         =   "Priore FingerPrint ActiveX Sample"
   ClientHeight    =   3705
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5565
   Icon            =   "frmSample.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3705
   ScaleWidth      =   5565
   StartUpPosition =   2  'CenterScreen
   Begin PFingerPrintCtl.FingerPrint FingerPrint1 
      Height          =   3015
      Left            =   180
      Top             =   180
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   5318
   End
   Begin VB.Image Image1 
      Height          =   3000
      Left            =   3120
      Picture         =   "frmSample.frx":000C
      Top             =   180
      Width           =   2250
   End
   Begin VB.Label lblInfo 
      Caption         =   "Please, put the finger on the sensor"
      Height          =   255
      Left            =   180
      TabIndex        =   0
      Top             =   3300
      Width           =   5175
   End
End
Attribute VB_Name = "frmSample"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim iStep As Integer
Dim Template() As Byte
Dim pPic As StdPicture

Private Sub FingerPrint1_Errors(Number As Long, Description As String)
    lblInfo.Caption = Description   ' show errors or bad position/pressure
End Sub

Private Sub FingerPrint1_FingerIn()
    Select Case iStep
        ' start step
        Case 0
            ' to save finger image, use this property for
            ' to save fingerprint data in your database
            Template = FingerPrint1.TemplateDataB
            ' show new msg
            lblInfo.Caption = "Finger ok, remove finger on sensor"
            iStep = 1   ' next step
        ' last step
        Case 2
            ' show wait msg
            lblInfo.Caption = "Wait, verify finger..."
            ' verify current finger with previous saved
            ' note: not need the keep the finger in to reader
            If FingerPrint1.VerifyFingerB(Template) Then
                ' finger it's same
                lblInfo.Caption = "Ok, finger it's same."
            Else
                ' finger NOT same
                lblInfo.Caption = "Error, finger is not same!"
            End If
    End Select
End Sub

Private Sub FingerPrint1_FingerOut()
    Select Case iStep
        ' start step
        Case 0
            ' show msg
            lblInfo.Caption = "Please, put the finger on the sensor"
        ' middle step
        Case 1
            ' show msg
            lblInfo.Caption = "Ok, put again the finger on the sensor"
            ' next step
            iStep = 2
        ' last step
        Case 2
            ' restart
            iStep = 0
    End Select
End Sub

Private Sub Form_Load()
    iStep = 0                       ' start step
    FingerPrint1.Interval = 1000    ' activate interval
End Sub
