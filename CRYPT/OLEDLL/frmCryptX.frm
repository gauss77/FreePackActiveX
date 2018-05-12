VERSION 5.00
Begin VB.Form frmCryptX 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Priore Crypto WEB VB Sample"
   ClientHeight    =   4935
   ClientLeft      =   1740
   ClientTop       =   1575
   ClientWidth     =   4350
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4935
   ScaleWidth      =   4350
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command2 
      Caption         =   "Exit"
      Height          =   375
      Left            =   3000
      TabIndex        =   16
      Top             =   4500
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Caption         =   "Encrypt Data"
      Height          =   2115
      Left            =   60
      TabIndex        =   0
      Top             =   120
      Width           =   4215
      Begin VB.TextBox txtEncryptOut 
         Height          =   435
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   7
         Top             =   1500
         Width           =   3915
      End
      Begin VB.CommandButton cmdEncrypt 
         Caption         =   "Run"
         Height          =   375
         Left            =   2700
         TabIndex        =   5
         Top             =   780
         Width           =   1155
      End
      Begin VB.TextBox txtPwdEncrypt 
         Height          =   315
         Left            =   1320
         TabIndex        =   4
         Top             =   780
         Width           =   1215
      End
      Begin VB.TextBox txtEncrypt 
         Height          =   315
         Left            =   1320
         TabIndex        =   2
         Top             =   300
         Width           =   2535
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Output:"
         Height          =   195
         Left            =   120
         TabIndex        =   6
         Top             =   1260
         Width           =   525
      End
      Begin VB.Label Label2 
         Caption         =   "Password for encrypt:"
         Height          =   495
         Left            =   120
         TabIndex        =   3
         Top             =   720
         Width           =   1095
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Data to encrypt:"
         Height          =   195
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   1140
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Decrypt Data"
      Height          =   2115
      Left            =   60
      TabIndex        =   8
      Top             =   2280
      Width           =   4215
      Begin VB.TextBox txtDecOut 
         Height          =   435
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   15
         Top             =   1500
         Width           =   3915
      End
      Begin VB.TextBox txtDec 
         Height          =   315
         Left            =   1320
         TabIndex        =   10
         Top             =   300
         Width           =   2535
      End
      Begin VB.TextBox txtPwdDec 
         Height          =   315
         Left            =   1320
         TabIndex        =   12
         Top             =   780
         Width           =   1215
      End
      Begin VB.CommandButton cmdDec 
         Caption         =   "Run"
         Height          =   375
         Left            =   2760
         TabIndex        =   13
         Top             =   780
         Width           =   1155
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Output:"
         Height          =   195
         Left            =   120
         TabIndex        =   14
         Top             =   1260
         Width           =   525
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Data to decrypt:"
         Height          =   195
         Left            =   120
         TabIndex        =   9
         Top             =   360
         Width           =   1140
      End
      Begin VB.Label Label3 
         Caption         =   "Password for decrypt:"
         Height          =   495
         Left            =   120
         TabIndex        =   11
         Top             =   720
         Width           =   1095
      End
   End
End
Attribute VB_Name = "frmCryptX"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private CryptX1 As New CryptoWeb.Functions

Private Sub cmdDec_Click()
    With CryptX1
        .EncryptionType = CRYPT_3DES_112
        .GenerateErrors = False
        .HashingType = CRYPT_SHA
        .Provider = CRYPT_MS_ENHANCED_PROV
        .SignatureType = CRYPT_PROV_RSA_FULL
        .Password = txtPwdDec.Text
        txtDecOut.Text = .Decrypt(txtDec.Text)
    End With
End Sub

Private Sub cmdEncrypt_Click()
    With CryptX1
        .EncryptionType = CRYPT_3DES_112
        .GenerateErrors = False
        .HashingType = CRYPT_SHA
        .Provider = CRYPT_MS_ENHANCED_PROV
        .SignatureType = CRYPT_PROV_RSA_FULL
        .Password = txtPwdEncrypt.Text
        txtEncryptOut.Text = .Encrypt(txtEncrypt.Text)
    End With
End Sub

Private Sub Command2_Click()
    End
End Sub

Private Sub Form_Load()
    CryptX1.About
End Sub
