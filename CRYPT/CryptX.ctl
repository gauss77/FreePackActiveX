VERSION 5.00
Begin VB.UserControl CryptX 
   CanGetFocus     =   0   'False
   ClientHeight    =   1500
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2280
   ClipControls    =   0   'False
   InvisibleAtRuntime=   -1  'True
   Picture         =   "CryptX.ctx":0000
   PropertyPages   =   "CryptX.ctx":08CA
   ScaleHeight     =   1500
   ScaleWidth      =   2280
   ToolboxBitmap   =   "CryptX.ctx":08DC
   Windowless      =   -1  'True
End
Attribute VB_Name = "CryptX"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "PropPageWizardRun" ,"Yes"
Option Explicit

Private bLimited As Boolean
Private Cry As New CryptoAPI

'Valori predefiniti proprietà:
Private Const m_def_SignatureType = PROV_RSA_FULL
Private Const m_def_GenerateErrors = False
Private Const m_def_Provider = 0
Private Const m_def_EncryptionType = CALG_RC4
Private Const m_def_HashingType = CALG_MD5
Private Const m_def_Password = vbNullString

'Variabili proprietà:
Private m_SignatureType As Long
Private m_GenerateErrors As Boolean
Private m_Provider As Long
Private m_EncryptionType As EncryptionAlgorithm
Private m_HashingType As HashingAlgorithm
Private m_Password As String

Public Enum HashingAlgorithm
    CRYPT_MD2 = CALG_MD2
    CRYPT_MD4 = CALG_MD4
    CRYPT_MD5 = CALG_MD5
    CRYPT_SHA = CALG_SHA
    'CRYPT_SHA_256 = CALG_SHA_256
    'CRYPT_SHA_384 = CALG_SHA_384
    'CRYPT_SHA_512 = CALG_SHA_512
    'CRYPT_SHA1 = CALG_SHA1
    'CRYPT_MAC = CALG_MAC
    'CRYPT_SSL3_SHAMD5 = CALG_SSL3_SHAMD5
    'CRYPT_HMAC = CALG_HMAC
    'CRYPT_TLS1PRF = CALG_TLS1PRF
    'CRYPT_HASH_REPLACE_OWF = CALG_HASH_REPLACE_OWF
End Enum

Public Enum EncryptionAlgorithm
    CRYPT_CALG_RC2 = CALG_RC2
    CRYPT_CALG_RC4 = CALG_RC4
    CRYPT_CALG_DES = CALG_DES
    CRYPT_CALG_3DES = CALG_3DES
    CRYPT_CALG_3DES_112 = CALG_3DES_112
    'CRYPT_CALG_AES_128 = CALG_AES_128
    'CRYPT_CALG_AES_192 = CALG_AES_192
    'CRYPT_CALG_AES_256 = CALG_AES_256
    'CRYPT_CALG_AES = CALG_AES
    'CRYPT_CALG_RC5 = CALG_RC5
    'CRYPT_CALG_DESX = CALG_DESX
    'CRYPT_CALG_SEAL = CALG_SEAL
    'CRYPT_CALG_SKIPJACK = CALG_SKIPJACK
    'CRYPT_CALG_TEK = CALG_TEK
    'CRYPT_CALG_CYLINK_MEK = CALG_CYLINK_MEK
End Enum

Public Enum EncryptionProvider
    CRYPT_MS_DEF_PROV = 0               ' "Microsoft Base Cryptographic Provider v1.0"
    CRYPT_MS_ENHANCED_PROV = 1          ' "Microsoft Enhanced Cryptographic Provider v1.0"
    CRYPT_MS_STRONG_PROV = 2            ' "Microsoft Strong Cryptographic Provider"
    'CRYPT_MS_DEF_RSA_SIG_PROV = 3       ' "Microsoft RSA Signature Cryptographic Provider"
    'CRYPT_MS_DEF_RSA_SCHANNEL_PROV = 4  ' "Microsoft RSA SChannel Cryptographic Provider"
    'CRYPT_MS_DEF_DSS_PROV = 5           ' "Microsoft Base DSS Cryptographic Provider"
    'CRYPT_MS_DEF_DSS_DH_PROV = 6        ' "Microsoft Base DSS and Diffie-Hellman Cryptographic Provider"
    'CRYPT_MS_ENH_DSS_DH_PROV = 7        ' "Microsoft Enhanced DSS and Diffie-Hellman Cryptographic Provider"
    'CRYPT_MS_DEF_DH_SCHANNEL_PROV = 8   ' "Microsoft DH SChannel Cryptographic Provider"
    'CRYPT_MS_SCARD_PROV = 9             ' "Microsoft Base Smart Card Crypto Provider"
    'CRYPT_MS_ENH_RSA_AES_PROV = 10      ' "Microsoft Enhanced RSA and AES Cryptographic Provider"
End Enum

Public Enum DigitalSignature
    CRYPT_PROV_RSA_FULL = PROV_RSA_FULL 'ok
    'CRYPT_PROV_RSA_SIG = PROV_RSA_SIG 'ok
    'CRYPT_PROV_DSS = PROV_DSS 'ok
    'CRYPT_PROV_FORTEZZA = PROV_FORTEZZA
    'CRYPT_PROV_MS_EXCHANGE = PROV_MS_EXCHANGE
    'CRYPT_PROV_SSL = PROV_SSL
    'CRYPT_PROV_RSA_SCHANNEL = PROV_RSA_SCHANNEL
    'CRYPT_PROV_DSS_DH = PROV_DSS_DH
    'CRYPT_PROV_EC_ECDSA_SIG = PROV_EC_ECDSA_SIG
    'CRYPT_PROV_EC_ECNRA_SIG = PROV_EC_ECNRA_SIG
    'CRYPT_PROV_EC_ECDSA_FULL = PROV_EC_ECDSA_FULL
    'CRYPT_PROV_EC_ECNRA_FULL = PROV_EC_ECNRA_FULL
    'CRYPT_PROV_DH_SCHANNEL = PROV_DH_SCHANNEL
    'CRYPT_PROV_SPYRUS_LYNKS = PROV_SPYRUS_LYNKS
    'CRYPT_PROV_RNG = PROV_RNG
    'CRYPT_PROV_INTEL_SEC = PROV_INTEL_SEC
    'CRYPT_PROV_REPLACE_OWF = PROV_REPLACE_OWF
    'CRYPT_PROV_RSA_AES = PROV_RSA_AES
End Enum

Public Sub About()
Attribute About.VB_Description = "Show About Box."
Attribute About.VB_UserMemId = -552
    frmAbout.Show vbModal
End Sub

'AVVISO: NON RIMUOVERE O MODIFICARE LE SEGUENTI RIGHE DI COMMENTO
'MemberInfo=14,0,0,0
Public Property Get Password() As String
Attribute Password.VB_Description = "Return or set a password for entrypt or decrypt string."
Attribute Password.VB_HelpID = 7
    Password = m_Password
End Property

Public Property Let Password(ByVal New_Password As String)
    m_Password = New_Password
    PropertyChanged "Password"
End Property

'AVVISO: NON RIMUOVERE O MODIFICARE LE SEGUENTI RIGHE DI COMMENTO
'MemberInfo=14
Public Function Encrypt(ByVal Data As String) As String
Attribute Encrypt.VB_Description = "Encrypt a string with MS-CryptoAPI algoritm."
Attribute Encrypt.VB_HelpID = 23
    Dim sProv As String
    
    Encrypt = vbNullString
    If bLimited Then Exit Function
    sProv = GetProviderName(m_Provider)
    With Cry
        .Errors = m_GenerateErrors
        .HashingType = m_HashingType
        .EncryptionType = m_EncryptionType
        .EncryptionCSPConnect sProv, m_SignatureType
        Encrypt = .EncryptData(Data, m_Password)
        .EncryptionCSPDisconnect
    End With
End Function

'AVVISO: NON RIMUOVERE O MODIFICARE LE SEGUENTI RIGHE DI COMMENTO
'MemberInfo=14
Public Function Decrypt(ByVal Data As String) As String
Attribute Decrypt.VB_Description = "Decrypt a string with MS-CryptoAPI algoritm."
Attribute Decrypt.VB_HelpID = 16
    Dim sProv As String
    
    Decrypt = vbNullString
    If bLimited Then Exit Function
    sProv = GetProviderName(m_Provider)
    With Cry
        .Errors = m_GenerateErrors
        .HashingType = m_HashingType
        .EncryptionType = m_EncryptionType
        .EncryptionCSPConnect sProv, m_SignatureType
        Decrypt = .DecryptData(Data, m_Password)
        .EncryptionCSPDisconnect
    End With
End Function

Private Function GetProviderName(ByVal idx As Integer) As String
    Dim sRet As String
    Select Case idx
        Case 1: sRet = MS_ENHANCED_PROV
        Case 2: sRet = MS_STRONG_PROV
        'Case 3: sRet = MS_DEF_RSA_SIG_PROV
        'Case 4: sRet = MS_DEF_RSA_SCHANNEL_PROV
        'Case 5: sRet = MS_DEF_DSS_PROV
        'Case 6: sRet = MS_DEF_DSS_DH_PROV
        'Case 7: sRet = MS_ENH_DSS_DH_PROV
        'Case 8: sRet = MS_DEF_DH_SCHANNEL_PROV
        'Case 9: sRet = MS_SCARD_PROV
        'Case 10: sRet = MS_ENH_RSA_AES_PROV
        Case Else
            sRet = MS_DEF_PROV
    End Select
    GetProviderName = sRet
End Function

Private Sub UserControl_Initialize()
    bLimited = False
    
    #If SHAREWARE = 1 Then
        frmAbout.Show vbModal
    #End If
End Sub

'Inizializza le proprietà di UserControl.
Private Sub UserControl_InitProperties()
    Width = 485
    Height = 485
    m_Password = m_def_Password
    m_EncryptionType = m_def_EncryptionType
    m_HashingType = m_def_HashingType
    m_Provider = m_def_Provider
    m_GenerateErrors = m_def_GenerateErrors
    m_SignatureType = m_def_SignatureType

    #If SHAREWARE = 1 Then
        bLimited = Ambient.UserMode
        If bLimited Then MsgBox "This is is a non-registered version of the Priore CryptX ActiveX, which should not be used in production environment. Reminder, some functions is not available in run-time mode.", vbInformation, App.Title
    #End If
End Sub

'Carica i valori delle proprietà dalla posizione di memorizzazione.
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    m_Password = PropBag.ReadProperty("Password", m_def_Password)
    m_EncryptionType = PropBag.ReadProperty("EncryptionType", m_def_EncryptionType)
    m_HashingType = PropBag.ReadProperty("HashingType", m_def_HashingType)
    m_Provider = PropBag.ReadProperty("Provider", m_def_Provider)
    m_GenerateErrors = PropBag.ReadProperty("GenerateErrors", m_def_GenerateErrors)
    m_SignatureType = PropBag.ReadProperty("SignatureType", m_def_SignatureType)
End Sub

Private Sub UserControl_Resize()
    Width = 485
    Height = 485
End Sub

'Scrive i valori delle proprietà nella posizione di memorizzazione.
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("Password", m_Password, m_def_Password)
    Call PropBag.WriteProperty("EncryptionType", m_EncryptionType, m_def_EncryptionType)
    Call PropBag.WriteProperty("HashingType", m_HashingType, m_def_HashingType)
    Call PropBag.WriteProperty("Provider", m_Provider, m_def_Provider)
    Call PropBag.WriteProperty("GenerateErrors", m_GenerateErrors, m_def_GenerateErrors)
    Call PropBag.WriteProperty("SignatureType", m_SignatureType, m_def_SignatureType)
End Sub

'AVVISO: NON RIMUOVERE O MODIFICARE LE SEGUENTI RIGHE DI COMMENTO
'MemberInfo=8,0,0,0
Public Property Get EncryptionType() As EncryptionAlgorithm
    EncryptionType = m_EncryptionType
End Property

Public Property Let EncryptionType(ByVal New_EncryptionType As EncryptionAlgorithm)
    m_EncryptionType = New_EncryptionType
    PropertyChanged "EncryptionType"
End Property

'AVVISO: NON RIMUOVERE O MODIFICARE LE SEGUENTI RIGHE DI COMMENTO
'MemberInfo=8,0,0,0
Public Property Get HashingType() As HashingAlgorithm
    HashingType = m_HashingType
End Property

Public Property Let HashingType(ByVal New_HashingType As HashingAlgorithm)
    m_HashingType = New_HashingType
    PropertyChanged "HashingType"
End Property

'AVVISO: NON RIMUOVERE O MODIFICARE LE SEGUENTI RIGHE DI COMMENTO
'MemberInfo=8,0,0,0
Public Property Get Provider() As EncryptionProvider
Attribute Provider.VB_Description = "Return or set the current crypto provider."
    Provider = m_Provider
End Property

Public Property Let Provider(ByVal New_Provider As EncryptionProvider)
    m_Provider = New_Provider
    PropertyChanged "Provider"
End Property

'AVVISO: NON RIMUOVERE O MODIFICARE LE SEGUENTI RIGHE DI COMMENTO
'MemberInfo=0,0,0,false
Public Property Get GenerateErrors() As Boolean
Attribute GenerateErrors.VB_Description = "Return or set generate the errors."
    GenerateErrors = m_GenerateErrors
End Property

Public Property Let GenerateErrors(ByVal New_GenerateErrors As Boolean)
    m_GenerateErrors = New_GenerateErrors
    PropertyChanged "GenerateErrors"
End Property

'AVVISO: NON RIMUOVERE O MODIFICARE LE SEGUENTI RIGHE DI COMMENTO
'MemberInfo=8,0,0,0
Public Property Get SignatureType() As DigitalSignature
Attribute SignatureType.VB_Description = "Return or set the digital signature."
    SignatureType = m_SignatureType
End Property

Public Property Let SignatureType(ByVal New_SignatureType As DigitalSignature)
    m_SignatureType = New_SignatureType
    PropertyChanged "SignatureType"
End Property

