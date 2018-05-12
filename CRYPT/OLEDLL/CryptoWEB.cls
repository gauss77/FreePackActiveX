VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Functions"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Const ALG_CLASS_DATA_ENCRYPT As Long = 24576
Private Const ALG_CLASS_HASH As Long = 32768
Private Const ALG_TYPE_ANY As Long = 0
Private Const ALG_TYPE_STREAM As Long = 2048
Private Const ALG_SID_RC4 As Long = 1
Private Const ALG_SID_SEAL As Long = 2
Private Const ALG_SID_MD2 As Long = 1
Private Const ALG_SID_MD4 As Long = 2
Private Const ALG_SID_MD5 As Long = 3
Private Const ALG_SID_SHA As Long = 4
Private Const CALG_MD2 As Long = ((ALG_CLASS_HASH Or ALG_TYPE_ANY) Or ALG_SID_MD2)
Private Const CALG_MD4 As Long = ((ALG_CLASS_HASH Or ALG_TYPE_ANY) Or ALG_SID_MD4)
Private Const CALG_MD5 As Long = ((ALG_CLASS_HASH Or ALG_TYPE_ANY) Or ALG_SID_MD5)
Private Const CALG_SHA As Long = ((ALG_CLASS_HASH Or ALG_TYPE_ANY) Or ALG_SID_SHA)
Private Const CALG_RC4 As Long = ((ALG_CLASS_DATA_ENCRYPT Or ALG_TYPE_STREAM) Or ALG_SID_RC4)
Private Const CALG_SEAL As Long = ((ALG_CLASS_DATA_ENCRYPT Or ALG_TYPE_STREAM) Or ALG_SID_SEAL)

Private Const m_def_SignatureType = PROV_RSA_FULL
Private Const m_def_GenerateErrors = False
Private Const m_def_Provider = 0
Private Const m_def_EncryptionType = CALG_RC4
Private Const m_def_HashingType = CALG_MD5
Private Const m_def_Password = vbNullString

Private m_SignatureType As Long
Private m_GenerateErrors As Boolean
Private m_Provider As Long
Private m_EncryptionType As EncryptionAlgorithm
Private m_HashingType As HashingAlgorithm
Private m_Password As String

Private Cry As CryptoAPI

#If SHAREWARE = 1 Then
    Private Const TRIAL_DAY = 30
    Private Const REG_KEY$ = "HKEY_LOCAL_MACHINE\Software\Priore\ActiveX\CryptoWeb"
#End If

Public Enum HashingAlgorithm
    CRYPT_MD2 = CALG_MD2
    CRYPT_MD4 = CALG_MD4
    CRYPT_MD5 = CALG_MD5
    CRYPT_SHA = CALG_SHA
End Enum

Public Enum EncryptionAlgorithm
    CRYPT_RC2 = CALG_RC2
    CRYPT_RC4 = CALG_RC4
    CRYPT_DES = CALG_DES
    CRYPT_SEAL = CALG_SEAL
    CRYPT_3DES = CALG_3DES
    CRYPT_3DES_112 = CALG_3DES_112
End Enum

Public Enum EncryptionProvider
    CRYPT_MS_DEF_PROV = 0               ' "Microsoft Base Cryptographic Provider v1.0"
    CRYPT_MS_ENHANCED_PROV = 1          ' "Microsoft Enhanced Cryptographic Provider v1.0"
    CRYPT_MS_STRONG_PROV = 2            ' "Microsoft Strong Cryptographic Provider"
End Enum

Public Enum DigitalSignature
    CRYPT_PROV_RSA_FULL = PROV_RSA_FULL 'ok
End Enum

Public Sub About()
    #If SHAREWARE = 1 Then
        frmSplash.Show vbModal
    #Else
        frmAbout.Show vbModal
    #End If
End Sub

Public Property Get Password() As String
    Password = m_Password
End Property

Public Property Let Password(ByVal New_Password As String)
    m_Password = New_Password
End Property

Public Function Encrypt(ByVal sData As String, Optional ByVal bToHex As Boolean = False) As String
    Dim s As String
    Dim sProv As String
    
    s = vbNullString
    If Cry Is Nothing Then Exit Function
    sProv = GetProviderName(m_Provider)
    With Cry
        .Errors = m_GenerateErrors
        .HashingType = m_HashingType
        .EncryptionType = m_EncryptionType
        .EncryptionCSPConnect sProv, m_SignatureType
        s = .EncryptData(sData, m_Password)
        .EncryptionCSPDisconnect
    End With
    If bToHex Then s = Str2Hex(s)
    
    Encrypt = s
End Function

Public Function Decrypt(ByVal sData As String, Optional ByVal bFromHex As Boolean = False) As String
    Dim s As String
    Dim r As Integer
    Dim sProv As String
    
    s = vbNullString
    If Cry Is Nothing Then Exit Function
    If bFromHex Then sData = Hex2Str(sData)
    sProv = GetProviderName(m_Provider)
    With Cry
        .Errors = m_GenerateErrors
        .HashingType = m_HashingType
        .EncryptionType = m_EncryptionType
        .EncryptionCSPConnect sProv, m_SignatureType
        s = .DecryptData(sData, m_Password)
        .EncryptionCSPDisconnect
    End With
    
    Decrypt = s
End Function

Public Property Get EncryptionType() As EncryptionAlgorithm
    EncryptionType = m_EncryptionType
End Property

Public Property Let EncryptionType(ByVal New_EncryptionType As EncryptionAlgorithm)
    m_EncryptionType = New_EncryptionType
End Property

Public Property Get HashingType() As HashingAlgorithm
    HashingType = m_HashingType
End Property

Public Property Let HashingType(ByVal New_HashingType As HashingAlgorithm)
    m_HashingType = New_HashingType
End Property

Public Property Get Provider() As EncryptionProvider
    Provider = m_Provider
End Property

Public Property Let Provider(ByVal New_Provider As EncryptionProvider)
    m_Provider = New_Provider
End Property

Public Property Get GenerateErrors() As Boolean
    GenerateErrors = m_GenerateErrors
End Property

Public Property Let GenerateErrors(ByVal New_GenerateErrors As Boolean)
    m_GenerateErrors = New_GenerateErrors
End Property

Public Property Get SignatureType() As DigitalSignature
    SignatureType = m_SignatureType
End Property

Public Property Let SignatureType(ByVal New_SignatureType As DigitalSignature)
    m_SignatureType = New_SignatureType
End Property

Public Function Hex2Str(ByVal sData As String) As String
    Dim i As Integer
    Dim s As String
    s = vbNullString
    For i = 1 To Len(sData) Step 2
        s = s & Chr$("&h" & Mid$(sData, i, 2))
    Next
    Hex2Str = s
End Function

Public Function Str2Hex(ByVal sData As String) As String
    Dim i As Integer
    Dim h As String
    Dim s As String
    s = vbNullString
    For i = 1 To Len(sData)
        h = Hex$(Asc(Mid$(sData, i, 1)))
        If Len(h) < 2 Then h = "0" & h
        s = s & h
    Next
    Str2Hex = s
End Function

Private Sub Class_Initialize()
    
    m_Password = m_def_Password
    m_EncryptionType = m_def_EncryptionType
    m_HashingType = m_def_HashingType
    m_Provider = m_def_Provider
    m_GenerateErrors = m_def_GenerateErrors
    m_SignatureType = m_def_SignatureType
    
    Set Cry = New CryptoAPI
    
    #If SHAREWARE = 1 Then
        Dim i As Integer
        Dim sval As String
        Dim aval() As Byte
        Dim reg As Registry
        Set reg = New Registry
        If reg.LeggiChiaveBinario(REG_KEY, "Trial", aval) Then
            sval = vbNullString
            For i = 0 To UBound(aval)
                sval = sval & Chr$(aval(i))
            Next
            
            With Cry
                .Errors = False
                .HashingType = CRYPT_SHA
                .EncryptionType = CRYPT_3DES
                .EncryptionCSPConnect GetProviderName(CRYPT_MS_ENHANCED_PROV), CRYPT_PROV_RSA_FULL
                sval = .DecryptData(sval, Replace$("1çç1ç0ççç2ç3ç7ççççstç", "ç", vbNullString))
                .EncryptionCSPDisconnect
            End With
            
            If DateDiff("d", Now, CDate(sval)) > TRIAL_DAY Then
                Err.Raise vbObjectError + 513, , "This OLE/DLL Control is a SHAREWARE version, and with time limit period terminated."
                Set Cry = Nothing
                Exit Sub
            End If
        Else
            Err.Raise vbObjectError + 512, , "Incorrect Installation."
            Set Cry = Nothing
            Exit Sub
        End If
        Set reg = Nothing
    #End If
    
End Sub

Private Sub Class_Terminate()
    On Local Error Resume Next
    Set Cry = Nothing
End Sub

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

