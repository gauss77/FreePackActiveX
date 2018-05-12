VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CryptoAPI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public HashingType As Long
Public EncryptionType As Long
Public Errors  As Boolean

Private Const KEY_CONTAINER As String = "CryptX"
Private Const PP_NAME As Long = 4
Private Const PP_CONTAINER As Long = 6
Private Const CRYPT_NEWKEYSET As Long = 8

Private hCryptProv As Long

Private Declare Function CryptAcquireContext Lib "advapi32.dll" Alias "CryptAcquireContextA" (ByRef phProv As Long, ByVal pszContainer As String, ByVal pszProvider As String, ByVal dwProvType As Long, ByVal dwFlags As Long) As Long
Private Declare Function CryptGetProvParam Lib "advapi32.dll" (ByVal hProv As Long, ByVal dwParam As Long, ByRef pbData As Any, ByRef pdwDataLen As Long, ByVal dwFlags As Long) As Long
Private Declare Function CryptCreateHash Lib "advapi32.dll" (ByVal hProv As Long, ByVal Algid As Long, ByVal hKey As Long, ByVal dwFlags As Long, ByRef phHash As Long) As Long
Private Declare Function CryptHashData Lib "advapi32.dll" (ByVal hHash As Long, ByVal pbData As String, ByVal dwDataLen As Long, ByVal dwFlags As Long) As Long
Private Declare Function CryptDeriveKey Lib "advapi32.dll" (ByVal hProv As Long, ByVal Algid As Long, ByVal hBaseData As Long, ByVal dwFlags As Long, ByRef phKey As Long) As Long
Private Declare Function CryptDestroyHash Lib "advapi32.dll" (ByVal hHash As Long) As Long
Private Declare Function CryptEncrypt Lib "advapi32.dll" (ByVal hKey As Long, ByVal hHash As Long, ByVal Final As Long, ByVal dwFlags As Long, ByVal pbData As String, ByRef pdwDataLen As Long, ByVal dwBufLen As Long) As Long
Private Declare Function CryptDestroyKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Private Declare Function CryptReleaseContext Lib "advapi32.dll" (ByVal hProv As Long, ByVal dwFlags As Long) As Long
Private Declare Function CryptDecrypt Lib "advapi32.dll" (ByVal hKey As Long, ByVal hHash As Long, ByVal Final As Long, ByVal dwFlags As Long, ByVal pbData As String, ByRef pdwDataLen As Long) As Long

Friend Function EncryptionCSPConnect(sProvider As String, ByVal iSignatureType As Long) As Boolean
    'Get handle to CSP
    If CryptAcquireContext(hCryptProv, KEY_CONTAINER, sProvider, iSignatureType, CRYPT_NEWKEYSET) = 0 Then
        If CryptAcquireContext(hCryptProv, KEY_CONTAINER, sProvider, iSignatureType, 0) = 0 Then
            If Errors Then Err.Raise vbObject + 1, , "A CryptX container probably already exists."
            EncryptionCSPConnect = False
            Exit Function
        End If
    End If
    
    EncryptionCSPConnect = True
End Function

Friend Sub EncryptionCSPDisconnect()
    'Release provider handle
    If hCryptProv <> 0 Then
        CryptReleaseContext hCryptProv, 0
    End If
End Sub

Friend Function EncryptData(Data As String, Password As String) As String
    Dim sEncrypted As String
    sEncrypted = EncryptDecrypt(Data, Password, True)
    'Build encrypted string, starting with number of encryption iterations
    EncryptData = sEncrypted
End Function

Friend Function DecryptData(Data As String, Password As String) As String
    Dim sDecrypted As String
    sDecrypted = EncryptDecrypt(Data, Password, False)
    DecryptData = sDecrypted
End Function

Friend Function GetCSPDetails() As String
    Dim lLength As Long
    Dim yContainer() As Byte
    
    If hCryptProv = 0 Then
        GetCSPDetails = "Not connected to CSP"
        Exit Function
    End If
    
    'For developer info, show what the CSP & container name is
    lLength = 1000
    ReDim yContainer(lLength)
    If CryptGetProvParam(hCryptProv, PP_NAME, yContainer(0), lLength, 0) <> 0 Then
        GetCSPDetails = "Cryptographic Service Provider name: " & ByteToStr(yContainer, lLength)
    End If
    lLength = 1000
    ReDim yContainer(lLength)
    If CryptGetProvParam(hCryptProv, PP_CONTAINER, yContainer(0), lLength, 0) <> 0 Then
        GetCSPDetails = GetCSPDetails & vbCrLf & "Key Container name: " & ByteToStr(yContainer, lLength)
    End If
End Function

Private Function EncryptDecrypt(Data As String, Password As String, ByVal Encrypt As Boolean) As String
    Dim lBufLen As Long
    Dim lDataLen As Long
    Dim sTemp As String
    Dim hHash As Long
    Dim hKey As Long
    
    If hCryptProv = 0 Then
        If Errors Then Err.Raise vbObject + 2, , "Not connected to CSP."
        Exit Function
    End If
    
    '--------------------------------------------------------------------
    'The data will be encrypted with a session key derived from the
    'password.
    'The session key will be recreated when the data is decrypted
    'only if the password used to create the key is available.
    '--------------------------------------------------------------------

    'Create a hash object
    If CryptCreateHash(hCryptProv, HashingType, 0, 0, hHash) = 0 Then
        If Errors Then Err.Raise vbObject + 3, , "Error during create hashing."
    End If

    'Hash the password
    If CryptHashData(hHash, Password, Len(Password), 0) = 0 Then
        If Errors Then Err.Raise vbObject + 4, , "Error during hashing data."
    End If
    
    'Derive a session key from the hash object
    If CryptDeriveKey(hCryptProv, EncryptionType, hHash, 0, hKey) = 0 Then
        If Errors Then Err.Raise vbObject + 5, , "Error during encryption type."
    End If
    
    'Do the work
    sTemp = Data
    lDataLen = Len(Data)
    If Encrypt Then
        'Encrypt data
        lBufLen = lDataLen
        
        'Get required buffer size
        If CryptEncrypt(hKey, 0&, 1, 0, vbNullString, lBufLen, 0) <> 0 Then
            
            'Inizialize buffer
            If lBufLen < lDataLen Then lBufLen = lDataLen
            If lDataLen < lBufLen Then sTemp = sTemp & String$(lBufLen - lDataLen, vbNullChar)
            
            'Encrypt data
            If CryptEncrypt(hKey, 0&, 1, 0, sTemp, lDataLen, lBufLen) <> 0 Then
                ' trunc the string if the encrypted size is <> than the data size
                EncryptDecrypt = Mid$(sTemp, 1, lBufLen)
            Else
                If Errors Then Err.Raise vbObject + 6, , "Error during encryption data."
            End If
        Else
            If Errors Then Err.Raise vbObject + 6, , "Error during encryption data."
        End If
    Else
        'Decrypt data
        If CryptDecrypt(hKey, 0, 1, 0, sTemp, lDataLen) <> 0 Then
            EncryptDecrypt = Mid$(sTemp, 1, lDataLen)
        Else
            If Errors Then Err.Raise vbObject + 7, , "Error during decryption data."
        End If
    End If

    'Destroy session key
    If hKey <> 0 Then
        CryptDestroyKey hKey
    End If

    'Destroy hash object
    If hHash <> 0 Then
        CryptDestroyHash hHash
    End If
End Function

Private Function ByteToStr(ByRef ByteArray() As Byte, ByVal lLength As Long) As String
    Dim i As Long
    For i = LBound(ByteArray) To (LBound(ByteArray) + lLength)
        ByteToStr = ByteToStr & Chr$(ByteArray(i))
    Next
End Function

