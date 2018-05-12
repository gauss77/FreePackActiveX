VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Registry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const HKEY_CLASSES_ROOT = &H80000000
Private Const HKEY_CURRENT_USER = &H80000001
Private Const HKEY_LOCAL_MACHINE = &H80000002
Private Const HKEY_USERS = &H80000003
Private Const HKEY_PERFORMANCE_DATA = &H80000004
Private Const HKEY_CURRENT_CONFIG = &H80000005
Private Const HKEY_DYN_DATA = &H80000006

Private Const REG_OPTION_NON_VOLATILE = 0&

Private Const KEY_QUERY_VALUE = &H1
Private Const KEY_SET_VALUE = &H2
Private Const KEY_CREATE_SUB_KEY = &H4
Private Const KEY_ENUMERATE_SUB_KEYS = &H8
Private Const KEY_NOTIFY = &H10
Private Const KEY_CREATE_LINK = &H20
Private Const KEY_ALL_ACCESS = &H3F

Private Const ERROR_SUCCESS = 0&
Private Const REG_SZ = 1&
Private Const REG_BINARY = 3&
Private Const REG_DWORD = 4&

Private Declare Function RegCreateKeyEx _
      Lib "advapi32.dll" Alias "RegCreateKeyExA" _
      (ByVal hKey As Long, _
      ByVal lpSubKey As String, _
      ByVal Reserved As Long, _
      ByVal lpClass As String, _
      ByVal dwOptions As Long, _
      ByVal samDesired As Long, _
      lpSecurityAttributes As Long, _
      phkResult As Long, _
      lpdwDisposition As Long) As Long

Private Declare Function RegDeleteKey _
      Lib "advapi32.dll" Alias "RegDeleteKeyA" _
      (ByVal hKey As Long, _
      ByVal lpSubKey As String) As Long

Private Declare Function RegOpenKeyEx _
      Lib "advapi32.dll" Alias "RegOpenKeyExA" _
      (ByVal hKey As Long, _
      ByVal lpSubKey As String, _
      ByVal ulOptions As Long, _
      ByVal samDesired As Long, _
      phkResult As Long) As Long

Private Declare Function RegCloseKey _
      Lib "advapi32.dll" _
      (ByVal hKey As Long) As Long

Private Declare Function RegQueryValueExByte _
      Lib "advapi32.dll" Alias "RegQueryValueExA" _
      (ByVal hKey As Long, _
      ByVal lpValueName As String, _
      ByVal lpReserved As Long, _
      lpType As Long, _
      lpData As Byte, _
      lpcbData As Long) As Long

Private Declare Function RegQueryValueExStr _
      Lib "advapi32.dll" Alias "RegQueryValueExA" _
      (ByVal hKey As Long, _
      ByVal lpValueName As String, _
      ByVal lpReserved As Long, _
      lpType As Long, _
      ByVal lpData As String, _
      lpcbData As Long) As Long

Private Declare Function RegQueryValueExLong _
      Lib "advapi32.dll" Alias "RegQueryValueExA" _
      (ByVal hKey As Long, _
      ByVal lpValueName As String, _
      ByVal lpReserved As Long, _
      lpType As Long, _
      lpData As Long, _
      lpcbData As Long) As Long

Private Declare Function RegSetValueExByte _
      Lib "advapi32.dll" Alias "RegSetValueExA" _
      (ByVal hKey As Long, _
      ByVal lpValueName As String, _
      ByVal Reserved As Long, _
      ByVal dwType As Long, _
      lpData As Byte, _
      ByVal cbData As Long) As Long

Private Declare Function RegSetValueExStr _
      Lib "advapi32.dll" Alias "RegSetValueExA" _
      (ByVal hKey As Long, _
      ByVal lpValueName As String, _
      ByVal Reserved As Long, _
      ByVal dwType As Long, _
      ByVal lpData As String, _
      ByVal cbData As Long) As Long

Private Declare Function RegSetValueExLong _
      Lib "advapi32.dll" Alias "RegSetValueExA" _
      (ByVal hKey As Long, _
      ByVal lpValueName As String, _
      ByVal Reserved As Long, _
      ByVal dwType As Long, _
      lpData As Long, _
      ByVal cbData As Long) As Long

Private Declare Function RegDeleteValue _
      Lib "advapi32.dll" Alias "RegDeleteValueA" _
      (ByVal hKey As Long, _
      ByVal lpValueName As String) As Long

Private Function TrovaBase(ByRef Nome As String) As Long
'Restituisce il codice relativo al ramo principale della chiave

   Dim Base As String

   Base = UCase$(Left$(Nome, InStr(Nome, "\") - 1))
   Nome = Mid$(Nome, InStr(Nome, "\") + 1)
   Select Case Base
   Case "HKEY_CLASSES_ROOT": TrovaBase = HKEY_CLASSES_ROOT
   Case "HKEY_CURRENT_USER": TrovaBase = HKEY_CURRENT_USER
   Case "HKEY_LOCAL_MACHINE": TrovaBase = HKEY_LOCAL_MACHINE
   Case "HKEY_USERS": TrovaBase = HKEY_USERS
   Case "HKEY_PERFORMANCE_DATA": TrovaBase = HKEY_PERFORMANCE_DATA
   Case "HKEY_CURRENT_CONFIG": TrovaBase = HKEY_CURRENT_CONFIG
   Case "HKEY_DYN_DATA": TrovaBase = HKEY_DYN_DATA
   Case Else: TrovaBase = &H88888888   'Valore non valido
   End Select

End Function

Friend Function CreaChiave(ByVal Nome As String) As Boolean
'Crea una nuova chiave (o la apre se già esiste) e imposta il suo codice

   Dim Chiave As Long
   Dim Base As Long
   Dim Ris As Long

   Base = TrovaBase(Nome)
   CreaChiave = (RegCreateKeyEx(Base, Nome, 0&, vbNullString, REG_OPTION_NON_VOLATILE, KEY_QUERY_VALUE, ByVal 0&, Chiave, Ris) = ERROR_SUCCESS)
   If CreaChiave Then CreaChiave = ChiudiChiave(Chiave)

End Function

Friend Function CancellaChiave(ByVal Nome As String) As Boolean
'Elimina una chiave

   Dim Base As Long

   Base = TrovaBase(Nome)
   CancellaChiave = (RegDeleteKey(Base, Nome) = ERROR_SUCCESS)

End Function

Private Function ApriChiave(ByVal Nome As String, ByRef hChiave As Long, ByVal Accesso As Long) As Boolean
'Apre una chiave e imposta il suo codice

   Dim Base As Long

   Base = TrovaBase(Nome)
   ApriChiave = (RegOpenKeyEx(Base, Nome, 0&, Accesso, hChiave) = ERROR_SUCCESS)

End Function

Private Function ChiudiChiave(ByVal hChiave As Long) As Boolean
'Chiude una chiave dato il codice

   ChiudiChiave = (RegCloseKey(hChiave) = ERROR_SUCCESS)

End Function

Friend Function LeggiChiaveStringa(ByVal Chiave As String, ByVal Nome As String, ByRef Valore As String, Optional ByVal Default As String = "") As Boolean
'Legge una stringa dalla chiave specificata, se non ci riesce restituisce il valore di dafault

   Dim hChiave As Long
   Dim Dimensione As Long
   Dim Tipo As Long

   LeggiChiaveStringa = False
   If ApriChiave(Chiave, hChiave, KEY_QUERY_VALUE) Then
      If (RegQueryValueExStr(hChiave, Nome, 0&, Tipo, ByVal 0&, Dimensione) = ERROR_SUCCESS) And (Tipo = REG_SZ) Then
         Valore = Space$(Dimensione)
         If RegQueryValueExStr(hChiave, Nome, 0&, ByVal 0&, ByVal Valore, Dimensione) = ERROR_SUCCESS Then
            Valore = Left$(Valore, InStr(Valore, Chr(0)) - 1)
            LeggiChiaveStringa = True
         End If
      End If
      ChiudiChiave hChiave
   End If
   If Not LeggiChiaveStringa Then Valore = Default

End Function

Friend Function LeggiChiaveBinario(ByVal Chiave As String, ByVal Nome As String, ByRef Valore() As Byte) As Boolean
'Legge un valore binario dalla chiave specificata, se non ci riesce restituisce il valore di dafault

   Dim hChiave As Long
   Dim Dimensione As Long
   Dim Tipo As Long

   LeggiChiaveBinario = False
   If ApriChiave(Chiave, hChiave, KEY_QUERY_VALUE) Then
      If (RegQueryValueExByte(hChiave, Nome, 0&, Tipo, ByVal 0&, Dimensione) = ERROR_SUCCESS) And (Tipo = REG_BINARY) Then
         ReDim Valore(0 To Dimensione - 1)
         LeggiChiaveBinario = (RegQueryValueExByte(hChiave, Nome, 0&, ByVal 0&, Valore(0), Dimensione) = ERROR_SUCCESS)
      End If
      ChiudiChiave hChiave
   End If
   
End Function

Friend Function LeggiChiaveNumero(ByVal Chiave As String, ByVal Nome As String, ByRef Valore As Long, Optional ByVal Default As Long = 0) As Boolean
'Legge un numero dalla chiave specificata, se non ci riesce restituisce il valore di dafault

   Dim hChiave As Long
   Dim Tipo As Long

   LeggiChiaveNumero = False
   If ApriChiave(Chiave, hChiave, KEY_QUERY_VALUE) Then
      LeggiChiaveNumero = ((RegQueryValueExLong(hChiave, Nome, 0&, Tipo, Valore, 4) = ERROR_SUCCESS) And (Tipo = REG_DWORD))
      ChiudiChiave hChiave
   End If
   If Not LeggiChiaveNumero Then Valore = Default

End Function

Friend Function ScriviChiaveStringa(ByVal Chiave As String, ByVal Nome As String, ByVal Valore As String, Optional ByVal Default As String = "") As Boolean
'Scrive una stringa nella chiave se il suo valore è diverso da quello di default, altrimenti cancella la stringa dalla chiave

   Dim hChiave As Long
   
   ScriviChiaveStringa = False
   If ApriChiave(Chiave, hChiave, KEY_SET_VALUE) Then
      If Valore <> Default Then
         ScriviChiaveStringa = (RegSetValueExStr(hChiave, Nome, 0&, REG_SZ, ByVal Valore, Len(Valore) + 1) = ERROR_SUCCESS)
      Else
         ScriviChiaveStringa = CancellaValore(hChiave, Nome)
      End If
      ChiudiChiave hChiave
   End If

End Function

Friend Function ScriviChiaveBinario(ByVal Chiave As String, ByVal Nome As String, ByRef Valore() As Byte) As Boolean
'Scrive un valore binario nella chiave

   Dim hChiave As Long
   Dim Dimensione As Long

   ScriviChiaveBinario = False
   If ApriChiave(Chiave, hChiave, KEY_SET_VALUE) Then
      Dimensione = UBound(Valore) - LBound(Valore) + 1
      ScriviChiaveBinario = (RegSetValueExByte(hChiave, Nome, 0&, REG_BINARY, Valore(0), Dimensione) = ERROR_SUCCESS)
      ChiudiChiave hChiave
   End If

End Function

Friend Function ScriviChiaveNumero(ByVal Chiave As String, ByVal Nome As String, ByVal Valore As Long, Optional ByVal Default As Long = 0) As Boolean
'Scrive un numero nella chiave se il suo valore è diverso da quello di default, altrimenti cancella il numero dalla chiave

   Dim hChiave As Long
   
   ScriviChiaveNumero = False
   If ApriChiave(Chiave, hChiave, KEY_SET_VALUE) Then
      If Valore <> Default Then
         ScriviChiaveNumero = (RegSetValueExLong(hChiave, Nome, 0&, REG_DWORD, Valore, 4) = ERROR_SUCCESS)
      Else
         ScriviChiaveNumero = CancellaValore(hChiave, Nome)
      End If
      ChiudiChiave hChiave
   End If
   
End Function

Private Function CancellaValore(ByVal hChiave As Long, ByVal Nome As String) As Boolean
'Cancella un valore (stringa o numero) da una chiave

   CancellaValore = (RegDeleteValue(hChiave, Nome) = ERROR_SUCCESS)

End Function


