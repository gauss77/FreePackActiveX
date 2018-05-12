VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DSN"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Const ODBC_ADD_SYS_DSN = 4
Const ODBC_CONFIG_SYS_DSN = 5
Const ODBC_REMOVE_SYS_DSN = 6

Private Declare Function SQLConfigDataSource Lib "ODBCCP32.DLL" (ByVal hwndParent As Long, ByVal fRequest As Long, ByVal lpszDriver As String, ByVal lpszAttributes As String) As Long

Public Function AddDSN(ByVal sDSN As String, ByVal sDriver As String, ByVal sServer As String, ByVal sDatabase As String, ByVal sUsername As String, ByVal sPassword As String) As Boolean
    AddDSN = SQLConfigDataSource(&H0, ODBC_ADD_SYS_DSN, sDriver & vbNullChar, "DSN=" & sDSN & vbNullChar & "SERVER=" & sServer & vbNullChar & "DATABASE=" & sDatabase & vbNullChar & "DBQ=" & sDatabase & vbNullChar & "UID=" & sUsername & vbNullChar & "PWD=" & sPassword & vbNullChar)
End Function

Public Function ConfigDSN(ByVal sDSN As String, ByVal sDriver As String, ByVal sServer As String, ByVal sDatabase As String, ByVal sUsername As String, ByVal sPassword As String) As Boolean
    ConfigDSN = SQLConfigDataSource(&H0, ODBC_CONFIG_SYS_DSN, sDriver & vbNullChar, "DSN=" & sDSN & vbNullChar & "SERVER=" & sServer & vbNullChar & "DATABASE=" & sDatabase & vbNullChar & "DBQ=" & sDatabase & vbNullChar & "UID=" & sUsername & vbNullChar & "PWD=" & sPassword & vbNullChar)
End Function

Public Function RemoveDSN(ByVal sDSN As String, ByVal sDriver As String) As Boolean
    RemoveDSN = SQLConfigDataSource(&H0, ODBC_REMOVE_SYS_DSN, sDriver & vbNullChar, "DSN=" & sDSN & vbNullChar)
End Function
