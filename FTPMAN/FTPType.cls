VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FTP"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private iIndex As Integer
Private sAlias As String
Private sPath As String
Private iAccessRead As Integer
Private iAccessWrite As Integer

Friend Property Get Index() As Integer
    Index = iIndex
End Property

Friend Property Let Index(ByVal New_Value As Integer)
    iIndex = New_Value
End Property

Public Property Get Alias() As String
Attribute Alias.VB_Description = "Return alas of VirtualDir."
Attribute Alias.VB_UserMemId = 0
    Alias = sAlias
End Property

Friend Property Let Alias(New_Alias As String)
    sAlias = New_Alias
End Property

Public Property Get Path() As String
Attribute Path.VB_Description = "Return complete path of VirtualDir."
    Path = sPath
End Property

Friend Property Let Path(New_Path As String)
    sPath = New_Path
End Property

Public Property Get AccessRead() As Integer
Attribute AccessRead.VB_Description = "Return if VirtualDir can be read."
    AccessRead = iAccessRead
End Property

Friend Property Let AccessRead(ByVal New_Access As Integer)
    iAccessRead = New_Access
End Property

Public Property Get AccessWrite() As Integer
Attribute AccessWrite.VB_Description = "Return if VirtualDir can be write."
    AccessWrite = iAccessWrite
End Property

Friend Property Let AccessWrite(ByVal New_Access As Integer)
    iAccessWrite = New_Access
End Property

