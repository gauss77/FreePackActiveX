VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "UserGroup"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private iIndex As Integer
Private sName As String
Attribute sName.VB_VarUserMemId = 0
Attribute sName.VB_VarDescription = "Return user or group name."
Private sAccess As String
Attribute sAccess.VB_VarDescription = "Return permission."

Friend Property Get Index() As Integer
    Index = iIndex
End Property

Friend Property Let Index(New_Value As Integer)
    iIndex = New_Value
End Property

Public Property Get Name() As String
    Name = sName
End Property

Friend Property Let Name(New_Name As String)
    sName = New_Name
End Property

Public Property Get Access() As String
    Access = sAccess
End Property

Friend Property Let Access(New_Access As String)
    sAccess = New_Access
End Property

