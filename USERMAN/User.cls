VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "User"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_HelpID = 7
Option Explicit

Private iIndex As Integer
Private sUsername As String
Private sDescription As String
Private sFullName As String

Friend Property Get Index() As Integer
    Index = iIndex
End Property

Friend Property Let Index(ByVal New_Value As Integer)
    iIndex = New_Value
End Property

Public Property Get Username() As String
Attribute Username.VB_Description = "Return the username."
Attribute Username.VB_HelpID = 9
Attribute Username.VB_UserMemId = 0
    Username = sUsername
End Property

Friend Property Let Username(New_Username As String)
    sUsername = New_Username
End Property

Public Property Get Description() As String
Attribute Description.VB_Description = "Return the user description."
Attribute Description.VB_HelpID = 10
    Description = sDescription
End Property

Friend Property Let Description(New_Description As String)
    sDescription = New_Description
End Property

Public Property Get Fullname() As String
    Fullname = sFullName
End Property

Friend Property Let Fullname(New_Fullname As String)
    sFullName = New_Fullname
End Property

