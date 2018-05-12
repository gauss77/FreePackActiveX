VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "UsersGroups"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mvarCollection As Collection

Public Property Get Count() As Long
    Count = mvarCollection.Count
End Property

Friend Sub Remove(Index As Variant)
    mvarCollection.Remove Index
End Sub

Public Function Item(Index As Variant) As UserGroup
Attribute Item.VB_UserMemId = 0
    Set Item = mvarCollection.Item(Index)
End Function

Friend Function Add(ByVal Name As String, Access As String) As UserGroup
    On Local Error Resume Next
    Dim mvarUser As UserGroup
    If Len(Trim$(Name)) > 0 Then
        Set mvarUser = New UserGroup
        With mvarUser
            .Index = mvarCollection.Count + 1
            .Name = Name
            .Access = Access
        End With
        mvarCollection.Add mvarUser, Name
    End If
    Set Add = mvarUser
    Set mvarUser = Nothing
End Function

Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    Set NewEnum = mvarCollection.[_NewEnum]
End Function

Private Sub Class_Initialize()
    Set mvarCollection = New Collection
End Sub

Private Sub Class_Terminate()
    Set mvarCollection = Nothing
End Sub


