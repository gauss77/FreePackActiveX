VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Users"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_HelpID = 8
Option Explicit

Private mvarCollection As Collection

Public Property Get Count() As Long
Attribute Count.VB_Description = "Return the numbers of users."
Attribute Count.VB_HelpID = 17
    Count = mvarCollection.Count
End Property

Friend Sub Remove(Index As Variant)
    mvarCollection.Remove Index
End Sub

Friend Sub Clear()
    Dim sUser As User
    For Each sUser In mvarCollection
        mvarCollection.Remove sUser.Index
    Next
End Sub

Public Function Item(Index As Variant) As User
Attribute Item.VB_Description = "Return a object user."
Attribute Item.VB_HelpID = 18
Attribute Item.VB_UserMemId = 0
    Set Item = mvarCollection.Item(Index)
End Function

Friend Function Add(Username As String, Optional Description As String = vbNullString, Optional Fullname As String = vbNullString) As User
    On Local Error Resume Next
    Dim sUser As User
    If Len(Trim$(Username)) > 0 Then
        Set sUser = New User
        With sUser
            .Index = mvarCollection.Count + 1
            .Username = Username
            .Description = Description
            .Fullname = Fullname
        End With
        mvarCollection.Add sUser, Username
    End If
    Set Add = sUser
    Set sUser = Nothing
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




