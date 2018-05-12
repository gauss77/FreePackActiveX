VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FTPs"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mvarCollection As Collection

Public Property Get Count() As Long
Attribute Count.VB_Description = "Return number of all VirtualDir in FTP Site."
    Count = mvarCollection.Count
End Property

Friend Sub Remove(Index As Variant)
    mvarCollection.Remove Index
End Sub

Friend Sub Clear()
    Dim mvarFTP As FTP
    For Each mvarFTP In mvarCollection
        mvarCollection.Remove mvarFTP.Index
    Next
End Sub

Public Function Item(ByVal Index As Variant) As FTP
Attribute Item.VB_Description = "Return a single VirtualDir."
Attribute Item.VB_UserMemId = 0
    Set Item = mvarCollection.Item(Index)
End Function

Friend Function Add(ByVal Alias As String, Optional ByVal Path As String = vbNullString, Optional ByVal AccessRead As Integer = 1, Optional ByVal AccessWrite As Integer = 0) As FTP
    On Local Error Resume Next
    Dim mvarFTP As FTP
    If Len(Trim$(Alias)) > 0 Then
        Set mvarFTP = New FTP
        With mvarFTP
            .Index = mvarCollection.Count + 1
            .Alias = Alias
            .Path = Path
            .AccessRead = AccessRead
            .AccessWrite = AccessWrite
        End With
        mvarCollection.Add mvarFTP, Alias
    End If
    Set Add = mvarFTP
    Set mvarFTP = Nothing
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



