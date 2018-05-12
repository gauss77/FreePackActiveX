VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Files"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mvarCollection As Collection

Public Property Get Count() As Long
Attribute Count.VB_HelpID = 34
    Count = mvarCollection.Count
End Property

Friend Sub Remove(Index As Variant)
    mvarCollection.Remove Index
End Sub

Friend Sub Clear()
    Dim sFile As File
    For Each sFile In mvarCollection
        mvarCollection.Remove sFile.Index
    Next
End Sub

Public Function Item(Index As Variant) As File
Attribute Item.VB_HelpID = 33
Attribute Item.VB_UserMemId = 0
    Set Item = mvarCollection.Item(Index)
End Function

Friend Function Add(Filename As String, Path As String) As File
    On Local Error Resume Next
    Dim sFile As File
    If Len(Trim$(Filename)) > 0 And Len(Trim$(Path)) Then
        Set sFile = New File
        With sFile
            .Index = mvarCollection.Count + 1
            .Filename = Filename
            .Path = Path
        End With
        mvarCollection.Add sFile, Filename
    End If
    Set Add = sFile
    Set sFile = Nothing
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





