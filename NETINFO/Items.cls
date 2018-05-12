VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Items"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Author:  Danilo Priore
'Email:   support@ prioregroup.com
'URL    : http://www.prioregroup.com
'
'This code is written and distributed under
'the GNU General Public License which means
'that its source code is freely-distributed
'and available to the general public.

Option Explicit

Private mvarCollection As Collection

Public Property Get Count() As Long
Attribute Count.VB_Description = "Return the numbersof items."
Attribute Count.VB_HelpID = 34
    Count = mvarCollection.Count
End Property

Friend Sub Remove(Index As Variant)
Attribute Remove.VB_Description = "Remove a single item."
    mvarCollection.Remove Index
End Sub

Friend Sub Clear()
Attribute Clear.VB_Description = "Remove all items."
    Dim sItem As NetInfoCtl.Item
    For Each sItem In mvarCollection
        mvarCollection.Remove sItem.Index
    Next
End Sub

Public Function Item(Index As Variant) As NetInfoCtl.Item
Attribute Item.VB_Description = "Return a object item."
Attribute Item.VB_UserMemId = 0
    Set Item = mvarCollection.Item(Index)
End Function

Friend Function Add(sName As String, Optional Description As String = vbNullString) As NetInfoCtl.Item
Attribute Add.VB_Description = "Append a new item in items list."
    On Local Error Resume Next
    Dim sItem As NetInfoCtl.Item
    If Len(Trim$(sName)) > 0 Then
        Set sItem = New NetInfoCtl.Item
        With sItem
            .Index = mvarCollection.Count + 1
            .ItemName = sName
            .Description = Description
        End With
        mvarCollection.Add sItem, sName
    End If
    Set Add = sItem
    Set sItem = Nothing
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





