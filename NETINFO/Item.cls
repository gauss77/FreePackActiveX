VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Item"
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

Private iIndex As Integer
Private sName As String
Private sDescription As String

Friend Property Get Index() As Integer
    Index = iIndex
End Property

Friend Property Let Index(ByVal New_Value As Integer)
    iIndex = New_Value
End Property

Public Property Get ItemName() As String
Attribute ItemName.VB_Description = "Return the item name."
Attribute ItemName.VB_HelpID = 32
Attribute ItemName.VB_UserMemId = 0
    ItemName = sName
End Property

Friend Property Let ItemName(New_name As String)
    sName = New_name
End Property

Public Property Get Description() As String
Attribute Description.VB_Description = "Return a description of item retrieves."
Attribute Description.VB_HelpID = 33
    Description = sDescription
End Property

Friend Property Let Description(New_Description As String)
    sDescription = New_Description
End Property
