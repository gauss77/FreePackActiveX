VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "File"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private iIndex As Integer
Private sFilename As String
Private sPath As String

Friend Property Get Index() As Integer
    Index = iIndex
End Property

Friend Property Let Index(ByVal New_Value As Integer)
    iIndex = New_Value
End Property

Public Property Get Filename() As String
Attribute Filename.VB_HelpID = 35
Attribute Filename.VB_UserMemId = 0
    Filename = sFilename
End Property

Friend Property Let Filename(New_Filename As String)
    sFilename = New_Filename
End Property

Public Property Get Path() As String
Attribute Path.VB_HelpID = 37
    Path = sPath
End Property

Friend Property Let Path(New_Path As String)
    sPath = New_Path
End Property

