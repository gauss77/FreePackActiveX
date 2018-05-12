VERSION 5.00
Begin VB.UserControl FileINI 
   CanGetFocus     =   0   'False
   ClientHeight    =   480
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   480
   ClipControls    =   0   'False
   InvisibleAtRuntime=   -1  'True
   Picture         =   "ManagerINI.ctx":0000
   PropertyPages   =   "ManagerINI.ctx":030A
   ScaleHeight     =   480
   ScaleWidth      =   480
   ToolboxBitmap   =   "ManagerINI.ctx":0322
   Windowless      =   -1  'True
End
Attribute VB_Name = "FileINI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "PropPageWizardRun" ,"Yes"
Option Explicit

'Valori predefiniti proprietà:
Const m_def_Default = ""
Const m_def_Filename = ""
Const m_def_Key = ""
Const m_def_Value = ""
Const m_def_Section = ""

'Variabili proprietà:
Dim m_Default As String
Dim m_Filename As String
Dim m_Key As String
Dim m_Value As String
Dim m_Section As String

Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long

Private Function RemoveNulls(lpString As String) As String
    If Left$(lpString, 1) <> vbNullChar And InStr(lpString, vbNullChar) > 0 Then
        RemoveNulls = Left$(lpString, InStr(lpString, vbNullChar) - 1)
    ElseIf Left$(lpString, 1) = vbNullChar Then
        RemoveNulls = vbNullString
    Else
        RemoveNulls = lpString
    End If
End Function

Private Sub UserControl_Initialize()
    #If SHAREWARE = 1 Then
        frmSplash.Show vbModal
    #End If
End Sub

Private Sub UserControl_Resize()
    On Local Error Resume Next
    Width = 480
    Height = 480
End Sub

Public Sub About()
Attribute About.VB_Description = "Show about box."
Attribute About.VB_UserMemId = -552
    frmSplash.Show vbModal
End Sub


Public Property Get Default() As String
Attribute Default.VB_Description = "Return or set the default value."
Attribute Default.VB_ProcData.VB_Invoke_Property = "INI"
    Default = m_Default
End Property

Public Property Let Default(ByVal New_Default As String)
    m_Default = New_Default
    PropertyChanged "Default"
End Property

Public Property Get Filename() As String
Attribute Filename.VB_Description = "Return or set ini file name with extension."
Attribute Filename.VB_ProcData.VB_Invoke_Property = "INI"
    Filename = m_Filename
End Property

Public Property Let Filename(ByVal New_Filename As String)
    m_Filename = New_Filename
    PropertyChanged "Filename"
End Property

Public Property Get Key() As String
Attribute Key.VB_Description = "Return or set key name."
Attribute Key.VB_ProcData.VB_Invoke_Property = "INI"
    Key = m_Key
End Property

Public Property Let Key(ByVal New_Key As String)
    m_Key = New_Key
    PropertyChanged "Key"
End Property

Public Property Get Value() As Variant
Attribute Value.VB_Description = "Return or set key value."
Attribute Value.VB_ProcData.VB_Invoke_Property = "INI"
    Dim sRet As String
    sRet = String$(8192, 0)
    If Len(m_Section) > 0 And Len(m_Key) > 0 And Len(m_Filename) > 0 Then
        GetPrivateProfileString m_Section, m_Key, m_Default, sRet, Len(sRet), m_Filename
    End If
    m_Value = RemoveNulls(sRet)
    Value = m_Value
End Property

Public Property Let Value(ByVal New_Value As Variant)
    If Len(m_Section) > 0 And Len(m_Key) > 0 And Len(m_Filename) > 0 Then
        m_Value = New_Value
        WritePrivateProfileString m_Section, m_Key, CStr(m_Value), m_Filename
        PropertyChanged "Value"
    End If
End Property

Public Property Get Section() As String
Attribute Section.VB_Description = "Return or set section name."
Attribute Section.VB_ProcData.VB_Invoke_Property = "INI"
    Section = m_Section
End Property

Public Property Let Section(ByVal New_Section As String)
    m_Section = New_Section
    PropertyChanged "Section"
End Property

Public Sub WriteINI(lpSection As String, lpKey As String, lpValue As Variant, lpFileINI As String)
Attribute WriteINI.VB_Description = "Write a value in to file INI."
    With Me
        .Section = lpSection
        .Key = lpKey
        .Filename = lpFileINI
        .Value = lpValue
    End With
End Sub

Public Function ReadINI(lpSection As String, lpKey As String, lpDefault As String, lpFileINI As String) As String
Attribute ReadINI.VB_Description = "Read a string value from ini file."
    With Me
        .Section = lpSection
        .Key = lpKey
        .Default = lpDefault
        .Filename = lpFileINI
        ReadINI = .Value
    End With
End Function

Public Function ReadIntINI(lpSection As String, lpKey As String, lpDefault As Integer, lpFileINI As String) As Integer
Attribute ReadIntINI.VB_Description = "Read a numeric value form ini file."
    On Local Error Resume Next
    With Me
        .Section = lpSection
        .Key = lpKey
        .Default = CStr(lpDefault)
        .Filename = lpFileINI
        ReadIntINI = CInt(.Value)
    End With
End Function

'Inizializza le proprietà di UserControl
Private Sub UserControl_InitProperties()
    m_Default = m_def_Default
    m_Filename = m_def_Filename
    m_Key = m_def_Key
    m_Value = m_def_Value
    m_Section = m_def_Section
End Sub

'Carica i valori della proprietà dalla memoria
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    m_Default = PropBag.ReadProperty("Default", m_def_Default)
    m_Filename = PropBag.ReadProperty("Filename", m_def_Filename)
    m_Key = PropBag.ReadProperty("Key", m_def_Key)
    m_Value = PropBag.ReadProperty("Value", m_def_Value)
    m_Section = PropBag.ReadProperty("Section", m_def_Section)
End Sub

'Scrive i valori della proprietà in memoria
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("Default", m_Default, m_def_Default)
    Call PropBag.WriteProperty("Filename", m_Filename, m_def_Filename)
    Call PropBag.WriteProperty("Key", m_Key, m_def_Key)
    Call PropBag.WriteProperty("Value", m_Value, m_def_Value)
    Call PropBag.WriteProperty("Section", m_Section, m_def_Section)
End Sub

