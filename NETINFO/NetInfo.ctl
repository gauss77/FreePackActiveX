VERSION 5.00
Begin VB.UserControl NetInfo 
   CanGetFocus     =   0   'False
   ClientHeight    =   480
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   495
   InvisibleAtRuntime=   -1  'True
   Picture         =   "NetInfo.ctx":0000
   PropertyPages   =   "NetInfo.ctx":08CA
   ScaleHeight     =   480
   ScaleWidth      =   495
   ToolboxBitmap   =   "NetInfo.ctx":08DC
End
Attribute VB_Name = "NetInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "PropPageWizardRun" ,"Yes"
'Author:  Danilo Priore
'Email:   support@ prioregroup.com
'URL    : http://www.prioregroup.com
'
'This code is written and distributed under
'the GNU General Public License which means
'that its source code is freely-distributed
'and available to the general public.

Option Explicit

Public Enum InfoTypeClass
    COMPUTERS_LIST = 0
    GROUPS_LIST = 1
    USERS_LIST = 2
    SERVICES_LIST = 3
    PRINTERS_LIST = 4
    ALL_LIST = 255
End Enum

'Valori predefiniti proprietà:
Private Const m_def_InfoType = COMPUTERS_LIST
Private Const m_def_Domain = ""
Private Const m_def_ComputerName = ""
Private Const m_def_UserName = ""

'Variabili proprietà:
Private m_InfoType As Integer
Private m_Domain As String
Private m_ComputerName As String
Private m_UserName As String

' variabile array per l'elenco dei computers
Private Computers() As String

' costanti per le api
Private Const RESOURCETYPE_DISK = &H1
Private Const LOGON32_LOGON_INTERACTIVE = 2
Private Const LOGON32_PROVIDER_DEFAULT = 0


' procedurae api
Private Declare Function GetComputerNameA Lib "kernel32" (ByVal lpBuffer As String, nSize As Long) As Long
Private Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Private Declare Function WNetConnectionDialog Lib "mpr.dll" (ByVal hwnd As Long, ByVal dwType As Long) As Long
Private Declare Function WNetDisconnectDialog Lib "mpr.dll" (ByVal hwnd As Long, ByVal dwType As Long) As Long
Private Declare Function LogonUser Lib "advapi32.dll" Alias "LogonUserA" (ByVal lpszUsername As String, ByVal lpszDomain As String, ByVal lpszPassword As String, ByVal dwLogonType As Long, ByVal dwLogonProvider As Long, phToken As Long) As Long
Private Declare Function ImpersonateLoggedOnUser Lib "advapi32.dll" (ByVal hToken As Long) As Long
Private Declare Function RevertToSelf Lib "advapi32.dll" () As Long

Private Function RemoveNulls(lpString As String) As String
    RemoveNulls = Replace$(lpString, vbNullChar, vbNullString)
End Function

Public Sub About()
Attribute About.VB_Description = "Show about box."
Attribute About.VB_UserMemId = -552
    frmSplash.Show vbModal
End Sub

'AVVISO: NON RIMUOVERE O MODIFICARE LE SEGUENTI RIGHE DI COMMENTO
'MemberInfo=13,0,0,0
Public Property Get InfoType() As InfoTypeClass
Attribute InfoType.VB_Description = "Set the selected service type."
Attribute InfoType.VB_HelpID = 30
    InfoType = m_InfoType
End Property

Public Property Let InfoType(ByVal New_InfoType As InfoTypeClass)
    m_InfoType = New_InfoType
    PropertyChanged "InfoType"
End Property


Public Property Get ComputerName() As String
Attribute ComputerName.VB_Description = "Return current net computer name."
Attribute ComputerName.VB_HelpID = 7
    Dim cpname As String
    cpname = String$(100, 0)
    GetComputerNameA cpname, Len(cpname)
    m_ComputerName = RemoveNulls(cpname)
    ComputerName = m_ComputerName
End Property

Public Property Let ComputerName(ByVal New_ComputerName As String)
    If Ambient.UserMode = False Then Err.Raise 394
    If Ambient.UserMode Then Err.Raise 393
    m_ComputerName = New_ComputerName
    PropertyChanged "ComputerName"
End Property

Public Property Get Username() As String
Attribute Username.VB_Description = "Return the current net user connection."
Attribute Username.VB_HelpID = 31
    Dim strBuffer As String
    strBuffer = String$(255, 0)
    GetUserName strBuffer, Len(strBuffer)
    m_UserName = RemoveNulls(strBuffer)
    Username = m_UserName
End Property

Public Property Let Username(ByVal New_Username As String)
    If Ambient.UserMode = False Then Err.Raise 394
    If Ambient.UserMode Then Err.Raise 393
    m_UserName = New_Username
    PropertyChanged "UserName"
End Property

Public Property Get Domain() As String
Attribute Domain.VB_Description = "Return or set net domain name to retrieve informations."
Attribute Domain.VB_HelpID = 37
    Domain = m_Domain
End Property

Public Property Let Domain(ByVal New_Domain As String)
    m_Domain = New_Domain
    PropertyChanged "Domain"
End Property

Public Function GetInfo() As NetInfoCtl.Items
Attribute GetInfo.VB_Description = "Return a object with all item retrieves."
Attribute GetInfo.VB_HelpID = 16
    Dim strDomain As String
    Dim dom As Variant
    Dim list As Variant
    Dim srv As String
    Dim itms As NetInfoCtl.Items
    Dim itm As NetInfoCtl.Item
    
    If Len(m_Domain) > 0 Then
        'Use the WinNT Directory Services
        strDomain = "WinNT://" & m_Domain
        
        'Create the Domain object
        Set dom = GetObject(strDomain)
        
        'Search for Computers in the Domain
        Select Case m_InfoType
            Case COMPUTERS_LIST: srv = "Computer"
            Case GROUPS_LIST: srv = "Group"
            Case USERS_LIST: srv = "User"
            Case SERVICES_LIST: srv = "Service"
            Case PRINTERS_LIST: srv = "PrintQueue"
        End Select
        If m_InfoType <> ALL_LIST Then dom.Filter = Array(srv)
        
        On Local Error Resume Next
        Set itms = New NetInfoCtl.Items
        For Each list In dom
            With list
                Set itm = itms.Add(.Name)
                If Not itm Is Nothing Then itm.Description = .Description
            End With
        Next
    End If

    Set GetInfo = itms
    Set list = Nothing
    Set dom = Nothing
End Function

Public Function NetConnection()
Attribute NetConnection.VB_Description = "Show standard dialog net connection drive."
Attribute NetConnection.VB_HelpID = 35
    NetConnection = WNetConnectionDialog(Parent.hwnd, RESOURCETYPE_DISK)
End Function

Public Function NetDisconnection()
Attribute NetDisconnection.VB_Description = "Show  standard dialog net disconnection drive."
Attribute NetDisconnection.VB_HelpID = 36
    NetDisconnection = WNetDisconnectDialog(Parent.hwnd, RESOURCETYPE_DISK)
End Function

Public Sub Logon(ByVal strAdminUser As String, ByVal strAdminPassword As String, ByVal strAdminDomain As String)
Attribute Logon.VB_Description = "Logon user on network."
Attribute Logon.VB_HelpID = 38
    Dim lngTokenHandle, lngLogonType, lngLogonProvider As Long
    Dim blnResult As Boolean
    
    lngLogonType = LOGON32_LOGON_INTERACTIVE
    lngLogonProvider = LOGON32_PROVIDER_DEFAULT
    
    blnResult = RevertToSelf()
    blnResult = LogonUser(strAdminUser, strAdminDomain, strAdminPassword, lngLogonType, lngLogonProvider, lngTokenHandle)
    blnResult = ImpersonateLoggedOnUser(lngTokenHandle)
End Sub

Public Sub Logoff()
Attribute Logoff.VB_Description = "Logoff user on network."
Attribute Logoff.VB_HelpID = 39
    Dim blnResult As Boolean
    blnResult = RevertToSelf()
End Sub

Private Sub UserControl_Initialize()
    #If SHAREWARE = 1 Then
        frmSplash.Show vbModal
    #End If
End Sub

'Inizializza le proprietà di UserControl
Private Sub UserControl_InitProperties()
    m_Domain = m_def_Domain
    m_InfoType = m_def_InfoType
End Sub

'Carica i valori della proprietà dalla memoria
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    m_Domain = PropBag.ReadProperty("Domain", m_def_Domain)
    m_InfoType = PropBag.ReadProperty("InfoType", m_def_InfoType)
End Sub

Private Sub UserControl_Resize()
    On Local Error Resume Next
    Width = 480
    Height = 480
End Sub

'Scrive i valori della proprietà in memoria
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("Domain", m_Domain, m_def_Domain)
    Call PropBag.WriteProperty("InfoType", m_InfoType, m_def_InfoType)
End Sub

