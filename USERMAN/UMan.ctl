VERSION 5.00
Begin VB.UserControl UserManager 
   ClientHeight    =   750
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1005
   InvisibleAtRuntime=   -1  'True
   Picture         =   "UMan.ctx":0000
   PropertyPages   =   "UMan.ctx":08CA
   ScaleHeight     =   750
   ScaleWidth      =   1005
   ToolboxBitmap   =   "UMan.ctx":08DC
   Windowless      =   -1  'True
End
Attribute VB_Name = "UserManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Enum AccessEnum
    AccessRead = 0
    AccessWrite = 1
    AccessChange = 2
    FullAccess = 3
End Enum

Public Enum ErrorsEnum
    DOMAIN_COMPUTER_NOT_FOUND = vbObjectError + 68
    GROUP_USER_NOT_FOUND = vbObjectError + 76
End Enum

Private Const LOGON32_LOGON_INTERACTIVE = 2
Private Const LOGON32_PROVIDER_DEFAULT = 0

Private Const NTService$ = "WinNT://"
Private Const USERService$ = "user"
Private Const GROUPService$ = "group"

Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Private Declare Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function LogonUser Lib "advapi32.dll" Alias "LogonUserA" (ByVal lpszUsername As String, ByVal lpszDomain As String, ByVal lpszPassword As String, ByVal dwLogonType As Long, ByVal dwLogonProvider As Long, phToken As Long) As Long
Attribute LogonUser.VB_HelpID = 30
Private Declare Function ImpersonateLoggedOnUser Lib "advapi32.dll" (ByVal hToken As Long) As Long
Private Declare Function RevertToSelf Lib "advapi32.dll" () As Long

Public Sub About()
Attribute About.VB_Description = "Show about box."
Attribute About.VB_UserMemId = -552
    #If SHAREWARE = 1 Then
        frmSplash.Show vbModal
    #Else
        frmAbout.Show vbModal
    #End If
End Sub

Public Sub Add(ByVal ComputerName As String, ByVal Username As String, Optional ByVal Password As String = vbNullString, Optional ByVal Fullname As String = vbNullString, Optional ByVal Description As String = vbNullString, Optional ByVal Groupname As String = vbNullString)
Attribute Add.VB_Description = "Create a new domain user."
Attribute Add.VB_HelpID = 14
    Dim DNS As Variant
    Dim User As Variant
    Dim Group As Variant
    Dim Service As String
    
    On Local Error Resume Next
    Service = NTService & ComputerName
    
    Set DNS = GetObject(Service)
    If Err.Number <> 0 Then
        On Local Error GoTo 0
        Err.Raise DOMAIN_COMPUTER_NOT_FOUND
    End If
    
    If Not DNS Is Nothing Then
        Set User = DNS.Create(USERService, Username)
        If Not User Is Nothing Then
            With User
                .SetInfo
                .SetPassword Password
                .SetInfo
                .Description = Description
                .SetInfo
                .Fullname = Fullname
                .SetInfo
            End With
            
            If Len(Groupname) > 0 Then
                Set Group = DNS.GetBoject(GROUPService, Groupname)
                If Err.Number <> 0 Then
                    On Local Error GoTo 0
                    Err.Raise GROUP_USER_NOT_FOUND
                End If
                
                If Not Group Is Nothing Then Group.Add User.ADsPath
            End If
        End If
    End If
    
    Set DNS = Nothing
    Set User = Nothing
    Set Group = Nothing
End Sub

Public Sub ChangePassword(ByVal ComputerName As String, ByVal Username As String, ByVal OldPassword As String, ByVal NewPassword As String)
Attribute ChangePassword.VB_Description = "Change password user."
Attribute ChangePassword.VB_HelpID = 15
    Dim DNS As Variant
    Dim User As Variant
    Dim Service As String
    
    On Local Error Resume Next
    Service = NTService & ComputerName
    
    Set DNS = GetObject(Service)
    If Err.Number <> 0 Then
        On Local Error GoTo 0
        Err.Raise DOMAIN_COMPUTER_NOT_FOUND
    End If
    
    Set User = DNS.GetObject(USERService, Username)
    If Err.Number <> 0 Then
        On Local Error GoTo 0
        Err.Raise GROUP_USER_NOT_FOUND
    End If
    
    User.ChangePassword OldPassword, NewPassword
    User.SetInfo

    Set DNS = Nothing
    Set User = Nothing
End Sub

Public Sub SetPassword(ByVal ComputerName As String, ByVal Username As String, ByVal NewPassword As String)
Attribute SetPassword.VB_HelpID = 27
    Dim DNS As Variant
    Dim User As Variant
    Dim Service As String
    
    On Local Error Resume Next
    Service = NTService & ComputerName
    
    Set DNS = GetObject(Service)
    If Err.Number <> 0 Then
        On Local Error GoTo 0
        Err.Raise DOMAIN_COMPUTER_NOT_FOUND
    End If
    
    Set User = DNS.GetObject(USERService, Username)
    If Err.Number <> 0 Then
        On Local Error GoTo 0
        Err.Raise GROUP_USER_NOT_FOUND
    End If
    
    User.SetPassword NewPassword
    User.SetInfo

    Set DNS = Nothing
    Set User = Nothing
End Sub

Public Sub Remove(ByVal ComputerName As String, ByVal Username As String)
Attribute Remove.VB_Description = "Remove exits user."
Attribute Remove.VB_HelpID = 20
    Dim DNS As Variant
    Dim User As Variant
    Dim Service As String
    Dim Coll As Variant
    
    On Local Error Resume Next
    Service = NTService & ComputerName
    
    Set DNS = GetObject(Service)
    If Err.Number <> 0 Then
        On Local Error GoTo 0
        Err.Raise DOMAIN_COMPUTER_NOT_FOUND
    End If
    
    Set User = DNS.Create(USERService, Username)
    
    Set Coll = GetObject(User.Parent)
    Call Coll.Delete(USERService, User.Name)
    
    Set DNS = Nothing
    Set User = Nothing
    Set Coll = Nothing
End Sub

Public Sub SetUserAccess(ByVal Username As String, ByVal Path As String, ByVal Access As AccessEnum, Optional ByVal Revoche As Boolean = False)
Attribute SetUserAccess.VB_Description = "Set folder/file permissions."
Attribute SetUserAccess.VB_HelpID = 21
    Dim acc As String
    
    Select Case Access
        Case AccessRead: acc = "R"
        Case AccessWrite: acc = "W"
        Case AccessChange: acc = "C"
        Case FullAccess: acc = "F"
    End Select
    ShellPlus "cacls " & Path & " /E " & IIf(Revoche, "/R ", "/G ") & Username & ":" & acc
End Sub

Public Function GetUsers(ByVal ComputerName As String, ByVal Groupname As String) As Users
Attribute GetUsers.VB_Description = "Return a object users with list all users in group."
Attribute GetUsers.VB_HelpID = 16
Attribute GetUsers.VB_UserMemId = 0
    Dim DNS As Variant
    Dim Group As Variant
    Dim Usr As Variant
    Dim Lst As Users
    Dim Service As String
    
    On Local Error Resume Next
    Service = NTService & ComputerName
    
    Set DNS = GetObject(Service)
    If Err.Number <> 0 Then
        On Local Error GoTo 0
        Err.Raise DOMAIN_COMPUTER_NOT_FOUND
    End If
    
    Set Group = DNS.GetObject(GROUPService, Groupname)
    If Err.Number <> 0 Then
        On Local Error GoTo 0
        Err.Raise GROUP_USER_NOT_FOUND
    End If
    
    Set Lst = New Users
    For Each Usr In Group.Members
        Lst.Add Usr.Name, Usr.Description, Usr.Fullname
        If Err.Number <> 0 Then Lst.Add Usr.Name
    Next
    Set GetUsers = Lst
    
    Set DNS = Nothing
    Set Group = Nothing
    Set Lst = Nothing
End Function

Public Sub Logon(ByVal strAdminUser As String, ByVal strAdminPassword As String, ByVal strAdminDomain As String)
Attribute Logon.VB_Description = "Logon user on network."
Attribute Logon.VB_HelpID = 29
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
    Dim blnResult As Boolean
    blnResult = RevertToSelf()
End Sub

Private Function ShellPlus(sCommand As String) As Long
    Dim i As Integer
    Dim idTask As Long
    Dim hProc As Long
    
    idTask = Shell(sCommand, vbHide)
    hProc = OpenProcess(&H100000, False, idTask)
    WaitForSingleObject hProc, &HFFFFFFFF
    CloseHandle hProc
    ShellPlus = idTask
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
