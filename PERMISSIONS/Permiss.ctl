VERSION 5.00
Begin VB.UserControl ACLEdit 
   ClientHeight    =   750
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1005
   InvisibleAtRuntime=   -1  'True
   ScaleHeight     =   750
   ScaleWidth      =   1005
   ToolboxBitmap   =   "Permiss.ctx":0000
   Windowless      =   -1  'True
End
Attribute VB_Name = "ACLEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Enum ACLEnum
    DeniedAccess = 0
    ReadAccess = 1
    WriteAccess = 2
    ChangeAccess = 3
    FullAccess = 4
End Enum

Public Enum AccessEnum
    GrantAccess = 0
    RevocheAccess = 1
    Denied = 2
End Enum

Public Enum ACLModeEnum
    NewACL = 0
    EditACL = 1
End Enum

Private Const DENIED_ACC$ = "N"
Private Const READ_ACC$ = "R"
Private Const WRITE_ACC$ = "W"
Private Const CHANGE_ACC$ = "C"
Private Const FULL_ACC$ = "F"

Private Const EDIT_ACL$ = "/E"
Private Const GRANT_ACCESS$ = "/G"
Private Const REVOCHE_ACCESS$ = "/E /R"
Private Const DENIED_ACCESS$ = "/D"

Private Const CACLS$ = "CACLS "

Private WScr As Object
Private mBlank() As String

Private Declare Function GetComputerNameA Lib "kernel32" (ByVal lpBuffer As String, nSize As Long) As Long
Private Declare Function GetTempPathA Lib "kernel32" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
Private Declare Function GetTempFileName Lib "kernel32" Alias "GetTempFileNameA" (ByVal lpszPath As String, ByVal lpPrefixString As String, ByVal wUnique As Long, ByVal lpTempFileName As String) As Long
Private Declare Function DeleteFile Lib "kernel32" Alias "DeleteFileA" (ByVal lpFileName As String) As Long

Public Sub About()
Attribute About.VB_Description = "Show about box."
Attribute About.VB_UserMemId = -552
    frmSplash.Show vbModal
End Sub

Public Function GetACL(ByVal Filename As String) As UsersGroups
Attribute GetACL.VB_Description = "Set the proprety User with users or groups have permissions of file or folder."
    Dim ret As Long
    Dim tmpfile As String
    Dim ACLcommand As String
    
    On Local Error GoTo Err_GetACL
    If Len(Filename) > 0 Then
        If Len(Dir$(Filename, vbNormal Or vbDirectory)) > 0 Then
            tmpfile = GetTempFile
            ACLcommand = "cmd /c echo y| " & CACLS
            ACLcommand = ACLcommand & Chr(34) & Filename & Chr(34)
            ACLcommand = ACLcommand & " >" & tmpfile
            WScr.Run ACLcommand, 0, True
            Set GetACL = ParseFile(tmpfile, Filename)
        Else
            Error 53    ' Impossibile trovare il file
        End If
    Else
        Error 449   ' Argomento non facoltativo
    End If
    DeleteFile tmpfile
    Exit Function
    
Err_GetACL:
    DeleteFile tmpfile
    Set GetACL = Nothing
End Function

Public Function SetACL(ByVal Filename As String, ByVal UserGroup As String, ByVal ACL As ACLEnum, ByVal Access As AccessEnum, Optional ByVal ACLMode As ACLModeEnum = 0) As Long
Attribute SetACL.VB_Description = "Set single user or group permission for file or folder."
    Dim strACL As String
    Dim ACLcommand As String
    
    On Local Error GoTo Err_SetACL
    If Len(Filename) > 0 And Len(UserGroup) > 0 Then
        If Len(Dir$(Filename, vbNormal Or vbDirectory)) > 0 Then
            ACLcommand = "cmd /c echo y| " & CACLS
            ACLcommand = ACLcommand & Filename & " /T "
            strACL = GetComputerName & "\" & UserGroup & ":"
            strACL = strACL & GetACLCommand(ACL) & Space$(1)
            strACL = strACL & GetAccessCommand(Access) & Space$(1)
            strACL = strACL & GetACLModeCommand(ACLMode)
            WScr.Run ACLcommand & strACL, 0, True
            WScr.Run ACLcommand & GRANT_ACCESS & GetComputerName & "\" & "Administrators:" & FULL_ACC, 0, True
            WScr.Run ACLcommand & GRANT_ACCESS & GetComputerName & "\" & "SYSTEM:" & FULL_ACC, 0, True
        Else
            Error 53    ' Impossibile trovare il file
        End If
    Else
        Error 449   ' Argomento non facoltativo
    End If
    SetACL = 0
    Exit Function
    
Err_SetACL:
    SetACL = Err.Number
End Function

Private Function GetComputerName() As String
    Dim ret As Long
    Dim buffer As String
    buffer = String$(256, vbNullChar)
    ret = GetComputerNameA(buffer, Len(buffer))
    GetComputerName = Replace$(buffer, vbNullChar, vbNullString)
End Function

Private Function GetACLCommand(ByVal ACL As ACLEnum) As String
    Dim tmp As String
    Select Case ACL
        Case DeniedAccess: tmp = DENIED_ACC
        Case ReadAccess: tmp = READ_ACC
        Case WriteAccess: tmp = WRITE_ACC
        Case ChangeAccess: tmp = CHANGE_ACC
        Case FullAccess: tmp = FULL_ACC
    End Select
    GetACLCommand = tmp
End Function

Private Function GetAccessCommand(ByVal Access As AccessEnum) As String
    Dim tmp As String
    Select Case Access
        Case GrantAccess: tmp = GRANT_ACCESS
        Case RevocheAccess: tmp = REVOCHE_ACCESS
        Case Denied: tmp = DENIED_ACCESS
    End Select
    GetAccessCommand = tmp
End Function

Private Function GetACLModeCommand(ByVal ACLMode As ACLModeEnum) As String
    Dim tmp As String
    Select Case ACLMode
        Case NewACL: tmp = vbNullString
        Case EditACL: tmp = EDIT_ACL
    End Select
    GetACLModeCommand = tmp
End Function

Private Function GetTempPath() As String
    Dim tmppath As String
    tmppath = String$(256, 0)
    GetTempPathA Len(tmppath), tmppath
    GetTempPath = Replace$(tmppath, vbNullChar, vbNullString)
End Function

Private Function GetTempFile() As String
    Dim tmppath As String
    Dim nomefile As String
    tmppath = GetTempPath
    nomefile = String$(256, 0)
    GetTempFileName tmppath, "tmp", 0, nomefile
    GetTempFile = Replace$(nomefile, vbNullChar, vbNullString)
End Function

Private Function ParseFile(Filename As String, AccessFName As String) As UsersGroups
    Dim i As Integer
    Dim p As Integer
    Dim ff As Integer
    Dim ln As String
    Dim user As String
    Dim Final As String
    Dim Access As String
    Dim Usr As UsersGroups
    
    On Local Error GoTo Err_ParseFile:
    Set Usr = New UsersGroups
    Final = vbNullString
    ff = FreeFile
    Open Filename For Input As #ff
    Do While Not EOF(ff)
        Line Input #ff, ln
        'Debug.Print ln
        ln = Trim$(ln)
        If Len(ln) > 0 Then
            ln = Replace$(ln, AccessFName, vbNullString, , , vbTextCompare)
            For i = 0 To UBound(mBlank)
                ln = Replace$(ln, mBlank(i), vbNullString, , , vbTextCompare)
            Next
            ln = Trim(ln)
            p = InStr(ln, ":")
            If p > 1 Then
                user = Trim(Mid$(ln, 1, p - 1))
                If Len(user) > 0 Then
                    If InStr(1, Final, user & "|", vbTextCompare) = 0 Then
                        Select Case UCase$(Mid(ln, p + 1, 1))
                            Case "N": Access = "None"
                            Case "W": Access = "Write"
                            Case "C": Access = "Change"
                            Case "F": Access = "Full"
                            Case Else: Access = "Read"
                        End Select
                        Usr.Add user, Access
                        Final = Final & user & "|"
                    End If
                End If
            End If
        End If
    Loop
    Close ff
    Kill Filename
    Set ParseFile = Usr
    Set Usr = Nothing
    Exit Function
    
Err_ParseFile:
    Set Usr = Nothing
End Function

Private Sub UserControl_Initialize()
    Set WScr = CreateObject("WScript.Shell")
    
    ReDim mBlank(7)
    mBlank(0) = "(IO)"
    mBlank(1) = "(OI)"
    mBlank(2) = "(CI)"
    mBlank(3) = "(special access:)"
    mBlank(4) = "(accesso speciale:)"
    mBlank(5) = "NT AUTHORITY\"
    mBlank(6) = "BUILTIN\"
    mBlank(7) = GetComputerName & "\"
    
    #If SHAREWARE = 1 Then
        frmSplash.Show vbModal
    #End If
End Sub

Private Sub UserControl_Resize()
    On Local Error Resume Next
    Width = 480
    Height = 480
End Sub

Private Sub UserControl_Terminate()
    Set WScr = Nothing
End Sub
