VERSION 5.00
Begin VB.UserControl FTPVirtualDir 
   ClientHeight    =   750
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1005
   InvisibleAtRuntime=   -1  'True
   Picture         =   "FTP.ctx":0000
   PropertyPages   =   "FTP.ctx":08CA
   ScaleHeight     =   750
   ScaleWidth      =   1005
   ToolboxBitmap   =   "FTP.ctx":08DC
   Windowless      =   -1  'True
End
Attribute VB_Name = "FTPVirtualDir"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Const NTService$ = "IIS://"
Private Const MSFTPSvc$ = "/MSFTPSVC/"
Private Const FTPRoot$ = "/Root"
Private Const FTPService$ = "IISFtpVirtualDir"

Private iEnumerate As Integer

Public Sub About()
Attribute About.VB_Description = "Show about box."
Attribute About.VB_UserMemId = -552
    frmSplash.Show vbModal
End Sub

Public Sub Add(ByVal ComputerName As String, ByVal Alias As String, ByVal Path As String, Optional ByVal AccessRead As Boolean = True, Optional ByVal AccessWrite As Boolean = False)
Attribute Add.VB_Description = "Create a new FTP VirtualDir."
Attribute Add.VB_UserMemId = 0
    Dim DNS As Variant
    Dim FTPx As Variant
    Dim Service As String
    
    Service = NTService & ComputerName & MSFTPSvc & iEnumerate & FTPRoot
    Set DNS = GetObject(Service)
    If Not DNS Is Nothing Then
        Set FTPx = DNS.CreateObject(FTPService, Alias)
        If Not FTPx Is Nothing Then
            With FTPx
                .Path = Path
                .AccessRead = IIf(AccessRead, 1, 0)
                .AccessWrite = IIf(AccessWrite, 1, 0)
                .SetInfo
            End With
        End If
    End If
    
    Set FTPx = Nothing
    Set DNS = Nothing
End Sub

Public Sub Remove(ByVal ComputerName As String, ByVal Alias As String)
Attribute Remove.VB_Description = "Remove FTP VirtualDir."
    Dim DNS As Variant
    Dim Service As String
    
    Service = NTService & ComputerName & MSFTPSvc & iEnumerate & FTPRoot
    Set DNS = GetObject(Service)
    If Not DNS Is Nothing Then DNS.Delete FTPService, Alias
    
    Set DNS = Nothing
End Sub

Public Function GetFTPs(ByVal ComputerName As String) As FTPs
Attribute GetFTPs.VB_Description = "Return all VirtualDirs in FTP Site."
    Dim DNS As Variant
    Dim mvarFTP As FTPs
    Dim Service As String
    Dim VD As Variant

    Set mvarFTP = New FTPs
    Service = NTService & ComputerName & MSFTPSvc & iEnumerate & FTPRoot
    Set DNS = GetObject(Service)
    If Not DNS Is Nothing Then
        mvarFTP.Clear
        For Each VD In DNS
            With VD
                mvarFTP.Add .Name, .Path, .AccessRead, .AccessWrite
            End With
        Next
    End If
    Set GetFTPs = mvarFTP

    Set mvarFTP = Nothing
    Set VD = Nothing
    Set DNS = Nothing
End Function

Public Property Get Enumerate() As Integer
Attribute Enumerate.VB_Description = "Return or set the enumerate FTP service (1:Defaul FTP Service)"
    Enumerate = iEnumerate
End Property

Public Property Let Enumerate(ByVal New_Enumerate As Integer)
    iEnumerate = New_Enumerate
    PropertyChanged "Enumerate"
End Property

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

Private Sub UserControl_InitProperties()
    iEnumerate = 1
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    iEnumerate = PropBag.ReadProperty("Enumerate", 1)
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("Enumerate", iEnumerate, 1)
End Sub

