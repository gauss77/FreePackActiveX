VERSION 5.00
Begin VB.UserControl DriveInfo 
   CanGetFocus     =   0   'False
   ClientHeight    =   480
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   480
   ClipControls    =   0   'False
   InvisibleAtRuntime=   -1  'True
   Picture         =   "DriveInfo.ctx":0000
   PropertyPages   =   "DriveInfo.ctx":0442
   ScaleHeight     =   480
   ScaleWidth      =   480
   ToolboxBitmap   =   "DriveInfo.ctx":045D
End
Attribute VB_Name = "DriveInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "The Drive Info Control enables you to retrieve information from drive, hard-disk and cd-rom units installed on system.\r\n"
Attribute VB_Ext_KEY = "PropPageWizardRun" ,"Yes"
Option Explicit

'Enumerazione usata per definire i tipi di formattazione
Public Enum FormatTypeConstants
    vbQuickFormat = 0
    vbCompleteFormat = 1
    vbBootFormat = 2
End Enum

'Valori predefiniti proprietà:
Private Const m_def_Drive = "C:\"
Private Const m_def_Type = ""
Private Const m_def_Size = 0
Private Const m_def_Free = 0
Private Const m_def_Serial = 0
Private Const m_def_Label = ""
Private Const m_def_FormatType = 0

'Variabili proprietà:
Private m_Drive As String
Private m_Type As String
Private m_Size As Long
Private m_Free As Long
Private m_Serial As Long
Private m_Label As String
Private m_FormatType As Integer

'Tipi per la dichiarazioni di varabili usate nelle API
Private Type OSVERSIONINFO
        dwOSVersionInfoSize As Long
        dwMajorVersion As Long
        dwMinorVersion As Long
        dwBuildNumber As Long
        dwPlatformId As Long
        szCSDVersion As String * 128
End Type

' costanti usate nelle funzioni API
Private Const RESOURCETYPE_DISK = &H1
Private Const VER_PLATFORM_WIN32_NT = 2

' procedurae api per ricavare le informazioni sull'unità
Private Declare Function GetDriveTypeA Lib "kernel32" (ByVal nDrive As String) As Long
Private Declare Function GetDiskFreeSpaceA Lib "kernel32" (ByVal lpRootPathName As String, lpSectorsPerCluster As Long, lpBytesPerSector As Long, lpNumberOfFreeClusters As Long, lpTtoalNumberOfClusters As Long) As Long
Private Declare Function GetVolumeInformation Lib "kernel32" Alias "GetVolumeInformationA" (ByVal lpRootPathName As String, ByVal lpVolumeNameBuffer As String, ByVal nVolumeNameSize As Long, lpVolumeSerialNumber As Long, lpMaximumComponentLength As Long, lpFileSystemFlags As Long, ByVal lpFileSystemNameBuffer As String, ByVal nFileSystemNameSize As Long) As Long
Private Declare Function SetVolumeLabel Lib "kernel32" Alias "SetVolumeLabelA" (ByVal lpRootPathName As String, ByVal lpVolumeName As String) As Long
Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Long
Private Declare Function SHFormatDrive Lib "shell32" (ByVal hwnd As Long, ByVal iDrive As Long, ByVal iCapacity As Long, ByVal iFormatType As Long) As Long

Public Sub About()
Attribute About.VB_Description = "Show about box."
Attribute About.VB_UserMemId = -552
    frmSplash.Show vbModal
End Sub

Public Property Get Drive() As String
Attribute Drive.VB_Description = "Returns or sets the drive to retrieve information or formatting."
Attribute Drive.VB_HelpID = 7
Attribute Drive.VB_ProcData.VB_Invoke_Property = "Drives"
    Drive = m_Drive
End Property

Public Property Let Drive(ByVal New_Drive As String)
    Const BACKSLASH$ = "\"
    Dim p As Integer
    m_Drive = Trim$(New_Drive)
    p = InStr(m_Drive, Space(1))
    If p > 0 Then m_Drive = Left$(m_Drive, p - 1)
    If Right$(m_Drive, 1) <> BACKSLASH Then m_Drive = m_Drive & BACKSLASH
    PropertyChanged "Drive"
End Property

Public Property Get DriveType() As String
Attribute DriveType.VB_Description = "Return the drive type."
Attribute DriveType.VB_HelpID = 8
    Dim Ret As Long
    Ret = GetDriveTypeA(m_Drive)
    m_Type = Choose(Ret + 1, "Unknow", "Invalid", "Floppy", "Hard-Disk", "Network", "CD-ROM", "RAM Disk", "DVD")
    DriveType = m_Type
End Property

Public Property Let DriveType(ByVal New_Type As String)
    If Ambient.UserMode = False Then Err.Raise 394
    If Ambient.UserMode Then Err.Raise 393
    m_Type = New_Type
    PropertyChanged "DriveType"
End Property

Public Property Get Size() As Long
Attribute Size.VB_Description = "Returns the total size of the drive."
Attribute Size.VB_HelpID = 13
    Call GetSize
    Size = m_Size
End Property

Public Property Let Size(ByVal New_Size As Long)
    If Ambient.UserMode = False Then Err.Raise 394
    If Ambient.UserMode Then Err.Raise 393
    m_Size = New_Size
    PropertyChanged "Size"
End Property

Public Property Get Free() As Long
Attribute Free.VB_Description = "Returns the free space available on the drive."
Attribute Free.VB_HelpID = 10
    Call GetSize
    Free = m_Free
End Property

Public Property Let Free(ByVal New_Free As Long)
    If Ambient.UserMode = False Then Err.Raise 394
    If Ambient.UserMode Then Err.Raise 393
    m_Free = New_Free
    PropertyChanged "Free"
End Property

Public Property Get Serial() As Long
Attribute Serial.VB_Description = "Returns the serial number of the drive."
Attribute Serial.VB_HelpID = 12
    Call GetVolumeInfo
    Serial = m_Serial
End Property

Public Property Let Serial(ByVal New_Serial As Long)
    If Ambient.UserMode = False Then Err.Raise 394
    If Ambient.UserMode Then Err.Raise 393
    m_Serial = New_Serial
    PropertyChanged "Serial"
End Property

Public Property Get Label() As String
Attribute Label.VB_Description = "Returns or sets the label name of drive."
Attribute Label.VB_HelpID = 11
Attribute Label.VB_ProcData.VB_Invoke_Property = "Drives"
    Call GetVolumeInfo
    Label = m_Label
End Property

Public Property Let Label(ByVal New_Label As String)
    m_Label = New_Label
    Call SetVolumeLabel(m_Drive, m_Label)
    PropertyChanged "Label"
End Property

Public Property Get FormatType() As FormatTypeConstants
Attribute FormatType.VB_Description = "Return or set the format type."
Attribute FormatType.VB_HelpID = 9
    FormatType = m_FormatType
End Property

Public Property Let FormatType(ByVal New_FormatType As FormatTypeConstants)
    m_FormatType = New_FormatType
    PropertyChanged "FormatType"
End Property

Public Function Start() As Long
Attribute Start.VB_Description = "Shows format standard dialog box window."
Attribute Start.VB_HelpID = 14
    Dim Ret As Long
    Dim drv As Integer
    
    Ret = -1
    drv = Asc(UCase$(Mid$(m_Drive, 1, 1)))
    If drv >= 65 And drv <= 91 Then
        Ret = SHFormatDrive(Parent.hwnd, drv - 65, 0, m_FormatType)
        If Not IsWindowNT Then
            Select Case Ret
                Case 6  ' no errors
                    Ret = 0
                Case 0  ' error in parameters
                    Ret = -1
            End Select
        End If
    End If
    Start = Ret
End Function

Private Sub UserControl_Initialize()
    #If SHAREWARE = 1 Then
        frmSplash.Show vbModal
    #End If
End Sub

'Inizializza le proprietà di UserControl
Private Sub UserControl_InitProperties()
    m_Drive = m_def_Drive
    m_Type = m_def_Type
    m_Size = m_def_Size
    m_Free = m_def_Free
    m_Serial = m_def_Serial
    m_Label = m_def_Label
    m_FormatType = m_def_FormatType
End Sub

'Carica i valori della proprietà dalla memoria
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    m_Drive = PropBag.ReadProperty("Drive", m_def_Drive)
    m_Type = PropBag.ReadProperty("DriveType", m_def_Type)
    m_Size = PropBag.ReadProperty("Size", m_def_Size)
    m_Free = PropBag.ReadProperty("Free", m_def_Free)
    m_Serial = PropBag.ReadProperty("Serial", m_def_Serial)
    m_Label = PropBag.ReadProperty("Label", m_def_Label)
    m_FormatType = PropBag.ReadProperty("FormatType", m_def_FormatType)
End Sub

Private Sub UserControl_Resize()
    On Local Error Resume Next
    Width = 480
    Height = 480
End Sub

'Scrive i valori della proprietà in memoria
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("Drive", m_Drive, m_def_Drive)
    Call PropBag.WriteProperty("DriveType", m_Type, m_def_Type)
    Call PropBag.WriteProperty("Size", m_Size, m_def_Size)
    Call PropBag.WriteProperty("Free", m_Free, m_def_Free)
    Call PropBag.WriteProperty("Serial", m_Serial, m_def_Serial)
    Call PropBag.WriteProperty("Label", m_Label, m_def_Label)
    Call PropBag.WriteProperty("FormatType", m_FormatType, m_def_FormatType)
End Sub

Private Sub GetVolumeInfo()
    Dim volmaxname As Long
    Dim volflag As Long
    Dim volsysname As String
    m_Label = String$(200, 0)
    volsysname = String$(200, 0)
    GetVolumeInformation m_Drive, m_Label, Len(m_Label), m_Serial, volmaxname, volflag, volsysname, Len(volsysname)
    If Mid$(m_Label, 1, 1) <> vbNullChar And InStr(m_Label, vbNullChar) > 0 Then
        m_Label = Left$(m_Label, InStr(m_Label, vbNullChar) - 1)
    ElseIf Mid$(m_Label, 1, 1) = vbNullChar Then
        m_Label = vbNullString
    End If
End Sub

Private Sub GetSize()
    Const m_def_Divisor = 1048576
    
    Dim sectorxcluster As Long
    Dim bytexcluster As Long
    Dim freecluster As Long
    Dim totalcluster As Long
    GetDiskFreeSpaceA m_Drive, sectorxcluster, bytexcluster, freecluster, totalcluster
    m_Free = ((freecluster * sectorxcluster) / m_def_Divisor) * bytexcluster
    m_Size = ((totalcluster * sectorxcluster) / m_def_Divisor) * bytexcluster
End Sub

Private Function IsWindowNT() As Boolean
    Dim OSINF As OSVERSIONINFO

    IsWindowNT = False
    With OSINF
        .dwOSVersionInfoSize = 148
        GetVersionEx OSINF
        If .dwPlatformId = VER_PLATFORM_WIN32_NT Then IsWindowNT = True
    End With
End Function

