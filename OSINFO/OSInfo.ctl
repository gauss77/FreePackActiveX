VERSION 5.00
Begin VB.UserControl OSInfo 
   CanGetFocus     =   0   'False
   ClientHeight    =   480
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   480
   ClipControls    =   0   'False
   InvisibleAtRuntime=   -1  'True
   Picture         =   "OSInfo.ctx":0000
   PropertyPages   =   "OSInfo.ctx":030A
   ScaleHeight     =   480
   ScaleWidth      =   480
   ToolboxBitmap   =   "OSInfo.ctx":0341
   Windowless      =   -1  'True
End
Attribute VB_Name = "OSInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "PropPageWizardRun" ,"Yes"
Option Explicit

' tipi di enumerazione usati per il riavvio
Public Enum ExitTypeClass
    osLogOff = 0
    osShutDown = 1
    osReboot = 2
    osForce = 4
End Enum

Public Enum DisplaySettingsClass
    os640x400 = 0
    os640x480 = 1
    os800x600 = 2
    os1024x768 = 3
    os1152x864 = 4
    os1280x1024 = 5
    os1600x1200 = 6
    osUndefinited = 7
End Enum

' evento generato durante il cambio di una modalita grafica
Public Event DisplayChanged(NeedRestart As Boolean, DisplayError As Boolean)
Attribute DisplayChanged.VB_Description = "Event fired with to change screen property."

'Vostanti utilizzate per le funzioni API
Private Const HWND_TOPMOST = -1
Private Const HWND_NOTOPMOST = -2
Private Const SWP_DRAWFRAME = &H20
Private Const SWP_NOMOVE = &H2
Private Const SWP_NOSIZE = &H1
Private Const SWP_NORMAL = 1
Private Const GW_CHILD = 5
Private Const SW_SHOW = 5
Private Const SW_HIDE = 0
Private Const SW_RESTORE = 9
Private Const SWP_HIDEWINDOW = &H80
Private Const SWP_SHOWWINDOW = &H40
Private Const COLOR_ACTIVECAPTION = 2
Private Const STR_PROGMAN$ = "Progman"
Private Const STR_STARTBAR$ = "Shell_TrayWnd"
Private Const STR_PROGRAMMANAGER$ = "Program Manager"
Private Const VER_PLATFORM_WIN32_NT = 2
Private Const VER_PLATFORM_WIN32_WINDOWS = 1
Private Const VER_PLATFORM_WIN32s = 0
Private Const CCDEVICENAME = 32
Private Const CCFORMNAME = 32
Private Const DM_BITSPERPEL = &H40000
Private Const DM_PELSWIDTH = &H80000
Private Const DM_PELSHEIGHT = &H100000
Private Const CDS_UPDATEREGISTRY = &H1
Private Const CDS_TEST = &H4
Private Const DISP_CHANGE_SUCCESSFUL = 0
Private Const DISP_CHANGE_RESTART = 1

'Valori predefiniti proprietà:
Const m_def_OnTop = False
Private Const m_def_ExitType = 0
Private Const m_def_TitleColor = 0
Private Const m_def_CreateTempFile = False
Private Const m_def_TempFileSuffix = "tmp"
Private Const m_def_ComputerName = ""
Private Const m_def_BuildNumber = 0
Private Const m_def_DesktopHandle = 0
Private Const m_def_MajorVersion = 0
Private Const m_def_MinorVersion = 0
Private Const m_def_PlatformID = ""
Private Const m_def_ProcessorType = 0
Private Const m_def_SoundCard = False
Private Const m_def_TempFile = ""
Private Const m_def_TempPath = ""
Private Const m_def_UserName = ""
Private Const m_def_WinPath = ""
Private Const m_def_WinSysPath = ""
Private Const m_def_WinStart = 0
Private Const m_def_SysKeysDisabled = False

'Strutture per il riperimento delle informazioni tramite le API
Private Type OSVERSIONINFO
        dwOSVersionInfoSize As Long
        dwMajorVersion As Long
        dwMinorVersion As Long
        dwBuildNumber As Long
        dwPlatformID As Long
        szCSDVersion As String * 128
End Type

Private Type SYSTEM_INFO
        dwOemID As Long
        dwPageSize As Long
        lpMinimumApplicationAddress As Long
        lpMaximumApplicationAddress As Long
        dwActiveProcessorMask As Long
        dwNumberOrfProcessors As Long
        dwProcessorType As Long
        dwAllocationGranularity As Long
        dwReserved As Long
End Type

Private Type DEVMODE
        dmDeviceName As String * CCDEVICENAME
        dmSpecVersion As Integer
        dmDriverVersion As Integer
        dmSize As Integer
        dmDriverExtra As Integer
        dmFields As Long
        dmOrientation As Integer
        dmPaperSize As Integer
        dmPaperLength As Integer
        dmPaperWidth As Integer
        dmScale As Integer
        dmCopies As Integer
        dmDefaultSource As Integer
        dmPrintQuality As Integer
        dmColor As Integer
        dmDuplex As Integer
        dmYResolution As Integer
        dmTTOption As Integer
        dmCollate As Integer
        dmFormName As String * CCFORMNAME
        dmUnusedPadding As Integer
        dmBitsPerPel As Long
        dmPelsWidth As Long
        dmPelsHeight As Long
        dmDisplayFlags As Long
        dmDisplayFrequency As Long
End Type

'Variabili proprietà:
Dim m_OnTop As Boolean
Private m_Display As DisplaySettingsClass
Private m_ExitType As Long
Private m_TitleColor As OLE_COLOR
Private m_CreateTempFile As Boolean
Private m_TempFileSuffix As String
Private m_ComputerName As String
Private m_BuildNumber As Long
Private m_DesktopHandle As Long
Private m_MajorVersion As Long
Private m_MinorVersion As Long
Private m_PlatformID As String
Private m_ProcessorType As Long
Private m_SoundCard As Boolean
Private m_TempFile As String
Private m_TempPath As String
Private m_UserName As String
Private m_WinPath As String
Private m_WinSysPath As String
Private m_WinStart As Long
Private m_SysKeysDisabled As Boolean

' variabili private varie
Private lWinID As Long

'Funzioni api per ricavare le informazioni sul sitema
Private Declare Function ChangeDisplaySettings Lib "user32" Alias "ChangeDisplaySettingsA" (lpDevMode As Any, ByVal dwFlags As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function DeleteFile Lib "kernel32" Alias "DeleteFileA" (ByVal lpFileName As String) As Long
Private Declare Function EnableWindow Lib "user32" (ByVal hwnd As Long, ByVal fEnable As Long) As Long
Private Declare Function EnumDisplaySettings Lib "user32" Alias "EnumDisplaySettingsA" (ByVal lpszDeviceName As Long, ByVal iModeNum As Long, lpDevMode As Any) As Boolean
Private Declare Function ExitWindowsEx Lib "user32" (ByVal uFlags As Long, ByVal dwReserved As Long) As Long
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Long
Private Declare Function GetComputerNameA Lib "kernel32" (ByVal lpBuffer As String, nSize As Long) As Long
Private Declare Function GetDesktopWindow Lib "user32" () As Long
Private Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
Private Declare Function GetTempFileName Lib "kernel32" Alias "GetTempFileNameA" (ByVal lpszPath As String, ByVal lpPrefixString As String, ByVal wUnique As Long, ByVal lpTempFileName As String) As Long
Private Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Private Declare Function GetWindow Lib "user32" (ByVal hwnd As Long, ByVal wCmd As Long) As Long
Private Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Private Declare Function GetTickCount Lib "kernel32" () As Long
Private Declare Function GetSystemDirectory Lib "kernel32" Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Private Declare Sub GetSystemInfo Lib "kernel32" (lpSystemInfo As SYSTEM_INFO)
Private Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long
Private Declare Function IsWindowEnabled Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function IsWindowVisible Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function waveOutGetNumDevs Lib "winmm.dll" () As Long
Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Private Declare Function SetSysColors Lib "user32" (ByVal nChanges As Long, lpSysColor As Long, lpColorValues As Long) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, ByVal lpvParam As Any, ByVal fuWinIni As Long) As Long
Private Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Private Declare Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long

Private Sub GetInfo()
    Dim OSINF As OSVERSIONINFO

    With OSINF
        .dwOSVersionInfoSize = 148
        GetVersionEx OSINF
        m_MajorVersion = .dwMajorVersion
        m_MinorVersion = .dwMinorVersion
        m_BuildNumber = .dwBuildNumber
        lWinID = .dwPlatformID
        If lWinID = 1 Then
            If .dwMinorVersion = 0 Then ' Win95
                m_PlatformID = "Windows 95"
            ElseIf .dwMinorVersion = 10 Then ' Win98
                m_PlatformID = "Windows 98"
            Else
                m_PlatformID = "Undefinited"
            End If
        ElseIf lWinID = 2 Then
            If .dwMajorVersion = 3 Then ' Win3.51
                m_PlatformID = "Windows NT 3.51"
            ElseIf .dwMajorVersion = 4 Then ' Win4
                m_PlatformID = "Windows NT 4.0"
            ElseIf .dwMajorVersion = 5 Then ' Win2000
                m_PlatformID = "Windows 2000"
            Else
                m_PlatformID = "Undefinited"
            End If
        Else
            m_PlatformID = "Undefinited"
        End If
    End With
End Sub

Private Function RemoveNulls(lpString As String) As String
    RemoveNulls = Replace$(lpString, vbNullChar, vbNullString)
End Function

Private Function RightSlash(lpString As String) As String
    Const STR_BACKSLASH$ = "\"
    RightSlash = vbNullString
    If Len(lpString) > 0 Then
        If Right$(lpString, 1) <> STR_BACKSLASH Then RightSlash = lpString & STR_BACKSLASH Else RightSlash = lpString
    End If
End Function

Public Sub About()
Attribute About.VB_Description = "Show about box."
Attribute About.VB_UserMemId = -552
    frmSplash.Show vbModal
End Sub

Public Property Get BuildNumber() As Long
Attribute BuildNumber.VB_Description = "Return system OEM version."
    BuildNumber = m_BuildNumber
End Property

Public Property Let BuildNumber(ByVal New_BuildNumber As Long)
    If Ambient.UserMode = False Then Err.Raise 394
    If Ambient.UserMode Then Err.Raise 393
    m_BuildNumber = New_BuildNumber
    PropertyChanged "BuildNumber"
End Property

Public Property Get SysKeysDisabled() As Boolean
Attribute SysKeysDisabled.VB_Description = "Enabled or disabled system keys events (CTRL-ALT-DEL, ALT-TAB and more)."
Attribute SysKeysDisabled.VB_ProcData.VB_Invoke_Property = "Desktop"
    SysKeysDisabled = m_SysKeysDisabled
End Property

Public Property Let SysKeysDisabled(New_Enabled As Boolean)
    Const STR_NAME$ = "§§PRIORE§"
    Const SPI_SCREENSAVERRUNNING = 97
    m_SysKeysDisabled = New_Enabled
    Call SystemParametersInfo(SPI_SCREENSAVERRUNNING, New_Enabled, STR_NAME, 0)
End Property

Public Property Get DesktopHandle() As Long
Attribute DesktopHandle.VB_Description = "Return desktop handle."
    m_DesktopHandle = GetDesktopWindow
    DesktopHandle = m_DesktopHandle
End Property

Public Property Let DesktopHandle(ByVal New_DesktopHandle As Long)
    If Ambient.UserMode = False Then Err.Raise 394
    If Ambient.UserMode Then Err.Raise 393
    m_DesktopHandle = New_DesktopHandle
    PropertyChanged "DesktopHandle"
End Property

Public Property Get DesktopIcons() As Boolean
Attribute DesktopIcons.VB_Description = "Hide or show desktop icons."
Attribute DesktopIcons.VB_ProcData.VB_Invoke_Property = "Desktop"
    Dim Handle As Long
    
    Handle = GetWindow(FindWindow(STR_PROGMAN, STR_PROGRAMMANAGER), GW_CHILD)
    DesktopIcons = IsWindowVisible(Handle) Or IsWindowEnabled(Handle)
End Property

Public Property Let DesktopIcons(New_Enabled As Boolean)
    Dim Handle As Long
    Handle = GetWindow(FindWindow(STR_PROGMAN, STR_PROGRAMMANAGER), GW_CHILD)
    If New_Enabled Then
        If lWinID = VER_PLATFORM_WIN32_NT Then Call ShowWindow(Handle, SW_SHOW Or SW_RESTORE)
        EnableWindow Handle, True
    Else
        If lWinID = VER_PLATFORM_WIN32_NT Then Call ShowWindow(Handle, SW_HIDE)
        EnableWindow Handle, False
    End If
End Property

Public Property Get MajorVersion() As Long
Attribute MajorVersion.VB_Description = "Return system major version."
    MajorVersion = m_MajorVersion
End Property

Public Property Let MajorVersion(ByVal New_MajorVersion As Long)
    If Ambient.UserMode = False Then Err.Raise 394
    If Ambient.UserMode Then Err.Raise 393
    m_MajorVersion = New_MajorVersion
    PropertyChanged "MajorVersion"
End Property

Public Property Get MinorVersion() As Long
Attribute MinorVersion.VB_Description = "Return system  minor version."
    MinorVersion = m_MinorVersion
End Property

Public Property Let MinorVersion(ByVal New_MinorVersion As Long)
    If Ambient.UserMode = False Then Err.Raise 394
    If Ambient.UserMode Then Err.Raise 393
    m_MinorVersion = New_MinorVersion
    PropertyChanged "MinorVersion"
End Property

Public Property Get PlatformID() As String
Attribute PlatformID.VB_Description = "Return system platform version."
    PlatformID = m_PlatformID
End Property

Public Property Let PlatformID(ByVal New_PlatformID As String)
    If Ambient.UserMode = False Then Err.Raise 394
    If Ambient.UserMode Then Err.Raise 393
    m_PlatformID = New_PlatformID
    PropertyChanged "PlatformID"
End Property

Public Property Get ProcessorType() As Long
Attribute ProcessorType.VB_Description = "Return CPU type."
    Dim SYSINF As SYSTEM_INFO
    GetSystemInfo SYSINF
    m_ProcessorType = SYSINF.dwProcessorType
    ProcessorType = m_ProcessorType
End Property

Public Property Let ProcessorType(ByVal New_ProcessorType As Long)
    If Ambient.UserMode = False Then Err.Raise 394
    If Ambient.UserMode Then Err.Raise 393
    m_ProcessorType = New_ProcessorType
    PropertyChanged "ProcessorType"
End Property

Public Property Get SoundCard() As Boolean
Attribute SoundCard.VB_Description = "Return if sound card is present."
    m_SoundCard = (waveOutGetNumDevs > 0)
    SoundCard = m_SoundCard
End Property

Public Property Let SoundCard(ByVal New_SoundCard As Boolean)
    If Ambient.UserMode = False Then Err.Raise 394
    If Ambient.UserMode Then Err.Raise 393
    m_SoundCard = New_SoundCard
    PropertyChanged "SoundCard"
End Property

Public Property Get StartBar() As Boolean
Attribute StartBar.VB_Description = "Hide or show application bar."
Attribute StartBar.VB_ProcData.VB_Invoke_Property = "Desktop"
    Dim Handle As Long
    
    Handle = FindWindow(STR_STARTBAR, vbNullString)
    StartBar = IsWindowVisible(Handle) ' Or IsWindowEnabled(Handle)
End Property

Public Property Let StartBar(New_Enabled As Boolean)
    Dim Handle As Long
    
    Handle = FindWindow(STR_STARTBAR, vbNullString)
    Call SetWindowPos(Handle, 0, 0, 0, 0, 0, IIf(New_Enabled, SWP_SHOWWINDOW, SWP_HIDEWINDOW))
    'If New_Enabled Then
    '    If lWinID = VER_PLATFORM_WIN32_NT Then Call ShowWindow(Handle, SW_SHOW Or SW_RESTORE)
    '    Call EnableWindow(Handle, True)
    'Else
    '    If lWinID = VER_PLATFORM_WIN32_NT Then Call ShowWindow(Handle, SW_HIDE)
    '    Call EnableWindow(Handle, False)
    'End If
End Property

Public Property Get TempFile() As String
Attribute TempFile.VB_Description = "Return  first temp file name available."
Attribute TempFile.VB_ProcData.VB_Invoke_Property = "TempFile"
    Dim tmppath As String
    Dim nomefile As String
    Dim nChar As Long
    Dim nfile As Long
    tmppath = String$(256, 0)
    nomefile = String$(256, 0)
    nChar = GetTempPath(Len(tmppath), tmppath)
    If nChar > 0 Then nfile = GetTempFileName(tmppath, m_TempFileSuffix, 0, nomefile)
    If nfile > 0 Then nomefile = Left$(nomefile, InStr(nomefile, vbNullChar) - 1)
    m_TempFile = RemoveNulls(nomefile)
    If Not m_CreateTempFile Then DeleteFile m_TempFile
    TempFile = m_TempFile
End Property

Public Property Let TempFile(ByVal New_TempFile As String)
    If Ambient.UserMode = False Then Err.Raise 394
    If Ambient.UserMode Then Err.Raise 393
    m_TempFile = New_TempFile
    PropertyChanged "TempFile"
End Property

Public Property Get TempPath() As String
Attribute TempPath.VB_Description = "Return  standard temp path name."
    Dim tmppath As String
    tmppath = String$(256, 0)
    GetTempPath Len(tmppath), tmppath
    m_TempPath = RightSlash(RemoveNulls(tmppath))
    TempPath = m_TempPath
End Property

Public Property Let TempPath(ByVal New_TempPath As String)
    If Ambient.UserMode = False Then Err.Raise 394
    If Ambient.UserMode Then Err.Raise 393
    m_TempPath = New_TempPath
    PropertyChanged "TempPath"
End Property

Public Property Get UserName() As String
Attribute UserName.VB_Description = "Return  net user name."
    Dim strBuffer As String
    strBuffer = String$(255, 0)
    GetUserName strBuffer, Len(strBuffer)
    m_UserName = RemoveNulls(strBuffer)
    UserName = m_UserName
End Property

Public Property Let UserName(ByVal New_UserName As String)
    If Ambient.UserMode = False Then Err.Raise 394
    If Ambient.UserMode Then Err.Raise 393
    m_UserName = New_UserName
    PropertyChanged "UserName"
End Property

Public Property Get WinPath() As String
Attribute WinPath.VB_Description = "Return windows standard path name."
    Dim spath As String
    spath = String$(256, 0)
    GetWindowsDirectory spath, Len(spath)
    m_WinPath = RightSlash(RemoveNulls(spath))
    WinPath = m_WinPath
End Property

Public Property Let WinPath(ByVal New_WinPath As String)
    If Ambient.UserMode = False Then Err.Raise 394
    If Ambient.UserMode Then Err.Raise 393
    m_WinPath = New_WinPath
    PropertyChanged "WinPath"
End Property

Public Property Get WinSysPath() As String
Attribute WinSysPath.VB_Description = "Return system standard path name."
    Dim spath As String
    spath = String$(256, 0)
    GetSystemDirectory spath, Len(spath)
    m_WinSysPath = RightSlash(RemoveNulls(spath))
    WinSysPath = m_WinSysPath
End Property

Public Property Let WinSysPath(ByVal New_WinSysPath As String)
    If Ambient.UserMode = False Then Err.Raise 394
    If Ambient.UserMode Then Err.Raise 393
    m_WinSysPath = New_WinSysPath
    PropertyChanged "WinSysPath"
End Property

Public Property Get WinStart() As Long
Attribute WinStart.VB_Description = "Return running windows time."
    m_WinStart = GetTickCount
    WinStart = m_WinStart
End Property

Public Property Let WinStart(ByVal New_WinStart As Long)
    If Ambient.UserMode = False Then Err.Raise 394
    If Ambient.UserMode Then Err.Raise 393
    m_WinStart = New_WinStart
    PropertyChanged "WinStart"
End Property

Private Sub UserControl_Initialize()
    #If SHAREWARE = 1 Then
        frmSplash.Show vbModal
    #End If
    
    Call GetInfo
End Sub

'Inizializza le proprietà di UserControl
Private Sub UserControl_InitProperties()
    m_BuildNumber = m_def_BuildNumber
    m_DesktopHandle = m_def_DesktopHandle
    m_MajorVersion = m_def_MajorVersion
    m_MinorVersion = m_def_MinorVersion
    m_PlatformID = m_def_PlatformID
    m_ProcessorType = m_def_ProcessorType
    m_SoundCard = m_def_SoundCard
    m_TempFile = m_def_TempFile
    m_TempPath = m_def_TempPath
    m_UserName = m_def_UserName
    m_WinPath = m_def_WinPath
    m_WinSysPath = m_def_WinSysPath
    m_WinStart = m_def_WinStart
    m_ComputerName = m_def_ComputerName
    m_TempFileSuffix = m_def_TempFileSuffix
    
    Call GetInfo
    
    m_SysKeysDisabled = m_def_SysKeysDisabled
    m_CreateTempFile = m_def_CreateTempFile
    m_TitleColor = m_def_TitleColor
    m_ExitType = m_def_ExitType
    m_Display = 0
    m_OnTop = m_def_OnTop
End Sub

'Carica i valori della proprietà dalla memoria
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    m_BuildNumber = PropBag.ReadProperty("BuildNumber", m_def_BuildNumber)
    m_DesktopHandle = PropBag.ReadProperty("DesktopHandle", m_def_DesktopHandle)
    m_MajorVersion = PropBag.ReadProperty("MajorVersion", m_def_MajorVersion)
    m_MinorVersion = PropBag.ReadProperty("MinorVersion", m_def_MinorVersion)
    m_PlatformID = PropBag.ReadProperty("PlatformID", m_def_PlatformID)
    m_ProcessorType = PropBag.ReadProperty("ProcessorType", m_def_ProcessorType)
    m_SoundCard = PropBag.ReadProperty("SoundCard", m_def_SoundCard)
    m_TempFile = PropBag.ReadProperty("TempFile", m_def_TempFile)
    m_TempPath = PropBag.ReadProperty("TempPath", m_def_TempPath)
    m_UserName = PropBag.ReadProperty("UserName", m_def_UserName)
    m_WinPath = PropBag.ReadProperty("WinPath", m_def_WinPath)
    m_WinSysPath = PropBag.ReadProperty("WinSysPath", m_def_WinSysPath)
    m_WinStart = PropBag.ReadProperty("WinStart", m_def_WinStart)
    m_ComputerName = PropBag.ReadProperty("ComputerName", m_def_ComputerName)
    m_TempFileSuffix = PropBag.ReadProperty("TempFileSuffix", m_def_TempFileSuffix)
    m_CreateTempFile = PropBag.ReadProperty("CreateTempFile", m_def_CreateTempFile)
    m_TitleColor = PropBag.ReadProperty("TitleColor", m_def_TitleColor)
    m_SysKeysDisabled = PropBag.ReadProperty("SysKeysDisabled", m_def_SysKeysDisabled)
    m_ExitType = PropBag.ReadProperty("ExitType", m_def_ExitType)
    m_Display = PropBag.ReadProperty("Display", 0)
    m_OnTop = PropBag.ReadProperty("OnTop", m_def_OnTop)
End Sub

Private Sub UserControl_Resize()
    On Local Error Resume Next
    Width = 480
    Height = 480
End Sub

'Scrive i valori della proprietà in memoria
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("BuildNumber", m_BuildNumber, m_def_BuildNumber)
    Call PropBag.WriteProperty("DesktopHandle", m_DesktopHandle, m_def_DesktopHandle)
    Call PropBag.WriteProperty("MajorVersion", m_MajorVersion, m_def_MajorVersion)
    Call PropBag.WriteProperty("MinorVersion", m_MinorVersion, m_def_MinorVersion)
    Call PropBag.WriteProperty("PlatformID", m_PlatformID, m_def_PlatformID)
    Call PropBag.WriteProperty("ProcessorType", m_ProcessorType, m_def_ProcessorType)
    Call PropBag.WriteProperty("SoundCard", m_SoundCard, m_def_SoundCard)
    Call PropBag.WriteProperty("TempFile", m_TempFile, m_def_TempFile)
    Call PropBag.WriteProperty("TempPath", m_TempPath, m_def_TempPath)
    Call PropBag.WriteProperty("UserName", m_UserName, m_def_UserName)
    Call PropBag.WriteProperty("WinPath", m_WinPath, m_def_WinPath)
    Call PropBag.WriteProperty("WinSysPath", m_WinSysPath, m_def_WinSysPath)
    Call PropBag.WriteProperty("WinStart", m_WinStart, m_def_WinStart)
    Call PropBag.WriteProperty("ComputerName", m_ComputerName, m_def_ComputerName)
    Call PropBag.WriteProperty("TempFileSuffix", m_TempFileSuffix, m_def_TempFileSuffix)
    Call PropBag.WriteProperty("CreateTempFile", m_CreateTempFile, m_def_CreateTempFile)
    Call PropBag.WriteProperty("TitleColor", m_TitleColor, m_def_TitleColor)
    Call PropBag.WriteProperty("SysKeysDisabled", m_SysKeysDisabled, m_def_SysKeysDisabled)
    Call PropBag.WriteProperty("ExitType", m_ExitType, m_def_ExitType)
    Call PropBag.WriteProperty("Display", m_Display, 0)
    Call PropBag.WriteProperty("OnTop", m_OnTop, m_def_OnTop)
End Sub

Public Property Get ComputerName() As String
Attribute ComputerName.VB_Description = "Return net computer name."
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

Public Property Get TempFileSuffix() As String
Attribute TempFileSuffix.VB_Description = "Return or set  suffix for temp files."
Attribute TempFileSuffix.VB_ProcData.VB_Invoke_Property = "TempFile"
    TempFileSuffix = m_TempFileSuffix
End Property

Public Property Let TempFileSuffix(ByVal New_TempFileSuffix As String)
    m_TempFileSuffix = New_TempFileSuffix
    PropertyChanged "TempFileSuffix"
End Property

Public Property Get CreateTempFile() As Boolean
Attribute CreateTempFile.VB_Description = "Enable or disable to create temp file."
Attribute CreateTempFile.VB_ProcData.VB_Invoke_Property = "TempFile"
    CreateTempFile = m_CreateTempFile
End Property

Public Property Let CreateTempFile(ByVal New_CreateTempFile As Boolean)
    m_CreateTempFile = New_CreateTempFile
    PropertyChanged "CreateTempFile"
End Property

Public Property Get TitleColor() As OLE_COLOR
Attribute TitleColor.VB_Description = "Return or set title bars color."
    m_TitleColor = GetSysColor(COLOR_ACTIVECAPTION)
    TitleColor = m_TitleColor
End Property

Public Property Let TitleColor(ByVal New_TitleColor As OLE_COLOR)
    m_TitleColor = New_TitleColor
    SetSysColors 1, COLOR_ACTIVECAPTION, m_TitleColor
    PropertyChanged "TitleColor"
End Property

Public Function ShellPlus(sCommand As String) As Long
Attribute ShellPlus.VB_Description = "Execute other application with pause."
    Dim i As Integer
    Dim idTask As Long
    Dim hProc As Long
    
    ' esegue l'applicazione
    idTask = Shell(sCommand, vbNormalFocus)
    ' ritorna handle del task
    hProc = OpenProcess(&H100000, False, idTask)
    ' controlla se deve rimanere in attesa della
    ' fine dell'applicazione prima di ritornare
    ' all'applicazione origine
    WaitForSingleObject hProc, &HFFFFFFFF
    CloseHandle hProc
    ShellPlus = idTask
End Function

Public Function Delay(iSeconds As Integer, Optional bDoEvents) As Long
Attribute Delay.VB_Description = "Run a  pause with or without other events."
    Dim iStart As Single
    iStart = Timer
    Delay = iSeconds
    Do While Timer < iStart + iSeconds
        If bDoEvents Then DoEvents
    Loop
End Function

'AVVISO: NON RIMUOVERE O MODIFICARE LE SEGUENTI RIGHE DI COMMENTO
'MemberInfo=8,0,0,0
Public Property Get ExitType() As ExitTypeClass
Attribute ExitType.VB_Description = "Return or set exit windows type."
    ExitType = m_ExitType
End Property

Public Property Let ExitType(ByVal New_ExitType As ExitTypeClass)
    m_ExitType = New_ExitType
    PropertyChanged "ExitType"
End Property

'AVVISO: NON RIMUOVERE O MODIFICARE LE SEGUENTI RIGHE DI COMMENTO
'MemberInfo=8
Public Sub ExitWindow()
Attribute ExitWindow.VB_Description = "Execute a exit windows."
    Call ExitWindowsEx(m_ExitType, 0)
End Sub

'AVVISO: NON RIMUOVERE O MODIFICARE LE SEGUENTI RIGHE DI COMMENTO
'MemberInfo=13,0,0,0
Public Property Get Display() As DisplaySettingsClass
Attribute Display.VB_Description = "Return or set size of screen and colors."
Attribute Display.VB_ProcData.VB_Invoke_Property = "Desktop"
    Select Case (Screen.Width / Screen.TwipsPerPixelX)
        Case 640
            If (Screen.Height / Screen.TwipsPerPixelY) = 400 Then m_Display = os640x400 Else m_Display = os640x480
        Case 800
            m_Display = os800x600
        Case 1024
            m_Display = os1024x768
        Case 1152
            m_Display = os1152x864
        Case 1280
            m_Display = os1280x1024
        Case 1600
            m_Display = os1600x1200
        Case Else
            m_Display = osUndefinited
    End Select
    Display = m_Display
End Property

Public Property Let Display(ByVal New_Display As DisplaySettingsClass)
    Dim devm As DEVMODE
    
    m_Display = New_Display
    Call EnumDisplaySettings(&H0, &H0, devm)
    With devm
        Select Case m_Display
            Case os640x400
                .dmPelsWidth = 640
                .dmPelsHeight = 400
            Case os640x480
                .dmPelsWidth = 640
                .dmPelsHeight = 480
            Case os800x600
                .dmPelsWidth = 800
                .dmPelsHeight = 60
            Case os1024x768
                .dmPelsWidth = 1024
                .dmPelsHeight = 768
            Case os1152x864
                .dmPelsWidth = 1152
                .dmPelsHeight = 864
            Case os1280x1024
                .dmPelsWidth = 1280
                .dmPelsHeight = 1024
            Case os1600x1200
                .dmPelsWidth = 1600
                .dmPelsHeight = 1200
            Case Else
                Exit Property
        End Select
        .dmFields = DM_PELSWIDTH Or DM_PELSHEIGHT
    End With
    Select Case ChangeDisplaySettings(devm, CDS_TEST)
        Case DISP_CHANGE_RESTART
            RaiseEvent DisplayChanged(True, False)
        Case DISP_CHANGE_SUCCESSFUL
            Call ChangeDisplaySettings(devm, CDS_UPDATEREGISTRY)
            RaiseEvent DisplayChanged(False, False)
        Case Else
            RaiseEvent DisplayChanged(False, True)
    End Select
    PropertyChanged "Display"
End Property

'AVVISO: NON RIMUOVERE O MODIFICARE LE SEGUENTI RIGHE DI COMMENTO
'MemberInfo=0,0,0,false
Public Property Get OnTop() As Boolean
Attribute OnTop.VB_Description = "Return ors set Windows stay on top state."
    OnTop = m_OnTop
End Property

Public Property Let OnTop(ByVal New_OnTop As Boolean)
    m_OnTop = New_OnTop
    SetWindowPos Parent.hwnd, IIf(m_OnTop, HWND_TOPMOST, HWND_NOTOPMOST), 0, 0, 0, 0, SWP_NORMAL Or SWP_NOSIZE Or SWP_NOMOVE Or SWP_DRAWFRAME
    PropertyChanged "OnTop"
End Property
