VERSION 5.00
Begin VB.UserControl KeyboardInfo 
   CanGetFocus     =   0   'False
   ClientHeight    =   480
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   480
   ClipControls    =   0   'False
   InvisibleAtRuntime=   -1  'True
   Picture         =   "KeyInfo.ctx":0000
   PropertyPages   =   "KeyInfo.ctx":0442
   ScaleHeight     =   480
   ScaleWidth      =   480
   ToolboxBitmap   =   "KeyInfo.ctx":045F
   Windowless      =   -1  'True
End
Attribute VB_Name = "KeyboardInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "PropPageWizardRun" ,"Yes"
Option Explicit
'Valori predefiniti proprietà:
Const m_def_Speed = 0
Const m_def_Delay = 0
Const m_def_CapsState = False
Const m_def_NumState = False
Const m_def_ScrollState = False
Const m_def_KeyboardFunctions = 0
Const m_def_KeyboardType = ""

' costanti per le API
Const VK_CAPITAL = &H14
Const VK_NUMLOCK = &H90
Const VK_SCROLL = &H91
Const SPI_GETKEYBOARDDELAY = 22
Const SPI_GETKEYBOARDSPEED = 10
Const SPI_SETKEYBOARDDELAY = 23
Const SPI_SETKEYBOARDSPEED = 11

'Variabili proprietà:
Dim m_Speed As Long
Dim m_Delay As Long
Dim m_CapsState As Boolean
Dim m_NumState As Boolean
Dim m_ScrollState As Boolean
Dim m_KeyboardFunctions As Long
Dim m_KeyboardType As String

Private Type KeyboardBytes
     kbByte(0 To 255) As Byte
End Type

Private Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer
Private Declare Function GetKeyboardType Lib "user32" (ByVal nTypeFlag As Long) As Long
Private Declare Function GetKeyboardLayout Lib "user32" (ByVal dwLayout As Long) As Long
Private Declare Function SetKeyboardState Lib "user32" (kbArray As KeyboardBytes) As Long
Private Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, ByVal lpvParam As Any, ByVal fuWinIni As Long) As Long

Public Sub About()
Attribute About.VB_Description = "Show about box."
Attribute About.VB_UserMemId = -552
    frmSplash.Show vbModal
End Sub

Public Property Get CapsState() As Boolean
Attribute CapsState.VB_Description = "Return or set CAPS LOCK key state."
Attribute CapsState.VB_ProcData.VB_Invoke_Property = "Keyboard"
    m_CapsState = GetKeyState(VK_CAPITAL) And 1
    CapsState = m_CapsState
End Property

Public Property Let CapsState(ByVal New_CapsState As Boolean)
    Dim kb As KeyboardBytes
    m_CapsState = New_CapsState
    kb.kbByte(VK_CAPITAL) = IIf(m_CapsState, 1, 0)
    Call SetKeyboardState(kb)
    PropertyChanged "CapsState"
End Property

Public Property Get NumState() As Boolean
Attribute NumState.VB_Description = "Return or set  NUM LOCK key state."
Attribute NumState.VB_ProcData.VB_Invoke_Property = "Keyboard"
    m_NumState = GetKeyState(VK_NUMLOCK) And 1
    NumState = m_NumState
End Property

Public Property Let NumState(ByVal New_NumState As Boolean)
    Dim kb As KeyboardBytes
    m_NumState = New_NumState
    kb.kbByte(VK_NUMLOCK) = IIf(m_NumState, 1, 0)
    Call SetKeyboardState(kb)
    PropertyChanged "NumState"
End Property

Public Property Get ScrollState() As Boolean
Attribute ScrollState.VB_Description = "Return or set SCROLL LOCK key state."
Attribute ScrollState.VB_ProcData.VB_Invoke_Property = "Keyboard"
    m_ScrollState = GetKeyState(VK_SCROLL) And 1
    ScrollState = m_ScrollState
End Property

Public Property Let ScrollState(ByVal New_ScrollState As Boolean)
    Dim kb As KeyboardBytes
    m_ScrollState = New_ScrollState
    kb.kbByte(VK_SCROLL) = IIf(m_ScrollState, 1, 0)
    Call SetKeyboardState(kb)
    PropertyChanged "ScrollState"
End Property

Public Property Get KeyboardFunctions() As Long
Attribute KeyboardFunctions.VB_Description = "Return the number of funtion keys."
    m_KeyboardFunctions = GetKeyboardType(2)
    KeyboardFunctions = m_KeyboardFunctions
End Property

Public Property Let KeyboardFunctions(ByVal New_KeyboardFunctions As Long)
    If Ambient.UserMode = False Then Err.Raise 394
    If Ambient.UserMode Then Err.Raise 393
    m_KeyboardFunctions = New_KeyboardFunctions
    PropertyChanged "KeyboardFunctions"
End Property

Public Property Get KeyboardType() As String
Attribute KeyboardType.VB_Description = "Return keyboard type."
    Select Case GetKeyboardType(0)
        Case 0
            m_KeyboardType = "PC 83 key"
        Case 3
            m_KeyboardType = "AT 84 key"
        Case 4
            m_KeyboardType = "Enhanced 101 or 102 key"
        Case Else
            m_KeyboardType = "Special"
    End Select
    KeyboardType = m_KeyboardType
End Property

Public Property Let KeyboardType(ByVal New_KeyboardType As String)
    If Ambient.UserMode = False Then Err.Raise 394
    If Ambient.UserMode Then Err.Raise 393
    m_KeyboardType = New_KeyboardType
    PropertyChanged "KeyboardType"
End Property

Private Sub UserControl_Initialize()
    #If SHAREWARE = 1 Then
        frmSplash.Show vbModal
    #End If
End Sub

'Inizializza le proprietà di UserControl
Private Sub UserControl_InitProperties()
    m_CapsState = m_def_CapsState
    m_NumState = m_def_NumState
    m_ScrollState = m_def_ScrollState
    m_KeyboardFunctions = m_def_KeyboardFunctions
    m_KeyboardType = m_def_KeyboardType
    Call SystemParametersInfo(SPI_GETKEYBOARDSPEED, 0, m_Speed, 0)
    Call SystemParametersInfo(SPI_GETKEYBOARDDELAY, 0, m_Delay, 0)
End Sub

'Carica i valori della proprietà dalla memoria
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    m_CapsState = PropBag.ReadProperty("CapsState", m_def_CapsState)
    m_NumState = PropBag.ReadProperty("NumState", m_def_NumState)
    m_ScrollState = PropBag.ReadProperty("ScrollState", m_def_ScrollState)
    m_KeyboardFunctions = PropBag.ReadProperty("KeyboardFunctions", m_def_KeyboardFunctions)
    m_KeyboardType = PropBag.ReadProperty("KeyboardType", m_def_KeyboardType)
    m_Speed = PropBag.ReadProperty("Speed", m_def_Speed)
    m_Delay = PropBag.ReadProperty("Delay", m_def_Delay)
End Sub

Private Sub UserControl_Resize()
    On Local Error Resume Next
    Width = 480
    Height = 480
End Sub

'Scrive i valori della proprietà in memoria
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("CapsState", m_CapsState, m_def_CapsState)
    Call PropBag.WriteProperty("NumState", m_NumState, m_def_NumState)
    Call PropBag.WriteProperty("ScrollState", m_ScrollState, m_def_ScrollState)
    Call PropBag.WriteProperty("KeyboardFunctions", m_KeyboardFunctions, m_def_KeyboardFunctions)
    Call PropBag.WriteProperty("KeyboardType", m_KeyboardType, m_def_KeyboardType)
    Call PropBag.WriteProperty("Speed", m_Speed, m_def_Speed)
    Call PropBag.WriteProperty("Delay", m_Delay, m_def_Delay)
End Sub

'AVVISO: NON RIMUOVERE O MODIFICARE LE SEGUENTI RIGHE DI COMMENTO
'MemberInfo=8,0,0,0
Public Property Get Speed() As Long
Attribute Speed.VB_Description = "Return or set key repeat  interval."
Attribute Speed.VB_ProcData.VB_Invoke_Property = "Keyboard"
    Call SystemParametersInfo(SPI_GETKEYBOARDSPEED, 0, m_Speed, 0)
    Speed = m_Speed
End Property

Public Property Let Speed(ByVal New_Speed As Long)
    m_Speed = New_Speed
    Call SystemParametersInfo(SPI_SETKEYBOARDSPEED, 0, m_Speed, 0)
    PropertyChanged "Speed"
End Property

'AVVISO: NON RIMUOVERE O MODIFICARE LE SEGUENTI RIGHE DI COMMENTO
'MemberInfo=8,0,0,0
Public Property Get Delay() As Long
Attribute Delay.VB_Description = "Return or set repeat keys  interval."
Attribute Delay.VB_ProcData.VB_Invoke_Property = "Keyboard"
    Call SystemParametersInfo(SPI_GETKEYBOARDDELAY, 0, m_Delay, 0)
    Delay = m_Delay
End Property

Public Property Let Delay(ByVal New_Delay As Long)
    m_Delay = New_Delay
    Call SystemParametersInfo(SPI_SETKEYBOARDDELAY, 0, m_Delay, 0)
    PropertyChanged "Delay"
End Property

