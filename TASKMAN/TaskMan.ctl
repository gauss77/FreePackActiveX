VERSION 5.00
Begin VB.UserControl TaskManager 
   CanGetFocus     =   0   'False
   ClientHeight    =   480
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   480
   ClipBehavior    =   0  'None
   HitBehavior     =   0  'None
   InvisibleAtRuntime=   -1  'True
   Picture         =   "TaskMan.ctx":0000
   PropertyPages   =   "TaskMan.ctx":08CA
   ScaleHeight     =   480
   ScaleWidth      =   480
   ToolboxBitmap   =   "TaskMan.ctx":08DC
   Windowless      =   -1  'True
End
Attribute VB_Name = "TaskManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "PropPageWizardRun" ,"Yes"
Option Explicit

'Valori predefiniti proprietà:
Const m_def_Service = False

'Variabili proprietà:
Dim m_Service As Boolean

Public Enum PriorityClass
    Normal = &H20
    Idle = &H40
    High = &H80
    RealTime = &H100
End Enum

Private Const PROCESS_DUP_HANDLE = &H40

Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function GetCurrentProcessId Lib "kernel32" () As Long
Private Declare Function GetPriorityClass Lib "kernel32" (ByVal hProcess As Long) As Long
Private Declare Function SetPriorityClass Lib "kernel32" (ByVal hProcess As Long, ByVal dwPriorityClass As Long)
Private Declare Function RegisterServiceProcess Lib "kernel32.dll" (ByVal dwProcessId As Long, ByVal dwType As Long) As Long

Private Sub ChangePriority(dwPriorityClass As Long)
    Dim pid As Long
    Dim hProcess As Long
    
    pid = GetCurrentProcessId() ' get my proccess id
    ' get a handle to the process
    hProcess = OpenProcess(PROCESS_DUP_HANDLE, True, pid)
    If hProcess = 0 Then Exit Sub

    ' change the priority
    Call SetPriorityClass(hProcess, dwPriorityClass)
    Call CloseHandle(hProcess)
End Sub

Private Function GetPriority() As Long
    Dim pid As Long
    Dim hProcess As Long
    
    pid = GetCurrentProcessId() ' get my proccess id
    ' get a handle to the process
    hProcess = OpenProcess(PROCESS_DUP_HANDLE, True, pid)
    If hProcess = 0 Then Exit Function

    ' change the priority
    GetPriority = GetPriorityClass(hProcess)
    Call CloseHandle(hProcess)
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

'Inizializza le proprietà di UserControl.
Private Sub UserControl_InitProperties()
    m_Service = m_def_Service
End Sub

'Carica i valori delle proprietà dalla posizione di memorizzazione.
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    m_Service = PropBag.ReadProperty("Service", m_def_Service)
End Sub

'Scrive i valori delle proprietà nella posizione di memorizzazione.
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("Service", m_Service, m_def_Service)
End Sub

Public Sub About()
Attribute About.VB_Description = "SHow about box."
Attribute About.VB_UserMemId = -552
    frmSplash.Show vbModal
End Sub

'AVVISO: NON RIMUOVERE O MODIFICARE LE SEGUENTI RIGHE DI COMMENTO
'MemberInfo=8,1,1,0
Public Property Get ProcessID() As Long
Attribute ProcessID.VB_Description = "Return process ID."
    ProcessID = GetCurrentProcessId()
End Property

Public Property Let ProcessID(ByVal New_ID As Long)
    If Ambient.UserMode = False Then Err.Raise 387
    If Ambient.UserMode Then Err.Raise 382
End Property

'AVVISO: NON RIMUOVERE O MODIFICARE LE SEGUENTI RIGHE DI COMMENTO
'MemberInfo=8,0,1,&H20
Public Property Get Priority() As PriorityClass
Attribute Priority.VB_Description = "Return or set process priority."
    Priority = GetPriority
End Property

Public Property Let Priority(ByVal New_Priority As PriorityClass)
    If Ambient.UserMode = False Then Err.Raise 387
    Call ChangePriority(New_Priority)
    PropertyChanged "Priority"
End Property

'AVVISO: NON RIMUOVERE O MODIFICARE LE SEGUENTI RIGHE DI COMMENTO
'MemberInfo=0,0,1,False
Public Property Get Service() As Boolean
Attribute Service.VB_Description = "Return or set current application running to a service (only Windows 95/98)."
    Service = m_Service
End Property

Public Property Let Service(ByVal New_Service As Boolean)
    If Ambient.UserMode = False Then Err.Raise 387
    m_Service = New_Service
    Call RegisterServiceProcess(GetCurrentProcessId, IIf(m_Service, 1, 0))
    PropertyChanged "Service"
End Property

