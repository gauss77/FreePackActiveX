VERSION 5.00
Begin VB.UserControl MemoryInfo 
   CanGetFocus     =   0   'False
   ClientHeight    =   480
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   480
   ClipControls    =   0   'False
   InvisibleAtRuntime=   -1  'True
   Picture         =   "MemoryInfo.ctx":0000
   PropertyPages   =   "MemoryInfo.ctx":030A
   ScaleHeight     =   480
   ScaleWidth      =   480
   ToolboxBitmap   =   "MemoryInfo.ctx":031C
   Windowless      =   -1  'True
End
Attribute VB_Name = "MemoryInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "PropPageWizardRun" ,"Yes"
Option Explicit

'Valori predefiniti proprietà:
Const m_def_Free = 0
Const m_def_Total = 0
Const m_def_Virtual = 0
Const m_def_TotalVirtual = 0

'Struttura per la funzione API
Private Type MEMORYSTATUS
        dwLength As Long
        dwMemoryLoad As Long
        dwTotalPhys As Long
        dwAvailPhys As Long
        dwTotalPageFile As Long
        dwAvailPageFile As Long
        dwTotalVirtual As Long
        dwAvailVirtual As Long
End Type

'Variabili proprietà:
Dim m_Free As Long
Dim m_Total As Long
Dim m_Virtual As Long
Dim m_TotalVirtual As Long

'Funzione api per ricava le informazioni sulla memoria
Private Declare Sub GlobalMemoryStatus Lib "kernel32" (lpBuffer As MEMORYSTATUS)

Public Sub About()
Attribute About.VB_Description = "Show about box."
Attribute About.VB_UserMemId = -552
    frmSplash.Show vbModal
End Sub

Public Property Get Free() As Long
Attribute Free.VB_Description = "Return  free memory available."
    Call GetMemory
    Free = m_Free
End Property

Public Property Let Free(ByVal New_Free As Long)
    If Ambient.UserMode = False Then Err.Raise 394
    If Ambient.UserMode Then Err.Raise 393
    m_Free = New_Free
    PropertyChanged "Free"
End Property

Public Property Get Total() As Long
Attribute Total.VB_Description = "Return  total memory available."
    Call GetMemory
    Total = m_Total
End Property

Public Property Let Total(ByVal New_Total As Long)
    If Ambient.UserMode = False Then Err.Raise 394
    If Ambient.UserMode Then Err.Raise 393
    m_Total = New_Total
    PropertyChanged "Total"
End Property

Public Property Get Virtual() As Long
Attribute Virtual.VB_Description = "Return  virtual memory available."
    Call GetMemory
    Virtual = m_Virtual
End Property

Public Property Let Virtual(ByVal New_Virtual As Long)
    If Ambient.UserMode = False Then Err.Raise 394
    If Ambient.UserMode Then Err.Raise 393
    m_Virtual = New_Virtual
    PropertyChanged "Virtual"
End Property

Public Property Get TotalVirtual() As Long
Attribute TotalVirtual.VB_Description = "Return  total virtual memory available."
    Call GetMemory
    TotalVirtual = m_TotalVirtual
End Property

Public Property Let TotalVirtual(ByVal New_TotalVirtual As Long)
    If Ambient.UserMode = False Then Err.Raise 394
    If Ambient.UserMode Then Err.Raise 393
    m_TotalVirtual = New_TotalVirtual
    PropertyChanged "TotalVirtual"
End Property

Private Sub UserControl_Initialize()
    #If SHAREWARE = 1 Then
        frmSplash.Show vbModal
    #End If
End Sub

'Inizializza le proprietà di UserControl
Private Sub UserControl_InitProperties()
    m_Free = m_def_Free
    m_Total = m_def_Total
    m_Virtual = m_def_Virtual
    m_TotalVirtual = m_def_TotalVirtual
End Sub

'Carica i valori della proprietà dalla memoria
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    m_Free = PropBag.ReadProperty("Free", m_def_Free)
    m_Total = PropBag.ReadProperty("Total", m_def_Total)
    m_Virtual = PropBag.ReadProperty("Virtual", m_def_Virtual)
    m_TotalVirtual = PropBag.ReadProperty("TotalVirtual", m_def_TotalVirtual)
End Sub

Private Sub UserControl_Resize()
    On Local Error Resume Next
    Width = 480
    Height = 480
End Sub

'Scrive i valori della proprietà in memoria
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("Free", m_Free, m_def_Free)
    Call PropBag.WriteProperty("Total", m_Total, m_def_Total)
    Call PropBag.WriteProperty("Virtual", m_Virtual, m_def_Virtual)
    Call PropBag.WriteProperty("TotalVirtual", m_TotalVirtual, m_def_TotalVirtual)
End Sub

Private Sub GetMemory()
    Const m_def_Divisor = 1048.576
    
    Dim lRet As Long
    Dim memstat As MEMORYSTATUS
    Call GlobalMemoryStatus(memstat)
    With memstat
        m_Free = .dwAvailPhys / m_def_Divisor
        m_Total = .dwTotalPhys / m_def_Divisor
        m_Virtual = .dwAvailVirtual / m_def_Divisor
        m_TotalVirtual = .dwTotalVirtual / m_def_Divisor
    End With
End Sub
