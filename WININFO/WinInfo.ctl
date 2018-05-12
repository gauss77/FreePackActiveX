VERSION 5.00
Begin VB.UserControl WindowsInfo 
   CanGetFocus     =   0   'False
   ClientHeight    =   480
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   480
   ClipControls    =   0   'False
   InvisibleAtRuntime=   -1  'True
   Picture         =   "WinInfo.ctx":0000
   PropertyPages   =   "WinInfo.ctx":030A
   ScaleHeight     =   480
   ScaleWidth      =   480
   ToolboxBitmap   =   "WinInfo.ctx":031C
End
Attribute VB_Name = "WindowsInfo"
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

'Valori predefiniti proprietà:
Const m_def_Caption = ""
Const m_def_Count = 0
Const m_def_WinIndex = 0

'Variabili proprietà:
Dim m_Count As Long
Dim m_WinIndex As Long
Dim m_WinLeft As Long
Dim m_WinTop As Long
Dim m_WinWidth As Long
Dim m_WinHeight As Long
Dim m_Caption As String

Private Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type

Private Const SWP_HIDEWINDOW = &H80
Private Const SWP_SHOWWINDOW = &H40

Private hWndWin() As Long

Private Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Private Declare Function GetWindow Lib "user32" (ByVal hwnd As Long, ByVal wCmd As Long) As Long
Private Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Private Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hwnd As Long) As Long
Private Declare Function GetActiveWindow Lib "user32" () As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Private Declare Function IsWindowEnabled Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function IsWindowVisible Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function IsZoomedA Lib "user32" Alias "IsZoomed" (ByVal hwnd As Long) As Long
Private Declare Function IsIconicA Lib "user32" Alias "IsIconic" (ByVal hwnd As Long) As Long
Private Declare Function IsWindow Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare Function SetWindowText Lib "user32" Alias "SetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String) As Long

Private Sub GetWinSize()
    Dim RET As RECT
    Call GetWindowRect(hWndWin(m_WinIndex), RET)
    With RET
        m_WinLeft = .Left
        m_WinTop = .Top
        m_WinWidth = .Right - .Left
        m_WinHeight = .Bottom - .Top
    End With
End Sub

Public Property Get Caption() As String
Attribute Caption.VB_Description = "Return ors set current windows title."
    Dim nChar As Long
    Dim sTmp As String
    m_Caption = m_def_Caption
    nChar = GetWindowTextLength(hWndWin(m_WinIndex))
    sTmp = String$(nChar + 1, 0)
    nChar = GetWindowText(hWndWin(m_WinIndex), sTmp, Len(sTmp))
    If nChar > 0 Then m_Caption = Replace$(sTmp, vbNullChar, vbNullString)
    Caption = m_Caption
End Property

Public Property Let Caption(ByVal New_Caption As String)
    m_Caption = New_Caption
    SetWindowText hWndWin(m_WinIndex), m_Caption
    PropertyChanged "Caption"
End Property

Public Property Get IsIconic() As Boolean
Attribute IsIconic.VB_Description = "Return  current window minimized state."
    IsIconic = IsIconicA(hWndWin(m_WinIndex))
End Property

Public Property Let IsIconic(ByVal New_IsIconic As Boolean)
    If Ambient.UserMode = False Then Err.Raise 394
    If Ambient.UserMode Then Err.Raise 393
End Property

Public Property Get IsEnabled() As Boolean
Attribute IsEnabled.VB_Description = "Return current window enabled state."
    IsEnabled = IsWindowEnabled(hWndWin(m_WinIndex))
End Property

Public Property Let IsEnabled(ByVal New_IsEnabled As Boolean)
    If Ambient.UserMode = False Then Err.Raise 394
    If Ambient.UserMode Then Err.Raise 393
End Property

Public Property Get IsVisible() As Boolean
Attribute IsVisible.VB_Description = "Return current window visible state."
    IsVisible = IsWindowVisible(hWndWin(m_WinIndex))
End Property

Public Property Let IsVisible(ByVal New_IsVisible As Boolean)
    If Ambient.UserMode = False Then Err.Raise 394
    If Ambient.UserMode Then Err.Raise 393
End Property

Public Property Get IsZoomed() As Boolean
Attribute IsZoomed.VB_Description = "Return current window maximized state."
    IsZoomed = IsZoomedA(hWndWin(m_WinIndex))
End Property

Public Property Let IsZoomed(ByVal New_IsZoomed As Boolean)
    If Ambient.UserMode = False Then Err.Raise 394
    If Ambient.UserMode Then Err.Raise 393
End Property

Public Property Get Count() As Long
Attribute Count.VB_Description = "Return numbers of windows available."
    Count = m_Count
End Property

Public Property Let Count(ByVal New_Count As Long)
    If Ambient.UserMode = False Then Err.Raise 394
    If Ambient.UserMode Then Err.Raise 393
End Property

Public Property Get WinIndex() As Long
Attribute WinIndex.VB_Description = "Return or set window index."
    WinIndex = m_WinIndex
End Property

Public Property Let WinIndex(ByVal New_WinIndex As Long)
    If New_WinIndex <= m_Count Then
        m_WinIndex = New_WinIndex
        Call GetWinSize
        PropertyChanged "WinIndex"
    End If
End Property

Public Property Get IsActive() As Boolean
Attribute IsActive.VB_Description = "Return current window active state."
    If hWndWin(m_WinIndex) = GetActiveWindow Then IsActive = True Else IsActive = False
End Property

Public Property Let IsActive(ByVal New_IsActive As Boolean)
    If Ambient.UserMode = False Then Err.Raise 394
    If Ambient.UserMode Then Err.Raise 393
End Property

Public Property Get WinLeft() As Long
Attribute WinLeft.VB_Description = "Return current window left position."
    WinLeft = m_WinLeft
End Property

Public Property Let WinLeft(ByVal New_WinLeft As Long)
    If Ambient.UserMode = False Then Err.Raise 394
    If Ambient.UserMode Then Err.Raise 393
End Property

Public Property Get WinTop() As Long
Attribute WinTop.VB_Description = "Return  current window top position."
    WinTop = m_WinTop
End Property

Public Property Let WinTop(ByVal New_WinTop As Long)
    If Ambient.UserMode = False Then Err.Raise 394
    If Ambient.UserMode Then Err.Raise 393
End Property

Public Property Get WinWidth() As Long
Attribute WinWidth.VB_Description = "Return  current window width size."
    WinWidth = m_WinWidth
End Property

Public Property Let WinWidth(ByVal New_WinWidth As Long)
    If Ambient.UserMode = False Then Err.Raise 394
    If Ambient.UserMode Then Err.Raise 393
End Property

Public Property Get WinHeight() As Long
Attribute WinHeight.VB_Description = "Return current window height."
    WinHeight = m_WinHeight
End Property

Public Property Let WinHeight(ByVal New_WinHeight As Long)
    If Ambient.UserMode = False Then Err.Raise 394
    If Ambient.UserMode Then Err.Raise 393
End Property

'Inizializza le proprietà di UserControl
Private Sub UserControl_InitProperties()
    m_Caption = ""
    m_Count = m_def_Count
    m_WinIndex = m_def_WinIndex
    m_WinLeft = 0
    m_WinTop = 0
    m_WinWidth = 0
    m_WinHeight = 0
    
    Me.Refresh
    If m_Count > 0 Then
        m_WinIndex = 1
        Call GetWinSize
    End If
End Sub

'Carica i valori della proprietà dalla memoria
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    m_Caption = PropBag.ReadProperty("Caption", "")
    m_Count = PropBag.ReadProperty("Count", m_def_Count)
    m_WinIndex = PropBag.ReadProperty("WinIndex", m_def_WinIndex)
    m_WinLeft = PropBag.ReadProperty("WinLeft", 0)
    m_WinTop = PropBag.ReadProperty("WinTop", 0)
    m_WinWidth = PropBag.ReadProperty("WinWidth", 0)
    m_WinHeight = PropBag.ReadProperty("WinHeight", 0)
End Sub

Private Sub UserControl_Resize()
    On Local Error Resume Next
    Width = 480
    Height = 480
End Sub

'Scrive i valori della proprietà in memoria
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("Caption", m_Caption, "")
    Call PropBag.WriteProperty("Count", m_Count, m_def_Count)
    Call PropBag.WriteProperty("WinIndex", m_WinIndex, m_def_WinIndex)
    Call PropBag.WriteProperty("WinLeft", m_WinLeft, 0)
    Call PropBag.WriteProperty("WinTop", m_WinTop, 0)
    Call PropBag.WriteProperty("WinWidth", m_WinWidth, 0)
    Call PropBag.WriteProperty("WinHeight", m_WinHeight, 0)
End Sub

Public Function Refresh() As Variant
Attribute Refresh.VB_Description = "Refresh windows list."
    Dim Hnd As Long
    m_Count = m_def_Count
    Hnd = GetWindow(Parent.hwnd, 0)
    Do While Hnd <> &H0
        If IsWindow(Hnd) And GetWindowTextLength(Hnd) > 0 Then
            m_Count = m_Count + 1
            ReDim Preserve hWndWin(m_Count)
            hWndWin(m_Count) = Hnd
        End If
        Hnd = GetWindow(Hnd, 2)
    Loop
End Function

Public Property Get Handle() As Long
Attribute Handle.VB_Description = "Return current window handle."
    Handle = hWndWin(m_WinIndex)
End Property

Public Property Let Handle(ByVal New_Handle As Long)
    If Ambient.UserMode = False Then Err.Raise 394
    If Ambient.UserMode Then Err.Raise 393
End Property

'AVVISO: NON RIMUOVERE O MODIFICARE LE SEGUENTI RIGHE DI COMMENTO
'MemberInfo=14
Public Sub Show()
Attribute Show.VB_Description = "Show current window."
    Call SetWindowPos(hWndWin(m_WinIndex), 0, 0, 0, 0, 0, SWP_SHOWWINDOW)
End Sub

'AVVISO: NON RIMUOVERE O MODIFICARE LE SEGUENTI RIGHE DI COMMENTO
'MemberInfo=14
Public Sub Hide()
Attribute Hide.VB_Description = "Hide current window."
    Call SetWindowPos(hWndWin(m_WinIndex), 0, 0, 0, 0, 0, SWP_HIDEWINDOW)
End Sub

'AVVISO: NON RIMUOVERE O MODIFICARE LE SEGUENTI RIGHE DI COMMENTO
'MemberInfo=13,1,1,
Public Property Get ClassName() As String
Attribute ClassName.VB_Description = "Return current window class."
    Dim n As String
    n = String(255, vbNullChar)
    Call GetClassName(hWndWin(m_WinIndex), n, Len(n))
    ClassName = Replace$(n, vbNullChar, vbNullString)
End Property

Public Property Let ClassName(ByVal New_ClassName As String)
    If Ambient.UserMode = False Then Err.Raise 387
    If Ambient.UserMode Then Err.Raise 382
End Property

