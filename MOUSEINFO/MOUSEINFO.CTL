VERSION 5.00
Begin VB.UserControl MouseInfo 
   CanGetFocus     =   0   'False
   ClientHeight    =   480
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   480
   ClipControls    =   0   'False
   InvisibleAtRuntime=   -1  'True
   Picture         =   "MouseInfo.ctx":0000
   PropertyPages   =   "MouseInfo.ctx":0442
   ScaleHeight     =   480
   ScaleMode       =   0  'User
   ScaleWidth      =   48
   ToolboxBitmap   =   "MouseInfo.ctx":045D
   Windowless      =   -1  'True
End
Attribute VB_Name = "MouseInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "PropPageWizardRun" ,"Yes"
Option Explicit

Private Const MOUSEEVENTF_LEFTDOWN = &H2
Private Const MOUSEEVENTF_LEFTUP = &H4
Private Const MOUSEEVENTF_MIDDLEDOWN = &H20
Private Const MOUSEEVENTF_MIDDLEUP = &H40
Private Const MOUSEEVENTF_RIGHTDOWN = &H8
Private Const MOUSEEVENTF_RIGHTUP = &H10

Private Const SM_CMOUSEBUTTONS = 43

'Valori predefiniti proprietà:
Private Const m_def_DblClickTime = 0
Private Const m_def_MouseLimit = False
Private Const m_def_X = 0
Private Const m_def_Y = 0

'Variabili proprietà:
Dim m_Control As Object
Private m_DblClickTime As Long
Private m_MouseLimit As Boolean
Private m_X As Long
Private m_Y As Long

Private Type POINTAPI
       x As Long
       y As Long
End Type

Private Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type

Private Point As POINTAPI

Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function SetCursorPos Lib "user32" (ByVal x As Long, ByVal y As Long) As Long
Private Declare Function ClipCursor Lib "user32" (lpRect As Any) As Long
Private Declare Function GetClientRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Private Declare Function ClientToScreen Lib "user32" (ByVal hwnd As Long, lpPoint As POINTAPI) As Long
Private Declare Function OffsetRect Lib "user32" (lpRect As RECT, ByVal x As Long, ByVal y As Long) As Long
Private Declare Function GetDoubleClickTime Lib "user32" () As Long
Private Declare Function SetDoubleClickTime Lib "user32" (ByVal wCount As Long) As Long
Private Declare Sub mouse_event Lib "user32" (ByVal dwFlags As Long, ByVal dx As Long, ByVal dy As Long, ByVal cButtons As Long, ByVal dwExtraInfo As Long)

Private Sub LimitMovement(Obj As Object, Enable As Boolean)
    Dim client As RECT
    Dim uleft As POINTAPI
    
    On Local Error Resume Next
    If Obj.hwnd > 0 Then
        If Enable Then
            Call GetClientRect(Obj.hwnd, client)
            With uleft
                .x = client.Left
                .y = client.Top
            End With
            Call ClientToScreen(Obj.hwnd, uleft)
            Call OffsetRect(client, uleft.x, uleft.y)
            Call ClipCursor(client)
            Call SetCursorPos(uleft.x, uleft.y)
        Else
            'Releases the cursor limits
            ClipCursor ByVal 0&
        End If
    End If
End Sub

Private Sub GetMousePos()
    Call GetCursorPos(Point)
    With Point
        m_X = .x
        m_Y = .y
    End With
End Sub

Public Sub About()
Attribute About.VB_Description = "Shwo about box."
Attribute About.VB_UserMemId = -552
    frmSplash.Show vbModal
End Sub

Public Property Get x() As Long
Attribute x.VB_Description = "Return or set cursor mouse horizzontal position."
Attribute x.VB_ProcData.VB_Invoke_Property = "Cursor"
    Call GetMousePos
    x = m_X
End Property

Public Property Let x(ByVal New_X As Long)
    m_X = New_X
    Call SetCursorPos(m_X, m_Y)
    PropertyChanged "X"
End Property

Public Property Get y() As Long
Attribute y.VB_Description = "Return or set cursor mouse vertical position."
Attribute y.VB_ProcData.VB_Invoke_Property = "Cursor"
    Call GetMousePos
    y = m_Y
End Property

Public Property Let y(ByVal New_Y As Long)
    m_Y = New_Y
    Call SetCursorPos(m_X, m_Y)
    PropertyChanged "Y"
End Property

Public Property Get Buttons() As Long
Attribute Buttons.VB_Description = "Return mouse buttons available."
    Buttons = GetSystemMetrics(SM_CMOUSEBUTTONS)
End Property

Private Sub UserControl_Initialize()
    #If SHAREWARE = 1 Then
        frmSplash.Show vbModal
    #End If
End Sub

'Inizializza le proprietà di UserControl
Private Sub UserControl_InitProperties()
    m_X = m_def_X
    m_Y = m_def_Y
    m_MouseLimit = m_def_MouseLimit
    m_DblClickTime = m_def_DblClickTime
End Sub

'Carica i valori della proprietà dalla memoria
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    m_X = PropBag.ReadProperty("X", m_def_X)
    m_Y = PropBag.ReadProperty("Y", m_def_Y)
    m_MouseLimit = PropBag.ReadProperty("MouseLimit", m_def_MouseLimit)
    m_DblClickTime = PropBag.ReadProperty("DblClickTime", m_def_DblClickTime)
    Set m_Control = PropBag.ReadProperty("Control", Nothing)
End Sub

Private Sub UserControl_Resize()
    On Local Error Resume Next
    Width = 480
    Height = 480
End Sub

'Scrive i valori della proprietà in memoria
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("X", m_X, m_def_X)
    Call PropBag.WriteProperty("Y", m_Y, m_def_Y)
    Call PropBag.WriteProperty("MouseLimit", m_MouseLimit, m_def_MouseLimit)
    Call PropBag.WriteProperty("DblClickTime", m_DblClickTime, m_def_DblClickTime)
    Call PropBag.WriteProperty("Control", m_Control, Nothing)
End Sub

'AVVISO: NON RIMUOVERE O MODIFICARE LE SEGUENTI RIGHE DI COMMENTO
'MemberInfo=0,0,1,False
Public Property Get MouseLimit() As Boolean
Attribute MouseLimit.VB_Description = "Enable or disable mouse limatate movements."
Attribute MouseLimit.VB_ProcData.VB_Invoke_Property = "Cursor"
    MouseLimit = m_MouseLimit
End Property

Public Property Let MouseLimit(ByVal New_MouseLimit As Boolean)
    m_MouseLimit = New_MouseLimit
    If m_MouseLimit And m_Control Is Nothing Then Set m_Control = Parent
    Call LimitMovement(m_Control, m_MouseLimit)
    PropertyChanged "MouseLimit"
End Property

'AVVISO: NON RIMUOVERE O MODIFICARE LE SEGUENTI RIGHE DI COMMENTO
'MemberInfo=8,0,0,0
Public Property Get DblClickTime() As Long
Attribute DblClickTime.VB_Description = "Return or set the double click interval."
Attribute DblClickTime.VB_ProcData.VB_Invoke_Property = "Cursor"
    m_DblClickTime = GetDoubleClickTime
    DblClickTime = m_DblClickTime
End Property

Public Property Let DblClickTime(ByVal New_DblClickTime As Long)
    m_DblClickTime = New_DblClickTime
    Call SetDoubleClickTime(m_DblClickTime)
    PropertyChanged "DblClickTime"
End Property

'AVVISO: NON RIMUOVERE O MODIFICARE LE SEGUENTI RIGHE DI COMMENTO
'MemberInfo=9,0,0,0
Public Property Get Control() As Object
Attribute Control.VB_Description = "Return or set  object with limit the mouse movements."
    Set Control = m_Control
End Property

Public Property Set Control(ByVal New_Control As Object)
    Set m_Control = New_Control
    PropertyChanged "Control"
End Property

'AVVISO: NON RIMUOVERE O MODIFICARE LE SEGUENTI RIGHE DI COMMENTO
'MemberInfo=14
Public Sub Click(Button As MouseButtonConstants, Optional x As Long = 0, Optional y As Long = 0)
Attribute Click.VB_Description = "Generate a mouse click."
    Dim UpFlag As Long
    Dim DownFlag As Long
    Select Case Button
        Case vbLeftButton
            UpFlag = MOUSEEVENTF_LEFTUP
            DownFlag = MOUSEEVENTF_LEFTDOWN
        Case vbMiddleButton
            UpFlag = MOUSEEVENTF_MIDDLEUP
            DownFlag = MOUSEEVENTF_MIDDLEDOWN
        Case vbRightButton
            UpFlag = MOUSEEVENTF_RIGHTUP
            DownFlag = MOUSEEVENTF_RIGHTDOWN
    End Select
    Call mouse_event(DownFlag, x, y, 0, 0)
    Call mouse_event(UpFlag, x, y, 0, 0)
End Sub

