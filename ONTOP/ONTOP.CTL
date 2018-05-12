VERSION 5.00
Begin VB.UserControl WindowSet 
   CanGetFocus     =   0   'False
   ClientHeight    =   1080
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1065
   ClipControls    =   0   'False
   InvisibleAtRuntime=   -1  'True
   Picture         =   "OnTop.ctx":0000
   PropertyPages   =   "OnTop.ctx":030A
   ScaleHeight     =   1080
   ScaleWidth      =   1065
   ToolboxBitmap   =   "OnTop.ctx":0351
   Begin VB.Timer tmrFlash 
      Left            =   540
      Top             =   420
   End
End
Attribute VB_Name = "WindowSet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "PropPageWizardRun" ,"Yes"
Option Explicit

Public Enum EffectClass
    wsExplode = 1
    wsUpDown = 2
    wsLeftRight = 3
    wsTopBottom = 4
End Enum

'Valori predefiniti proprietà:
Const m_def_Effect = 1
Const m_def_Steps = 100
Const m_def_Center = False
Const m_def_Value = False
Const m_def_RemoveClose = False
Const m_def_Flash = False
Const m_def_Interval = 1000
Const m_def_CloseCaption = "&Close"

'Variabili proprietà:
Dim m_Effect As EffectClass
Dim m_Steps As Long
Dim m_Center As Boolean
Dim m_Value As Boolean
Dim m_RemoveClose As Boolean
Dim m_Flash As Boolean
Dim m_Interval As Long
Dim m_CloseCaption As String

'Varibili interne
Dim bFlash As Boolean

'tipi per le API
Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

' costanti per le API
Const HWND_TOPMOST = -1
Const HWND_NOTOPMOST = -2
Const SWP_DRAWFRAME = &H20
Const SWP_NOMOVE = &H2
Const SWP_NOSIZE = &H1
Const SWP_NOZORDER = &H4
Const SWP_NORMAL = 1
Const SWP_HIDEWINDOW = &H80
Const SWP_SHOWWINDOW = &H40
Const SC_CLOSE = &HF060&
Const MF_END = &H80
Const MF_SEPARATOR = &H800&
Const MF_BYPOSITION = &H400&
Const WM_NCACTIVATE As Long = &H86

Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare Function FlashWindow Lib "user32" (ByVal hwnd As Long, ByVal bInvert As Long) As Long
Private Declare Function GetActiveWindow Lib "user32" () As Long
Private Declare Function RemoveMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
Private Declare Function AppendMenu Lib "user32" Alias "AppendMenuA" (ByVal hMenu As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, ByVal lpNewItem As Any) As Long
Private Declare Function GetSystemMenu Lib "user32" (ByVal hwnd As Long, ByVal bRevert As Long) As Long
Private Declare Function GetMenuItemCount Lib "user32" (ByVal hMenu As Long) As Long
Private Declare Function GetMenuItemID Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Private Declare Function GetMenuString Lib "user32" Alias "GetMenuStringA" (ByVal hMenu As Long, ByVal wIDItem As Long, ByVal lpString As String, ByVal nMaxCount As Long, ByVal wFlag As Long) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Private Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hdc As Long) As Long
Private Declare Function Rectangle Lib "gdi32" (ByVal hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long

Private Sub Effects()
    Dim x As Long
    Dim y As Long
    Dim i As Long
    Dim cx As Long
    Dim cy As Long
    Dim Brush As Long
    Dim myRect As RECT
    Dim TheScreen As Long
    Dim formWidth As Long
    Dim formHeight As Long
    
    Call GetWindowRect(Parent.hwnd, myRect)
    formWidth = (myRect.Right - myRect.Left)
    formHeight = myRect.Bottom - myRect.Top
    TheScreen = GetDC(0)
    Brush = CreateSolidBrush(UserControl.FillColor)
    Call SelectObject(TheScreen, Brush)
    For i = 1 To m_Steps
        With myRect
            Select Case m_Effect
                Case wsUpDown
                    cx = formWidth
                    cy = formHeight * (i / m_Steps)
                    x = .Left
                    y = .Top
                Case wsLeftRight
                    cx = formWidth * (i / m_Steps)
                    cy = formHeight
                    x = .Left
                    y = .Top
                Case wsTopBottom
                    cx = formWidth * (i / m_Steps)
                    cy = formHeight * (i / m_Steps)
                    x = .Left
                    y = .Top
                Case wsExplode
                    cx = formWidth * (i / m_Steps)
                    cy = formHeight * (i / m_Steps)
                    x = .Left + (formWidth - cx) * 0.5
                    y = .Top + (formHeight - cy) * 0.5
            End Select
        End With
        Rectangle TheScreen, x, y, x + cx, y + cy
    Next
    Call ReleaseDC(0, TheScreen)
    Call DeleteObject(Brush)
End Sub

Private Sub RemoveFormClose(bRemove As Boolean)
    Dim i As Long
    Dim hmnu As Long
    Dim iditm As Long
    Static mnuStr As String
    
    ' ricava l'handle del menu di sistema del form
    hmnu = GetSystemMenu(Parent.hwnd, 0)
    
    
    If bRemove Then
        ' scorre le voci del menu
        For i = 0 To GetMenuItemCount(hmnu) - 1
            ' cerca quella che corrisponde a Close
            iditm = GetMenuItemID(hmnu, i)
            If iditm = SC_CLOSE Then
                ' memorizza la sua descrizione
                mnuStr = String(127, 0)
                Call GetMenuString(hmnu, iditm, mnuStr, Len(mnuStr), MF_BYPOSITION)
                ' la rimuove
                Call RemoveMenu(hmnu, i, MF_BYPOSITION)
                ' rimuova anche la barra di separazione
                Call RemoveMenu(hmnu, i - 1, MF_BYPOSITION)
                Exit For
            End If
        Next
    Else
        If Len(mnuStr) > 0 Then
            If Asc(Mid$(mnuStr, 1, 1)) = 0 Then mnuStr = m_CloseCaption & vbNullChar
        Else
            mnuStr = m_CloseCaption
        End If
        AppendMenu hmnu, MF_SEPARATOR, &H0, vbNullString
        AppendMenu hmnu, MF_END, SC_CLOSE, mnuStr
    End If
    Call SendMessage(Parent.hwnd, WM_NCACTIVATE, True, 0)
End Sub

Public Sub About()
Attribute About.VB_Description = "Show about box."
Attribute About.VB_UserMemId = -552
    frmSplash.Show vbModal
End Sub


Public Property Get Flash() As Boolean
Attribute Flash.VB_Description = "Return or set title bar flashing state."
Attribute Flash.VB_ProcData.VB_Invoke_Property = "VisualEffects"
    Flash = m_Flash
End Property

Public Property Let Flash(ByVal New_Value As Boolean)
    m_Flash = New_Value
    If m_Flash Then
        tmrFlash.Interval = m_Interval
    Else
        bFlash = False
        tmrFlash.Interval = 0
        If GetActiveWindow = Parent.hwnd Then Call FlashWindow(Parent.hwnd, False)
    End If
    PropertyChanged "Flash"
End Property

Public Property Get Interval() As Long
Attribute Interval.VB_Description = "Return or set title bar flashing interval."
Attribute Interval.VB_ProcData.VB_Invoke_Property = "VisualEffects"
    Interval = m_Interval
End Property

Public Property Let Interval(ByVal New_Value As Long)
    If New_Value >= 0 And New_Value <= 65535 Then
        m_Interval = New_Value
        If tmrFlash.Interval > 0 Then tmrFlash.Interval = m_Interval
        PropertyChanged "Interval"
    End If
End Property

Public Property Get Value() As Boolean
Attribute Value.VB_Description = "Return or set form on top state."
Attribute Value.VB_ProcData.VB_Invoke_Property = "OnTop"
    Value = m_Value
End Property

Public Property Let Value(ByVal New_Value As Boolean)
    m_Value = New_Value
    SetWindowPos Parent.hwnd, IIf(m_Value, HWND_TOPMOST, HWND_NOTOPMOST), 0, 0, 0, 0, SWP_NORMAL Or SWP_NOSIZE Or SWP_NOMOVE Or SWP_DRAWFRAME
    PropertyChanged "Value"
End Property

Public Property Get RemoveClose() As Boolean
Attribute RemoveClose.VB_Description = "Remove close system menu item."
Attribute RemoveClose.VB_ProcData.VB_Invoke_Property = "SystemMenu"
    RemoveClose = m_RemoveClose
End Property

Public Property Let RemoveClose(ByVal New_RemoveClose As Boolean)
    m_RemoveClose = New_RemoveClose
    Call RemoveFormClose(m_RemoveClose)
    PropertyChanged "RemoveClose"
End Property

Public Property Get CloseCaption() As String
Attribute CloseCaption.VB_Description = "Return or set close system menu item caption."
Attribute CloseCaption.VB_ProcData.VB_Invoke_Property = "SystemMenu"
    CloseCaption = m_CloseCaption
End Property

Public Property Let CloseCaption(ByVal New_Value As String)
    m_CloseCaption = New_Value
    PropertyChanged "CloseCaption"
End Property

Private Sub tmrFlash_Timer()
    bFlash = Not bFlash
    Call FlashWindow(Parent.hwnd, bFlash)
End Sub

Private Sub UserControl_Initialize()
    #If SHAREWARE = 1 Then
        frmSplash.Show vbModal
    #End If
End Sub

'Inizializza le proprietà di UserControl
Private Sub UserControl_InitProperties()
    bFlash = False
    
    m_Value = m_def_Value
    m_RemoveClose = m_def_RemoveClose
    m_Center = m_def_Center
    m_Flash = m_def_Flash
    m_Interval = m_def_Interval
    m_CloseCaption = m_def_CloseCaption
    m_Effect = m_def_Effect
    m_Steps = m_def_Steps

    UserControl.Width = 480
    UserControl.Height = 480
    UserControl.FillColor = &HC0C0C0
End Sub

'Carica i valori della proprietà dalla memoria
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    m_Value = PropBag.ReadProperty("Value", m_def_Value)
    m_RemoveClose = PropBag.ReadProperty("RemoveClose", m_def_RemoveClose)
    m_Center = PropBag.ReadProperty("Center", m_def_Center)
    m_Flash = PropBag.ReadProperty("Flash", m_def_Flash)
    m_Interval = PropBag.ReadProperty("Interval", m_def_Interval)
    m_CloseCaption = PropBag.ReadProperty("CloseCaption", m_def_CloseCaption)
    m_Effect = PropBag.ReadProperty("Effect", m_def_Effect)
    m_Steps = PropBag.ReadProperty("Steps", m_def_Steps)
    UserControl.FillColor = PropBag.ReadProperty("FillColor", &HC0C0C0)
End Sub

Private Sub UserControl_Resize()
    On Local Error Resume Next
    Width = 480
    Height = 480
End Sub

'Scrive i valori della proprietà in memoria
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("Value", m_Value, m_def_Value)
    Call PropBag.WriteProperty("RemoveClose", m_RemoveClose, m_def_RemoveClose)
    Call PropBag.WriteProperty("Center", m_Center, m_def_Center)
    Call PropBag.WriteProperty("Flash", m_Flash, m_def_Flash)
    Call PropBag.WriteProperty("Interval", m_Interval, m_def_Interval)
    Call PropBag.WriteProperty("CloseCaption", m_CloseCaption, m_def_CloseCaption)
    Call PropBag.WriteProperty("Effect", m_Effect, m_def_Effect)
    Call PropBag.WriteProperty("Steps", m_Steps, m_def_Steps)
    Call PropBag.WriteProperty("FillColor", UserControl.FillColor, &HC0C0C0)
End Sub

Public Property Get Center() As Boolean
Attribute Center.VB_Description = "Center current form."
    Center = m_Center
End Property

Public Property Let Center(ByVal New_Center As Boolean)
    m_Center = New_Center
    If m_Center Then Parent.Move (Screen.Width - Parent.Width) * 0.5, (Screen.Height - Parent.Height) * 0.5
    PropertyChanged "Center"
End Property

'AVVISO: NON RIMUOVERE O MODIFICARE LE SEGUENTI RIGHE DI COMMENTO
'MemberInfo=8,0,0,0
Public Property Get Effect() As EffectClass
Attribute Effect.VB_Description = "Return or set show effect type."
    Effect = m_Effect
End Property

Public Property Let Effect(ByVal New_Effect As EffectClass)
    m_Effect = New_Effect
    PropertyChanged "Effect"
End Property

'AVVISO: NON RIMUOVERE O MODIFICARE LE SEGUENTI RIGHE DI COMMENTO
'MemberInfo=14,0,0,0
Public Property Get Steps() As Long
Attribute Steps.VB_Description = "Return or set show effects steps."
Attribute Steps.VB_ProcData.VB_Invoke_Property = "VisualEffects"
    Steps = m_Steps
End Property

Public Property Let Steps(ByVal New_Steps As Long)
    m_Steps = New_Steps
    PropertyChanged "Steps"
End Property

'AVVISO: NON RIMUOVERE O MODIFICARE LE SEGUENTI RIGHE DI COMMENTO
'MemberInfo=8
Public Sub Show()
Attribute Show.VB_Description = "Show form with effects."
    On Local Error Resume Next
    Call Effects
    Parent.Show
End Sub

'AVVISO: NON RIMUOVERE O MODIFICARE LE SEGUENTI RIGHE DI COMMENTO
'MappingInfo=UserControl,UserControl,-1,FillColor
Public Property Get FillColor() As OLE_COLOR
Attribute FillColor.VB_Description = "Restituisce o imposta il colore utilizzato per applicare riempimenti a forme, cerchi e caselle."
    FillColor = UserControl.FillColor
End Property

Public Property Let FillColor(ByVal New_FillColor As OLE_COLOR)
    UserControl.FillColor() = New_FillColor
    PropertyChanged "FillColor"
End Property

