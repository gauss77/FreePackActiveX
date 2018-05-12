VERSION 5.00
Begin VB.UserControl PictureTile 
   CanGetFocus     =   0   'False
   ClientHeight    =   1500
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1515
   ClipControls    =   0   'False
   InvisibleAtRuntime=   -1  'True
   Picture         =   "PicTile.ctx":0000
   PropertyPages   =   "PicTile.ctx":098A
   ScaleHeight     =   1500
   ScaleWidth      =   1515
   ToolboxBitmap   =   "PicTile.ctx":09AE
End
Attribute VB_Name = "PictureTile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "PropPageWizardRun" ,"Yes"
Option Explicit
'Valori predefiniti proprietà:
Const m_def_StartY = 0
'Variabili proprietà:
Dim m_StartY As Long

Private Sub TitleBackground(DrawObj As Object, Src As Object, StartY As Long)
    Dim foo As Integer
    Dim foobar As Integer
    
    DrawObj.AutoRedraw = True
    For foo = 0 To DrawObj.ScaleWidth Step Src.ScaleWidth - Screen.TwipsPerPixelX
        For foobar = StartY To DrawObj.ScaleHeight Step Src.ScaleHeight
            DrawObj.PaintPicture Src.Picture, foo, foobar, Src.ScaleWidth, Src.ScaleHeight
        Next
    Next
    DrawObj.Line (0, StartY)-(DrawObj.Width, StartY), vbButtonShadow
    DrawObj.Line (0, StartY + 15)-(DrawObj.Width, StartY + 15), vb3DHighlight
    DrawObj.AutoRedraw = False
End Sub

Public Sub About()
Attribute About.VB_Description = "Show about box."
Attribute About.VB_UserMemId = -552
    frmSplash.Show vbModal
End Sub

Public Sub Refresh()
Attribute Refresh.VB_Description = "Run the paint picture of entire form."
    Call TitleBackground(Parent, Me, m_StartY)
End Sub

'AVVISO: NON RIMUOVERE O MODIFICARE LE SEGUENTI RIGHE DI COMMENTO
'MappingInfo=UserControl,UserControl,-1,hDC
Public Property Get hDC() As Long
Attribute hDC.VB_Description = "Restituisce un handle fornito in Microsoft Windows al contesto di periferica di un oggetto."
    hDC = UserControl.hDC
End Property

'AVVISO: NON RIMUOVERE O MODIFICARE LE SEGUENTI RIGHE DI COMMENTO
'MappingInfo=UserControl,UserControl,-1,hWnd
Public Property Get hWnd() As Long
Attribute hWnd.VB_Description = "Restituisce un handle (da Microsoft Windows) alla finestra di un oggetto."
Attribute hWnd.VB_UserMemId = -515
    hWnd = UserControl.hWnd
End Property

'AVVISO: NON RIMUOVERE O MODIFICARE LE SEGUENTI RIGHE DI COMMENTO
'MappingInfo=UserControl,UserControl,-1,Picture
Public Property Get Picture() As Picture
Attribute Picture.VB_Description = "Restituisce o imposta un elemento grafico da visualizzare in un controllo."
    Set Picture = UserControl.Picture
End Property

Public Property Set Picture(ByVal New_Picture As Picture)
    Set UserControl.Picture = New_Picture
    PropertyChanged "Picture"
End Property

'AVVISO: NON RIMUOVERE O MODIFICARE LE SEGUENTI RIGHE DI COMMENTO
'MappingInfo=UserControl,UserControl,-1,ScaleHeight
Public Property Get ScaleHeight() As Single
Attribute ScaleHeight.VB_Description = "Restituisce o imposta il numero di unità per la misurazione verticale dell'area interna di un oggetto."
    ScaleHeight = UserControl.ScaleHeight
End Property

Public Property Let ScaleHeight(ByVal New_ScaleHeight As Single)
    UserControl.ScaleHeight() = New_ScaleHeight
    PropertyChanged "ScaleHeight"
End Property

'AVVISO: NON RIMUOVERE O MODIFICARE LE SEGUENTI RIGHE DI COMMENTO
'MappingInfo=UserControl,UserControl,-1,ScaleLeft
Public Property Get ScaleLeft() As Single
Attribute ScaleLeft.VB_Description = "Restituisce o imposta le coordinate orizzontali per il bordo sinistro di un oggetto."
    ScaleLeft = UserControl.ScaleLeft
End Property

Public Property Let ScaleLeft(ByVal New_ScaleLeft As Single)
    UserControl.ScaleLeft() = New_ScaleLeft
    PropertyChanged "ScaleLeft"
End Property

'AVVISO: NON RIMUOVERE O MODIFICARE LE SEGUENTI RIGHE DI COMMENTO
'MappingInfo=UserControl,UserControl,-1,ScaleMode
Public Property Get ScaleMode() As Integer
Attribute ScaleMode.VB_Description = "Restituisce o imposta un valore che indica le unità di misura per le coordinate di un oggetto quando si utilizzano metodi grafici o si posizionano controlli."
    ScaleMode = UserControl.ScaleMode
End Property

Public Property Let ScaleMode(ByVal New_ScaleMode As Integer)
    UserControl.ScaleMode() = New_ScaleMode
    PropertyChanged "ScaleMode"
End Property

'AVVISO: NON RIMUOVERE O MODIFICARE LE SEGUENTI RIGHE DI COMMENTO
'MappingInfo=UserControl,UserControl,-1,ScaleTop
Public Property Get ScaleTop() As Single
Attribute ScaleTop.VB_Description = "Restituisce o imposta le coordinate verticali per il bordo superiore di un oggetto."
    ScaleTop = UserControl.ScaleTop
End Property

Public Property Let ScaleTop(ByVal New_ScaleTop As Single)
    UserControl.ScaleTop() = New_ScaleTop
    PropertyChanged "ScaleTop"
End Property

'AVVISO: NON RIMUOVERE O MODIFICARE LE SEGUENTI RIGHE DI COMMENTO
'MappingInfo=UserControl,UserControl,-1,ScaleWidth
Public Property Get ScaleWidth() As Single
Attribute ScaleWidth.VB_Description = "Restituisce o imposta il numero di unità per la misurazione orizzontale dell'area interna di un oggetto."
    ScaleWidth = UserControl.ScaleWidth
End Property

Public Property Let ScaleWidth(ByVal New_ScaleWidth As Single)
    UserControl.ScaleWidth() = New_ScaleWidth
    PropertyChanged "ScaleWidth"
End Property

Private Sub UserControl_Initialize()
    #If SHAREWARE = 1 Then
        frmSplash.Show vbModal
    #End If
End Sub

'Carica i valori della proprietà dalla memoria
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    Set Picture = PropBag.ReadProperty("Picture", Nothing)
    UserControl.ScaleHeight = PropBag.ReadProperty("ScaleHeight", 1065)
    UserControl.ScaleLeft = PropBag.ReadProperty("ScaleLeft", 0)
    UserControl.ScaleMode = PropBag.ReadProperty("ScaleMode", 1)
    UserControl.ScaleTop = PropBag.ReadProperty("ScaleTop", 0)
    UserControl.ScaleWidth = PropBag.ReadProperty("ScaleWidth", 1920)
    m_StartY = PropBag.ReadProperty("StartY", m_def_StartY)
End Sub

'Scrive i valori della proprietà in memoria
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("Picture", Picture, Nothing)
    Call PropBag.WriteProperty("ScaleHeight", UserControl.ScaleHeight, 1065)
    Call PropBag.WriteProperty("ScaleLeft", UserControl.ScaleLeft, 0)
    Call PropBag.WriteProperty("ScaleMode", UserControl.ScaleMode, 1)
    Call PropBag.WriteProperty("ScaleTop", UserControl.ScaleTop, 0)
    Call PropBag.WriteProperty("ScaleWidth", UserControl.ScaleWidth, 1920)
    Call PropBag.WriteProperty("StartY", m_StartY, m_def_StartY)
End Sub

Public Property Get StartY() As Long
Attribute StartY.VB_Description = "Return or set start vertical position."
    StartY = m_StartY
End Property

Public Property Let StartY(ByVal New_StartY As Long)
    m_StartY = New_StartY
    PropertyChanged "StartY"
End Property

'Inizializza le proprietà di UserControl
Private Sub UserControl_InitProperties()
    m_StartY = m_def_StartY
End Sub

