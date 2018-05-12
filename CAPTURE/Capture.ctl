VERSION 5.00
Begin VB.UserControl Capture 
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   PropertyPages   =   "Capture.ctx":0000
   ScaleHeight     =   3600
   ScaleWidth      =   4800
   ToolboxBitmap   =   "Capture.ctx":0043
End
Attribute VB_Name = "Capture"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "PropPageWizardRun" ,"Yes"
Option Explicit

'--------------------------------------------------------------------
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' Visual Basic 4.0 16/32 Capture Routines
'
' This module contains several routines for capturing windows into a
' picture. All the routines work on both 16 and 32 bit Windows
' platforms.
' The routines also have palette support.
'
' CreateBitmapPicture - Creates a picture object from a bitmap and
' palette
' CaptureWindow - Captures any window given a window handle
' CaptureActiveWindow - Captures the active window on the desktop
' CaptureForm - Captures the entire form
' CaptureClient - Captures the client area of a form
' CaptureScreen - Captures the entire screen
' PrintPictureToFitPage - prints any picture as big as possible on
' the page
'
' NOTES
' - No error trapping is included in these routines
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
Option Base 0

Public Enum ApparenceClass
    vbFlat = 0
    vb3D = 1
End Enum

Public Enum BackStyleClass
    vbTransparent = 0
    vbOpaque = 1
End Enum

Public Enum BorderStyleClass
    vbNone = 0
    vbFixedSingle
End Enum

Public Enum CaptureClass
    vbActiveWindow = 0
    vbClient = 1
    vbForm = 2
    vbScreen = 3
End Enum

Private Type PALETTEENTRY
    peRed As Byte
    peGreen As Byte
    peBlue As Byte
    peFlags As Byte
End Type

Private Type LOGPALETTE
    palVersion As Integer
    palNumEntries As Integer
    palPalEntry(255) As PALETTEENTRY ' Enough for 256 colors
End Type

Private Type GUID
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(7) As Byte
End Type

Private Const RASTERCAPS As Long = 38
Private Const RC_PALETTE As Long = &H100
Private Const SIZEPALETTE As Long = 104

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private Declare Function CreateCompatibleDC Lib "GDI32" (ByVal hDC As Long) As Long
Private Declare Function CreateCompatibleBitmap Lib "GDI32" (ByVal hDC As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function GetDeviceCaps Lib "GDI32" (ByVal hDC As Long, ByVal iCapabilitiy As Long) As Long
Private Declare Function GetSystemPaletteEntries Lib "GDI32" (ByVal hDC As Long, ByVal wStartIndex As Long, ByVal wNumEntries As Long, lpPaletteEntries As PALETTEENTRY) As Long
Private Declare Function CreatePalette Lib "GDI32" (lpLogPalette As LOGPALETTE) As Long
Private Declare Function SelectObject Lib "GDI32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Private Declare Function BitBlt Lib "GDI32" (ByVal hDCDest As Long, ByVal XDest As Long, ByVal YDest As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hDCSrc As Long, ByVal XSrc As Long, ByVal YSrc As Long, ByVal dwRop As Long) As Long
Private Declare Function DeleteDC Lib "GDI32" (ByVal hDC As Long) As Long
Private Declare Function GetForegroundWindow Lib "USER32" () As Long
Private Declare Function SelectPalette Lib "GDI32" (ByVal hDC As Long, ByVal hPalette As Long, ByVal bForceBackground As Long) As Long
Private Declare Function RealizePalette Lib "GDI32" (ByVal hDC As Long) As Long
Private Declare Function GetWindowDC Lib "USER32" (ByVal hwnd As Long) As Long
Private Declare Function GetDC Lib "USER32" (ByVal hwnd As Long) As Long
Private Declare Function GetWindowRect Lib "USER32" (ByVal hwnd As Long, lpRect As RECT) As Long
Private Declare Function ReleaseDC Lib "USER32" (ByVal hwnd As Long, ByVal hDC As Long) As Long
Private Declare Function GetDesktopWindow Lib "USER32" () As Long

Private Type PicBmp
    Size As Long
    Type As Long
    hBmp As Long
    hPal As Long
    Reserved As Long
End Type

Private Declare Function OleCreatePictureIndirect Lib "olepro32.dll" (PicDesc As PicBmp, RefIID As GUID, ByVal fPictureOwnsHandle As Long, IPic As IPicture) As Long

'Valori predefiniti proprietà:
Const m_def_AutoSize = False
Const m_def_CaptureType = 3

'Variabili proprietà:
Dim m_AutoSize As Boolean
Dim m_CaptureType As CaptureClass

'Dichiarazioni di eventi:
Event Click() 'MappingInfo=UserControl,UserControl,-1,Click
Attribute Click.VB_Description = "Viene generato quando si preme e quindi si rilascia un pulsante del mouse su un oggetto."
Attribute Click.VB_UserMemId = -600
Event DblClick() 'MappingInfo=UserControl,UserControl,-1,DblClick
Attribute DblClick.VB_Description = "Viene generato quando si preme e si rilascia due volte in rapida successione un pulsante del mouse su un oggetto."
Attribute DblClick.VB_UserMemId = -601
Event HitTest(X As Single, Y As Single, HitResult As Integer) 'MappingInfo=UserControl,UserControl,-1,HitTest
Attribute HitTest.VB_Description = "Viene generato in un controllo utente privo di finestra in risposta alle operazioni con il mouse."
Event KeyDown(KeyCode As Integer, Shift As Integer) 'MappingInfo=UserControl,UserControl,-1,KeyDown
Attribute KeyDown.VB_Description = "Viene generato quando si preme un tasto mentre lo stato attivo si trova su un oggetto."
Attribute KeyDown.VB_UserMemId = -602
Event KeyPress(KeyAscii As Integer) 'MappingInfo=UserControl,UserControl,-1,KeyPress
Attribute KeyPress.VB_Description = "Viene generato quando si preme e si rilascia un tasto ANSI."
Attribute KeyPress.VB_UserMemId = -603
Event KeyUp(KeyCode As Integer, Shift As Integer) 'MappingInfo=UserControl,UserControl,-1,KeyUp
Attribute KeyUp.VB_Description = "Viene generato quando si rilascia un tasto mentre lo stato attivo si trova su un oggetto."
Attribute KeyUp.VB_UserMemId = -604
Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseDown
Attribute MouseDown.VB_Description = "Viene generato quando si preme il pulsante del mouse mentre lo stato attivo si trova su un oggetto."
Attribute MouseDown.VB_UserMemId = -605
Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseMove
Attribute MouseMove.VB_Description = "Viene generato quando si sposta il mouse."
Attribute MouseMove.VB_UserMemId = -606
Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseUp
Attribute MouseUp.VB_Description = "Viene generato quando si rilascia il pulsante del mouse mentre lo stato attivo si trova su un oggetto."
Attribute MouseUp.VB_UserMemId = -607
Event Paint() 'MappingInfo=UserControl,UserControl,-1,Paint
Attribute Paint.VB_Description = "Viene generato quando una parte qualsiasi di un form o di un controllo PictureBox viene spostata, allargata o esposta."
Event Resize() 'MappingInfo=UserControl,UserControl,-1,Resize
Attribute Resize.VB_Description = "Viene generato non appena un form viene visualizzato o quando le dimensioni di un oggetto vengono modificate."



'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' CreateBitmapPicture
' - Creates a bitmap type Picture object from a bitmap and palette
'
' hBmp
' - Handle to a bitmap
'
' hPal
' - Handle to a Palette
' - Can be null if the bitmap doesn't use a palette
'
' Returns
' - Returns a Picture object containing the bitmap
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
Private Function CreateBitmapPicture(ByVal hBmp As Long, ByVal hPal As Long) As Picture

    Dim r As Long
    Dim pic As PicBmp
    ' IPicture requires a reference to "Standard OLE Types"
    Dim IPic As IPicture
    Dim IID_IDispatch As GUID
    
    ' Fill in with IDispatch Interface ID
    With IID_IDispatch
        .Data1 = &H20400
        .Data4(0) = &HC0
        .Data4(7) = &H46
    End With
    
    ' Fill Pic with necessary parts
    With pic
        .Size = Len(pic) ' Length of structure
        .Type = vbPicTypeBitmap ' Type of Picture (bitmap)
        .hBmp = hBmp ' Handle to bitmap
        .hPal = hPal ' Handle to palette (may be null)
    End With
    
    ' Create Picture object
    r = OleCreatePictureIndirect(pic, IID_IDispatch, 1, IPic)
    
    ' Return the new Picture object
    Set CreateBitmapPicture = IPic
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' CaptureWindow
' - Captures any portion of a window
'
' hWndSrc
' - Handle to the window to be captured
'
' Client
' - If True CaptureWindow captures from the client area of the
' window
' - If False CaptureWindow captures from the entire window
'
' LeftSrc, TopSrc, WidthSrc, HeightSrc
' - Specify the portion of the window to capture
' - Dimensions need to be specified in pixels
'
' Returns
' - Returns a Picture object containing a bitmap of the specified
' portion of the window that was captured
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''
'
Private Function CaptureWindow(ByVal hWndSrc As Long, ByVal Client As Boolean, ByVal LeftSrc As Long, ByVal TopSrc As Long, ByVal WidthSrc As Long, ByVal HeightSrc As Long) As Picture
    
    Dim hDCMemory As Long
    Dim hBmp As Long
    Dim hBmpPrev As Long
    Dim r As Long
    Dim hDCSrc As Long
    Dim hPal As Long
    Dim hPalPrev As Long
    Dim RasterCapsScrn As Long
    Dim HasPaletteScrn As Long
    Dim PaletteSizeScrn As Long
    Dim LogPal As LOGPALETTE
    
    ' Depending on the value of Client get the proper device context
    If Client Then
        hDCSrc = GetDC(hWndSrc) ' Get device context for client area
    Else
        hDCSrc = GetWindowDC(hWndSrc) ' Get device context for entire
        ' window
    End If
    
    ' Create a memory device context for the copy process
    hDCMemory = CreateCompatibleDC(hDCSrc)
    ' Create a bitmap and place it in the memory DC
    hBmp = CreateCompatibleBitmap(hDCSrc, WidthSrc, HeightSrc)
    hBmpPrev = SelectObject(hDCMemory, hBmp)
    
    ' Get screen properties
    RasterCapsScrn = GetDeviceCaps(hDCSrc, RASTERCAPS) ' Raster
    'capabilities
    HasPaletteScrn = RasterCapsScrn And RC_PALETTE ' Palette
    'support
    PaletteSizeScrn = GetDeviceCaps(hDCSrc, SIZEPALETTE) ' Size of
    ' palette
    
    ' If the screen has a palette make a copy and realize it
    If HasPaletteScrn And (PaletteSizeScrn = 256) Then
        ' Create a copy of the system palette
        LogPal.palVersion = &H300
        LogPal.palNumEntries = 256
        r = GetSystemPaletteEntries(hDCSrc, 0, 256, LogPal.palPalEntry(0))
        hPal = CreatePalette(LogPal)
        ' Select the new palette into the memory DC and realize it
        hPalPrev = SelectPalette(hDCMemory, hPal, 0)
        r = RealizePalette(hDCMemory)
    End If
    
    ' Copy the on-screen image into the memory DC
    r = BitBlt(hDCMemory, 0, 0, WidthSrc, HeightSrc, hDCSrc, LeftSrc, TopSrc, vbSrcCopy)
    
    ' Remove the new copy of the on-screen image
    hBmp = SelectObject(hDCMemory, hBmpPrev)
    
    ' If the screen has a palette get back the palette that was
    ' selected in previously
    If HasPaletteScrn And (PaletteSizeScrn = 256) Then
        hPal = SelectPalette(hDCMemory, hPalPrev, 0)
    End If
    
    ' Release the device context resources back to the system
    r = DeleteDC(hDCMemory)
    r = ReleaseDC(hWndSrc, hDCSrc)
    
    ' Call CreateBitmapPicture to create a picture object from the
    ' bitmap and palette handles. Then return the resulting picture
    ' object.
    Set CaptureWindow = CreateBitmapPicture(hBmp, hPal)
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' CaptureScreen
' - Captures the entire screen
'
' Returns
' - Returns a Picture object containing a bitmap of the screen
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
Private Function CaptureScreen() As Picture
    Dim hWndScreen As Long

    ' Get a handle to the desktop window
    hWndScreen = GetDesktopWindow()

    ' Call CaptureWindow to capture the entire desktop give the handle
    ' and return the resulting Picture object

    Set CaptureScreen = CaptureWindow(hWndScreen, False, 0, 0, Screen.Width \ Screen.TwipsPerPixelX, Screen.Height \ Screen.TwipsPerPixelY)
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' CaptureForm
' - Captures an entire form including title bar and border
'
' frmSrc
' - The Form object to capture
'
' Returns
' - Returns a Picture object containing a bitmap of the entire
' form
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
Private Function CaptureForm(frmSrc As Form) As Picture
    ' Call CaptureWindow to capture the entire form given it's window
    ' handle and then return the resulting Picture object
    Set CaptureForm = CaptureWindow(frmSrc.hwnd, False, 0, 0, frmSrc.ScaleX(frmSrc.Width, vbTwips, vbPixels), frmSrc.ScaleY(frmSrc.Height, vbTwips, vbPixels))
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' CaptureClient
' - Captures the client area of a form
'
' frmSrc
' - The Form object to capture
'
' Returns
' - Returns a Picture object containing a bitmap of the form's
' client area
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
Private Function CaptureClient(frmSrc As Form) As Picture
    ' Call CaptureWindow to capture the client area of the form given
    ' it's window handle and return the resulting Picture object
    Set CaptureClient = CaptureWindow(frmSrc.hwnd, True, 0, 0, frmSrc.ScaleX(frmSrc.ScaleWidth, frmSrc.ScaleMode, vbPixels), frmSrc.ScaleY(frmSrc.ScaleHeight, frmSrc.ScaleMode, vbPixels))
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' CaptureActiveWindow
' - Captures the currently active window on the screen
'
' Returns
' - Returns a Picture object containing a bitmap of the active
' window
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
Private Function CaptureActiveWindow() As Picture
    Dim hWndActive As Long
    Dim r As Long
    Dim RectActive As RECT

    ' Get a handle to the active/foreground window
    hWndActive = GetForegroundWindow()
    
    ' Get the dimensions of the window
    r = GetWindowRect(hWndActive, RectActive)
    
    ' Call CaptureWindow to capture the active window given it's
    ' handle and return the Resulting Picture object
    Set CaptureActiveWindow = CaptureWindow(hWndActive, False, 0, 0, RectActive.Right - RectActive.Left, RectActive.Bottom - RectActive.Top)
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' PrintPictureToFitPage
' - Prints a Picture object as big as possible
'
' Prn
' - Destination Printer object
'
' Pic
' - Source Picture object
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
Private Sub PrintPictureToFitPage(Prn As Printer, pic As Picture)
    Const vbHiMetric As Integer = 8
    Dim PicRatio As Double
    Dim PrnWidth As Double
    Dim PrnHeight As Double
    Dim PrnRatio As Double
    Dim PrnPicWidth As Double
    Dim PrnPicHeight As Double
    
    ' Determine if picture should be printed in landscape or portrait
    ' and set the orientation
    If pic.Height >= pic.Width Then
        Prn.Orientation = vbPRORPortrait ' Taller than wide
    Else
        Prn.Orientation = vbPRORLandscape ' Wider than tall
    End If
    
    ' Calculate device independent Width to Height ratio for picture
    PicRatio = pic.Width / pic.Height
    
    ' Calculate the dimentions of the printable area in HiMetric
    PrnWidth = Prn.ScaleX(Prn.ScaleWidth, Prn.ScaleMode, vbHiMetric)
    PrnHeight = Prn.ScaleY(Prn.ScaleHeight, Prn.ScaleMode, vbHiMetric)
    ' Calculate device independent Width to Height ratio for printer
    PrnRatio = PrnWidth / PrnHeight
    
    ' Scale the output to the printable area
    If PicRatio >= PrnRatio Then
        ' Scale picture to fit full width of printable area
        PrnPicWidth = Prn.ScaleX(PrnWidth, vbHiMetric, Prn.ScaleMode)
        PrnPicHeight = Prn.ScaleY(PrnWidth / PicRatio, vbHiMetric, Prn.ScaleMode)
    Else
        ' Scale picture to fit full height of printable area
        PrnPicHeight = Prn.ScaleY(PrnHeight, vbHiMetric, Prn.ScaleMode)
        PrnPicWidth = Prn.ScaleX(PrnHeight * PicRatio, vbHiMetric, Prn.ScaleMode)
    End If
    
    ' Print the picture using the PaintPicture method
    Prn.PaintPicture pic, 0, 0, PrnPicWidth, PrnPicHeight
End Sub

Public Sub About()
Attribute About.VB_Description = "Show about box."
Attribute About.VB_UserMemId = -552
    frmSplash.Show vbModal
End Sub


'AVVISO: NON RIMUOVERE O MODIFICARE LE SEGUENTI RIGHE DI COMMENTO
'MemberInfo=8,0,0,0
Public Property Get CaptureType() As CaptureClass
Attribute CaptureType.VB_Description = "Return or set the capture type."
Attribute CaptureType.VB_HelpID = 7
    CaptureType = m_CaptureType
End Property

Public Property Let CaptureType(ByVal New_CaptureType As CaptureClass)
    m_CaptureType = New_CaptureType
    PropertyChanged "CaptureType"
End Property

'Inizializza le proprietà di UserControl.
Private Sub UserControl_InitProperties()
    m_CaptureType = m_def_CaptureType
    Set UserControl.Font = Ambient.Font
    m_AutoSize = m_def_AutoSize
End Sub

'Carica i valori delle proprietà dalla posizione di memorizzazione.
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    m_CaptureType = PropBag.ReadProperty("CaptureType", m_def_CaptureType)
    UserControl.Appearance = PropBag.ReadProperty("Appearance", 1)
    UserControl.AutoRedraw = PropBag.ReadProperty("AutoRedraw", False)
    UserControl.BackColor = PropBag.ReadProperty("BackColor", &H8000000F)
    UserControl.BackStyle = PropBag.ReadProperty("BackStyle", 1)
    UserControl.BorderStyle = PropBag.ReadProperty("BorderStyle", 0)
    UserControl.DrawMode = PropBag.ReadProperty("DrawMode", 13)
    UserControl.DrawStyle = PropBag.ReadProperty("DrawStyle", 0)
    UserControl.DrawWidth = PropBag.ReadProperty("DrawWidth", 1)
    UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
    UserControl.FillColor = PropBag.ReadProperty("FillColor", &H0&)
    UserControl.FillStyle = PropBag.ReadProperty("FillStyle", 1)
    Set UserControl.Font = PropBag.ReadProperty("Font", Ambient.Font)
    UserControl.ForeColor = PropBag.ReadProperty("ForeColor", &H80000012)
    UserControl.MaskColor = PropBag.ReadProperty("MaskColor", -2147483633)
    Set MaskPicture = PropBag.ReadProperty("MaskPicture", Nothing)
    Set MouseIcon = PropBag.ReadProperty("MouseIcon", Nothing)
    UserControl.MousePointer = PropBag.ReadProperty("MousePointer", 0)
    Set Palette = PropBag.ReadProperty("Palette", Nothing)
    UserControl.PaletteMode = PropBag.ReadProperty("PaletteMode", 3)
    Set Picture = PropBag.ReadProperty("Picture", Nothing)
    UserControl.ScaleHeight = PropBag.ReadProperty("ScaleHeight", 3600)
    UserControl.ScaleMode = PropBag.ReadProperty("ScaleMode", 1)
    UserControl.ScaleWidth = PropBag.ReadProperty("ScaleWidth", 4800)
    m_AutoSize = PropBag.ReadProperty("AutoSize", m_def_AutoSize)
    Set MaskPicture = PropBag.ReadProperty("MaskPicture", Nothing)
End Sub

'Scrive i valori delle proprietà nella posizione di memorizzazione.
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("CaptureType", m_CaptureType, m_def_CaptureType)
    Call PropBag.WriteProperty("Appearance", UserControl.Appearance, 1)
    Call PropBag.WriteProperty("AutoRedraw", UserControl.AutoRedraw, False)
    Call PropBag.WriteProperty("BackColor", UserControl.BackColor, &H8000000F)
    Call PropBag.WriteProperty("BackStyle", UserControl.BackStyle, 1)
    Call PropBag.WriteProperty("BorderStyle", UserControl.BorderStyle, 0)
    Call PropBag.WriteProperty("DrawMode", UserControl.DrawMode, 13)
    Call PropBag.WriteProperty("DrawStyle", UserControl.DrawStyle, 0)
    Call PropBag.WriteProperty("DrawWidth", UserControl.DrawWidth, 1)
    Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
    Call PropBag.WriteProperty("FillColor", UserControl.FillColor, &H0&)
    Call PropBag.WriteProperty("FillStyle", UserControl.FillStyle, 1)
    Call PropBag.WriteProperty("Font", UserControl.Font, Ambient.Font)
    Call PropBag.WriteProperty("ForeColor", UserControl.ForeColor, &H80000012)
    Call PropBag.WriteProperty("MaskColor", UserControl.MaskColor, -2147483633)
    Call PropBag.WriteProperty("MaskPicture", MaskPicture, Nothing)
    Call PropBag.WriteProperty("MouseIcon", MouseIcon, Nothing)
    Call PropBag.WriteProperty("MousePointer", UserControl.MousePointer, 0)
    Call PropBag.WriteProperty("Palette", Palette, Nothing)
    Call PropBag.WriteProperty("PaletteMode", UserControl.PaletteMode, 3)
    Call PropBag.WriteProperty("Picture", Picture, Nothing)
    Call PropBag.WriteProperty("ScaleHeight", UserControl.ScaleHeight, 3600)
    Call PropBag.WriteProperty("ScaleMode", UserControl.ScaleMode, 1)
    Call PropBag.WriteProperty("ScaleWidth", UserControl.ScaleWidth, 4800)
    Call PropBag.WriteProperty("AutoSize", m_AutoSize, m_def_AutoSize)
    Call PropBag.WriteProperty("MaskPicture", MaskPicture, Nothing)
End Sub

'AVVISO: NON RIMUOVERE O MODIFICARE LE SEGUENTI RIGHE DI COMMENTO
'MemberInfo=14
Public Sub Capture()
Attribute Capture.VB_Description = "Start a capture."
Attribute Capture.VB_HelpID = 16
    Dim pic As Picture
    
    Select Case m_CaptureType
        Case vbActiveWindow
            Set pic = CaptureActiveWindow
        Case vbClient
            Set pic = CaptureClient(Parent)
        Case vbForm
            Set pic = CaptureForm(Parent)
        Case vbScreen
            Set pic = CaptureScreen
    End Select
    If m_AutoSize Then
        UserControl.Width = pic.Width
        UserControl.Height = pic.Height
        RaiseEvent Resize
    End If
    If pic > 0 Then Set UserControl.Picture = pic
End Sub

'AVVISO: NON RIMUOVERE O MODIFICARE LE SEGUENTI RIGHE DI COMMENTO
'MappingInfo=UserControl,UserControl,-1,Appearance
Public Property Get Appearance() As ApparenceClass
Attribute Appearance.VB_Description = "Restituisce o imposta un valore che indica se un oggetto viene ridisegnato con effetti tridimensionali in fase di esecuzione."
    Appearance = UserControl.Appearance
End Property

Public Property Let Appearance(ByVal New_Appearance As ApparenceClass)
    UserControl.Appearance() = New_Appearance
    PropertyChanged "Appearance"
End Property

'AVVISO: NON RIMUOVERE O MODIFICARE LE SEGUENTI RIGHE DI COMMENTO
'MappingInfo=UserControl,UserControl,-1,AutoRedraw
Public Property Get AutoRedraw() As Boolean
Attribute AutoRedraw.VB_Description = "Restituisce o imposta l'output di un metodo grafico in una bitmap fissa."
    AutoRedraw = UserControl.AutoRedraw
End Property

Public Property Let AutoRedraw(ByVal New_AutoRedraw As Boolean)
    UserControl.AutoRedraw() = New_AutoRedraw
    PropertyChanged "AutoRedraw"
End Property

'AVVISO: NON RIMUOVERE O MODIFICARE LE SEGUENTI RIGHE DI COMMENTO
'MappingInfo=UserControl,UserControl,-1,BackColor
Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Restituisce o imposta il colore di sfondo utilizzato per la visualizzazione di testo e grafica in un oggetto."
Attribute BackColor.VB_UserMemId = -501
    BackColor = UserControl.BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    UserControl.BackColor() = New_BackColor
    PropertyChanged "BackColor"
End Property

'AVVISO: NON RIMUOVERE O MODIFICARE LE SEGUENTI RIGHE DI COMMENTO
'MappingInfo=UserControl,UserControl,-1,BackStyle
Public Property Get BackStyle() As BackStyleClass
Attribute BackStyle.VB_Description = "Indica se il controllo Label o lo sfondo di un controllo Shape è trasparente oppure opaco."
    BackStyle = UserControl.BackStyle
End Property

Public Property Let BackStyle(ByVal New_BackStyle As BackStyleClass)
    UserControl.BackStyle() = New_BackStyle
    PropertyChanged "BackStyle"
End Property

'AVVISO: NON RIMUOVERE O MODIFICARE LE SEGUENTI RIGHE DI COMMENTO
'MappingInfo=UserControl,UserControl,-1,BorderStyle
Public Property Get BorderStyle() As BorderStyleClass
Attribute BorderStyle.VB_Description = "Restituisce o imposta lo stile del bordo di un oggetto."
Attribute BorderStyle.VB_UserMemId = -504
    BorderStyle = UserControl.BorderStyle
End Property

Public Property Let BorderStyle(ByVal New_BorderStyle As BorderStyleClass)
    UserControl.BorderStyle() = New_BorderStyle
    PropertyChanged "BorderStyle"
End Property

'Il carattere di sottolineatura che segue "Circle" è necessario in quanto
'si tratta di una parola riservata di VBA.
'AVVISO: NON RIMUOVERE O MODIFICARE LE SEGUENTI RIGHE DI COMMENTO
'MappingInfo=UserControl,UserControl,-1,Circle
Public Sub Circle_(X As Single, Y As Single, Radius As Single, Color As Long, StartPos As Single, EndPos As Single, Aspect As Single)
    UserControl.Circle (X, Y), Radius, Color, StartPos, EndPos, Aspect
End Sub

Private Sub UserControl_Click()
    RaiseEvent Click
End Sub

'AVVISO: NON RIMUOVERE O MODIFICARE LE SEGUENTI RIGHE DI COMMENTO
'MappingInfo=UserControl,UserControl,-1,Cls
Public Sub Cls()
Attribute Cls.VB_Description = "Cancella le immagini e il testo generati in fase di esecuzione da un form o da un controllo Image o PictureBox."
    UserControl.Cls
End Sub

Private Sub UserControl_DblClick()
    RaiseEvent DblClick
End Sub

'AVVISO: NON RIMUOVERE O MODIFICARE LE SEGUENTI RIGHE DI COMMENTO
'MappingInfo=UserControl,UserControl,-1,DrawMode
Public Property Get DrawMode() As DrawModeConstants
Attribute DrawMode.VB_Description = "Imposta l'aspetto dell'output dei metodi grafici o di un controllo Shape o Line."
Attribute DrawMode.VB_UserMemId = -507
    DrawMode = UserControl.DrawMode
End Property

Public Property Let DrawMode(ByVal New_DrawMode As DrawModeConstants)
    UserControl.DrawMode() = New_DrawMode
    PropertyChanged "DrawMode"
End Property

'AVVISO: NON RIMUOVERE O MODIFICARE LE SEGUENTI RIGHE DI COMMENTO
'MappingInfo=UserControl,UserControl,-1,DrawStyle
Public Property Get DrawStyle() As DrawStyleConstants
Attribute DrawStyle.VB_Description = "Determina lo stile della linea per l'output di metodi grafici."
    DrawStyle = UserControl.DrawStyle
End Property

Public Property Let DrawStyle(ByVal New_DrawStyle As DrawStyleConstants)
    UserControl.DrawStyle() = New_DrawStyle
    PropertyChanged "DrawStyle"
End Property

'AVVISO: NON RIMUOVERE O MODIFICARE LE SEGUENTI RIGHE DI COMMENTO
'MappingInfo=UserControl,UserControl,-1,DrawWidth
Public Property Get DrawWidth() As Integer
Attribute DrawWidth.VB_Description = "Restituisce o imposta lo spessore della linea per l'output di metodi grafici."
    DrawWidth = UserControl.DrawWidth
End Property

Public Property Let DrawWidth(ByVal New_DrawWidth As Integer)
    UserControl.DrawWidth() = New_DrawWidth
    PropertyChanged "DrawWidth"
End Property

'AVVISO: NON RIMUOVERE O MODIFICARE LE SEGUENTI RIGHE DI COMMENTO
'MappingInfo=UserControl,UserControl,-1,Enabled
Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Restituisce o imposta un valore che determina se un oggetto è in grado di rispondere agli eventi generati dall'utente."
    Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    UserControl.Enabled() = New_Enabled
    PropertyChanged "Enabled"
End Property

'AVVISO: NON RIMUOVERE O MODIFICARE LE SEGUENTI RIGHE DI COMMENTO
'MappingInfo=UserControl,UserControl,-1,FillColor
Public Property Get FillColor() As OLE_COLOR
Attribute FillColor.VB_Description = "Restituisce o imposta il colore utilizzato per applicare riempimenti a forme, cerchi e caselle."
Attribute FillColor.VB_UserMemId = -510
    FillColor = UserControl.FillColor
End Property

Public Property Let FillColor(ByVal New_FillColor As OLE_COLOR)
    UserControl.FillColor() = New_FillColor
    PropertyChanged "FillColor"
End Property

'AVVISO: NON RIMUOVERE O MODIFICARE LE SEGUENTI RIGHE DI COMMENTO
'MappingInfo=UserControl,UserControl,-1,FillStyle
Public Property Get FillStyle() As FillStyleConstants
Attribute FillStyle.VB_Description = "Restituisce o imposta lo stile di riempimento di una forma."
    FillStyle = UserControl.FillStyle
End Property

Public Property Let FillStyle(ByVal New_FillStyle As FillStyleConstants)
    UserControl.FillStyle() = New_FillStyle
    PropertyChanged "FillStyle"
End Property

'AVVISO: NON RIMUOVERE O MODIFICARE LE SEGUENTI RIGHE DI COMMENTO
'MappingInfo=UserControl,UserControl,-1,Font
Public Property Get Font() As Font
Attribute Font.VB_Description = "Restituisce un oggetto Font."
Attribute Font.VB_UserMemId = -512
    Set Font = UserControl.Font
End Property

Public Property Set Font(ByVal New_Font As Font)
    Set UserControl.Font = New_Font
    PropertyChanged "Font"
End Property

'AVVISO: NON RIMUOVERE O MODIFICARE LE SEGUENTI RIGHE DI COMMENTO
'MappingInfo=UserControl,UserControl,-1,ForeColor
Public Property Get ForeColor() As OLE_COLOR
Attribute ForeColor.VB_Description = "Restituisce o imposta il colore di primo piano utilizzato per la visualizzazione di testo e grafica in un oggetto."
    ForeColor = UserControl.ForeColor
End Property

Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
    UserControl.ForeColor() = New_ForeColor
    PropertyChanged "ForeColor"
End Property

'AVVISO: NON RIMUOVERE O MODIFICARE LE SEGUENTI RIGHE DI COMMENTO
'MappingInfo=UserControl,UserControl,-1,HasDC
Public Property Get HasDC() As Boolean
Attribute HasDC.VB_Description = "Determina se al controllo viene assegnato un contesto di visualizzazione univoco."
    HasDC = UserControl.HasDC
End Property

'AVVISO: NON RIMUOVERE O MODIFICARE LE SEGUENTI RIGHE DI COMMENTO
'MappingInfo=UserControl,UserControl,-1,hDC
Public Property Get hDC() As Long
Attribute hDC.VB_Description = "Restituisce un handle (da Microsoft Windows) al contesto di periferica di un oggetto."
    hDC = UserControl.hDC
End Property

Private Sub UserControl_HitTest(X As Single, Y As Single, HitResult As Integer)
    RaiseEvent HitTest(X, Y, HitResult)
End Sub

'AVVISO: NON RIMUOVERE O MODIFICARE LE SEGUENTI RIGHE DI COMMENTO
'MappingInfo=UserControl,UserControl,-1,hWnd
Public Property Get hwnd() As Long
Attribute hwnd.VB_Description = "Restituisce un handle (da Microsoft Windows) alla finestra di un oggetto."
Attribute hwnd.VB_UserMemId = -515
    hwnd = UserControl.hwnd
End Property

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyDown(KeyCode, Shift)
End Sub

Private Sub UserControl_KeyPress(KeyAscii As Integer)
    RaiseEvent KeyPress(KeyAscii)
End Sub

Private Sub UserControl_KeyUp(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyUp(KeyCode, Shift)
End Sub

'AVVISO: NON RIMUOVERE O MODIFICARE LE SEGUENTI RIGHE DI COMMENTO
'MappingInfo=UserControl,UserControl,-1,MaskColor
Public Property Get MaskColor() As OLE_COLOR
Attribute MaskColor.VB_Description = "Restituisce o imposta il colore che specifica le aree trasparenti in MaskPicture."
    MaskColor = UserControl.MaskColor
End Property

Public Property Let MaskColor(ByVal New_MaskColor As OLE_COLOR)
    UserControl.MaskColor() = New_MaskColor
    PropertyChanged "MaskColor"
End Property

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseDown(Button, Shift, X, Y)
End Sub

'AVVISO: NON RIMUOVERE O MODIFICARE LE SEGUENTI RIGHE DI COMMENTO
'MappingInfo=UserControl,UserControl,-1,MouseIcon
Public Property Get MouseIcon() As Picture
Attribute MouseIcon.VB_Description = "Imposta un'icona personalizzata per il puntatore del mouse."
    Set MouseIcon = UserControl.MouseIcon
End Property

Public Property Set MouseIcon(ByVal New_MouseIcon As Picture)
    Set UserControl.MouseIcon = New_MouseIcon
    PropertyChanged "MouseIcon"
End Property

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

'AVVISO: NON RIMUOVERE O MODIFICARE LE SEGUENTI RIGHE DI COMMENTO
'MappingInfo=UserControl,UserControl,-1,MousePointer
Public Property Get MousePointer() As MousePointerConstants
Attribute MousePointer.VB_Description = "Restituisce o imposta il tipo di puntatore del mouse visualizzato quando il puntatore si trova su una parte specifica di un oggetto."
    MousePointer = UserControl.MousePointer
End Property

Public Property Let MousePointer(ByVal New_MousePointer As MousePointerConstants)
    UserControl.MousePointer() = New_MousePointer
    PropertyChanged "MousePointer"
End Property

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseUp(Button, Shift, X, Y)
End Sub

Private Sub UserControl_Paint()
    RaiseEvent Paint
End Sub

'AVVISO: NON RIMUOVERE O MODIFICARE LE SEGUENTI RIGHE DI COMMENTO
'MappingInfo=UserControl,UserControl,-1,PaintPicture
Public Sub PaintPicture(ByVal Picture As Picture, ByVal X1 As Single, ByVal Y1 As Single, Optional ByVal Width1 As Variant, Optional ByVal Height1 As Variant, Optional ByVal X2 As Variant, Optional ByVal Y2 As Variant, Optional ByVal Width2 As Variant, Optional ByVal Height2 As Variant, Optional ByVal Opcode As Variant)
Attribute PaintPicture.VB_Description = "Disegna il contenuto di un file di grafica su un oggetto Form, PictureBox o Printer."
    UserControl.PaintPicture Picture, X1, Y1, Width1, Height1, X2, Y2, Width2, Height2, Opcode
End Sub

'AVVISO: NON RIMUOVERE O MODIFICARE LE SEGUENTI RIGHE DI COMMENTO
'MappingInfo=UserControl,UserControl,-1,Palette
Public Property Get Palette() As Picture
Attribute Palette.VB_Description = "Restituisce o imposta un'immagine che contiene la tavolozza da utilizzare su un oggetto quando PaletteMode viene impostata su Custom."
    Set Palette = UserControl.Palette
End Property

Public Property Set Palette(ByVal New_Palette As Picture)
    Set UserControl.Palette = New_Palette
    PropertyChanged "Palette"
End Property

'AVVISO: NON RIMUOVERE O MODIFICARE LE SEGUENTI RIGHE DI COMMENTO
'MappingInfo=UserControl,UserControl,-1,PaletteMode
Public Property Get PaletteMode() As PaletteModeConstants
Attribute PaletteMode.VB_Description = "Restituisce o imposta un valore che determina quale tavolozza utilizzare per i controlli in un oggetto."
    PaletteMode = UserControl.PaletteMode
End Property

Public Property Let PaletteMode(ByVal New_PaletteMode As PaletteModeConstants)
    UserControl.PaletteMode() = New_PaletteMode
    PropertyChanged "PaletteMode"
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

'Il carattere di sottolineatura che segue "PSet" è necessario in quanto
'si tratta di una parola riservata di VBA.
'AVVISO: NON RIMUOVERE O MODIFICARE LE SEGUENTI RIGHE DI COMMENTO
'MappingInfo=UserControl,UserControl,-1,PSet
Public Sub PSet_(X As Single, Y As Single, Color As Long)
    UserControl.PSet Step(X, Y), Color
End Sub

'AVVISO: NON RIMUOVERE O MODIFICARE LE SEGUENTI RIGHE DI COMMENTO
'MappingInfo=UserControl,UserControl,-1,Refresh
Public Sub Refresh()
Attribute Refresh.VB_Description = "Ridisegna completamente un oggetto."
    UserControl.Refresh
End Sub

Private Sub UserControl_Resize()
    RaiseEvent Resize
End Sub

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
'
''AVVISO: NON RIMUOVERE O MODIFICARE LE SEGUENTI RIGHE DI COMMENTO
''MappingInfo=UserControl,UserControl,-1,ScaleLeft
'Public Property Get ScaleLeft() As Single
'    ScaleLeft = UserControl.ScaleLeft
'End Property
'
'Public Property Let ScaleLeft(ByVal New_ScaleLeft As Single)
'    UserControl.ScaleLeft() = New_ScaleLeft
'    PropertyChanged "ScaleLeft"
'End Property

'AVVISO: NON RIMUOVERE O MODIFICARE LE SEGUENTI RIGHE DI COMMENTO
'MappingInfo=UserControl,UserControl,-1,ScaleMode
Public Property Get ScaleMode() As ScaleModeConstants
Attribute ScaleMode.VB_Description = "Restituisce o imposta un valore che indica le unità di misura per le coordinate di un oggetto quando si utilizzano metodi grafici o si posizionano controlli."
    ScaleMode = UserControl.ScaleMode
End Property

Public Property Let ScaleMode(ByVal New_ScaleMode As ScaleModeConstants)
    UserControl.ScaleMode() = New_ScaleMode
    PropertyChanged "ScaleMode"
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

'AVVISO: NON RIMUOVERE O MODIFICARE LE SEGUENTI RIGHE DI COMMENTO
'MemberInfo=0,0,0,False
Public Property Get AutoSize() As Boolean
Attribute AutoSize.VB_Description = "Return or set the control autosize."
Attribute AutoSize.VB_UserMemId = -500
    AutoSize = m_AutoSize
End Property

Public Property Let AutoSize(ByVal New_AutoSize As Boolean)
    m_AutoSize = New_AutoSize
    PropertyChanged "AutoSize"
End Property

'AVVISO: NON RIMUOVERE O MODIFICARE LE SEGUENTI RIGHE DI COMMENTO
'MemberInfo=11
Public Sub PrintPicture()
    Call PrintPictureToFitPage(Printer, UserControl.Picture)
End Sub

'AVVISO: NON RIMUOVERE O MODIFICARE LE SEGUENTI RIGHE DI COMMENTO
'MappingInfo=UserControl,UserControl,-1,MaskPicture
Public Property Get MaskPicture() As Picture
Attribute MaskPicture.VB_Description = "Restituisce o imposta l'immagine che specifica l'area di un controllo disegnabile o su cui si può fare clic quando BackStyle è 0 (trasparente)."
    Set MaskPicture = UserControl.MaskPicture
End Property

Public Property Set MaskPicture(ByVal New_MaskPicture As Picture)
    Set UserControl.MaskPicture = New_MaskPicture
    PropertyChanged "MaskPicture"
End Property

