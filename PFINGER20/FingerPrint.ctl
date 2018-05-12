VERSION 5.00
Begin VB.UserControl FingerPrint 
   BorderStyle     =   1  'Fixed Single
   CanGetFocus     =   0   'False
   ClientHeight    =   1500
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1500
   PropertyPages   =   "FingerPrint.ctx":0000
   ScaleHeight     =   100
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   100
   ToolboxBitmap   =   "FingerPrint.ctx":0012
   Begin VB.Timer tmrChkFinger 
      Left            =   540
      Top             =   780
   End
End
Attribute VB_Name = "FingerPrint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Valori predefiniti pubblici
Public Enum ErrorsConstants
    PBBOTTOM = &HC
    PBCANCEL = &H5
    PBDARKIMAGE = &H9
    PBINVALID_PARAMETERS = &H3
    PBINVALID_SMARTCARD = &H7
    PBINVALID_TEMPLATE = &H4
    PBLEFT = &HD
    PBLITTLEPRESS = &HA
    PBNOCOMMUNICATION = &H8
    PBNOREADER = &H1
    PBNORESOURCES = &H2
    PBNOSENSOR = &HF
    PBNOSMARTCARD = &H6
    PBOK = &H0
    PBRIGHT = &HE
    PBUP = &HB
    PBIENOTLOADED = &HFD
    PBNOBETTERQUALITY = &HFE
End Enum

'Dichiarazioni di eventi:
Public Event FingerIn()
Attribute FingerIn.VB_HelpID = 40
Public Event FingerOut()
Attribute FingerOut.VB_HelpID = 41
Public Event Errors(ByVal Number As Long, ByVal Description As String)
Attribute Errors.VB_HelpID = 42
Attribute Errors.VB_UserMemId = -608

'Costanti varie
Private Const PBSENSOR = &H2
Private Const SRCCOPY = &HCC0020

Private Const HTMLFormName As String = "frmFingerPrint"
Private Const HTMLRawDataName As String = "rawdata"

'Variabili uso interno
Private lBak As Long
Private lInit As Long
Private iImgRow As Long
Private iImgCol As Long
Private iSecLvl As Long
Private lFinger As Long
Private lRawImageSize As Long
Private lTemplateSize As Long
Private aRawImage() As Byte
Private aTemplate() As Byte
Private aRawBackup() As Byte
Private iCounter As Integer
Private GDI As cGDI
'Private ZIP As clsZlibWrapper

'Dichiarazioni WinAPI
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hDC As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Private Declare Function StretchBlt Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hDC As Long) As Long

'Dichiarazioni PB.DLL
Private Declare Function pbClose Lib "pb.dll" () As Long
Private Declare Function pbEnroll Lib "pb.dll" (template As Any, image As Any, valid As Long) As Long
Private Declare Function pbFingerPosition Lib "pb.dll" (image As Any, high As Long, low As Long, Left As Long, Right As Long) As Long
Private Declare Function pbFingerPresent Lib "pb.dll" (image As Any, Present As Long) As Long
Private Declare Function pbGetActualTemplateSize Lib "pb.dll" (template As Any, Size As Long) As Long
Private Declare Function pbGetCapabilities Lib "pb.dll" (Capabilities As Long) As Long
Private Declare Function pbGetRawImage Lib "pb.dll" (image As Any) As Long
Private Declare Function pbGetRawImageSize Lib "pb.dll" (nof_row As Long, nof_col As Long) As Long
Private Declare Function pbGetTemplateSize Lib "pb.dll" (Size As Long) As Long
Private Declare Function pbInitialize Lib "pb.dll" () As Long
Private Declare Function pbSetSecurityLevel Lib "pb.dll" (ByVal security_level As Long) As Long
Private Declare Function pbTextureQuality Lib "pb.dll" (image As Any, brightness As Long, measure As Long) As Long
Private Declare Function pbVerifyFingerprintEx Lib "pb.dll" (template As Any, image As Any, match As Long) As Long

'Valori predefiniti proprietà:
'Private Const m_def_Compressed = False
Private Const m_def_Quality = 64
Private Const m_def_SecurityLevel = 4

'Variabili proprietà:
'Private m_Compressed As Boolean
Private m_Quality As Long
Private m_SecurityLevel As Long
Private m_Sensor As Boolean
Private m_Status As Long
Private m_RawData As String
Private m_TemplateData As String

Public Sub About()
Attribute About.VB_UserMemId = -552
    frmAbout.Show vbModal
End Sub

'MemberInfo=8,0,0,0
Public Property Get Status() As Long
    Status = m_Status
End Property

Public Property Let Status(ByVal New_Status As Long)
    If Ambient.UserMode = False Then Err.Raise 387
    If Ambient.UserMode Then Err.Raise 382
End Property

'MemberInfo=0,1,2,False
Public Property Get Sensor() As Boolean
    Sensor = m_Sensor
End Property

Public Property Let Sensor(ByVal New_Sensor As Boolean)
    If Ambient.UserMode = False Then Err.Raise 387
    If Ambient.UserMode Then Err.Raise 382
End Property

'MemberInfo=0,0,0,0
Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Restituisce o imposta un valore che determina se un oggetto è in grado di rispondere agli eventi generati dall'utente."
Attribute Enabled.VB_UserMemId = -514
    Enabled = tmrChkFinger.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    tmrChkFinger.Enabled = New_Enabled
    PropertyChanged "Enabled"
End Property

'MemberInfo=8,0,0,0
Public Property Get Interval() As Long
Attribute Interval.VB_HelpID = 29
    Interval = tmrChkFinger.Interval
End Property

Public Property Let Interval(ByVal New_Interval As Long)
    ' se non c'è un lettore di impronte genera un errore e non imposta la proprietà
    If Ambient.UserMode And New_Interval > 0 And tmrChkFinger.Enabled And lInit <> PBOK Then
        RaiseEvent Errors(PBNOREADER, GetPBErrorDescription(PBNOREADER))
        New_Interval = 0
    End If
    
#If VERIFYSENSOR = 1 Then
    ' se non c'è il sensore genera un errore e non imposta la proprietà
    If Ambient.UserMode And New_Interval > 0 And tmrChkFinger.Enabled And Not m_Sensor Then
        RaiseEvent Errors(PBNOSENSOR, GetPBErrorDescription(PBNOSENSOR))
        New_Interval = 0
    End If
#End If
    
    tmrChkFinger.Interval = New_Interval
    PropertyChanged "Interval"
End Property

'MemberInfo=14,1,1,0
Public Property Get TemplateData() As String
Attribute TemplateData.VB_HelpID = 34
    'If m_Compressed Then
    '    ' comprime il template se richiesto
    '    TemplateData = ZipBinary2Str(aTemplate)
    'Else
        ' altrimenti lo restituisce integro
        TemplateData = m_TemplateData
    'End If
End Property

Public Property Let TemplateData(ByVal NewValue As String)
    If Ambient.UserMode = False Then Err.Raise 387 ' Progettazione
    If Ambient.UserMode Then Err.Raise 382         ' Esecuzione
End Property

'MemberInfo=1,1,1,0
Public Property Get TemplateDataB() As Byte()
Attribute TemplateDataB.VB_HelpID = 33
    'If m_Compressed Then
    '    ' comprime il template se richiesto
    '    TemplateDataB = ZipBinary(aTemplate)
    'Else
        ' altrimenti lo restituisce integro
        TemplateDataB = aTemplate
    'End If
End Property

Public Property Let TemplateDataB(ByRef New_TemplateDataB() As Byte)
    If Ambient.UserMode = False Then Err.Raise 387
    If Ambient.UserMode Then Err.Raise 382
End Property

'MemberInfo=14,1,1,0
Public Property Get RawData() As String
    'If m_Compressed Then
    '    ' comprime i dati raw se richiesto
    '    RawData = ZipBinary2Str(aRawBackup)
    'Else
        ' altrimenti li restituisce integri
        RawData = m_RawData
    'End If
End Property

Public Property Let RawData(ByVal NewValue As String)
    If Ambient.UserMode = False Then Err.Raise 387 ' Progettazione
    If Ambient.UserMode Then Err.Raise 382         ' Esecuzione
End Property

'MemberInfo=1,1,1,0
Public Property Get RawDataB() As Byte()
    'If m_Compressed Then
    '    ' comprime i dati raw se richiesto
    '    RawDataB = ZipBinary(aRawBackup)
    'Else
        ' altrimenti li restituisce integri
        RawDataB = aRawBackup
    'End If
End Property

Public Property Let RawDataB(ByRef New_RawDataB() As Byte)
    If Ambient.UserMode = False Then Err.Raise 387
    If Ambient.UserMode Then Err.Raise 382
End Property

'MappingInfo=UserControl,UserControl,-1,Appearance
Public Property Get Appearance3D() As Boolean
    Appearance3D = CBool(UserControl.Appearance)
End Property

Public Property Let Appearance3D(ByVal New_Appearance As Boolean)
    UserControl.Appearance() = Abs(CInt(New_Appearance))
    PropertyChanged "Appearance3D"
End Property

'MappingInfo=UserControl,UserControl,-1,BorderStyle
Public Property Get Border() As Boolean
    Border = CBool(UserControl.BorderStyle)
End Property

Public Property Let Border(ByVal New_BorderStyle As Boolean)
    UserControl.BorderStyle() = Abs(CInt(New_BorderStyle))
    PropertyChanged "Border"
End Property

'MappingInfo=UserControl,UserControl,-1,hDC
Public Property Get hDC() As Long
    hDC = UserControl.hDC
    hDC = 0
End Property

Public Property Let hDC(ByVal New_HDC As Long)
    If Ambient.UserMode = False Then Err.Raise 387 ' Progettazione
    If Ambient.UserMode Then Err.Raise 382         ' Esecuzione
End Property

'MappingInfo=UserControl,UserControl,-1,hWnd
Public Property Get hwnd() As Long
    hwnd = UserControl.hwnd
End Property

Public Property Let hwnd(ByVal New_hWnd As Long)
    If Ambient.UserMode = False Then Err.Raise 387 ' Progettazione
    If Ambient.UserMode Then Err.Raise 382         ' Esecuzione
End Property

'MappingInfo=UserControl,UserControl,-1,Picture
Public Property Get Picture() As Picture
    Set Picture = UserControl.Picture
End Property

Public Property Set Picture(ByVal New_Picture As Picture)
    Set UserControl.Picture = New_Picture
    PropertyChanged "Picture"
End Property

'MemberInfo=8,0,0,4
Public Property Get SecurityLevel() As Long
    SecurityLevel = m_SecurityLevel
End Property

Public Property Let SecurityLevel(ByVal New_SecurityLevel As Long)
    If New_SecurityLevel > 0 And New_SecurityLevel < 8 Then
        iSecLvl = New_SecurityLevel
        m_SecurityLevel = New_SecurityLevel
        pbSetSecurityLevel New_SecurityLevel
        PropertyChanged "SecurityLevel"
    End If
End Property

'MemberInfo=8,0,0,64
Public Property Get Quality() As Long
    Quality = m_Quality
End Property

Public Property Let Quality(ByVal New_Quality As Long)
    m_Quality = New_Quality
    PropertyChanged "Quality"
End Property

'MemberInfo=0,0,0,False
'Public Property Get Compressed() As Boolean
'    Compressed = m_Compressed
'End Property

'Public Property Let Compressed(ByVal New_Compressed As Boolean)
'    m_Compressed = New_Compressed
'    PropertyChanged "Compressed"
'End Property

'MemberInfo=0
Public Function VerifyFinger(ByRef lpTemplate As String, Optional ByVal iSecurityLevel As Long = 4) As Boolean
Attribute VerifyFinger.VB_HelpID = 36
    Dim lvalid As Long
    Dim templateByte() As Byte
    
    ' cambia il livello di sicurezza (1..7)
    If iSecLvl <> iSecurityLevel And (iSecurityLevel > 0 And iSecurityLevel < 8) Then
        iSecLvl = iSecurityLevel
        pbSetSecurityLevel iSecurityLevel
    End If
    
    ' converte la stringa in un array
    ReDim templateByte(Len(lpTemplate))
    CopyMemory templateByte(0), ByVal StrPtr(lpTemplate), Len(lpTemplate)
    
    ' decomprime il template se richiesto
    'If m_Compressed Then templateByte = UnzipBinary(templateByte)
    
    ' verifica il template con l'ultima impronta presente nel lettore
    pbVerifyFingerprintEx templateByte(0), aRawBackup(0), lvalid
    VerifyFinger = CBool(lvalid)
End Function

'MemberInfo=0
Public Function VerifyFingerB(ByRef lpTemplate() As Byte, Optional ByVal iSecurityLevel As Long = 4) As Boolean
Attribute VerifyFingerB.VB_HelpID = 35
    Dim lvalid As Long
    'Dim tmpzip() As Byte
    
    ' cambia il livello di sicurezza (1..7)
    If iSecLvl <> iSecurityLevel And (iSecurityLevel > 0 And iSecurityLevel < 8) Then
        iSecLvl = iSecurityLevel
        pbSetSecurityLevel iSecurityLevel
    End If
    
    'If m_Compressed Then
    '    ' decomprime il template
    '    tmpzip = UnzipBinary(lpTemplate)
        
        ' verifica il template con l'ultima impronta presente sul lettore
    '    pbVerifyFingerprintEx tmpzip(0), aRawBackup(0), lValid
    'Else
        ' verifica il template con l'ultima impronta presente sul lettore
        pbVerifyFingerprintEx lpTemplate(0), aRawBackup(0), lvalid
    'End If
    VerifyFingerB = CBool(lvalid)
End Function

'MemberInfo=0
Public Function VerifyFingerEx(ByRef lpTemplate() As Byte, ByRef lpRawImage() As Byte, Optional ByVal iSecurityLevel As Long = 4) As Boolean
Attribute VerifyFingerEx.VB_HelpID = 37
    Dim lvalid As Long
    'Dim tmpzip() As Byte
    'Dim rawzip() As Byte
    
    ' cambia il livello di sicurezza (1..7)
    If iSecLvl <> iSecurityLevel And (iSecurityLevel > 0 And iSecurityLevel < 8) Then
        iSecLvl = iSecurityLevel
        pbSetSecurityLevel iSecurityLevel
    End If
    
    'If m_Compressed Then
    '    ' decomprime i dati
    '    tmpzip = UnzipBinary(lpTemplate)
    '    rawzip = UnzipBinary(lpRawImage)
        
        ' verifica il template con i dati raw
    '    pbVerifyFingerprintEx tmpzip(0), rawzip(0), lValid
    'Else
        ' verifica il template con i dati raw
        pbVerifyFingerprintEx lpTemplate(0), lpRawImage(0), lvalid
    'End If
    VerifyFingerEx = CBool(lvalid)
End Function

'MemberInfo=0
Public Function VerifyFingerPic(ByRef lpTemplate() As Byte, ByRef Pic As StdPicture, Optional ByVal iSecurityLevel As Long = 4) As Boolean
    Dim RawData() As Byte
    
    ' cambia il livello di sicurezza (1..7)
    If iSecLvl <> iSecurityLevel And (iSecurityLevel > 0 And iSecurityLevel < 8) Then
        iSecLvl = iSecurityLevel
        pbSetSecurityLevel iSecurityLevel
    End If
    
    ' converte la picture in dati raw e li verifica
    RawData = GDI.PictureToRaw(Pic)
    VerifyFingerPic = VerifyFingerEx(lpTemplate, RawData)
End Function

Public Function RawToPicture(ByRef RawData() As Byte) As StdPicture
    'Dim rawzip() As Byte
    
    'If m_Compressed Then
    '    ' decomprime i dati raw
    '    rawzip = UnzipBinary(RawData)
        
        ' converte dei dati raw in una picture
    '    Set RawToPicture = ToPicture(rawzip, False)
    'Else
        ' converte dei dati raw in una picture
        Set RawToPicture = ToPicture(RawData, False)
    'End If
End Function

Public Function PictureToRaw(ByVal Picture As StdPicture) As Byte()
    Dim raw() As Byte
    
    ' trsforma la picture in dati raw
    raw = GDI.PictureToRaw(Picture)
    
    ' restituisce il risultato de/compresso
    'If m_Compressed Then
    '    PictureToRaw = ZipBinary(raw)
    'Else
        PictureToRaw = raw
    'End If
End Function

Public Function RawToTemplate(ByRef RawData() As Byte) As Byte()
    Dim lvalid As Long
    Dim aRet() As Byte
    
    ReDim aRet(lTemplateSize) As Byte       ' ridimensiona il buffer
    pbEnroll aRet(0), RawData(0), lvalid    ' ricava il template e la validazione
    If Not CBool(lvalid) Then               ' se NON è una impronta valida
        ' genera l'evento di errore
        RaiseEvent Errors(PBINVALID_TEMPLATE, vbNullString)
    End If
    ' restituisce il template
    RawToTemplate = aRet
End Function

Public Function Load(ByVal sFilename As String) As StdPicture
    ' carica un immagine d un file
    Set Load = GDI.LoadImage(sFilename)
End Function

Public Sub Save(ByVal Picture As StdPicture, ByVal sFilename As String, Optional ByVal iQuality As Integer = 0)
    Dim itype As Integer
    
    ' salva una immagine in un file
    If Len(sFilename) > 0 Then                      ' se ha scelto un nome
        Select Case UCase$(Right$(sFilename, 4))    ' controlla la sua estensione
            Case ".JPG"                             ' se è JPEG
                itype = IMAGETYPE_JPEG              ' imposta il tipo di file
            Case ".GIF"                             ' se è GIF
                itype = IMAGETYPE_GIF               ' imposta il tipo di file
            Case ".TIF"                             ' se è TIFF
                itype = IMAGETYPE_TIFF              ' imposta il tipo di file
            Case ".PNG"                             ' se è PNG
                itype = IMAGETYPE_PNG               ' imposta il tipo di file
            Case Else                               ' se non è nessuna riconosciuta
                ' se ha la qualità impostata allora salva una JPEG altrimenti salva una BMP
                If iQuality > 0 Then itype = IMAGETYPE_JPEG Else itype = IMAGETYPE_BITMAP
        End Select
        Set GDI.Picture = Picture
        GDI.SaveImage sFilename, itype, iQuality
    End If
End Sub

Public Sub ShowSavePicture(ByVal Picture As StdPicture, _
    Optional sFilter As String = vbNullString, _
    Optional sTitle As String = vbNullString, _
    Optional sInitDir As String = vbNullString, _
    Optional ByVal iQuality As Integer = 0)

    Dim sfile As String
    Dim dlg As cDialog
    
    Set dlg = New cDialog
    
    ' visualizza la finestra standard per la selezione di un nome di file
    sfile = dlg.SaveDialog(UserControl.hwnd, sFilter, sTitle, sInitDir)
    Me.Save Picture, sfile, iQuality
    Set dlg = Nothing
End Sub

' **** ATTENZIONE!!! QUESTA FUNZIONE RESTITUISCE UN ERRORE IN VB.NET ****
Public Function ShowLoadPicture(Optional sFilter As String = vbNullString, _
    Optional sTitle As String = vbNullString, _
    Optional sInitDir As String = vbNullString) As StdPicture
    
    Dim sfile As String
    Dim dlg As cDialog
   
    Set dlg = New cDialog
    Set ShowLoadPicture = Nothing
    
    ' visualizza la finestra standard per la selezione di un nome di file
    sfile = dlg.OpenDialog(UserControl.hwnd, sFilter, sTitle, sInitDir)
    If Len(sfile) > 0 Then Set ShowLoadPicture = GDI.LoadImage(sfile)
    Set dlg = Nothing
End Function

Public Function ConvertToBW(ByVal Picture As StdPicture) As StdPicture
    Dim Pic As cPicture24
    
    ' converte una picture in bianco e nero
    Set Pic = New cPicture24
    Set Pic.Picture = Picture
    Pic.BlackWhite
    Set ConvertToBW = Pic.Picture
    Set Pic = Nothing
End Function

Public Function HTTPSend(ByVal URL As String, ByRef RawData() As Byte, Optional bNewWindow As Boolean = True) As Object
    Dim ie As Object
    Dim frm As Object
    Dim stmp As String
    Dim tID As String
    
    If StrComp(TypeName(UserControl.Parent), "HTMLDocument") = 0 Then
        tID = App.ThreadID
        
        ' converte l'array in una stringa
        stmp = String$(UBound(RawData), vbNullChar)
        CopyMemory ByVal StrPtr(stmp), RawData(0), UBound(RawData)
        
        ' invia i dati tramite la pagina corrente
        Set ie = UserControl.Parent.script.document
        If Not ie Is Nothing Then
            ie.write "<form id=""" & HTMLFormName & tID & """ method=""POST"" enctype=""multipart/form-data"" action=""" & URL & """ style=""display:none"">"
            ie.write "<textarea name=""" & HTMLRawDataName & tID & """>" & stmp & "</textarea>"
            ie.write "</form>"
            Set frm = ie.getElementById("frmFingerPrint" & tID)
            frm.submit
            Set frm = Nothing
        End If
    Else
        On Local Error Resume Next
        Set ie = CreateObject("InternetExplorer.Application")
        If Err.Number <> 0 Then
            RaiseEvent Errors(PBIENOTLOADED, Err.Description)
            Exit Function
        End If
        
        ie.Visible = bNewWindow
        
        'invia i dati
        ie.Navigate2 URL, , , RawData, "Content-Type: multipart/form-data"
        
        ' attende che ha finito
        Do While ie.Busy
        Loop
        
        ' restituisce il document
        Set HTTPSend = ie.document
        
        ' scarica IE
        If Not bNewWindow Then ie.Quit
    End If
    Set ie = Nothing
End Function

Private Sub tmrChkFinger_Timer()
    Dim lHigh As Long
    Dim lLow As Long
    Dim lLeft As Long
    Dim lRight As Long
    Dim lBright As Long
    Dim lQuality As Long
    Dim lvalid As Long
    Dim lRet As Long
    Dim ltype As Long
    Dim lSize As Long
    
    lBak = lFinger                                                  ' memorizza lo stato attuale dell'impronta
    m_RawData = vbNullString                                        ' svuota la stringa della rawimage
    m_TemplateData = vbNullString                                   ' svuota la stringa del template
    lRet = pbGetRawImage(aRawImage(0))                              ' recupera l'immagine del sensore
    If lRet = PBOK Then pbFingerPresent aRawImage(0), lFinger       ' recupera se è presente l'impronta
    If CBool(lFinger) Then                                          ' se è presente l'impronta sul sensore
        lRet = pbTextureQuality(aRawImage(0), lBright, lQuality)    ' recupera la qualità della luce e della pressione
        If lBright = -1 Then                                        ' immagine scura
            m_Status = PBDARKIMAGE                                  ' imposta lo stato corrente
            RaiseEvent Errors(PBDARKIMAGE, "Too much pressure")     ' genera l'evento di errore
        ElseIf lBright = 1 Then                                     ' poca pressione
            m_Status = PBLITTLEPRESS                                ' imposta lo stato corrente
            RaiseEvent Errors(PBLITTLEPRESS, "Little pressure")     ' genera l'evento di errore
        ElseIf lQuality < m_Quality Then                            ' se non ha una qualita soddisfacente
            m_Status = PBNOBETTERQUALITY                            ' imposta lo stato corrente
            RaiseEvent Errors(PBNOBETTERQUALITY, "No better quality")
        Else
            ' recupera la posizione dell'impronta
            lRet = pbFingerPosition(aRawImage(0), lHigh, lLow, lLeft, lRight)
            If Not CBool(lHigh) Then                        ' se è troppo in basso
                m_Status = PBUP                             ' imposta lo stato corrente
                RaiseEvent Errors(PBUP, "Move up")          ' genera l'evento di errore
            ElseIf Not CBool(lLow) Then                     ' se è troppo in alto
                m_Status = PBBOTTOM                         ' imposta lo stato corrente
                RaiseEvent Errors(PBBOTTOM, "Move down")    ' genera l'evento di errore
            ElseIf Not CBool(lRight) Then                   ' se è troppo a sinistra
                m_Status = PBRIGHT                          ' imposta lo stato corrente
                RaiseEvent Errors(PBRIGHT, "Move right")    ' genera l'evento di errore
            ElseIf Not CBool(lLeft) Then                    ' se è troppo a destra
                m_Status = PBLEFT                           ' imposta lo stato corrente
                RaiseEvent Errors(PBLEFT, "Move left")      ' genera l'evento di errore
            Else
                ' backup dell'impronta attuale
                CopyMemory aRawBackup(0), aRawImage(0), lRawImageSize
                ' copia il backup convertendolo in una stringa
                m_RawData = String$(lRawImageSize, vbNullChar)
                CopyMemory ByVal StrPtr(m_RawData), aRawBackup(0), lRawImageSize
                pbEnroll aTemplate(0), aRawImage(0), lvalid     ' ricava il template e la validazione
                If CBool(lvalid) Then                           ' se è una impronta valida
                    ToPicture aRawImage, True                   ' disegna l'impronta nell'UserControl
                    m_Status = PBOK                             ' imposta lo stato corrente
                    pbGetActualTemplateSize aTemplate(0), lSize ' ricava la dimensione del template
                    m_TemplateData = String$(lSize, vbNullChar) ' inizializza la stringa per il template
                    ' copia il template convertendolo in una stringa
                    CopyMemory ByVal StrPtr(m_TemplateData), aTemplate(0), lSize
                    If lBak <> lFinger Then ' se lo stato è diverso dal precedente
                        RaiseEvent FingerIn ' genera l'evento impronta presente
                    End If
                End If
            End If
        End If
    ElseIf lBak <> lFinger Then             ' se lo stato è diverso dal precedente
        UserControl.Picture = LoadPicture() ' cancella l'immagine
        RaiseEvent FingerOut                ' genera l'evento impronta non presente
    End If
End Sub

Private Sub UserControl_Initialize()
    ' inizializza il contatore dei numero massimo
    ' di utilizzo se è in versione SHAREWARE
    iCounter = 0
    
    ' inizializza il livello di sicurezza per
    ' il controllo dell'impronta con VerifyFinger
    iSecLvl = 4
    
    ' inizializza gli oggetti
    Set GDI = New cGDI
    'Set ZIP = New clsZlibWrapper
    
    #If SHAREWARE = 1 Then
        Randomize Timer
        If Rnd * 10 Mod 2 Then frmSplash.Show vbModal
    #End If

    Call InitReader
End Sub

'Inizializza le proprietà di UserControl.
Private Sub UserControl_InitProperties()
    m_RawData = vbNullString
    m_TemplateData = vbNullString
    m_SecurityLevel = m_def_SecurityLevel
    m_Quality = m_def_Quality
    m_Sensor = False
    m_Status = 0
    'm_Compressed = m_def_Compressed
End Sub

Private Sub UserControl_Paint()
    DrawDemoVersion
End Sub

'Carica i valori delle proprietà dalla posizione di memorizzazione.
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    tmrChkFinger.Enabled = PropBag.ReadProperty("Enabled", True)
    tmrChkFinger.Interval = PropBag.ReadProperty("Interval", 0)
    UserControl.Appearance = Abs(CInt(PropBag.ReadProperty("Appearance3D", True)))
    UserControl.BorderStyle = Abs(CInt(PropBag.ReadProperty("Border", True)))
    Set Picture = PropBag.ReadProperty("Picture", Nothing)
    m_TemplateData = PropBag.ReadProperty("TemplateData", vbNullString)
    m_RawData = PropBag.ReadProperty("RawData", vbNullString)
    m_SecurityLevel = PropBag.ReadProperty("SecurityLevel", m_def_SecurityLevel)
    m_Quality = PropBag.ReadProperty("Quality", m_def_Quality)
    m_Status = PropBag.ReadProperty("Status", 0)
    'm_Compressed = PropBag.ReadProperty("Compressed", m_def_Compressed)
End Sub

Private Sub UserControl_Terminate()
    Set GDI = Nothing
    'Set ZIP = Nothing
    
    ' chiude la connessione al lettore
    If lInit = PBOK Then pbClose
End Sub

'Scrive i valori delle proprietà nella posizione di memorizzazione.
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("Enabled", tmrChkFinger.Enabled, True)
    Call PropBag.WriteProperty("Interval", tmrChkFinger.Interval, 0)
    Call PropBag.WriteProperty("Appearance3D", CBool(UserControl.Appearance), True)
    Call PropBag.WriteProperty("Border", CBool(UserControl.BorderStyle), True)
    Call PropBag.WriteProperty("Picture", Picture, Nothing)
    Call PropBag.WriteProperty("TemplateData", m_TemplateData, vbNullString)
    Call PropBag.WriteProperty("RawData", m_RawData, vbNullString)
    Call PropBag.WriteProperty("SecurityLevel", m_SecurityLevel, m_def_SecurityLevel)
    Call PropBag.WriteProperty("Quality", m_Quality, m_def_Quality)
    Call PropBag.WriteProperty("Status", m_Status, 0)
    'Call PropBag.WriteProperty("Compressed", m_Compressed, m_def_Compressed)
End Sub

Private Function GetPBErrorDescription(Error As Long) As String
    Select Case Error
        Case PBNOREADER
            GetPBErrorDescription = "There is no reader connected."
        Case PBNORESOURCES
            GetPBErrorDescription = "We have no resources."
        Case PBINVALID_PARAMETERS
            GetPBErrorDescription = "The parameters are invalid."
        Case PBINVALID_TEMPLATE
            GetPBErrorDescription = "The template is invalid."
        Case PBCANCEL
            GetPBErrorDescription = "The user pressed cancel."
        Case PBNOSMARTCARD
            GetPBErrorDescription = "There is no smart card in the reader."
        Case PBINVALID_SMARTCARD
            GetPBErrorDescription = "The inserted smart card is invalid."
        Case PBNOCOMMUNICATION
            GetPBErrorDescription = "There is no device for Precise 100."
        Case PBNOSENSOR
            GetPBErrorDescription = "There is no sensor in the reader."
        Case Else
            GetPBErrorDescription = "Unknown error code" & CStr(Error)
    End Select
End Function

Private Sub InitReader()
    Dim ltype As Long
    
    'If Not Ambient.UserMode Then Exit Sub
    
    ' inizializza il lettore
    lInit = pbInitialize
    If lInit = PBOK Then
        ' controlla se è presente il sensore
        pbGetCapabilities ltype
        m_Sensor = (ltype And PBSENSOR)
#If VERIFYSENSOR = 1 Then
        If m_Sensor Then
            ' ridimensiona i buffers
            pbGetRawImageSize iImgRow, iImgCol
            pbGetTemplateSize lTemplateSize
            lRawImageSize = iImgRow * iImgCol
            ReDim aRawImage(lRawImageSize)
            ReDim aRawBackup(lRawImageSize)
            ReDim aTemplate(lTemplateSize)
        End If
#Else
            ' ridimensiona i buffers
            pbGetRawImageSize iImgRow, iImgCol
            pbGetTemplateSize lTemplateSize
            lRawImageSize = iImgRow * iImgCol
            ReDim aRawImage(lRawImageSize)
            ReDim aRawBackup(lRawImageSize)
            ReDim aTemplate(lTemplateSize)
#End If
    End If
End Sub

Private Function ToPicture(ByRef RawData() As Byte, ByVal bToUserControl As Boolean) As StdPicture
    Dim c As Long
    Dim iy As Long
    Dim ix As Long
    Dim lColor As Long
    Dim lhdc As Long
    Dim lhwnd As Long
    
    ' crea una bitmap in memoria per il disegno
    lhdc = CreateCompatibleDC(0)
    lhwnd = CreateCompatibleBitmap(lhdc, iImgCol, iImgRow)
    SelectObject lhdc, lhwnd

    ' disegna l'impronta
    GDI.DrawRawImage lhdc, aRawImage, iImgCol, iImgRow
    
    If bToUserControl Then
        ' copia l'impronta nell'UserControl
        StretchBlt UserControl.hDC, 0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight, lhdc, 0, 0, iImgCol, iImgRow, SRCCOPY
        DeleteObject lhwnd
    Else
        ' restituisce l'immagine
        Set ToPicture = GDI.GetPicture(lhwnd)
    End If
    
    ' libera le risorse
    DeleteDC lhdc
    DrawDemoVersion
End Function

Private Sub DrawDemoVersion()
    ' fa apparire la scritta DEMO VERSION
    ' per la versione SHAREWARE del prodotto
    #If SHAREWARE = 1 Then
        
        On Local Error Resume Next
        
        Dim x As Long
        Dim isize As Integer
        Dim s As String
        
        x = 0
        isize = 32
        s = "DEMO VERSION"
        
        UserControl.FontBold = True
        Do While x <= 0
            isize = isize - 2
            UserControl.FontSize = isize
            x = (UserControl.ScaleWidth - UserControl.TextWidth(s)) / 2
        Loop
        
        UserControl.ForeColor = vbRed
        UserControl.CurrentX = x
        UserControl.CurrentY = (UserControl.ScaleHeight - UserControl.TextHeight(s)) / 2
        UserControl.Print s
        UserControl.ForeColor = vbBlack
    #End If
End Sub

'Private Function ZipBinary(aArray() As Byte) As Byte()
'    ' comprime un array di byte
'    ReDim tmpzip(UBound(aArray)) As Byte
'    CopyMemory tmpzip(0), aArray(0), UBound(aArray)
'    ZIP.CompressByteArray tmpzip, Z_BEST_COMPRESSION
'    ZipBinary = tmpzip
'End Function

'Private Function ZipBinary2Str(aArray() As Byte) As String
'    Dim stmp As String
'    Dim tmpzip() As Byte
'
'    ' comprime un array di byte e restituisce un stringa
'    tmpzip = ZipBinary(aArray)
'    stmp = String$(UBound(tmpzip), vbNullChar)
'    CopyMemory ByVal StrPtr(stmp), tmpzip(0), UBound(tmpzip)
'    ZipBinary2Str = stmp
'End Function

'Private Function UnzipBinary(aArray() As Byte) As Byte()
'    ' decomprime un array di byte
'    ReDim tmpzip(UBound(aArray)) As Byte
'    CopyMemory tmpzip(0), aArray(0), UBound(aArray)
'    ZIP.DecompressByteArray tmpzip, UBound(tmpzip) * 1.4142
'    UnzipBinary = tmpzip
'End Function
