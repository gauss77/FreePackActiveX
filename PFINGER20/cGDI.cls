VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cGDI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Enum IMAGETYPEENUM
    IMAGETYPE_BITMAP = 0
    IMAGETYPE_JPEG = 1
    IMAGETYPE_GIF = 2
    IMAGETYPE_TIFF = 3
    IMAGETYPE_PNG = 4
End Enum

Private Type GUID
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(0 To 7) As Byte
End Type

Private Type GdiplusStartupInput
    GdiplusVersion As Long
    DebugEventCallback As Long
    SuppressBackgroundThread As Long
    SuppressExternalCodecs As Long
End Type

Private Type EncoderParameter
    GUID As GUID
    NumberOfValues As Long
    type As Long
    Value As Long
End Type

Private Type EncoderParameters
    Count As Long
    Parameter As EncoderParameter
End Type

Private Type PICTDESC
   cbSizeOfStruct As Long
   picType As Long
   hgdiObj As Long
   hPalOrXYExt As Long
End Type

Private Type BITMAP
    type As Long
    Width As Long
    Height As Long
    WidthBytes As Long
    Planes As Integer
    BitsPixel As Integer
    Bits As Long
End Type

Private Type PICBMPBASE
    Size As Long
    type As Long
    hBmp As Long
End Type

Private m_hWnd As Long

Private Declare Function GetObject Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long
Private Declare Function GdiplusStartup Lib "gdiplus.dll" (token As Long, inputbuf As GdiplusStartupInput, Optional ByVal outputbuf As Long = 0) As Long
Private Declare Function GdiplusShutdown Lib "gdiplus.dll" (ByVal token As Long) As Long
Private Declare Function GdipCreateBitmapFromHBITMAP Lib "gdiplus.dll" (ByVal hbm As Long, ByVal hPal As Long, BITMAP As Long) As Long
Private Declare Function GdipDisposeImage Lib "gdiplus.dll" (ByVal image As Long) As Long
Private Declare Function GdipSaveImageToFile Lib "gdiplus.dll" (ByVal image As Long, ByVal filename As Long, clsidEncoder As GUID, encoderParams As Any) As Long
Private Declare Function CLSIDFromString Lib "ole32.dll" (ByVal str As Long, id As GUID) As Long
Private Declare Function GdipCreateBitmapFromFile Lib "GDIPlus" (ByVal filename As Long, BITMAP As Long) As Long
Private Declare Function GdipCreateHBITMAPFromBitmap Lib "GDIPlus" (ByVal BITMAP As Long, hbmReturn As Long, ByVal background As Long) As Long
Private Declare Function OleCreatePictureIndirect Lib "olepro32.dll" (PicDesc As PICBMPBASE, RefIID As GUID, ByVal fPictureOwnsHandle As Long, IPic As IPicture) As Long
Private Declare Function SetPixel Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long, ByVal crColor As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Private Declare Function GetPixel Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hDC As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long

Friend Property Get hwnd() As Long
    hwnd = m_hWnd
End Property

Friend Property Let hwnd(ByVal lHandle As Long)
    m_hWnd = lHandle
End Property

Friend Property Get Picture() As StdPicture
    If m_hWnd = 0 Then
        Set Picture = Nothing
    Else
        Set Picture = GetPicture(m_hWnd)
    End If
End Property

Friend Property Set Picture(pPicture As StdPicture)
    If pPicture Is Nothing Then m_hWnd = 0 Else m_hWnd = pPicture.handle
End Property

Friend Function SaveImage(filename As String, _
    Optional ByVal imagetype As IMAGETYPEENUM = IMAGETYPE_BITMAP, _
    Optional ByVal Quality As Byte = 70) As Long
   
    Dim tSI As GdiplusStartupInput
    Dim lRes As Long
    Dim lGDIP As Long
    Dim lBitmap As Long
    Dim sID As String
    
    ' Initialize GDI+
    tSI.GdiplusVersion = 1
    lRes = GdiplusStartup(lGDIP, tSI)
    
    If lRes = 0 Then
    
        ' Create the GDI+ bitmap
        ' from the image handle
        lRes = GdipCreateBitmapFromHBITMAP(m_hWnd, 0, lBitmap)
        
        If lRes = 0 Then
            Dim tImgEncoder As GUID
            Dim tParams As EncoderParameters
            
            Select Case imagetype
                Case IMAGETYPE_JPEG
                    sID = "{557cf401-1a04-11D3-9A73-0000f81ef32e}"
                Case IMAGETYPE_GIF
                    sID = "{557cf402-1a04-11d3-9a73-0000f81ef32e}"
                Case IMAGETYPE_TIFF
                    sID = "{557cf405-1a04-11d3-9a73-0000f81ef32e}"
                Case IMAGETYPE_PNG
                    sID = "{557cf406-1a04-11d3-9a73-0000f81ef32e}"
                Case Else
                    sID = "{557cf400-1a04-11d3-9a73-0000f81ef32e}"
            End Select
            
            ' Initialize the encoder GUID
            CLSIDFromString StrPtr(sID), tImgEncoder
            
            ' Initialize the encoder parameters
            If imagetype = IMAGETYPE_JPEG Or IMAGETYPE_PNG Then
                tParams.Count = 1
                With tParams.Parameter ' Quality
                    ' Set the Quality GUID
                    CLSIDFromString StrPtr("{1D5BE4B5-FA4A-452D-9CDD-5DB35105E7EB}"), .GUID
                    .NumberOfValues = 1
                    .type = 4
                    .Value = VarPtr(Quality)
                End With
            End If
            
            ' Save the image
            lRes = GdipSaveImageToFile(lBitmap, StrPtr(filename), tImgEncoder, tParams)
            
            ' Destroy the bitmap
            GdipDisposeImage lBitmap
            
        End If
            
        ' Shutdown GDI+
        GdiplusShutdown lGDIP
    
    End If
    SaveImage = lRes
End Function

Friend Function LoadImage(filename As String) As StdPicture
    
    Dim tSI As GdiplusStartupInput
    Dim lGDIP As Long
    Dim lRes As Long
    Dim lBitmap As Long
    
    ' Initialize GDI+
    tSI.GdiplusVersion = 1
    lRes = GdiplusStartup(lGDIP, tSI)
    
    If lRes = 0 Then
        ' Open the image file
        lRes = GdipCreateBitmapFromFile(StrPtr(filename), lBitmap)
        If lRes = 0 Then
            ' Create a GDI bitmap
            lRes = GdipCreateHBITMAPFromBitmap(lBitmap, m_hWnd, 0)
            ' Create the StdPicture object
            Set LoadImage = GetPicture(m_hWnd)
            ' Dispose the image
            GdipDisposeImage lBitmap
        End If
    
        ' Shutdown GDI+
        GdiplusShutdown lGDIP
    End If
    If lRes <> 0 Then Set LoadImage = Nothing
End Function

Friend Function GetImageWidth() As Long
    Dim tBI As BITMAP
    
    ' recupera le dimensioni orizzontali effettive dell'immagine
    Call GetObject(m_hWnd, Len(tBI), tBI)
    GetImageWidth = tBI.Width
End Function

Friend Function GetImageHeight() As Long
    Dim tBI As BITMAP
    
    ' recupera le dimensioni verticali effettive dell'immagine
    Call GetObject(m_hWnd, Len(tBI), tBI)
    GetImageHeight = tBI.Height
End Function

Friend Sub DrawRawImage(ByVal lhdc As Long, ByRef RawImage() As Byte, Optional ByVal lWidth As Long = 300, Optional ByVal lHeight As Long = 300)
    Dim c As Long
    Dim iy As Long
    Dim ix As Long
    
    c = -1
    For iy = 0 To lHeight - 1
        For ix = 0 To lWidth - 1
            c = c + 1
            SetPixel lhdc, ix, iy, RGB(RawImage(c), RawImage(c), RawImage(c))
        Next
    Next
End Sub

Friend Function GetPicture(ByVal lHandle As Long) As StdPicture

    Dim vPic As PICBMPBASE
    Dim vIDispatch As GUID
    
    ' Initialize the PICTDESC structure
    With vPic
        .Size = Len(vPic)
        .type = vbPicTypeBitmap
        .hBmp = lHandle
    End With
    
    ' Initialize the IPicture interface ID
    With vIDispatch
        .Data1 = &H7BF80981
        .Data2 = &HBF32
        .Data3 = &H101A
        .Data4(0) = &H8B
        .Data4(1) = &HBB
        .Data4(3) = &HAA
        .Data4(5) = &H30
        .Data4(6) = &HC
        .Data4(7) = &HAB
    End With
    
    ' Create the object
    OleCreatePictureIndirect vPic, vIDispatch, 1, GetPicture
End Function

Friend Function PictureToRaw(ByRef Picture As StdPicture) As Byte()
    Dim c As Long
    Dim iy As Long
    Dim ix As Long
    Dim RawData() As Byte
    Dim lhdc As Long
    Dim tBI As BITMAP
    
    ' crea un DC
    lhdc = CreateCompatibleDC(0)
    SelectObject lhdc, Picture.handle
    
    ' recupera le dimensioni dell'immagine
    Call GetObject(Picture.handle, Len(tBI), tBI)
    
    ' ridimensiona il buffer
    c = -1
    ReDim RawData(tBI.Height * tBI.Width)
    
    ' riempie il buffer
    For iy = 0 To tBI.Height - 1
        For ix = 0 To tBI.Width - 1
            c = c + 1
            RawData(c) = GetPixel(lhdc, ix, iy) / 65793
        Next
    Next
    PictureToRaw = RawData  ' restituisce il buffer
    DeleteDC lhdc           ' libera risorse
End Function

Friend Function RawToPicture(ByRef RawData() As Byte, Optional ByVal lWidth As Long = 300, Optional ByVal lHeight As Long = 300) As StdPicture
    Dim c As Long
    Dim iy As Long
    Dim ix As Long
    Dim lColor As Long
    Dim lhdc As Long
    Dim lhwnd As Long
    
    ' crea una bitmap in memoria per il disegno
    lhdc = CreateCompatibleDC(0)
    lhwnd = CreateCompatibleBitmap(lhdc, lWidth, lHeight)
    SelectObject lhdc, lhwnd

    ' disegna l'impronta
    DrawRawImage lhdc, RawData, lWidth, lHeight
    
    ' restituisce l'immagine
    Set RawToPicture = GetPicture(lhwnd)
    
    ' libera le risorse
    DeleteDC lhdc
End Function

