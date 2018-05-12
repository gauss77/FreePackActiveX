VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cPicture24"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const LR_COPYRETURNORG = &H4
Private Const RGN_XOR = &H3

Private Type BITMAP
    type As Long
    Width As Long
    Height As Long
    WidthBytes As Long
    Planes As Integer
    BitsPixel As Integer
    Bits As Long
End Type

Private Type BITMAPBASEINFO
    Size As Long
    Width As Long
    Height As Long
    Planes As Integer
    BitCount As Integer
    Compression As Long
    SizeImage As Long
    XPelsPerMeter As Long
    YPelsPerMeter As Long
    ClrUsed As Long
    ClrImportant As Long
End Type

Private Type GUID
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(7) As Byte
End Type

Private Type PICBMPBASE
    Size As Long
    type As Long
    hBmp As Long
End Type

Private Type PICTDESC
   cbSizeOfStruct As Long
   picType As Long
   hgdiObj As Long
   hPalOrXYExt As Long
End Type

Private Type RGBQUAD
        rgbBlue As Byte
        rgbGreen As Byte
        rgbRed As Byte
        rgbReserved As Byte
End Type

Private Type BITMAPINFO
        bmiHeader As BITMAPBASEINFO
        bmiColors As RGBQUAD
End Type

Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function CreateDIBSection Lib "gdi32" (ByVal hDC As Long, pBitmapInfo As BITMAPBASEINFO, ByVal un As Long, ByVal lplpVoid As Long, ByVal handle As Long, ByVal dw As Long) As Long
Private Declare Function GetDIBits Lib "gdi32" (ByVal hDC As Long, ByVal hBitmap As Long, ByVal nStartScan As Long, ByVal nNumScans As Long, lpBits As Any, lpBI As BITMAPBASEINFO, ByVal wUsage As Long) As Long
Private Declare Function SetDIBitsToDevice Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long, ByVal dX As Long, ByVal dy As Long, ByVal SrcX As Long, ByVal SrcY As Long, ByVal Scan As Long, ByVal NumScans As Long, Bits As Any, BitsInfo As BITMAPBASEINFO, ByVal wUsage As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function GetObject Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long
Private Declare Function CopyImage Lib "user32" (ByVal handle As Long, ByVal imagetype As Long, ByVal NewWidth As Long, ByVal NewHeight As Long, ByVal lFlags As Long) As Long
Private Declare Function CreateRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function CombineRgn Lib "gdi32" (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long
Private Declare Function SetWindowRgn Lib "user32" (ByVal hwnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
Private Declare Function GetRegionData Lib "gdi32" (ByVal hRgn As Long, ByVal dwCount As Long, lpRgnData As Any) As Long
Private Declare Function ExtCreateRegion Lib "gdi32" (lpXform As Any, ByVal nCount As Long, lpRgnData As Any) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDst As Any, pSrc As Any, ByVal ByteLen As Long)
Private Declare Function OleCreatePictureIndirect Lib "olepro32.dll" (PicDesc As PICBMPBASE, RefIID As GUID, ByVal fPictureOwnsHandle As Long, IPic As IPicture) As Long

Private mBits() As Byte ' DIB bits buffer
Private mHeight As Long ' Height
Private mWidth As Long ' Width
Private mRgn As Long ' Region pointer

Public Event Progress(ByVal Percent As Integer)

Friend Property Set Picture(vData As StdPicture)
   Dim vBMP As BITMAP ' Basic image info
   Dim lBmp As Long ' Copy of image
   Dim lDC As Long ' Temporary device context
   Dim vBMI As BITMAPBASEINFO ' Basic DIB info
   
   Me.Reset
   If Not (vData Is Nothing) Then
      Call GetObject(vData.handle, Len(vBMP), vBMP)  ' Get bitmap info
      If (vBMP.Width > 0) And (vBMP.Height > 0) And (vBMP.BitsPixel >= 1) Then
         lBmp = CopyImage(vData.handle, 0, 0, 0, LR_COPYRETURNORG) ' Create temp copy
         lDC = CreateCompatibleDC(0) ' Create temp device context
         Call SelectObject(lDC, lBmp) ' Select image into temp dc
         ReDim mBits((vBMP.Width * 3) + CalcPadding(vBMP.Width) - 1, vBMP.Height - 1)
         With vBMI
            .Size = Len(vBMI) ' Set up for image
            .BitCount = 24
            .Width = vBMP.Width
            .Height = -vBMP.Height
            .Planes = 1
            .SizeImage = (UBound(mBits, 1) + 1) * (UBound(mBits, 2) + 1)
         End With
         Call GetDIBits(lDC, lBmp, 0, vBMP.Height, mBits(0, 0), vBMI, 0) ' Get bits
         mWidth = vBMP.Width ' Store width
         mHeight = vBMP.Height ' Store height
         Call DeleteObject(lBmp) ' Delete temp image
         Call DeleteDC(lDC) ' Delete temp dc
      End If
   End If
End Property

Friend Property Get Picture() As StdPicture
   Dim lDC As Long ' Temporary device context
   Dim vBMI As BITMAPBASEINFO ' Basic DIB info
   Dim lBmp As Long ' Copy of image
    
   If mWidth > 0 Then
        lDC = CreateCompatibleDC(0) ' Create temp dc
        With vBMI ' Set up for image
            .BitCount = 24
            .Planes = 1
            .Width = mWidth
            .Height = -mHeight
            .SizeImage = (UBound(mBits, 1) + 1) * (UBound(mBits, 2) + 1)
            .Size = Len(vBMI)
        End With
        lBmp = CreateDIBSection(lDC, vBMI, 0, 0, 0, 0) ' Create a temp blank image
        Call SelectObject(lDC, lBmp) ' Select image into temp DC
        Call SetDIBitsToDevice(lDC, 0, 0, mWidth, mHeight, 0, 0, 0, mHeight, mBits(0, 0), vBMI, 0) ' Set bits
        Set Picture = GetPicture(lBmp)
        Call DeleteDC(lDC) ' Delete temp dc
    Else
        Set Picture = Nothing
    End If
End Property

Friend Function Resize(ByVal NewWidth As Long, ByVal NewHeight As Long) As Boolean
   Dim lDC As Long ' Temporary device context
   Dim vBMI As BITMAPBASEINFO ' Basic DIB info
   Dim lBmp As Long ' Copy of image
   Dim lNewBmp As Long ' Resized copy of image
   
   If mWidth > 0 Then
      lDC = CreateCompatibleDC(0) ' Create temp dc
      With vBMI ' Setup for image
         .BitCount = 24
         .Planes = 1
         .Width = mWidth
         .Height = -mHeight
         .SizeImage = (UBound(mBits, 1) + 1) * (UBound(mBits, 2) + 1)
         .Size = Len(vBMI)
      End With
      lBmp = CreateDIBSection(lDC, vBMI, 0, 0, 0, 0) ' Create temp blank image
      Call SelectObject(lDC, lBmp) ' Select the image into the dc
      Call SetDIBitsToDevice(lDC, 0, 0, mWidth, mHeight, 0, 0, 0, mHeight, mBits(0, 0), vBMI, 0) ' Set bits
      
      ' Reset
      Erase mBits
      mWidth = 0
      mHeight = 0
      Me.DeleteRegion
      
      lNewBmp = CopyImage(lBmp, 0, NewWidth, NewHeight, LR_COPYRETURNORG) ' Copy resized temp image
      Call DeleteObject(lBmp) ' Delete temp image
      Call SelectObject(lDC, lNewBmp) ' Select new image into temp dc
      ReDim mBits((NewWidth * 3) + CalcPadding(NewWidth) - 1, NewHeight - 1)
      With vBMI ' Setup for resized image
         .Width = NewWidth
         .Height = -NewHeight
         .SizeImage = (UBound(mBits, 1) + 1) * (UBound(mBits, 2) + 1)
      End With
      Call GetDIBits(lDC, lNewBmp, 0, NewHeight, mBits(0, 0), vBMI, 0) ' Get new bits
      mWidth = NewWidth ' Store width
      mHeight = NewHeight ' Store height
      Resize = True ' Return true
      Call DeleteObject(lNewBmp) ' Delete new image
      Call DeleteDC(lDC) ' Delete temp dc
   End If
End Function

Friend Function Crop(ByVal x As Long, ByVal y As Long, ByVal Width As Long, ByVal Height As Long) As Boolean
    Dim lDC As Long ' Temporary device context
    Dim vBMI As BITMAPBASEINFO ' Basic DIB info
    Dim lBmp As Long ' Copy of original image
    Dim lNewBmp As Long ' Cropped copy of image
    Dim lNewDC As Long ' Cropped image's temporary device context
    
    If mWidth > 0 Then
        lDC = CreateCompatibleDC(0) ' Create temp dc
        With vBMI ' Setup for image
            .BitCount = 24
            .Planes = 1
            .Width = mWidth
            .Height = -mHeight
            .SizeImage = UBound(mBits) + 1
            .Size = Len(vBMI)
        End With
        lBmp = CreateDIBSection(lDC, vBMI, 0, 0, 0, 0) ' Create temp blank image
        Call SelectObject(lDC, lBmp) ' Select image into temp dc
        Call SetDIBitsToDevice(lDC, 0, 0, mWidth, mHeight, 0, 0, 0, mHeight, mBits(0, 0), vBMI, 0) ' Set bits
        
        ' Reset
        Erase mBits
        mWidth = 0
        mHeight = 0
        Me.DeleteRegion
        
        lNewDC = CreateCompatibleDC(0) ' Create new temp dc
         ReDim mBits((Width * 3) + CalcPadding(Width) - 1, Height - 1)
         With vBMI ' Setup for cropped image
            .SizeImage = (UBound(mBits, 1) + 1) * (UBound(mBits, 2) + 1)
            .Width = Width
            .Height = -Height
        End With
        lNewBmp = CreateDIBSection(lDC, vBMI, 0, 0, 0, 0) ' Create new blank image
        SelectObject lNewDC, lNewBmp ' Select blank into new dc
        Call BitBlt(lNewDC, 0, 0, Width, Height, lDC, x, y, vbSrcCopy) ' Transfer specified section
        DeleteObject lBmp ' Delete copy
        DeleteDC lDC ' Delete temp dc
         Call GetDIBits(lNewDC, lNewBmp, 0, Height, mBits(0, 0), vBMI, 0)  ' Get new bits
         mWidth = Width ' Store width
         mHeight = Height ' Store height
         Crop = True ' Return true
        Call DeleteObject(lNewBmp) ' Delete cropped image
        Call DeleteDC(lNewDC) ' Delete new dc
    End If
End Function

Friend Sub FitInto(ByVal Width As Long, ByVal Height As Long, Optional ByVal EnlargeOK As Boolean)
   Dim nP As Single
   If mWidth > 0 Then
      If (Width <> mWidth) Or (Height <> mHeight) Then
         nP = mWidth / mHeight
         If nP < (Width / Height) Then
            Width = Height * nP
         Else
            Height = Width / nP
         End If
         If ((Width <= mWidth) And (Height <= mHeight)) Or EnlargeOK Then
            Me.Resize Width, Height
         End If
      End If
   End If
End Sub

Friend Sub Rotate()
    Dim bTmp() As Byte
    Dim x As Long
    Dim y As Long
    Dim lTmp As Long
    
    If mWidth > 0 Then
        ReDim bTmp((mHeight * 3) + CalcPadding(mHeight) - 1, mWidth - 1)
        For y = 0 To mHeight - 1
            For x = 0 To mWidth - 1
                CopyMemory bTmp(y * 3, x), mBits(x * 3, mHeight - y - 1), 3
            Next
            RaiseEvent Progress(y / (mHeight - 1) * 100)
        Next
        
        lTmp = mWidth
        mWidth = mHeight
        mHeight = lTmp
        mBits = bTmp
        Erase bTmp
    End If
End Sub

Friend Sub FlipX()
    Dim bBits() As Byte
    Dim x As Long
    Dim y As Long
    
    If mWidth > 0 Then
        bBits = mBits
        For y = 0 To mHeight - 1
            For x = 0 To mWidth - 1
                CopyMemory mBits(x * 3, y), bBits((mWidth - x - 1) * 3, y), 3
            Next
        
            RaiseEvent Progress(y / (mHeight - 1) * 100)
        Next
        Erase bBits
    End If
End Sub

Friend Sub FlipY()
   Dim bBits() As Byte
   Dim y As Long
   
   If mWidth > 0 Then
      bBits = mBits
      For y = 0 To mHeight - 1
        CopyMemory mBits(0, y), bBits(0, mHeight - y - 1), mWidth * 3
        RaiseEvent Progress(y / (mHeight - 1) * 100)
      Next
      Erase bBits
   End If
End Sub

Friend Sub Invert(Optional ByVal Red As Boolean = True, Optional ByVal Green As Boolean = True, Optional ByVal Blue As Boolean = True)
    Dim x As Long
    Dim y As Long
    Dim bRGB() As Byte
   
    If (mWidth > 0) And (Red Or Green Or Blue) Then
        For y = 0 To mHeight - 1
            For x = 0 To mWidth - 1
                bRGB = PixelB(x, y)
                If Red Then bRGB(2) = 255 - bRGB(2)
                If Green Then bRGB(1) = 255 - bRGB(1)
                If Blue Then bRGB(0) = 255 - bRGB(0)
                PixelB(x, y) = bRGB
            Next
            RaiseEvent Progress(y / (mHeight - 1) * 100)
        Next
    End If
End Sub

Friend Property Let Bits(vData() As Byte)
    mBits = vData
End Property

Friend Property Get Bits() As Byte()
    Bits = mBits
End Property

Friend Property Get Width() As Long
    Width = mWidth
End Property

Friend Property Get Height() As Long
    Height = mHeight
End Property

Friend Property Let Pixel(ByVal x As Long, ByVal y As Long, ByVal vData As Long)
   CopyMemory mBits(x * 3, y), vData, 3
End Property

Friend Property Get Pixel(ByVal x As Long, ByVal y As Long) As Long
   Dim lMem As Long
   CopyMemory lMem, mBits(x * 3, y), 3
   Pixel = lMem
End Property

Friend Property Let PixelB(ByVal x As Long, ByVal y As Long, vData() As Byte)
   CopyMemory mBits(x * 3, y), vData(0), 3
End Property

Friend Property Get PixelB(ByVal x As Long, ByVal y As Long) As Byte()
   Dim bMem(0 To 2) As Byte
   CopyMemory bMem(0), mBits(x * 3, y), 3
   PixelB = bMem
End Property

Friend Property Let PixelL(ByVal x As Long, ByVal y As Long, vData() As Integer)
    Dim i As Integer
    Dim bData(0 To 2) As Byte
    For i = 0 To 2
        If vData(i) > 255 Then vData(i) = 255 Else If vData(i) < 0 Then vData(i) = 0
        bData(i) = CByte(vData(i))
    Next
    CopyMemory mBits(x * 3, y), bData(0), 3
End Property

Friend Property Get PixelL(ByVal x As Long, ByVal y As Long) As Integer()
    Dim i As Byte
    Dim iData(0 To 2) As Integer
    Dim bData(0 To 2) As Byte
    CopyMemory bData(0), mBits(x * 3, y), 3
    For i = 0 To 2
        iData(i) = CInt(bData(i))
    Next
    PixelL = iData
End Property

Friend Sub CreateRegion(ByVal TransColor As ColorConstants)
    Dim lX As Long ' X iteration
    Dim lY As Long ' Y iteration
    Dim lXa As Long ' XOR region start
    Dim lXb As Long ' XOR region end
    Dim lXor As Long ' XOR region
    
    ' Reset region
    Me.DeleteRegion
    
   If mWidth > 0 Then
      mRgn = CreateRectRgn(0, 0, mWidth, mHeight) ' Create bounding rect region
      For lY = 0 To mHeight - 1
         lXa = -1
         For lX = 0 To mWidth - 1
            If Me.Pixel(lX, lY) = TransColor Then ' Check for pixel transparency
               If lXa = -1 Then lXa = lX ' Set start if needed
               lXb = lX + 1 ' Set/reset end
               If lX = (mWidth - 1) Then ' Check for last horizontal pixel
                  lXor = CreateRectRgn(lXa, lY, lXb, lY + 1) ' Create XOR region (point or line)
                  Call CombineRgn(mRgn, lXor, mRgn, RGN_XOR) ' XOR from rect region
                  Call DeleteObject(lXor) ' Delete XOR region
               End If
            Else ' Pixel not transparent
               If lXa > -1 Then ' Check for transparent pixels found
                  lXor = CreateRectRgn(lXa, lY, lXb, lY + 1) ' Create XOR region (point or line)
                  Call CombineRgn(mRgn, lXor, mRgn, RGN_XOR) ' XOR from rect region
                  Call DeleteObject(lXor) ' Delete XOR region
               End If
               lXa = -1 ' Reset start for following pixels in scanline
            End If
         Next
      Next
    End If
End Sub

Friend Sub CreateBlank(ByVal Width As Long, ByVal Height As Long, Optional ByVal Fill)
   Dim x As Long
   Dim y As Long
   Dim lFill As Long
   Dim bFill() As Byte
   
   If (Width > 0) And (Height > 0) Then
      Erase mBits
      DeleteRegion
      ReDim mBits((Width * 3) + CalcPadding(Width) - 1, Height - 1)
      mWidth = Width
      mHeight = Height
      If Not IsMissing(Fill) Then
         If TypeName(Fill) = "Byte()" Then
            bFill = Fill
            For y = 0 To Height - 1
                For x = 0 To Width - 1
                   CopyMemory mBits(x * 3, y), bFill(0), 3
                Next
            Next
            Erase bFill
         Else
            lFill = Fill
            For y = 0 To Height - 1
                For x = 0 To Width - 1
                   CopyMemory mBits(x * 3, y), lFill, 3
                Next
            Next
         End If
      End If
   End If
End Sub

Friend Function HasRegion() As Boolean
   HasRegion = (mRgn <> 0)
End Function

Friend Sub ApplyRegion(ByVal hwnd As Long)
    Dim lData As Long ' Size of region data
    Dim bData() As Byte ' Region data buffer
    Dim lRgn As Long ' Temporary region copy
    
    If mRgn <> 0 Then ' Check for region
        lData = GetRegionData(mRgn, 0, ByVal 0&) ' Get data size
        ReDim bData(lData - 1) As Byte ' Allocate region buffer
        Call GetRegionData(mRgn, lData, bData(0)) ' Get region data
        lRgn = ExtCreateRegion(ByVal 0&, lData, bData(0)) ' Create temp copy
        Erase bData ' Empty region buffer
        Call SetWindowRgn(hwnd, lRgn, True) ' Set window's region
        Call DeleteObject(lRgn) ' Delete temp region
    End If
End Sub

Friend Sub ClearRegion(ByVal hwnd As Long)
    Call SetWindowRgn(hwnd, 0, True) ' Clear window's region
End Sub

Friend Sub DeleteRegion()
    If mRgn <> 0 Then ' Check for region
        Call DeleteObject(mRgn) ' Delete region
        mRgn = 0 ' Clear pointer
    End If
End Sub

Friend Function ComparePicture(Picture) As Boolean
   Dim oPic As cPicture24
   Dim bBits() As Byte
   Dim sBits1 As String
   Dim sBits2 As String
   
   Select Case TypeName(Picture)
      Case "StdPicture", "IPictureDsp", "Picture"
         Set oPic = New cPicture24
         oPic.Picture = Picture
      Case "cPicture24"
         Set oPic = Picture
   End Select
   If Not (oPic Is Nothing) Then
      If (oPic.Width = mWidth) And (oPic.Height = mHeight) Then
         sBits1 = Space((UBound(mBits, 1) + 1) * (UBound(mBits, 2) + 1))
         CopyMemory ByVal sBits1, mBits(0, 0), Len(sBits1)
         bBits = oPic.Bits
         Set oPic = Nothing
         sBits2 = Space((UBound(bBits, 1) + 1) * (UBound(bBits, 2) + 1))
         CopyMemory ByVal sBits2, bBits(0, 0), Len(sBits2)
         Erase bBits
         ComparePicture = (sBits1 = sBits2)
         sBits1 = Empty
         sBits2 = Empty
      Else
         Set oPic = Nothing
      End If
   End If
End Function

Friend Property Get Padding() As Integer
   Padding = CalcPadding(mWidth)
End Property

Friend Sub Reset()
   Erase mBits
   mWidth = 0
   mHeight = 0
   Me.DeleteRegion
End Sub

Friend Sub GrayScale()
    Dim x As Long
    Dim y As Long
    Dim lGray As Long
    Dim bRGB() As Byte
   
    For y = 0 To mHeight - 1
        For x = 0 To mWidth - 1
            bRGB = PixelB(x, y)
            
            lGray = (222 * CLng(bRGB(2)) + 707 * CLng(bRGB(1)) + 71 * CLng(bRGB(0))) * 0.001
            
            ' International Telecommunications Union standard - recommended Shade = (0.2125 * Red + 0.7154 * Green + 0.0721 * Blue)
            ' NTSC and PAL Shade = (0.299 * Red + 0.587 * Green + 0.114 * Blue)
            ' Simple average Shade = (Red + Green + Blue) / 3
            ' Weighted average - common Shade = (3 * Red + 4 * Green + 2 * Blue) / 9
            ' Distance of color vector in color cube - not recommended Shade = Sqr((Red ^ 2 + Green ^ 2 + Blue ^ 2) / 3)
            ' Human eye responsive - not recommended (ignores red & blue) Shade = Green Anyway
            
            bRGB(0) = lGray
            bRGB(1) = lGray
            bRGB(2) = lGray
            
            PixelB(x, y) = bRGB
        Next
    
        RaiseEvent Progress(y / (mHeight - 1) * 100)
    Next
End Sub

Friend Sub BlackWhite()
    Dim x As Long
    Dim y As Long
    Dim bRGB() As Byte
   
    For y = 0 To mHeight - 1
        For x = 0 To mWidth - 1
            bRGB = PixelB(x, y)
            If bRGB(0) >= 192 Or bRGB(1) >= 192 Or bRGB(2) >= 192 Then
                bRGB(0) = 255
                bRGB(1) = 255
                bRGB(2) = 255
            Else
                bRGB(0) = 0
                bRGB(1) = 0
                bRGB(2) = 0
            End If
            PixelB(x, y) = bRGB
        Next
    
        RaiseEvent Progress(y / (mHeight - 1) * 100)
    Next
End Sub

Friend Sub Sharpen(ByVal lAmount As Single)
    Dim x As Long
    Dim y As Long
    Dim iRGB() As Integer
    Dim iRGB2() As Integer
   
    For y = 0 To mHeight - 2
        For x = 0 To mWidth - 2
            iRGB = PixelL(x, y)
            iRGB2 = PixelL(x + 1, y + 1)
            iRGB(0) = iRGB(0) + lAmount * (iRGB(0) - iRGB2(0))
            iRGB(1) = iRGB(1) + lAmount * (iRGB(1) - iRGB2(1))
            iRGB(2) = iRGB(2) + lAmount * (iRGB(2) - iRGB2(2))
            PixelL(x, y) = iRGB
        Next
    
        RaiseEvent Progress(y / (mHeight - 1) * 100)
    Next
End Sub

Friend Sub Blur()
    Dim x As Long
    Dim y As Long
    Dim iRGB() As Integer
    Dim iRGB1() As Integer
    Dim iRGB2() As Integer
    Dim iRGB3() As Integer
    Dim iRGB4() As Integer
   
    For y = 1 To mHeight - 2
        For x = 1 To mWidth - 2
            iRGB = PixelL(x, y)
            iRGB1 = PixelL(x - 1, y - 1)
            iRGB2 = PixelL(x + 1, y + 1)
            iRGB3 = PixelL(x + 1, y)
            iRGB4 = PixelL(x, y + 1)
            iRGB(0) = (iRGB(0) + iRGB2(0) + iRGB1(0) + iRGB3(0) + iRGB4(0)) * 0.2
            iRGB(1) = (iRGB(1) + iRGB2(1) + iRGB1(1) + iRGB3(1) + iRGB4(1)) * 0.2
            iRGB(2) = (iRGB(2) + iRGB2(2) + iRGB1(2) + iRGB3(2) + iRGB4(2)) * 0.2
            PixelL(x, y) = iRGB
        Next
    
        RaiseEvent Progress(y / (mHeight - 1) * 100)
    Next
End Sub

Friend Sub Contrast(ByVal lAmount As Integer)
    Dim x As Long
    Dim y As Long
    Dim iRGB() As Integer
   
    For y = 0 To mHeight - 2
        For x = 0 To mWidth - 2
            iRGB = PixelL(x, y)
            iRGB(0) = ((iRGB(0) - 128) * lAmount) + 128
            iRGB(1) = ((iRGB(1) - 128) * lAmount) + 128
            iRGB(2) = ((iRGB(2) - 128) * lAmount) + 128
            PixelL(x, y) = iRGB
        Next
    
        RaiseEvent Progress(y / (mHeight - 1) * 100)
    Next
End Sub

Friend Sub Fade(ByVal lAmount As Long)
    Dim x As Long
    Dim y As Long
    Dim bRGB() As Byte
   
    For y = 0 To mHeight - 1
        For x = 0 To mWidth - 1
            bRGB = PixelB(x, y)
            
            bRGB(0) = lAmount * bRGB(0) * 3.92156862745098E-03
            bRGB(1) = lAmount * bRGB(1) * 3.92156862745098E-03
            bRGB(2) = lAmount * bRGB(2) * 3.92156862745098E-03
            
            PixelB(x, y) = bRGB
        Next
    
        RaiseEvent Progress(y / (mHeight - 1) * 100)
    Next
End Sub

Friend Sub Lighten(ByVal lAmount As Long)
    Dim x As Long
    Dim y As Long
    Dim lB As Long
    Dim lG As Long
    Dim lR As Long
    Dim h As Single
    Dim s As Single
    Dim l As Single
    Dim bRGB() As Byte
   
    For y = 0 To mHeight - 1
        For x = 0 To mWidth - 1
            bRGB = PixelB(x, y)
            
            RGBToHLS bRGB(2), bRGB(1), bRGB(0), h, s, l
            l = l * (1 + (lAmount / 100))
            If (l > 1) Then l = 1
            HLSToRGB h, s, l, lR, lG, lB
            bRGB(0) = lB
            bRGB(1) = lG
            bRGB(2) = lR
            
            PixelB(x, y) = bRGB
        Next
        RaiseEvent Progress(y / (mHeight - 1) * 100)
    Next
End Sub

Private Function CalcPadding(ByVal Width As Long) As Integer
   CalcPadding = ((((Width * 3) + 3) \ 4) * 4) - (Width * 3)
End Function

Private Function GetPicture(ByVal lHandle As Long) As StdPicture

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

Private Sub RGBToHLS(ByVal r As Long, ByVal g As Long, ByVal b As Long, h As Single, s As Single, l As Single)
    
    Dim Max As Single
    Dim Min As Single
    Dim delta As Single
    Dim rR As Single
    Dim rG As Single
    Dim rB As Single

    rR = r * 3.92156862745098E-03
    rG = g * 3.92156862745098E-03
    rB = b * 3.92156862745098E-03

    '{Given: rgb each in [0,1].
    ' Desired: h in [0,360] and s in [0,1], except if s=0, then h=UNDEFINED.}
    Max = Maximum(rR, rG, rB)
    Min = Minimum(rR, rG, rB)
    l = (Max + Min) * 0.5 '{This is the lightness}
    '{Next calculate saturation}
    If Max = Min Then
        'begin {Acrhomatic case}
        s = 0
        h = 0
        'end {Acrhomatic case}
        Else
        'begin {Chromatic case}
        '{First calculate the saturation.}
        If l <= 0.5 Then
            s = (Max - Min) / (Max + Min)
        Else
            s = (Max - Min) / (2 - Max - Min)
        End If
        '{Next calculate the hue.}
        delta = Max - Min
        If rR = Max Then
            h = (rG - rB) / delta    '{Resulting color is between yellow and magenta}
        ElseIf rG = Max Then
            h = 2 + (rB - rR) / delta '{Resulting color is between cyan and yellow}
        ElseIf rB = Max Then
            h = 4 + (rR - rG) / delta '{Resulting color is between magenta and cyan}
        End If
        'Debug.Print h
        'h = h * 60
        'If h < 0# Then
        '     h = h + 360            '{Make degrees be nonnegative}
        'End If
        'end {Chromatic Case}
    End If
    'end {RGB_to_HLS}
End Sub

Private Sub HLSToRGB(ByVal h As Single, ByVal s As Single, ByVal l As Single, r As Long, g As Long, b As Long)

    Dim rR As Single
    Dim rG As Single
    Dim rB As Single
    Dim Min As Single
    Dim Max As Single
    
    If s = 0 Then
       ' Achromatic case:
       rR = l
       rG = l
       rB = l
    Else
       ' Chromatic case:
       ' delta = Max-Min
       If l <= 0.5 Then
          's = (Max - Min) / (Max + Min)
          ' Get Min value:
          Min = l * (1 - s)
       Else
          's = (Max - Min) / (2 - Max - Min)
          ' Get Min value:
          Min = l - s * (1 - l)
       End If
       ' Get the Max value:
       Max = 2 * l - Min
       
       ' Now depending on sector we can evaluate the h,l,s:
       If (h < 1) Then
          rR = Max
          If (h < 0) Then
             rG = Min
             rB = rG - h * (Max - Min)
          Else
             rB = Min
             rG = h * (Max - Min) + rB
          End If
       ElseIf (h < 3) Then
          rG = Max
          If (h < 2) Then
             rB = Min
             rR = rB - (h - 2) * (Max - Min)
          Else
             rR = Min
             rB = (h - 2) * (Max - Min) + rR
          End If
       Else
          rB = Max
          If (h < 4) Then
             rR = Min
             rG = rR - (h - 4) * (Max - Min)
          Else
             rG = Min
             rR = (h - 4) * (Max - Min) + rG
          End If
          
       End If
             
    End If
    r = rR * 255
    g = rG * 255
    b = rB * 255
End Sub

Private Function Maximum(rR As Single, rG As Single, rB As Single) As Single
    If (rR > rG) Then
       If (rR > rB) Then
          Maximum = rR
       Else
          Maximum = rB
       End If
    Else
       If (rB > rG) Then
          Maximum = rB
       Else
          Maximum = rG
       End If
    End If
End Function

Private Function Minimum(rR As Single, rG As Single, rB As Single) As Single
    If (rR < rG) Then
       If (rR < rB) Then
          Minimum = rR
       Else
          Minimum = rB
       End If
    Else
       If (rB < rG) Then
          Minimum = rB
       Else
          Minimum = rG
       End If
    End If
End Function

Private Sub Class_Terminate()
    Erase mBits ' Empty buffer
    DeleteRegion ' Delete region, if any
End Sub
