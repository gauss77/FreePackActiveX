Private Const CCHDEVICENAME = 32
Private Const CCHFORMNAME = 32

Private Type DOCINFO
        'cbSize As Long
        lpszDocName As String
        lpszOutput As String
        lpszDataType As String
End Type

Private Type DEVMODE
        dmDeviceName As String * CCHDEVICENAME
        dmSpecVersion As Integer
        dmDriverVersion As Integer
        dmSize As Integer
        dmDriverExtra As Integer
        dmFields As Long
        dmOrientation As Integer
        dmPaperSize As Integer
        dmPaperLength As Integer
        dmPaperWidth As Integer
        dmScale As Integer
        dmCopies As Integer
        dmDefaultSource As Integer
        dmPrintQuality As Integer
        dmColor As Integer
        dmDuplex As Integer
        dmYResolution As Integer
        dmTTOption As Integer
        dmCollate As Integer
        dmFormName As String * CCHFORMNAME
        dmUnusedPadding As Integer
        dmBitsPerPel As Long
        dmPelsWidth As Long
        dmPelsHeight As Long
        dmDisplayFlags As Long
        dmDisplayFrequency As Long
End Type

Private Type PRINTER_DEFAULTS
        pDatatype As String
        pDevMode As DEVMODE
        DesiredAccess As Long
End Type

Private Declare Function WritePrinter Lib "winspool.drv" (ByVal hPrinter As Long, pBuf As Any, ByVal cdBuf As Long, pcWritten As Long) As Long
Private Declare Function OpenPrinter Lib "winspool.drv" Alias "OpenPrinterA" (ByVal pPrinterName As String, phPrinter As Long, pDefault As PRINTER_DEFAULTS) As Long
Private Declare Function ClosePrinter Lib "winspool.drv" (ByVal hPrinter As Long) As Long
Private Declare Function StartDocPrinter Lib "winspool.drv" Alias "StartDocPrinterA" (ByVal hPrinter As Long, ByVal Level As Long, pDocInfo As DOCINFO) As Long
Private Declare Function EndDocPrinter Lib "winspool.drv" (ByVal hPrinter As Long) As Long

Private Sub WritePrn(ByVal Dati As String)
    Dim ph As Long
    Dim ret As Long
    Dim lpW As Long
    Dim pdef As PRINTER_DEFAULTS
    Dim DocInf As DOCINFO
    'Dati = "questa è una prova lunga che dura un eternità"
    ret = OpenPrinter(Printer.DeviceName, ph, pdef)
    With DocInf
        .lpszDataType = vbNullString
        .lpszOutput = vbNullString
        .lpszDocName = App.Title
    End With
    ret = StartDocPrinter(ph, 1, DocInf)
    ret = WritePrinter(ph, ByVal Dati, Len(Dati), lpW)
    ret = EndDocPrinter(ph)
    ret = ClosePrinter(ph)
End Sub