VERSION 5.00
Begin VB.UserControl PSPrint 
   ClientHeight    =   750
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1005
   InvisibleAtRuntime=   -1  'True
   Picture         =   "PSPrint.ctx":0000
   PropertyPages   =   "PSPrint.ctx":08CA
   ScaleHeight     =   750
   ScaleWidth      =   1005
   ToolboxBitmap   =   "PSPrint.ctx":08DC
   Windowless      =   -1  'True
End
Attribute VB_Name = "PSPrint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Valori predefiniti proprietà:
Const m_def_PrinterDeviceName = ""
Const m_def_FileName = ""
Const m_def_DocumentTitle = ""

'Variabili proprietà:
Dim m_PrinterDeviceName As String
Dim m_FileName As String
Dim m_DocumentTitle As String

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

Private Sub WritePrn(ByVal buffer As String)
    Dim ph As Long
    Dim ret As Long
    Dim lpW As Long
    Dim pdef As PRINTER_DEFAULTS
    Dim DocInf As DOCINFO
    ret = OpenPrinter(m_PrinterDeviceName, ph, pdef)
    With DocInf
        .lpszDataType = vbNullString
        .lpszOutput = vbNullString
        .lpszDocName = m_DocumentTitle
    End With
    ret = StartDocPrinter(ph, 1, DocInf)
    ret = WritePrinter(ph, ByVal buffer, Len(buffer), lpW)
    ret = EndDocPrinter(ph)
    ret = ClosePrinter(ph)
End Sub

Public Sub About()
Attribute About.VB_Description = "Show about box."
Attribute About.VB_UserMemId = -552
    frmSplash.Show vbModal
End Sub

Private Sub UserControl_Initialize()
    #If SHAREWARE = 1 Then
        frmSplash.Show vbModal
    #End If
End Sub

Private Sub UserControl_Resize()
    On Local Error Resume Next
    Width = 480
    Height = 480
End Sub

'AVVISO: NON RIMUOVERE O MODIFICARE LE SEGUENTI RIGHE DI COMMENTO
'MemberInfo=13,0,0,
Public Property Get PrinterDeviceName() As String
Attribute PrinterDeviceName.VB_Description = "Return or set printer device name."
Attribute PrinterDeviceName.VB_HelpID = 17
    PrinterDeviceName = m_PrinterDeviceName
End Property

Public Property Let PrinterDeviceName(ByVal New_PrinterDeviceName As String)
    m_PrinterDeviceName = New_PrinterDeviceName
    PropertyChanged "PrinterDeviceName"
End Property

'AVVISO: NON RIMUOVERE O MODIFICARE LE SEGUENTI RIGHE DI COMMENTO
'MemberInfo=13,0,0,
Public Property Get Filename() As String
Attribute Filename.VB_Description = "Return or set PostScript file name."
Attribute Filename.VB_HelpID = 10
    Filename = m_FileName
End Property

Public Property Let Filename(ByVal New_FileName As String)
    m_FileName = New_FileName
    PropertyChanged "FileName"
End Property

'AVVISO: NON RIMUOVERE O MODIFICARE LE SEGUENTI RIGHE DI COMMENTO
'MemberInfo=13,0,0,
Public Property Get DocumentTitle() As String
Attribute DocumentTitle.VB_Description = "Return or set document title."
Attribute DocumentTitle.VB_HelpID = 9
    DocumentTitle = m_DocumentTitle
End Property

Public Property Let DocumentTitle(ByVal New_DocumentTitle As String)
    m_DocumentTitle = New_DocumentTitle
    PropertyChanged "DocumentTitle"
End Property

'AVVISO: NON RIMUOVERE O MODIFICARE LE SEGUENTI RIGHE DI COMMENTO
'MemberInfo=8
Public Function StartPrint(Optional ByVal sFilename As String = vbNullString, Optional ByVal sPrinterDeviceName As String = vbNullString) As Long
Attribute StartPrint.VB_Description = "Start print PostScript."
Attribute StartPrint.VB_HelpID = 16
    On Local Error GoTo Err_Print
    
    Dim ff As Integer
    Dim lin As String
    Dim buffer As String
    
    If Len(sFilename) > 0 Then Filename = sFilename
    If Len(sPrinterDeviceName) > 0 Then PrinterDeviceName = sPrinterDeviceName
    
    buffer = vbNullString
    ff = FreeFile
    Open m_FileName For Input As #ff
    Do While Not EOF(ff)
        Line Input #ff, lin
        buffer = buffer & lin & vbCrLf
    Loop
    Close #ff
    
    Call WritePrn(buffer)
    StartPrint = 0
    Exit Function
    
Err_Print:
    StartPrint = Err.Number
End Function

'Inizializza le proprietà di UserControl.
Private Sub UserControl_InitProperties()
    m_PrinterDeviceName = m_def_PrinterDeviceName
    m_FileName = m_def_FileName
    m_DocumentTitle = m_def_DocumentTitle
End Sub

'Carica i valori delle proprietà dalla posizione di memorizzazione.
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    m_PrinterDeviceName = PropBag.ReadProperty("PrinterDeviceName", m_def_PrinterDeviceName)
    m_FileName = PropBag.ReadProperty("FileName", m_def_FileName)
    m_DocumentTitle = PropBag.ReadProperty("DocumentTitle", m_def_DocumentTitle)
End Sub

'Scrive i valori delle proprietà nella posizione di memorizzazione.
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("PrinterDeviceName", m_PrinterDeviceName, m_def_PrinterDeviceName)
    Call PropBag.WriteProperty("FileName", m_FileName, m_def_FileName)
    Call PropBag.WriteProperty("DocumentTitle", m_DocumentTitle, m_def_DocumentTitle)
End Sub

