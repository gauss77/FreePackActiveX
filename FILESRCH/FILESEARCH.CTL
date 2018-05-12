VERSION 5.00
Begin VB.UserControl FileSearch 
   CanGetFocus     =   0   'False
   ClientHeight    =   480
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   480
   ClipControls    =   0   'False
   InvisibleAtRuntime=   -1  'True
   Picture         =   "FileSearch.ctx":0000
   PropertyPages   =   "FileSearch.ctx":08CA
   ScaleHeight     =   480
   ScaleWidth      =   480
   ToolboxBitmap   =   "FileSearch.ctx":08E3
   Windowless      =   -1  'True
End
Attribute VB_Name = "FileSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "PropPageWizardRun" ,"Yes"
Option Explicit

'Valori predefiniti proprietà:
Const m_def_Path = ""
Const m_def_Filename = ""

'Variabili proprietà:
Dim m_Path As String
Dim m_Filename As String

'Dichiarazioni degli oggetti
Dim lstFiles As Files

'Dichiarazioni di eventi:
Public Event Progress(PathName As String, NameFile As String)
Attribute Progress.VB_Description = "Event generate with found a file."

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

Public Sub About()
Attribute About.VB_Description = "Show about box."
Attribute About.VB_UserMemId = -552
    frmSplash.Show vbModal
End Sub

Public Property Get Path() As String
    Path = m_Path
End Property

Public Property Let Path(ByVal New_Path As String)
    m_Path = New_Path
    PropertyChanged "Path"
End Property

Public Property Get Filename() As String
    Filename = m_Filename
End Property

Public Property Let Filename(ByVal New_Filename As String)
    m_Filename = New_Filename
    PropertyChanged "Filename"
End Property

'Inizializza le proprietà di UserControl
Private Sub UserControl_InitProperties()
    m_Path = m_def_Path
    m_Filename = m_def_Filename
End Sub

'Carica i valori della proprietà dalla memoria
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    m_Path = PropBag.ReadProperty("Path", m_def_Path)
    m_Filename = PropBag.ReadProperty("Filename", m_def_Filename)
End Sub

'Scrive i valori della proprietà in memoria
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("Path", m_Path, m_def_Path)
    Call PropBag.WriteProperty("Filename", m_Filename, m_def_Filename)
End Sub

Public Function Search(Optional Filename As String = vbNullString, Optional Path As String = vbNullString) As Files
Attribute Search.VB_Description = "Start find files."
Attribute Search.VB_HelpID = 16
    If Len(Trim$(Filename)) > 0 Then m_Filename = Filename
    If Len(Trim$(Path)) > 0 Then m_Path = Path
    If Len(Trim$(m_Filename)) > 0 And Len(Trim$(m_Path)) > 0 Then
        Set lstFiles = New Files
        Call FilesSearch(m_Path, m_Filename)
        Set Search = lstFiles
        Set lstFiles = Nothing
    End If
End Function

Private Sub FilesSearch(DrivePath As String, Ext As String)

    Dim XDir() As String
    Dim TmpDir As String
    Dim FFound As String
    Dim DirCount As Integer
    Dim x As Integer
    '     'Initialises Variables
    DirCount = 0
    ReDim XDir(0) As String
    XDir(DirCount) = ""

    If Right(DrivePath, 1) <> "\" Then
        DrivePath = DrivePath & "\"
    End If

    '     'Enter here the code for showing the path being
    '     'search. Example: Form1.label2 = DrivePath
    '     'Search for all directories and store in the
    '     'XDir() variable

    'DoEvents
    TmpDir = Dir$(DrivePath, vbDirectory)

    Do While TmpDir <> ""

        If TmpDir <> "." And TmpDir <> ".." Then

            If (GetAttr(DrivePath & TmpDir) And vbDirectory) = vbDirectory Then
                XDir(DirCount) = DrivePath & TmpDir & "\"
                RaiseEvent Progress(XDir(DirCount), vbNullString)
                DirCount = DirCount + 1
                ReDim Preserve XDir(DirCount) As String
            End If

        End If

        TmpDir = Dir$
    Loop

    '     'Searches for the files given by extension Ext
    FFound = Dir(DrivePath & Ext)

    Do Until FFound = ""
        '     'Code in here for the actions of the files found.
        '     'Files found stored in the variable FFound.
        '     'Example: Form1.list1.AddItem DrivePath & FFound
        lstFiles.Add FFound, DrivePath
        RaiseEvent Progress(DrivePath, FFound)
        FFound = Dir$
    Loop

    '     'Recursive searches through all sub directories
    For x = 0 To (UBound(XDir) - 1)
        FilesSearch XDir(x), Ext
    Next x
End Sub

