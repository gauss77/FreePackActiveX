VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.UserControl ListUser 
   ClientHeight    =   1725
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3630
   ScaleHeight     =   1725
   ScaleWidth      =   3630
   Begin MSComctlLib.ImageList imlUser 
      Left            =   2760
      Top             =   900
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ListUser.ctx":0000
            Key             =   "User"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView lvwUser 
      DragIcon        =   "ListUser.ctx":059A
      Height          =   1155
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3075
      _ExtentX        =   5424
      _ExtentY        =   2037
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      SmallIcons      =   "imlUser"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
End
Attribute VB_Name = "ListUser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Const STR_KEY$ = "USER"
Private Const STR_UPDATE$ = "UPDATE "
Private Const STR_SET$ = " SET "
Private Const STR_WHERE$ = " WHERE "
Private Const STR_SELECT$ = "SELECT "
Private Const STR_DELETE$ = "DELETE FROM "
Private Const STR_INSERT$ = "INSERT INTO "
Private Const STR_FROM$ = " FROM "
Private Const STR_ORDERBY$ = " ORDER BY "
Private Const STR_IN$ = "IN("
Private Const STR_EQ$ = " = "

'Valori predefiniti proprietà:
Const m_def_DataList = "Cognome, Nome, Telefono"
Const m_def_IDFolder = 0
Const m_def_IDFieldName = "ID"
Const m_def_FolderFieldName = "IDCartella"
Const m_def_TableName = "Nominativi"
Const m_def_UseTransaction = False

'Variabili proprietà:
Dim m_IDFolder As Long
Dim m_Connection As Object
Dim m_UseTransaction As Boolean

'Dichiarazioni di eventi:
Event Error(Number As Long, Description As String)
Event ItemClick(ByVal Item As ListItem) 'MappingInfo=lvwUser,lvwUser,-1,ItemClick
Attribute ItemClick.VB_Description = "Viene generato quando si fa clic su un oggetto ListItem o quando lo si seleziona."

Private Sub RecuperaElencoNominativi()
    Dim i As Long
    Dim r As Integer
    Dim sql As String
    Dim rs As Object
    Dim rec As Variant
    Dim itm As ListItem
    
    #If ERRORDEBUG = 0 Then
        On Local Error GoTo Err_RecuperaElencoNominativi
    #End If
    ' recupera la lista dei nominativi dal database
    sql = STR_SELECT & m_def_IDFieldName & "," & m_def_DataList
    sql = sql & STR_FROM & m_def_TableName
    sql = sql & STR_WHERE & m_def_FolderFieldName & STR_EQ & m_IDFolder
    Set rs = m_Connection.Execute(sql)
    ' crea le colonne per la listview
    With lvwUser
        .Visible = False
        .ListItems.Clear
        .ColumnHeaders.Clear
        For r = 0 To rs.Fields.Count - 1
            If rs.Fields(r).Name <> m_def_IDFieldName Then .ColumnHeaders.Add , , rs.Fields(r).Name
        Next
        If Not rs.EOF Then              ' se ci sono nominativi
            rec = rs.GetRows            ' recupera i nominativi in una botta sola
            For i = 0 To UBound(rec, 2) ' cicla tutti i nominativi recuperati
                ' riempie la listview con ogni singolo nominativo
                Set itm = .ListItems.Add(, STR_KEY & rec(0, i), rec(1, i), , 1)
                ' ridimensiona la colonna principale se più piccola
                If TextWidth(rec(1, i)) * 2 > .ColumnHeaders(i + 1).Width Then .ColumnHeaders(i + 1).Width = TextWidth(rec(1, i)) * 2
                For r = 2 To UBound(rec)            ' cicla tutti gli altri campi
                    itm.SubItems(r - 1) = rec(r, i) ' li aggiunge alla listview
                    ' ridimensiona la colonna se più piccola
                    If TextWidth(rec(r, i)) * 2 > .ColumnHeaders(i + 1).Width Then .ColumnHeaders(i + 1).Width = TextWidth(rec(r, i)) * 2
                Next
            Next
            ' ordina tutti i nominativi
            .SortKey = 0
            .Sorted = True
            ' seleziona il primo nominativo della lista
            .ListItems(1).Selected = True
            Call lvwUser_ItemClick(.SelectedItem)
            Set itm = Nothing
        End If
        rs.Close
        Set rs = Nothing
        .Visible = True
    End With
    Exit Sub
    
Err_RecuperaElencoNominativi:
    Set itm = Nothing
    Set rs = Nothing
    lvwUser.Visible = True
    RaiseEvent Error(Err.Number, Err.Description)
End Sub

Private Sub SpostaNominativo(ByVal ID As Long, ByVal destID As Long)
    Dim sql As String

    #If ERRORDEBUG = 0 Then
        On Local Error GoTo Err_SpostaNominativo
    #End If
    ' sposta il nominativo nel database
    sql = STR_UPDATE & m_def_TableName & STR_SET
    sql = sql & m_def_FolderFieldName & STR_EQ & destID
    sql = sql & STR_WHERE & m_def_IDFieldName & STR_EQ & ID
    If m_UseTransaction Then
        m_Connection.BeginTrans
        m_Connection.Execute sql
        m_Connection.CommitTrans
    Else
        m_Connection.Execute sql
    End If
    ' rimuove il nominativo dalla lista attuale
    lvwUser.ListItems.Remove lvwUser.ListItems(STR_KEY & ID).Index
    Exit Sub
    
Err_SpostaNominativo:
    If m_UseTransaction Then m_Connection.RollbackTrans
    RaiseEvent Error(Err.Number, Err.Description)
End Sub

Public Sub Refresh()
Attribute Refresh.VB_Description = "Ridisegna completamente un form o un controllo."
    Call RecuperaElencoNominativi
End Sub

Private Sub lvwUser_ColumnClick(ByVal ColumnHeader As ColumnHeader)
    ' ordina la visualizzazione dei nominativi
    lvwUser.SortKey = ColumnHeader.Index - 1
    lvwUser.Sorted = True
End Sub

Public Property Get HideColumnHeaders() As Boolean
Attribute HideColumnHeaders.VB_Description = "Restituisce o imposta un valore che determina se le intestazioni di colonna di un controllo ListView sono nascoste in visualizzazione report."
    HideColumnHeaders = lvwUser.HideColumnHeaders
End Property

Public Property Let HideColumnHeaders(ByVal New_HideColumnHeaders As Boolean)
    lvwUser.HideColumnHeaders() = New_HideColumnHeaders
    PropertyChanged "HideColumnHeaders"
End Property

Private Sub lvwUser_DblClick()
    If Not lvwUser.SelectedItem Is Nothing Then Call ViewUser(ID(lvwUser.SelectedItem))
End Sub

Private Sub lvwUser_ItemClick(ByVal Item As ListItem)
    RaiseEvent ItemClick(Item)
End Sub

Public Property Get Connection() As Object
    Set Connection = m_Connection
End Property

Public Property Set Connection(ByVal New_Connection As Object)
    Set m_Connection = New_Connection
    PropertyChanged "Connection"
End Property

Public Sub DeleteUser()

End Sub

Public Property Get SelectedItem() As ListItem
    Set SelectedItem = lvwUser.SelectedItem
End Property

Public Property Let SelectedItem(ByVal New_SelectedItem As ListItem)
    Set lvwUser.SelectedItem = New_SelectedItem
    PropertyChanged "SelectedItem"
End Property

Public Property Get UseTransaction() As Boolean
    UseTransaction = m_UseTransaction
End Property

Public Property Let UseTransaction(ByVal New_UseTransaction As Boolean)
    m_UseTransaction = New_UseTransaction
    PropertyChanged "UseTransaction"
End Property

Private Sub lvwUser_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Not lvwUser.SelectedItem Is Nothing Then
        ' inizia il drag/drop solo se premuto il tasto sinistro del mouse
        ' e la cartella non è la cartella principale o il cestino
        If Button = vbLeftButton Then lvwUser.Drag vbBeginDrag
    End If
End Sub

'Inizializza le proprietà di UserControl.
Private Sub UserControl_InitProperties()
    m_UseTransaction = m_def_UseTransaction
    m_IDFolder = m_def_IDFolder
End Sub

'Carica i valori delle proprietà dalla posizione di memorizzazione.
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    lvwUser.HideColumnHeaders = PropBag.ReadProperty("HideColumnHeaders", False)
    Set m_Connection = PropBag.ReadProperty("Connection", Nothing)
    m_UseTransaction = PropBag.ReadProperty("UseTransaction", m_def_UseTransaction)
    m_IDFolder = PropBag.ReadProperty("IDFolder", m_def_IDFolder)
End Sub

Private Sub UserControl_Resize()
    On Local Error Resume Next
    lvwUser.Move 0, 0, Width, Height
End Sub

Private Sub UserControl_Show()
    Call UserControl_Resize
End Sub

Private Sub UserControl_Terminate()
    Set m_Connection = Nothing
End Sub

'Scrive i valori delle proprietà nella posizione di memorizzazione.
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("HideColumnHeaders", lvwUser.HideColumnHeaders, False)
    Call PropBag.WriteProperty("UseTransaction", m_UseTransaction, m_def_UseTransaction)
    Call PropBag.WriteProperty("IDFolder", m_IDFolder, m_def_IDFolder)
End Sub

Public Property Get Count() As Long
    Count = lvwUser.ListItems.Count
End Property

Public Property Let Count(ByVal New_Count As Long)
    If Ambient.UserMode = False Then Err.Raise 387
    If Ambient.UserMode Then Err.Raise 382
    PropertyChanged "Count"
End Property

Public Property Get IDFolder() As Long
    IDFolder = m_IDFolder
End Property

Public Property Let IDFolder(ByVal New_IDFolder As Long)
    m_IDFolder = New_IDFolder
    PropertyChanged "IDFolder"
End Property

Public Property Get DataList() As String
    DataList = m_def_DataList
End Property

Public Function ID(Item As ListItem) As Long
    Dim i As Integer
    Dim p As Integer
    Dim sKey As String
    
    ID = -1
    If Not Item Is Nothing Then
        ' cerca il primo numero
        p = 0
        sKey = Item.Key
        For i = 1 To Len(sKey)
            Select Case Asc(Mid$(sKey, i, 1))
                Case 48 To 57
                    p = i
                    Exit For
            End Select
        Next
        If p > 0 Then ID = CLng(Mid$(sKey, p))
    End If
End Function

Public Sub MoveUser(ByVal UserID As Long, ByVal destFolderID As Long)
    Call SpostaNominativo(UserID, destFolderID)
End Sub

Public Function Image(ByVal Item As ListItem) As ListImage
    Set Image = imlUser.ListImages(Item.SmallIcon)
End Function

Public Sub ViewUser(ByVal ID As Long)
    Dim frm As frmDati
    
    Set frm = New frmDati
    With frm
        .ID = ID
        .TableName = m_def_TableName
        .IDFieldName = m_def_IDFieldName
        .Create = False
        .ReadOnly = True
        Set .Connection = m_Connection
        .Show vbModal
    End With
    Set frm = Nothing
End Sub

Public Function ModifyUser(ByVal ID As Long)
    Dim frm As frmDati
    
    Set frm = New frmDati
    With frm
        .ID = ID
        .TableName = m_def_TableName
        .IDFieldName = m_def_IDFieldName
        .Create = False
        .ReadOnly = False
        Set .Connection = m_Connection
        .Show vbModal
        If .Changed Then Call RecuperaElencoNominativi
    End With
    Set frm = Nothing
End Function

Public Function AddNewUser(ByVal IDFolder As Long)
    Dim frm As frmDati
    
    Set frm = New frmDati
    With frm
        .ID = 0
        .TableName = m_def_TableName
        .IDFieldName = m_def_IDFieldName
        .IDFolder = IDFolder
        .Create = True
        .ReadOnly = False
        Set .Connection = m_Connection
        .Show vbModal
        If .Changed Then Call RecuperaElencoNominativi
    End With
    Set frm = Nothing
End Function
