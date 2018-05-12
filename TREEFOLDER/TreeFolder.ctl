VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.UserControl TreeFolder 
   Alignable       =   -1  'True
   BackStyle       =   0  'Transparent
   ClientHeight    =   3105
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2610
   ScaleHeight     =   3105
   ScaleWidth      =   2610
   Begin MSComctlLib.ImageList imlCat 
      Left            =   1560
      Top             =   1680
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "TreeFolder.ctx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "TreeFolder.ctx":2382
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "TreeFolder.ctx":4704
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.TreeView tvwCat 
      DragIcon        =   "TreeFolder.ctx":485E
      Height          =   2775
      Left            =   120
      TabIndex        =   0
      Top             =   60
      Width           =   1755
      _ExtentX        =   3096
      _ExtentY        =   4895
      _Version        =   393217
      HideSelection   =   0   'False
      Sorted          =   -1  'True
      Style           =   7
      ImageList       =   "imlCat"
      Appearance      =   1
   End
End
Attribute VB_Name = "TreeFolder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Const STR_KEY$ = "FOLDER"
Private Const STR_KEYCARTELLE$ = "FOLDER0"
Private Const STR_KEYCESTINO$ = "BASKET"
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

Private lstNodes As Nodes
Private sElencoNodi As String

'Valori predefiniti proprietà:
Const m_def_ChildrenBasket = 0
Const m_def_Prompt = ""
Const m_def_Title = ""
Const m_def_DefaultNewFolderName = "New Folder"
Const m_def_UseTransaction = False
Const m_def_TableName = "Folders"
Const m_def_MainFolderName = "Folders"
Const m_def_BasketName = "Basket"
Const m_def_IDFieldName = "ID"
Const m_def_IDParentFieldName = "IDParent"
Const m_def_DescriptionFieldName = "Description"

'Variabili proprietà:
Dim m_ChildrenBasket As Long
Dim m_Prompt As String
Dim m_Title As String
Dim m_DefaultNewFolderName As String
Dim m_UseTransaction As Boolean
Dim m_Connection As Object
Dim m_TableName As String
Dim m_MainFolderName As String
Dim m_BasketName As String
Dim m_IDFieldName As String
Dim m_IDParentFieldName As String
Dim m_DescriptionFieldName As String

'Dichiarazioni di eventi:
Event BeforeDeleteFolder(ByVal Node As Node, Cancel As Boolean)
Event BeforeEmptyBasket(ByVal Nodes As Nodes, Cancel As Boolean)
Event Error(Number As Long, Description As String)
Event NodeClick(ByVal Node As Node) 'MappingInfo=tvwCat,tvwCat,-1,NodeClick

Private Function Str2SQL(sStr As String) As String
    ' formatta una stringa per renderla compatibile con SQL (doppi apici)
    Str2SQL = Replace$(sStr, "'", "''")
End Function

Private Function SpostaCartellaDB(ByVal ID As Long, ByVal Dest As Long) As Boolean
    Dim sql As String
    
    #If ERRORDEBUG = 0 Then
        On Local Error GoTo Err_SpostaCartellaDB
    #End If
    ' sposta la cartella nella destinazione (DB)
    SpostaCartellaDB = False
    If ID = Dest Then Exit Function
    sql = STR_UPDATE & TableName & STR_SET & IDParentFieldName & STR_EQ & Dest
    sql = sql & STR_WHERE & IDFieldName & STR_EQ & ID
    If m_UseTransaction Then
        m_Connection.BeginTrans
        m_Connection.Execute sql
        m_Connection.CommitTrans
    Else
        m_Connection.Execute sql
    End If
    SpostaCartellaDB = True
    Exit Function
    
Err_SpostaCartellaDB:
    If m_UseTransaction Then m_Connection.RollbackTrans
    RaiseEvent Error(Err.Number, Err.Description)
End Function

Private Sub CreaNuovaCartella()
    Dim rel As Long
    Dim ID As Long
    Dim itm As Node
    Dim sql As String
    Dim rs As Object
    
    #If ERRORDEBUG = 0 Then
        On Local Error GoTo Err_CreaNuovaCartella
    #End If
    ' recupera l'id della cartella corrente (padre)
    If tvwCat.SelectedItem.Parent Is Nothing Then rel = 0 Else rel = CLng(Mid$(tvwCat.SelectedItem.Parent.Key, Len(STR_KEY) + 1))
    ' inserisce la nuova cartella nel database e recupera l'id della stessa
    sql = STR_SELECT & " * " & STR_FROM & m_TableName & STR_WHERE & m_IDFieldName & STR_EQ & "0"
    Set rs = m_Connection.Execute(sql)
    With rs
        ' apre keyset/pessimistic
        .Open TableName, Connection, 1, 2
        .AddNew
        .Fields(IDParentFieldName) = rel
        .Fields(DescriptionFieldName) = m_DefaultNewFolderName
        .Update
        ID = .Fields(IDFieldName)
        .Close
    End With
    Set rs = Nothing
    ' aggiunge la nuova cartella all'albero delle cartelle e permette la modifica
    With tvwCat
        Set itm = .Nodes.Add(STR_KEY & rel, tvwChild, STR_KEY & ID, m_DefaultNewFolderName, 1, 2)
        .Nodes(STR_KEY & rel).Sorted = True
        itm.EnsureVisible
        itm.Selected = True
        .StartLabelEdit
    End With
    Set itm = Nothing
    Exit Sub
    
Err_CreaNuovaCartella:
    Set itm = Nothing
    Set rs = Nothing
    RaiseEvent Error(Err.Number, Err.Description)
End Sub

Private Sub RecuperaElencoCartelle()
    Dim sql As String
    Dim rs As Object
    Dim itm As Node
    
    ' prepara l'albero delle cartelle
    With tvwCat
        .Visible = False
        .Nodes.Clear
        Set itm = .Nodes.Add(, , STR_KEYCARTELLE, MainFolderName, 2)
        itm.Selected = True
        itm.Expanded = True
        Set tvwCat.SelectedItem = itm
        .Nodes.Add STR_KEYCARTELLE, tvwChild, STR_KEYCESTINO, BasketName, 3
    End With
    #If ERRORDEBUG = 0 Then
        On Local Error GoTo Err_RecuperaElencoCartelle
    #End If
    ' recupera dal database le cartelle
    sql = STR_SELECT & " * " & STR_FROM & TableName & STR_ORDERBY & IDFieldName & "," & IDParentFieldName & "," & DescriptionFieldName
    Set rs = m_Connection.Execute(sql)
    Do While Not rs.EOF
        ' riempie l'albero delle cartelle
        tvwCat.Nodes.Add STR_KEY & rs(IDParentFieldName), tvwChild, STR_KEY & rs(IDFieldName), rs(DescriptionFieldName), 1, 2
        rs.MoveNext
    Loop
    rs.Close
    Set rs = Nothing
    ' ordina le cartelle
    For Each itm In tvwCat.Nodes
        If itm.Children > 0 Then itm.Sorted = True
    Next
    tvwCat.Visible = True
    Exit Sub
    
Err_RecuperaElencoCartelle:
    tvwCat.Visible = True
    Set rs = Nothing
    RaiseEvent Error(Err.Number, Err.Description)
End Sub

Private Sub SpostaCartellaAct(Source As Node, Dest As Node)
    Dim ID As Long
    Dim IDdest As Long
    
    ' sposta la cartella nella destinazione scelta
    With tvwCat
        If Source Is Nothing Or Dest Is Nothing Then Exit Sub
        ID = CLng(Mid$(Source.Key, Len(STR_KEY) + 1))
        IDdest = CLng(Mid$(Dest.Key, Len(STR_KEY) + 1))
        If IDdest <> ID And SpostaCartellaDB(ID, IDdest) Then
            Set Source.Parent = Dest
            Dest.Sorted = True
        End If
    End With
End Sub

Private Sub SpostaCartella()
    Dim bCancel As Boolean
    Dim frm As frmSposta
    
    ' visualizza il form per la scelta della cartella di destinazione
    Set frm = New frmSposta
    With frm
        .Prompt = m_Prompt
        .Title = m_Title
        Set .Nodes = tvwCat.Nodes
        .Show vbModal
        If Not .Node Is Nothing Then
            bCancel = False
            If Not bCancel Then Call SpostaCartellaAct(tvwCat.SelectedItem, .Node)
        End If
    End With
    Set frm = Nothing
End Sub

Private Sub EliminaCartella()
    Dim ret As Long
    Dim bCancel As Boolean
    
    ' sposta la cartella selezionata nel cestino
    If tvwCat.SelectedItem Is Nothing Then Exit Sub
    bCancel = False
    RaiseEvent BeforeDeleteFolder(tvwCat.SelectedItem, bCancel)
    If Not bCancel Then Call SpostaCartellaAct(tvwCat.SelectedItem, tvwCat.Nodes(STR_KEYCESTINO))
End Sub

Private Sub ElencoNodi(objNode As Node, Optional ByVal bStart As Boolean = False)
    Static iIDLevel As Integer
    ' procedura ricorsiva che elenca tutte le key dei nodi presenti sotto al nodo di partenza
    ' nella variabile sElencoNodi saranno elencate tutte le key dei nodi separati da virgola
    If bStart Then iIDLevel = 0
    If StrComp(Mid$(objNode.Key, 1, Len(STR_KEY)), STR_KEY) = 0 Then sElencoNodi = sElencoNodi & objNode.Key & ","
    lstNodes.Add , , objNode.Key, objNode.Text
    If objNode.Children > 0 Then
        iIDLevel = iIDLevel + 1
        Call ElencoNodi(objNode.Child)
    End If
    Set objNode = objNode.Next
    If TypeName(objNode) <> "Nothing" Then Call ElencoNodi(objNode) Else iIDLevel = iIDLevel - 1
End Sub

Private Sub SvuotaCestino()
    Dim bCancel As Boolean
    Dim sql As String
    
    #If ERRORDEBUG = 0 Then
        On Local Error GoTo Err_SvuotaCestino
    #End If
    lstNodes.Clear
    sElencoNodi = vbNullString
    Call ElencoNodi(tvwCat.Nodes(STR_KEYCESTINO).Child, True)
    sElencoNodi = Replace$(Mid$(sElencoNodi, 1, Len(sElencoNodi) - 1), STR_KEY, vbNullString)
    bCancel = False
    RaiseEvent BeforeEmptyBasket(lstNodes, bCancel)
    Set lstNodes = Nothing
    If bCancel Then Exit Sub
    ' elimina le cartelle presenti nel database
    sql = STR_DELETE & TableName
    sql = sql & STR_WHERE & IDFieldName & STR_IN & sElencoNodi & ")"
    If m_UseTransaction Then
        m_Connection.BeginTrans
        m_Connection.Execute sql
        m_Connection.CommitTrans
    Else
        m_Connection.Execute sql
    End If
    ' elimina le cartelle presenti nel cestino
    With tvwCat
        .Nodes.Remove tvwCat.Nodes(STR_KEYCESTINO).Index
        .Nodes.Add STR_KEYCARTELLE, tvwChild, STR_KEYCESTINO, BasketName, 3
        .Nodes(STR_KEYCARTELLE).Sorted = True
    End With
    Exit Sub
    
Err_SvuotaCestino:
    If m_UseTransaction Then m_Connection.RollbackTrans
    RaiseEvent Error(Err.Number, Err.Description)
End Sub

Private Sub tvwCat_AfterLabelEdit(Cancel As Integer, NewString As String)
    Dim ID As Long
    Dim sql As String
    Dim itm As Node
    
    #If ERRORDEBUG = 0 Then
        On Local Error GoTo Err_AfterLabelEdit
    #End If
    ' aggiorna la descrizione della cartella
    Set itm = tvwCat.SelectedItem
    ID = CLng(Mid$(itm.Key, Len(STR_KEY) + 1))
    sql = STR_UPDATE & TableName & STR_SET
    sql = sql & DescriptionFieldName & STR_EQ & "'" & Str2SQL(NewString) & "'"
    sql = sql & STR_WHERE & IDFieldName & STR_EQ & ID
    If m_UseTransaction Then
        m_Connection.BeginTrans
        m_Connection.Execute sql
        m_Connection.CommitTrans
    Else
        m_Connection.Execute sql
    End If
    itm.Text = NewString
    itm.Parent.Sorted = True
    Set itm = Nothing
    Exit Sub
    
Err_AfterLabelEdit:
    Cancel = True
    Set itm = Nothing
    If m_UseTransaction Then m_Connection.RollbackTrans
    RaiseEvent Error(Err.Number, Err.Description)
End Sub

Private Sub tvwCat_Collapse(ByVal Node As MSComctlLib.Node)
    ' non chiude la cartella principale
    If Node.Key = STR_KEYCARTELLE Then Node.Expanded = True
End Sub

Private Sub tvwCat_DragDrop(Source As Control, x As Single, y As Single)
    ' se un nodo di questo treeview
    If TypeOf Source Is Node Then
        ' sposta la cartella nella destinazione scelta con il drag/drop
        Call SpostaCartellaAct(tvwCat.SelectedItem, tvwCat.DropHighlight)
    End If
    Set tvwCat.DropHighlight = Nothing
End Sub

Private Sub tvwCat_DragOver(Source As Control, x As Single, y As Single, State As Integer)
    ' evidenzia la destinazione del drag/drop
    Set tvwCat.DropHighlight = tvwCat.HitTest(x, y)
End Sub

Private Sub tvwCat_KeyDown(KeyCode As Integer, Shift As Integer)
    ' sposta nel cestino la cartella selezionata se viene premuto CANC
    If KeyCode = vbKeyDelete Then Call EliminaCartella
End Sub

Private Sub tvwCat_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    ' seleziona il nodo appena viene premuto il testo del mouse
    Set tvwCat.SelectedItem = tvwCat.HitTest(x, y)
End Sub

Private Sub tvwCat_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Not tvwCat.SelectedItem Is Nothing Then
        ' inizia il drag/drop solo se premuto il tasto sinistro del mouse
        ' e la cartella non è la cartella principale o il cestino
        If Button = vbLeftButton And ID(tvwCat.SelectedItem) > 0 Then tvwCat.Drag vbBeginDrag
    End If
End Sub

Private Sub tvwCat_NodeClick(ByVal Node As MSComctlLib.Node)
    RaiseEvent NodeClick(Node)
End Sub

Private Sub UserControl_Resize()
    On Local Error Resume Next
    tvwCat.Move 0, 0, Width, Height
End Sub

Private Sub UserControl_Show()
    Call UserControl_Resize
End Sub

Public Sub Refresh()
Attribute Refresh.VB_Description = "Ridisegna completamente un form o un controllo."
    Call RecuperaElencoCartelle
End Sub

Public Property Get Connection() As Object
    Set Connection = m_Connection
End Property

Public Property Set Connection(ByVal New_Connection As Object)
    Set m_Connection = New_Connection
    PropertyChanged "Connection"
End Property

Public Property Get TableName() As String
    TableName = m_TableName
End Property

Public Property Let TableName(ByVal New_TableName As String)
    m_TableName = New_TableName
    PropertyChanged "TableName"
End Property

Public Property Get MainFolderName() As String
    MainFolderName = m_MainFolderName
End Property

Public Property Let MainFolderName(ByVal New_MainFolderName As String)
    On Local Error Resume Next
    m_MainFolderName = New_MainFolderName
    tvwCat.Nodes(STR_KEYCARTELLE).Text = m_MainFolderName
    PropertyChanged "MainFolderName"
End Property

Public Property Get BasketName() As String
    BasketName = m_BasketName
End Property

Public Property Let BasketName(ByVal New_BasketName As String)
    On Local Error Resume Next
    m_BasketName = New_BasketName
    tvwCat.Nodes(STR_KEYCESTINO).Text = m_BasketName
    PropertyChanged "BasketName"
End Property

Public Property Get IDFieldName() As String
    IDFieldName = m_IDFieldName
End Property

Public Property Let IDFieldName(ByVal New_IDFieldName As String)
    m_IDFieldName = New_IDFieldName
    PropertyChanged "IDFieldName"
End Property

Public Property Get IDParentFieldName() As String
    IDParentFieldName = m_IDParentFieldName
End Property

Public Property Let IDParentFieldName(ByVal New_IDParentFieldName As String)
    m_IDParentFieldName = New_IDParentFieldName
    PropertyChanged "IDParentFieldName"
End Property

Public Property Get DescriptionFieldName() As String
    DescriptionFieldName = m_DescriptionFieldName
End Property

Public Property Let DescriptionFieldName(ByVal New_DescriptionFieldName As String)
    m_DescriptionFieldName = New_DescriptionFieldName
    PropertyChanged "DescriptionFieldName"
End Property

'Inizializza le proprietà di UserControl.
Private Sub UserControl_InitProperties()
    m_TableName = m_def_TableName
    m_MainFolderName = m_def_MainFolderName
    m_BasketName = m_def_BasketName
    m_IDFieldName = m_def_IDFieldName
    m_IDParentFieldName = m_def_IDParentFieldName
    m_DescriptionFieldName = m_def_DescriptionFieldName
    m_UseTransaction = m_def_UseTransaction
    m_DefaultNewFolderName = m_def_DefaultNewFolderName
    m_Prompt = m_def_Prompt
    m_Title = m_def_Title
    m_ChildrenBasket = m_def_ChildrenBasket
End Sub

'Carica i valori delle proprietà dalla posizione di memorizzazione.
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    Set m_Connection = PropBag.ReadProperty("Connection", Nothing)
    m_TableName = PropBag.ReadProperty("TableName", m_def_TableName)
    m_MainFolderName = PropBag.ReadProperty("MainFolderName", m_def_MainFolderName)
    m_BasketName = PropBag.ReadProperty("BasketName", m_def_BasketName)
    m_IDFieldName = PropBag.ReadProperty("IDFieldName", m_def_IDFieldName)
    m_IDParentFieldName = PropBag.ReadProperty("IDParentFieldName", m_def_IDParentFieldName)
    m_DescriptionFieldName = PropBag.ReadProperty("DescriptionFieldName", m_def_DescriptionFieldName)
    m_UseTransaction = PropBag.ReadProperty("UseTransaction", m_def_UseTransaction)
    m_DefaultNewFolderName = PropBag.ReadProperty("DefaultNewFolderName", m_def_DefaultNewFolderName)
    m_Prompt = PropBag.ReadProperty("Prompt", m_def_Prompt)
    m_Title = PropBag.ReadProperty("Title", m_def_Title)
    m_ChildrenBasket = PropBag.ReadProperty("ChildrenBasket", m_def_ChildrenBasket)
End Sub

Private Sub UserControl_Terminate()
    Set m_Connection = Nothing
End Sub

'Scrive i valori delle proprietà nella posizione di memorizzazione.
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("Connection", m_Connection, Nothing)
    Call PropBag.WriteProperty("TableName", m_TableName, m_def_TableName)
    Call PropBag.WriteProperty("MainFolderName", m_MainFolderName, m_def_MainFolderName)
    Call PropBag.WriteProperty("BasketName", m_BasketName, m_def_BasketName)
    Call PropBag.WriteProperty("IDFieldName", m_IDFieldName, m_def_IDFieldName)
    Call PropBag.WriteProperty("IDParentFieldName", m_IDParentFieldName, m_def_IDParentFieldName)
    Call PropBag.WriteProperty("DescriptionFieldName", m_DescriptionFieldName, m_def_DescriptionFieldName)
    Call PropBag.WriteProperty("UseTransaction", m_UseTransaction, m_def_UseTransaction)
    Call PropBag.WriteProperty("DefaultNewFolderName", m_DefaultNewFolderName, m_def_DefaultNewFolderName)
    Call PropBag.WriteProperty("Prompt", m_Prompt, m_def_Prompt)
    Call PropBag.WriteProperty("Title", m_Title, m_def_Title)
    Call PropBag.WriteProperty("ChildrenBasket", m_ChildrenBasket, m_def_ChildrenBasket)
End Sub

Public Property Get UseTransaction() As Boolean
    UseTransaction = m_UseTransaction
End Property

Public Property Let UseTransaction(ByVal New_UseTransaction As Boolean)
    m_UseTransaction = New_UseTransaction
    PropertyChanged "UseTransaction"
End Property

Public Property Get DefaultNewFolderName() As String
    DefaultNewFolderName = m_DefaultNewFolderName
End Property

Public Property Let DefaultNewFolderName(ByVal New_DefaultNewFolderName As String)
    m_DefaultNewFolderName = New_DefaultNewFolderName
    PropertyChanged "DefaultNewFolderName"
End Property

Public Sub AddFolder()
    Call CreaNuovaCartella
End Sub

Public Sub MoveFolder(Optional Prompt As String = vbNullString, Optional Title As String = vbNullString)
    If Len(Prompt) > 0 Then
        m_Prompt = Prompt
        PropertyChanged "Prompt"
    End If
    If Len(Title) > 0 Then
        m_Title = Title
        PropertyChanged "Title"
    End If
    Call SpostaCartella
End Sub

Public Sub DeleteFolder()
    Call EliminaCartella
End Sub

Public Sub EmptyBasket()
    Call SvuotaCestino
End Sub

Public Sub RenameFolder()
    tvwCat.StartLabelEdit
End Sub

Public Sub AddItemOnBasket(ByVal sKey As String, ByVal sText As String, ByVal lstImage As ListImage)
    ' aggiunge l'item al cestino
    tvwCat.Nodes.Add STR_KEYCESTINO, tvwChild, sKey, sText, imlCat.ListImages.Add(, lstImage.Key, lstImage.Picture).Index
End Sub

Public Function ShowFolders(Optional Prompt As String = vbNullString, Optional Title As String = vbNullString) As Node
    Dim frm As frmSposta
    
    ' visualizza il form per la scelta di una cartella
    Set frm = New frmSposta
    With frm
        If Len(Prompt) > 0 Then .Prompt = Prompt
        If Len(Title) > 0 Then .Title = Title
        Set .Nodes = tvwCat.Nodes
        .Show vbModal
        Set ShowFolders = .Node
    End With
    Set frm = Nothing
End Function

Public Property Get SelectedItem() As Node
    Set SelectedItem = tvwCat.SelectedItem
End Property

Public Property Set SelectedItem(ByVal New_SelectedItem As Node)
    Set tvwCat.SelectedItem = New_SelectedItem
    PropertyChanged "SelectedItem"
End Property

Public Property Get ChildrenBasket() As Long
    On Local Error Resume Next
    ChildrenBasket = tvwCat.Nodes(STR_KEYCESTINO).Children
End Property

Public Function ID(Node As Node) As Long
    Dim i As Integer
    Dim p As Integer
    Dim sKey As String
    
    ID = -1
    sKey = Node.Key
    If sKey <> STR_KEYCESTINO Then
        ' cerca il primo numero
        p = 0
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

Public Property Get Basket() As Node
    Set Basket = tvwCat.Nodes(STR_KEYCESTINO)
End Property

