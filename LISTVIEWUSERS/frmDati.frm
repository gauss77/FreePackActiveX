VERSION 5.00
Object = "{918AABEF-81E0-11D6-9068-0080C88A8003}#2.0#0"; "WMInterfaceXPFree.ocx"
Begin VB.Form frmDati 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Nominativo"
   ClientHeight    =   2685
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4770
   Icon            =   "frmDati.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2685
   ScaleWidth      =   4770
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin WMInterfaceXPFree.GroupXPFree grpDati 
      Height          =   2535
      Left            =   60
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   60
      Width           =   4635
      _ExtentX        =   8176
      _ExtentY        =   4471
      Caption         =   "Dati principali"
      BackColor       =   -2147483633
      Begin WMInterfaceXPFree.CommandXPFree cdmAdvanced 
         Height          =   315
         Left            =   3240
         TabIndex        =   8
         Top             =   1620
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   556
         Caption         =   "Avanzate"
      End
      Begin WMInterfaceXPFree.CommandXPFree cmdCancel 
         Cancel          =   -1  'True
         Height          =   315
         Left            =   3240
         TabIndex        =   7
         Top             =   1140
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   556
         Caption         =   "Annulla"
      End
      Begin WMInterfaceXPFree.CommandXPFree cmdOK 
         Default         =   -1  'True
         Height          =   315
         Left            =   3240
         TabIndex        =   6
         Top             =   660
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   556
         Caption         =   "Ok"
      End
      Begin VB.TextBox txtPhone 
         Height          =   315
         Left            =   300
         MaxLength       =   24
         TabIndex        =   5
         Top             =   1980
         Width           =   2355
      End
      Begin VB.TextBox txtSurname 
         Height          =   315
         Left            =   300
         MaxLength       =   50
         TabIndex        =   3
         Top             =   1320
         Width           =   2355
      End
      Begin VB.TextBox txtName 
         Height          =   315
         Left            =   300
         MaxLength       =   50
         TabIndex        =   1
         Top             =   660
         Width           =   2355
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "&Telefono :"
         Height          =   195
         Left            =   300
         TabIndex        =   4
         Top             =   1740
         Width           =   720
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "&Cognome :"
         Height          =   195
         Left            =   300
         TabIndex        =   2
         Top             =   1080
         Width           =   765
      End
      Begin VB.Label lblName 
         AutoSize        =   -1  'True
         Caption         =   "&Nome :"
         Height          =   195
         Left            =   300
         TabIndex        =   0
         Top             =   420
         Width           =   510
      End
   End
End
Attribute VB_Name = "frmDati"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public ID As Long
Public IDFolder As Long
Public Changed As Boolean
Public ReadOnly As Boolean
Public Create As Boolean
Public TableName As String
Public IDFieldName As String
Public Connection As Object

Private Function Str2SQL(sStr As String) As String
    ' formatta una stringa per renderla compatibile con SQL (doppi apici)
    Str2SQL = Replace$(Trim$(sStr), "'", "''")
End Function

Private Sub cmdCancel_Click()
    ' scarica il form senza effettuare altre operazioni
    MousePointer = vbHourglass
    Changed = False
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Dim sql As String
    
    #If ERRORDEBUG = 0 Then
        On Local Error GoTo Err_OK
    #End If
    MousePointer = vbHourglass
    ' se i dati devono essere creati
    If Create And Changed Then
        ' crea la query di inserimento dei dati
        sql = "INSERT INTO " & TableName
        sql = sql & "(nome,cognome,telefono,idcartella)"
        sql = sql & " VALUES("
        sql = sql & "'" & Str2SQL(txtName.Text) & "',"
        sql = sql & "'" & Str2SQL(txtSurname.Text) & "',"
        sql = sql & "'" & txtPhone.Text & "',"
        sql = sql & CStr(IDFolder) & ")"
        Connection.Execute sql
    ' se i dati potevano essere modificati e sono stati modificati
    ElseIf Not ReadOnly And Changed Then
        ' esegue l'aggiornamento dei dati sul database
        sql = "UPDATE " & TableName & " SET "
        sql = sql & "Nome='" & Str2SQL(txtName.Text) & "',"
        sql = sql & "Cognome='" & Str2SQL(txtSurname.Text) & "',"
        sql = sql & "Telefono='" & txtPhone.Text & "'"
        sql = sql & " WHERE ID=" & ID
        Connection.Execute sql
    End If
    GoTo Unload_Me
    Exit Sub
    
Err_OK:
    MsgBox Err.Description, vbExclamation
    MousePointer = vbDefault
    Exit Sub
    
Unload_Me:
    Unload Me
End Sub

Private Sub Form_Load()
    Dim sql As String
    Dim rs As Object
    
    #If ERRORDEBUG = 0 Then
        On Local Error GoTo Err_Load
    #End If
    ' esegue la query di selezione dei dati del nominativo
    If Not Create And ID > 0 Then
        sql = "SELECT * FROM " & TableName & " WHERE " & IDFieldName & "=" & ID
        Set rs = Connection.Execute(sql)
        If Not rs.EOF Then
            ' riempie gli oggetti dle form con i dati del nominativo
            txtName.Text = Trim$(rs("nome"))
            txtSurname.Text = Trim$(rs("cognome"))
            txtPhone.Text = Trim$(rs("telefono"))
        End If
        rs.Close
        Set rs = Nothing
    End If
    
    ' dis/abilita la modifica dei dati
    Changed = False
    txtName.Locked = ReadOnly
    txtSurname.Locked = ReadOnly
    txtPhone.Locked = ReadOnly
    cmdOK.Enabled = False
    Exit Sub
    
Err_Load:
    Set rs = Nothing
    MsgBox Err.Description, vbExclamation
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set Connection = Nothing
End Sub

Private Sub txtName_Change()
    Changed = True
    cmdOK.Enabled = True
End Sub

Private Sub txtPhone_Change()
    Changed = True
    cmdOK.Enabled = True
End Sub

Private Sub txtPhone_KeyPress(KeyAscii As Integer)
    ' se non è numerico annulla
    Select Case KeyAscii
        Case 8, 48 To 57
            ' NOP
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub txtSurname_Change()
    Changed = True
    cmdOK.Enabled = True
End Sub
