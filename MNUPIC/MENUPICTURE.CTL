VERSION 5.00
Begin VB.UserControl MenuImage 
   CanGetFocus     =   0   'False
   ClientHeight    =   480
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   480
   ClipControls    =   0   'False
   InvisibleAtRuntime=   -1  'True
   Picture         =   "MenuPicture.ctx":0000
   PropertyPages   =   "MenuPicture.ctx":030A
   ScaleHeight     =   480
   ScaleWidth      =   480
   ToolboxBitmap   =   "MenuPicture.ctx":0335
End
Attribute VB_Name = "MenuImage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "PropPageWizardRun" ,"Yes"
Option Explicit

'Valori predefiniti proprietà:
Const m_def_MenuCaption = ""
Const m_def_ItemCaption = ""

'Variabili proprietà:
Dim m_Image As Picture
Dim m_ImageUnchecked As Picture
Dim m_ImageChecked As Picture
Dim m_MenuCaption As String
Dim m_ItemCaption As String

Private Type BITMAP
        bmType As Long
        bmWidth As Long
        bmHeight As Long
        bmWidthBytes As Long
        bmPlanes As Integer
        bmBitsPixel As Integer
        bmBits As Long
End Type

Private Declare Function GetMenu Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function GetSubMenu Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Private Declare Function GetMenuItemCount Lib "user32" (ByVal hMenu As Long) As Long
Private Declare Function GetMenuString Lib "user32" Alias "GetMenuStringA" (ByVal hMenu As Long, ByVal wIDItem As Long, ByVal lpString As String, ByVal nMaxCount As Long, ByVal wFlag As Long) As Long
Private Declare Function SetMenuItemBitmaps Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long, ByVal hBitmapUnchecked As Long, ByVal hBitmapChecked As Long) As Long
Private Declare Function ModifyMenu Lib "user32" Alias "ModifyMenuA" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, ByVal lpString As Any) As Long
Private Declare Function GetMenuItemID Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long

Private Function RemoveAmpersand(sStr As String) As String
    Dim p As Integer
    
    RemoveAmpersand = sStr
    
    p = InStr(sStr, "&")
    Do While p > 0
        If Mid$(sStr, p, 2) <> "&&" Then
            RemoveAmpersand = Mid$(sStr, 1, p - 1) & Mid$(sStr, p + 1)
            Exit Function
        Else
            p = p + 1
        End If
        p = InStr(p + 1, sStr, "&")
    Loop
End Function

Private Function GetMenuPos() As Long
    Dim hMnu As Long
    Dim MnuCount As Long
    Dim sName As String
    Dim i As Long
    hMnu = GetMenu(Parent.hwnd)
    MnuCount = GetMenuItemCount(hMnu)
    For i = 0 To MnuCount
        sName = String$(127, 0)
        Call GetMenuString(hMnu, i, sName, Len(sName), &H400)
        If Left$(RemoveAmpersand(sName), Len(m_MenuCaption)) = m_MenuCaption Then
            GetMenuPos = i
            Exit Function
        End If
    Next
    GetMenuPos = -1
End Function

Private Sub SetBitmaps(Sostitute As Boolean)
    Const MF_BITMAP = &H4&
    Const MF_BYPOSITION = &H400&
    
    Dim hMnu As Long
    Dim MnuPos As Long
    Dim hSubMnu As Long
    Dim MnuCount As Long
    Dim sName As String
    Dim i As Long
    
    MnuPos = GetMenuPos
    If MnuPos > -1 Then
        hMnu = GetMenu(Parent.hwnd)
        hSubMnu = GetSubMenu(hMnu, MnuPos)
        MnuCount = GetMenuItemCount(hSubMnu)
        For i = 0 To MnuCount
            sName = String$(127, 0)
            Call GetMenuString(hSubMnu, i, sName, Len(sName), &H400)
            If Left$(RemoveAmpersand(sName), Len(m_ItemCaption)) = m_ItemCaption Then
                If Sostitute Then
                    Call ModifyMenu(hSubMnu, i, MF_BITMAP Or MF_BYPOSITION, GetMenuItemID(hSubMnu, i), m_Image.Handle)
                Else
                    Call SetMenuItemBitmaps(hSubMnu, i, MF_BITMAP Or MF_BYPOSITION, m_ImageUnchecked.Handle, m_ImageChecked.Handle)
                End If
                Exit Sub
            End If
        Next
    End If
End Sub

Public Sub About()
Attribute About.VB_Description = "Show about box."
Attribute About.VB_UserMemId = -552
    frmSplash.Show vbModal
End Sub

Public Sub Modify()
Attribute Modify.VB_Description = "Changed  menu item with picture."
    If m_Image <> 0 Then Call SetBitmaps(True)
End Sub

Public Sub Refresh()
Attribute Refresh.VB_Description = "Apply picture to menu."
    If m_ImageChecked <> 0 Or m_ImageUnchecked <> 0 Then
        Call SetBitmaps(False)
    ElseIf m_Image <> 0 Then
        Call SetBitmaps(True)
    End If
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

Public Property Get Image() As Picture
Attribute Image.VB_Description = "Return or set  menu picture (13x13 pixel)."
    Set Image = m_Image
End Property

Public Property Set Image(ByVal New_Image As Picture)
    Set m_Image = New_Image
    PropertyChanged "Image"
End Property

Public Property Get ImageUnchecked() As Picture
Attribute ImageUnchecked.VB_Description = "Return or set  unchecked menu picture (13x13 pixel)."
    Set ImageUnchecked = m_ImageUnchecked
End Property

Public Property Set ImageUnchecked(ByVal New_ImageUnchecked As Picture)
    Set m_ImageUnchecked = New_ImageUnchecked
    PropertyChanged "ImageUnchecked"
End Property

Public Property Get ImageChecked() As Picture
Attribute ImageChecked.VB_Description = "Return or set  checked menu picture (13x13 pixel)."
    Set ImageChecked = m_ImageChecked
End Property

Public Property Set ImageChecked(ByVal New_ImageChecked As Picture)
    Set m_ImageChecked = New_ImageChecked
    PropertyChanged "ImageChecked"
End Property

Public Property Get MenuCaption() As String
Attribute MenuCaption.VB_Description = "Return or set  menu caption."
Attribute MenuCaption.VB_ProcData.VB_Invoke_Property = "Item"
    MenuCaption = m_MenuCaption
End Property

Public Property Let MenuCaption(ByVal New_MenuCaption As String)
    m_MenuCaption = RemoveAmpersand(New_MenuCaption)
    PropertyChanged "MenuCaption"
End Property

Public Property Get ItemCaption() As String
Attribute ItemCaption.VB_Description = "Return or set menu  item caption."
Attribute ItemCaption.VB_ProcData.VB_Invoke_Property = "Item"
    ItemCaption = m_ItemCaption
End Property

Public Property Let ItemCaption(ByVal New_ItemCaption As String)
    m_ItemCaption = RemoveAmpersand(New_ItemCaption)
    PropertyChanged "ItemCaption"
End Property

'Inizializza le proprietà di UserControl
Private Sub UserControl_InitProperties()
    Set m_Image = LoadPicture()
    Set m_ImageUnchecked = LoadPicture()
    Set m_ImageChecked = LoadPicture()
    m_MenuCaption = m_def_MenuCaption
    m_ItemCaption = m_def_ItemCaption
End Sub

'Carica i valori della proprietà dalla memoria
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    Set m_Image = PropBag.ReadProperty("Image", LoadPicture())
    Set m_ImageUnchecked = PropBag.ReadProperty("ImageUnchecked", LoadPicture())
    Set m_ImageChecked = PropBag.ReadProperty("ImageChecked", LoadPicture())
    m_MenuCaption = PropBag.ReadProperty("MenuCaption", m_def_MenuCaption)
    m_ItemCaption = PropBag.ReadProperty("ItemCaption", m_def_ItemCaption)
End Sub

'Scrive i valori della proprietà in memoria
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("Image", m_Image, LoadPicture())
    Call PropBag.WriteProperty("ImageUnchecked", m_ImageUnchecked, LoadPicture())
    Call PropBag.WriteProperty("ImageChecked", m_ImageChecked, LoadPicture())
    Call PropBag.WriteProperty("MenuCaption", m_MenuCaption, m_def_MenuCaption)
    Call PropBag.WriteProperty("ItemCaption", m_ItemCaption, m_def_ItemCaption)
End Sub

