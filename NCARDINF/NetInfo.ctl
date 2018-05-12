VERSION 5.00
Begin VB.UserControl NCardInfo 
   CanGetFocus     =   0   'False
   ClientHeight    =   480
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   495
   InvisibleAtRuntime=   -1  'True
   Picture         =   "NetInfo.ctx":0000
   PropertyPages   =   "NetInfo.ctx":030A
   ScaleHeight     =   480
   ScaleWidth      =   495
   ToolboxBitmap   =   "NetInfo.ctx":031C
End
Attribute VB_Name = "NCardInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "PropPageWizardRun" ,"Yes"
Option Explicit

Private m_objIpHelper As CIpHelper

Public Sub About()
Attribute About.VB_Description = "Show about box."
Attribute About.VB_UserMemId = -552
    frmSplash.Show vbModal
End Sub

Public Function GetInfo(iInterfaceIndex As Integer) As CInterface
Attribute GetInfo.VB_UserMemId = 0
    If iInterfaceIndex < 0 Or iInterfaceIndex > m_objIpHelper.Interfaces.Count Then Exit Function
    Set GetInfo = m_objIpHelper.Interfaces.Item(iInterfaceIndex)
End Function

'AVVISO: NON RIMUOVERE O MODIFICARE LE SEGUENTI RIGHE DI COMMENTO
'MemberInfo=7,1,1,0
Public Property Get Count() As Integer
Attribute Count.VB_Description = "Return number of interface mached."
    Count = m_objIpHelper.Interfaces.Count
End Property

Private Sub UserControl_Initialize()
    Set m_objIpHelper = New CIpHelper
    #If SHAREWARE = 1 Then
        frmSplash.Show vbModal
    #End If
End Sub

Private Sub UserControl_Resize()
    On Local Error Resume Next
    Width = 480
    Height = 480
End Sub

Private Sub UserControl_Terminate()
    Set m_objIpHelper = Nothing
End Sub

