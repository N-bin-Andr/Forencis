VERSION 5.00
Begin VB.MDIForm frmMDI 
   BackColor       =   &H00808000&
   ClientHeight    =   9450
   ClientLeft      =   60
   ClientTop       =   855
   ClientWidth     =   12405
   LinkTopic       =   "MDIForm1"
   WindowState     =   2  '����������
   Begin VB.Menu mnuFile 
      Caption         =   "&�������"
      Begin VB.Menu mnuTanatology 
         Caption         =   "&���������� �� ��������"
      End
   End
   Begin VB.Menu mnuChange 
      Caption         =   "&��������"
      Begin VB.Menu mnuFolder 
         Caption         =   "&����� ��� ���������� ����������"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&�������"
   End
End
Attribute VB_Name = "frmMDI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Main Form
'@author Andr.N@_bin
'@E-mail Andr.Nab.n@gmail.com
Option Explicit
Private newFrm As Form
Private tmpcategory As String
Private SvcService As Object        '������ ���������� svcsvc.dll
Private tmpnewRoot As String
'
'++++++++++++++++++  � � � � � � � � � � � � � ++++++++++++++
'
Private Sub MDIForm_Initialize()
    Me.Caption = "��� �3 �������� ����� ����������."
End Sub
'
Private Sub MDIForm_Unload(Cancel As Integer)
'��������� �������� �����
    If MsgBox("������� ����������?", vbYesNo, "�����?") = vbYes Then
'        Set newFrm = Nothing
    Else
        Cancel = 1
    End If
End Sub
'
'++++++++++++++++++ � � � � � � � � � � � � ++++++++++++++++++
Private Property Let category(ByVal vData As String)
'��������� ������������ ���������
    tmpcategory = vData
End Property
'
Friend Property Get category() As String
'��������� ������������ ���������
    category = tmpcategory
'Debug.Print "category: " & category
End Property
'
Private Property Let newRoot(ByVal vData As String)
'����� ����������, ��������� �������������
    tmpnewRoot = vData
End Property
'
Friend Property Get newRoot() As String
'����� ����������, ��������� �������������
    newRoot = tmpnewRoot
'Debug.Print "����� ����������, ��������� �������������: " & newRoot
End Property
'
'++++++++++++++++++ � � � � � � ++++++++++++++++++++++++++++++
Private Sub mnuHelp_Click()
frmAbout.Show
End Sub
'
Public Sub mnuTanatology_Click()
    category = "�����������"
'    Set newFrm = frmNewResearch
        With frmNewResearch
'            Call .addExpert
            .Show
        End With
End Sub
'
Private Sub mnuFolder_Click()
    Set SvcService = CreateObject("Svcsvc.Service") '������ ����������  svcsvc.dll
        newRoot = SvcService.SelectFolder("�������� ����� ��� ���������� ����������", "", &H10 + &H4000, "")
    Set SvcService = Nothing    '������ ����������  svcsvc.dll
'Debug.Print "����� ����������, ��������� �������������: " & newRoot
End Sub

