VERSION 5.00
Begin VB.Form frmEvidences 
   BackColor       =   &H00CFC2AC&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "������������ ��������������"
   ClientHeight    =   9540
   ClientLeft      =   45
   ClientTop       =   465
   ClientWidth     =   15795
   BeginProperty Font 
      Name            =   "Calibri"
      Size            =   11.25
      Charset         =   204
      Weight          =   400
      Underline       =   0   'False
      Italic          =   -1  'True
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9540
   ScaleWidth      =   15795
   Begin VB.CheckBox chbNoMedCrim 
      BackColor       =   &H00CFC2AC&
      Caption         =   "�������� � ��������� �� ��� ���.����. ����."
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   28
      Top             =   8280
      Width           =   4935
   End
   Begin VB.TextBox txtBxStamp 
      Alignment       =   2  '���������
      BackColor       =   &H00FFE8DF&
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   12
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1485
      Left            =   1440
      MultiLine       =   -1  'True
      TabIndex        =   4
      Top             =   3840
      Width           =   3855
   End
   Begin VB.ComboBox cmbBxPackage 
      BackColor       =   &H00FFE8DF&
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   12
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   1440
      TabIndex        =   3
      Top             =   3360
      Width           =   3855
   End
   Begin VB.CommandButton cmdBxClear 
      BackColor       =   &H008080FF&
      Caption         =   "&��������"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   8760
      Width           =   1455
   End
   Begin VB.TextBox txtBxFirstDate 
      Alignment       =   2  '���������
      BackColor       =   &H00FFE8DF&
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   12
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   406
      Left            =   3360
      MaxLength       =   10
      TabIndex        =   1
      Top             =   840
      Width           =   1935
   End
   Begin VB.CommandButton cmdSaveOK 
      BackColor       =   &H00C0FFC0&
      Caption         =   "&���������"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3720
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   8760
      Width           =   1575
   End
   Begin VB.TextBox txtBxOutTake 
      Alignment       =   2  '���������
      BackColor       =   &H00FFE8DF&
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   12
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   406
      Left            =   2760
      MaxLength       =   10
      TabIndex        =   0
      Top             =   360
      Width           =   2535
   End
   Begin VB.ListBox lstEvList 
      Columns         =   1
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   12
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   8610
      Left            =   5640
      TabIndex        =   20
      Top             =   720
      Width           =   9615
   End
   Begin VB.TextBox txtBxPlace 
      Alignment       =   2  '���������
      BackColor       =   &H00FFE8DF&
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   12
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1365
      Left            =   360
      MultiLine       =   -1  'True
      TabIndex        =   2
      Top             =   1800
      Width           =   4935
   End
   Begin VB.CommandButton cmdBxCancel 
      BackColor       =   &H00C0C0FF&
      Caption         =   "&������"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2040
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   8760
      Width           =   1455
   End
   Begin VB.Frame fraInjuredPerson1 
      BackColor       =   &H00C0CFC5&
      Caption         =   "�������������� ��"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808000&
      Height          =   2655
      Left            =   360
      TabIndex        =   13
      Top             =   5520
      Width           =   4935
      Begin VB.TextBox txtOwSurName 
         BackColor       =   &H00DDEADB&
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   12
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   406
         Left            =   1560
         TabIndex        =   5
         Top             =   480
         Width           =   3135
      End
      Begin VB.TextBox txtOwName 
         BackColor       =   &H00DDEADB&
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   12
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   406
         Left            =   1560
         TabIndex        =   6
         Top             =   960
         Width           =   3135
      End
      Begin VB.TextBox txtOwMidName 
         BackColor       =   &H00DDEADB&
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   12
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   406
         Left            =   1560
         TabIndex        =   7
         Top             =   1440
         Width           =   3135
      End
      Begin VB.TextBox txtOwBirthday 
         BackColor       =   &H00DDEADB&
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   12
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   406
         Left            =   1920
         TabIndex        =   8
         Top             =   2010
         Width           =   1335
      End
      Begin VB.ListBox lstOwSex 
         BackColor       =   &H00DDEADB&
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   12
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   3840
         TabIndex        =   9
         Top             =   2040
         Width           =   855
      End
      Begin VB.Label lbl6 
         BackColor       =   &H00C0CFC5&
         Caption         =   "��� ��������"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12.75
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   240
         TabIndex        =   18
         Top             =   2050
         Width           =   1575
      End
      Begin VB.Label lbl5 
         BackColor       =   &H00C0CFC5&
         Caption         =   "��������"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12.75
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   240
         TabIndex        =   17
         Top             =   1440
         Width           =   1335
      End
      Begin VB.Label lbl4 
         BackColor       =   &H00C0CFC5&
         Caption         =   "���"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12.75
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   240
         TabIndex        =   16
         Top             =   960
         Width           =   1335
      End
      Begin VB.Label lbl3 
         BackColor       =   &H00C0CFC5&
         Caption         =   "�������"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12.75
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   240
         TabIndex        =   15
         Top             =   480
         Width           =   1335
      End
      Begin VB.Label Label5 
         Alignment       =   2  '���������
         BackColor       =   &H00C0CFC5&
         Caption         =   "���"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12.75
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   3240
         TabIndex        =   14
         Top             =   2040
         Width           =   615
      End
   End
   Begin VB.Label lblboxEvSum 
      BackColor       =   &H00CFC2AC&
      Caption         =   "���������� �� � ��������:"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   9480
      TabIndex        =   27
      Top             =   240
      Width           =   3255
   End
   Begin VB.Label lblBoxCounter 
      BackColor       =   &H00CFC2AC&
      Caption         =   "���������� ��������:"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   375
      Left            =   5640
      TabIndex        =   26
      Top             =   240
      Width           =   2775
   End
   Begin VB.Label Label3 
      BackColor       =   &H00CFC2AC&
      Caption         =   "������"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12.75
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   360
      TabIndex        =   25
      Top             =   4080
      Width           =   1095
   End
   Begin VB.Label Label2 
      BackColor       =   &H00CFC2AC&
      Caption         =   "��������"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12.75
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   360
      TabIndex        =   24
      Top             =   3415
      Width           =   1095
   End
   Begin VB.Label lbl7 
      BackColor       =   &H00CFC2AC&
      Caption         =   "���� �������������� ��"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12.75
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   360
      TabIndex        =   23
      Top             =   895
      Width           =   2895
   End
   Begin VB.Label lbl1 
      BackColor       =   &H00CFC2AC&
      Caption         =   "���� ������� ��:"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12.75
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   360
      TabIndex        =   22
      Top             =   390
      Width           =   2055
   End
   Begin VB.Label lblAllEvSumCounter 
      BackColor       =   &H00CFC2AC&
      Caption         =   "����� ����� ��:"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004000&
      Height          =   375
      Left            =   12960
      TabIndex        =   21
      Top             =   240
      Width           =   2295
   End
   Begin VB.Label lbl2 
      BackColor       =   &H00CFC2AC&
      Caption         =   "����� ������� ��:"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12.75
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   405
      Left            =   360
      TabIndex        =   19
      Top             =   1440
      Width           =   2295
   End
End
Attribute VB_Name = "frmEvidences"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'����� "frmEvidences"
'���� ������ "������������ ��������������"
'���� ��������: 01.06.2016
'@author Andr.Nab.n@gmail.com
Option Explicit
'Collection (������ �������� ������� � �������� ��  � ���)
'Public colEvidences As New Collection '��������� "��" (� ��������)
'Fields
'Private mvarstrBoxKey As String '���� ��� �������� � �� ="BX****"
'Private mvarstrEvKey As String  '���� ��� �������� (��)= "EV****"
'Private mvarstrKey As String    '����� ���� = "BX****" + "EV****" ("BX****EV****")
'Class
Public newEvOwner As clmEntrants '��������� ������ "�������� ��"
Private newEvDate As clmCaseDate '��������� ������ ����
Private tmpevKey As String
'Errors:
Const Msg As String = "������ ����� ������!"
'Colors
Private Enum frmColor
    LIGHT_YELLOW = &HC0FFFF 'RGB(102, 102, 153)
    CITRIC = &HFFFF& '������
    RED = &HFF&
    Black = &H0&
    LIGHT_GREEN = &HC0FFC0  '��������� RGB(200, 256, 200)
    LIGHT_OLIVE = &HDDEADB '��������� �������
    LIGHT_BLUE = &HFFE8DF   '&HFEF7F1
    PURPLE = &HFFC0C0
    BROWN = &H80FF&
End Enum
'Cancel
Dim Cancel As Integer
'
'=================== C O N S T R U C T O R ===================
'
Private Sub Form_Initialize()
'������������� ����� frmEvidences
     With Me
        .Caption = "����������� ��" 'frmNewEF.newEF.categories
        .Width = 5580
        .cmdBxClear.Visible = False
    End With
    Call EvClass_Initialize
End Sub
'
Private Sub Form_Load()
'��������� �������� �����
Move (Screen.Width - Width) / 2, (Screen.Height - Height) / 3.7 '/ 17.7
End Sub
'
Private Sub EvClass_Initialize()
'��������� ������������� ����� ����������� ������� "Evidences", "Owner":
    Set newEvOwner = New clmEntrants        '��������� ������ "�������� ��"
    Set newEvDate = New clmCaseDate         '��������� ������ ����
    Call frmNewEF.addNewBox
End Sub
'
Private Sub EvClass_Terminate()
'��������� ����������� ����� ����������� ������� "Evidences", "Owner":
    Set newEvOwner = Nothing   '����������� ���������� ������ "�������� ��"
    Set newEvDate = Nothing '��������� ������ ����
End Sub
'
Public Sub Form_Terminate()
    Call EvClass_Terminate
End Sub


'================== I C A P S U L A T I O N =============================
'
'
'==================== � � � � � � ======================================
'
Private Sub cmdEvClear_Click()
'��������� ������� ������ "��������"
'1) ����������� ����������� �������;
Call EvClass_Terminate
'2)������� ����� �����:
    Dim x As Object
        For Each x In Me.Controls
            If TypeName(x) = "TextBox" Then
                x.Text = ""
            ElseIf TypeName(x) = "ComboBox" Then
                x.Clear
            End If
        Next x
Call Ev_Enabled
txtBxOutTake.SetFocus
End Sub
'
Private Sub Ev_Enabled()
'��������� ��������� ����� "���������� ��������"
    Dim x As Object
        For Each x In Me.Controls
            If TypeName(x) = "TextBox" Or TypeName(x) = "ComboBox" Then
                x.BackColor = &HF1F9FA
            End If
        Next x
    txtBxFirstDate.SetFocus
End Sub
''
Private Sub cmdEvCancel_Click()
'��������� ������� ������ "������"
MsgBox "���� ��������!", vbOKOnly   '����� ���� ���������:
'    Set newEvDate = Nothing             '����������� ������� "NewEvid"
    Set newEvOwner = Nothing            '����������� ������� "newEvOwner"
'    Set mdCount.boxCounter = Nothing    '����������� ������� �������� "��������� �������� � ���.������"
'frmDocList.Show '����� ��������� �����
'    Call Open_DocListForm ??????? - ��������� � ������ �����.
    Call EvClass_Terminate
Me.Hide '�������� ����� "Evidences"
End Sub
'
Private Sub cmdSaveOK_Click()
'��������� ������� ������ "���������"
'1)'�������������� �����
    Me.Width = 15675
    With frmNewEF
        Dim str As String
'���������� ������ ++++++++++++++++++++++++
'        .newBox.strBxEntrants = create_strOwner
'         str = "�������� � ��������� " & .newBox.strBxEntrants & ":" '������������ �������� �������� ��:
'+++++++++++++++++++++++++++++++++++++++++
        Dim str1 As String
        str1 = create_strOwner(2, True)
        Me.Caption = str1
'        lstEvList.AddItem str1  '���������� ������ � ���� ����� frmEvidences:
        Me.lblBoxCounter.Caption = " ������ �������� �" & .colBoxes.Count + 1
    End With
    Call InputEvid_MsBxShow '����� ����������� ���� �������� ����� ������ ��������
End Sub
'
Private Static Sub InputEvid_MsBxShow()
'����� ����������� ���� "�������� ��?"
    If MsgBox("������� ����� ���.���. � ������?", vbYesNo) = vbYes Then
        Call InputEvid
    Else:
        With frmNewEF
            .newBox.strBxEntrants = create_strOwner(.newBox.colEvidences.Count, True)
            .colBoxes.Add .newBox, .strBoxKey '���������� ������� � ��  � ��������� �������
            Call .show_MsgCreateNewBox '����� ����������� ���� "������� ����� ������� � ��"
'           Me.Hide
            Unload Me
        End With
    End If
End Sub
'
Private Static Sub InputEvid()
'��������� �������� ������ ��
'1) ������������ ����� ��� ��
    With frmNewEF
        .strEvKey = .newBox.colEvidences.Count + 1
'        Debug.Print "���� ��� �������� = " & .strEvKey
    End With
'NB!!!! ������ ���� strKey ����������� � �������� Get_strKey
'3) ���� �������� ��
    With frmNewEF.newBox
        Dim strEvName As String
            strEvName = CStr(InputBox("������� �������� ������������� ��������������", "������ ������ �������", , 4000, 4300))
     '   �������� �� ������ ��������
        If strEvName = "" Then '
            Call InputEvid_MsBxShow
        Else
            .colEvidences.Add strEvName, frmNewEF.strKey '�������� �� � ��������� �� frmNewEF.newBox
        End If
'4) ������ �� ���������� � ��������� �� �����
        Me.lblboxEvSum.Caption = "���������� �� � ��������: " & .colEvidences.Count
        If chbNoMedCrim.Value = 0 Then
            frmNewEF.allEvSumCounter.increment '���������� �������� ����� ����� ��
            Me.lblAllEvSumCounter.Caption = "����� ����� ��: " & frmNewEF.allEvSumCounter.getTale
        End If
'5)  ��� � ���� �����:
        lstEvList.AddItem .colEvidences.Count & ". " & frmNewEF.newBox.colEvidences(frmNewEF.strKey)
    End With
 Call InputEvid_MsBxShow
End Sub
'
Private Static Function create_strOwner(Optional cnt As Long = 1, Optional version As Boolean = False) As String
'������� �������� ������ "�������������� ��"
'cnt - count ���������� �� � �������
'version - ��������� ������ ������ (False  - ������ 1, True - ������ 2)
Dim strTmp1 As String, strTmp2 As String, strTmp3 As String
    If newEvOwner.surName <> "" Then
        strTmp2 = newEvOwner.create_InitialslName  '" " &
    Else: strTmp2 = "(�������������� �� �������)"
    End If
    
    If cnt = 1 Then
        strTmp1 = "�������� "
    Else: strTmp1 = "��������� "
    End If
    
    If frmNewEF.newBox.strBxPlace <> "" Then
        If cnt = 1 Then
            If version = False Then
                strTmp3 = ", ������� " & frmNewEF.newBox.strBxPlace
            Else: strTmp3 = ", ������� " & frmNewEF.newBox.strBxPlace
            End If
        ElseIf cnt > 1 Then
            If version = False Then
                strTmp3 = ", ������� " & frmNewEF.newBox.strBxPlace
            Else: strTmp3 = ", �������� " & frmNewEF.newBox.strBxPlace
            End If
        End If
    Else: strTmp3 = ", (����� ������� �� �������)"
    End If
create_strOwner = frmNewEF.newBox.strBxPackage & " � " & strTmp1 & strTmp2 & strTmp3   '"�������" & strTmp1 & strTmp2
End Function
'
'==================== � � � �  � � � � � ===============================
'
Private Sub txtEnter(tmpObj As Object)
'��������� �������� ����� ��� ��������� ������
    With tmpObj
        .Text = ""
        .BackColor = frmColor.LIGHT_YELLOW 'RGB(102, 102, 153) 'frmColor.����1
        .ForeColor = frmColor.Black
    End With
End Sub
'
 Private Sub txt_Exit(tmpObj As Object)
'��������� �������� ����� ��� ������ ������
    With tmpObj
        .BackColor = frmColor.BROWN
        If .name = "txtOwBirthday" Or _
                .name = "lstOwSex" _
            Then
                .Text = "�� ������"
        ElseIf .name = "txtBxOutTake" _
                Or .name = "txtBxFirstDate" _
                Or .name = "cmbBxPackage" _
                Or .name = "txtBxStamp" _
                Or .name = "txtOwSurName" _
            Then
                .Text = "�� �������"
        ElseIf .name = "txtBxPlace" _
                Or .name = "txtOwName" _
                Or .name = "txtOwMidName" _
            Then
                .Text = "�� �������"
        Else
            .Text = "�� �������"
        End If
   End With
 End Sub
'
Private Sub txtBxOutTake_GotFocus()
'���� ������� ��:
    Call txtEnter(txtBxOutTake)
   '���� ������� �� = ���� ��������� �������������
    If frmNewEF.newEF.rulingDate <> 0 Then
         txtBxOutTake.Text = frmNewEF.newEF.rulingDate
    End If
End Sub
'
Private Sub txtBxOutTake_LostFocus()
'���� ������� ��:
 Dim str As String, dt As Date
 With txtBxOutTake
    If .Text = "" Then
        Call txt_Exit(txtBxOutTake)
    Else
        With newEvDate
            dt = .validateDate(.ExamDate(txtBxOutTake.Text)) '�������� ��������� ���� �� ����������
            frmNewEF.newBox.DtmBxOutTake = dt
            str = .dateToString(dt)
        End With
        .BackColor = frmColor.LIGHT_GREEN 'RGB(200, 256, 200) 'frmColor.���������
        .Text = str
    End If
End With
'Debug.Print "���� ������� ��: " & frmNewEF.newBox.DtmBxOutTake
End Sub
'
Private Sub txtBxFirstDate_GotFocus()
'���� �������������� ��
      Call txtEnter(txtBxFirstDate)
      txtBxFirstDate.Text = frmNewEF.newBox.DtmBxOutTake + 1
End Sub
'
Private Sub txtBxFirstDate_LostFocus()
'���� �������������� ��
 Dim tmp As Double
 Dim dt As Date
 With txtBxFirstDate
    If .Text = "" Then
        Call txt_Exit(txtBxFirstDate)
    Else
        With newEvDate
            dt = .validateDate(.ExamDate(txtBxFirstDate.Text))
'            '��������� ���: '���� �������������� �� >= ���� ������� ��
            Do
                tmp = .compareDt(frmNewEF.newBox.DtmBxOutTake, dt)
                If tmp < 0 Then
                    MsgBox "���� ������� �� ������ ���� �������������� ��!", _
                            vbCritical, Msg
                    dt = .ExamDate(InputBox("��������� ������� ���� �������������� ��!", _
                                        "���� ����", frmNewEF.newBox.DtmBxOutTake + 1))
                Else
                    frmNewEF.newBox.DtmBxFirstDate = dt
                    Exit Do
                End If
            Loop
        End With
            With txtBxFirstDate
                .BackColor = RGB(200, 256, 200) 'frmColor.���������
                .Text = newEvDate.dateToString(dt)
            End With
    End If
End With
Debug.Print "���� �������������� ��: " & frmNewEF.newBox.DtmBxFirstDate
End Sub
'
Private Sub txtBxPlace_GotFocus()
'����� ������� ��:
     Call txtEnter(txtBxPlace)
End Sub
'
Private Sub txtBxPlace_LostFocus()
'����� ������� ��:
      With txtBxPlace
        If .Text = "" Then
           Call txt_Exit(txtBxPlace)
        Else
            frmNewEF.newBox.strBxPlace = .Text
            .BackColor = frmColor.LIGHT_GREEN 'RGB(200, 256, 200) 'frmColor.���������
        End If
    End With
Debug.Print "����� ������� ��: " & frmNewEF.newBox.strBxPlace
End Sub
'
Private Sub cmbBxPackage_GotFocus()
     Call txtEnter(cmbBxPackage)
 With cmbBxPackage
        .Clear
        .Text = "��������� �������"
        .AddItem "��������� �������"
        .AddItem "����������� �����"
        .AddItem "���������"
        .AddItem "�������� �������"
        .AddItem "�������� �������"
        .AddItem "�������� �����"
    End With
End Sub
'
Private Sub cmbBxPackage_LostFocus()
'��������
    With cmbBxPackage
        If .Text = "" Then
           Call txt_Exit(cmbBxPackage)
        Else
'            frmNewEF.colBoxes.Item(frmNewEF.strBoxKey).strBxPackage = .Text
        frmNewEF.newBox.strBxPackage = .Text
            .BackColor = frmColor.LIGHT_GREEN
        End If
    End With
Debug.Print "�������� - " & frmNewEF.newBox.strBxPackage
End Sub
'
Private Sub txtBxStamp_GotFocus()
'������
    Call txtEnter(txtBxStamp)
    With txtBxStamp
        .Text = "������������� �������� ���������� �������� " & "��� �������"
    End With
End Sub
'
Private Sub txtBxStamp_LostFocus()
'������
    With txtBxStamp
        If .Text = "" Then
           Call txt_Exit(txtBxStamp)
        Else
'            frmNewEF.colBoxes.Item(frmNewEF.strBoxKey).strBxStamp = .Text
        frmNewEF.newBox.strBxStamp = Chr(171) & .Text & Chr(187)
            .BackColor = frmColor.LIGHT_GREEN
        End If
    End With
Debug.Print "������ - " & frmNewEF.newBox.strBxStamp
End Sub
'
Private Sub txtOwSurName_GotFocus()
'������� ��������� �� (Owner)
    Call txtEnter(txtOwSurName)
'    Dim str As String
'    str = frmNewEF.newInjPr.surName
'        With txtOwSurName
'            If str <> "" Then
'                .Text = str
'            End If
'        End With
End Sub
'
Private Sub txtOwSurName_LostFocus()
'������� ��������� �� (InjPrSurName)
    With txtOwSurName
        If .Text = "" Then
             Call txt_Exit(txtOwSurName)
        Else: newEvOwner.surName = StrConv(.Text, vbProperCase)
            .BackColor = RGB(200, 256, 200) 'frmColor.���������
            .Text = newEvOwner.surName
        End If
    End With
Debug.Print "������� ��������� ��: " & newEvOwner.surName
End Sub
'
Private Sub txtOwName_GotFocus()
'��� ��������� �� (InjPrName)
    Call txtEnter(txtOwName)
'    Dim str As String
'    str = frmNewEF.newInjPr.name
'        With txtOwName
'            If str <> "" Then
'                .Text = str
'            End If
'        End With
End Sub
'
Private Sub txtOwName_LostFocus()
'��� ��������� �� (InjPrName)
      With txtOwName
        If .Text = "" Then
             Call txt_Exit(txtOwName)
        Else: newEvOwner.name = StrConv(.Text, vbProperCase)
            .BackColor = RGB(200, 256, 200) 'frmColor.���������
            .Text = newEvOwner.name
        End If
    End With
Debug.Print "��� ��������� ��: " & newEvOwner.name
End Sub
''
Private Sub txtOwMidName_GotFocus()
'�������� ��������� �� (Owner)
    Call txtEnter(txtOwMidName)
'    Dim str As String
'    str = frmNewEF.newInjPr.midName
'        With txtOwMidName
'            If str <> "" Then
'                .Text = str
'            End If
'        End With
End Sub
'
Private Sub txtOwMidName_LostFocus()
'�������� ��������� �� (InjPrMidName)
     With txtOwMidName
        If .Text = "" Then
             Call txt_Exit(txtOwMidName)
        Else: newEvOwner.midName = StrConv(.Text, vbProperCase)
            .BackColor = RGB(200, 256, 200) 'frmColor.���������
            .Text = newEvOwner.midName
        End If
    End With
Debug.Print "�������� ��������� ��: " & newEvOwner.midName
End Sub
'
Private Sub txtOwBirthday_GotFocus()
'���� �������� ��������� �� (InjPrBirthday):
    Call txtEnter(txtOwBirthday)
'    Dim str As String
'    str = frmNewEF.newInjPr.birthday
'        With txtOwBirthday
'            If str <> "" Then
'                .Text = str
'            End If
'        End With
End Sub
'
Private Sub txtOwBirthday_LostFocus()
'���� �������� ��������� �� (InjPrBirthday):
    With txtOwBirthday
        If .Text = "" Then
             Call txt_Exit(txtOwBirthday)
        Else: newEvOwner.birthday = StrConv(.Text, vbProperCase)
            .BackColor = RGB(200, 256, 200) 'frmColor.���������
            .Text = newEvOwner.birthday
        End If
    End With
Debug.Print "���� �������� ��������� ��: " & newEvOwner.birthday
End Sub
'
Private Sub lstOwSex_GotFocus()
'��� ��������� ��
    Dim str As String
    str = frmNewEF.newInjPr.sex
         With lstOwSex
            .Clear
            .AddItem "���."
            .AddItem "���."
            .AddItem "��� �� ������"
            .BackColor = &HC0FFFF
            If str <> "" Then
                .Text = str
            End If
        End With
End Sub
'
Private Sub lstOwSex_LostFocus()
'��� ��������� ��
    With lstOwSex
        newEvOwner.sex = .Text
        .BackColor = RGB(200, 256, 200)
    End With
Debug.Print "��� ��������� ��: " & newEvOwner.sex
End Sub
'




'================= L I B ==========================.
'Private Sub cmbBxCategory_GotFocus()
''��������� ��
' Call txtEnter(cmbBxCategory)
' With cmbBxCategory
'        .Clear
'        .Text = "�������� ������-������������������"
''   ������ � �������� �������� ���������
'    Dim i As Long
'    If UBound(mdPrintDoc.arrEF) <> 0 Then
'            For i = LBound(mdPrintDoc.arrEF) To UBound(mdPrintDoc.arrEF)
'                .AddItem mdPrintDoc.arrEF(i)
'            Next i
'        End If
'    End With
'End Sub
''Private Sub cmbBxCategory_LostFocus()
''��������� ��
'  With cmbBxCategory
'        If .Text = "" Then
'           Call txt_Exit(cmbBxCategory)
'        Else
'            frmNewEF.newBox.strBxCategory = .Text
''            frmNewEF.colBoxes.Item(frmNewEF.strBoxKey).strBxCategory = .Text
'            .BackColor = frmColor.LIGHT_GREEN
'        End If
'    End With
'Debug.Print "cmbBxCategory - " & frmNewEF.newBox.strBxCategory
'End Sub

'


'=================== C O L L E C T I O N =================================
'1) ������ � ����������
'
'Function keyExists(coll As Collection, _
'                key As String) As Boolean
''���������,���������� �� ����.
'On Error GoTo EH
'    coll.Item key
'    keyExistsExists = True
'EH: '"key is not exists"
'End Function
''
'Sub QuickSort(coll As Collection, _
'                first As Long, _
'                last As Long)
''������� ���������� QuickSort
'Dim vCentreVal As Variant
'Dim vTemp As Variant
'Dim lTempLow As Long
'Dim lTempHi As Long
'    lTempLow = first
'    lTempHi = last
'    vCentreVal = coll((first + last) \ 2)
'    Do While lTempLow <= lTempHi
'
'        Do While coll(lTempLow) < vCentreVal And lTempLow < last
'            lTempLow = lTempLow + 1
'        Loop
'
'        Do While vCentreVal < coll(lTempHi) And lTempHi > first
'            lTempHi = lTempHi - 1
'        Loop
'
'        If lTempLow <= lTempHi Then
'    '�������� ��������
'            vTemp = coll(lTempLow)
'            coll.Add coll(lTempHi), After:=lTempLow
'            coll.Remove lTempLow
'            coll.Add vTemp, Before:=lTempHi
'            coll.Remove lTempHi + 1
'    '������� � ��������� ��������
'            TempLow = lTempLow + 1
'            lTempHi = lTempHi - 1
'        End If
'    Loop
'    If first < lTempHi Then
'            QuickSort coll, first, lTempHi
'            If lTempLow < last Then
'                QuickSort coll, lTempLow, last
'            End If
'    End If
'End Sub
'

'Private Sub Form_Initialize()
''������������� �����
'     With Me
'        .Caption = "����������� ��" 'frmNewEF.newEF.categories
'        .Width = 5580
'        .cmdBxClear.Visible = False
'    End With
'    Call EvClass_Initialize
'    Call zeroControlStatus
'End Sub
'
'Private Sub zeroControlStatus()
''��������� ����������� ����� ����� � ��������� ���������
'    With Me
'        Dim X As Object
'            For Each X In .Controls
'                If TypeName(X) = "TextBox" Then
'                    X.BackColor = frmColor.LIGHT_BLUE '&HFFE8DF   '&HF1F9FA
'                    X.Text = ""
'                        If X.name = "txtOwSurName" Or _
'                            X.name = "txtOwName" Or _
'                            X.name = "txtOwMidName" Or _
'                            X.name = "txtOwBirthday" _
'                            Then
'                                X.BackColor = frmColor.LIGHT_OLIVE
'                        End If
'                ElseIf TypeName(X) = "ListBox" Then
'                    X.BackColor = frmColor.LIGHT_OLIVE 'If TypeName(X) = "ListBox" Then
'                ElseIf TypeName(X) = "ComboBox" Then
'                    X.Clear
'                    X.BackColor = frmColor.LIGHT_BLUE
'                End If
'            Next X
'    End With
'End Sub
'
'Private Sub EvClass_Initialize()
''��������� ������������� ����� ����������� ������� "Evidences", "Owner":
'    Set newEvOwner = New clmEntrants        '��������� ������ "�������� ��"
'    Set newEvDate = New clmCaseDate         '��������� ������ ����
'    Set frmNewEF.newBox = New clmEvBox '��������� ������ "�������� � ��"
''   ��������:
''    Set mdCount.boxEvSumCounter = New clmCounter '��������� ������ �������� "����� ���.����� (��) � ��������"
''        With mdCount.boxEvSumCounter
''            .name = "��������� ���.����� � �������� "
''            Debug.Print .name & Chr(32) & .getTale
''        End With
''������� ����� ����� (���������� ��� ������������ ����������� �����)
'Call zeroControlStatus
'End Sub
'
'=====================� � � � � �  � � � � �: ===========================
'Private Sub cmdBxCancel_Click()
''��������� ������� ������ "������"
'MsgBox "���� ��������!", vbOKOnly '����� ���� ���������:
'Set newBxid = Nothing '����������� ������� "NewEvid"
'Set newEvOwner = Nothing '����������� ������� "newEvOwner"
''frmDocList.Show '����� ��������� �����
'Call Open_DocListForm
'Me.Hide '�������� ����� "Evidences"
'End Sub
''
'Private Sub cmdEvClear_Click()
''��������� ������� ������ "��������"
''1) ����������� ����������� �������;
'Call EvClass_Terminate
''2)������� ����� �����:
'    Dim X As Object
'        For Each X In Me.Controls
'            If TypeName(X) = "TextBox" Then
'                X.Text = ""
'            ElseIf TypeName(X) = "ComboBox" Then
'                X.Clear
'            End If
'        Next X
'Call Ev_Enabled
'txtEvOutTake.SetFocus
'End Sub
'

'Private Static Sub addNewEvBox_MsgShow()
''����� ����������� ���� "������� ����� ������ ��"
'    If MsgBox("������� ����� ������ �������� (��)?", vbYesNo) = vbYes Then
'        Call cmdEvClear_Click
'        Call EvClass_Initialize
'    Else:
''������� ����� frmDocList
'        Dim tmGr As String
'            tmGr = CStr(Format(Me.lblEvidCount, "#000")) & "000" & "0000"
'            frmNewEF.colEvidences.Add Me.lblEvGRSum, "GrCount" & tmGr
'        MsgBox "���� �������� ��������!", vbOKOnly
'        Call EvClass_Terminate
'        Call Open_DocListForm
'        Me.Hide
'    End If
'''Debug.Print "��������� ����� ��= " & frmNewEF.EvGrCount
'End Sub

' Private Sub changeBoxCounter()
'   With mdCount.boxCounter
'        .increment '���������� (+1) �������� ��������
'            Debug.Print .name & Chr(32) & .getTale
''       ��������� ������� "lblBoxCounter" �� �����
'        lblBoxCounter.Caption = .name & " - " & .getTale
''       �������� �������� ����� ��� ��������
'        strBoxKey = CStr(Format(.getTale, "#0000"))
'            Debug.Print "���� ��� �������� = " & strBoxKey
''       ������ ������ � ���������
'        Dim tmp As String
'            tmp = Create_strOwner
'        frmNewEF.colBoxes.Add tmp, strBoxKey
'Debug.Print "�������� �������� - " & frmNewEF.colBoxes.Item(strBoxKey)
'Debug.Print "�������� ����� = " & strBoxKey
''       ���������� ������ � ���� ����� frmEvidences:
'        lstEvList.AddItem tmp
'    End With
' End Sub

'Private Static Sub InputEvid()
''��������� �������� ������ ��
''1) ���� �������� ��
'    With frmNewEF.newBox
'        .strBxName = CStr(InputBox("������� �������� ������������� ��������������", "������ ������ �������", , 4000, 4300))
'     '   �������� �� ������ ��������
'        If .strBxName = "" Then '
'           Call InputEvid_MsBxShow
'        Else
''       1)������� "����� ��" � ��������
'            With mdCount.boxEvSumCounter
'                .increment '���������� (+1) �������� �������� "����� ��" � ��������
'                    Debug.Print .name & " = " & .getTale
''               �������� �������� ����� ��� ��������
'                strEvKey = CStr(Format(.getTale, "#0000"))
'                    Debug.Print "���� ��� �������� = " & strEvKey
''               ��������� ������� "lblBoxCounter.Caption" �� �����
'                lblboxEvSum.Caption = .name & " - " & .getTale
'            End With
''        2)������� "����� ����� ��"
'            With mdCount.allEvSumCounter
'                .increment '���������� (+1) �������� �������� "����� ��" � ��������
'                   Debug.Print .name & Chr(32) & .getTale
''               ��������� ������� "lblAllEvSumCounter" �� �����
'            lblAllEvSumCounter.Caption = .name & " - " & .getTale
'            End With
''        3)������������ ������� �����:
'            Dim tmpEv As String
'            tmpEv = strBoxKey & strEvKey
'                Debug.Print "������ ���� ��� �������� = " & tmpEv
''        4)������ �� � ���������:
'            frmNewEF.newBox.colEvidences.Add .strBxName, tmpEv
'                Debug.Print "������ �� � ��������� -" & frmNewEF.newBox.colEvidences.Item(tmpEv) & Chr(10) _
'                & "��������� �������� � ��������� - " & frmNewEF.newBox.colEvidences.Count
''           ����������� �� � ���� �����:
''        5)������ �� � ����� ��������� Box & ��
'             frmNewEF.colBoxes.Add .strBxName, tmpEv
'             Debug.Print "���������� ������� = " & frmNewEF.colBoxes.Item(tmpEv)
'             Debug.Print "�������� ����� " & tmpEv
'            lstEvList.AddItem (mdCount.boxEvSumCounter.getTale) & ". " & frmNewEF.newBox.colEvidences(tmpEv)
'        End If
'    End With
' Call InputEvid_MsBxShow
'End Sub

'���������� ������ ���������� ������ "��":
'Private NewEvid As clmEvidences
'Private newEvOwner As Entrants
''
'Public Property Let fEvCount(ByVal vData As Integer)
''������� ����������� ���� "��"
'mvarfEvCount = vData
'End Property
''
'Public Property Get fEvCount() As Integer
''������� ����������� ���� "��"
'fEvCount = mvarfEvCount
''Debug.Print "������� ����������� ���� �� = ", fEvCount
'End Property
''

''
''
'Private Sub lstOwSex_LostFocus()
''��� ������������
'    With lstOwSex
'        newEvOwner.En_Sex = .Text
'        .BackColor = RGB(200, 256, 200)
'    End With
'End Sub
''

''
'Private Sub lstOwSex_GotFocus()
''��� ������������
'    With lstOwSex
'        .Clear
'        .AddItem "���."
'        .AddItem "���."
'        .AddItem "��� �� ������"
'        .BackColor = &HC0FFFF
'    End With
'End Sub


'Private Sub txtEvOutTake_GotFocus()
''���� ������� �� (EvOutTake)
'    With txtEvOutTake
'        .Text = ""
'        .BackColor = &HC0FFFF
'        .ForeColor = &H80000008
'    End With
'End Sub
''
'Private Sub txtEvOutTake_LostFocus()
''���� ������� �� (EvOutTake)
'        With txtEvOutTake
'            If .Text = "" Then
'                .Text = "�� �������"
'                .BackColor = &HC0E0FF
'            Else:
'                Do
'                    If Not IsDate(.Text) Then
'                        Beep
'                        .BackColor = RGB(256, 0, 0)
'                        MsgBox "������������ ������� ����!", vbCritical, "������ �����!"
'                        .Text = InputBox("������� ��������� ����!", _
'                            "���� ���������� ����")
'                    Else: NewEvid.DtmEvOutTake = CDate(.Text)
'                        .BackColor = RGB(200, 256, 200)
'                Exit Do
'                    End If
'                Loop
'            End If
'    End With
'    cmdEvClear.Visible = True
''�������� ��������� ��
''NewEvid.strEvCategory = Me.Caption
'End Sub
''
'Private Sub txtEvPlace_GotFocus()
''����� ������� �� (EvPlace)
'    With txtEvPlace
'        .Text = ""
'        .BackColor = &HC0FFFF
'        .ForeColor = &H80000008
'    End With
'End Sub
''
'Private Sub txtEvPlace_LostFocus()
''����� ������� �� (EvPlace)
'    With txtEvPlace
'        If Len(.Text) = 0 Then
'            .Text = "(����� ������� �������� �� �������)"
'            .BackColor = &HC0E0FF
'        Else: NewEvid.strEvPlace = .Text
'            .BackColor = RGB(200, 256, 200)
'        End If
'    End With
'End Sub

''
'Private Sub txtOwName_GotFocus()
''��� ��������� �� (InjPrName)
'    With txtOwName
'        .BackColor = &HC0FFFF
'        .ForeColor = &H80000008
'        If .Text = "(��� �� �������)" Then
'            .Text = ""
'        End If
'    End With
'End Sub
''
'Private Sub txtOwName_LostFocus()
''��� ��������� �� (InjPrName)
'    With txtOwName
'        If Len(.Text) = 0 Then
'            .Text = "(��� �� �������)"
'            .BackColor = &HC0E0FF
'        Else
'            If Len(.Text) = 1 Then
'                newEvOwner.En_Name = StrConv(.Text, vbProperCase) & "."
'            Else: newEvOwner.En_Name = StrConv(.Text, vbProperCase)
'            End If
'        .Text = newEvOwner.En_Name
'        .BackColor = RGB(200, 256, 200)
'        End If
'    End With
'End Sub
''
'Private Sub txtOwMidName_GotFocus()
''�������� ��������� �� (Owner)
'    With txtOwMidName
'        .BackColor = &HC0FFFF
'        .ForeColor = &H80000008
'        .BackColor = &HC0FFFF
'        If .Text = "(�������� �� �������)" Then
'            .Text = ""
'        End If
'    End With
'End Sub
''
'Private Sub txtOwMidName_LostFocus()
''�������� ��������� �� (InjPrMidName)
'    With txtOwMidName
'        If Len(.Text) = 0 Then
'            .Text = "(�������� �� �������)"
'            .BackColor = &HC0E0FF
'        Else
'            If Len(.Text) = 1 Then
'                newEvOwner.En_MidName = StrConv(.Text, vbProperCase) & "."
'            Else: newEvOwner.En_MidName = StrConv(.Text, vbProperCase)
'            End If
'        .Text = newEvOwner.En_MidName
'        .BackColor = RGB(200, 256, 200)
'        End If
'    End With
'End Sub
''
'Private Sub txtOwBirthday_GotFocus()
''���� �������� ��������� �� (InjPrBirthday):
'    With txtOwBirthday
'        .BackColor = &HC0FFFF
'        .ForeColor = &H80000008
'    End With
'End Sub
''
'Private Sub txtOwBirthday_LostFocus()
''���� �������� ��������� �� (InjPrBirthday):
'    With txtOwBirthday
'        If .Text = "" Then
'            .Text = "�� ������"
'            .BackColor = &HC0E0FF
'        Else: newEvOwner.En_Birthday = .Text
'            .BackColor = RGB(200, 256, 200)
'        End If
'    End With
'End Sub
''
'Private Sub txtEvFirstDate_GotFocus()
''���� �������������� ��
'    With txtEvFirstDate
'        .BackColor = &HC0FFFF
'        .ForeColor = &H80000008
'    End With
'End Sub
''
'Private Sub txtEvFirstDate_LostFocus()
''���� �������������� ��
'    With txtEvFirstDate
'            If .Text = "" Then
'                .Text = "�� �������"
'                .BackColor = &HC0E0FF
'            Else
'                Do
'                    If Not IsDate(.Text) Then  'Or (CDate(.Text) - NewEvid.DtmEvOutTake) < 0
'                        Beep
'                        .BackColor = RGB(256, 0, 0)
'                        MsgBox "����������� ������� ����!", vbCritical, "������ �����!"
'                        .Text = InputBox("������� ��������� ����!", _
'                            "���� ���������� ����")
'                    Else: NewEvid.DtmEvFirstDate = CDate(.Text)
'                        NewEvid.blnEvProvide = True '������������ = "��"
'                        .BackColor = RGB(200, 256, 200)
'                Exit Do
'                    End If
'                Loop
'            End If
'    End With
'End Sub
''
'Private Static Sub addNewEvBox_MsgShow()
''����� ����������� ���� "������� ����� ������ ��"
'    If MsgBox("������� ����� ������ �������� (��)?", vbYesNo) = vbYes Then
'        Call cmdEvClear_Click
'        Call EvClass_Initialize
'    Else:
''������� ����� frmDocList
'        Dim tmGr As String
'            tmGr = CStr(Format(Me.lblEvidCount, "#000")) & "000" & "0000"
'            frmNewEF.colEvidences.Add Me.lblEvGRSum, "GrCount" & tmGr
'        MsgBox "���� �������� ��������!", vbOKOnly
'        Call EvClass_Terminate
'        Call Open_DocListForm
'        Me.Hide
'    End If
'''Debug.Print "��������� ����� ��= " & frmNewEF.EvGrCount
'End Sub
''
'Private Sub Open_DocListForm() '��������� �������� ����� "�������� ����������"
'    mdCount.fDocListCount = mdCount.fDocListCount + 1
'    Dim tmp As String '��������� ����������-�������
'        tmp = Me.lblEvidCount '��������� ������-������ ����� �������� ����� frmEvidences
'    Dim frmD As frmDocList '��������� ����� �����
'    Set frmD = New frmDocList
''���������� ����� � ������ ��������� ����
'        frmD.Caption = mdPrintDoc.DocCat & Chr(32) & tmp
'        frmD.lblDocListCount.Caption = tmp
'        mdCount.colForms.Add frmD, "DocListform" & tmp
'    frmD.Show
'    mdCount.colForms("Evidform" & tmp).Hide
'End Sub
''

'Private Sub Create_colOwner()
''��������� ������ ������� ����� ����� � ��������� ��
''1)��������� ���������:
'newEvOwner.EnCount = newEvOwner.EnCount + 1
'Me.lblEvGRSum = EvGRCount
''������������ ������-�����: 000_000_0000(������������_��������������_���������)
'    Dim tmp As String
'        tmp = CStr(Format(Me.lblEvidCount, "#000")) & CStr(Format(mdCount.EvGRCount, "#000")) & "0000"
'    With newEvOwner '"��������"
'        frmNewEF.colEvidences.Add .Create_InitialsFullName, "newEvOwnerFullName" & tmp '��� ��������� ��
'        frmNewEF.colEvidences.Add .En_Birthday, "newEvOwnerBirthday" & tmp '��� �������� ��
'        frmNewEF.colEvidences.Add .En_Sex, "newEvOwnerSex" & tmp '��� ��������� ��
'    End With
'    With NewEvid '"������������ ��������������"
'        frmNewEF.colEvidences.Add .DtmEvOutTake, "EvOutTake" & tmp  '���� �������
'        frmNewEF.colEvidences.Add .DtmEvFirstDate, "EvFirstDate" & tmp  '���� ��������������
'        frmNewEF.colEvidences.Add .strEvPlace, "EvPlace" & tmp  '����� �������
'        frmNewEF.colEvidences.Add .strEvPackage, "EvPackage" & tmp  '��������
'        frmNewEF.colEvidences.Add .strEvStamp, "EvStamp" & tmp  '������
'        NewEvid.Print_EvNumArr1
'    End With
''���������� ������ "�������������� ��������"
'       frmNewEF.colEvidences.Add Create_strOwner, "Owner" & tmp
''���������� ������ � ������:
'        lstEvList.AddItem Create_strOwner
''Debug.Print "���� =" & tmp
'Debug.Print "�������� = " & frmNewEF.colEvidences("Owner" & tmp)
'End Sub
''
'Private Static Function Create_strOwner() As String
''������� �������� ������ "�������������� ��"
'Dim strTmp1 As String, strTmp2 As String
'    If newEvOwner.En_SurName <> "" Then
'        strTmp1 = " " & newEvOwner.Create_InitialsFullName
'    Else: strTmp1 = " (�������������� �������� �� �������)"
'    End If
'    If NewEvid.strEvPlace <> "" Then
'        strTmp2 = ", ������� �� ������: " & NewEvid.strEvPlace & ":"
'    Else: strTmp2 = ", (����� ������� �� �������):"
'    End If
'Create_strOwner = "�������" & strTmp1 & strTmp2
'End Function
''


''Private Sub Ev_Disabled()
'''��������� ���������� ����� "���������� ��������"
''    Dim X As Object
''        For Each X In Me.Controls
''            If InStrRev(X.Name, "txtEE", 5) > 0 Or InStrRev(X.Name, "cboEF", 5) > 0 Then
''                X.Enabled = False
''                X.BackColor = &HFDEADB
''            End If
''        Next X
''End Sub
'
'
'
'
''
''Private Sub Erase_ListItem()
'''�������� � ������� ������ �� ������ �����
''    For Item = lstEvList.ListCount - 1 To 0 Step -1
''        If List1.List(Item) = "" Then
''            List1.RemoveItem Item
''        End If
''    Next
''End Sub
'
''������� �������� ������������� �������
''Dim X As Object
''Dim EvCount As Integer
''EvCount = LowBound
''Debug.Print "�������� ������� frmFlArr: "
''    Select Case EvCount
''        Case Is <= 5
''            txtEvOutTake.SetFocus
''            For Each X In Me.Controls
''                If TypeName(X) = "TextBox" Then
''                    ReDim Preserve frmFlArr(EvCount)
''                    frmFlArr(EvCount) = X.Text
''                    Debug.Print EvCount & "." & frmFlArr(EvCount)
''                    EvCount = EvCount + 1
''                End If
''            Next X
''        Case Is = 6
''            frmFlArr(EvCount) = lstOwSex.Text
''            Debug.Print EvCount & "." & frmFlArr(EvCount)
''        Case Is = 7
''            frmFlArr(EvCount) = newEvOwner.Create_InitialsInjPrFullName
''            Debug.Print frmFlArr(EvCount)
''        End Select
'
'''���������� ������ �����
''lstEvList.AddItem Create_strOwner1
''lblEvGRSum.Caption = NewEvid.mdCount.EvGRCount
'''������� �������:
''Dim v As Integer, d As Integer
''    If d = 0 Then
''        For v = LowBound To n
''            Debug.Print "������ �" & d & "_" & v & " ������� = " & arrEvidences(d, v)
''        Next v
''    Else
''        For d = LowBound To NewEvid.mdCount.EvGRCount
''            For v = LowBound To n
''            Debug.Print "������ �" & d & "_" & v & " ������� = " & arrEvidences(d, v)
''        Next v
''    Next d
''    End If
''End Sub

''Function IsNotEmptyArray(parArray As Variant) As Boolean
'''������� �������� ������������� �������
''  On Error Resume Next
''  IsNotEmptyArray = LBound(parArray) <= UBound(parArray)
''End Function
'
