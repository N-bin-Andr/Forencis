VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmCraniumDimensions 
   Caption         =   "����������������� ����� ������������ ������"
   ClientHeight    =   9165
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6630
   OleObjectBlob   =   "frmCraniumDimensions.frx":0000
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "frmCraniumDimensions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


'@author Andr.Nab.n@gmail.com
'@01/01/2015
'
Option Explicit
'
Public �� As Counter
Public �� As Counter
Public ����� As Counter
Public �� As Counter
Public �� As Counter
Public ��� As Counter
Public myForm As Object
Public N As Counter
'���������� ���������� ������� ��������� ���������� (������� "����� ������������")
Private ������(3, 24) As String
Private dimension(24) As Currency
'���������� ���������� ����������, �������� �����
Const IMGdir As String = "D:\Crime\MasterForm\Bons\Cranium\VB\"
'
Private CurDimensions As Currency
'
Private Static Sub Select_Dimensions(ByVal CurDimensions As Currency)
'��������� ���������� "���������� �������: ��������-�������������" � ���������� �������
    Dim tmp As Integer
        tmp = N.Count
    Select Case tmp '������������ ����������, � ��� ����������� ���������
        Case Is = 0
            dimension(tmp) = CurDimensions
            lblSex1.Caption = Select_����������_�������(ByVal CurDimensions)
        Case Is = 1
            dimension(tmp) = CurDimensions
            lblSex1.Caption = Select_����������_�������(ByVal CurDimensions)
        Case Is = 2
            dimension(tmp) = CurDimensions
            lblSex1.Caption = Select_��������_�������(ByVal CurDimensions)
        Case Is = 3
            dimension(tmp) = CurDimensions
            lblSex1.Caption = Select_�����_���������(ByVal CurDimensions)
        Case Is = 4
            dimension(tmp) = CurDimensions
            lblSex1.Caption = Select_Min_������_���(ByVal CurDimensions)
        Case Is = 5
            dimension(tmp) = CurDimensions
            lblSex1.Caption = Select_������_���������(ByVal CurDimensions)
        Case Is = 6
            dimension(tmp) = CurDimensions
            lblSex1.Caption = Select_������_�������(ByVal CurDimensions)
        Case Is = 7
            dimension(tmp) = CurDimensions
            lblSex1.Caption = Select_�����������_������(ByVal CurDimensions)
        Case Is = 8
            dimension(tmp) = CurDimensions
            lblSex1.Caption = Select_����������_������(ByVal CurDimensions)
        Case Is = 9
            dimension(tmp) = CurDimensions
            lblSex1.Caption = Select_������������_�����(ByVal CurDimensions)
        Case Is = 10
            dimension(tmp) = CurDimensions
            lblSex1.Caption = Select_������_�����(ByVal CurDimensions)
        Case Is = 11
            dimension(tmp) = CurDimensions
            lblSex1.Caption = Select_��������_�����(ByVal CurDimensions)
        Case Is = 12
            dimension(tmp) = CurDimensions
            lblSex1.Caption = Select_�����_���(ByVal CurDimensions)
        Case Is = 13
            dimension(tmp) = CurDimensions
            lblSex1.Caption = Select_������_���(ByVal CurDimensions)
        Case Is = 14
            dimension(tmp) = CurDimensions
            lblSex1.Caption = Select_��������_�������(ByVal CurDimensions)
        Case Is = 15
            dimension(tmp) = CurDimensions
            lblSex1.Caption = Select_�����_���������_����(ByVal CurDimensions)
        Case Is = 16
            dimension(tmp) = CurDimensions
            lblSex1.Caption = Select_�������_������_����(ByVal CurDimensions)
        Case Is = 17
            dimension(tmp) = CurDimensions
            lblSex1.Caption = Select_������_������_����(ByVal CurDimensions)
        Case Is = 18
            dimension(tmp) = CurDimensions
            lblSex1.Caption = Select_�������_������_����(ByVal CurDimensions)
        Case Is = 19
            dimension(tmp) = CurDimensions
            lblSex1.Caption = Select_�������_������_����(ByVal CurDimensions)
        Case Is = 20
            dimension(tmp) = CurDimensions
            lblSex1.Caption = Select_������_����(ByVal CurDimensions)
        Case Is = 21
            dimension(tmp) = CurDimensions
            lblSex1.Caption = Select_������_������(ByVal CurDimensions)
        Case Is = 22
            dimension(tmp) = CurDimensions
            lblSex1.Caption = Select_����������_������(ByVal CurDimensions)
        Case Is = 23
            dimension(tmp) = CurDimensions
            lblSex1.Caption = Select_������������_������(ByVal CurDimensions)
        Case Is = 24
            dimension(tmp) = CurDimensions
            lblSex1.Caption = Select_������_������(ByVal CurDimensions)
        End Select
    Call lblVisible
End Sub
'
Private Sub cmdEraseData_Click()
    Call CrDimensions.Main
    Me.Hide
End Sub
'
Private Sub lblVisible()
    lblSex.Visible = True '������� ���� "������������ ���"
    lblSex1.Visible = True '������� ���� "������������ ���"
    cmdOK.Visible = True '������� ������ "��"
End Sub
'
Private Sub lblUnVisible()
    lblSex.Visible = False '������� ���� "������������ ���"
    lblSex1.Visible = False '������� ���� "������������ ���"
    cmdOK.Visible = False '������� ������ "��"
End Sub
'
Private Sub cmdOK_Click()
'�������� ��������� ��� ������� ������ "��"
    Call lblUnVisible
'�������� ������� "����� ������������" � ��������������� ��������
    If N.Count < 24 Then
        ������(3, N.Count) = lblSex1.Caption
        dimension(N.Count) = txtDimensions.Text
        Call Select_CrParameter
        N.increment
        lblMetod.Caption = ������(0, N.Count)
        lblMetod1.Caption = ������(1, N.Count)
        imgCraniumInv.Picture = LoadPicture(������(2, N.Count))  ' �������� �����������
        txtDimensions.Text = "" '"���������"
    Else
        ������(3, N.Count) = lblSex1.Caption
        dimension(N.Count) = txtDimensions.Text
        Call Select_CrParameter
        With Me
'������ ����� ��� ��������� ������������
        .Height = 408
        .Width = 343.5
        .lblMetod.Height = 32 '������ ������� ��� �������� ������������
        .lblMetod.Caption = "���������� ������������"
'���������� �������
        .lblMetod.BackColor = &H80FFFF
        .lblMetod.ForeColor = &HFF&
        .lblMetod.Font.name = "Arial"
        .lblMetod.Font.Size = 18
'������� �������� ��������� �����
        .imgCraniumInv.Visible = False
        Call lblUnVisible
        .lbltxtDimensions.Visible = False
        .txtDimensions.Visible = False
'��������� ������ ���������
        .lblCrForms.Visible = True
            .lblCrForms.Caption = selectCrForm(ByVal getCrForm)
        .lblCrHeight.Visible = True
            .lblCrHeight.Caption = selectCrHeight(ByVal getCrHeight)
        .lblParameter.Visible = True
        .lbl��.Visible = True
            .lbl��.Caption = ��.toString
        .lbl��.Visible = True
            .lbl��.Caption = ��.toString
        .lbl�����.Visible = True
            .lbl�����.Caption = �����.toString
        .lbl��.Visible = True
            .lbl��.Caption = ��.toString
        .lbl��.Visible = True
            .lbl��.Caption = ��.toString
        .lbl���.Visible = True
            .lbl���.Caption = ���.toString
        .cmdToWord.Visible = True
        .cmdEraseData.Visible = True
        .lblMetod1.Caption = getSex
        End With
    End If
End Sub
'
Private Sub cmdToWord_Click()
'������ ������� � �������� �������� Word
    With ActiveDocument
'����������� ������ �������
        Dim TabN As Byte
'        TabN = .Tables.Count + 1 - 4
'��������� �������� �������
Selection.ParagraphFormat.Alignment = wdAlignParagraphRight
    Selection.TypeText Text:= _
        "������� �" & TabN & ". ����������� ���� ����������� �� ������."
    Selection.TypeParagraph
    Selection.ParagraphFormat.Alignment = wdAlignParagraphJustify
    Selection.TypeText Text:="� �/�." & vbTab & "������������ �������� ������" _
         & vbTab & "��������������� ������� "
    Selection.TypeText Text:="(��.)" & vbTab & "���������� (�/�)"
'������ ����� �������
 Dim i As Integer
        For i = 1 To 25
            Selection.TypeParagraph
            Selection.TypeText Text:=i & "." & vbTab & ������(0, (i - 1)) & vbTab & _
                dimension(i - 1) & vbTab & ������(3, (i - 1))
        Next i
'���������� �������
    Selection.TypeParagraph
    Selection.MoveUp Unit:=wdLine, Count:=27, Extend:=wdExtend
    Selection.ConvertToTable Separator:=wdSeparateByTabs, NumColumns:=4, _
        NumRows:=26, AutoFitBehavior:=wdAutoFitContent
    With Selection.Tables(1)
        .Rows(1).Height = CentimetersToPoints(1.1)
        .Rows(2).Height = CentimetersToPoints(0.6)
        .Columns(1).Width = CentimetersToPoints(1.19)
        .Columns(2).Width = CentimetersToPoints(7.5)
        .Columns(3).Width = CentimetersToPoints(3.75)
        .Columns(4).Width = CentimetersToPoints(4.25)
        .Style = "����� �������"
        .ApplyStyleHeadingRows = True
        .ApplyStyleLastRow = False
        .ApplyStyleFirstColumn = True
        .ApplyStyleLastColumn = False
        .Columns(1).Select
            Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
            Selection.Cells.VerticalAlignment = wdCellAlignVerticalCenter
        .Columns(2).Select
            Selection.Cells.VerticalAlignment = wdCellAlignVerticalCenter
        .Columns(3).Select
            Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
            Selection.Cells.VerticalAlignment = wdCellAlignVerticalCenter
        .Columns(4).Select
            Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
            Selection.Cells.VerticalAlignment = wdCellAlignVerticalCenter
    End With
    Selection.MoveDown Unit:=wdLine, Count:=1
    Selection.TypeText Text:=selectCrForm(ByVal getCrForm) & ", " & selectCrHeight(ByVal getCrHeight) & ". " & _
    "���������� ����������� ���������: " & ��.toString & ", " & ��.toString & ", " & �����.toString & ", " _
    & ��.toString & ", " & ��.toString & ", " & ���.toString & ". " & _
    "�������������, " & getSex
    End With
End Sub
'
Private Sub txtDimensions_Change()
'������� ��� ���������� ���������� ���� "���������� �������� � ��"
    With txtDimensions
        If .Text <> "" Then
            Do
                If Not IsNumeric(.Text) Then ' Or Len(.Text) = 0
                    Beep
                    .BackColor = RGB(256, 0, 0)
                    MsgBox "C������ ������� �����", vbCritical, "������ �����"
                    .Text = InputBox("������� ��������� ������ ���������!", "����������� ������ �����")
                Else: CurDimensions = CCur(.Text)
                    .BackColor = RGB(200, 256, 200)
                    If CurDimensions = 0 Then
                        lblSex1.Caption = "��� ���������� ����������"
                        Call lblVisible
                    Else: Select_Dimensions ByVal CurDimensions
                    End If
            Exit Do
                End If
            Loop
        End If
    End With
'    CurDimensions = CCur(.Text)
'        If CurDimensions = 0 Then
'            lblSex1.Caption = "��� ���������� ����������"
'            Call lblVisible
'        Else: Select_Dimensions ByVal CurDimensions
'        End If
End Sub
'
Private Sub UserForm_Initialize()
'��������� ������������� ����:
'���������� �������
    Set �� = New Counter
        ��.name = "���������� �������"
'�������� �������
    Set �� = New Counter
        ��.name = "�������� �������"
'��������������
    Set ����� = New Counter
        �����.name = "��������������"
'�������� �������
    Set �� = New Counter
        ��.name = "�������� �������"
'���������� �������
    Set �� = New Counter
        ��.name = "���������� �������"
'���
    Set ��� = New Counter
        ���.name = "���"
'��������
    Set N = New Counter
        N.name = "Ncounter"
'������������� �������
'�������:
    ������(0, 0) = "���������� �������"
        ������(1, 0) = "glabella (g.) - opistokranion (op.)"
    ������(0, 1) = "���������� �������"
        ������(1, 1) = "euryon (eu.) - euryon (eu.)"
    ������(0, 2) = "�������� �������"
        ������(1, 2) = "basion (ba.) - bregma (ba.)"
    ������(0, 3) = "����� ��������� ������"
        ������(1, 3) = "basion (ba.) - basion (ba.)"
    ������(0, 4) = "���������� ������ ���"
        ������(1, 4) = "fronto-temporale (ft.) - fronto-temporale (ft.)"
    ������(0, 5) = "������ ��������� ������"
        ������(1, 5) = "auriculare (au.) - auriculare (au.)"
    ������(0, 6) = "������ �������"
        ������(1, 6) = "asterion (ast.) - asterion (ast.)"
    ������(0, 7) = "����������� ������"
        ������(1, 7) = "mastoidale (ms.) - mastoidale (ms.)"
    ������(0, 8) = "���������� ������"
        ������(1, 8) = ""
    ������(0, 9) = "������������ �����"
        ������(1, 9) = "nasion (n.) - opistion (o.)"
    ������(0, 10) = "������ �����"
        ������(1, 10) = "nasion (n.) - bregma (b.)"
    ������(0, 11) = "�������� �����"
        ������(1, 11) = "bregma (b.) - lambda (l.)"
    ������(0, 12) = "����� ���"
        ������(1, 12) = "basion (ba.) - opistion (o.)"
    ������(0, 13) = "������ ���"
        ������(1, 13) = ""
    ������(0, 14) = "�������� �������"
        ������(1, 14) = "zygion (zg.) - zygion (zg.)"
    ������(0, 15) = "����� ��������� ����"
        ������(1, 15) = "basion (ba.) - prostion (pr.)"
    ������(0, 16) = "������� ������ ����"
        ������(1, 16) = "nasion (n.) - alveolare (al.)"
    ������(0, 17) = "������ ������ ����"
        ������(1, 17) = "nasion (n.) - gnation (gn.)"
    ������(0, 18) = "������� ������ ����"
        ������(1, 18) = "fronto-malare-temporale (fmt.) - fronto-malare-temporale (fmt.)"
    ������(0, 19) = "������� ������ ����"
        ������(1, 19) = "zygomaxillare (zm.) - zygomaxillare (zm.)"
    ������(0, 20) = "������ ����"
        ������(1, 20) = "nasion (n.) - nasospinale (ns.)"
    ������(0, 21) = "������ ����� ������"
        ������(1, 21) = "maxillofrantale (mf.) - ektokonchion (ek.)"
    ������(0, 22) = "���������� ������"
        ������(1, 22) = ""
    ������(0, 23) = "������������ ������"
        ������(1, 23) = "gonion (go.) - gonijon (go.)"
    ������(0, 24) = "������ ���� ��"
        ������(1, 24) = "gnation (gn.) - infradentale (id.)"
'��������:
    ������(2, 0) = IMGdir & "25.jpg"
    ������(2, 1) = IMGdir & "1.jpg"
    ������(2, 2) = IMGdir & "2.jpg"
    ������(2, 3) = IMGdir & "3.jpg"
    ������(2, 4) = IMGdir & "4.jpg"
    ������(2, 5) = IMGdir & "5.jpg"
    ������(2, 6) = IMGdir & "6.jpg"
    ������(2, 7) = IMGdir & "7.jpg"
    ������(2, 8) = IMGdir & "8.jpg"
    ������(2, 9) = IMGdir & "9.jpg"
    ������(2, 10) = IMGdir & "10.jpg"
    ������(2, 11) = IMGdir & "11.jpg"
    ������(2, 12) = IMGdir & "12.jpg"
    ������(2, 13) = IMGdir & "13.jpg"
    ������(2, 14) = IMGdir & "14.jpg"
    ������(2, 15) = IMGdir & "15.jpg"
    ������(2, 16) = IMGdir & "16.jpg"
    ������(2, 17) = IMGdir & "17.jpg"
    ������(2, 18) = IMGdir & "18.jpg"
    ������(2, 19) = IMGdir & "19.jpg"
    ������(2, 20) = IMGdir & "20.jpg"
    ������(2, 21) = IMGdir & "21.jpg"
    ������(2, 22) = IMGdir & "22.jpg"
    ������(2, 23) = IMGdir & "23.jpg"
    ������(2, 24) = IMGdir & "24.jpg"
'�������� ����� � ��������������� ����������
    lblMetod.Height = 38
    cmdToWord.Visible = False
    imgCraniumInv.BackColor = RGB(0, 164, 157)
'����� ������ �������� ��� �������� �����
    lblMetod.Caption = ������(0, N.Count)
    lblMetod1.Caption = ������(1, N.Count)
    imgCraniumInv.Picture = LoadPicture(������(2, N.Count))
End Sub
'
Private Static Function getCrForm() As Currency
    If dimension(1) > 0 And dimension(0) > 0 Then
        getCrForm = Format(dimension(1) * 100 / dimension(0), 0#)
    Else: getCrForm = 0
    End If
'Debug.Print "���������� ����� =" & getCrForm
End Function
'
Private Static Function selectCrForm(ByVal getCrForm As Currency) As String
'������������� ������� �� �����(���������� �������*100/���������� �������)
    '������� (��������) - ������������ - >80%
    '������� - ����������� - �� 75% �� 79,9%;
    '����� (�������) - ������������� - <75%;
    Dim curTmp As Currency
        curTmp = getCrForm
        Select Case curTmp  '��������� ���������� � �����������
            Case Is = 0
                selectCrForm = "������ ���������������� ����� �� �����"
            Case 1 To 74, 9
                selectCrForm = "����������� ����� ������������� ����� " & "(" & curTmp & ")" '��������, ������������ ��������
            Case 75 To 79.9
                selectCrForm = "����������� ����� ����������� ����� " & "(" & curTmp & ")"
            Case Is > 80
                selectCrForm = "����������� ����� ������������ ����� " & "(" & curTmp & ")"
        End Select
End Function
'
Private Static Function getCrHeight() As Currency
    If dimension(2) > 0 And dimension(0) > 0 Then
        getCrHeight = dimension(2) * 100 / dimension(0)
    Else: getCrHeight = 0
    End If
'Debug.Print "�������� ���������� = " & getCrHeight
End Function
'
Private Static Function selectCrHeight(ByVal getCrHeight As Currency) As String
    Dim curTmp As Currency
        curTmp = getCrHeight
    Select Case curTmp  '��������� ���������� � �����������
        Case Is = 0
            selectCrHeight = "������ ���������������� ����� �� ������"
        Case 1 To 69, 9
            selectCrHeight = "������ (�����������) " & "(" & curTmp & ")"
        Case 70 To 74, 9
            selectCrHeight = "������������� (�����������) " & "(" & curTmp & ")"
        Case Is > 75
            selectCrHeight = "������� (������������) " & "(" & curTmp & ")"
    End Select
'Debug.Print "�� ������ -" & selectCrHeight
End Function
'
Private Static Function getSex() As String
    If ��.Count > 12 Or ��.Count > ��.Count Then
        getSex = "����������� ����� ����������� ����������� �������� ����."
    ElseIf ��.Count > 12 Or ��.Count > ��.Count Then
        getSex = "����������� ����� ����������� ����������� �������� ����."
    Else
        If (��.Count + ��.Count) > (��.Count + ��.Count) Then
            getSex = "����������� ����� ��� ������������ ����������� �������� ����."
        ElseIf (��.Count + ��.Count) > (��.Count + ��.Count) Then
            getSex = "����������� ����� ��� ������������ ����������� �������� ����."
        Else: getSex = "���������� ��� ������������ ������ �� �������������� ���������."
        End If
    End If
End Function
'1
Private Static Function Select_����������_�������(ByVal CurDimensions As Currency) As String
'��������� ���������� "���������� �������: ��������-�������������" � ���������� �������
    If CurDimensions <= 0 Then '���������� �������� ��������
        Select_����������_������� = "��� ���������� ����������"
    Else
        Select Case CurDimensions '������������ ����������, � ��� ����������� ���������
            Case Is <= 160
                Select_����������_������� = "���������� �������"
            Case 160.1 To 172
                Select_����������_������� = "�������� �������"
            Case 172.1 To 178.5
                Select_����������_������� = "��������������"
            Case 178.6 To 187
                Select_����������_������� = "�������� �������"
            Case Is >= 187.1
                Select_����������_������� = "���������� �������"
            Case Else: MsgBox "������! ��������� �������� �� �������� ������ " _
                & "��� ������� �� ������� ���������� ��������!" & Err.Description
        End Select
    End If
End Function
'2
Private Static Function Select_����������_�������(ByVal CurDimensions As Currency) As String
'��������� ���������� "����������_�������" � ���������� �������
    If CurDimensions <= 0 Then '���������� �������� ��������
        Select_����������_������� = "��� ���������� ����������"
    Else
        Select Case CurDimensions '������������ ����������, � ��� ����������� ���������
            Case Is <= 127
                Select_����������_������� = "���������� �������"
            Case 127.1 To 138
                Select_����������_������� = "�������� �������"
            Case 138.1 To 143
                Select_����������_������� = "��������������"
            Case 143.1 To 152
                Select_����������_������� = "�������� �������"
            Case Is >= 152.1
                Select_����������_������� = "���������� �������"
            Case Else: MsgBox "������! ��������� �������� �� �������� ������ " _
                & "��� ������� �� ������� ���������� ��������!" & Err.Description
        End Select
    End If
End Function
'3
Private Static Function Select_��������_�������(ByVal CurDimensions As Currency) As String
'��������� ���������� "��������_�������" � ���������� �������
    If CurDimensions <= 0 Then '���������� �������� ��������
        Select_��������_������� = "��� ���������� ����������"
    Else
        Select Case CurDimensions '������������ ����������, � ��� ����������� ���������
            Case Is <= 121
                Select_��������_������� = "���������� �������"
            Case 121.1 To 128
                Select_��������_������� = "�������� �������"
            Case 128.1 To 134
                Select_��������_������� = "��������������"
            Case 134.1 To 140.5
                Select_��������_������� = "�������� �������"
            Case Is >= 140.6
                Select_��������_������� = "���������� �������"
           Case Else: MsgBox "������! ��������� �������� �� �������� ������ " _
                & "��� ������� �� ������� ���������� ��������!" & Err.Description
        End Select
    End If
End Function
'4
Private Static Function Select_�����_���������(ByVal CurDimensions As Currency) As String
'��������� ���������� "�����_���������" � ���������� �������
    If CurDimensions <= 0 Then '���������� �������� ��������
        Select_�����_��������� = "��� ���������� ����������"
    Else
        Select Case CurDimensions '������������ ����������, � ��� ����������� ���������
            Case Is <= 90
                Select_�����_��������� = "���������� �������"
            Case 90.1 To 96
                Select_�����_��������� = "�������� �������"
            Case 96.1 To 101
                Select_�����_��������� = "��������������"
            Case 101.1 To 109
                Select_�����_��������� = "�������� �������"
            Case Is >= 109.1
                Select_�����_��������� = "���������� �������"
            Case Else: MsgBox "������! ��������� �������� �� �������� ������ " _
                & "��� ������� �� ������� ���������� ��������!" & Err.Description
        End Select
    End If
End Function
'5
Private Static Function Select_Min_������_���(ByVal CurDimensions As Currency) As String
'��������� ���������� "Min_������_���" � ���������� �������
    If CurDimensions <= 0 Then '���������� �������� ��������
        Select_Min_������_��� = "��� ���������� ����������"
    Else
        Select Case CurDimensions '������������ ����������, � ��� ����������� ���������
            Case Is <= 86
                Select_Min_������_��� = "���������� �������"
            Case 86.1 To 95
                Select_Min_������_��� = "�������� �������"
            Case 95.1 To 98
                Select_Min_������_��� = "��������������"
            Case 98.1 To 108
                Select_Min_������_��� = "�������� �������"
            Case Is >= 108.1
                Select_Min_������_��� = "���������� �������"
           Case Else: MsgBox "������! ��������� �������� �� �������� ������ " _
                & "��� ������� �� ������� ���������� ��������!" & Err.Description
        End Select
    End If
End Function
'6
Private Static Function Select_������_���������(ByVal CurDimensions As Currency) As String
'��������� ���������� "������_���������" � ���������� �������
    If CurDimensions <= 0 Then '���������� �������� ��������
        Select_������_��������� = "��� ���������� ����������"
    Else
        Select Case CurDimensions '������������ ����������, � ��� ����������� ���������
            Case Is <= 112
                Select_������_��������� = "���������� �������"
            Case 112.1 To 117
                Select_������_��������� = "�������� �������"
            Case 117.1 To 123
                Select_������_��������� = "��������������"
            Case 123.1 To 133
                Select_������_��������� = "�������� �������"
            Case Is >= 133.1
                Select_������_��������� = "���������� �������"
            Case Else: MsgBox "������! ��������� �������� �� �������� ������ " _
                & "��� ������� �� ������� ���������� ��������!" & Err.Description
        End Select
    End If
End Function
'7
Private Static Function Select_������_�������(ByVal CurDimensions As Currency) As String
'��������� ���������� "������_�������" � ���������� �������
    If CurDimensions <= 0 Then '���������� �������� ��������
        Select_������_������� = "��� ���������� ����������"
    Else
        Select Case CurDimensions '������������ ����������, � ��� ����������� ���������
            Case Is <= 99
                Select_������_������� = "���������� �������"
            Case 99.1 To 106.9
                Select_������_������� = "�������� �������"
            Case 107 To 110.4
                Select_������_������� = "��������������"
            Case 110.5 To 120
                Select_������_������� = "�������� �������"
            Case Is >= 120.1
                Select_������_������� = "���������� �������"
            Case Else: MsgBox "������! ��������� �������� �� �������� ������ " _
                & "��� ������� �� ������� ���������� ��������!" & Err.Description
        End Select
    End If
End Function
'8
Private Static Function Select_�����������_������(ByVal CurDimensions As Currency) As String
'��������� ���������� "�����������_������" � ���������� �������
    If CurDimensions <= 0 Then '���������� �������� ��������
        Select_�����������_������ = "��� ���������� ����������"
    Else
        Select Case CurDimensions '������������ ����������, � ��� ����������� ���������
            Case Is <= 92
                Select_�����������_������ = "���������� �������"
            Case 92.1 To 100
                Select_�����������_������ = "�������� �������"
            Case 100.1 To 105
                Select_�����������_������ = "��������������"
            Case 105.1 To 116
                Select_�����������_������ = "�������� �������"
            Case Is >= 116.1
                Select_�����������_������ = "���������� �������"
            Case Else: MsgBox "������! ��������� �������� �� �������� ������ " _
                & "��� ������� �� ������� ���������� ��������!" & Err.Description
        End Select
    End If
End Function
'9
Private Static Function Select_����������_������(ByVal CurDimensions As Currency) As String
'��������� ���������� "����������_������" � ���������� �������
    If CurDimensions <= 0 Then '���������� �������� ��������
        Select_����������_������ = "��� ���������� ����������"
    Else
        Select Case CurDimensions '������������ ����������, � ��� ����������� ���������
            Case Is <= 476
                Select_����������_������ = "���������� �������"
            Case 476.1 To 500.5
                Select_����������_������ = "�������� �������"
            Case 500.6 To 516.5
                Select_����������_������ = "��������������"
            Case 516.6 To 540
                Select_����������_������ = "�������� �������"
            Case Is >= 540.1
                Select_����������_������ = "���������� �������"
           Case Else: MsgBox "������! ��������� �������� �� �������� ������ " _
                & "��� ������� �� ������� ���������� ��������!" & Err.Description
        End Select
    End If
End Function
'10
Private Static Function Select_������������_�����(ByVal CurDimensions As Currency) As String
'��������� ���������� "������������_�����" � ���������� �������
    If CurDimensions <= 0 Then '���������� �������� ��������
        Select_������������_����� = "��� ���������� ����������"
    Else
        Select Case CurDimensions '������������ ����������, � ��� ����������� ���������
            Case Is <= 123
                Select_������������_����� = "���������� �������"
            Case 123.1 To 128.5
                Select_������������_����� = "�������� �������"
            Case 128.6 To 134.5
                Select_������������_����� = "��������������"
            Case 134.6 To 145
                Select_������������_����� = "�������� �������"
            Case Is >= 145.1
                Select_������������_����� = "���������� �������"
            Case Else: MsgBox "������! ��������� �������� �� �������� ������ " _
                & "��� ������� �� ������� ���������� ��������!" & Err.Description
        End Select
    End If
End Function
'11
Private Static Function Select_������_�����(ByVal CurDimensions As Currency) As String
'��������� ���������� "������_�����" � ���������� �������
    If CurDimensions <= 0 Then '���������� �������� ��������
        Select_������_����� = "��� ���������� ����������"
    Else
        Select Case CurDimensions  '������������ ����������, � ��� ����������� ���������
            Case Is <= 99
                Select_������_����� = "���������� �������"
            Case 99.1 To 107
                Select_������_����� = "�������� �������"
            Case 107.1 To 111.5
                Select_������_����� = "��������������"
            Case 111.6 To 121
                Select_������_����� = "�������� �������"
            Case Is >= 121.1
                Select_������_����� = "���������� �������"
            Case Else: MsgBox "������! ��������� �������� �� �������� ������ " _
                & "��� ������� �� ������� ���������� ��������!" & Err.Description
        End Select
    End If
End Function
'12
Private Static Function Select_��������_�����(ByVal CurDimensions As Currency) As String
'��������� ���������� " ��������_�����" � ���������� �������
    If CurDimensions <= 0 Then '���������� �������� ��������
        Select_��������_����� = "��� ���������� ����������"
    Else
        Select Case CurDimensions  '������������ ����������, � ��� ����������� ���������
            Case Is <= 94
                Select_��������_����� = "���������� �������"
            Case 94.1 To 107
                Select_��������_����� = "�������� �������"
            Case 107.1 To 110.5
                Select_��������_����� = "��������������"
            Case 110.6 To 124
                Select_��������_����� = "�������� �������"
            Case Is >= 124.1
                Select_��������_����� = "���������� �������"
            Case Else: MsgBox "������! ��������� �������� �� �������� ������ " _
                & "��� ������� �� ������� ���������� ��������!" & Err.Description
        End Select
    End If
End Function
'13
Private Static Function Select_�����_���(ByVal CurDimensions As Currency) As String
'��������� ���������� "�����_���" � ���������� �������
    If CurDimensions <= 0 Then '���������� �������� ��������
        Select_�����_��� = "��� ���������� ����������"
    Else
        Select Case CurDimensions '������������ ����������, � ��� ����������� ���������
            Case Is <= 30
                Select_�����_��� = "���������� �������"
            Case 30.1 To 34
                Select_�����_��� = "�������� �������"
            Case 34.1 To 36
                Select_�����_��� = "��������������"
            Case 36.1 To 41
                Select_�����_��� = "�������� �������"
            Case Is >= 41.1
                Select_�����_��� = "���������� �������"
        Case Else: MsgBox "������! ��������� �������� �� �������� ������ " _
                & "��� ������� �� ������� ���������� ��������!" & Err.Description
        End Select
    End If
End Function
'14
Private Static Function Select_������_���(ByVal CurDimensions As Currency) As String
'��������� ���������� "������_���" � ���������� �������
    If CurDimensions <= 0 Then '���������� �������� ��������
        Select_������_��� = "��� ���������� ����������"
    Else
        Select Case CurDimensions '������������ ����������, � ��� ����������� ���������
            Case Is <= 25
                Select_������_��� = "���������� �������"
            Case 25.1 To 28.5
                Select_������_��� = "�������� �������"
            Case 28.6 To 30.5
                Select_������_��� = "��������������"
            Case 30.6 To 35
                Select_������_��� = "�������� �������"
            Case Is >= 35.1
                Select_������_��� = "���������� �������"
            Case Else: MsgBox "������! ��������� �������� �� �������� ������ " _
                & "��� ������� �� ������� ���������� ��������!" & Err.Description
        End Select
    End If
End Function
'15
Private Static Function Select_��������_�������(ByVal CurDimensions As Currency) As String
'��������� ���������� "��������_�������" � ���������� �������
    If CurDimensions <= 0 Then '���������� �������� ��������
        Select_��������_������� = "��� ���������� ����������"
    Else
        Select Case CurDimensions '������������ ����������, � ��� ����������� ���������
            Case Is <= 120
                Select_��������_������� = "���������� �������"
            Case 120.1 To 124
                Select_��������_������� = "�������� �������"
            Case 124.1 To 132
                Select_��������_������� = "��������������"
            Case 132.1 To 139
                Select_��������_������� = "�������� �������"
            Case Is >= 139.1
                Select_��������_������� = "���������� �������"
           Case Else: MsgBox "������! ��������� �������� �� �������� ������ " _
                & "��� ������� �� ������� ���������� ��������!" & Err.Description
        End Select
    End If
End Function
'16
Private Static Function Select_�����_���������_����(ByVal CurDimensions As Currency) As String
'��������� ���������� "�����_���������_����" � ���������� �������
    If CurDimensions <= 0 Then '���������� �������� ��������
        Select_�����_���������_���� = "��� ���������� ����������"
    Else
        Select Case CurDimensions '������������ ����������, � ��� ����������� ���������
            Case Is <= 82
                Select_�����_���������_���� = "���������� �������"
            Case 82.1 To 93
                Select_�����_���������_���� = "�������� �������"
            Case 93.1 To 97.5
                Select_�����_���������_���� = "��������������"
            Case 97.6 To 107
                Select_�����_���������_���� = "�������� �������"
            Case Is >= 107.1
                Select_�����_���������_���� = "���������� �������"
            Case Else: MsgBox "������! ��������� �������� �� �������� ������ " _
                & "��� ������� �� ������� ���������� ��������!" & Err.Description
        End Select
    End If
End Function
'17
Private Static Function Select_�������_������_����(ByVal CurDimensions As Currency) As String
'��������� ���������� "�������_������_����" � ���������� �������
    If CurDimensions <= 0 Then '���������� �������� ��������
        Select_�������_������_���� = "��� ���������� ����������"
    Else
        Select Case CurDimensions '������������ ����������, � ��� ����������� ���������
            Case Is <= 59
                Select_�������_������_���� = "���������� �������"
            Case 59.1 To 66.5
                Select_�������_������_���� = "�������� �������"
            Case 66.6 To 71
                Select_�������_������_���� = "��������������"
            Case 71.1 To 78
                Select_�������_������_���� = "�������� �������"
            Case Is >= 78.1
                Select_�������_������_���� = "���������� �������"
            Case Else: MsgBox "������! ��������� �������� �� �������� ������ " _
                & "��� ������� �� ������� ���������� ��������!" & Err.Description
        End Select
    End If
End Function
'18
Private Static Function Select_������_������_����(ByVal CurDimensions As Currency) As String
'��������� ���������� "������_������_����" � ���������� �������
    If CurDimensions <= 0 Then '���������� �������� ��������
        Select_������_������_���� = "��� ���������� ����������"
    Else
        Select Case CurDimensions '������������ ����������, � ��� ����������� ���������
            Case Is <= 100
                Select_������_������_���� = "���������� �������"
            Case 100.1 To 111
                Select_������_������_���� = "�������� �������"
            Case 111.1 To 119
                Select_������_������_���� = "��������������"
            Case 119.1 To 132
                Select_������_������_���� = "�������� �������"
            Case Is >= 132.1
                Select_������_������_���� = "���������� �������"
            Case Else: MsgBox "������! ��������� �������� �� �������� ������ " _
                & "��� ������� �� ������� ���������� ��������!" & Err.Description
        End Select
    End If
End Function
'19
Private Static Function Select_�������_������_����(ByVal CurDimensions As Currency) As String
'��������� ���������� "�������_������_����" � ���������� �������
    If CurDimensions <= 0 Then '���������� �������� ��������
        Select_�������_������_���� = "��� ���������� ����������"
    Else
        Select Case CurDimensions '������������ ����������, � ��� ����������� ���������
            Case Is <= 93
                Select_�������_������_���� = "���������� �������"
            Case 93.1 To 101
                Select_�������_������_���� = "�������� �������"
            Case 101.1 To 105
                Select_�������_������_���� = "��������������"
            Case 105.1 To 113
                Select_�������_������_���� = "�������� �������"
            Case Is >= 113.1
                Select_�������_������_���� = "���������� �������"
            Case Else: MsgBox "������! ��������� �������� �� �������� ������ " _
                & "��� ������� �� ������� ���������� ��������!" & Err.Description
        End Select
    End If
End Function
'20
Private Static Function Select_�������_������_����(ByVal CurDimensions As Currency) As String
'��������� ���������� "�������_������_����"� ���������� �������
    If CurDimensions <= 0 Then '���������� �������� ��������
        Select_�������_������_���� = "��� ���������� ����������"
    Else
        Select Case CurDimensions '������������ ����������, � ��� ����������� ���������
            Case Is <= 78
                Select_�������_������_���� = "���������� �������"
            Case 78.1 To 89
                Select_�������_������_���� = "�������� �������"
            Case 89.1 To 93.5
                Select_�������_������_���� = "��������������"
            Case 93.6 To 104
                Select_�������_������_���� = "�������� �������"
            Case Is >= 104.1
                Select_�������_������_���� = "���������� �������"
           Case Else: MsgBox "������! ��������� �������� �� �������� ������ " _
                & "��� ������� �� ������� ���������� ��������!" & Err.Description
        End Select
    End If
End Function
'21
Private Static Function Select_������_����(ByVal CurDimensions As Currency) As String
'��������� ���������� "������_����"� ���������� �������
    If CurDimensions <= 0 Then '���������� �������� ��������
        Select_������_���� = "��� ���������� ����������"
    Else
        Select Case CurDimensions '������������ ����������, � ��� ����������� ���������
            Case Is <= 44
                Select_������_���� = "���������� �������"
            Case 44.1 To 48.5
                Select_������_���� = "�������� �������"
            Case 48.6 To 52
                Select_������_���� = "��������������"
            Case 52.1 To 56
                Select_������_���� = "�������� �������"
            Case Is >= 56.1
                Select_������_���� = "���������� �������"
            Case Else: MsgBox "������! ��������� �������� �� �������� ������ " _
                & "��� ������� �� ������� ���������� ��������!" & Err.Description
        End Select
    End If
End Function
'22
Private Static Function Select_������_������(ByVal CurDimensions As Currency) As String
'��������� ���������� "������_������" � ���������� �������
    If CurDimensions <= 0 Then '���������� �������� ��������
        Select_������_������ = "��� ���������� ����������"
    Else
        Select Case CurDimensions '������������ ����������, � ��� ����������� ���������
            Case Is <= 38
                Select_������_������ = "���������� �������"
            Case 38.1 To 42
                Select_������_������ = "�������� �������"
            Case 42.1 To 43.5
                Select_������_������ = "��������������"
            Case 43.6 To 48
                Select_������_������ = "�������� �������"
            Case Is >= 48.1
                Select_������_������ = "���������� �������"
            Case Else: MsgBox "������! ��������� �������� �� �������� ������ " _
                & "��� ������� �� ������� ���������� ��������!" & Err.Description
        End Select
    End If
End Function
'23
Private Static Function Select_����������_������(ByVal CurDimensions As Currency) As String
'��������� ���������� "����������_������"� ���������� �������
    If CurDimensions <= 0 Then '���������� �������� ��������
        Select_����������_������ = "��� ���������� ����������"
    Else
        Select Case CurDimensions '������������ ����������, � ��� ����������� ���������
            Case Is <= 105
            Select_����������_������ = "���������� �������"
            Case 105.1 To 113.5
                Select_����������_������ = "�������� �������"
            Case 113.6 To 118.5
                Select_����������_������ = "��������������"
            Case 118.6 To 127
                Select_����������_������ = "�������� �������"
            Case Is >= 127.1
                Select_����������_������ = "���������� �������"
           Case Else: MsgBox "������! ��������� �������� �� �������� ������ " _
                & "��� ������� �� ������� ���������� ��������!" & Err.Description
        End Select
    End If
End Function
'24
Private Static Function Select_������������_������(ByVal CurDimensions As Currency) As String
'��������� ���������� "������������_������"� ���������� �������
    If CurDimensions <= 0 Then '���������� �������� ��������
        Select_������������_������ = "��� ���������� ����������"
    Else
        Select Case CurDimensions '������������ ����������, � ��� ����������� ���������
            Case Is <= 85
                Select_������������_������ = "���������� �������"
            Case 85.1 To 95
                Select_������������_������ = "�������� �������"
            Case 95.1 To 102.5
                Select_������������_������ = "��������������"
            Case 102.6 To 112
                Select_������������_������ = "�������� �������"
            Case Is >= 112.1
                Select_������������_������ = "���������� �������"
            Case Else: MsgBox "������! ��������� �������� �� �������� ������ " _
                & "��� ������� �� ������� ���������� ��������!" & Err.Description
        End Select
    End If
End Function
'25
Private Static Function Select_������_������(ByVal CurDimensions As Currency) As String
'��������� ���������� "������_������" � ���������� �������
    If CurDimensions <= 0 Then '���������� �������� ��������
        Select_������_������ = "��� ���������� ����������"
    Else
        Select Case CurDimensions '������������ ����������, � ��� ����������� ���������
            Case Is <= 27
                Select_������_������ = "���������� �������"
            Case 27.1 To 31
                Select_������_������ = "�������� �������"
            Case 31.1 To 33.5
                Select_������_������ = "��������������"
            Case 33.6 To 41
                Select_������_������ = "�������� �������"
            Case Is >= 41.1
                Select_������_������ = "���������� �������"
            Case Else: MsgBox "������! ��������� �������� �� �������� ������ " _
                & "��� ������� �� ������� ���������� ��������!" & Err.Description
        End Select
    End If
End Function
'
Private Static Sub Select_CrParameter()
    Select Case ������(3, N.Count) '������������ ����������, � ��� ����������� ���������
        Case Is = "���������� �������"
            ��.increment
        Case Is = "�������� �������"
            ��.increment
        Case Is = "��������������"
            �����.increment
        Case Is = "�������� �������"
            ��.increment
        Case Is = "���������� �������"
            ��.increment
        Case Else: ���.increment
    End Select
End Sub
'
Private Sub UserForm_Terminate()
'��������� ������������� ����:
'���������� �������
    Set �� = Nothing
    Set �� = Nothing
    Set ����� = Nothing
    Set �� = Nothing
    Set �� = Nothing
    Set ��� = Nothing
    Set N = Nothing
    Unload Me
End Sub
