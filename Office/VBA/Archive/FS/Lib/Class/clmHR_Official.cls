VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clmHR_Official"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "Class HR_Official - �����, ��������������� ��� ������ � ����������� � ������ ������ �� �������� ����������. �������� ���� � ������, ����������� ��������� ������� ������, ��������, ��� ""�����������"" �/��� ""�������""."
'����� Hight rank Officials (�����������/�������)
'@author Andr.N@_bin
'@E-mail Andr.Nab.n@gmail.com
Option Explicit
Private tmpid As Long               '������������� (���������� ��������� ��������)
Const MY_EXPERIENCE As Date = #1/8/2005#
'���� (����������) ������
Private tmpsurname As String        '�������
Private tmpname As String           '���
Private tmpmidName As String        '��������
Private tmpsex As String            '���
Private tmplegalStatus              '����������� ������ (�����������/�������/������� � �.�.)
Private tmppost As String           '���������
Private tmprank As String           '������
Private tmpconformation As String   '��������� �������������
Private tmpoffice As String         '���������� �������������
Private tmpdepartment As String     '����� �������������
Private tmpPCR As String            '��������� + ��������� + ������
Private tmpexperience As String     '���� ������ �� �������������
Private tmppi As String             '������������ ������
Private tmpcomments As String       '����������
'
Friend Function create_Rank() As String
'������ = ��������� + ���������� + ������
    If post = "" Then
        create_Rank = conformation & Chr(32) & rank & "(��������� �� �������)"
    ElseIf conformation = "" Then
        create_Rank = post & Chr(32) & rank & "(���������� �� �������)"
    ElseIf rank = "" Then
        create_Rank = post & Chr(32) & conformation & "(������ �� �������)"
    Else: create_Rank = post & Chr(32) & conformation & Chr(32) & rank
    End If
Debug.Print "��������� + ���������� + ������= " & create_Rank
End Function
'
Friend Function print_Cor() As String
'������ = ��������� + ��������� + ������ "�����������"
        PCR = create_Rank '��������� + ��������� + ������
    If PCR = "" Then
       print_Cor = "����������� (��������� � ������ �� �������)  " & create_InitialslName
    Else:
        print_Cor = PCR & Chr(32) & create_InitialslName
    End If
Debug.Print "print_Cor- " & print_Cor
End Function
'
'Friend Function Create_CERank() As String
''������� �������� ������ ��������� � ������ ��������
''   �������� �� �������
'    If post = "" Then
'        Create_CERank = conformation & Chr(32) & rank & "(��������� �� �������)"
'    ElseIf conformation = "" Then
'        create_Rank = post & Chr(32) & rank & "(���������� �� �������)"
'    ElseIf rank = "" Then
'        create_Rank = post & Chr(32) & conformation & "(������ �� �������)"
'    Else: create_Rank = post & Chr(32) & conformation & Chr(32) & rank
'    End If
'Debug.Print "��������� + ���������� + ������= " & create_Rank
'
'End Function
'
Friend Function print_Expert() As String
'������� ������ ������ ����������� "������ ��������"
'    �������� �� �������:
       print_Expert = " ���������������� ������������ ��������� �������� " & create_InitialslName
End Function
'
'Friend Function print_Exp() As String
''������ = ��������� + ��������� + ������ "��������"
'    print_Exp = "���������������� ������������ ��������� �������� " & create_InitialslName
'Debug.Print "print_Exp- " & print_Exp
'End Function
'
Public Function create_Initials(initials As String) As String
'������� �������� �������� � ������
    If initials <> "" Then
        Dim tmp As String
            tmp = Left(initials, 1) & "."
        create_Initials = StrConv(tmp, vbProperCase)
    End If
Debug.Print "������� � ������-" & tmp
End Function
'
Friend Function create_InitialslName() As String
'������� = ������� + �������� ����� � �������� (� ��������� ����� � �������)
Dim tmp1 As String, tmp2 As String, tmp3 As String
'   �������
    If surname = "" Then
        tmp1 = "(������� �� �������)"
    Else: tmp1 = surname
    End If
'   ���:
    If name = "" Then
        tmp2 = "(��� �� �������) "
    Else: tmp2 = create_Initials(name)
    End If
'   ��������
    If midName = "" Then
        tmp3 = "(�������� �� �������)"
    Else: tmp3 = create_Initials(midName)
    End If
create_InitialslName = tmp1 & Chr(32) & tmp2 & tmp3
Debug.Print "clmHR_Official create_InitialslName- " & create_InitialslName
End Function
'
Friend Function create_ReverseInitialslName() As String
'������� = �������� ����� � �������� (� ��������� ����� � �������) + �������
Dim tmp1 As String, tmp2 As String, tmp3 As String
'   �������
    If surname = "" Then
        tmp1 = "(������� �� �������)"
    Else: tmp1 = surname
    End If
'   ���:
    If name = "" Then
        tmp2 = "(��� �� �������)"
    Else: tmp2 = create_Initials(name)
    End If
'   ��������:
    If midName = "" Then
        tmp3 = "(�������� �� �������)"
    Else: tmp3 = create_Initials(midName)
    End If
create_ReverseInitialslName = tmp2 & tmp3 & Chr(32) & tmp1
Debug.Print "create_ReverseInitialslName- " & create_ReverseInitialslName
End Function
'
Friend Function create_FullName() As String
'������� �������� �������, ������ ����� � ��������
Dim str As String
    str = name
        If Len(str) > 2 Then
            str = name & Chr(32)
        End If
create_FullName = surname & Chr(32) & str & midName
Debug.Print "create_FullName- " & create_FullName
End Function
'
Friend Function create_ReverseFullName() As String
'������� �������� ������ �����, �������� � �������
Dim str As String
    str = name
        If Len(str) > 2 Then
            str = name & Chr(32)
        End If
create_ReverseFullName = str & midName & Chr(32) & surname
Debug.Print "create_ReverseFullName- " & create_ReverseFullName
End Function
'
Friend Function create_surName(tmp As String) As String
'������� �������� ������� ��� ���������
        If tmp <> "" Then
            create_surName = Left(tmp, Len(tmp) - 5)
        End If
Debug.Print "create_surName- " & create_surName
End Function
'
Friend Function create_revExpert(fio As String) As String
'������ ������� ��������� � ������� (�� ������� �.�. =  �.�.������� )
    Dim tmpsurname As String, tmpIO As String
        tmpIO = Right(fio, 4)
        tmpsurname = Left(fio, Len(fio) - 5)
    create_revExpert = tmpIO & Chr(32) & tmpsurname
End Function

'������������
'Let_____________________________________________
Private Property Let id(ByVal vData As Long)
'��������� ��������
    tmpid = vData
End Property
'
Friend Property Let surname(ByVal vData As String)
Attribute surname.VB_Description = "���� ""�������"""
'�������
tmpsurname = vData
'    If Len(vData) = 0 Then
'        tmpsurName = "(������� �� �������)"
'    Else: tmpsurName = StrConv(vData, vbProperCase)
'    End If
End Property
'
Friend Property Let name(ByVal vData As String)
'���
tmpname = vData
'    If Len(vData) = 0 Then
'        tmpname = "(��� �� �������)"
'    Else
'        If Len(vData) = 1 Then
'            Item = vData
'            tmpname = create_Initials(Item)
'        Else: tmpname = StrConv(vData, vbProperCase)
'        End If
'    End If
End Property
'
Friend Property Let midName(ByVal vData As String)
Attribute midName.VB_Description = "���� ""��������"""
'��������
tmpmidName = vData
'    If Len(vData) = 0 Then
'        tmpmidName = "(�������� �� �������)"
'    Else
'        If Len(vData) = 1 Then
'            Item = vData
'            tmpmidName = create_Initials(Item)
'        Else: tmpmidName = StrConv(vData, vbProperCase)
'        End If
'    End If
End Property
'
Friend Property Let sex(ByVal vData As String)
Attribute sex.VB_Description = "���� ""���"" (�������/�������)"
'���
    If Len(vData) = 0 Then
        tmpsex = "��� ������"
    Else: tmpsex = vData
    End If
End Property
'
Friend Property Let legalStatus(ByVal vData As String)
'����������� ������
    If Len(vData) = 0 Then
        tmplegalStatus = ""
    Else: tmplegalStatus = vData
    End If
End Property
'
Friend Property Let experience(ByVal vData As String)
'���� ������ �� �������������
    If Len(vData) = 0 Then
        tmpexperience = "��� ������"
    Else: tmpexperience = vData
    End If
End Property
'
Friend Property Let post(ByVal vData As String)
Attribute post.VB_Description = "���� ""���������"""
'���������
    If Len(vData) = 0 Then
        tmppost = "��� ������"
    Else: tmppost = vData
    End If
End Property
'
Friend Property Let rank(ByVal vData As String)
Attribute rank.VB_Description = "���� ""������"""
'������
    If Len(vData) = 0 Then
        tmprank = "��� ������"
    Else: tmprank = vData
    End If
End Property
'
Friend Property Let conformation(ByVal vData As String)
Attribute conformation.VB_Description = "���� ""��������� �������������"""
'��������� �������������
    If Len(vData) = 0 Then
        tmpconformation = "��� ������"
    Else: tmpconformation = vData
    End If
End Property
'
Friend Property Let office(ByVal vData As String)
Attribute office.VB_Description = "���� ""���������� (�����)  �������������"""
'���������� �������������
    If Len(vData) = 0 Then
        tmpoffice = "��� ������"
    Else: tmpoffice = vData
    End If
End Property
'
Friend Property Let department(ByVal vData As String)
Attribute department.VB_Description = "���� ""����� �������������"""
'����� �������������
    If Len(vData) = 0 Then
        tmpdepartment = "��� ������"
    Else: tmpdepartment = vData
    End If
End Property
'
Friend Property Let PCR(ByVal vData As String)
Attribute PCR.VB_Description = "���� ""����� ������������"""
'��������� + ��������� + ������
    If Len(vData) = 0 Then
        tmpPCR = "��� ������"
    Else: tmpPCR = vData
    End If
End Property
'
Friend Property Let comments(ByVal vData As String)
'����������
    If Len(vData) = 0 Then
        tmpcomments = "��� ������"
    Else: tmpcomments = vData
    End If
End Property
'
Friend Property Let pi(ByVal vData As String)
'������������ ������
    If Len(vData) = 0 Then
        tmppi = "��� ������"
    Else: tmppi = vData
    End If
End Property
'
'Get========================================
Friend Property Get id() As Long
'��������� ��������
    id = tmpid
'Debug.Print "���������� ���������- " & id
End Property
'
Friend Property Get surname() As String
'�������.
    surname = tmpsurname
'Debug.Print "�������- " & surName
End Property
'
Friend Property Get name() As String
'���.
    name = tmpname
'Debug.Print "���- " & name
End Property
'
Friend Property Get midName() As String
'��������.
    midName = tmpmidName
'Debug.Print "��������- " & midName
End Property
'
Friend Property Get sex() As String
'���.
    sex = tmpsex
'Debug.Print "���- " & sex
End Property
'
Friend Property Get legalStatus() As String
'����������� ������
    legalStatus = tmplegalStatus
'Debug.Print "legalStatus- " & legalStatus
End Property
'
Friend Property Get post() As String
'���������.
    post = tmppost
'Debug.Print "���������- " & post
End Property
'
Friend Property Get rank() As String
'������.
    rank = tmprank
'Debug.Print "������- " & rank
End Property
'
Friend Property Get conformation() As String
'��������� �������������.
    conformation = tmpconformation
'Debug.Print "��������� �������������- " & conformation
End Property
'
Friend Property Get office() As String
'���������� �������������
    office = tmpoffice
'Debug.Print "���������� �������������- " & office
End Property
'
Friend Property Get department() As String
'����� �������������.
    department = tmpdepartment
'Debug.Print "����� �������������- " & department
End Property
'
Friend Property Get PCR() As String
'��������� + ��������� + ������
    PCR = tmpPCR
'Debug.Print "��������� + ��������� + ������- " & PCR
End Property
'
Friend Property Get experience() As String
Attribute experience.VB_Description = "���� ""���� ������"""
'���� ������ �� �������������
    experience = tmpexperience
'Debug.Print "���� ������ �� �������������: " & experience
End Property
'
Friend Property Get comments() As String
Attribute comments.VB_Description = "���� ""����������"""
'����������.
    comments = tmpcomments
'Debug.Print "����������: " & comments
End Property
'
Friend Property Get pi() As String
Attribute pi.VB_Description = "���� ""������������ ������"""
'������������ ������.
    pi = tmppi
'Debug.Print "������������ ������: " & pi
End Property
'
Private Sub Class_Initialize()
'�����������
    id = id + 1
'������������� ���������� "���� ������ �� �������������" ����� �������
    experience = MY_EXPERIENCE
End Sub
'
'+++++++++++++++++++++++++++++ L I B +++++++++++++++++++
'Private Static Function reverse_Name(expertName As String) As String
''������ = �������� + �������
''1)���������� ����� ������:
'    Dim i As Integer
'        i = Len(expertName)
''2)�������� ��������:
'    Dim str1 As String
'        str1 = Right(expertName, 4)
''3)�������� ������ = �������� + �������
'    reverse_Name = str1 & Chr(32) & Left(expertName, (i - 5))
'Debug.Print "reverse_Name: " & reverse_Name
'End Function

