Attribute VB_Name = "mdCount"
'������ "mdCount" ��� ������ �� ����������
'�������� ���������� � ������� ��� ������ �� ����������
'NB!!! �������� � ������� clmCounter
'�������� �������� � ��������� ������, �.�. ���������� � ������ �
'������� ����� �������� � ������?
'���� ��������: 01.06.2016
'@version 0.0.1
'@author Andr.Nab.n@gmail.com
Option Explicit
'���������� ���������
Public boxCounter As clmCounter    '��������� �������� � ��
Public boxEvSumCounter As clmCounter  '������� "����� ��" � ��������
Public allEvSumCounter As clmCounter  '������� "����� ����� ��"
'Public colForms As Collection       '��������� ����


'Public counterEF As clmCounter       '������� ���� "���������� ��������"
'Public counterEvCat As clmCounter    '������� "��������� �� (� ����������)"
'Public mvarfEvidCount As clmCounter      '������� ���� "������������ ��������������"
'Public mvarfDocListCount As clmCounter   '������� ���� "������ ����������"
'
''���������� ���������
'Private mvarfEFCount As Long        '������� ���� "���������� ��������"
'Private mvarfEvCat_Count As Long    '������� "��������� �� (� ����������)"
'Private mvarfEvidCount As Long      '������� ���� "������������ ��������������"
'Private mvarfDocListCount As Long   '������� ���� "������ ����������"
'Private mvarEvGRCount As Integer    '��������� ����� ��
'Private mvarEvSumCount As Long      '������� "����� ����� ��"
''�������� ��������� ����������
''Private mvarDocCategory As String   '��������� ���������
''Private mvarstrDOC As String        '�
''Private mvarstrDOT As String        '������ ���������
'
'Public Property Let fEvidCount(ByVal vData As Integer)
''������� ���� "������������ ��������������"
'mvarfEvidCount = vData
'End Property
''
'Public Property Get fEvidCount() As Integer
''������� ���� "������������ ��������������"
'fEvidCount = mvarfEvidCount
'Debug.Print "������� ������������ �������������� = ", fEvidCount
'End Property
''
'Public Property Let EvGRCount(ByVal vData As Integer)
''������� ����� ��
'    mvarEvGRCount = vData
'End Property
''
'Public Property Get EvGRCount() As Integer
''������� ����� ��
'    EvGRCount = mvarEvGRCount
'Debug.Print "'������� ����� ��= ", EvGRCount
'End Property
''
'Public Property Let fEFCount(ByVal vData As Long)
''������� ���� "���������� ��������".
'    mvarfEFCount = vData
'End Property
''
'Public Property Get fEFCount() As Long
''������� ���� "���������� ��������"
'    fEFCount = mvarfEFCount
''Debug.Print "������� ���� "���������� ��������" = ", EvSumCount
'End Property
''
'Public Property Let fEvCat_Count(ByVal vData As Long)
''"��������� �� (� ����������)"
'    mvarfEvCat_Count = vData
'End Property
''
'Public Property Get fEvCat_Count() As Long
''������� ���� "��������� ����������".
'    fEvCat_Count = mvarfEvCat_Count
'Debug.Print "������� ���� ���� �������� = ", fEvCat_Count
'End Property
''
'Public Property Let EvSumCount(ByVal vData As Long)
''������� "����� ����� ��".
'    mvarEvSumCount = vData
'End Property
''
'Public Property Get EvSumCount() As Long
''������� "����� ����� ��".
'    EvSumCount = mvarEvSumCount
''Debug.Print "������� "����� ����� ��" = ", EvSumCount
'End Property
''
'Public Property Let fDocListCount(ByVal vData As Long)
''������� ���� frmDocList
'mvarfDocListCount = vData
'End Property
''
'Public Property Get fDocListCount() As Long
''������� ���� frmDocList
'fDocListCount = mvarfDocListCount
'End Property

'Public Property Let DocCategory(ByVal vData As String)
''���������� ��������� ����������.
'    mvarDocCategory = vData
'End Property
''
'Public Property Get DocCategory() As String
''���������� ��������� ����������.
'    DocCategory = mvarDocCategory
''Debug.Print "���������� ��������� ���������� = ", DocCategory
'End Property
''
'Public Property Let strDOC(ByVal vData As String)
''���������� ������ ���������.
'    mvarstrDOC = vData
'End Property
''
'Public Property Get strDOC() As String
''���������� ������ ���������.
'    strDOC = mvarstrDOC
''Debug.Print "������ ��������� = ", strDOC
'End Property

