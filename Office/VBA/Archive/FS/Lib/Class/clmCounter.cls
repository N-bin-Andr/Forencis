VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clmCounter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "�������"
'����� "Counter"
'��� �������� �������� "�������"
'���� ��������: 18.10.2017
'@version 0.0.001
'@author Andr.N@_bin
'
Option Explicit
Private tmpid As Long       '������������� (���������� ��������� ��������)
Private tmpmeCount As Long  '��������� ��������
Private tmpname As String   '��� ��������
Public tmpstep As Integer		'шаг счетчика.
'
Private Sub Class_Initialize()
    meCount = 0
'    id = id + 1
End Sub
'
Friend Function increment()
'���������� �������� ��������
    meCount = meCount + 1
End Function
'
Friend Function reset()
'����� �������� �� ����
    meCount = 0
End Function
'
Friend Function getTale() As Long
'����������� �������� �������� ��������
    getTale = meCount
'Debug.Print "������� �������� ��������- " & getTale
End Function
'
Friend Function toStirng(Optional str As String = "") As String
'��������� ������������� ��������
If str = "" Then
    Dim tmp As Long, tmp1 As Long
    tmp = getTale
        Select Case tmp
            Case Is = 0
                str = "��������"
            Case Is = 1
                str = "������"
            Case 2 To 4
                str = "�������"
            Case 5 To 20
                str = "��������"
            Case Else
            tmp1 = Right(tmp, 1)
                Select Case tmp1
                    Case Is = 0
                        str = "��������"
                    Case Is = 1
                        str = "������"
                    Case 2 To 4
                        str = "�������"
                    Case 5 To 9
                        str = "��������"
                    End Select
            End Select
   End If
toStirng = name & " (" & meCount & Chr(32) & str & ")"
Debug.Print "��������� ������������� ��������- " & toStirng
End Function
'
'������������
'Let+++++++++++++++++++++++++++++++++++++++
Private Property Let id(ByVal vData As Long)
'��������� ��������
    tmpid = vData
End Property
'
Friend Property Let name(ByVal vData As String)
'��� ��������
    If Len(vData) = 0 Then
        tmpname = "��� �������� �� ���������� "
    Else: tmpname = vData
    End If
End Property
'
Private Property Let meCount(ByVal vData As Long)
'��������� ��������
    tmpmeCount = vData
End Property

Friend Property Let step (ByVal vData As Integer)
'счетчик
    tmpstep = vData
End Property
'
'Get+++++++++++++++++++++++++++++++++++
Friend Property Get id() As Long
'��������� ��������
    id = tmpid
'Debug.Print "���������� ���������- " & id
End Property
'
Friend Property Get name() As String
'��� ��������
    name = tmpname
'Debug.Print "��� ��������- " & name
End Property
'
Private Property Get meCount() As Long
'��������� ��������
    meCount = tmpmeCount
'Debug.Print "��������� ��������- " & meCount
End Property

Private Property Get step() As Integer
'��������� ��������
    step = tmpstep
'Debug.Print "шаг счетчика = " & step
End Property
'Get========================================
