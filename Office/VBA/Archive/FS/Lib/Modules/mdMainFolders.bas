Attribute VB_Name = "mdMainFolders"
'mdMainFolders -> ��������� ������ � ������������
'@author Andr.Nab.n@gmail.com
'Lib:
'newEF.getNumber = ����_YY
'newEF.getNum_Cat = ����_YY_�����
Option Explicit
Private fs As FileSystemObject  '���������  FileSystemObject
Private SvcService As Object    '������ ���������� svcsvc.dll
Private Const MAIN_FOLDER_NOT_EXIST = "����� ��� ���������� ����� ���������� �� �������!"
Public Const USER_DOC_DIR As String = "D:\Crime\Soft_��������\VBA\1_6\Resources\user\usrDocDir.txt" '���� ������, ���������� ���������� ��� ����� ����������, ��������� �������������
Public Const USER_DOT_DIR As String = "D:\Crime\Soft_��������\VBA\1_6\Resources\user\usrDotDir.txt"
Private tmpdirDOT As String '���������� �������� ����������
Private tmpdirDOC As String '���������� ��� ���������� ������� ����������

Public Const DOC_INDEX As Integer = 0
Public Const DOT_INDEX As Integer = 1
Private fileDescr As Integer
'������ �������� �����
Public arrDocDir(0 To 6) As String
'arrDocDir (0) = �������� ���������� ��� ����� ����������(��������� �������������) ��:"D:\Crime\"
'arrDocDir (1) = �������� ���������� � ��������� ����������� ����������(��������� �������������) ��:"D:\Crime\DOT\"

'arrDocDir (1) = 0 + \YYYY  (D:\Crime\2020)
'arrDocDir (2) = 1 + \newEF.getNum_Cat  (D:\Crime\YYYY\����_YY_�����)
'arrDocDir (3) = 2+ \����_" & criateDocName(����)
'arrDocDir (4) = 2+ \��������_" &
'arrDocDir (5) = 2+"\�����_"  &
'arrDocDir (6) = 2+"\��������_"  &
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'Private Property Let dirDOT(ByVal vData As String)
''���������� �������� ����������
'    tmpdirDOT = vData
'End Property
''
'Public Property Get dirDOT(tmpRoot As String) As String
''���������� �������� ����������
'   fileDescr = FreeFile
'    Open tmpRoot For Input As #fileDescr
'        While Not EOF(1)
'            Input #1, str
'            Debug.Print "������ �� ����� = " & tmpdirDOT
'        Wend
'    Close #fileDescr
'    dirDOT = tmpdirDOT
'Debug.Print "dirDOT = " & dirDOT
'End Property

Public Sub getMainDir(tmpRoot As String, i As Integer)
'������� ��������� ���� �� ���������� �����.
'String tmpRoot -> ���������� ���� � ������: usrDocDir.txt/usrDotDir.txt ����������: ���������� ��� ����� ����������/�������� ����������, ��������� �������������
'Integer i -> ������ ������� ��� ������ ���������� ��� ����� ����������(0)/�������� ���������� (1),
'1) ��������� ���� �� ���������� �����.
    Dim str As String
    fileDescr = FreeFile
    Open tmpRoot For Input As #fileDescr
        While Not EOF(1)
            Input #1, str
            Debug.Print "������ �� ����� = " & str
        Wend
    Close #fileDescr
'   �������� �� ������ ��������
    If str <> "" Then
        arrDocDir(i) = str
    Else
        Call inputUserDir(i)
    End If

'    Dim tmpRoot As String
'    fileDescr = FreeFile
'    Open USER_DOC_DIR For Input As #fileDescr
'        While Not EOF(1)
'            Input #1, tmpRoot
'            Debug.Print "������ �� ����� = " & tmpRoot
'        Wend
'    Close #fileDescr
''   �������� �� ������ ��������
'    If tmpRoot <> "" Then
'        arrDocDir(0) = tmpRoot
'    Else
'        Call inputUserDir
'    End If
End Sub
'
Public Sub inputUserDir(i As Integer)
'����� ���������� ��� ����� ����������
fileDescr = FreeFile
Set SvcService = CreateObject("Svcsvc.Service") '������ ����������  svcsvc.dll
Set fs = New FileSystemObject                   '���������  FileSystemObject

Dim str1 As String, str2 As String
    If i = 0 Then
        str1 = "������� ����� ��� ���������� ����� ����������!"
        str2 = USER_DOC_DIR
    ElseIf i = 1 Then
        str1 = "������� �����,  ���������� ������� ����������� ����������!"
        str2 = USER_DOT_DIR
    End If
    MsgBox str1, vbExclamation, "����� ����� �����"
       arrDocDir(i) = SvcService.SelectFolder("����� �����", "", &H10 + &H4000, "")
'           ������ ���� � ��������� ���� ���������:
            Open str2 For Output As #fileDescr
                Print #fileDescr, arrDocDir(i)    '������ ������ ����;
                    Debug.Print "newUserDir = " & arrDocDir(i)
                    MsgBox "������� ����� " & arrDocDir(i) & "!", vbOKOnly, "����� �����"
            Close #fileDescr
        Set SvcService = Nothing    '����������� ������� ����������  svcsvc.dll
        Set fs = Nothing            '����������� ����������  FileSystemObject
'������ ������ +++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'fileDescr = FreeFile
'Set SvcService = CreateObject("Svcsvc.Service") '������ ����������  svcsvc.dll
'Set fs = New FileSystemObject                   '���������  FileSystemObject
'    MsgBox "������� ����� ��� ���������� ����� ����������!", vbExclamation, "����� ����� �����"
'       arrDocDir(0) = SvcService.SelectFolder("�������� ����� ��� ���������� ����������", "", &H10 + &H4000, "")
''           ������ ���� � ��������� ���� ���������:
'            Open USER_DOC_DIR For Output As #fileDescr
'                Print #fileDescr, arrDocDir(0)    '������ ������ ����;
'                    Debug.Print "newUserDir = " & arrDocDir(0)
'                    MsgBox "������� ����� " & arrDocDir(0) & "!", vbOKOnly, "����� �����"
'            Close #fileDescr
'        Set SvcService = Nothing    '����������� ������� ����������  svcsvc.dll
'        Set fs = Nothing            '����������� ����������  FileSystemObject
'������ ������ +++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
End Sub
'
Public Sub makeDocDir(Optional fn As String = "", _
                        Optional fnCat As String = "�����")
'fn - ���������� ����� ���������_��� = newEF.getNum_Cat
'fnCat - ��������� ���������� �� ����������� �������� = newEF.getNum_Cat
'1) ����� ����������:
    ChDrive "D"
'   ���������� ������ �� ����� � ��������� ���� (������������ �������������)
    Call getMainDir(USER_DOC_DIR, 0)
'2)�������� ������� ����������
    If arrDocDir(0) = "" Then
    MsgBox MAIN_FOLDER_NOT_EXIST, vbExclamation, "������ ����������!"
        Call inputUserDir(0)
    End If
'�������� ����� � ������� ����������
    If fn <> "" Then
        arrDocDir(2) = arrDocDir(0) & "\" & fn & "_" & Right(CStr(Year(Now)), 2) & "_" & fnCat  '(D:\Crime\YYYY\����_YY_�����)
        arrDocDir(3) = arrDocDir(2) & "\����_" & fn & "_" & Right(CStr(Year(Now)), 2)           '(D:\Crime\YYYY\����_YY_�����\����_����_YY)
        arrDocDir(4) = arrDocDir(2) & "\��������_" & fn & "_" & Right(CStr(Year(Now)), 2)       '(D:\Crime\YYYY\����_YY_�����\��������_����_YY)
        arrDocDir(5) = arrDocDir(2) & "\�����_" & fn & "_" & Right(CStr(Year(Now)), 2)          '(D:\Crime\YYYY\����_YY_�����\�����_����_YY)
        arrDocDir(6) = arrDocDir(2) & "\��������_" & fn & "_" & Right(CStr(Year(Now)), 2)       '(D:\Crime\YYYY\����_YY_�����\��������_����_YY)
    End If
    
    Dim i As Integer
        For i = 2 To UBound(arrDocDir)
            MkDir arrDocDir(i)
        Next i
�������:
    Dim x As Integer
     For x = LBound(arrDocDir) To UBound(arrDocDir)
        Debug.Print "�������� ������� " & x & " = " & arrDocDir(x) & Chr(10)
    Next x
End Sub
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

