Attribute VB_Name = "mdDocCreate"
Option Explicit

'����� ��� �������� ����� � ������ ����������.
'@author Andr.N@_bin
'@E-mail Andr.Nab.n@gmail.com
'
'���������� ��������� ����������
Private fs As FileSystemObject  '���������  FileSystemObject
Private SvcService As Object    '������ ���������� svcsvc.dll
'Public ourStamp As String       '���� ������ ��� �������� ��������
'Word
Public MyWdApp As Word.Application  '��������� ����������
Public DocWord As Word.Document     ' ��������� ���������
'������
Const CRITICAL As String = "��������, ������!"

'+++++++++++++++++++++++++ � � � � � � ++++++++++++++++++++++++++++

Public Function Print_OurStamp(lngEvCount As Long) As String
'������="������������/�� ��������������/� ���������/�, ���������/� ��������� �������� ������ �����
'           ������� ������ + "���� ������"
    Dim str As String
    If lngEvCount > 0 Then
        If lngEvCount = 1 Then
            str = "������������ �������������� ���������, ��������� "
        Else: str = "������������ �������������� ���������, ��������� "
        End If
    End If
Dim strTx As String
    strTx = "��������� �������� ������ ����� � ������ ������ "
Print_OurStamp = str & strTx & Chr(171) & STAMP & Chr(187)
'OurStamp = "Print_OurStamp " & Print_OurStamp
End Function
'
Public Function deliveryEvToString(lngEvCount As Long) As String
'������� ������ �������� ��������/��
'������ = "������������/�� ��������������/� ����������/� ��������, �����������/�� � + �������� + ������"
    Dim str As String
    If lngEvCount > 0 Then
        If lngEvCount = 1 Then
            str = "������������ �������������� ���������� ��������, ����������� � "
        Else: str = "������������ �������������� ���������� ��������, ����������� � "
        End If
    End If
    
    Dim strTx1 As String
    If StrComp(strEvPackage, "��������� �������") = True Then
        strTx1 = "����������� "
    Else: strTx1 = "����������� "
    Debug.Print "strTx1 = "
    End If
deliveryEvToString = str & strTx1 & strEvPackage & "(" & "���� ��" & ")."
Debug.Print "deliveryEvToString = " & deliveryEvToString
End Function
'
Public Function boxIntegralityToString(lngEvCount As Long) As String
'����������� ��������
    Dim str As String, strTx1 As String
    If lngEvCount > 0 Then
        If lngEvCount = 1 Then
            str = "���������������� ������� "
            strTx1 = "��� �������� "
            
        Else
            str = "��������������� �������� "
            strTx1 = "���� ���������: "
        End If
    End If
boxIntegralityToString = "����������� �������� �� ��������, ����������" & str & _
                        "��� ����������� ����������� �������� ����������. ��� �������� �������� �� ��� " & strTx1
Debug.Print "boxIntegralityToString = " & boxIntegralityToString
End Function
'
Public Function accordanceEvObjToString(lngEvCount As Long) As String
'������= "������/�, ���������������/�� ��� ������-������������������� ������������, �������������/�� "
'       ��������, ���������� � ����������� � � ���������������� ������� � ������������ ���������������. "
 Dim str As String, strTx1 As String
    If lngEvCount > 0 Then
        If lngEvCount = 1 Then
            str = "������, ��������������� �� ������������, "
            strTx1 = "������������� "
            
        Else
            str = "�������, ��������������� �� ������������, "
            strTx1 = "������������ "
        End If
    End If
accordanceEvObjToString = str & strTx1 & "��������, ���������� � ����������� � � ���������������� ������� � ������������ ���������������." & Chr(10)
Debug.Print "accordanceEvObjToString = " & accordanceEvObjToString
End Function
'
'Public Property Get ourStamp() As String
'
'
'
'End Property


Public Static Sub printEF(nameDOT As String, nameDOC As String, tmpDir As String)
'�������� � ���������� ���������� ��������
 Set SvcService = CreateObject("Svcsvc.Service") '������ ����������  svcsvc.dll
 Set MyWdApp = New Word.Application
    With MyWdApp
        .Visible = True  'False ��������  ��������
        '1)�������� ������� �� ����� � ���������
        If nameDOT = "" Then
           MsgBox "����� c ��������� ���������� �� �������!", vbExclamation, CRITICAL
           nameDOT = SvcService.SelectFolder("������� ����� � ��������� ����������!", "", &H10 + &H4000, "")
        End If
'       2)�������� �������
        Set DocWord = .Documents.Open(nameDOT) '��������� ������"
'       ���������� ���
        DocWord.Activate
'       ���������� ������ ��������� ��� ����������� ������
        With DocWord  '�������� �������� ��������
'       ������ � ������ ����������:
'       �������� � ���������
'       �������� �� ������������� ��������
        If ActiveDocument.Bookmarks.Exists("temp") = True Then
            ActiveDocument.Bookmarks("temp").Select
'       frmEvidences.colEvidences(index)

            
        End If

'
'
'



'       �������� - lb(lable)
'                .FormFields("lb_InjPrIniName").Result = frmNewResearch.newInjPr.create_InitialslName
'                .FormFields("lb_fNumberEF").Result = frmNewResearch.newEF.getFullNumber 'fNumberEF
'                .FormFields("lb_EFFirstDay").Result = frmNewResearch.newEF.firstDay 'firstDate
'                .FormFields("lb_AutopsyDate").Result = frmNewResearch.newInjPr.autopsyDate 'autopsyDate
'                .FormFields("lb_expertName").Result = frmNewResearch.newExpert 'expertName
'       ������ - fm (Form)
'                .FormFields("fm_InjPrFullName").Result = frmNewResearch.newInjPr.create_FullName 'injPrFullName
'                .FormFields("fm_InjPrBirthday").Result = frmNewResearch.newInjPr.birthday 'injPrBirthday
'                .FormFields("fm_InjPrSex").Result = frmNewResearch.newInjPr.sex 'injPrSex
'                .FormFields("fm_deceaseDate").Result = frmNewResearch.newInjPr.decease 'deceaseDate
'                .FormFields("fm_autopsyDate").Result = frmNewResearch.newInjPr.autopsyDate 'autopsy
'                .FormFields("fm_fNumberEF").Result = frmNewResearch.newEF.getFullNumber 'fNumberEF
'                .FormFields("fm_factCase").Result = frmNewResearch.txtFactCase.Text 'factCase
'                .FormFields("fm_legalGround").Result = frmNewResearch.create_legalGround 'legalGround
'                .FormFields("fm_expertName").Result = reverse_Name(frmNewResearch.newExpert)
'      ���������������� �������  nt (note)
'                .FormFields("nt_EFFirstDay").Result = frmNewResearch.newEF.firstDay 'firstDate
'                .FormFields("nt_fNumberEF").Result = frmNewResearch.newEF.getFullNumber 'fNumberEF
'                .FormFields("nt_autopsy").Result = frmNewResearch.newInjPr.print_Autopsy 'autopsy info
'                .FormFields("nt_stringEF").Result = frmNewResearch.newEF.toStringEF 'stringEF
'                .FormFields("nt_expertName").Result = reverse_Name(frmNewResearch.newExpert)

'            ����������:
'            .SaveAs FileName:=tmpdir & nameDOC, FileFormat:= _
''                wdFormatXMLDocumentMacroEnabled, LockComments:=False, Password:="", _
''                AddToRecentFiles:=True, WritePassword:="", ReadOnlyRecommended:=False, _
''                EmbedTrueTypeFonts:=False, SaveNativePictureFormat:=False, SaveFormsData _
''                :=False, SaveAsAOCELetter:=False
'            '��������
            .Close
        End With
    End With
Set SvcService = Nothing    '������ ����������  svcsvc.dll
'���������� ������ - ��������
Set DocWord = Nothing
'���������� ������ ����������
MyWdApp.Quit
Set MyWdApp = Nothing
End Sub
































'Public Static Sub print_Blank(nameDOT As String, nameDOC As String, tmpdir As String)
''���������� ������� �����������
' Set SvcService = CreateObject("Svcsvc.Service") '������ ����������  svcsvc.dll
'' Set MyWdApp = New Word.Application
'    With MyWdApp
'        .Visible = False 'True '��������  ��������
'        '1)�������� ������� �� ����� � ���������
'        If DOTDIR = "" Then
'           MsgBox "����� c ��������� ���������� �� �������!", vbExclamation, CRITICAL
'           DOTDIR = SvcService.SelectFolder("������� ����� � ��������� ����������!", "", &H10 + &H4000, "")
'        End If
'        '2)�������� �������
'        Dim tmpDot As String
'            tmpDot = DOTDIR & nameDOT & ".dotm" '������������ ����� ������� (� �����������)
'        '��������� ������ ���������
'        Set DocWord = .Documents.Open(tmpDot)  '��������� ������"
'        '���������� ���
'        DocWord.Activate
'        '���������� ������ ��������� ��� ����������� ������
'        With DocWord  '�������� �������� ��������
'            '������ � ������ ����������:
'            If nameDOT = "�������" Then
'                .FormFields("WdsNumberEF").Result = frmNewResearch.newEF.number    'sNumberEF
'                .FormFields("WdyearEF").Result = Right(frmNewResearch.newInjPr.autopsyDate, 4)  'Right(autopsyDate, 4)
'                '������ ��������� ���������
'                If MsgBox("���������� " & nameDOC & " ?", vbYesNo) = vbYes Then
'                    .PrintOut Range:=wdPrintAllDocument, item:= _
'                        wdPrintDocumentContent, Copies:=1, Pages:="", PageType:=wdPrintAllPages, _
'                        ManualDuplexPrint:=False, Collate:=False, Background:=True, PrintToFile:= _
'                        False, PrintZoomColumn:=0, PrintZoomRow:=0, PrintZoomPaperWidth:=0, _
'                        PrintZoomPaperHeight:=0
'
'                End If
'            ElseIf nameDOT = "����� ������" Or _
'                                nameDOT = "����� �����" Or _
'                                nameDOT = "�������� �����" Or _
'                                nameDOT = "�����������" Or _
'                                nameDOT = "���� ���������� ���" Then
'                .FormFields("WdstringEF").Result = frmNewResearch.newEF.toStringEF
'                .FormFields("WdexpertName").Result = reverse_Name(frmNewResearch.newExpert)
'                 '������ ��������� ���������
'                If MsgBox("���������� " & nameDOC & " ?", vbYesNo) = vbYes Then
'                    .PrintOut Range:=wdPrintAllDocument, item:= _
'                        wdPrintDocumentContent, Copies:=1, Pages:="", PageType:=wdPrintAllPages, _
'                        ManualDuplexPrint:=False, Collate:=False, Background:=True, PrintToFile:= _
'                        False, PrintZoomColumn:=0, PrintZoomRow:=0, PrintZoomPaperWidth:=0, _
'                        PrintZoomPaperHeight:=0
'                End If
'            ElseIf nameDOT = "��������" Then
'                .FormFields("WdexpertName").Result = frmNewResearch.newExpert 'expertName
'                .FormFields("WdsNumberEF").Result = frmNewResearch.newEF.getNumber 'sNumberEF
'                .FormFields("Wdautopsy").Result = frmNewResearch.newInjPr.print_Autopsy 'autopsy
'                 '������ ��������� ���������
'                If MsgBox("���������� " & nameDOC & " ?", vbYesNo) = vbYes Then
'                    .PrintOut Range:=wdPrintAllDocument, item:= _
'                        wdPrintDocumentContent, Copies:=1, Pages:="", PageType:=wdPrintAllPages, _
'                        ManualDuplexPrint:=True, Collate:=False, Background:=True, PrintToFile:= _
'                        False, PrintZoomColumn:=0, PrintZoomRow:=0, PrintZoomPaperWidth:=0, _
'                        PrintZoomPaperHeight:=0
'                End If
'            Else
'                '�������� - lb(lable)
'                .FormFields("lb_InjPrIniName").Result = frmNewResearch.newInjPr.create_InitialslName
'                .FormFields("lb_fNumberEF").Result = frmNewResearch.newEF.getFullNumber 'fNumberEF
'                .FormFields("lb_EFFirstDay").Result = frmNewResearch.newEF.firstDay 'firstDate
'                .FormFields("lb_AutopsyDate").Result = frmNewResearch.newInjPr.autopsyDate 'autopsyDate
'                .FormFields("lb_expertName").Result = frmNewResearch.newExpert 'expertName
'                '������ - fm (Form)
'                .FormFields("fm_InjPrFullName").Result = frmNewResearch.newInjPr.create_FullName 'injPrFullName
'                .FormFields("fm_InjPrBirthday").Result = frmNewResearch.newInjPr.birthday 'injPrBirthday
'                .FormFields("fm_InjPrSex").Result = frmNewResearch.newInjPr.sex 'injPrSex
'                .FormFields("fm_deceaseDate").Result = frmNewResearch.newInjPr.decease 'deceaseDate
'                .FormFields("fm_autopsyDate").Result = frmNewResearch.newInjPr.autopsyDate 'autopsy
'                .FormFields("fm_fNumberEF").Result = frmNewResearch.newEF.getFullNumber 'fNumberEF
'                .FormFields("fm_factCase").Result = frmNewResearch.txtFactCase.Text 'factCase
'                .FormFields("fm_legalGround").Result = frmNewResearch.create_legalGround 'legalGround
'                .FormFields("fm_expertName").Result = reverse_Name(frmNewResearch.newExpert)
'                '���������������� �������  nt (note)
'                .FormFields("nt_EFFirstDay").Result = frmNewResearch.newEF.firstDay 'firstDate
'                .FormFields("nt_fNumberEF").Result = frmNewResearch.newEF.getFullNumber 'fNumberEF
'                .FormFields("nt_autopsy").Result = frmNewResearch.newInjPr.print_Autopsy 'autopsy info
'                .FormFields("nt_stringEF").Result = frmNewResearch.newEF.toStringEF 'stringEF
'                .FormFields("nt_expertName").Result = reverse_Name(frmNewResearch.newExpert)
''                ������ ��������� ���������
'                If MsgBox("���������� " & nameDOC & " ?", vbYesNo) = vbYes Then
'                    If nameDOT = "�� ����������" Then
'                        .PrintOut Range:=wdPrintRangeOfPages, item:= _
'                            wdPrintDocumentContent, Copies:=1, Pages:="1-2", PageType:= _
'                            wdPrintAllPages, ManualDuplexPrint:=True, Collate:=False, Background:= _
'                            True, PrintToFile:=False, PrintZoomColumn:=0, PrintZoomRow:=0, _
'                            PrintZoomPaperWidth:=0, PrintZoomPaperHeight:=0
'                        .PrintOut Range:=wdPrintRangeOfPages, item:= _
'                            wdPrintDocumentContent, Copies:=2, Pages:="3", PageType:=wdPrintAllPages, _
'                            ManualDuplexPrint:=False, Collate:=False, Background:=True, PrintToFile _
'                            :=False, PrintZoomColumn:=0, PrintZoomRow:=0, PrintZoomPaperWidth:=0, _
'                            PrintZoomPaperHeight:=0
'                        .PrintOut Range:=wdPrintRangeOfPages, item:= _
'                            wdPrintDocumentContent, Copies:=1, Pages:="4", PageType:=wdPrintAllPages, _
'                            ManualDuplexPrint:=False, Collate:=False, Background:=True, PrintToFile _
'                            :=False, PrintZoomColumn:=0, PrintZoomRow:=0, PrintZoomPaperWidth:=0, _
'                            PrintZoomPaperHeight:=0
'                    Else
'                        .PrintOut Range:=wdPrintRangeOfPages, item:= _
'                            wdPrintDocumentContent, Copies:=1, Pages:="1-2", PageType:= _
'                            wdPrintAllPages, ManualDuplexPrint:=True, Collate:=False, Background:= _
'                            True, PrintToFile:=False, PrintZoomColumn:=0, PrintZoomRow:=0, _
'                            PrintZoomPaperWidth:=0, PrintZoomPaperHeight:=0
'                         .PrintOut Range:=wdPrintRangeOfPages, item:= _
'                            wdPrintDocumentContent, Copies:=2, Pages:="3", PageType:=wdPrintAllPages, _
'                            ManualDuplexPrint:=False, Collate:=False, Background:=True, PrintToFile _
'                            :=False, PrintZoomColumn:=0, PrintZoomRow:=0, PrintZoomPaperWidth:=0, _
'                            PrintZoomPaperHeight:=0
'                    End If
'                End If
'            End If
'            '����������:
'            .SaveAs FileName:=tmpdir & nameDOC, FileFormat:= _
'                wdFormatXMLDocumentMacroEnabled, LockComments:=False, Password:="", _
'                AddToRecentFiles:=True, WritePassword:="", ReadOnlyRecommended:=False, _
'                EmbedTrueTypeFonts:=False, SaveNativePictureFormat:=False, SaveFormsData _
'                :=False, SaveAsAOCELetter:=False
'            '��������
'            .Close
'        End With
'    End With
'Set SvcService = Nothing    '������ ����������  svcsvc.dll
''���������� ������ - ��������
'Set DocWord = Nothing
''���������� ������ ����������
''MyWdApp.Quit
''Set MyWdApp = Nothing
'End Sub
