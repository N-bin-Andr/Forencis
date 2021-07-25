Attribute VB_Name = "mdPrintDoc"
'������ "mdPrint" ��� ������������ ��������� ������ � ������ �� � ���������
'���� ��������: 01.06.2016
'@version 0.0.1
'@author Andr.Nab.n@gmail.com
Option Explicit
'!!�������� �� ��������� ������ �� ����� �������
Const DOT = "D:\Crime\Soft_��������\VBA\1_6\Resources\doc" '��� ����������" '���������� �������� (�� ���������)
Const doc = "D:" '���������� ������� ���������� (�� ���������).

Const CRITICAL As String = "��������, ������!"
Private Const BOX As String = "BX"  '������� ����� ��� �������
Private Const EVID As String = "EV" '������� ����� ��� ��������.
'
Public tmpBx As clmEvBox
Private bxKey  As Variant
'
Public colDocCat As Collection '��������� "��������� ���������"
Public colTextDoc As Collection ' ��������� ����������� �������,����������� � ���������
Public colExcelData As Collection '��������� ������ ��� ������ � ����� Excel
Public strMonth As String            '������� �����
'
Private mvarDocCat As String   '���������� "��������� ���������"
Private mvarstrDOT As String   '��� ������� ���������
'
Private mvarstrDOC As String   '��� ���������� ���������
'������������ �������:
Public arrEF() As String             '������ �������� ��������� (�������� � ��������)
Public arrExpert() As String         '������ ���� ���������
'
'������ ���������� ���������� "MsOffice"
Public MyWdApp As New Word.Application      '��������� ���������� MsOffice Word
Public MyWdDoc As Word.Document            '��������� ��������� MsOffice Word
Public myExApp As Excel.Application        '��������� ���������� MsOffice Excel
Public myExDoc As Excel.Workbook           '��������� ��������� MsOffice Excel
'Public MyOlApp As Outlook.Application      '��������� ���������� MsOffice Outlook
'Public MyOlTask As Outlook.TaskItem        '��������� ������ MsOffice Outlook
Private fs As New FileSystemObject         '���������� ������ ������� (�����)
Private fll As Variant
Private SvcService As Object    '������ ���������� svcsvc.dll
Public monthNow As String       '������� ����� (������.../�������)��� ������ � Excel
'
'========================== I N C A P S U L A T I O N ==================================
'
Public Property Let DocCat(ByVal vData As String)
'���������� ��������� ����������.
    mvarDocCat = vData
End Property
'
Public Property Get DocCat() As String
'���������� ��������� ����������.
    DocCat = mvarDocCat
'Debug.Print "���������� ��������� ���������� = ", DocCategory
End Property
'
Public Property Let strDOC(ByVal vData As String)
'���������� ������ ���������.
    mvarstrDOC = vData
End Property
'
Public Property Get strDOC() As String
'���������� ������ ���������.
    strDOC = mvarstrDOC
'Debug.Print "������ ��������� = ", strDOC
End Property
'
Public Property Let strDOT(ByVal vData As String)
'���������� ������ ���������.
    mvarstrDOT = vData
End Property
'
Public Property Get strDOT() As String
'���������� ������ ���������.
    strDOT = mvarstrDOT
'Debug.Print "������ ��������� = ", strDOC
End Property
'
Public Function puckEVup(tmpStamp As String, lngEvCount As Long) As String
'������ - "������������/��  ��������������/�� ���������/��, ���������/��" + ���� ������
   Dim str1 As String, str2 As String, str3 As String
    If lngEvCount > 0 Then
        If lngEvCount = 1 Then
                str1 = "������������ �������������� "
            str2 = "���������, "
            str3 = "��������� "
        Else:
        str1 = "������������ �������������� "
         str2 = "���������, "
            str3 = "��������� "
        End If
puckEVup = str1 & str2 & str3 & "��������� �������� ������ ����� ������� ������ " & tmpStamp
    End If
End Function
'
'====================================
Public Static Sub withApplWD(dirDOT As String, dirDOC As String)
'��������� ��� ������ � ����������� msWord
'String dirDOT = ����� � �������� ������������ ���������
'String dirDOC = ����� ��� ����������� ���������� ���������
 Set SvcService = CreateObject("Svcsvc.Service") '������ ����������  svcsvc.dll
 Set colTextDoc = New Collection '��������� ��������� ��� ����������� ������� �������
 
'===================�������� ��� ��������� ������ ���������=======================
'���������� ��������� colTextDoc!!!
'    Dim str1 As String, str2 As String, str3 As String
'            str 1 = frmNewEF.newEF.getLegalGround(frmNewEF.newCor.print_Cor, frmNewEF.newCase.printCsData)
'        colTextDoc.Add str1, "legalGround"
''       ����������� ��������:
'        If frmNewEF.chkAddExData.Value = 1 Then
'            str2 = frmNewEF.newAExperts.print_Expert
'                colTextDoc.Add str2, "Print_Expert"
'        str3 = frmNewEF.newAEF.print_AEdoc(frmNewEF.newAExperts.print_Expert, frmNewEF.newInjPr.print_Autopsy)
'            colTextDoc.Add str2, "print_AEdoc"
'        End If
'===================�������� ��� ��������� ������ ���������=======================

    Set MyWdApp = New Word.Application
        With MyWdApp
            .Visible = False '��������� ���������� �� ����� ��������� ���������
''           1)�������� �� ������������� ����� � ��������� � ����� ��� ����������� ����������
'                If dirDOT = "" Then
'                    MsgBox "����� � ��������� �� �������!", vbExclamation, CRITICAL
'                    dirDOT = SvcService.SelectFolder("������� ����� � ���������!", "", &H10 + &H4000, "")
'            ElseIf dirDOC = "" Then
'                MsgBox "����� ��� ����������� �� �������!", vbExclamation, CRITICAL
'                    dirDOC = SvcService.SelectFolder("������� ����� ��� �����������!", "", &H10 + &H4000, "")
'                End If
'          ������ ����������:
'           1) ���������� ��������
            Call printEF(mdMainFolders.arrDocDir(1), mdMainFolders.arrDocDir(0), DocCat)
'           2)�����������
            Call printFotoList(mdMainFolders.arrDocDir(1))
'           3)��������
            Call print_Labels(mdMainFolders.arrDocDir(1))
'           4)���������������� ����� � �������� �� ��� ��������
                With frmDocuments
                    Dim x As Object
                        For Each x In .Controls
                            If TypeName(x) = "CheckBox" Then
                                If x.Value = 1 Then
                                    If x.Tag = "Description" Then
                                        Call print_Description(mdMainFolders.arrDocDir(1))
                                    ElseIf x.Tag = "Hodataistvo" Then
                                        Call print_Hodataistvo(mdMainFolders.arrDocDir(1))
                                    ElseIf x.Tag = "Nesootvetstvie" Then
                                        Call print_Nesootvetstvie(mdMainFolders.arrDocDir(1), x.Tag, x.Caption)
                                    Else
                                        Call print_Soprovod(mdMainFolders.arrDocDir(1), x.Tag, x.Caption)
                                    End If 'Tag
                                End If 'X.Value
                            End If 'TypeName
                        Next x 'Controls
                End With 'frmDocuments
            .Quit '�������� ����������
        End With
Set MyWdApp = Nothing
Set SvcService = Nothing   '�������� ������� ����������  svcsvc.dll
End Sub
'
Public Static Sub printEF(dirDOT As String, dirDOC As String, ByVal DocCat As String)
'�������� ��������� "���������� ��������"
'String dirDOT = ����� � �������� ������������ ���������
'String dirDOC = ����� ��� ����������� ���������� ���������
'String tmpDocCat = ��������� ������������ ��������� (������/��������� ���� � �.�.)
'    Debug.Print "����� � �������� = " & dirDOT & "\" & strDOT
    Set MyWdDoc = MyWdApp.Documents.Open(dirDOT & "\" & strDOT)    '�������� �������
    MyWdDoc.Activate ' ���������� ������
    With ActiveDocument
'   1)  ���������� �������������� ����� ���������
'           ���� ������ ������ ��������:
'            Dim i As Long, str As String
'                For i = LBound(arrExpert) To UBound(arrExpert)
'                    If i = UBound(arrExpert) Then
'                        str = str & frmNewEF.newExperts.create_revExpert(arrExpert(i)) '& Chr(10)
'                    Else
'                        str = str & frmNewEF.newExperts.create_revExpert(arrExpert(i)) & vbCrLf '& Chr(10)
'                    End If
'                Next i
'                ActiveDocument.Bookmarks("experts").Range.Text = str
'                Debug.Print "������ ��������� - " & str
        If .Bookmarks.Exists("WdArchDocName") = True Then
            .FormFields("WdArchDocName").Result = frmNewEF.newEF.getArchName(frmNewEF.newEF.number, "�������")
        End If
        If .Bookmarks.Exists("WdDirDOC") = True Then
            .FormFields("WdDirDOC").Result = mdMainFolders.arrDocDir(2)
        End If
        If .Bookmarks.Exists("WdEFNum") = True Then
            .FormFields("WdEFNum").Result = frmNewEF.newEF.number
        End If
'   2)���������� ��������� ��� ����������� ������
        .SaveAs2 FileName:=mdMainFolders.arrDocDir(2) & "\����������_" & frmNewEF.newEF.getNum_Cat, FileFormat:= _
            wdFormatXMLDocumentMacroEnabled, LockComments:=False, Password:="", _
            AddToRecentFiles:=True, WritePassword:="", ReadOnlyRecommended:=False, _
            EmbedTrueTypeFonts:=False, SaveNativePictureFormat:=False, SaveFormsData _
            :=False, SaveAsAOCELetter:=False
'    3)������ � ������ ���������:
'       2.1 ������ ����� ����������:
        If .Bookmarks.Exists("WdEFfullNum") = True Then '�������� �� ������������� ��������:
            .Bookmarks("WdEFfullNum").Range.Text = frmNewEF.newEF.getFullNumber
        End If
'       2.2 ���� ������ ����������
        .FormFields("WdFirstDay").Result = frmNewEF.newEF.firstDay
'       2.3 ���� ������ �� �������������
    If .Bookmarks.Exists("getExperience") = True Then '�������� �� ������������� ��������:
        .Bookmarks("getExperience").Range.Text = frmNewEF.newDate.getExperience(frmNewEF.newDate.experience, frmNewEF.newDate.DateNow)
    End If
'       2.4 ����������� ���������:
    If .Bookmarks.Exists("legalGround") = True Then 'legalGround
        Dim lg1 As String, lg2 As String, lg3 As String, lg4 As String '������������ ������:
            lg3 = ", ������ �������� ������-������������������ ����������."
        With frmNewEF
            Dim str1 As String, str2 As String
                str1 = .newEF.getLegalGround(.newCor.print_Cor, .newCase.printCsData)
            colTextDoc.Add str1, "legalGround"
            lg1 = "�� " & str1 'lg1 = "�� " & colTextDoc("legalGround")
            If frmNewEF.chkAddExData.Value = 1 Then
                str2 = .newAExperts.print_Expert
                colTextDoc.Add str2, "print_AEdoc"
                lg2 = "�����������" & str2
                lg4 = .newAEF.print_AEdoc(lg2, .newInjPr.print_Autopsy)
                If ActiveDocument.Bookmarks.Exists("print_AEdoc") = True Then
                    ActiveDocument.Bookmarks("print_AEdoc").Range.Text = lg4
                End If
                ActiveDocument.Bookmarks("legalGround").Range.Text = lg1 & "� " & lg2 & lg3
            Else
                ActiveDocument.Bookmarks("legalGround").Range.Text = lg1 & lg3
            End If
        End With 'frmNewEF
    End If 'legalGround
    
    Dim tmpBx As clmEvBox
    Dim x As Long, bxKey As String, tmpStr1 As String
'       2.5 ������ ��������������� �� � ���� ������.
    If .Bookmarks.Exists("wdEvidences_L") = True Then
            With frmNewEF.colBoxes
                For x = 1 To .Count
                    bxKey = BOX & CStr(Format(x, "#0000"))
                    Set tmpBx = .Item(bxKey)
                       tmpStr1 = tmpStr1 & tmpBx.print_LineEvidBxEntrants
                Next x
            End With
           ActiveDocument.Bookmarks("wdEvidences_L").Range.Text = tmpStr1
        End If
' �������
     If .Bookmarks.Exists("WdBoxes") = True Then
     Dim tmpStr2 As String
            With frmNewEF.colBoxes
                For x = 1 To .Count
                    bxKey = BOX & CStr(Format(x, "#0000"))
                    Set tmpBx = .Item(bxKey)
                       tmpStr2 = tmpStr2 & tmpBx.print_DeliveryEv & Chr(10)
                Next x
            End With
          ActiveDocument.Bookmarks("WdBoxes").Range.Text = tmpStr2
        End If
    Set tmpBx = Nothing
    .Close
    End With 'ActiveDocument
End Sub
'
Public Static Sub print_Description(dirDOT As String)
Set MyWdDoc = MyWdApp.Documents.Open(dirDOT & "\" & "Description.docm")    '�������� �������
    MyWdDoc.Activate ' ���������� ������
    With ActiveDocument
'   ���������� ��������� ��� ����������� ������
        .SaveAs2 FileName:=mdMainFolders.arrDocDir(2) & "\��������_" & frmNewEF.newEF.getNumber, FileFormat:= _
            wdFormatXMLDocumentMacroEnabled, LockComments:=False, Password:="", _
            AddToRecentFiles:=True, WritePassword:="", ReadOnlyRecommended:=False, _
            EmbedTrueTypeFonts:=False, SaveNativePictureFormat:=False, SaveFormsData _
            :=False, SaveAsAOCELetter:=False
''       ������ � ������ ���������:
        If .Bookmarks.Exists("InfoEF") = True Then
            ActiveDocument.Bookmarks("InfoEF").Range.Text = frmNewEF.newEF.toString_InfoEF(colTextDoc("legalGround"), frmNewEF.newExpert) & " (���/���� (017) 308-67-53)." & Chr(10)
        End If
'   ������� � ��
  Dim tmpBx As clmEvBox
    Dim x As Long, bxKey As String, tmpStr1 As String
     If .Bookmarks.Exists("WdBoxes") = True Then
            With frmNewEF.colBoxes
                For x = 1 To .Count
                    bxKey = BOX & CStr(Format(x, "#0000"))
                    Set tmpBx = .Item(bxKey)
                       tmpStr1 = tmpStr1 & tmpBx.print_NumericColumnBoxes & Chr(10)
                Next x
            End With
          ActiveDocument.Bookmarks("WdBoxes").Range.Text = tmpStr1
        End If
    Set tmpBx = Nothing
    .Close
    End With 'ActiveDocument
End Sub
'
Public Sub print_Hodataistvo(dirDOT As String)
Set MyWdDoc = MyWdApp.Documents.Open(dirDOT & "\" & "Hodataistvo.docm")    '�������� �������
    MyWdDoc.Activate ' ���������� ������
    With ActiveDocument
'   ���������� ��������� ��� ����������� ������
        .SaveAs2 FileName:=mdMainFolders.arrDocDir(2) & "\�����������_" & frmNewEF.newEF.getNumber, FileFormat:= _
            wdFormatXMLDocumentMacroEnabled, LockComments:=False, Password:="", _
            AddToRecentFiles:=True, WritePassword:="", ReadOnlyRecommended:=False, _
            EmbedTrueTypeFonts:=False, SaveNativePictureFormat:=False, SaveFormsData _
            :=False, SaveAsAOCELetter:=False
'    ������ � ������ ���������:
        If .Bookmarks.Exists("wdCoroner") = True Then
            ActiveDocument.Bookmarks("wdCoroner").Range.Text = frmNewEF.newCor.print_Cor
        End If

        If .Bookmarks.Exists("wdFullNumber") = True Then
            .FormFields("wdFullNumber").Result = frmNewEF.newEF.getFullNumber
        End If
        
        If .Bookmarks.Exists("wdLegalGround") = True Then
            ActiveDocument.Bookmarks("wdLegalGround").Range.Text = frmNewEF.newEF.getLegalGround(, frmNewEF.newCase.printCsData)
        End If
        
        If .Bookmarks.Exists("wsEFfirstDay") = True Then
            ActiveDocument.Bookmarks("wsEFfirstDay").Range.Text = frmNewEF.newEF.firstDay
        End If
        
        If .Bookmarks.Exists("wdFinDate") = True Then
            Dim dt As Date, strDt As String
            With frmNewEF.newDate
                dt = .getPeriod(.DateNow, 30)
                strDt = .dateToString(dt)
            End With
            ActiveDocument.Bookmarks("wdFinDate").Range.Text = strDt
        End If
         If .Bookmarks.Exists("experts") = True Then
             ActiveDocument.Bookmarks("experts").Range.Text = frmNewEF.newExpert
        End If
    .Close
    End With 'ActiveDocument
End Sub
'
Public Static Sub print_Soprovod(dirDOT As String, dotName As String, docName As String)
'dirDOT - ���������� �������
'dotNam - ��� ������� (����: "Biology")
'docNam - ��� ���������, ��� ������� �� ����� �������� (����.: "��������_��������_���_YY")
If dotName <> "" Or docName <> "" Then
    Set MyWdDoc = MyWdApp.Documents.Open(dirDOT & "\" & dotName & ".docm")    '�������� �������
    MyWdDoc.Activate ' ���������� ������
    With ActiveDocument
'   ���������� ��������� ��� ����������� ������
        .SaveAs2 FileName:=mdMainFolders.arrDocDir(6) & "\" & frmNewEF.newEF.getNumber & "_" & docName & ".docm", FileFormat:= _
            wdFormatXMLDocumentMacroEnabled, LockComments:=False, Password:="", _
            AddToRecentFiles:=True, WritePassword:="", ReadOnlyRecommended:=False, _
            EmbedTrueTypeFonts:=False, SaveNativePictureFormat:=False, SaveFormsData _
            :=False, SaveAsAOCELetter:=False
'       ������ � ������ ���������:
        If .Bookmarks.Exists("InfoEF") = True Then
            ActiveDocument.Bookmarks("InfoEF").Range.Text = colTextDoc("legalGround")
        End If
'       ��������:
'       !!! �������� ��������� ������� (���� ������� ����� ����� ��������)
        If .Bookmarks.Exists("experts") = True Then
             ActiveDocument.Bookmarks("experts").Range.Text = frmNewEF.newExpert
        End If
'       ������� � ��
     Dim tmpStr As String, tmpStr1 As String, x As Long, bxKey As String
            With frmNewEF.colBoxes
                If .Count > 0 Then
                    If .Count = 1 Then
                        tmpStr = "������"
                    Else: tmpStr = "�������"
                    End If

                    For x = 1 To .Count
                        bxKey = BOX & CStr(Format(x, "#0000"))
                        Set tmpBx = .Item(bxKey)
                        tmpStr1 = tmpStr1 & tmpBx.print_NumericColumnBoxes & Chr(10)
                    Next x
                End If '.Count > 0
            End With
        With ActiveDocument
            If .Bookmarks.Exists("wdObject") = True Then
                .Bookmarks("wdObject").Range.Text = tmpStr
                tmpStr = ""
            End If
            If .Bookmarks.Exists("WdBoxes") = True Then
                .Bookmarks("WdBoxes").Range.Text = tmpStr1
                tmpStr1 = ""
            End If
        End With
    Set tmpBx = Nothing
    .Close
    End With 'ActiveDocument
End If
End Sub
'
Public Static Sub print_Nesootvetstvie(dirDOT As String, dotName As String, docName As String)
'������ ���� ��������������
'dirDOT - ���������� �������
'dotNam - ��� ������� (����: "Biology")
'docNam - ��� ���������, ��� ������� �� ����� �������� (����.: "��������_��������_���_YY")
If dotName <> "" Or docName <> "" Then
    Set MyWdDoc = MyWdApp.Documents.Open(dirDOT & "\" & dotName & ".docm")    '�������� �������
    MyWdDoc.Activate ' ���������� ������
    With ActiveDocument
'   ���������� ��������� ��� ����������� ������
        .SaveAs2 FileName:=mdMainFolders.arrDocDir(2) & "\" & frmNewEF.newEF.getNumber & "_" & docName & ".docm", FileFormat:= _
            wdFormatXMLDocumentMacroEnabled, LockComments:=False, Password:="", _
            AddToRecentFiles:=True, WritePassword:="", ReadOnlyRecommended:=False, _
            EmbedTrueTypeFonts:=False, SaveNativePictureFormat:=False, SaveFormsData _
            :=False, SaveAsAOCELetter:=False
'       ������ � ������ ���������:
'       2.1 ������ ����� ����������:
        If .Bookmarks.Exists("WdEFfullNum") = True Then '�������� �� ������������� ��������:
            .FormFields("WdEFfullNum").Result = frmNewEF.newEF.getFullNumber
        End If
'       2.2 ���� ������ ����������
        If .Bookmarks.Exists("WdFirstDay") = True Then '�������� �� ������������� ��������:
             .Bookmarks("WdFirstDay").Range.Text = frmNewEF.newEF.firstDay
        End If
'       2.3 ���� ����������� ��
        If .Bookmarks.Exists("wdBxFirstDate") = True Then '�������� �� ������������� ��������:
             .Bookmarks("wdBxFirstDate").Range.Text = frmNewEF.newBox.DtmBxFirstDate
        End If
'       2.4 ��������� ����
        If .Bookmarks.Exists("wdCsData") = True Then '�������� �� ������������� ��������:
            .FormFields("wdCsData").Result = frmNewEF.newCase.printCsData
        End If
'       2.5 ����������� ���������
         If .Bookmarks.Exists("InfoEF") = True Then
            ActiveDocument.Bookmarks("InfoEF").Range.Text = colTextDoc("legalGround")
        End If
'       2.6 �������
        If .Bookmarks.Exists("experts") = True Then
             ActiveDocument.Bookmarks("experts").Range.Text = frmNewEF.newExpert
        End If
'       2.7 �����������
        If .Bookmarks.Exists("wdCoroner") = True Then
            ActiveDocument.Bookmarks("wdCoroner").Range.Text = frmNewEF.newCor.print_Cor
        End If
'       2.6������� � ��
     Dim tmpStr As String, tmpStr1 As String, x As Long, bxKey As String
            With frmNewEF.colBoxes
                If .Count > 0 Then
                    If .Count = 1 Then
                        tmpStr = "������"
                    Else: tmpStr = "�������"
                    End If

                    For x = 1 To .Count
                        bxKey = BOX & CStr(Format(x, "#0000"))
                        Set tmpBx = .Item(bxKey)
                        tmpStr1 = tmpStr1 & tmpBx.print_NumericColumnBoxes & Chr(10)
                    Next x
                End If '.Count > 0
            End With
        With ActiveDocument
            If .Bookmarks.Exists("wdObject") = True Then
                .Bookmarks("wdObject").Range.Text = tmpStr
                tmpStr = ""
            End If
            If .Bookmarks.Exists("WdBoxes") = True Then
                .Bookmarks("WdBoxes").Range.Text = tmpStr1
                tmpStr1 = ""
            End If
        End With
    Set tmpBx = Nothing
    .Close
    End With 'ActiveDocument
End If
End Sub
'
Public Static Sub printFotoList(dirDOT As String, Optional dotName As String = "FotoList", Optional docName As String = "�����������")
'�������� � ���������� �����������
'dirDOT - ���������� �������
'dotNam - ��� ������� (����: "Biology")
'docNam - ��� ���������, ��� ������� �� ����� �������� (����.: "��������_��������_���_YY")
Set MyWdDoc = MyWdApp.Documents.Open(dirDOT & "\" & dotName & ".docm")    '�������� �������
    MyWdDoc.Activate ' ���������� ������
    With ActiveDocument
'   ���������� ��������� ��� ����������� ������
        .SaveAs2 FileName:=mdMainFolders.arrDocDir(3) & "\" & docName & "_" & frmNewEF.newEF.getNumber, FileFormat:= _
            wdFormatXMLDocumentMacroEnabled, LockComments:=False, Password:="", _
            AddToRecentFiles:=True, WritePassword:="", ReadOnlyRecommended:=False, _
            EmbedTrueTypeFonts:=False, SaveNativePictureFormat:=False, SaveFormsData _
            :=False, SaveAsAOCELetter:=False
'       ������ � ������ ���������:
        If .Bookmarks.Exists("WdEFfullNum") = True Then
            ActiveDocument.Bookmarks("WdEFfullNum").Range.Text = frmNewEF.newEF.getFullNumber
        End If
'       ��������:
'     ��������� ������� (���� ������� ����� ����� ��������)
        If .Bookmarks.Exists("experts") = True Then
            Dim i As Long, str As String
                For i = LBound(arrExpert) To UBound(arrExpert)
                    If i = UBound(arrExpert) Then
                        str = str & frmNewEF.newExperts.create_revExpert(arrExpert(i)) '& Chr(10)
                    Else
                        str = str & frmNewEF.newExperts.create_revExpert(arrExpert(i)) & vbCrLf '& Chr(10)
                    End If
                Next i
                ActiveDocument.Bookmarks("experts").Range.Text = str
                Debug.Print "������ ��������� - " & str
        End If
'           ������ �������
        If .Bookmarks.Exists("WdBoxes") = True Then
            Dim tmpBx As clmEvBox
            Dim x As Long, bxKey As String, tmpStr As String
            With frmNewEF.colBoxes
                For x = 1 To .Count
                    bxKey = BOX & CStr(Format(x, "#0000"))
                    Set tmpBx = .Item(bxKey)
                       tmpStr = tmpStr & tmpBx.print_BoxFotoList & Chr(10)
                Next x
            End With
          ActiveDocument.Bookmarks("WdBoxes").Range.Text = tmpStr
        End If
    .Close
    End With 'ActiveDocument
'LIB
'   ActiveDocument.Bookmarks("experts").Range.Text = frmNewEF.newExperts.create_revExpert(frmNewEF.newExpert)
End Sub
'
Public Static Sub print_Labels(dirDOT As String, _
                                Optional dotName As String = "Label", _
                                Optional docName As String = "��������")
Set MyWdDoc = MyWdApp.Documents.Open(dirDOT & "\" & dotName & ".docm")    '�������� �������
    MyWdDoc.Activate ' ���������� ������
    With ActiveDocument
'   ���������� ��������� ��� ����������� ������
        .SaveAs2 FileName:=mdMainFolders.arrDocDir(4) & "\" & docName & "_" & frmNewEF.newEF.getNumber, FileFormat:= _
            wdFormatXMLDocumentMacroEnabled, LockComments:=False, Password:="", _
            AddToRecentFiles:=True, WritePassword:="", ReadOnlyRecommended:=False, _
            EmbedTrueTypeFonts:=False, SaveNativePictureFormat:=False, SaveFormsData _
            :=False, SaveAsAOCELetter:=False
'       ������ � ������ ���������:
        If .Bookmarks.Exists("wdLabels") = True Then
            Dim tmpBx As clmEvBox
            Dim x As Long, bxKey As String, tmpStr As String
            Dim arrStr(4) As String
                arrStr(0) = frmNewEF.newEF.getFullNumber '������ ����� ���������
                arrStr(1) = frmNewEF.newCase.printCsData & Chr(10) '������ �� �������������
                arrStr(2) = frmNewEF.newEF.firstDay '���� ������ ����������
                Dim i As Long, str As String
                For i = LBound(arrExpert) To UBound(arrExpert)
                    If i = UBound(arrExpert) Then
                        str = str & frmNewEF.newExperts.create_revExpert(arrExpert(i)) & Chr(10)
                    Else
                        str = str & frmNewEF.newExperts.create_revExpert(arrExpert(i)) & vbCrLf & Chr(10)
                    End If
                Next i
                arrStr(3) = str
            With frmNewEF.colBoxes
                For x = 1 To .Count
                    bxKey = BOX & CStr(Format(x, "#0000"))
                    Set tmpBx = .Item(bxKey)
'                    tmpStr = tmpStr & tmpBx.print_NumericColumnBoxes, arrStr(1), arrStr(2), arrStr(3)) & Chr(10) & Chr(10)
                       tmpStr = tmpStr & tmpBx.print_Labels(arrStr(0), arrStr(1), arrStr(2), arrStr(3)) & Chr(10) & Chr(10)
                Next x
            End With
          ActiveDocument.Bookmarks("wdLabels").Range.Text = tmpStr
        End If
         .Close
    End With 'ActiveDocument
End Sub

Public Static Sub openEx()
'��������� �������� msExcel
 Dim n As Long '�������
 Dim numRow, tmpRow As Integer '�������� ������ ������ ��� ������ ������
 Dim arrStr(8) As String
    arrStr(0) = "������� ����� -> " & strMonth
    arrStr(1) = "A"
    arrStr(2) = "B"
    arrStr(3) = "C"
    arrStr(4) = "D"
    arrStr(5) = "E"
    arrStr(6) = "G"
    arrStr(7) = "j"
'   ������ � Excel
    Set myExApp = New Excel.Application
        myExApp.Visible = False
    Dim tpmDocName As String
        tpmDocName = mdMainFolders.arrDocDir(0) & "\�����_��������������_ " & Year(Now) & ".xlsm"
'   �������� ��������� Excel
    Workbooks.Open FileName:=tpmDocName ', Password:="���7004"
'   �������� ����� � ������ �������� ������
        Worksheets.Item(strMonth).Activate
        With ActiveWorkbook.Sheets
'            .Item(monthNow).Activate ' ��������� ����� � ������ �������� ������
            With .Application
'               �������� �������� ���������� ���������:
                tmpRow = .Cells(3, 3).Value
                numRow = tmpRow + 13
                For n = 1 To 7
                    .Cells(numRow, arrStr(n)).Activate   '���������� ������ ��� ������ ������:
                    .ActiveCell.Value = mdPrintDoc.colExcelData.Item(arrStr(n)) '������ ������ �� ������� � ������:
                Next n
            End With
        End With
        ActiveWorkbook.Save
    myExApp.Quit
    Set myExApp = Nothing
End Sub
    
    
Public Sub closeEx()
'�������� ����������  msExcel

'       .Save

Set myExDoc = Nothing
Set myExApp = Nothing
End Sub

'Private Static Sub Create_ReportEF()
''������ � �������� Excel
''MyExApp.Visible = False
'''CellArr(1) = "A"
'''CellArr(2) = "B"
'''CellArr(3) = "C"
'''CellArr(4) = "D"
'''CellArr(5) = "E"
'''CellArr(6) = "F"
'''CellArr(7) = "G"
'''CellArr(8) = "H"
'''CellArr(9) = "I"
'''CellArr(10) = "J"
'''CellArr(11) = "K"
'''CellArr(12) = "L"
'''CellArr(13) = "M"
'''CellArr(14) = "N"
'''CellArr(15) = "O"
'''CellArr(16) = "P"
'''CellArr(17) = "Q"
'''CellArr(18) = "R"
'''CellArr(19) = "S"
'''CellArr(20) = "T"
'''CellArr(21) = "U"
'''CellArr(22) = "V"
'''CellArr(23) = "W"
'''CellArr(24) = "X"
'''CellArr(25) = "Y"
'''CellArr(26) = "Z"
'Dim intX As Integer, intNEF As Integer, intTmpDt As Integer, curAlc As Currency
''�������� ������� ����� "����� ��������������"
'With Workbooks
'    .Open FileName:="E:\Crime\Reports\����� ��������������_" & Year(Now) & ".xlsm", Password:="���7004"
''��������� ����� ��������� �����
''    ������ "�������� ����� ���������"
'        Range("D10").Select
'        intNEF = ActiveCell.Value
'        intX = intNEF + 30
''    ������ "���������� �����" "A"
'        Range(CellArr(1) & intX).Select
'        ActiveCell.FormulaR1C1 = intX - 29
''    ������ "����� ����������" "B"
'        Range(CellArr(2) & intX).Select
'        ActiveCell.FormulaR1C1 = CInt(NewExpertFindings.strEFNum)
''    ������ "���� ������" "C"
'        Range(CellArr(3) & intX).Select
'        ActiveCell.FormulaR1C1 = NewExpertFindings.DtmEFFirstDay
''    '������ "���������" "D"
'        Range(CellArr(4) & intX).Select
'        ActiveCell.FormulaR1C1 = newCase.strCsCat
''    ������ "� ����" "G"
'        Range(CellArr(7) & intX).Select
'        ActiveCell.FormulaR1C1 = newCase.strCsNum
''    ������ "�������" "I"
''        Range(CellArr(9) & intX).Select
''        ActiveCell.FormulaR1C1 = i
''   ������ "��������������" "J"
'        Range(CellArr(10) & intX).Select
'''            If frmDocList.chkAEFInquiry = 1 Then
'''                NewExpertFindings.DtmEFSuspDate = DateTime.Date
'''                ActiveCell.FormulaR1C1 = NewExpertFindings.DtmEFSuspDate
'''            End If
''   ������ "���� 15" "N"
''        Range(CellArr(14) & intX).Select
''            If frmMDI.strCrEFCat = "������" Then
''              intTmpDt = 15
''              ActiveCell.FormulaR1C1 = NewExpertFindings.Calculate_EFPeriod(ByVal intTmpDt)
''            End If
''   ������ "����" "P"
'        Range(CellArr(16) & intX).Select
'        intTmpDt = 30
'        ActiveCell.FormulaR1C1 = NewExpertFindings.Calculate_EFPeriod(ByVal intTmpDt)
''   �������� ��������� ������ � ������ "�������� ����� ���������" (�.�. +1)
'            Range("D10").Select
'            ActiveCell.FormulaR1C1 = intNEF + 1
''   �������� ������ "S"
'            Range(CellArr(19) & intX).Select
'            ActiveCell.FormulaR1C1 = NewExpertFindings.DtmEFFirstDay
''  ������ ������� ������
'        Range(CellArr(20) & intX - 1).Select
'        curAlc = ActiveCell.Value '����������� ����������� ������� ������
'        Range(CellArr(20) & intX).Select
'        ActiveCell.FormulaR1C1 = curAlc - 0.1
'    End With
'ActiveWorkbook.Save
'Workbooks.Close
'MyExApp.Quit
'End Sub


