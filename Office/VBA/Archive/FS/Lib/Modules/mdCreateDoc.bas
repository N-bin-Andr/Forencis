Attribute VB_Name = "mdCreateDoc"
'@author Andr.Nab.n@gmail.com
Option Explicit
'Public Const DOCDir As String = "D:\Crime\"
'Public MyFields As FormFields '���������� ���������������� �����
'Private fn As String 'FN - ��������� ���������� �������� ���������� ���� "WdEFNum" ��������� ���������
'Private docCat As String
'Private Const Msg As String = "�������, ����������, ����� ����������"
'Private Const Title As String = "���� ��� ����� ������ ���"
'Private Const PUBLICDIR As String = "D:\��� ������\" '���������� ����� ����� "��� ������"
'Private strDocDir1 As String '���������� ������� ������
'Private strDocDir2 As String '���������� ������� ������
'Private MyOlApp As New Outlook.Application
'Private MyOlTask As Outlook.TaskItem
'Private MyExApp As New Excel.Application
'
'Lib
''������ ���������� �������� �������� ����������
''For Each aDoc In Documents
''aName = aName & aDoc.Name & vbCr 'vbCr - ��� ���������, ������������ ������ �������� ������� (��� 13)
''Next aDoc
''MsgBox aName
'___________________________________________________________________
'����� ��������� � ��������� �������� ����������
'Dim oDoc1 As Document
'For i = 1 To Documents.Count
'    Set oDoc1 = Documents.Item(i)
'        If oDoc1.Name = "doc1.doc" Then
'            Exit For
'        End If
'    Set oDoc1 = Nothing
'Next



'Public Static Function Create_mainFolders(Optional tmpRoot As String, Optional tmpnameFolder As String)
''�������� ������ �������� ����� � �������� ����������
'Set SvcService = CreateObject("Svcsvc.Service") '������ ����������  svcsvc.dll
'Set fs = New FileSystemObject                   '���������  FileSystemObject
''1)�������� �� ������� ����� �� ���������� ����:
'Dim newDir As String
'    If tmpRoot = "" Then
'        tmpRoot = MAIN_ROOT
'     '�������� ������� ���������� MAIN_ROOT
'        If Not fs.FolderExists(tmpRoot) Then  ' ���� ����� ����� �� ���������� � ���� �����
'            MsgBox MAIN_FOLDER_NOT_EXIST & "������� ����� ����� ��� ���������� ����������!", vbCritical, CRITICAL
'            newDir = SvcService.SelectFolder("�������� ����� ��� ���������� ����������", "", &H10 + &H4000, "")
'        Else: newDir = tmpRoot
'        End If
'    Else
'        newDir = tmpRoot
'    End If
'    Err.Clear ' ������� ����� �� ������ ��� ��������� ��������!
''2)�������� ����� ����� � ������� ���� � ���:
'    '���������� �������� �����:
'    newDir = newDir & tmpnameFolder & "\"   '������: D:\�������_��������������_���\
''3)�������� ������� ����������� ����������
'    If Not fs.FolderExists(newDir) Then  ' ���� ����� ����� �� ���������� � ���� �����
'        MkDir newDir
'    End If
'    Create_mainFolders = newDir
'Set SvcService = Nothing    '������ ����������  svcsvc.dll
'Set fs = Nothing            '���������  FileSystemObject
'Err.Clear ' ������� ����� �� ������ ��� ���������� ��������
'End Function
''

'
'
'
'Public Sub makeDocDir()
'With ActiveDocument
'    fn = InputBox(Msg, Title)
'    .FormFields("WdEFNum").Result = fn
'    .FormFields("WdArchDocName").Result = criateArchDocName(fn)
'    docCat = .FormFields("WdDocCat").Result
'    Call caseDate.caseDate
'    '�������� ����� � ������� ����������
'    ChDrive "D"
'        strDocDir1 = DOCDir & "\" & CStr(Year(Now)) & "\"
'        strDocDir2 = strDocDir1 & fn & "_" & Right(CStr(Year(Now)), 2) & "_" & docCat
'        MkDir strDocDir2 '�������� ����������
'        .FormFields("WdDirDOC").Result = strDocDir2 '& "\"
'        Dim strDocDir3 As String, strDocDir4 As String, strDocDir5 As String '���������� �������� ������
'            strDocDir3 = strDocDir2 & "\����_" & criateDocName(fn)
'        MkDir strDocDir3 '�������� ����������
'            strDocDir4 = strDocDir2 & "\��������_" & criateDocName(fn)
'        MkDir strDocDir4 '�������� ����������
'            strDocDir5 = strDocDir2 & "\�����_" & criateDocName(fn)
'        MkDir strDocDir5 '�������� ����������
''���������� ��������� ��������� � ����� "��� ������"
'  ChangeFileOpenDirectory PUBLICDIR
'    ActiveDocument.SaveAs2 FileName:="����������_" & criateDocName(fn), FileFormat:= _
'        wdFormatXMLDocumentMacroEnabled, LockComments:=False, Password:="", AddToRecentFiles _
'        :=True, WritePassword:="", ReadOnlyRecommended:=False, EmbedTrueTypeFonts _
'        :=False, SaveNativePictureFormat:=False, SaveFormsData:=False, _
'        SaveAsAOCELetter:=False, CompatibilityMode:=14
'End With
''    '���������� ��������� ��������� � ������������� ������ �����
'    ChangeFileOpenDirectory strDocDir2 & "\"
'    ActiveDocument.SaveAs2 FileName:="����������_" & criateDocName(FN), FileFormat:= _
'        wdFormatXMLDocumentMacroEnabled, LockComments:=False, Password:="", AddToRecentFiles _
'        :=True, WritePassword:="", ReadOnlyRecommended:=False, EmbedTrueTypeFonts _
'        :=False, SaveNativePictureFormat:=False, SaveFormsData:=False, _
'        SaveAsAOCELetter:=False, CompatibilityMode:=14
'End With
'    With Documents
'        .Add ("D:\Crime\MasterForm\Word\Fotolist.dotm")
'        With Application.ActiveDocument
'            .Bookmarks("WdEFfullNum1").Range = fn
'            .Bookmarks("WdEFfullNum2").Range = fn
'            .Bookmarks("WdEFfullNum3").Range = fn
'        End With
'        ChangeFileOpenDirectory strDocDir2 & "\"
'            ActiveDocument.SaveAs2 FileName:="�����������_" & criateDocName(fn) & ".docm", FileFormat:= _
'                wdFormatXMLDocumentMacroEnabled, LockComments:=False, Password:="", AddToRecentFiles _
'            :=True, WritePassword:="", ReadOnlyRecommended:=False, EmbedTrueTypeFonts _
'            :=False, SaveNativePictureFormat:=False, SaveFormsData:=False, _
'            SaveAsAOCELetter:=False, CompatibilityMode:=14
'         .Item("�����������_" & criateDocName(fn) & ".docm").Close
'    End With
''Call CreateTaskItem
'End Sub
''
'Public Function criateDocName(Optional strFN As String = "") As String
''������������ ����� �������� ���������
'    criateDocName = fn & "_" & Right(CStr(Year(Now)), 2)
'End Function
''
'Public Function criateArchDocName(Optional strFN As String = "") As String
''������������ ��������� ����� ���������
'        criateArchDocName = "012_" & Format(fn, "#00000") & "�_49_" & Right(CStr(Year(Now)), 2) & "_�������_�����"
'End Function
''

'Private Static Sub CreateTaskItem()
'�������� ������ Outlook
'Dim newDate As clmCaseDate
'Set newDate = New clmCaseDate
'Set MyOlTask = MyOlApp.CreateItem(ItemType:=olTaskItem) 'CreateItem(ItemType:=olTaskItem)
'    With MyOlTask
'        .StartDate = ActiveDocument.FormFields("WdFirstDay").Result '���� ������
'        .dueDate = newDate.getPeriod(newDate.ExamDate(.StartDate), 30) '����
'        .categories = ActiveDocument.FormFields("WdDocCat").Result '���������
'        .Subject = criateDocName(fn) '& .Bookmarks("WdEF..�����������").Range.Text  '����
'        '����������
'
'        .Body = "NB!!! �������� �� " & newDate.getReminder(.StartDate, 30) & vbCr
        
        
        
'            If optClothes.Value = True Then
'                .Body = "��������� " & strCsCat & " �" & strCsNum & Chr(10) & _
'                "�������������: " & strCrPost & Chr(10) & _
'                strCrOprAr & Chr(10) & _
'                strCrRank & Chr(32) & strCrSurName & Chr(32) & strCrName & Chr(32) & strCrMidName & Chr(10) & _
'                Chr(10) & " ������������:" & Chr(10) & Create_EFReminder15(ByVal CrEFFirstDay)
'            Else: .Body = "��������� " & strCsCat & " �" & strCsNum & Chr(10) & _
'                "�������������: " & strCrPost & Chr(10) & _
'                strCrOprAr & Chr(10) & _
'                strCrRank & Chr(32) & strCrSurName & Chr(32) & strCrName & Chr(32) & strCrMidName & Chr(10) & _
'                Chr(10) & " ������������:" & Chr(10)
''            PrintArray1 i, N
'            End If




          '���������
'        .Status = olTaskInProgress
'        .PercentComplete = 15 '������� ����������
'        .Companies = "���� �� ���������� ������-������������������ ���������" '�����������
'        .BillingInformation = "������� " & ActiveDocument.FormFields("WdEFNum").Result '�������
'        .importance = olImportanceNormal '��������
'''�����������
'        .ReminderSet = True
''        .ReminderTime = newDate.getReminder(.StartDate, 30)
'        .Save
'    End With
'MyOlTask.Display
'End Sub






'Private Static Sub CreateTaskItem(ByVal strDN As String, ByVal strDocName5 As String, ByVal CrEFFirstDay As Date)
''�������� ������ Outlook
'Set MyOlTask = MyOlApp.CreateItem(ItemType:=olTaskItem)
'    With MyOlTask
'        .Categories = strDocName5 '���������
'        .Subject = strDN & "/" & Year(Now) & " " & strInjPrSurName & " " & strInjPrName & " " & strInjPrMidName  '����
'        .StartDate = CrEFFirstDay '���� ������
'        .DueDate = Str(CrPeriod(ByVal CrEFFirstDay, ByVal dtDueDate))  '����
'        '����������
'            If optClothes.Value = True Then
'                .Body = "��������� " & strCsCat & " �" & strCsNum & Chr(10) & _
'                "�������������: " & strCrPost & Chr(10) & _
'                strCrOprAr & Chr(10) & _
'                strCrRank & Chr(32) & strCrSurName & Chr(32) & strCrName & Chr(32) & strCrMidName & Chr(10) & _
'                Chr(10) & " ������������:" & Chr(10) & Create_EFReminder15(ByVal CrEFFirstDay)
'            Else: .Body = "��������� " & strCsCat & " �" & strCsNum & Chr(10) & _
'                "�������������: " & strCrPost & Chr(10) & _
'                strCrOprAr & Chr(10) & _
'                strCrRank & Chr(32) & strCrSurName & Chr(32) & strCrName & Chr(32) & strCrMidName & Chr(10) & _
'                Chr(10) & " ������������:" & Chr(10)
''            PrintArray1 i, N
'            End If
'          '���������
'        .Status = olTaskInProgress
'        .PercentComplete = 15 '������� ����������
'        .Companies = "���� �� ���������� ������-������������������ ���������" '�����������
'        .BillingInformation = "������� " & strDN & "_" & Year(Now) '�������
'        .Importance = olImportanceNormal '��������
'''�����������
''        .ReminderSet = True
''        .ReminderTime = CrEFFirstDay + 7
'        .Save
'    End With
'MyOlTask.Display
'End Sub

Private Static Sub Create_FotoList()
'�������� �����������
'�������� ������� Fotolist
'With Application
'    Dim MyWdDoc As New Word.Document
'        MyWdDoc = .Documents.Add("D:\Crime\MasterForm\Word\FotoList_new.dotm")
'        MyWdDoc.Name = "�����������" & criateDocName(FN)
'
''���������� �����������:
''MyWdDoc.SaveAs2 FileName:=NewExpertFindings.DOC & "�����������_" & NewExpertFindings.strEFNum _
''    & "_" & Right(Year(Now), 2), FileFormat:=wdFormatXMLDocument, _
''    LockComments:=False, Password:="", AddToRecentFiles:=True, WritePassword:="", _
'    ReadOnlyRecommended:=False, EmbedTrueTypeFonts:=False, SaveNativePictureFormat:=False, _
'    SaveFormsData:=False, SaveAsAOCELetter:=False








'    MyWdDoc.SaveAs2 FileName:=NewExpertFindings.DOC & "�����������_" & NewExpertFindings.strEFNum _
'            & "_" & Right(Year(Now), 2), FileFormat:=wdFormatXMLDocument, _
'            LockComments:=False, Password:="", AddToRecentFiles:=True, WritePassword:="", _
'            ReadOnlyRecommended:=False, EmbedTrueTypeFonts:=False, SaveNativePictureFormat:=False, _
'            SaveFormsData:=False, SaveAsAOCELetter:=False
'    With MyWdApp.ActiveDocument  '�������� �������� ��������
'        On Error Resume Next ' ��������� ������
''������ � ������� ������������:
'    If .ActiveWindow.View.SplitSpecial <> wdPaneNone Then
'        .ActiveWindow.Panes(2).Close
'    End If
'    If .ActiveWindow.ActivePane.View.Type = wdNormalView Or ActiveWindow. _
'        ActivePane.View.Type = wdOutlineView Then
'        .ActiveWindow.ActivePane.View.Type = wdPrintView
'    End If
'    .ActiveWindow.ActivePane.View.SeekView = wdSeekCurrentPageHeader
'    .Bookmarks("WdEFfullNum").Range.Text = NewExpertFindings.Criate_FullEFNum
'    .ActiveWindow.ActivePane.View.SeekView = wdSeekMainDocument
'''������ � ��������� ��������� ���������:
''i = 1
''    For i = 1 To n - 1 Step 1
''        .Bookmarks("WdEvColumnArr").Range.Text = "���� � " & "." & Chr(32) & EvidArray(i) & Chr(10) & Chr(10)
''    Next i
''���������� ������ � ����������
'        .Close SaveChanges:=wdSaveChanges
'End With
'Set MyWdDoc = Nothing
End Sub


'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'Public Static Sub Create_EF()
'��������� �������� ������ ��������� Ecspert Findings
'    MyWdDoc.SaveAs2 FileName:=NewExpertFindings.DOC & "����������_" & NewExpertFindings.strEFNum _
'            & "_" & Right(Year(Now), 2), FileFormat:=wdFormatXMLDocument, _
'            LockComments:=False, Password:="", AddToRecentFiles:=True, WritePassword:="", _
'            ReadOnlyRecommended:=False, EmbedTrueTypeFonts:=False, SaveNativePictureFormat:=False, _
'            SaveFormsData:=False, SaveAsAOCELetter:=False
'    With MyWdApp.ActiveDocument  '�������� �������� ��������
'        On Error Resume Next
''����:
'       .FormFields("WdDirDOC").Result = NewExpertFindings.DOC
'       .FormFields("WdEFNum").Result = NewExpertFindings.strEFNum
'       .FormFields("WdFirstDay").Result = NewExpertFindings.DtmEFFirstDay
''��������:
'        .Bookmarks("WdEFfullNum").Range.Text = NewExpertFindings.Criate_FullEFNum '��� ����������
'        .Bookmarks("WdExperience").Range.Text = NewExperts.Calculate_Experience '������ ����� ������
'        .Bookmarks("WdRsnComplain").Range.Text = Print_RsnComplain '��.���������
'        .Bookmarks("FactCase").Range.Text = Print_FactCase '�������������� ����
'    If chkAddExData.Value = 1 Then
'        .Bookmarks("WdAEDirection").Range.Text = Print_AEDirection '����������� ������ ��������
'    End If
'    Dim tmp As String
'        tmp = Me.lblEFCount
'    If mdCount.colForms("DocListform" & tmp).chkAEFInquiry.Value = 1 Then
'        NewExpertFindings.DtmEFSuspDate = DateTime.Date
'        .Bookmarks("WdEFSuspDate").Range.Text = NewExpertFindings.DtmEFSuspDate
'        .Bookmarks("WdAEFInquiry").Range.Text = Print_AEFInquiry '������������ �����������
'    End If
'
'������ �������� �������� � �.�. ������������ �������������
'        .Bookmarks("WdEvNumArr1").Range.Text = NewEvid.Print_EvNumArr1
'        .Bookmarks("WdEvNumArr2").Range.Text = NewEvid.Print_EvNumArr2
'        .Bookmarks("WdEvNumArr3").Range.Text = NewEvid.Print_EvNumArr3
''������ ��:
'    .Bookmarks("WdEvColumnArr").Range.Select
'            Call Print_NumColumn
''         For i = 1 To n - 1
''            .Bookmarks("WdEvColumnArr").Range.Text = EvidArray(i) & Chr(10)
''            .Bookmarks("WdEvLineArr").Range.Text = EvidArray(i) & ", "
''            .Bookmarks("WdEvLineArr1").Range.Text = EvidArray(i) & ", "
''            .Bookmarks("WdEvLineArr3").Range.Text = EvidArray(i) & ", "
''        Next i
''    On Error GoTo 999 ' �������� ��������� ������
'    .Close SaveChanges:=wdSaveChanges
'    End With
'Set MyWdDoc = Nothing
'999:
'    MsgBox Err.Description  '������
'    Err.Clear
'End Sub
'
'Private Static Sub Create_ReportEF()
'������ � �������� Excel
'MyExApp.Visible = False
''CellArr(1) = "A"
''CellArr(2) = "B"
''CellArr(3) = "C"
''CellArr(4) = "D"
''CellArr(5) = "E"
''CellArr(6) = "F"
''CellArr(7) = "G"
''CellArr(8) = "H"
''CellArr(9) = "I"
''CellArr(10) = "J"
''CellArr(11) = "K"
''CellArr(12) = "L"
''CellArr(13) = "M"
''CellArr(14) = "N"
''CellArr(15) = "O"
''CellArr(16) = "P"
''CellArr(17) = "Q"
''CellArr(18) = "R"
''CellArr(19) = "S"
''CellArr(20) = "T"
''CellArr(21) = "U"
''CellArr(22) = "V"
''CellArr(23) = "W"
''CellArr(24) = "X"
''CellArr(25) = "Y"
''CellArr(26) = "Z"
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
'        ActiveCell.FormulaR1C1 = NewCase.strCsCat
''    ������ "� ����" "G"
'        Range(CellArr(7) & intX).Select
'        ActiveCell.FormulaR1C1 = NewCase.strCsNum
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
''




