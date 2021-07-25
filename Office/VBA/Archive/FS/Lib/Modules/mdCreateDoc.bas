Attribute VB_Name = "mdCreateDoc"
'@author Andr.Nab.n@gmail.com
Option Explicit
'Public Const DOCDir As String = "D:\Crime\"
'Public MyFields As FormFields 'Объявление пользовательских полей
'Private fn As String 'FN - строковая переменная значения текстового поля "WdEFNum" активного документа
'Private docCat As String
'Private Const Msg As String = "Введите, пожалуйста, номер экспертизы"
'Private Const Title As String = "Окно для ввода номера СМЭ"
'Private Const PUBLICDIR As String = "D:\Для обмена\" 'директория общей папки "Для обмена"
'Private strDocDir1 As String 'директория первого уровня
'Private strDocDir2 As String 'директория второго уровня
'Private MyOlApp As New Outlook.Application
'Private MyOlTask As Outlook.TaskItem
'Private MyExApp As New Excel.Application
'
'Lib
''пример показывает названия открытых документов
''For Each aDoc In Documents
''aName = aName & aDoc.Name & vbCr 'vbCr - это константа, определяющая символ возврата каретки (код 13)
''Next aDoc
''MsgBox aName
'___________________________________________________________________
'поиск документа в коллекции открытых документов
'Dim oDoc1 As Document
'For i = 1 To Documents.Count
'    Set oDoc1 = Documents.Item(i)
'        If oDoc1.Name = "doc1.doc" Then
'            Exit For
'        End If
'    Set oDoc1 = Nothing
'Next



'Public Static Function Create_mainFolders(Optional tmpRoot As String, Optional tmpnameFolder As String)
''Создание нового каталога папок в заданной директории
'Set SvcService = CreateObject("Svcsvc.Service") 'объект библиотеки  svcsvc.dll
'Set fs = New FileSystemObject                   'экземпляр  FileSystemObject
''1)проверка на наличие папки по указанному пути:
'Dim newDir As String
'    If tmpRoot = "" Then
'        tmpRoot = MAIN_ROOT
'     'Проверка наличия директории MAIN_ROOT
'        If Not fs.FolderExists(tmpRoot) Then  ' если такая папка не существует в этом месте
'            MsgBox MAIN_FOLDER_NOT_EXIST & "Укажите новую папку для сохранения документов!", vbCritical, CRITICAL
'            newDir = SvcService.SelectFolder("Выбираем папку для сохранения документов", "", &H10 + &H4000, "")
'        Else: newDir = tmpRoot
'        End If
'    Else
'        newDir = tmpRoot
'    End If
'    Err.Clear ' Очищаем поток от ошибки при отсутсвии элемента!
''2)создание новой папки и возврат пути к ней:
'    'директория корневой папки:
'    newDir = newDir & tmpnameFolder & "\"   'Пример: D:\Эксперт_НомерДокумента_Год\
''3)Проверка наличия создаваемой директории
'    If Not fs.FolderExists(newDir) Then  ' если такая папка не существует в этом месте
'        MkDir newDir
'    End If
'    Create_mainFolders = newDir
'Set SvcService = Nothing    'объект библиотеки  svcsvc.dll
'Set fs = Nothing            'экземпляр  FileSystemObject
'Err.Clear ' Очищаем поток от ошибки при отсутствии элемента
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
'    'создание папок в рабочей директории
'    ChDrive "D"
'        strDocDir1 = DOCDir & "\" & CStr(Year(Now)) & "\"
'        strDocDir2 = strDocDir1 & fn & "_" & Right(CStr(Year(Now)), 2) & "_" & docCat
'        MkDir strDocDir2 'создание директории
'        .FormFields("WdDirDOC").Result = strDocDir2 '& "\"
'        Dim strDocDir3 As String, strDocDir4 As String, strDocDir5 As String 'директории третьего уровня
'            strDocDir3 = strDocDir2 & "\Фото_" & criateDocName(fn)
'        MkDir strDocDir3 'создание директории
'            strDocDir4 = strDocDir2 & "\Упаковки_" & criateDocName(fn)
'        MkDir strDocDir4 'создание директории
'            strDocDir5 = strDocDir2 & "\Сканы_" & criateDocName(fn)
'        MkDir strDocDir5 'создание директории
''сохранение активного документа в папке "Для обмена"
'  ChangeFileOpenDirectory PUBLICDIR
'    ActiveDocument.SaveAs2 FileName:="Заключение_" & criateDocName(fn), FileFormat:= _
'        wdFormatXMLDocumentMacroEnabled, LockComments:=False, Password:="", AddToRecentFiles _
'        :=True, WritePassword:="", ReadOnlyRecommended:=False, EmbedTrueTypeFonts _
'        :=False, SaveNativePictureFormat:=False, SaveFormsData:=False, _
'        SaveAsAOCELetter:=False, CompatibilityMode:=14
'End With
''    'сохранение активного документа с присваиванием нового имени
'    ChangeFileOpenDirectory strDocDir2 & "\"
'    ActiveDocument.SaveAs2 FileName:="Заключение_" & criateDocName(FN), FileFormat:= _
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
'            ActiveDocument.SaveAs2 FileName:="Фототаблица_" & criateDocName(fn) & ".docm", FileFormat:= _
'                wdFormatXMLDocumentMacroEnabled, LockComments:=False, Password:="", AddToRecentFiles _
'            :=True, WritePassword:="", ReadOnlyRecommended:=False, EmbedTrueTypeFonts _
'            :=False, SaveNativePictureFormat:=False, SaveFormsData:=False, _
'            SaveAsAOCELetter:=False, CompatibilityMode:=14
'         .Item("Фототаблица_" & criateDocName(fn) & ".docm").Close
'    End With
''Call CreateTaskItem
'End Sub
''
'Public Function criateDocName(Optional strFN As String = "") As String
''Формирование имени рабочего документа
'    criateDocName = fn & "_" & Right(CStr(Year(Now)), 2)
'End Function
''
'Public Function criateArchDocName(Optional strFN As String = "") As String
''Формирование архивного имени документа
'        criateArchDocName = "012_" & Format(fn, "#00000") & "з_49_" & Right(CStr(Year(Now)), 2) & "_Набебин_Комар"
'End Function
''

'Private Static Sub CreateTaskItem()
'Создание задачи Outlook
'Dim newDate As clmCaseDate
'Set newDate = New clmCaseDate
'Set MyOlTask = MyOlApp.CreateItem(ItemType:=olTaskItem) 'CreateItem(ItemType:=olTaskItem)
'    With MyOlTask
'        .StartDate = ActiveDocument.FormFields("WdFirstDay").Result 'дата начала
'        .dueDate = newDate.getPeriod(newDate.ExamDate(.StartDate), 30) 'срок
'        .categories = ActiveDocument.FormFields("WdDocCat").Result 'категория
'        .Subject = criateDocName(fn) '& .Bookmarks("WdEF..потерпевший").Range.Text  'тема
'        'примечания
'
'        .Body = "NB!!! Окончить до " & newDate.getReminder(.StartDate, 30) & vbCr
        
        
        
'            If optClothes.Value = True Then
'                .Body = "Материалы " & strCsCat & " №" & strCsNum & Chr(10) & _
'                "Постановление: " & strCrPost & Chr(10) & _
'                strCrOprAr & Chr(10) & _
'                strCrRank & Chr(32) & strCrSurName & Chr(32) & strCrName & Chr(32) & strCrMidName & Chr(10) & _
'                Chr(10) & " Предоставили:" & Chr(10) & Create_EFReminder15(ByVal CrEFFirstDay)
'            Else: .Body = "Материалы " & strCsCat & " №" & strCsNum & Chr(10) & _
'                "Постановление: " & strCrPost & Chr(10) & _
'                strCrOprAr & Chr(10) & _
'                strCrRank & Chr(32) & strCrSurName & Chr(32) & strCrName & Chr(32) & strCrMidName & Chr(10) & _
'                Chr(10) & " Предоставили:" & Chr(10)
''            PrintArray1 i, N
'            End If




          'состояние
'        .Status = olTaskInProgress
'        .PercentComplete = 15 'процент выполнения
'        .Companies = "ГКСЭ РБ Управление медико-криминалистических экспертиз" 'организация
'        .BillingInformation = "Справка " & ActiveDocument.FormFields("WdEFNum").Result 'расходы
'        .importance = olImportanceNormal 'важность
'''напоминание
'        .ReminderSet = True
''        .ReminderTime = newDate.getReminder(.StartDate, 30)
'        .Save
'    End With
'MyOlTask.Display
'End Sub






'Private Static Sub CreateTaskItem(ByVal strDN As String, ByVal strDocName5 As String, ByVal CrEFFirstDay As Date)
''Создание задачи Outlook
'Set MyOlTask = MyOlApp.CreateItem(ItemType:=olTaskItem)
'    With MyOlTask
'        .Categories = strDocName5 'категория
'        .Subject = strDN & "/" & Year(Now) & " " & strInjPrSurName & " " & strInjPrName & " " & strInjPrMidName  'тема
'        .StartDate = CrEFFirstDay 'дата начала
'        .DueDate = Str(CrPeriod(ByVal CrEFFirstDay, ByVal dtDueDate))  'срок
'        'примечания
'            If optClothes.Value = True Then
'                .Body = "Материалы " & strCsCat & " №" & strCsNum & Chr(10) & _
'                "Постановление: " & strCrPost & Chr(10) & _
'                strCrOprAr & Chr(10) & _
'                strCrRank & Chr(32) & strCrSurName & Chr(32) & strCrName & Chr(32) & strCrMidName & Chr(10) & _
'                Chr(10) & " Предоставили:" & Chr(10) & Create_EFReminder15(ByVal CrEFFirstDay)
'            Else: .Body = "Материалы " & strCsCat & " №" & strCsNum & Chr(10) & _
'                "Постановление: " & strCrPost & Chr(10) & _
'                strCrOprAr & Chr(10) & _
'                strCrRank & Chr(32) & strCrSurName & Chr(32) & strCrName & Chr(32) & strCrMidName & Chr(10) & _
'                Chr(10) & " Предоставили:" & Chr(10)
''            PrintArray1 i, N
'            End If
'          'состояние
'        .Status = olTaskInProgress
'        .PercentComplete = 15 'процент выполнения
'        .Companies = "ГКСЭ РБ Управление медико-криминалистических экспертиз" 'организация
'        .BillingInformation = "Справка " & strDN & "_" & Year(Now) 'расходы
'        .Importance = olImportanceNormal 'важность
'''напоминание
''        .ReminderSet = True
''        .ReminderTime = CrEFFirstDay + 7
'        .Save
'    End With
'MyOlTask.Display
'End Sub

Private Static Sub Create_FotoList()
'Создание фототаблицы
'Открытие шаблона Fotolist
'With Application
'    Dim MyWdDoc As New Word.Document
'        MyWdDoc = .Documents.Add("D:\Crime\MasterForm\Word\FotoList_new.dotm")
'        MyWdDoc.Name = "Фототаблица" & criateDocName(FN)
'
''Сохранение фототаблицы:
''MyWdDoc.SaveAs2 FileName:=NewExpertFindings.DOC & "Фототаблица_" & NewExpertFindings.strEFNum _
''    & "_" & Right(Year(Now), 2), FileFormat:=wdFormatXMLDocument, _
''    LockComments:=False, Password:="", AddToRecentFiles:=True, WritePassword:="", _
'    ReadOnlyRecommended:=False, EmbedTrueTypeFonts:=False, SaveNativePictureFormat:=False, _
'    SaveFormsData:=False, SaveAsAOCELetter:=False








'    MyWdDoc.SaveAs2 FileName:=NewExpertFindings.DOC & "Фототаблица_" & NewExpertFindings.strEFNum _
'            & "_" & Right(Year(Now), 2), FileFormat:=wdFormatXMLDocument, _
'            LockComments:=False, Password:="", AddToRecentFiles:=True, WritePassword:="", _
'            ReadOnlyRecommended:=False, EmbedTrueTypeFonts:=False, SaveNativePictureFormat:=False, _
'            SaveFormsData:=False, SaveAsAOCELetter:=False
'    With MyWdApp.ActiveDocument  'Выбираем активный документ
'        On Error Resume Next ' Отключаем ошибки
''Работа с верхним колонтитулом:
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
'''работа с закладкой основного документа:
''i = 1
''    For i = 1 To n - 1 Step 1
''        .Bookmarks("WdEvColumnArr").Range.Text = "Фото № " & "." & Chr(32) & EvidArray(i) & Chr(10) & Chr(10)
''    Next i
''Завершение работы с документом
'        .Close SaveChanges:=wdSaveChanges
'End With
'Set MyWdDoc = Nothing
End Sub


'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'Public Static Sub Create_EF()
'Процедура создания нового документа Ecspert Findings
'    MyWdDoc.SaveAs2 FileName:=NewExpertFindings.DOC & "Заключение_" & NewExpertFindings.strEFNum _
'            & "_" & Right(Year(Now), 2), FileFormat:=wdFormatXMLDocument, _
'            LockComments:=False, Password:="", AddToRecentFiles:=True, WritePassword:="", _
'            ReadOnlyRecommended:=False, EmbedTrueTypeFonts:=False, SaveNativePictureFormat:=False, _
'            SaveFormsData:=False, SaveAsAOCELetter:=False
'    With MyWdApp.ActiveDocument  'Выбираем активный документ
'        On Error Resume Next
''Поля:
'       .FormFields("WdDirDOC").Result = NewExpertFindings.DOC
'       .FormFields("WdEFNum").Result = NewExpertFindings.strEFNum
'       .FormFields("WdFirstDay").Result = NewExpertFindings.DtmEFFirstDay
''Закладки:
'        .Bookmarks("WdEFfullNum").Range.Text = NewExpertFindings.Criate_FullEFNum 'Код экспертизы
'        .Bookmarks("WdExperience").Range.Text = NewExperts.Calculate_Experience 'расчет стажа работы
'        .Bookmarks("WdRsnComplain").Range.Text = Print_RsnComplain 'юр.основание
'        .Bookmarks("FactCase").Range.Text = Print_FactCase 'Обстоятельства дела
'    If chkAddExData.Value = 1 Then
'        .Bookmarks("WdAEDirection").Range.Text = Print_AEDirection 'Направление общего эксперта
'    End If
'    Dim tmp As String
'        tmp = Me.lblEFCount
'    If mdCount.colForms("DocListform" & tmp).chkAEFInquiry.Value = 1 Then
'        NewExpertFindings.DtmEFSuspDate = DateTime.Date
'        .Bookmarks("WdEFSuspDate").Range.Text = NewExpertFindings.DtmEFSuspDate
'        .Bookmarks("WdAEFInquiry").Range.Text = Print_AEFInquiry 'направленные ходатайства
'    End If
'
'Печать описания упаковок и т.д. вещественных доказательств
'        .Bookmarks("WdEvNumArr1").Range.Text = NewEvid.Print_EvNumArr1
'        .Bookmarks("WdEvNumArr2").Range.Text = NewEvid.Print_EvNumArr2
'        .Bookmarks("WdEvNumArr3").Range.Text = NewEvid.Print_EvNumArr3
''Печать ВД:
'    .Bookmarks("WdEvColumnArr").Range.Select
'            Call Print_NumColumn
''         For i = 1 To n - 1
''            .Bookmarks("WdEvColumnArr").Range.Text = EvidArray(i) & Chr(10)
''            .Bookmarks("WdEvLineArr").Range.Text = EvidArray(i) & ", "
''            .Bookmarks("WdEvLineArr1").Range.Text = EvidArray(i) & ", "
''            .Bookmarks("WdEvLineArr3").Range.Text = EvidArray(i) & ", "
''        Next i
''    On Error GoTo 999 ' Включаем обработку ошибки
'    .Close SaveChanges:=wdSaveChanges
'    End With
'Set MyWdDoc = Nothing
'999:
'    MsgBox Err.Description  'Ошибка
'    Err.Clear
'End Sub
'
'Private Static Sub Create_ReportEF()
'Работа с таблицей Excel
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
''открытие рабочей книги "отчет криминалистика"
'With Workbooks
'    .Open FileName:="E:\Crime\Reports\Отчет Криминалистика_" & Year(Now) & ".xlsm", Password:="фтв7004"
''Заполнние ячеек активного листа
''    ячейка "получено новых экспертиз"
'        Range("D10").Select
'        intNEF = ActiveCell.Value
'        intX = intNEF + 30
''    ячейка "порядковый номер" "A"
'        Range(CellArr(1) & intX).Select
'        ActiveCell.FormulaR1C1 = intX - 29
''    ячейка "номер экспертизы" "B"
'        Range(CellArr(2) & intX).Select
'        ActiveCell.FormulaR1C1 = CInt(NewExpertFindings.strEFNum)
''    ячейка "Дата начала" "C"
'        Range(CellArr(3) & intX).Select
'        ActiveCell.FormulaR1C1 = NewExpertFindings.DtmEFFirstDay
''    'ячейка "Материалы" "D"
'        Range(CellArr(4) & intX).Select
'        ActiveCell.FormulaR1C1 = NewCase.strCsCat
''    ячейка "№ дела" "G"
'        Range(CellArr(7) & intX).Select
'        ActiveCell.FormulaR1C1 = NewCase.strCsNum
''    ячейка "Объекты" "I"
''        Range(CellArr(9) & intX).Select
''        ActiveCell.FormulaR1C1 = i
''   ячейка "приостановлена" "J"
'        Range(CellArr(10) & intX).Select
'''            If frmDocList.chkAEFInquiry = 1 Then
'''                NewExpertFindings.DtmEFSuspDate = DateTime.Date
'''                ActiveCell.FormulaR1C1 = NewExpertFindings.DtmEFSuspDate
'''            End If
''   ячейка "срок 15" "N"
''        Range(CellArr(14) & intX).Select
''            If frmMDI.strCrEFCat = "Одежда" Then
''              intTmpDt = 15
''              ActiveCell.FormulaR1C1 = NewExpertFindings.Calculate_EFPeriod(ByVal intTmpDt)
''            End If
''   ячейка "срок" "P"
'        Range(CellArr(16) & intX).Select
'        intTmpDt = 30
'        ActiveCell.FormulaR1C1 = NewExpertFindings.Calculate_EFPeriod(ByVal intTmpDt)
''   итоговое изменение данных в ячейке "получено новых экспертиз" (т.е. +1)
'            Range("D10").Select
'            ActiveCell.FormulaR1C1 = intNEF + 1
''   списание спирта "S"
'            Range(CellArr(19) & intX).Select
'            ActiveCell.FormulaR1C1 = NewExpertFindings.DtmEFFirstDay
''  расчет остатка спирта
'        Range(CellArr(20) & intX - 1).Select
'        curAlc = ActiveCell.Value 'запоминание предыдущего остатка спирта
'        Range(CellArr(20) & intX).Select
'        ActiveCell.FormulaR1C1 = curAlc - 0.1
'    End With
'ActiveWorkbook.Save
'Workbooks.Close
'MyExApp.Quit
'End Sub
''




