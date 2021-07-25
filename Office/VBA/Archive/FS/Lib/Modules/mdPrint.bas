Attribute VB_Name = "mdPrintDoc"
'модуль "mdPrint" для формирования строковых данных и печать их в документы
'Дата создания: 01.06.2016
'@version 0.0.1
'@author Andr.Nab.n@gmail.com
Option Explicit
'!!вставить по умолчанию ссылку на папку проекта
Const DOT = "D:\Crime\Soft_поРаботе\VBA\1_6\Resources\doc" 'Дом разработка" 'директория шаблонов (по умолчанию)
Const doc = "D:" 'директория готовых документов (по умолчанию).

Const CRITICAL As String = "Внимание, ошибка!"
Private Const BOX As String = "BX"  'префикс ключа для коробок
Private Const EVID As String = "EV" 'префикс ключа для вещдоков.
'
Public tmpBx As clmEvBox
Private bxKey  As Variant
'
Public colDocCat As Collection 'коллекция "категория документа"
Public colTextDoc As Collection ' коллекция создаваемых абзацев,вставляемых в документы
Public colExcelData As Collection 'коллекция данных для записи в книгу Excel
Public strMonth As String            'текущий месяц
'
Private mvarDocCat As String   'переменная "категория документа"
Private mvarstrDOT As String   'имя шаблона документа
'
Private mvarstrDOC As String   'имя созданного документа
'Динамические массивы:
Public arrEF() As String             'массив названия экспертиз (входящих в комплекс)
Public arrExpert() As String         'массив имен экспертов
'
'раннее связывание приложений "MsOffice"
Public MyWdApp As New Word.Application      'экземпляр приложения MsOffice Word
Public MyWdDoc As Word.Document            'экземпляр документа MsOffice Word
Public myExApp As Excel.Application        'экземпляр приложения MsOffice Excel
Public myExDoc As Excel.Workbook           'экземпляр документа MsOffice Excel
'Public MyOlApp As Outlook.Application      'экземпляр приложения MsOffice Outlook
'Public MyOlTask As Outlook.TaskItem        'экземпляр задачи MsOffice Outlook
Private fs As New FileSystemObject         'Объявление нового объекта (папки)
Private fll As Variant
Private SvcService As Object    'объект библиотеки svcsvc.dll
Public monthNow As String       'текущий месяц (Январь.../Декабрь)для работы с Excel
'
'========================== I N C A P S U L A T I O N ==================================
'
Public Property Let DocCat(ByVal vData As String)
'Переменная Категогия документов.
    mvarDocCat = vData
End Property
'
Public Property Get DocCat() As String
'Переменная Категогия документов.
    DocCat = mvarDocCat
'Debug.Print "Переменная Категогия документов = ", DocCategory
End Property
'
Public Property Let strDOC(ByVal vData As String)
'Переменная Шаблон документа.
    mvarstrDOC = vData
End Property
'
Public Property Get strDOC() As String
'Переменная Шаблон документа.
    strDOC = mvarstrDOC
'Debug.Print "Шаблон документа = ", strDOC
End Property
'
Public Property Let strDOT(ByVal vData As String)
'Переменная Шаблон документа.
    mvarstrDOT = vData
End Property
'
Public Property Get strDOT() As String
'Переменная Шаблон документа.
    strDOT = mvarstrDOT
'Debug.Print "Шаблон документа = ", strDOC
End Property
'
Public Function puckEVup(tmpStamp As String, lngEvCount As Long) As String
'строка - "вещественное/ые  доказательство/ва упаковано/ны, опечатано/ны" + Наша печать
   Dim str1 As String, str2 As String, str3 As String
    If lngEvCount > 0 Then
        If lngEvCount = 1 Then
                str1 = "Вещественное доказательство "
            str2 = "упаковано, "
            str3 = "опечатано "
        Else:
        str1 = "Вещественные доказательства "
         str2 = "упакованы, "
            str3 = "опечатаны "
        End If
puckEVup = str1 & str2 & str3 & "мастичным оттиском синего цвета круглой печати " & tmpStamp
    End If
End Function
'
'====================================
Public Static Sub withApplWD(dirDOT As String, dirDOC As String)
'процедура для работы с приложением msWord
'String dirDOT = папка с шаблоном создаваемого документа
'String dirDOC = папка для сохраннения созданного документа
 Set SvcService = CreateObject("Svcsvc.Service") 'объект библиотеки  svcsvc.dll
 Set colTextDoc = New Collection 'обнуление коллекции для сохраннения текстов абзацев
 
'===================доделать для ускорения работы программы=======================
'Заполнение коллекции colTextDoc!!!
'    Dim str1 As String, str2 As String, str3 As String
'            str 1 = frmNewEF.newEF.getLegalGround(frmNewEF.newCor.print_Cor, frmNewEF.newCase.printCsData)
'        colTextDoc.Add str1, "legalGround"
''       направление эксперта:
'        If frmNewEF.chkAddExData.Value = 1 Then
'            str2 = frmNewEF.newAExperts.print_Expert
'                colTextDoc.Add str2, "Print_Expert"
'        str3 = frmNewEF.newAEF.print_AEdoc(frmNewEF.newAExperts.print_Expert, frmNewEF.newInjPr.print_Autopsy)
'            colTextDoc.Add str2, "print_AEdoc"
'        End If
'===================доделать для ускорения работы программы=======================

    Set MyWdApp = New Word.Application
        With MyWdApp
            .Visible = False 'видимость приложения во время выполения программы
''           1)проверка на существование папки с шаблонами и папки для сохраннения документов
'                If dirDOT = "" Then
'                    MsgBox "Папка с шаблонами не указана!", vbExclamation, CRITICAL
'                    dirDOT = SvcService.SelectFolder("Укажите пакку с шаблонами!", "", &H10 + &H4000, "")
'            ElseIf dirDOC = "" Then
'                MsgBox "Папка для сохраннения не указана!", vbExclamation, CRITICAL
'                    dirDOC = SvcService.SelectFolder("Укажите пакку для сохраннения!", "", &H10 + &H4000, "")
'                End If
'          печать документов:
'           1) заключение эксперта
            Call printEF(mdMainFolders.arrDocDir(1), mdMainFolders.arrDocDir(0), DocCat)
'           2)фототаблица
            Call printFotoList(mdMainFolders.arrDocDir(1))
'           3)Этикетки
            Call print_Labels(mdMainFolders.arrDocDir(1))
'           4)сопроводительные листы и описание ВД для биологов
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
            .Quit 'закрытие приложения
        End With
Set MyWdApp = Nothing
Set SvcService = Nothing   'закрытие объекта библиотеки  svcsvc.dll
End Sub
'
Public Static Sub printEF(dirDOT As String, dirDOC As String, ByVal DocCat As String)
'создание документа "Заключение эксперта"
'String dirDOT = папка с шаблоном создаваемого документа
'String dirDOC = папка для сохраннения созданного документа
'String tmpDocCat = категория создаваемого документа (одежда/препараты кожи и т.д.)
'    Debug.Print "папка с шаблоном = " & dirDOT & "\" & strDOT
    Set MyWdDoc = MyWdApp.Documents.Open(dirDOT & "\" & strDOT)    'открытие шаблона
    MyWdDoc.Activate ' активируем шаблон
    With ActiveDocument
'   1)  заполнение информационных полей документа
'           Если больше одного эксперта:
'            Dim i As Long, str As String
'                For i = LBound(arrExpert) To UBound(arrExpert)
'                    If i = UBound(arrExpert) Then
'                        str = str & frmNewEF.newExperts.create_revExpert(arrExpert(i)) '& Chr(10)
'                    Else
'                        str = str & frmNewEF.newExperts.create_revExpert(arrExpert(i)) & vbCrLf '& Chr(10)
'                    End If
'                Next i
'                ActiveDocument.Bookmarks("experts").Range.Text = str
'                Debug.Print "Список экспертов - " & str
        If .Bookmarks.Exists("WdArchDocName") = True Then
            .FormFields("WdArchDocName").Result = frmNewEF.newEF.getArchName(frmNewEF.newEF.number, "Набебин")
        End If
        If .Bookmarks.Exists("WdDirDOC") = True Then
            .FormFields("WdDirDOC").Result = mdMainFolders.arrDocDir(2)
        End If
        If .Bookmarks.Exists("WdEFNum") = True Then
            .FormFields("WdEFNum").Result = frmNewEF.newEF.number
        End If
'   2)сохранение документа под необходимым именем
        .SaveAs2 FileName:=mdMainFolders.arrDocDir(2) & "\Заключение_" & frmNewEF.newEF.getNum_Cat, FileFormat:= _
            wdFormatXMLDocumentMacroEnabled, LockComments:=False, Password:="", _
            AddToRecentFiles:=True, WritePassword:="", ReadOnlyRecommended:=False, _
            EmbedTrueTypeFonts:=False, SaveNativePictureFormat:=False, SaveFormsData _
            :=False, SaveAsAOCELetter:=False
'    3)работа с полями документа:
'       2.1 полный номер экспертизы:
        If .Bookmarks.Exists("WdEFfullNum") = True Then 'проверка на существование закладки:
            .Bookmarks("WdEFfullNum").Range.Text = frmNewEF.newEF.getFullNumber
        End If
'       2.2 дата начала экспертизы
        .FormFields("WdFirstDay").Result = frmNewEF.newEF.firstDay
'       2.3 стаж работы по специальности
    If .Bookmarks.Exists("getExperience") = True Then 'проверка на существование закладки:
        .Bookmarks("getExperience").Range.Text = frmNewEF.newDate.getExperience(frmNewEF.newDate.experience, frmNewEF.newDate.DateNow)
    End If
'       2.4 юридическое основание:
    If .Bookmarks.Exists("legalGround") = True Then 'legalGround
        Dim lg1 As String, lg2 As String, lg3 As String, lg4 As String 'формирование строки:
            lg3 = ", провел судебную медико-криминалистическую экспертизу."
        With frmNewEF
            Dim str1 As String, str2 As String
                str1 = .newEF.getLegalGround(.newCor.print_Cor, .newCase.printCsData)
            colTextDoc.Add str1, "legalGround"
            lg1 = "на " & str1 'lg1 = "на " & colTextDoc("legalGround")
            If frmNewEF.chkAddExData.Value = 1 Then
                str2 = .newAExperts.print_Expert
                colTextDoc.Add str2, "print_AEdoc"
                lg2 = "направления" & str2
                lg4 = .newAEF.print_AEdoc(lg2, .newInjPr.print_Autopsy)
                If ActiveDocument.Bookmarks.Exists("print_AEdoc") = True Then
                    ActiveDocument.Bookmarks("print_AEdoc").Range.Text = lg4
                End If
                ActiveDocument.Bookmarks("legalGround").Range.Text = lg1 & "и " & lg2 & lg3
            Else
                ActiveDocument.Bookmarks("legalGround").Range.Text = lg1 & lg3
            End If
        End With 'frmNewEF
    End If 'legalGround
    
    Dim tmpBx As clmEvBox
    Dim x As Long, bxKey As String, tmpStr1 As String
'       2.5 печать предоставленных ВД в одну строку.
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
' Коробки
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
Set MyWdDoc = MyWdApp.Documents.Open(dirDOT & "\" & "Description.docm")    'открытие шаблона
    MyWdDoc.Activate ' активируем шаблон
    With ActiveDocument
'   сохранение документа под необходимым именем
        .SaveAs2 FileName:=mdMainFolders.arrDocDir(2) & "\Описание_" & frmNewEF.newEF.getNumber, FileFormat:= _
            wdFormatXMLDocumentMacroEnabled, LockComments:=False, Password:="", _
            AddToRecentFiles:=True, WritePassword:="", ReadOnlyRecommended:=False, _
            EmbedTrueTypeFonts:=False, SaveNativePictureFormat:=False, SaveFormsData _
            :=False, SaveAsAOCELetter:=False
''       работа с полями документа:
        If .Bookmarks.Exists("InfoEF") = True Then
            ActiveDocument.Bookmarks("InfoEF").Range.Text = frmNewEF.newEF.toString_InfoEF(colTextDoc("legalGround"), frmNewEF.newExpert) & " (тел/факс (017) 308-67-53)." & Chr(10)
        End If
'   коробки с ВД
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
Set MyWdDoc = MyWdApp.Documents.Open(dirDOT & "\" & "Hodataistvo.docm")    'открытие шаблона
    MyWdDoc.Activate ' активируем шаблон
    With ActiveDocument
'   сохранение документа под необходимым именем
        .SaveAs2 FileName:=mdMainFolders.arrDocDir(2) & "\Ходатайство_" & frmNewEF.newEF.getNumber, FileFormat:= _
            wdFormatXMLDocumentMacroEnabled, LockComments:=False, Password:="", _
            AddToRecentFiles:=True, WritePassword:="", ReadOnlyRecommended:=False, _
            EmbedTrueTypeFonts:=False, SaveNativePictureFormat:=False, SaveFormsData _
            :=False, SaveAsAOCELetter:=False
'    работа с полями документа:
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
'dirDOT - директория шаблона
'dotNam - имя шаблона (Напр: "Biology")
'docNam - имя документа, под которым он будет сохранен (Напр.: "Сопровод_Биология_№№№_YY")
If dotName <> "" Or docName <> "" Then
    Set MyWdDoc = MyWdApp.Documents.Open(dirDOT & "\" & dotName & ".docm")    'открытие шаблона
    MyWdDoc.Activate ' активируем шаблон
    With ActiveDocument
'   сохранение документа под необходимым именем
        .SaveAs2 FileName:=mdMainFolders.arrDocDir(6) & "\" & frmNewEF.newEF.getNumber & "_" & docName & ".docm", FileFormat:= _
            wdFormatXMLDocumentMacroEnabled, LockComments:=False, Password:="", _
            AddToRecentFiles:=True, WritePassword:="", ReadOnlyRecommended:=False, _
            EmbedTrueTypeFonts:=False, SaveNativePictureFormat:=False, SaveFormsData _
            :=False, SaveAsAOCELetter:=False
'       работа с полями документа:
        If .Bookmarks.Exists("InfoEF") = True Then
            ActiveDocument.Bookmarks("InfoEF").Range.Text = colTextDoc("legalGround")
        End If
'       эксперты:
'       !!! добавить обработку массива (если выбрано более одноо эксперта)
        If .Bookmarks.Exists("experts") = True Then
             ActiveDocument.Bookmarks("experts").Range.Text = frmNewEF.newExpert
        End If
'       коробки с ВД
     Dim tmpStr As String, tmpStr1 As String, x As Long, bxKey As String
            With frmNewEF.colBoxes
                If .Count > 0 Then
                    If .Count = 1 Then
                        tmpStr = "объект"
                    Else: tmpStr = "объекты"
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
'печать акта несоответствия
'dirDOT - директория шаблона
'dotNam - имя шаблона (Напр: "Biology")
'docNam - имя документа, под которым он будет сохранен (Напр.: "Сопровод_Биология_№№№_YY")
If dotName <> "" Or docName <> "" Then
    Set MyWdDoc = MyWdApp.Documents.Open(dirDOT & "\" & dotName & ".docm")    'открытие шаблона
    MyWdDoc.Activate ' активируем шаблон
    With ActiveDocument
'   сохранение документа под необходимым именем
        .SaveAs2 FileName:=mdMainFolders.arrDocDir(2) & "\" & frmNewEF.newEF.getNumber & "_" & docName & ".docm", FileFormat:= _
            wdFormatXMLDocumentMacroEnabled, LockComments:=False, Password:="", _
            AddToRecentFiles:=True, WritePassword:="", ReadOnlyRecommended:=False, _
            EmbedTrueTypeFonts:=False, SaveNativePictureFormat:=False, SaveFormsData _
            :=False, SaveAsAOCELetter:=False
'       работа с полями документа:
'       2.1 полный номер экспертизы:
        If .Bookmarks.Exists("WdEFfullNum") = True Then 'проверка на существование закладки:
            .FormFields("WdEFfullNum").Result = frmNewEF.newEF.getFullNumber
        End If
'       2.2 дата начала экспертизы
        If .Bookmarks.Exists("WdFirstDay") = True Then 'проверка на существование закладки:
             .Bookmarks("WdFirstDay").Range.Text = frmNewEF.newEF.firstDay
        End If
'       2.3 дата поступления ВД
        If .Bookmarks.Exists("wdBxFirstDate") = True Then 'проверка на существование закладки:
             .Bookmarks("wdBxFirstDate").Range.Text = frmNewEF.newBox.DtmBxFirstDate
        End If
'       2.4 материалы дела
        If .Bookmarks.Exists("wdCsData") = True Then 'проверка на существование закладки:
            .FormFields("wdCsData").Result = frmNewEF.newCase.printCsData
        End If
'       2.5 юридическое основание
         If .Bookmarks.Exists("InfoEF") = True Then
            ActiveDocument.Bookmarks("InfoEF").Range.Text = colTextDoc("legalGround")
        End If
'       2.6 эксперт
        If .Bookmarks.Exists("experts") = True Then
             ActiveDocument.Bookmarks("experts").Range.Text = frmNewEF.newExpert
        End If
'       2.7 следователь
        If .Bookmarks.Exists("wdCoroner") = True Then
            ActiveDocument.Bookmarks("wdCoroner").Range.Text = frmNewEF.newCor.print_Cor
        End If
'       2.6коробки с ВД
     Dim tmpStr As String, tmpStr1 As String, x As Long, bxKey As String
            With frmNewEF.colBoxes
                If .Count > 0 Then
                    If .Count = 1 Then
                        tmpStr = "объект"
                    Else: tmpStr = "объекты"
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
Public Static Sub printFotoList(dirDOT As String, Optional dotName As String = "FotoList", Optional docName As String = "Фототаблица")
'Создание и заполнение фототаблицы
'dirDOT - директория шаблона
'dotNam - имя шаблона (Напр: "Biology")
'docNam - имя документа, под которым он будет сохранен (Напр.: "Сопровод_Биология_№№№_YY")
Set MyWdDoc = MyWdApp.Documents.Open(dirDOT & "\" & dotName & ".docm")    'открытие шаблона
    MyWdDoc.Activate ' активируем шаблон
    With ActiveDocument
'   сохранение документа под необходимым именем
        .SaveAs2 FileName:=mdMainFolders.arrDocDir(3) & "\" & docName & "_" & frmNewEF.newEF.getNumber, FileFormat:= _
            wdFormatXMLDocumentMacroEnabled, LockComments:=False, Password:="", _
            AddToRecentFiles:=True, WritePassword:="", ReadOnlyRecommended:=False, _
            EmbedTrueTypeFonts:=False, SaveNativePictureFormat:=False, SaveFormsData _
            :=False, SaveAsAOCELetter:=False
'       работа с полями документа:
        If .Bookmarks.Exists("WdEFfullNum") = True Then
            ActiveDocument.Bookmarks("WdEFfullNum").Range.Text = frmNewEF.newEF.getFullNumber
        End If
'       эксперты:
'     обработка массива (если выбрано более одноо эксперта)
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
                Debug.Print "Список экспертов - " & str
        End If
'           список коробок
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
                                Optional docName As String = "Этикетки")
Set MyWdDoc = MyWdApp.Documents.Open(dirDOT & "\" & dotName & ".docm")    'открытие шаблона
    MyWdDoc.Activate ' активируем шаблон
    With ActiveDocument
'   сохранение документа под необходимым именем
        .SaveAs2 FileName:=mdMainFolders.arrDocDir(4) & "\" & docName & "_" & frmNewEF.newEF.getNumber, FileFormat:= _
            wdFormatXMLDocumentMacroEnabled, LockComments:=False, Password:="", _
            AddToRecentFiles:=True, WritePassword:="", ReadOnlyRecommended:=False, _
            EmbedTrueTypeFonts:=False, SaveNativePictureFormat:=False, SaveFormsData _
            :=False, SaveAsAOCELetter:=False
'       работа с полями документа:
        If .Bookmarks.Exists("wdLabels") = True Then
            Dim tmpBx As clmEvBox
            Dim x As Long, bxKey As String, tmpStr As String
            Dim arrStr(4) As String
                arrStr(0) = frmNewEF.newEF.getFullNumber 'полный номер документа
                arrStr(1) = frmNewEF.newCase.printCsData & Chr(10) 'данные из постановления
                arrStr(2) = frmNewEF.newEF.firstDay 'дата начала экспертизы
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
'процедура открытие msExcel
 Dim n As Long 'счетчик
 Dim numRow, tmpRow As Integer 'значение номера строки для записи данных
 Dim arrStr(8) As String
    arrStr(0) = "Текущий месяц -> " & strMonth
    arrStr(1) = "A"
    arrStr(2) = "B"
    arrStr(3) = "C"
    arrStr(4) = "D"
    arrStr(5) = "E"
    arrStr(6) = "G"
    arrStr(7) = "j"
'   работа с Excel
    Set myExApp = New Excel.Application
        myExApp.Visible = False
    Dim tpmDocName As String
        tpmDocName = mdMainFolders.arrDocDir(0) & "\Отчет_Криминалистика_ " & Year(Now) & ".xlsm"
'   открытие документа Excel
    Workbooks.Open FileName:=tpmDocName ', Password:="фтв7004"
'   открытие листа с именем текущего месяца
        Worksheets.Item(strMonth).Activate
        With ActiveWorkbook.Sheets
'            .Item(monthNow).Activate ' активация листа с именем текущего месяца
            With .Application
'               получаем значение количества экспертиз:
                tmpRow = .Cells(3, 3).Value
                numRow = tmpRow + 13
                For n = 1 To 7
                    .Cells(numRow, arrStr(n)).Activate   'активируем ячейку для записи данных:
                    .ActiveCell.Value = mdPrintDoc.colExcelData.Item(arrStr(n)) 'запись данных из массива в ячейки:
                Next n
            End With
        End With
        ActiveWorkbook.Save
    myExApp.Quit
    Set myExApp = Nothing
End Sub
    
    
Public Sub closeEx()
'закрытие приложения  msExcel

'       .Save

Set myExDoc = Nothing
Set myExApp = Nothing
End Sub

'Private Static Sub Create_ReportEF()
''Работа с таблицей Excel
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
'        ActiveCell.FormulaR1C1 = newCase.strCsCat
''    ячейка "№ дела" "G"
'        Range(CellArr(7) & intX).Select
'        ActiveCell.FormulaR1C1 = newCase.strCsNum
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


