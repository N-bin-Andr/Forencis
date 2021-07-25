Attribute VB_Name = "mdDocCreate"
Option Explicit

'Класс для создания папок и печати документов.
'@author Andr.N@_bin
'@E-mail Andr.Nab.n@gmail.com
'
'объявление системных переменных
Private fs As FileSystemObject  'экземпляр  FileSystemObject
Private SvcService As Object    'объект библиотеки svcsvc.dll
'Public ourStamp As String       'наша печать для упаковки объектов
'Word
Public MyWdApp As Word.Application  'экземпляр приложения
Public DocWord As Word.Document     ' экземпляр документа
'Ошибки
Const CRITICAL As String = "Внимание, ошибка!"

'+++++++++++++++++++++++++ М Е Т О Д Ы ++++++++++++++++++++++++++++

Public Function Print_OurStamp(lngEvCount As Long) As String
'строка="вещественное/ые доказательство/а упаковано/ы, опечатано/ы мастичным оттиском синего цвета
'           круглой печати + "наша печать"
    Dim str As String
    If lngEvCount > 0 Then
        If lngEvCount = 1 Then
            str = "вещественное доказательство упаковано, опечатано "
        Else: str = "вещественные доказательства упакованы, опечатаны "
        End If
    End If
Dim strTx As String
    strTx = "мастичным оттиском синего цвета к руглой печати "
Print_OurStamp = str & strTx & Chr(171) & STAMP & Chr(187)
'OurStamp = "Print_OurStamp " & Print_OurStamp
End Function
'
Public Function deliveryEvToString(lngEvCount As Long) As String
'Функция печати доставки упаковки/ок
'строка = "Вещественное/ые доказательство/а доставлено/ы нарочным, упакованное/ые в + упаковка + печать"
    Dim str As String
    If lngEvCount > 0 Then
        If lngEvCount = 1 Then
            str = "Вещественное доказательство доставлено нарочным, упакованное в "
        Else: str = "Вещественные доказательства доставлены нарочным, упакованные в "
        End If
    End If
    
    Dim strTx1 As String
    If StrComp(strEvPackage, "картонную коробку") = True Then
        strTx1 = "опечатанную "
    Else: strTx1 = "опечатанный "
    Debug.Print "strTx1 = "
    End If
deliveryEvToString = str & strTx1 & strEvPackage & "(" & "Фото №№" & ")."
Debug.Print "deliveryEvToString = " & deliveryEvToString
End Function
'
Public Function boxIntegralityToString(lngEvCount As Long) As String
'целостность упаковки
    Dim str As String, strTx1 As String
    If lngEvCount > 0 Then
        If lngEvCount = 1 Then
            str = "предоставленного объекта "
            strTx1 = "был извлечен "
            
        Else
            str = "предоставленных объектов "
            strTx1 = "были извлечены: "
        End If
    End If
boxIntegralityToString = "Целостность упаковки не нарушена, извлечение" & str & _
                        "без повреждения целостности упаковки невозможно. При вскрытии упаковки из нее " & strTx1
Debug.Print "boxIntegralityToString = " & boxIntegralityToString
End Function
'
Public Function accordanceEvObjToString(lngEvCount As Long) As String
'строка= "объект/ы, предоставленный/ые для медико-криминалистического исследования, соответствует/ют "
'       перечную, указанному в направлении и в сопроводительной надписи к вещественным доказательствам. "
 Dim str As String, strTx1 As String
    If lngEvCount > 0 Then
        If lngEvCount = 1 Then
            str = "Объект, предоставленный на исследование, "
            strTx1 = "соответствует "
            
        Else
            str = "Объекты, предоставленные на исследование, "
            strTx1 = "сответствуют "
        End If
    End If
accordanceEvObjToString = str & strTx1 & "перечную, указанному в направлении и в сопроводительной надписи к вещественным доказательствам." & Chr(10)
Debug.Print "accordanceEvObjToString = " & accordanceEvObjToString
End Function
'
'Public Property Get ourStamp() As String
'
'
'
'End Property


Public Static Sub printEF(nameDOT As String, nameDOC As String, tmpDir As String)
'Создание и заполнение заключения эксперта
 Set SvcService = CreateObject("Svcsvc.Service") 'объект библиотеки  svcsvc.dll
 Set MyWdApp = New Word.Application
    With MyWdApp
        .Visible = True  'False скрываем  документ
        '1)Проверка указана ли папка с шаблонами
        If nameDOT = "" Then
           MsgBox "Папка c шаблонами документов не указана!", vbExclamation, CRITICAL
           nameDOT = SvcService.SelectFolder("Укажите папку с шаблонами документов!", "", &H10 + &H4000, "")
        End If
'       2)открытие шаблона
        Set DocWord = .Documents.Open(nameDOT) 'Открываем шаблон"
'       активируем его
        DocWord.Activate
'       Сохранение нового документа под необходимым именем
        With DocWord  'Выбираем активный документ
'       работа с полями документов:
'       Упаковки с вещдоками
'       проверка на существование заклаюки
        If ActiveDocument.Bookmarks.Exists("temp") = True Then
            ActiveDocument.Bookmarks("temp").Select
'       frmEvidences.colEvidences(index)

            
        End If

'
'
'



'       Этикетки - lb(lable)
'                .FormFields("lb_InjPrIniName").Result = frmNewResearch.newInjPr.create_InitialslName
'                .FormFields("lb_fNumberEF").Result = frmNewResearch.newEF.getFullNumber 'fNumberEF
'                .FormFields("lb_EFFirstDay").Result = frmNewResearch.newEF.firstDay 'firstDate
'                .FormFields("lb_AutopsyDate").Result = frmNewResearch.newInjPr.autopsyDate 'autopsyDate
'                .FormFields("lb_expertName").Result = frmNewResearch.newExpert 'expertName
'       Бланки - fm (Form)
'                .FormFields("fm_InjPrFullName").Result = frmNewResearch.newInjPr.create_FullName 'injPrFullName
'                .FormFields("fm_InjPrBirthday").Result = frmNewResearch.newInjPr.birthday 'injPrBirthday
'                .FormFields("fm_InjPrSex").Result = frmNewResearch.newInjPr.sex 'injPrSex
'                .FormFields("fm_deceaseDate").Result = frmNewResearch.newInjPr.decease 'deceaseDate
'                .FormFields("fm_autopsyDate").Result = frmNewResearch.newInjPr.autopsyDate 'autopsy
'                .FormFields("fm_fNumberEF").Result = frmNewResearch.newEF.getFullNumber 'fNumberEF
'                .FormFields("fm_factCase").Result = frmNewResearch.txtFactCase.Text 'factCase
'                .FormFields("fm_legalGround").Result = frmNewResearch.create_legalGround 'legalGround
'                .FormFields("fm_expertName").Result = reverse_Name(frmNewResearch.newExpert)
'      Сопроводительная записка  nt (note)
'                .FormFields("nt_EFFirstDay").Result = frmNewResearch.newEF.firstDay 'firstDate
'                .FormFields("nt_fNumberEF").Result = frmNewResearch.newEF.getFullNumber 'fNumberEF
'                .FormFields("nt_autopsy").Result = frmNewResearch.newInjPr.print_Autopsy 'autopsy info
'                .FormFields("nt_stringEF").Result = frmNewResearch.newEF.toStringEF 'stringEF
'                .FormFields("nt_expertName").Result = reverse_Name(frmNewResearch.newExpert)

'            Сохранение:
'            .SaveAs FileName:=tmpdir & nameDOC, FileFormat:= _
''                wdFormatXMLDocumentMacroEnabled, LockComments:=False, Password:="", _
''                AddToRecentFiles:=True, WritePassword:="", ReadOnlyRecommended:=False, _
''                EmbedTrueTypeFonts:=False, SaveNativePictureFormat:=False, SaveFormsData _
''                :=False, SaveAsAOCELetter:=False
'            'Закрытие
            .Close
        End With
    End With
Set SvcService = Nothing    'объект библиотеки  svcsvc.dll
'уничтожаем обьект - документ
Set DocWord = Nothing
'уничтожаем объект приложения
MyWdApp.Quit
Set MyWdApp = Nothing
End Sub
































'Public Static Sub print_Blank(nameDOT As String, nameDOC As String, tmpdir As String)
''Заполнение бланков направлений
' Set SvcService = CreateObject("Svcsvc.Service") 'объект библиотеки  svcsvc.dll
'' Set MyWdApp = New Word.Application
'    With MyWdApp
'        .Visible = False 'True 'скрываем  документ
'        '1)Проверка указана ли папка с шаблонами
'        If DOTDIR = "" Then
'           MsgBox "Папка c шаблонами документов не указана!", vbExclamation, CRITICAL
'           DOTDIR = SvcService.SelectFolder("Укажите папку с шаблонами документов!", "", &H10 + &H4000, "")
'        End If
'        '2)открытие шаблона
'        Dim tmpDot As String
'            tmpDot = DOTDIR & nameDOT & ".dotm" 'формирование имени шаблона (с директорией)
'        'открываем шаблон документа
'        Set DocWord = .Documents.Open(tmpDot)  'Открываем шаблон"
'        'активируем его
'        DocWord.Activate
'        'Сохранение нового документа под необходимым именем
'        With DocWord  'Выбираем активный документ
'            'работа с полями документов:
'            If nameDOT = "Линейка" Then
'                .FormFields("WdsNumberEF").Result = frmNewResearch.newEF.number    'sNumberEF
'                .FormFields("WdyearEF").Result = Right(frmNewResearch.newInjPr.autopsyDate, 4)  'Right(autopsyDate, 4)
'                'Печать активного документа
'                If MsgBox("Напечатать " & nameDOC & " ?", vbYesNo) = vbYes Then
'                    .PrintOut Range:=wdPrintAllDocument, item:= _
'                        wdPrintDocumentContent, Copies:=1, Pages:="", PageType:=wdPrintAllPages, _
'                        ManualDuplexPrint:=False, Collate:=False, Background:=True, PrintToFile:= _
'                        False, PrintZoomColumn:=0, PrintZoomRow:=0, PrintZoomPaperWidth:=0, _
'                        PrintZoomPaperHeight:=0
'
'                End If
'            ElseIf nameDOT = "Схема Голова" Or _
'                                nameDOT = "Схема ребер" Or _
'                                nameDOT = "Ростовая схема" Or _
'                                nameDOT = "Фототаблица" Or _
'                                nameDOT = "План проведения СМЭ" Then
'                .FormFields("WdstringEF").Result = frmNewResearch.newEF.toStringEF
'                .FormFields("WdexpertName").Result = reverse_Name(frmNewResearch.newExpert)
'                 'Печать активного документа
'                If MsgBox("Напечатать " & nameDOC & " ?", vbYesNo) = vbYes Then
'                    .PrintOut Range:=wdPrintAllDocument, item:= _
'                        wdPrintDocumentContent, Copies:=1, Pages:="", PageType:=wdPrintAllPages, _
'                        ManualDuplexPrint:=False, Collate:=False, Background:=True, PrintToFile:= _
'                        False, PrintZoomColumn:=0, PrintZoomRow:=0, PrintZoomPaperWidth:=0, _
'                        PrintZoomPaperHeight:=0
'                End If
'            ElseIf nameDOT = "Черновик" Then
'                .FormFields("WdexpertName").Result = frmNewResearch.newExpert 'expertName
'                .FormFields("WdsNumberEF").Result = frmNewResearch.newEF.getNumber 'sNumberEF
'                .FormFields("Wdautopsy").Result = frmNewResearch.newInjPr.print_Autopsy 'autopsy
'                 'Печать активного документа
'                If MsgBox("Напечатать " & nameDOC & " ?", vbYesNo) = vbYes Then
'                    .PrintOut Range:=wdPrintAllDocument, item:= _
'                        wdPrintDocumentContent, Copies:=1, Pages:="", PageType:=wdPrintAllPages, _
'                        ManualDuplexPrint:=True, Collate:=False, Background:=True, PrintToFile:= _
'                        False, PrintZoomColumn:=0, PrintZoomRow:=0, PrintZoomPaperWidth:=0, _
'                        PrintZoomPaperHeight:=0
'                End If
'            Else
'                'Этикетки - lb(lable)
'                .FormFields("lb_InjPrIniName").Result = frmNewResearch.newInjPr.create_InitialslName
'                .FormFields("lb_fNumberEF").Result = frmNewResearch.newEF.getFullNumber 'fNumberEF
'                .FormFields("lb_EFFirstDay").Result = frmNewResearch.newEF.firstDay 'firstDate
'                .FormFields("lb_AutopsyDate").Result = frmNewResearch.newInjPr.autopsyDate 'autopsyDate
'                .FormFields("lb_expertName").Result = frmNewResearch.newExpert 'expertName
'                'Бланки - fm (Form)
'                .FormFields("fm_InjPrFullName").Result = frmNewResearch.newInjPr.create_FullName 'injPrFullName
'                .FormFields("fm_InjPrBirthday").Result = frmNewResearch.newInjPr.birthday 'injPrBirthday
'                .FormFields("fm_InjPrSex").Result = frmNewResearch.newInjPr.sex 'injPrSex
'                .FormFields("fm_deceaseDate").Result = frmNewResearch.newInjPr.decease 'deceaseDate
'                .FormFields("fm_autopsyDate").Result = frmNewResearch.newInjPr.autopsyDate 'autopsy
'                .FormFields("fm_fNumberEF").Result = frmNewResearch.newEF.getFullNumber 'fNumberEF
'                .FormFields("fm_factCase").Result = frmNewResearch.txtFactCase.Text 'factCase
'                .FormFields("fm_legalGround").Result = frmNewResearch.create_legalGround 'legalGround
'                .FormFields("fm_expertName").Result = reverse_Name(frmNewResearch.newExpert)
'                'Сопроводительная записка  nt (note)
'                .FormFields("nt_EFFirstDay").Result = frmNewResearch.newEF.firstDay 'firstDate
'                .FormFields("nt_fNumberEF").Result = frmNewResearch.newEF.getFullNumber 'fNumberEF
'                .FormFields("nt_autopsy").Result = frmNewResearch.newInjPr.print_Autopsy 'autopsy info
'                .FormFields("nt_stringEF").Result = frmNewResearch.newEF.toStringEF 'stringEF
'                .FormFields("nt_expertName").Result = reverse_Name(frmNewResearch.newExpert)
''                Печать активного документа
'                If MsgBox("Напечатать " & nameDOC & " ?", vbYesNo) = vbYes Then
'                    If nameDOT = "На гистологию" Then
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
'            'Сохранение:
'            .SaveAs FileName:=tmpdir & nameDOC, FileFormat:= _
'                wdFormatXMLDocumentMacroEnabled, LockComments:=False, Password:="", _
'                AddToRecentFiles:=True, WritePassword:="", ReadOnlyRecommended:=False, _
'                EmbedTrueTypeFonts:=False, SaveNativePictureFormat:=False, SaveFormsData _
'                :=False, SaveAsAOCELetter:=False
'            'Закрытие
'            .Close
'        End With
'    End With
'Set SvcService = Nothing    'объект библиотеки  svcsvc.dll
''уничтожаем обьект - документ
'Set DocWord = Nothing
''уничтожаем объект приложения
''MyWdApp.Quit
''Set MyWdApp = Nothing
'End Sub
