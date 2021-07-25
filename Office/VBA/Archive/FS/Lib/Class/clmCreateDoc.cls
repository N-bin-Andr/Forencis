VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clmCreateDoc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Класс для создания папок и печати документов.
'@author Andr.N@_bin
'@E-mail Andr.Nab.n@gmail.com
'
Option Explicit
'объявление системных переменных
Private tmpid As Long
Private fs As FileSystemObject  'экземпляр  FileSystemObject
Private SvcService As Object    'объект библиотеки svcsvc.dll
''Директории:
Const MAIN_ROOT As String = "C:\Users\Admin\Desktop\Для вскрытия\"  '"D:\Locker\Development\GitHub\VBA\ООЭ_3\Сборка проекта\Отладка\"
Private tmpdirExpert As String      'директория корневой папки
Private tmpdirAnalysis As String    'директория  папки "Направлений на исследования"
Private tmpdirDOC As String         'директория  папки "Схемы"
Private tmpDOTDIR As String         'директория папки "Шаблоны документов" по умолчанию
'Word
Public MyWdApp As Word.Application  'экземпляр приложения
Public DocWord As Word.Document     ' экземпляр документа
'Ошибки
Const MAIN_FOLDER_NOT_EXIST As String = "Папка для сохранения документов отсутствует! "
Const CRITICAL As String = "Внимание, ошибка!"
''
'++++++++++++++++++++ И Н И Ц И А Л И З А Ц И Я++++++++++++++++++++++++++++++++++++++
Private Sub Class_Initialize()
'Constructor
'    DOTDIR = "C:\Users\Admin\Documents\MasterForm\" '"D:\Locker\Development\GitHub\VBA\ООЭ_3\Сборка проекта\Documents\MasterForm\"
End Sub
'
Private Sub Class_Terminate()
'Destructor
End Sub
'
'+++++++++++++++++++++++++ М Е Т О Д Ы ++++++++++++++++++++++++++++


'Private Static Function reverse_Name(expertName As String) As String
''строка = инициалы + фамилия
''1)определяем длину строки:
'    Dim i As Integer
'        i = Len(expertName)
''2)отсекаем инициалы:
'    Dim str1 As String
'        str1 = Right(expertName, 4)
''3)собираем строку = инициалы + фамилия
'    reverse_Name = str1 & Chr(32) & Left(expertName, (i - 5))
'Debug.Print "reverse_Name: " & reverse_Name
'End Function
''
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
''
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
''+++++++++++++++++++++ И Н К А П С У Л Я Ц И Я +++++++++++++++++++++++++++++++++++++++
'Private Property Let id(ByVal vData As Long)
''ID класса
'    tmpid = vData
'End Property
''
'Friend Property Get id() As Long
''ID класса
'    id = tmpid
''Debug.Print "ID: " & id
'End Property
''
'Friend Property Let dirExpert(ByVal vData As String)
''директория корневой папки
'    tmpdirExpert = vData
'End Property
''
'Friend Property Get dirExpert() As String
''директория корневой папки
'     dirExpert = tmpdirExpert
''Debug.Print "директория корневой папки: " &  dirExpert
'End Property
''
'Friend Property Let dirAnalysis(ByVal vData As String)
''директория  папки "Направлений на исследования"
'    tmpdirAnalysis = vData
'End Property
''
'Friend Property Get dirAnalysis() As String
''директория  папки "Направлений на исследования"
'     dirAnalysis = tmpdirAnalysis
''Debug.Print "директория  папки "Направлений на исследования": " &  dirAnalysis
'End Property
'
'Friend Property Let dirDoc(ByVal vData As String)
''директория  папки "Таблицы и запросы"
'    tmpdirDoc = vData
'End Property
''
'Friend Property Get dirDoc() As String
''директория  папки "Таблицы и запросы"
'     dirDoc = tmpdirDoc
''Debug.Print "директория  папки "Таблицы и запросы": " &  dirDoc
'End Property
'
'Private Property Let DOTDIR(ByVal vData As String)
''директория  папки "Шаблоны документов"
'    tmpDOTDIR = vData
'End Property
''
'Friend Property Get DOTDIR() As String
''директория  папки "Шаблоны документов"
'     DOTDIR = tmpDOTDIR
''Debug.Print "директория  папки "Шаблоны документов": " &  DOTDIR
'End Property
'
'
