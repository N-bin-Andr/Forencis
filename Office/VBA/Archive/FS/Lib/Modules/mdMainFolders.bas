Attribute VB_Name = "mdMainFolders"
'mdMainFolders -> модульдля работы с директориями
'@author Andr.Nab.n@gmail.com
'Lib:
'newEF.getNumber = №СМЭ_YY
'newEF.getNum_Cat = №СМЭ_YY_Общий
Option Explicit
Private fs As FileSystemObject  'экземпляр  FileSystemObject
Private SvcService As Object    'объект библиотеки svcsvc.dll
Private Const MAIN_FOLDER_NOT_EXIST = "Папка для сохранения новых документов не указана!"
Public Const USER_DOC_DIR As String = "D:\Crime\Soft_поРаботе\VBA\1_6\Resources\user\usrDocDir.txt" 'файл данных, содержащий директорию для новых документов, введенную пользователем
Public Const USER_DOT_DIR As String = "D:\Crime\Soft_поРаботе\VBA\1_6\Resources\user\usrDotDir.txt"
Private tmpdirDOT As String 'директория шаблонов документов
Private tmpdirDOC As String 'директория для сохранения готовых документов

Public Const DOC_INDEX As Integer = 0
Public Const DOT_INDEX As Integer = 1
Private fileDescr As Integer
'массив значений путей
Public arrDocDir(0 To 6) As String
'arrDocDir (0) = основная директроия для новых документов(введенная пользователем) Пр:"D:\Crime\"
'arrDocDir (1) = основная директроия с шаблонами создаваемых документов(введенная пользователем) Пр:"D:\Crime\DOT\"

'arrDocDir (1) = 0 + \YYYY  (D:\Crime\2020)
'arrDocDir (2) = 1 + \newEF.getNum_Cat  (D:\Crime\YYYY\№СМЭ_YY_Общий)
'arrDocDir (3) = 2+ \Фото_" & criateDocName(№СМЭ)
'arrDocDir (4) = 2+ \Упаковки_" &
'arrDocDir (5) = 2+"\Сканы_"  &
'arrDocDir (6) = 2+"\Сопровод_"  &
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'Private Property Let dirDOT(ByVal vData As String)
''директория шаблонов документов
'    tmpdirDOT = vData
'End Property
''
'Public Property Get dirDOT(tmpRoot As String) As String
''директория шаблонов документов
'   fileDescr = FreeFile
'    Open tmpRoot For Input As #fileDescr
'        While Not EOF(1)
'            Input #1, str
'            Debug.Print "Чтение из файла = " & tmpdirDOT
'        Wend
'    Close #fileDescr
'    dirDOT = tmpdirDOT
'Debug.Print "dirDOT = " & dirDOT
'End Property

Public Sub getMainDir(tmpRoot As String, i As Integer)
'функция получения пути из текстового файла.
'String tmpRoot -> переменная пути к файлам: usrDocDir.txt/usrDotDir.txt содержащим: директории для новых документов/шаблонов документов, введенных пользователем
'Integer i -> индекс массива для записи директории для новых документов(0)/шаблонов документов (1),
'1) получение пути из текстового файла.
    Dim str As String
    fileDescr = FreeFile
    Open tmpRoot For Input As #fileDescr
        While Not EOF(1)
            Input #1, str
            Debug.Print "Чтение из файла = " & str
        Wend
    Close #fileDescr
'   проверка на пустое значение
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
'            Debug.Print "Чтение из файла = " & tmpRoot
'        Wend
'    Close #fileDescr
''   проверка на пустое значение
'    If tmpRoot <> "" Then
'        arrDocDir(0) = tmpRoot
'    Else
'        Call inputUserDir
'    End If
End Sub
'
Public Sub inputUserDir(i As Integer)
'выбор директории для новых документов
fileDescr = FreeFile
Set SvcService = CreateObject("Svcsvc.Service") 'объект библиотеки  svcsvc.dll
Set fs = New FileSystemObject                   'экземпляр  FileSystemObject

Dim str1 As String, str2 As String
    If i = 0 Then
        str1 = "Укажите папку для сохранения новых документов!"
        str2 = USER_DOC_DIR
    ElseIf i = 1 Then
        str1 = "Укажите папку,  содержащую шаблоны создаваемых документов!"
        str2 = USER_DOT_DIR
    End If
    MsgBox str1, vbExclamation, "Выбор новой папки"
       arrDocDir(i) = SvcService.SelectFolder("Выбор папки", "", &H10 + &H4000, "")
'           запись пути в текстовый файл программы:
            Open str2 For Output As #fileDescr
                Print #fileDescr, arrDocDir(i)    'печать нового пути;
                    Debug.Print "newUserDir = " & arrDocDir(i)
                    MsgBox "Выбрана папка " & arrDocDir(i) & "!", vbOKOnly, "Новая папка"
            Close #fileDescr
        Set SvcService = Nothing    'уничтожение объекта библиотеки  svcsvc.dll
        Set fs = Nothing            'уничтожение экземпляра  FileSystemObject
'старая версия +++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'fileDescr = FreeFile
'Set SvcService = CreateObject("Svcsvc.Service") 'объект библиотеки  svcsvc.dll
'Set fs = New FileSystemObject                   'экземпляр  FileSystemObject
'    MsgBox "Укажите папку для сохранения новых документов!", vbExclamation, "Выбор новой папки"
'       arrDocDir(0) = SvcService.SelectFolder("Выбираем папку для сохранения документов", "", &H10 + &H4000, "")
''           запись пути в текстовый файл программы:
'            Open USER_DOC_DIR For Output As #fileDescr
'                Print #fileDescr, arrDocDir(0)    'печать нового пути;
'                    Debug.Print "newUserDir = " & arrDocDir(0)
'                    MsgBox "Выбрана папка " & arrDocDir(0) & "!", vbOKOnly, "Новая папка"
'            Close #fileDescr
'        Set SvcService = Nothing    'уничтожение объекта библиотеки  svcsvc.dll
'        Set fs = Nothing            'уничтожение экземпляра  FileSystemObject
'старая версия +++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
End Sub
'
Public Sub makeDocDir(Optional fn As String = "", _
                        Optional fnCat As String = "Общий")
'fn - журнальный номер документа_Год = newEF.getNum_Cat
'fnCat - категория экспертизы по исследуемым объектам = newEF.getNum_Cat
'1) смена директории:
    ChDrive "D"
'   считывание данных из папки с указанием пути (сохраненного пользователем)
    Call getMainDir(USER_DOC_DIR, 0)
'2)проверка главной директории
    If arrDocDir(0) = "" Then
    MsgBox MAIN_FOLDER_NOT_EXIST, vbExclamation, "Ошибка сохранения!"
        Call inputUserDir(0)
    End If
'создание папок в рабочей директории
    If fn <> "" Then
        arrDocDir(2) = arrDocDir(0) & "\" & fn & "_" & Right(CStr(Year(Now)), 2) & "_" & fnCat  '(D:\Crime\YYYY\№СМЭ_YY_Общий)
        arrDocDir(3) = arrDocDir(2) & "\Фото_" & fn & "_" & Right(CStr(Year(Now)), 2)           '(D:\Crime\YYYY\№СМЭ_YY_Общий\Фото_№СМЭ_YY)
        arrDocDir(4) = arrDocDir(2) & "\Упаковки_" & fn & "_" & Right(CStr(Year(Now)), 2)       '(D:\Crime\YYYY\№СМЭ_YY_Общий\Упаковки_№СМЭ_YY)
        arrDocDir(5) = arrDocDir(2) & "\Сканы_" & fn & "_" & Right(CStr(Year(Now)), 2)          '(D:\Crime\YYYY\№СМЭ_YY_Общий\Сканы_№СМЭ_YY)
        arrDocDir(6) = arrDocDir(2) & "\Сопровод_" & fn & "_" & Right(CStr(Year(Now)), 2)       '(D:\Crime\YYYY\№СМЭ_YY_Общий\Сопровод_№СМЭ_YY)
    End If
    
    Dim i As Integer
        For i = 2 To UBound(arrDocDir)
            MkDir arrDocDir(i)
        Next i
Отладка:
    Dim x As Integer
     For x = LBound(arrDocDir) To UBound(arrDocDir)
        Debug.Print "Значение массива " & x & " = " & arrDocDir(x) & Chr(10)
    Next x
End Sub
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

