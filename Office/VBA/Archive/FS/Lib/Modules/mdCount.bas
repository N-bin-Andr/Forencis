Attribute VB_Name = "mdCount"
'модуль "mdCount" для работы со счетчиками
'содержит переменные и функции для работы со счетчиками
'NB!!! работает с классом clmCounter
'счетчики вынесены в отдельный модуль, т.к. объявление и работа с
'классом проло работают в формах?
'Дата создания: 01.06.2016
'@version 0.0.1
'@author Andr.Nab.n@gmail.com
Option Explicit
'Объявление счетчиков
Public boxCounter As clmCounter    'Количесво упаковок с ВД
Public boxEvSumCounter As clmCounter  'счетчик "сумма ВД" в упаковке
Public allEvSumCounter As clmCounter  'счетчик "Общая сумма ВД"
'Public colForms As Collection       'коллекция форм


'Public counterEF As clmCounter       'Счетчик форм "Заключение эксперта"
'Public counterEvCat As clmCounter    'Счетчик "Категогия ВД (и документов)"
'Public mvarfEvidCount As clmCounter      'Счетчик форм "Вещественные доказательства"
'Public mvarfDocListCount As clmCounter   'Счетчик форм "Список документов"
'
''Объявление счетчиков
'Private mvarfEFCount As Long        'Счетчик форм "Заключение эксперта"
'Private mvarfEvCat_Count As Long    'Счетчик "Категогия ВД (и документов)"
'Private mvarfEvidCount As Long      'Счетчик форм "Вещественные доказательства"
'Private mvarfDocListCount As Long   'Счетчик форм "Список документов"
'Private mvarEvGRCount As Integer    'Количесво групп ВД
'Private mvarEvSumCount As Long      'счетчик "Общая сумма ВД"
''Создание локальных переменных
''Private mvarDocCategory As String   'категория документа
''Private mvarstrDOC As String        'с
''Private mvarstrDOT As String        'шаблон документа
'
'Public Property Let fEvidCount(ByVal vData As Integer)
''Счетчик форм "Вещественные доказательства"
'mvarfEvidCount = vData
'End Property
''
'Public Property Get fEvidCount() As Integer
''Счетчик форм "Вещественные доказательства"
'fEvidCount = mvarfEvidCount
'Debug.Print "Счетчик Вещественные доказательства = ", fEvidCount
'End Property
''
'Public Property Let EvGRCount(ByVal vData As Integer)
''Счетчик групп ВД
'    mvarEvGRCount = vData
'End Property
''
'Public Property Get EvGRCount() As Integer
''Счетчик групп ВД
'    EvGRCount = mvarEvGRCount
'Debug.Print "'Счетчик групп ВД= ", EvGRCount
'End Property
''
'Public Property Let fEFCount(ByVal vData As Long)
''Счетчик форм "Заключение эксперта".
'    mvarfEFCount = vData
'End Property
''
'Public Property Get fEFCount() As Long
''Счетчик форм "Заключение эксперта"
'    fEFCount = mvarfEFCount
''Debug.Print "Счетчик форм "Заключение эксперта" = ", EvSumCount
'End Property
''
'Public Property Let fEvCat_Count(ByVal vData As Long)
''"Категогия ВД (и документов)"
'    mvarfEvCat_Count = vData
'End Property
''
'Public Property Get fEvCat_Count() As Long
''Счетчик форм "Категогия документов".
'    fEvCat_Count = mvarfEvCat_Count
'Debug.Print "Счетчик форм Закл эксперта = ", fEvCat_Count
'End Property
''
'Public Property Let EvSumCount(ByVal vData As Long)
''Счетчик "Общая сумма ВД".
'    mvarEvSumCount = vData
'End Property
''
'Public Property Get EvSumCount() As Long
''Счетчик "Общая сумма ВД".
'    EvSumCount = mvarEvSumCount
''Debug.Print "Счетчик "Общая сумма ВД" = ", EvSumCount
'End Property
''
'Public Property Let fDocListCount(ByVal vData As Long)
''Счетчик форм frmDocList
'mvarfDocListCount = vData
'End Property
''
'Public Property Get fDocListCount() As Long
''Счетчик форм frmDocList
'fDocListCount = mvarfDocListCount
'End Property

'Public Property Let DocCategory(ByVal vData As String)
''Переменная Категогия документов.
'    mvarDocCategory = vData
'End Property
''
'Public Property Get DocCategory() As String
''Переменная Категогия документов.
'    DocCategory = mvarDocCategory
''Debug.Print "Переменная Категогия документов = ", DocCategory
'End Property
''
'Public Property Let strDOC(ByVal vData As String)
''Переменная Шаблон документа.
'    mvarstrDOC = vData
'End Property
''
'Public Property Get strDOC() As String
''Переменная Шаблон документа.
'    strDOC = mvarstrDOC
''Debug.Print "Шаблон документа = ", strDOC
'End Property

