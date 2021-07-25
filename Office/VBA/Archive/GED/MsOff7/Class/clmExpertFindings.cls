VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clmExpertFindings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "Класс-модуль ""ExpertFindings"" - описывает свойства и методы, используемые в ""Заключении Эксперта"""
'Класс "Заключение эксперта"
'@author Andr.N@_bin
'@E-mail Andr.Nab.n@gmail.com
Option Explicit
Private tmpid As Long                  'идентификатор (количество созданных объектов класса)
Private tmpname As String              'название экспертизы (медико-криминалистическая/и т.д.)
Private tmpnumber As String            'журнальный номер заключения эксперта
Private tmpdefinition As String        'Основание назначения экспертизы: постановление/определение
Private tmprulingDate As Date          'Дата вынесения постановления
Private tmpreceptDate As Date          'Дата получения постановления
Private tmpfirstDay As Date            'дата начала экспертизы
Private tmpsuspentDate As Date         'дата приостановления производства экспертизы
Private tmprenewDate As Date           'Дата возобновления производства экспертизы
Private tmpdueDate As Date             'Скрок окончания производства экспертизы
Private tmpfinishDate                  'Дата окончания экспертизы
Private tmplegalGround As String       'Юридическое основание проведения экспертизы
Private tmpreadiness As Currency       'Процент готовности экспертизы
Private tmpimportance As String        'Категория "Важность" экспертизы
Private tmpcategories As String        'Категория экспертизы (первичная, дополнительная и т.д.)
Private tmptanatology As Boolean       'категория экспертизы Танатология/Амбулатория
Private tmpevidensCategories As String 'Категория экспертизы по исследуемым объктам
Private tmpcondition As String         'Состояние экспертизы (не начата, в работе и т.д.)
Private tmpcomments As String          'Примечания
'константы:
Const TANATOLOGY_EF As String = "25-6-4.8/" 'Полный префикс номера заключения по Танатологии (ООЭ №3)
Const AMBULANT_EF As String = "25-6-4.7/"   'Полный префикс номера заключения по Амбулатории (ООЭ №3)
Const TANATOLOGY_ARH As String = "700_7"    'префикс архивного номера по Танатологии (ООЭ №3)
Const AMBULANT_ARH As String = "700_6"      'префикс архивного номера по Амбулатории(ООЭ №3)
'
Private Sub Class_Initialize()
'конструктор
    tanatology = True 'категория экспертизы Танатология/Амбулатория
End Sub
'
Private Sub Class_Terminate()
End Sub
'
Friend Function getNumber() As String
'строка =  номер_год (2цифры)начала экспертизы
Dim dt As Date
    If firstDay = 0 Then
        dt = Year(Now)
    Else
        dt = firstDay
    End If
    
    If number = "" Then
        getNumber = "X_" & Right(CStr(dt), 2)
    Else
        getNumber = number & "_" & Right(CStr(dt), 2)
    End If
Debug.Print "getNumber: " & getNumber
End Function
'
Friend Function getFullNumber() As String
'Функция создания полного номера документа (экспертизы)
    If tanatology = True Then
        getFullNumber = TANATOLOGY_EF & number
    Else
        getFullNumber = AMBULANT_EF & number
    End If
Debug.Print "getFullNumber: " & getFullNumber
End Function
'
Friend Function getArchName(Optional expert As String = "", _
                            Optional manager As String = "") As String
'Функция создания архивного номера экспертизы
    If tanatology = True Then
        getArchName = TANATOLOGY_ARH & Format(number, "#0000") & "з_28_" & Right(CStr(Year(Now)), 2) & _
                   "_" & expert & "_" & manager
    Else
        getArchName = AMBULANT_EF & Format(number, "#0000") & "з_47_" & Right(CStr(Year(Now)), 2) & _
                     "_" & expert & "_" & manager
    End If
Debug.Print "getArchName: " & getArchName
End Function
'
Friend Function toStringEF() As String
'строка = № + полный номер+ от + Дата начала
    toStringEF = "№" & getFullNumber & " от " & firstDay & "."
Debug.Print "toStringEF- " & toStringEF
End Function
'
Friend Function toString_InfoEF() As String
    toString_InfoEF = name & Chr(32) & getFullNumber & ". Дата начала: " & firstDay
Debug.Print "toStringEF- " & toStringEF
End Function
'
Public Static Function printDefinition() As String
'строка = "на основании постановления/определения"
    printDefinition = "на основании " & definition & Chr(32)
End Function
'
'инкапсуляция
'Let=================================================================================
'
Private Property Let id(ByVal vData As Long)
'показания счетчика
    tmpid = vData
End Property
'
Friend Property Let name(ByVal vData As String)
'название экспертизы (медико-криминалистическая/и т.д.)
    If Len(vData) = 0 Then
        tmpname = "Медико-криминалистическая экспертиза "
    Else: tmpname = vData
    End If
End Property
'
Friend Property Let number(ByVal vData As String)
'Журнальный номер документа (экспертизы).
Dim tmp As String
    tmp = vData
    Do
        If Not IsNumeric(tmp) Or Len(tmp) = 0 Then
        Beep
            MsgBox "Cледует ввести цифры!", vbCritical, "Ошибка ввода"
            tmp = InputBox("Введите правильно номер экспертизы!", "Исправление ошибки ввода")
        Else: tmpnumber = tmp
            Exit Do
        End If
    Loop
End Property
'
Friend Property Let definition(ByVal vData As String)
'Основание назначения экспертизы (постановление/определение).
    If Len(vData) = 0 Then
        tmpdefinition = "нет данных"
    Else: tmpdefinition = vData
    End If
End Property
'
Friend Property Let rulingDate(ByVal vData As Date)
'Дата вынесения постановления.
    tmprulingDate = vData
End Property
'
Friend Property Let receptDate(ByVal vData As Date)
'Дата получения постановления.
    tmpreceptDate = vData
End Property
'
Friend Property Let firstDay(ByVal vData As Date)
'Дата начала производства экспертизы.
    tmpfirstDay = vData
End Property
'
Friend Property Let suspentDate(ByVal vData As Date)
'Дата приостановления производства экспертизы.
    tmpsuspentDate = vData
End Property
'
Friend Property Let renewDate(ByVal vData As Date)
'Дата возобновления производства экспертизы.
    tmprenewDate = vData
End Property
'
Friend Property Let dueDate(ByVal vData As Date)
'Срок окончания экспертизы.
    tmpdueDate = vData
End Property
'
Friend Property Let finishDate(ByVal vData As Date)
'Дата окончания экспертизы.
    tmpfinishDate = vData
End Property
'
Friend Property Let legalGround(ByVal vData As String)
'Юридическое основание
'строка = "на основании " + definition + Cor + case:
    If Len(vData) = 0 Then
        tmplegalGround = "нет данных"
    Else: tmplegalGround = vData
    End If
End Property
'
Friend Property Let readiness(ByVal vData As Currency)
'Процент готовности экспертизы.
    tmpreadiness = vData
End Property
'
Friend Property Let importance(ByVal vData As String)
'Категория "Важность" экспертизы.
    If Len(vData) = 0 Then
        tmpimportance = "нет данных"
    Else: tmpimportance = vData
    End If
End Property
'
Friend Property Let categories(ByVal vData As String)
'Категория экспертизы (первичная и т.д.).
    If Len(vData) = 0 Then
        tmpcategories = "первичная"
    Else: tmpcategories = vData
    End If
End Property
'
Friend Property Let tanatology(ByVal vData As Boolean)
 'категория экспертизы Танатология/Амбулатория
    tmptanatology = vData
End Property
'
Friend Property Let evidensCategories(ByVal vData As String)
'Категория экспертизы по исследуемым объктам.
    If Len(vData) = 0 Then
        tmpevidensCategories = "нет данных"
    Else: tmpevidensCategories = vData
    End If
End Property
'
Friend Property Let condition(ByVal vData As String)
'Состояние экспертизы
    If Len(vData) = 0 Then
        tmpcondition = "не начата"
    Else: tmpcondition = vData
    End If
End Property
'
Friend Property Let comments(ByVal vData As String)
'Примечания
    If Len(vData) = 0 Then
        tmpcomments = "нет данных"
    Else: tmpcomments = vData
    End If
End Property
'
'Get==================================================================================
Friend Property Get id() As Long
'показания счетчика
    id = tmpid
'Debug.Print "количество счетчиков- " & id
End Property

Friend Property Get name() As String
'название экспертизы (медико-криминалистическая/и т.д.)
    name = tmpname
'Debug.Print "name- " & name
End Property
'
Friend Property Get number() As String
'Журнальный номер документа (экспертизы).
    number = tmpnumber
'Debug.Print "Журнальный номер документа (экспертизы) = ", number
End Property
'
Friend Property Get definition() As String
'Основание назначения экспертизы (постановление/определение).
    definition = tmpdefinition
'Debug.Print "Основание назначения экспертизы= ", definition
End Property
'
Friend Property Get rulingDate() As Date
'Дата вынесения постановления.
    rulingDate = tmprulingDate
'Debug.Print "Дата вынесения постановления= ", rulingDate
End Property
'
Friend Property Get receptDate() As Date
'Дата получения постановления.
    receptDate = tmpreceptDate
'Debug.Print "Дата получения постановления = ", receptDate
End Property
'
Friend Property Get firstDay() As Date
'Дата начала производства экспертизы.
    firstDay = tmpfirstDay
'Debug.Print "Дата начала производства экспертизы = ", firstDay
End Property
'
Friend Property Get suspentDate() As Date
'Дата приостановления производства экспертизы.
    suspentDate = tmpsuspentDate
'Debug.Print "Дата возобновления производства экспертизы = ", suspentDate
End Property
'
Friend Property Get renewDate() As Date
'Дата возобновления производства экспертизы.
    renewDate = tmprenewDate
'Debug.Print "'Дата возобновления производства экспертизы = ", renewDate
End Property
'
Friend Property Get dueDate() As Date
'Срок окончания экспертизы.
    dueDate = tmpdueDate
'Debug.Print "Срок окончания экспертизы = ", dueDate
End Property
'
Friend Property Get finishDate() As Date
'Дата окончания экспертизы.
    finishDate = tmpfinishDate
'Debug.Print "Дата окончания экспертизы = ", finishDate
End Property
'
Friend Property Get legalGround() As String
'Юридическое основание
'строка = "на основании " + definition + Cor + case
    legalGround = tmplegalGround
'Debug.Print "Юридическое основание = ", legalGround
End Property
'
Friend Property Get readiness() As Currency
'Процент готовности экспертизы.
    readiness = tmpreadiness
'Debug.Print "Процент готовности экспертизы = ", readiness
End Property
'
Friend Property Get importance() As String
'Категория "Важность" экспертизы.
    importance = tmpimportance
'Debug.Print "Категория Важность экспертизы = ", importance
End Property
'
Friend Property Get categories() As String
'Категория экспертизы (первичная и т.д.).
    categories = tmpcategories
'Debug.Print "Категория экспертизы = ", categories
End Property
'
Friend Property Get tanatology() As Boolean
'Категория экспертизы Танатология/Амбулатория
    tanatology = tmptanatology
'Debug.Print "Танатология: " & tanatology
End Property
'
Friend Property Get evidensCategories() As String
'Категория экспертизы по исследуемым объктам.
    evidensCategories = tmpevidensCategories
'Debug.Print "Категория экспертизы по исследуемым объктам = ", evidensCategories
End Property
'
Friend Property Get condition() As String
'Состояние экспертизы.
    condition = tmpcondition
'Debug.Print "Состояние экспертизы = ", condition
End Property
'
Friend Property Get comments() As String
'Примечания.
    comments = tmpcomments
'Debug.Print "Примечания = "; comments
End Property
'Get+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
