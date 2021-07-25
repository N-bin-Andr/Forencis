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
'Класс clmExpertFindings "Заключение эксперта"
'@author Andr.N@_bin
'@E-mail Andr.Nab.n@gmail.com
Option Explicit
Private tmpid As Long                  'идентификатор (количество созданных объектов класса)
Private Const OUR As String = "Судебная медико-криминалистическая экспертиза"
Private tmpname As String              'название экспертизы (медико-криминалистическая/и т.д.)
Private tmpnumber As String            'журнальный номер заключения эксперта
Private tmpEFcode As String            'код отдела в номере экспертизы
Private tmpdefinition As String        'Основание назначения экспертиз  ы: постановление/определение
Private tmprulingDate As Date          'Дата вынесения постановления
Private tmpreceptDate As Date          'Дата получения постановления
Private tmpfirstDay As Date            'дата начала экспертизы
Private tmpsuspentDate As Date         'дата приостановления производства экспертизы
Private tmprenewDate As Date           'Дата возобновления производства экспертизы
Private tmpdueDate As Date             'Скрок окончания производства экспертизы
Private tmpfinishDate                  'Дата окончания экспертизы
Private tmpLegalGround As String       'Юридическое основание проведения экспертизы
Private tmpreadiness As Currency       'Процент готовности экспертизы
Private tmpimportance As String        'Категория "Важность" экспертизы
Private tmpcategories As String        'Категория экспертизы (первичная, дополнительная и т.д.)
Private tmptanatology As Boolean       'категория экспертизы Танатология/Амбулатория
Private tmpevidensCategories As String 'Категория экспертизы по исследуемым объктам
Private tmpcondition As String         'Состояние экспертизы (не начата, в работе и т.д.)
Private tmpcomments As String          'Примечания
'КОЛЛЕКЦИЯ ВД

'константы:
Const CRIMINALISTICA_EF As String = "5.2"
'
'Const TANATOLOGY_EF As String = "25-6-4.8/" 'Полный префикс номера заключения по Танатологии (ООЭ №3)
'Const AMBULANT_EF As String = "25-6-4.7/"   'Полный префикс номера заключения по Амбулатории (ООЭ №3)
'Const TANATOLOGY_ARH As String = "700_7"    'префикс архивного номера по Танатологии (ООЭ №3)
'Const AMBULANT_ARH As String = "700_6"      'префикс архивного номера по Амбулатории(ООЭ №3)
'

'==================================== I N I T I A L I S A T I O N ====================================
'
Private Sub Class_Initialize()
End Sub
'
Private Sub Class_Terminate()
End Sub
'
Friend Function getNumber() As String
'строка =  NNN_YY (номер_год (2цифры)начала экспертизы)
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
Friend Function getNum_Cat() As String
'строка = №СМЭ_YY_Общий
    getNum_Cat = getNumber & "_" & evidensCategories
Debug.Print "getNum_Cat = " & getNum_Cat
End Function
'
Friend Function getFullNumber(Optional strEFcode As String = CRIMINALISTICA_EF) As String
'Функция создания полного номера документа (экспертизы)
'строка=5.2/NNN
    If strEFcode <> "" Then
        getFullNumber = "№" & strEFcode & "/" & number
    Else: getFullNumber = "№" & number
    End If
'    Debug.Print "getFullNumber: " & getFullNumber
End Function
'
Friend Function getArchName(Optional fn As String = "", _
                            Optional expert As String = "", _
                            Optional manager As String = "Комар") As String
'Функция создания архивного номера экспертизы
'fn - номер экспертизы;expert - фамилия эксперта; manager-фамилия руководителя
'строка=012_***fnз_49_expert_manager
getArchName = "012_" & Format(fn, "#00000") & "з_49_" & Right(CStr(Year(Now)), 2) & "_" & expert & "_" & manager
Debug.Print "getArchName: " & getArchName
End Function
'
Friend Function toStringEF() As String
'строка = № + полный номер+ от + Дата начала
    toStringEF = "№" & getFullNumber & " от " & firstDay & "."
Debug.Print "toStringEF- " & toStringEF
End Function
'
Friend Function toString_InfoEF(Optional tmpLegalGround As String = "", Optional expert As String = "Набебин А.В.") As String
'строка = Mедико-криминалистическая экспертиза №5.2/NNN + Дата начала: DD.MM.YYYY. + ГМСЭ " + expert
    If tmpLegalGround <> "" Then
        toString_InfoEF = OUR & Chr(32) & getFullNumber & _
        ", на " & tmpLegalGround & _
        ". Дата начала: " & firstDay & ". ГМСЭ " & expert
    Else
        toString_InfoEF = name & getFullNumber & ". Дата начала: " & firstDay & ". ГМСЭ " & expert
    End If
Debug.Print "toStringEF- " & toStringEF
End Function
'
Public Static Function printDefinition() As String
'строка = "на основании постановления/определения"
    printDefinition = "основании " & definition & Chr(32)
End Function
'
Friend Function getLegalGround(Optional tmpCor As String = "", _
                                Optional tmpCase As String = "") As String
 'строка =   на основании постановления/определения + следователя ... + вынесенного +
' дата вынесения постановления + по материалам уголовного дела №**************

getLegalGround = printDefinition & "о назначении " & categories & _
                " экспертизы " & tmpCor & ", вынесенного " & rulingDate & Chr(32) & tmpCase
End Function
'
Friend Function print_AEdoc(aeExpert As String, autopsy As String) As String
'строка = "Из направления ГМСЭ ФИО следует: и т.д.
'String aeExpert (данные общего эксперта) = clmExperts.print_Expert
'String autopsy  (данные по вскрытию) = clmEntrants.print_Autopsy
print_AEdoc = "Из " & aeExpert & " следует: " & Chr(171) & "... материал от трупа " & _
            autopsy & Chr(32) & _
            "Заключение эксперта: " & getFullNumber(EFcode) & Chr(32) & _
            "Краткие обстоятельства и предполагаемая причина смерти " & Chr(10) & _
            "Судебно-медицинский диагноз " & "..." & Chr(187) & Chr(10)
Debug.Print "направление ОЭ - " & print_AEdoc
End Function
'
'==================================== I N C A P S U L A T I O N ====================================
'++++++++++++++++++++++++++++ L E T ++++++++++++++++++++++++++++
Private Property Let id(ByVal vData As Long)
'показания счетчика
    tmpid = vData
End Property
'
Friend Property Let name(ByVal vData As String)
'название экспертизы (медико-криминалистическая/и т.д.)
'    If Len(vData) = 0 Then
'        tmpname = "Медико-криминалистическая экспертиза "
'    Else:
    tmpname = vData
'    End If
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
1545            tmp = InputBox("Введите правильно номер экспертизы!", "Исправление ошибки ввода")
        Else: tmpnumber = tmp
            Exit Do
        End If
    Loop
End Property
'
Friend Property Let EFcode(ByVal vData As String)
'код отдела в номере экспертизы
tmpEFcode = vData
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
        tmpLegalGround = "нет данных"
    Else: tmpLegalGround = vData
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
'++++++++++++++++++++++++++++ G E T ++++++++++++++++++++++++++++
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
Friend Property Get EFcode() As String
'код отдела в номере экспертизы
EFcode = tmpEFcode
'Debug.Print "код отдела в номере экспертизы -> " & EFcode
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
    legalGround = tmpLegalGround
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
