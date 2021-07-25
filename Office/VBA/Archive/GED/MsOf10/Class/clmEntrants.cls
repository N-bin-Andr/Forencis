VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clmEntrants"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "Class Entrants (участники) - класс, предназначенный для работы с документами и базами данных по судебной экспертизе. Содержит поля и методы, позволяющие создавать объекты класса, например, как ""Потерпевший (Injured person)"" и/или ""Владелец вещественных доказательств (Owner)""."
'Класс clmEntrants "Участники дела"
'@author Andr.N@_bin
'@E-mail Andr.Nab.n@gmail.com
Option Explicit
Private tmpid As Long            'идентификатор (количество созданных объектов)
Private tmpitem As String        'обменная переменная (для передачи значения в функцию)
'Переменные (поля) класса
Private tmpsurName As String     'Фамилия
Private tmpname As String        'Имя
Private tmpmidName As String     'Отчество
Private tmpsex As String         'Пол
Private tmplegalStatus As String 'юридический статус (обвиняемый/подозреваемый/потерпевший/свидетель и т.д.)
Private tmpbirthday As String    'Дата/год рождения
Private tmptoHospital As Date    'Дата госпитализации
Private tmpfromHospital As Date  'Дата выписки из ЛПУ
Private tmpdecease As Date       'Дата смерти
Private tmpautopsyDate As Date       'Дата вскрытия
Private tmppreDiagnosis As String 'Предварительный диагноз
Private tmppi As String          'Персональные данные
Private tmpcomments As String    'Примечания
'
Private Sub Class_Initialize()
'конструктор
    id = id + 1
End Sub
'
Public Function create_Initials(item As String) As String
'Функция создания инициала с точкой
    Dim tmp As String
        tmp = Left(item, 1) & "."
    create_Initials = StrConv(tmp, vbProperCase)
'Debug.Print "Инициал с точкой-" & tmp
End Function
'
Friend Function create_InitialslName() As String
'функция = фамилия + инициалы имени и отчества (с заглавной буквы с точками)
Dim tmp1 As String, tmp2 As String
item = name
    If item = "(имя не указано)" Then
        tmp1 = item
    Else: tmp1 = create_Initials(item)
    End If
item = midName
    If item = "(отчество не указано)" Then
        tmp2 = item
    Else: tmp2 = create_Initials(item)
    End If
create_InitialslName = surName & Chr(32) & tmp1 & tmp2
'Debug.Print "clmEntrants фамилия + инициалы имени и отчества: " & create_InitialslName
End Function
'
Friend Function create_FullName() As String
'строка = Фамилия + полное Имя + полное Отчество
Dim str As String
    If name <> "" Then
        str = name
        If Len(str) > 2 Then
            str = name & Chr(32)
        End If
    End If
create_FullName = surName & Chr(32) & str & midName
'Debug.Print "create_FullName- " & Create_FullName
End Function
'
Friend Function print_InjPersons() As String
'строка = Фамилия + Имя + Отчество + год рождения (если указан)
'1) проверка длины строки
    Dim str As String
    str = name
        If Len(str) > 2 Then
            str = name & Chr(32)
        End If
'
    If birthday = "" Then
        print_InjPersons = surName & Chr(32) & str & midName
    Else: print_InjPersons = surName & Chr(32) & str & midName _
        & Chr(32) & birthday & " г.р."
    End If
'Debug.Print "print_InjPersons Фамилия + Имя + Отчество + год рождения - " & print_InjPersons
End Function
'
Friend Function print_Autopsy() As String
'печать данных по вскрытию
    print_Autopsy = print_InjPersons & " Дата смерти: " & decease & "." & _
                        " Дата вскрытия: " & autopsyDate
'Debug.Print "print_Autopsy- " & print_Autopsy
End Function
'
'инкапсуляция
'Let=======================================================
Friend Property Let surName(ByVal vData As String)
'Фамилия
    tmpsurName = StrConv(vData, vbProperCase)
End Property
'
Friend Property Let name(ByVal vData As String)
'Имя
        If Len(vData) = 1 Then
            item = vData
            tmpname = create_Initials(item)
        Else: tmpname = StrConv(vData, vbProperCase)
        End If
End Property
'
Friend Property Let midName(ByVal vData As String)
'Отчество
        If Len(vData) = 1 Then
            item = vData
            tmpmidName = create_Initials(item)
        Else: tmpmidName = StrConv(vData, vbProperCase)
        End If
End Property
'
Friend Property Let sex(ByVal vData As String)
'Пол.
    If vData = "" Then ' Len(vData) = 0
        tmpsex = "пол не указан"
    Else: tmpsex = vData
    End If
End Property
'
Friend Property Let legalStatus(ByVal vData As String)
'юридический статус
    If vData = "" Then
        tmplegalStatus = ""
    Else: tmplegalStatus = vData
    End If
End Property
'
Friend Property Let birthday(ByVal vData As String)
'Дата/год рождения
    tmpbirthday = vData
End Property
'
Friend Property Let toHospital(ByVal vData As Date)
'Дата госпитализации
    tmptoHospital = vData
End Property
'
Friend Property Let fromHospital(ByVal vData As Date)
'Дата выписки из ЛПУ
    tmpfromHospital = vData
End Property
'
Friend Property Let decease(ByVal vData As Date)
'Дата смерти
    tmpdecease = vData
End Property
'
Friend Property Let autopsyDate(ByVal vData As Date)
'Дата вскрытия
    tmpautopsyDate = vData
End Property
'
Friend Property Let preDiagnosis(ByVal vData As String)
'Предварительный диагноз
    If vData = "" Then
        tmppreDiagnosis = "нет данных"
    Else: tmppreDiagnosis = vData
    End If
End Property
'
Friend Property Let pi(ByVal vData As String)
'персональные данные
    If vData = "" Then
        tmppi = "нет данных"
    Else: tmppi = vData
    End If
End Property
'
Friend Property Let сomments(ByVal vData As String)
'примечания
    If vData = "" Then
        tmpcomments = "нет данных"
    Else: tmpcomments = vData
    End If
End Property
'
Private Property Let id(ByVal vData As Long)
'идентификатор
    tmpid = vData
End Property
'
Private Property Let item(ByVal vData As String)
'обменная переменная
    tmpitem = vData
End Property
'
'Get================================================
Friend Property Get surName() As String
'Фамилия
    surName = tmpsurName
'Debug.Print "Фамилия владельца = " & surName
End Property
'
Friend Property Get name() As String
'Имя.
    name = tmpname
'Debug.Print "Имя владельца = " & name
End Property
'
Friend Property Get midName() As String
'Отчество
    midName = tmpmidName
'Debug.Print "Отчество владельца = " & midName
End Property
'
Friend Property Get sex() As String
'Пол.
    sex = tmpsex
'Debug.Print "Пол владельца= " & En_Sex
End Property
'
Friend Property Get legalStatus() As String
'юридический статус
    legalStatus = tmplegalStatus
'Debug.Print "legalStatus- " & legalStatus
End Property
'
Friend Property Get birthday() As String
'Дата/год рождения
    birthday = tmpbirthday
'Debug.Print "Дата/год рождения владельца = " & birthday
End Property
'
Friend Property Get toHospital() As Date
'Дата госпитализации
    toHospital = tmptoHospital
'Debug.Print "'Дата госпитализации = " & En_toHospital
End Property
'
Friend Property Get fromHospital() As Date
'Дата выписки из ЛПУ
    fromHospital = tmpfromHospital
'Debug.Print "'Дата выписки из ЛПУ = " & En_fromHospital
End Property
'
Friend Property Get decease() As Date
'Дата смерти
    decease = tmpdecease
'Debug.Print "Дата смерти = " & decease
End Property
'
Friend Property Get autopsyDate() As Date
'Дата вскрытия
    autopsyDate = tmpautopsyDate
'Debug.Print "Дата вскрытия = " & autopsyDate
End Property
'
Friend Property Get preDiagnosis() As String
'Предварительный диагноз
    preDiagnosis = tmppreDiagnosis
'Debug.Print "Предварительный диагноз = " & En_PreDiagnosis
End Property
'
Friend Property Get pi() As String
'Персональные данные
    pi = tmppi
'Debug.Print "Персональные данные = " & En_pi
End Property
'
Friend Property Get id() As Long
'идентификатор
    id = tmpid
'Debug.Print "количество счетчиков- " & id
End Property
'
Friend Property Get сomments() As String
'примечания
    сomments = tmpcomments
'Debug.Print "Коментарии = " & En_Comments
End Property
'
Private Property Get item() As String
'обменная переменная
    item = tmpitem
'Debug.Print "обменная переменная - " & item
End Property
