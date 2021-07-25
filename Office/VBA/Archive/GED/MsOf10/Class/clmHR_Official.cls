VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clmHR_Official"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "Class HR_Official - класс, предназначенный для работы с документами и базами данных по судебной экспертизе. Содержит поля и методы, позволяющие создавать объекты класса, например, как ""Следователь"" и/или ""Эксперт""."
'Класс Hight rank Officials (следователь/эксперт)
'@author Andr.N@_bin
'@E-mail Andr.Nab.n@gmail.com
Option Explicit
Private tmpid As Long               'идентификатор (количество созданных объектов)
Private tmpitem As String           'обменная переменная (для передачи значения в функцию)
Const MY_EXPERIENCE As Date = #1/8/2005#
'Поля (переменные) класса
Private tmpsurName As String        'Фамилия
Private tmpname As String           'Имя
Private tmpmidName As String        'Отчество
Private tmpsex As String            'пол
Private tmplegalStatus              'юридический статус (следователь/эксперт/понятой и т.д.)
Private tmppost As String           'Должность
Private tmprank As String           'Звание
Private tmpconformation As String   'Структура подразделения
Private tmpoffice As String         'Управление подразделения
Private tmpdepartment As String     'Отдел подразделения
Private tmpPCR As String            'должность + структура + звание
Private tmpexperience As String     'Стаж работы по специальности
Private tmppi As String             'Персональные данные
Private tmpcomments As String       'Коментарии
'
Friend Function create_Rank() As String
'строка = должность + управление + звание
    If post = "" Then
        create_Rank = conformation & Chr(32) & rank & "(долность не указана)"
    ElseIf conformation = "" Then
        create_Rank = post & Chr(32) & rank & "(управление не указано)"
    ElseIf rank = "" Then
        create_Rank = post & Chr(32) & conformation & "(звание не указано)"
    Else: create_Rank = post & Chr(32) & conformation & Chr(32) & rank
    End If
Debug.Print "должность + управление + звание= " & create_Rank
End Function
'
Friend Function print_Cor() As String
'строка = должность + структура + звание "Следователя"
        PCR = create_Rank 'должность + структура + звание
    If PCR = "" Then
       print_Cor = "следователя (должность и звание не указаны)  " & create_InitialslName
    Else:
        print_Cor = PCR & Chr(32) & create_InitialslName
    End If
Debug.Print "print_Cor- " & print_Cor
End Function
'
Friend Function print_Exp() As String
'строка = должность + структура + звание "Эксперта"
    print_Exp = "Государственного медицинского судебного эксперта " & create_InitialslName
Debug.Print "print_Exp- " & print_Exp
End Function
'
Public Function create_Initials(ByVal item As String) As String
'Функция создания инициала с точкой
    Dim tmp As String
        tmp = Left(item, 1) & "."
    create_Initials = StrConv(tmp, vbProperCase)
Debug.Print "Инициал с точкой-" & tmp
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
Debug.Print "clmHR_Official create_InitialslName- " & create_InitialslName
End Function
'
Friend Function create_ReverseInitialslName() As String
'функция = инициалы имени и отчества (с заглавной буквы с точками) + фамилия
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
create_ReverseInitialslName = tmp1 & tmp2 & Chr(32) & surName
Debug.Print "create_ReverseInitialslName- " & create_ReverseInitialslName
End Function
'
Friend Function create_FullName() As String
'Функция создания Фамилии, полных Имени и Отчества следователя
Dim str As String
    str = name
        If Len(str) > 2 Then
            str = name & Chr(32)
        End If
create_FullName = surName & Chr(32) & str & midName
Debug.Print "create_FullName- " & create_FullName
End Function
'
Friend Function create_ReverseFullName() As String
'Функция создания Фамилии, полных Имени и Отчества следователя
Dim str As String
    str = name
        If Len(str) > 2 Then
            str = name & Chr(32)
        End If
create_ReverseFullName = str & midName & Chr(32) & surName
Debug.Print "create_ReverseFullName- " & create_ReverseFullName
End Function
'
'инкапсуляция
'Let_____________________________________________
Private Property Let id(ByVal vData As Long)
'показания счетчика
    tmpid = vData
End Property
'
Private Property Let item(ByVal vData As String)
'обменная переменная
    tmpitem = vData
End Property
'
Friend Property Let surName(ByVal vData As String)
Attribute surName.VB_Description = "поле ""фамилия"""
'Фамилия
    If Len(vData) = 0 Then
        tmpsurName = "(фамилия не указана)"
    Else: tmpsurName = StrConv(vData, vbProperCase)
    End If
End Property
'
Friend Property Let name(ByVal vData As String)
'Имя
    If Len(vData) = 0 Then
        tmpname = "(имя не указано)"
    Else
        If Len(vData) = 1 Then
            item = vData
            tmpname = create_Initials(item)
        Else: tmpname = StrConv(vData, vbProperCase)
        End If
    End If
End Property
'
Friend Property Let midName(ByVal vData As String)
Attribute midName.VB_Description = "поле ""отчество"""
'Отчество
    If Len(vData) = 0 Then
        tmpmidName = "(отчество не указано)"
    Else
        If Len(vData) = 1 Then
            item = vData
            tmpmidName = create_Initials(item)
        Else: tmpmidName = StrConv(vData, vbProperCase)
        End If
    End If
End Property
'
Friend Property Let sex(ByVal vData As String)
Attribute sex.VB_Description = "поле ""пол"" (мужской/женский)"
'пол
    If Len(vData) = 0 Then
        tmpsex = "нет данных"
    Else: tmpsex = vData
    End If
End Property
'
Friend Property Let legalStatus(ByVal vData As String)
'юридический статус
    If Len(vData) = 0 Then
        tmplegalStatus = ""
    Else: tmplegalStatus = vData
    End If
End Property
'
Friend Property Let experience(ByVal vData As String)
'Стаж работы по специальности
    If Len(vData) = 0 Then
        tmpexperience = "нет данных"
    Else: tmpexperience = vData
    End If
End Property
'
Friend Property Let post(ByVal vData As String)
Attribute post.VB_Description = "поле ""должность"""
'Должность
    If Len(vData) = 0 Then
        tmppost = "нет данных"
    Else: tmppost = vData
    End If
End Property
'
Friend Property Let rank(ByVal vData As String)
Attribute rank.VB_Description = "поле ""звание"""
'Звание
    If Len(vData) = 0 Then
        tmprank = "нет данных"
    Else: tmprank = vData
    End If
End Property
'
Friend Property Let conformation(ByVal vData As String)
Attribute conformation.VB_Description = "поле ""структура подразделения"""
'Структура подразделения
    If Len(vData) = 0 Then
        tmpconformation = "нет данных"
    Else: tmpconformation = vData
    End If
End Property
'
Friend Property Let office(ByVal vData As String)
Attribute office.VB_Description = "поле ""управление (отдел)  подразделения"""
'Управление подразделения
    If Len(vData) = 0 Then
        tmpoffice = "нет данных"
    Else: tmpoffice = vData
    End If
End Property
'
Friend Property Let department(ByVal vData As String)
Attribute department.VB_Description = "поле ""отдел подразделения"""
'Отдел подразделения
    If Len(vData) = 0 Then
        tmpdepartment = "нет данных"
    Else: tmpdepartment = vData
    End If
End Property
'
Friend Property Let PCR(ByVal vData As String)
Attribute PCR.VB_Description = "поле ""район обслуживания"""
'должность + структура + звание
    If Len(vData) = 0 Then
        tmpPCR = "нет данных"
    Else: tmpPCR = vData
    End If
End Property
'
Friend Property Let comments(ByVal vData As String)
'Коментарии
    If Len(vData) = 0 Then
        tmpcomments = "нет данных"
    Else: tmpcomments = vData
    End If
End Property
'
Friend Property Let pi(ByVal vData As String)
'Персональные данные
    If Len(vData) = 0 Then
        tmppi = "нет данных"
    Else: tmppi = vData
    End If
End Property
'
'Get========================================
Friend Property Get id() As Long
'показания счетчика
    id = tmpid
'Debug.Print "количество счетчиков- " & id
End Property
'
Friend Property Get item() As String
'обменная переменная
    item = tmpitem
'Debug.Print "обменная переменная - " & item
End Property
'
Friend Property Get surName() As String
'Фамилия.
    surName = tmpsurName
'Debug.Print "Фамилия- " & surName
End Property
'
Friend Property Get name() As String
'Имя.
    name = tmpname
'Debug.Print "Имя- " & name
End Property
'
Friend Property Get midName() As String
'Отчество.
    midName = tmpmidName
'Debug.Print "Отчество- " & midName
End Property
'
Friend Property Get sex() As String
'пол.
    sex = tmpsex
'Debug.Print "пол- " & sex
End Property
'
Friend Property Get legalStatus() As String
'юридический статус
    legalStatus = tmplegalStatus
'Debug.Print "legalStatus- " & legalStatus
End Property
'
Friend Property Get post() As String
'Должность.
    post = tmppost
'Debug.Print "Должность- " & post
End Property
'
Friend Property Get rank() As String
'Звание.
    rank = tmprank
'Debug.Print "Звание- " & rank
End Property
'
Friend Property Get conformation() As String
'Структура подразделения.
    conformation = tmpconformation
'Debug.Print "Структура подразделения- " & conformation
End Property
'
Friend Property Get office() As String
'Управление подразделения
    office = tmpoffice
'Debug.Print "Управление подразделения- " & office
End Property
'
Friend Property Get department() As String
'Отдел подразделения.
    department = tmpdepartment
'Debug.Print "Отдел подразделения- " & department
End Property
'
Friend Property Get PCR() As String
'должность + структура + звание
    PCR = tmpPCR
'Debug.Print "должность + структура + звание- " & PCR
End Property
'
Friend Property Get experience() As String
Attribute experience.VB_Description = "поле ""опыт работы"""
'Стаж работы по специальности
    experience = tmpexperience
'Debug.Print "Стаж работы по специальности: " & experience
End Property
'
Friend Property Get comments() As String
Attribute comments.VB_Description = "поле ""примечания"""
'Коментарии.
    comments = tmpcomments
'Debug.Print "Коментарии: " & comments
End Property
'
Friend Property Get pi() As String
Attribute pi.VB_Description = "поле ""персональные данные"""
'Персональные данные.
    pi = tmppi
'Debug.Print "Персональные данные: " & pi
End Property
'
Private Sub Class_Initialize()
'конструктор
    id = id + 1
'инициализация переменной "Стаж работы по специальности" моими данными
    experience = MY_EXPERIENCE
End Sub
