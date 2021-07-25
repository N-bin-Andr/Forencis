VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "HR_Official"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "Class HR_Official - класс, предназначенный для работы с документами и базами данных по судебной экспертизе. Содержит поля и методы, позволяющие создавать объекты класса, например, как ""Следователь"" и/или ""Эксперт""."
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'Класс Hight Rank Officials (следователь/эксперт)
'Работа с программами VBA или базами данных Access
'Обработка данных при создании медико-криминалистических заключений
'Дата создания: 01.06.2016
'@version 0.0.1
'@author Andr.Nab.n@gmail.com
'
Option Explicit
'Поля (переменные) класса
Private mvarHR_name As String           'Имя
Private mvarHR_sex As String            'пол
Private mvarHR_SurName As String        'Фамилия
Private mvarHR_MidName As String        'Отчество
Private mvarHR_Post As String           'Должность
Private mvarHR_OprAr As String          'Район обслуживания
Private mvarHR_Rank As String           'Звание
Private mvarHR_PI As String             'Персональные данные
Private mvarHR_Comments As String       'Коментарии
Private mvarHR_Conformation As String   'Структура подразделения
Private mvarHR_Office As String         'Управление подразделения
Private mvarHR_Department As String     'Отдел подразделения
Private mvarHR_Experience As String     'Стаж работы по специальности
Const MyExperience As String = "01.08.2005"
'
Friend Function Create_FullRank() As String
'Функция создания текста должности и звания Эксперта
Create_FullRank = HR_Post & Chr(32) & HR_OprAr & Chr(32) & HR_Rank
End Function
'
Friend Function Print_Expert() As String
Attribute Print_Expert.VB_Description = "функция печати обзаца ""эксперт"" на титульном листе заключения эксперта"
'Функция набота текста направление "Общего эксперта"
Print_Expert = "Государственного медицинского судебного эксперта " & Create_FullRank & Chr(32) & Create_InitialsFullName
End Function
'
Friend Function Calculate_Experience() As String
'Функция расчета стажа работы
Dim lngEx1 As Long, strEx2 As String
    lngEx1 = DateDiff("yyyy", CDate(HR_Experience), Now) 'расчет стажа работы
    Select Case lngEx1
        Case Is = 1
            Calculate_Experience = lngEx1 & " год,"
        Case 2 To 4
            Calculate_Experience = lngEx1 & " года,"
        Case 5 To 20
            Calculate_Experience = lngEx1 & " лет,"
          Case Else
            strEx2 = Right(lngEx1, 1)
                Select Case CStr(HR_Experience)
                    Case Is = 0
                        Calculate_Experience = lngEx1 & " лет,"
                    Case Is = 1
                        Calculate_Experience = lngEx1 & " год,"
                    Case 2 To 4
                        Calculate_Experience = lngEx1 & " года,"
                    Case 5 To 9
                        Calculate_Experience = lngEx1 & " лет,"
                End Select
    End Select
End Function
'
Friend Function Print_Experience(ByVal HR_Experience As String) As String
'Функция печати "Стаж работы по специальности с (Пр: 2005) года"
Print_Experience = Right(HR_Experience, 4) & " года"
End Function
'
Friend Property Let HR_Experience(ByVal vData As String)
Attribute HR_Experience.VB_Description = "поле ""опыт работы"""
'Стаж работы по специальности
    mvarHR_Experience = vData
End Property
'
Friend Property Get HR_Experience() As String
'Стаж работы по специальности
    HR_Experience = mvarHR_Experience
'Debug.Print "Стаж работы по специальности: " & HR_Experience
End Property
'
Friend Function Print_Coroner() As String
Attribute Print_Coroner.VB_Description = "функция набора  обзаца ""следователь"" на титульном листе заключения эксперта"
'Функция набора текста "Следователь"
    If Create_FullRank = "" Then
       Print_Coroner = "следователя (должность и звание не указаны)  " & Create_InitialsFullName
    Else:
        Print_Coroner = Create_FullRank & Chr(32) & Create_InitialsFullName
    End If
End Function
'
Friend Function Create_InitialsFullName() As String
Attribute Create_InitialsFullName.VB_Description = "функция создания фамилии и инициалов имени и отчества с точками следователя/эксперта"
'Функция создания фамилии, инициалов имени и отчества следователя с точками
Create_InitialsFullName = HR_SurName & Chr(32) & Create_InitialsName(ByVal HR_name) & _
    Create_InitialsMidName(ByVal HR_MidName)
End Function
'
Friend Function Create_InitialsName(ByVal HR_name As String) As String
Attribute Create_InitialsName.VB_Description = "функция создания инициала с точкой из полного имени следователя/эксперта"
'Функция создания инициала имени следователя с точкой
Create_InitialsName = Left(HR_name, 1) & "."
End Function
'
Friend Function Create_InitialsMidName(ByVal HR_MidName As String) As String
Attribute Create_InitialsMidName.VB_Description = "функция создания инициала с точкой из полного отчества следователя/эксперта"
'Функция создания инициала отчества следователя
    Create_InitialsMidName = Left(HR_MidName, 1) & "."
End Function
'
Friend Function Create_FullName() As String
Attribute Create_FullName.VB_Description = "функция создания фамилии, полных имени и отчества следователя/эксперта"
'Функция создания Фамилии, полных Имени и Отчества следователя
Create_FullName = HR_SurName & Chr(32) & HR_name & Chr(32) & HR_MidName
End Function
'
Friend Property Let HR_Department(ByVal vData As String)
Attribute HR_Department.VB_Description = "поле ""отдел подразделения"""
'Отдел подразделения
    mvarHR_Department = vData
End Property
'
Friend Property Get HR_Department() As String
'Отдел подразделения.
    HR_Department = mvarHR_Department
'Debug.Print "Отдел подразделения: " & HR_Department
End Property
'
Friend Property Let HR_Office(ByVal vData As String)
Attribute HR_Office.VB_Description = "поле ""управление (отдел)  подразделения"""
'Управление подразделения
    mvarHR_Office = vData
End Property
'
Friend Property Get HR_Office() As String
'Управление подразделения
    HR_Office = mvarHR_Office
'Debug.Print "Управление подразделения: " & HR_Office
End Property
'
Friend Property Let HR_Conformation(ByVal vData As String)
Attribute HR_Conformation.VB_Description = "поле ""структура подразделения"""
'Структура подразделения
    mvarHR_Conformation = vData
End Property
'
Friend Property Get HR_Conformation() As String
'Структура подразделения.
    HR_Conformation = mvarHR_Conformation
'Debug.Print "Структура подразделения: " & HR_Conformation
End Property
'
Friend Property Let HR_Comments(ByVal vData As String)
Attribute HR_Comments.VB_Description = "поле ""примечания"""
'Коментарии
    mvarHR_Comments = vData
End Property
'
Friend Property Get HR_Comments() As String
'Коментарии.
    HR_Comments = mvarHR_Comments
'Debug.Print "Коментарии: " & HR_Comments
End Property
'
Friend Property Let HR_PI(ByVal vData As String)
Attribute HR_PI.VB_Description = "поле ""персональные данные"""
'Персональные данные
    mvarHR_PI = vData
End Property
'
Friend Property Get HR_PI() As String
'Персональные данные.
    HR_PI = mvarHR_PI
'Debug.Print "Персональные данные: " & HR_PI
End Property
'
Friend Property Let HR_Rank(ByVal vData As String)
Attribute HR_Rank.VB_Description = "поле ""звание"""
'Звание
    mvarHR_Rank = vData
End Property
'
Friend Property Get HR_Rank() As String
'Звание.
    HR_Rank = mvarHR_Rank
'Debug.Print "Звание: " & HR_Rank
End Property
'
Friend Property Let HR_OprAr(ByVal vData As String)
Attribute HR_OprAr.VB_Description = "поле ""район обслуживания"""
'Район обслуживания
    mvarHR_OprAr = vData
End Property
'
Friend Property Get HR_OprAr() As String
'Район обслуживания.
    HR_OprAr = mvarHR_OprAr
'Debug.Print "Район обслуживания: " & HR_OprAr
End Property
'
Friend Property Let HR_Post(ByVal vData As String)
Attribute HR_Post.VB_Description = "поле ""должность"""
'Должность
    mvarHR_Post = vData
End Property
'
Friend Property Get HR_Post() As String
'Должность.
    HR_Post = mvarHR_Post
'Debug.Print "Должность: " & HR_Post
End Property
'
Friend Property Let HR_MidName(ByVal vData As String)
Attribute HR_MidName.VB_Description = "поле ""отчество"""
'Отчество
    mvarHR_MidName = vData
End Property
'
Friend Property Get HR_MidName() As String
'Отчество.
    HR_MidName = mvarHR_MidName
'Debug.Print "Отчество: " & HR_MidName
End Property
'
Friend Property Let HR_SurName(ByVal vData As String)
Attribute HR_SurName.VB_Description = "поле ""фамилия"""
'Фамилия
    mvarHR_SurName = vData
End Property
'
Friend Property Get HR_SurName() As String
'Фамилия.
    HR_SurName = mvarHR_SurName
 'Debug.Print "Фамилия: " & HR_SurName
End Property
'
Friend Property Let HR_sex(ByVal vData As String)
Attribute HR_sex.VB_Description = "поле ""пол"" (мужской/женский)"
'пол
    mvarHR_sex = vData
End Property
'
Friend Property Get HR_sex() As String
'пол.
    HR_sex = mvarHR_sex
'Debug.Print "пол: " & HR_sex
End Property
'
Friend Property Let HR_name(ByVal vData As String)
Attribute HR_name.VB_Description = "поле ""имя"""
'Имя
    mvarHR_name = vData
End Property
'
Friend Property Get HR_name() As String
'Имя.
    HR_name = mvarHR_name
'Debug.Print "Имя: " & HR_name
End Property
'
Private Sub Class_Initialize()
'инициализация переменной "Стаж работы по специальности" моими данными
HR_Experience = MyExperience
End Sub
