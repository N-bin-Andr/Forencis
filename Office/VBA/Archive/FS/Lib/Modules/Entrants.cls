VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Entrants"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "Class Entrants (участники) - класс, предназначенный для работы с документами и базами данных по судебной экспертизе. Содержит поля и методы, позволяющие создавать объекты класса, например, как ""Потерпевший (Injured person)"" и/или ""Владелец вещественных доказательств (Owner)""."
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'Класс "Участники дела"
'Работа с программами VBA или базами данных Access
'Обработка данных при создании медико-криминалистических заключений
'Дата создания: 01.06.2016
'@version 0.0.1
'@author Andr.Nab.n@gmail.com
'
Option Explicit
'Переменные (поля) класса
Private mvarEn_Sex As String            'Пол
Private mvarEn_SurName As String        'Фамилия
Private mvarEn_Name As String           'Имя
Private mvarEn_MidName As String        'Отчество
Private mvarEn_Birthday As String       'Дата/год рождения
Private mvarEn_Decease As Date          'Дата смерти
Private mvarEn_Autopsy As Date          'Дата вскрытия
Private mvarEn_PreDiagnosis As String   'Предварительный диагноз
Private mvarEn_PI As String             'Персональные данные
Private mvarEn_Comments As String       'Коментарии
Private mvarEn_toHospital As Date       'Дата госпитализации
Private mvarEn_fromHospital As Date     'Дата выписки из ЛПУ
Private mvarEnCount As Integer          'Счетчик
'
Friend Property Let EnCount(ByVal vData As Integer)
Attribute EnCount.VB_Description = "Переменная-счетчик количества потерпевших/владельцев вещественных доказательств"
'Счетчик
    mvarEnCount = vData
End Property
'
Friend Property Get EnCount() As Integer
'Счетчик.
    EnCount = mvarEnCount
'Debug.Print "Счетчик= " & EnCount
End Property
'
Friend Function Print_Owner() As String
'
End Function
'
Friend Function Print_InjPersons() As String
'Функция печати текста "Потерпевший"
    If Val(En_Birthday) = 0 Then
        Print_InjPersons = En_SurName & Chr(32) & En_Name & En_MidName
    Else: Print_InjPersons = En_SurName & Chr(32) & En_Name & En_MidName _
                    & Chr(32) & En_Birthday & " г.р."
    End If
End Function
'
Friend Function Print_Autopsy() As String
'Функция печати данных по вскрытию
    Print_Autopsy = "... материал от трупа " & Print_InjPersons & " Дата смерти: " & En_Decease & _
                        " Дата забора материала: " & En_Autopsy
End Function
'
Friend Function Create_InitialsMidName() As String
'Функция создания инициала отчества
    Create_InitialsMidName = Left(En_MidName, 1) & "."
End Function

Friend Function Create_InitialsName() As String
'Функция создания инициала имени
    Create_InitialsName = Left(En_Name, 1) & "."
End Function
'
Friend Function Create_InitialsFullName() As String
'Функция создания фамилии и инициалов потерпевшего
    Create_InitialsFullName = En_SurName & Chr(32) & Create_InitialsName() _
        & Create_InitialsMidName()
End Function
'
Friend Function Create_FullName() As String
'Функция создания Фамилии, полных Имени и Отчества.
    Create_FullName = En_SurName & Chr(32) & En_Name & Chr(32) & En_MidName
End Function
'
Friend Property Let En_fromHospital(ByVal vData As Date)
'Дата выписки из ЛПУ
    mvarEn_fromHospital = vData
End Property
'
Friend Property Get En_fromHospital() As Date
'Дата выписки из ЛПУ
    En_fromHospital = mvarEn_fromHospital
'Debug.Print "'Дата выписки из ЛПУ = " & En_fromHospital
End Property
'
Friend Property Let En_toHospital(ByVal vData As Date)
'Дата госпитализации
    mvarEn_toHospital = vData
End Property
'
Friend Property Get En_toHospital() As Date
'Дата госпитализации
    En_toHospital = mvarEn_toHospital
'Debug.Print "'Дата госпитализации = " & En_toHospital
End Property
'
Friend Property Let En_Comments(ByVal vData As String)
'Коментарии
    mvarEn_Comments = vData
End Property
'
Friend Property Get En_Comments() As String
'Коментарии
    En_Comments = mvarEn_Comments
'Debug.Print "Коментарии = " & En_Comments
End Property
'
Friend Property Let En_PI(ByVal vData As String)
'Персональные данные
    mvarEn_PI = vData
End Property
'
Friend Property Get En_PI() As String
'Персональные данные
    En_PI = mvarEn_PI
'Debug.Print "Персональные данные = " & En_PI
End Property
'
Friend Property Let En_PreDiagnosis(ByVal vData As String)
'Предварительный диагноз
    mvarEn_PreDiagnosis = vData
End Property
'
Friend Property Get En_PreDiagnosis() As String
'Предварительный диагноз
    En_PreDiagnosis = mvarEn_PreDiagnosis
'Debug.Print "Предварительный диагноз = " & En_PreDiagnosis
End Property
'
Friend Property Let En_Autopsy(ByVal vData As Date)
'Дата вскрытия
    mvarEn_Autopsy = vData
End Property
'
Friend Property Get En_Autopsy() As Date
'Дата вскрытия
    En_Autopsy = mvarEn_Autopsy
'Debug.Print "Дата вскрытия = " & En_Autopsy
End Property
'
Friend Property Let En_Decease(ByVal vData As Date)
'Дата смерти
    mvarEn_Decease = vData
End Property
'
Friend Property Get En_Decease() As Date
'Дата смерти
    En_Decease = mvarEn_Decease
'Debug.Print "Дата смерти = " & En_Decease
End Property
'
Friend Property Let En_Birthday(ByVal vData As String)
'Дата/год рождения
    mvarEn_Birthday = vData
End Property
'
Friend Property Get En_Birthday() As String
'Дата/год рождения
    En_Birthday = mvarEn_Birthday
'Debug.Print "Дата/год рождения владельца = " & En_Birthday
End Property
'
Friend Property Let En_MidName(ByVal vData As String)
'Отчество
    mvarEn_MidName = vData
End Property
'
Friend Property Get En_MidName() As String
'Отчество
    En_MidName = mvarEn_MidName
'Debug.Print "Отчество владельца = " & En_MidName
End Property
'
Friend Property Let En_Name(ByVal vData As String)
'Имя
    mvarEn_Name = vData
End Property
'
Friend Property Get En_Name() As String
'Имя.
    En_Name = mvarEn_Name
'Debug.Print "Имя владельца = " & En_Name
End Property
'
Friend Property Let En_SurName(ByVal vData As String)
'Отчество
    mvarEn_SurName = vData
End Property
'
Friend Property Get En_SurName() As String
'Отчество
    En_SurName = mvarEn_SurName
'Debug.Print "Отчество владельца = " & En_SurName
End Property
'
Friend Property Let En_Sex(ByVal vData As String)
'Пол.
    mvarEn_Sex = vData
End Property
'
Friend Property Get En_Sex() As String
'Пол.
    En_Sex = mvarEn_Sex
'Debug.Print "Пол владельца= " & En_Sex
End Property
'
Private Sub Class_Initialize()
EnCount = EnCount + 1 'Инициализация счетчика (кол-во владельцев ВД)
End Sub
