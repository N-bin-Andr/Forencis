VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clmCoroner"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'Класс "Следователь"
'Цель: работа с программами VBA или базами данных Access
'Обработка данных при создании медико-криминалистических заключений
'Версия: 1.0
'Дата создания: 01.06.2016
'Автор: Набебин Андрей Владиславович
Option Explicit
Private mvarstrCrSex As String          'Пол следователя
Private mvarstrCrSurName As String      'Фамилия следователя
Private mvarstrCrName As String         'Имя следователя
Private mvarstrCrMidName As String      'Отчество следователя
Private mvarstrCrPost As String         'Должность следователя
Private mvarstrCrOprAr As String        'Район обслуживания следователя
Private mvarstrCrRank As String         'Звание следователя
Private mvarstrCrPI As String           'Контактные данные следователя
Private mvarstrCrComments As String     'Примечания
Private mvarstrCrConformation As String 'Структура подразделения
Private mvarstrCrOffice As String       'Управление подразделения
Private mvarstrCrDepartment As String   'Отдел подразделения
'
Friend Function Create_CrFullName() As String
Attribute Create_CrFullName.VB_Description = "Функция создания Фамилии, полных Имени и Отчества следователя"
'Функция создания Фамилии, полных Имени и Отчества следователя
    Create_CrFullName = strCrSurName & Chr(32) & strCrName & Chr(32) & strCrMidName
End Function
'
Friend Function Create_InitialsCrMidName(ByVal strCrMidName As String) As String
'Функция создания инициала отчества следователя
    Create_InitialsCrMidName = Left(strCrMidName, 1) & "."
End Function
'
Friend Function Create_InitialsCrName(ByVal strCrName As String) As String
'Функция создания инициала имени следователя
    Create_InitialsCrName = Left(strCrName, 1) & "."
End Function
'
Friend Function Create_InitialsCrFullName() As String
'Функция создания фамилии и инициалов следователя
Create_InitialsCrFullName = strCrSurName & Chr(32) & Create_InitialsCrName(ByVal strCrName) _
& Create_InitialsCrMidName(ByVal strCrMidName)
End Function
'
Friend Function Print_Coroner()
'Функция набора текста "Следователь"
Print_Coroner = strCrPost & Chr(32) & strCrOprAr & Chr(32) & strCrRank & Chr(32) & Create_InitialsCrFullName
End Function
'
Friend Property Let strCrDepartment(ByVal vData As String)
'Отдел
    mvarstrCrDepartment = vData
End Property
'
Friend Property Get strCrDepartment() As String
'Отдел
    strCrDepartment = mvarstrCrDepartment
Debug.Print "Отдел= ", strCrDepartment
End Property
'
Friend Property Let strCrOffice(ByVal vData As String)
'Управление
    mvarstrCrOffice = vData
End Property
'
Friend Property Get strCrOffice() As String
'Управление
    strCrOffice = mvarstrCrOffice
Debug.Print "Управление= ", strCrOffice
End Property
'
Friend Property Let strCrConformation(ByVal vData As String)
'Структура
    mvarstrCrConformation = vData
End Property
'
Friend Property Get strCrConformation() As String
'Структура
    strCrConformation = mvarstrCrConformation
Debug.Print "Структура= ", strCrConformation
End Property
'
Friend Property Let strCrComments(ByVal vData As String)
'Примечания
    mvarstrCrComments = vData
End Property

Friend Property Get strCrComments() As String
'Примечания
    strCrComments = mvarstrCrComments
Debug.Print "Примечания = ", strCrComments
End Property
'
Friend Property Let strCrPI(ByVal vData As String)
'Контактные данные
    mvarstrCrPI = vData
End Property
'
Friend Property Get strCrPI() As String
'Контактные данные
    strCrPI = mvarstrCrPI
Debug.Print "Контактные данные = ", strCrPI
End Property
'
Friend Property Let strCrRank(ByVal vData As String)
'Звание следователя
    mvarstrCrRank = vData
End Property
'
Friend Property Get strCrRank() As String
'Звание следователя
    strCrRank = mvarstrCrRank
Debug.Print "Звание следователя= ", strCrRank
End Property
'
Friend Property Let strCrOprAr(ByVal vData As String)
'Район обслуживания
    mvarstrCrOprAr = vData
End Property
'
Friend Property Get strCrOprAr() As String
'Район обслуживания
    strCrOprAr = mvarstrCrOprAr
Debug.Print "Район обслуживания= ", strCrOprAr
End Property
'
Friend Property Let strCrPost(ByVal vData As String)
'Должность следователя
    mvarstrCrPost = vData
End Property
'
Friend Property Get strCrPost() As String
'Должность следователя
    strCrPost = mvarstrCrPost
Debug.Print "Должность следователя= ", strCrPost
End Property
'
Friend Property Let strCrMidName(ByVal vData As String)
'Отчество следователя
    mvarstrCrMidName = vData
End Property
'
Friend Property Get strCrMidName() As String
'Отчество следователя
    strCrMidName = mvarstrCrMidName
Debug.Print "Отчество следователя= ", strCrMidName
End Property
'
Friend Property Let strCrName(ByVal vData As String)
'Имя следователя
    mvarstrCrName = vData
End Property
'
Friend Property Get strCrName() As String
'Имя следователя
    strCrName = mvarstrCrName
Debug.Print "Имя следователя= ", strCrName
End Property
'
Friend Property Let strCrSurName(ByVal vData As String)
'Фамилия следователя
    mvarstrCrSurName = vData
End Property
'
Friend Property Get strCrSurName() As String
'Фамилия следователя
    strCrSurName = mvarstrCrSurName
Debug.Print "Фамилия следователя= ", strCrSurName
End Property
'
Friend Property Let strCrSex(ByVal vData As String)
'Пол следователя
    mvarstrCrSex = vData
End Property

Friend Property Get strCrSex() As String
'Пол следователя
    strCrSex = mvarstrCrSex
Debug.Print "Пол следователя= ", strCrSex
End Property
