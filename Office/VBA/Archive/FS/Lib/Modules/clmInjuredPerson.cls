VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clmInjuredPerson"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "Класс-модуль ""Потерпевший"" - описывает соответствующие  свойства и методы"
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'Класс "Потерпевший(-ие)"
'Цель: работа с программами VBA или базами данных Access
'Обработка данных при создании медико-криминалистических заключений
'Версия: 1.0
'Дата создания: 01.06.2016
'Автор: Набебин Андрей Владиславович
Option Explicit
Private mvarstrInjPrSex As String           'Пол потерпевшего
Private mvarstrInjPrSurName As String       'Фамилия потерпевшего
Private mvarstrInjPrName As String          'Имя потерпевшего
Private mvarstrInjPrMidName As String       'Отчество потерпевшего
Private mvarstrInjPrBirthday As Variant     'Дата (год) рождения
Private mvarDtmInjPrDecease As Date         'Дата смерти
Private mvarDtmInjPrAutopsy As Date         'Дата вскрытия
Private mvarstrInjPrPreDiagnosis As String  'Предполагаемый диагноз
Private mvarstrInjPrPI As String            'Персональные данные потерпевшего
Private mvarstrComments As String           'Примечания
Private mvarDtmInjPrDatetoHospital As Date	'Дата поступления в лечебное учреждение
Private mvarDtmInjPrDatefrHospital As Date	'Дата выписки из лечебного учреждения
'
Friend Function Create_InjPrFullName() As String
Attribute Create_InjPrFullName.VB_Description = "Функция создания Фамилии, полных Имени и Отчества потерпевшего."
'Функция создания Фамилии, полных Имени и Отчества потерпевшего.
    If Len(strInjPrSurName) <> 0 Then
     Create_InjPrFullName = strInjPrSurName & Chr(32) & strInjPrName & Chr(32) & strInjPrMidName
    Else: Create_InjPrFullName = "(нет полных данных о владельце объекта(ов))"
    End If
End Function
'
Friend Function Create_InitialsInjPrFullName() As String
'Функция создания фамилии и инициалов потерпевшего
    If Len(strInjPrSurName) <> 0 Then
        Create_InitialsInjPrFullName = strInjPrSurName & Chr(32) & Create_InitialsInjPrName(ByVal strInjPrName) _
        & Create_InitialsInjPrMidName(ByVal strInjPrMidName)
    Else: Create_InitialsInjPrFullName = "(нет полных данных о владельце объекта(ов))"
    End If
End Function
'
Friend Function Create_InitialsInjPrName(ByVal strInjPrName As String) As String
'Функция создания инициала имени потерпевшего
    Create_InitialsInjPrName = Left(strInjPrName, 1) & "."
End Function
'
Friend Function Create_InitialsInjPrMidName(ByVal strInjPrMidName As String) As String
'Функция создания инициала отечтва потерпевшего
    Create_InitialsInjPrMidName = Left(strInjPrMidName, 1) & "."
End Function
'
Friend Function Print_InjPrAutopsy() As String
'Функция печати данных по вскрытию
    Print_InjPrAutopsy = "... материал от трупа " & Print_InjPr & " Дата смерти: " & DtmInjPrDecease & _
                        " Дата забора материала: " & DtmInjPrAutopsy
End Function
'
Friend Function Print_InjPr() As String
'Функция печати текста "Потерпевший"
    If Val(strInjPrBirthday) = 0 Then
        Print_InjPr = strInjPrSurName & Chr(32) & strInjPrName & strInjPrMidName
    Else: Print_InjPr = strInjPrSurName & Chr(32) & strInjPrName & strInjPrMidName _
                    & Chr(32) & strInjPrBirthday & " г.р."
    End If
End Function
'
Friend Property Let strComments(ByVal vData As String)
'Примечания
    mvarstrComments = vData
End Property
'
Friend Property Get strComments() As String
'Примечания
    strComments = mvarstrComments
'Debug.Print "Примечания= ", strComments
End Property
'
Friend Property Let strInjPrPI(ByVal vData As String)
'Персональные данные.
    mvarstrInjPrPI = vData
End Property
'
Friend Property Get strInjPrPI() As String
'Персональные данные.
    strInjPrPI = mvarstrInjPrPI
'Debug.Print "Персональные данные= ", strInjPrPI
End Property
'
Friend Property Let strInjPrPreDiagnosis(ByVal vData As String)
'Предварительный диагноз.
    mvarstrInjPrPreDiagnosis = vData
End Property
'
Friend Property Get strInjPrPreDiagnosis() As String
'Предварительный диагноз.
    strInjPrPreDiagnosis = mvarstrInjPrPreDiagnosis
'Debug.Print "Предварительный диагноз= ", strInjPrPreDiagnosis
End Property
'
Friend Property Let DtmInjPrAutopsy(ByVal vData As Date)
'Дата вскрытия.
    mvarDtmInjPrAutopsy = vData
End Property
'
Friend Property Get DtmInjPrAutopsy() As Date
'Дата вскрытия.
    DtmInjPrAutopsy = mvarDtmInjPrAutopsy
'Debug.Print "Дата вскрытия= ", DtmInjPrAutopsy
End Property
'
Friend Property Let DtmInjPrDecease(ByVal vData As Date)
'Дата смерти.
    mvarDtmInjPrDecease = vData
End Property
'
Friend Property Get DtmInjPrDecease() As Date
'Дата смерти.
    DtmInjPrDecease = mvarDtmInjPrDecease
'Debug.Print "Дата смерти= ", DtmInjPrDecease
End Property
'
Friend Property Let strInjPrBirthday(ByVal vData As String)
'Дата (год) рождения.
    mvarstrInjPrBirthday = vData
End Property
'
Friend Property Get strInjPrBirthday() As String
'Дата (год) рождения.
    strInjPrBirthday = mvarstrInjPrBirthday
'Debug.Print "Дата (год) рождения= ", strInjPrBirthday
End Property
'
Friend Property Let strInjPrMidName(ByVal vData As String)
'Фамилия потерпевшего.
    mvarstrInjPrMidName = vData
End Property
'
Friend Property Get strInjPrMidName() As String
'Фамилия потерпевшего.
    strInjPrMidName = mvarstrInjPrMidName
'Debug.Print "Фамилия потерпевшего= ", strInjPrMidName
End Property
'
Friend Property Let strInjPrName(ByVal vData As String)
'Имя потерпевшего.
    mvarstrInjPrName = vData
End Property
'
Friend Property Get strInjPrName() As String
'Имя потерпевшего.
'Debug.Print "Имя потерпевшего= ", strInjPrName
    strInjPrName = mvarstrInjPrName
End Property
'
Friend Property Let strInjPrSurName(ByVal vData As String)
'Отчество потерпевшего.
    mvarstrInjPrSurName = vData
End Property
'
Friend Property Get strInjPrSurName() As String
'Отчество потерпевшего.
    strInjPrSurName = mvarstrInjPrSurName
'Debug.Print "Отчество потерпевшего= ", strInjPrSurName
End Property
'
Friend Property Let strInjPrSex(ByVal vData As String)
'Пол потерпевшего.
    mvarstrInjPrSex = vData
End Property
'
Friend Property Get strInjPrSex() As String
'Пол потерпевшего.
    strInjPrSex = mvarstrInjPrSex
'Debug.Print "Пол потерпевшего= ", strInjPrSex
End Property
