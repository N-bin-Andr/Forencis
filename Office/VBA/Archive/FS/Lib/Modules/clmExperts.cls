VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clmExperts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "Класс-модуль ""Эксперт"""
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'Класс "Эксперт(-ы)"
'Цель: работа с программами VBA или базами данных Access
'Обработка данных при создании медико-криминалистических заключений
'Версия: 1.0
'Дата создания: 01.06.2016
'Автор: Набебин Андрей Владиславович
Option Explicit
Private mvarstrCESex As String          'Пол эксперта
Private mvarstrCEPost As String         'Должность эксперта
Private mvarstrCERank As String         'Звание эксперта
Private mvarstrCESurName As String      'Фамилия эксперта
Private mvarstrCEName As String         'Имя эксперта
Private mvarstrCEMidName As String      'Отчество эксперта
Private mvarstrCEConformation As String 'Структура (где работает)
Private mvarstrCEOffice As String       'Управление (где работает)
Private mvarstrCEDepartment As String   'Отдел (где работает)
Private mvarstrCEPI As String           'Контактные данные эксперта
Private mvarstrCEComment As String      'Примечания
'
Friend Function Create_CEFullName() As String
Attribute Create_CEFullName.VB_Description = "Функция создания Фамилии, полных Имени и Отчества эксперта."
'Функция создания Фамилии, полных Имени и Отчества эксперта.
    Create_CEFullName = strCESurName & Chr(32) & strCEName & Chr(32) & strCEMidName
End Function

Friend Function Calculate_Experience() As String
'Функция расчета стажа работы
Dim bytEx1 As Byte, strEx2 As String
strEx2 = "01.08.2005"
    bytEx1 = DateDiff("yyyy", CDate(strEx2), Now) 'расчет стажа работы
    Select Case bytEx1
        Case Is = 1
            Calculate_Experience = bytEx1 & " год,"
        Case 2 To 4
            Calculate_Experience = bytEx1 & " года,"
        Case 5 To 20
            Calculate_Experience = bytEx1 & " лет,"
          Case Else
            strEx2 = Right(bytEx1, 1)
                Select Case CByte(strEx2)
                    Case Is = 0
                        Calculate_Experience = bytEx1 & " лет,"
                    Case Is = 1
                        Calculate_Experience = bytEx1 & " год,"
                    Case 2 To 4
                        Calculate_Experience = bytEx1 & " года,"
                    Case 5 To 9
                        Calculate_Experience = bytEx1 & " лет,"
                End Select
    End Select
End Function
'
Friend Function Create_InitialsCEMidName(ByVal strCEMidName As String) As String
'Функция создания инициала отества эксперта
    Create_InitialsCEMidName = Left(strCEMidName, 1) & "."
End Function
'
Friend Function Create_InitialsCEName(ByVal strCEName As String) As String
'Функция создания инициала имени эксперта
    Create_InitialsCEName = Left(strCEName, 1) & "."
End Function
'
Friend Function Create_InitialsCEFullName() As String
'Функция создания фамилии и инициалов эксперта
Create_InitialsCEFullName = strCESurName & Chr(32) & Create_InitialsCEName(ByVal strCEName) _
& Create_InitialsCEMidName(ByVal strCEMidName)
End Function
'
Friend Function Print_Expert() As String
'Функция набота текста направление "Общего эксперта"
'    Проверка на условие:
    If Create_CERank(ByVal strCEPost, ByVal strCERank) = "" Then
       Print_Expert = "Государственного медицинского судебного эксперта " & Create_InitialsCEFullName
    Else: Print_Expert = "Государственного медицинского судебного эксперта " & _
    Create_CERank(ByVal strCEPost, ByVal strCERank) & Chr(32) & Create_InitialsCEFullName
    End If
End Function
'
Friend Function Create_CERank(ByVal strCEPost As String, ByVal strCERank As String) As String
'Функция создания текста должности и звания Эксперта
'   Проверка на условие
    If Val(strCEPost) = 0 Then
        Create_CERank = ""
    ElseIf Val(strCERank) = 0 Then
        Create_CERank = ""
    Else: Create_CERank = strCEPost & Chr(32) & strCERank
    End If
End Function
'
Friend Property Let strCEComment(ByVal vData As String)
'Примечания.
    mvarstrCEComment = vData
End Property
'
Friend Property Get strCEComment() As String
'Примечания.
    strCEComment = mvarstrCEComment
Debug.Print "Примечания= ", strCEComment
End Property
'
Friend Property Let strCEPI(ByVal vData As String)
'Контактные данные.
    mvarstrCEPI = vData
End Property
'
Friend Property Get strCEPI() As String
'Контактные данные.
    strCEPI = mvarstrCEPI
Debug.Print "Контактные данные= ", strCEPI
End Property
'
Friend Property Let strCEDepartment(ByVal vData As String)
'Отдел.
    mvarstrCEDepartment = vData
End Property
'
Friend Property Get strCEDepartment() As String
'Отдел.
    strCEDepartment = mvarstrCEDepartment
Debug.Print "Отдел= ", strCEDepartment
End Property
'
Friend Property Let strCEOffice(ByVal vData As String)
'Управление.
    mvarstrCEOffice = vData
End Property
'
Friend Property Get strCEOffice() As String
'Управление.
    strCEOffice = mvarstrCEOffice
Debug.Print "Управление= ", strCEOffice
End Property
'
Friend Property Let strCEConformation(ByVal vData As String)
'Структура
mvarstrCEConformation = vData
End Property
'
Friend Property Get strCEConformation() As String
'Структура
    strCEConformation = mvarstrCEConformation
Debug.Print "Структура= ", strCEConformation
End Property
'
Friend Property Let strCEMidName(ByVal vData As String)
'Отчество эксперта
    mvarstrCEMidName = vData
End Property
'
Friend Property Get strCEMidName() As String
'Отчество эксперта
    strCEMidName = mvarstrCEMidName
Debug.Print "Отчество эксперта= ", strCEMidName
End Property
'
Friend Property Let strCEName(ByVal vData As String)
'Имя эксперта
    mvarstrCEName = vData
End Property
'
Friend Property Get strCEName() As String
'Имя эксперта
    strCEName = mvarstrCEName
Debug.Print "Имя эксперта= ", strCEName
End Property
'
Friend Property Let strCESurName(ByVal vData As String)
'Фамилия эксперта
    mvarstrCESurName = vData
End Property
'
Friend Property Get strCESurName() As String
'Фамилия эксперта
    strCESurName = mvarstrCESurName
Debug.Print "Фамилия эксперта= ", strCESurName
End Property
'
Friend Property Let strCERank(ByVal vData As String)
'Звание эксперта
    mvarstrCERank = vData
End Property
'
Friend Property Get strCERank() As String
'Звание эксперта
    strCERank = mvarstrCERank
Debug.Print "Звание эксперта=", strCERank
End Property
'
Friend Property Let strCEPost(ByVal vData As String)
'Должность эксперта
    mvarstrCEPost = vData
End Property
'
Friend Property Get strCEPost() As String
'Должность эксперта
    strCEPost = mvarstrCEPost
Debug.Print "Должность эксперта= ", strCEPost
End Property
'
Friend Property Let strCESex(ByVal vData As String)
'Пол эксперта
    mvarstrCESex = vData
End Property
'
Friend Property Get strCESex() As String
'Пол эксперта
    strCESex = mvarstrCESex
Debug.Print "Пол эксперта= ", strCESex
End Property
'
Private Sub Class_Initialize()

End Sub
