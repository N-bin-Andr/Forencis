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
Option Explicit
'
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
Friend Function Create_InitialsCEMidName(ByVal strCEMidName As String) As String
Attribute Create_InitialsCEMidName.VB_Description = "Функция создания инициала отчества эксперта"
'Функция создания инициала отества эксперта
    Create_InitialsCEMidName = Left(strCEMidName, 1) & "."
End Function
'
Friend Function Create_InitialsCEName(ByVal strCEName As String) As String
Attribute Create_InitialsCEName.VB_Description = "Функция создания инициала имени эксперта"
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
Friend Function print_Expert() As String
Attribute print_Expert.VB_Description = "Функция набора текста """"Общий эксперт"""
'Функция набота текста направление "Общего эксперта"
'    Проверка на условие:
    If Val(Create_CERank(strCEPost, strCERank)) = 0 Then
       print_Expert = " Государственного медицинского судебного эксперта " & strCESurName & Chr(32) & _
            strCEName & Chr(32) & strCEMidName
    Else: print_Expert = " Государственного медицинского судебного эксперта " & _
    Create_CERank(ByVal strCEPost, ByVal strCERank) & Chr(32) & strCESurName & Chr(32) & _
            strCEName & Chr(32) & strCEMidName
    End If
End Function
'
Friend Function Create_CERank(ByVal strCEPost As String, ByVal strCERank As String) As String
Attribute Create_CERank.VB_Description = "Функция создания строки: должность и звание эксперта"
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
Attribute strCEComment.VB_Description = "Примечания (эксперт)"
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
Attribute strCEPI.VB_Description = "Персональные данные"
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
Attribute strCEDepartment.VB_Description = "Отдел"
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
Attribute strCEOffice.VB_Description = "Управление (подразделения)"
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
Attribute strCEConformation.VB_Description = "Структура (подазделения)"
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
Attribute strCEName.VB_Description = "Имя эксперта"
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
Attribute strCESurName.VB_Description = "Фамилия эксперта"
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
Attribute strCERank.VB_Description = "Звание эксперта"
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
Attribute strCEPost.VB_Description = "Должность эксперта"
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
Attribute strCESex.VB_Description = "Пол эксперта"
'Пол эксперта
    mvarstrCESex = vData
End Property
'
Friend Property Get strCESex() As String
'Пол эксперта
    strCESex = mvarstrCESex
Debug.Print "Пол эксперта= ", strCESex
End Property
