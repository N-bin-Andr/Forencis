VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CrimExperts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = " Эксперты	-CrimExperts (CE)\r\nCEID		-Счетчик\r\nCESex		-Пол\r\nCEPost		-Должность\r\nCESurName	-Фамилия\r\nCEName		-Имя\r\nCEMidName	-Отчество\r\nCEConformation	-Структура\r\nCEDepartment	-Отдел/Управление ГКСЭ РБ\r\nCEPersInf	-Контактыне данные (Personal Information)\r\nCEComment	-Примечания\r\nCEFoto		-Фото эксперта\r\n"
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'local variable(s) to hold property value(s)
Private mvarCESex As Boolean 'local copy
Private mvarCEPost As String 'local copy
Private mvarCESurName As String 'local copy
Private mvarCEName As String 'local copy
Private mvarCEMidName As String 'local copy
Private mvarCEConformation As String 'local copy
Private mvarCEDepartment As String 'local copy
Private mvarCEPersInf As String 'local copy
Private mvarCEComment As String 'local copy
Public Property Let CEComment(ByVal vData As String)
Attribute CEComment.VB_Description = "Примечания (Эксперты)"
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.CEComment = 5
    mvarCEComment = vData
End Property


Public Property Get CEComment() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.CEComment
    CEComment = mvarCEComment
End Property



Public Property Let CEPersInf(ByVal vData As String)
Attribute CEPersInf.VB_Description = "Контактные данные (Эксперты)"
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.CEPersInf = 5
    mvarCEPersInf = vData
End Property


Public Property Get CEPersInf() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.CEPersInf
    CEPersInf = mvarCEPersInf
End Property



Public Property Let CEDepartment(ByVal vData As String)
Attribute CEDepartment.VB_Description = "Отдел / управление (Эксперты)"
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.CEDepartment = 5
    mvarCEDepartment = vData
End Property


Public Property Get CEDepartment() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.CEDepartment
    CEDepartment = mvarCEDepartment
End Property



Public Property Let CEConformation(ByVal vData As String)
Attribute CEConformation.VB_Description = "Структура (подразделение) (Эксперты)"
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.CEConformation = 5
    mvarCEConformation = vData
End Property


Public Property Get CEConformation() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.CEConformation
    CEConformation = mvarCEConformation
End Property



Public Property Let CEMidName(ByVal vData As String)
Attribute CEMidName.VB_Description = "Отчество (Эксперты)"
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.CEMidName = 5
    mvarCEMidName = vData
End Property


Public Property Get CEMidName() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.CEMidName
    CEMidName = mvarCEMidName
End Property



Public Property Let CEName(ByVal vData As String)
Attribute CEName.VB_Description = "Имя (Эксперты)"
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.CEName = 5
    mvarCEName = vData
End Property


Public Property Get CEName() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.CEName
    CEName = mvarCEName
End Property



Public Property Let CESurName(ByVal vData As String)
Attribute CESurName.VB_Description = "Фамилия (Эксперты)"
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.CESurName = 5
    mvarCESurName = vData
End Property


Public Property Get CESurName() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.CESurName
    CESurName = mvarCESurName
End Property



Public Property Let CEPost(ByVal vData As String)
Attribute CEPost.VB_Description = "Должность (Эксперты)"
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.CEPost = 5
    mvarCEPost = vData
End Property


Public Property Get CEPost() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.CEPost
    CEPost = mvarCEPost
End Property



Public Property Let CESex(ByVal vData As Boolean)
Attribute CESex.VB_Description = "Пол (Эксперты)"
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.CESex = 5
    mvarCESex = vData
End Property


Public Property Get CESex() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.CESex
    CESex = mvarCESex
End Property



