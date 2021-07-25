VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Owner"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = " Владелец ВД (принадлежность ВД)\r\nOwSex		-пол (М/Ж/нет данных)\r\nOwSurName	-Фамилия\r\nOwName		-Имя\r\nOwMidName	-Отчество\r\nOwBirthday	-Год рождения\r\nOwEvidences	-Принадлежащие ВД\r\nOwPI		-Контактные данные\r\nOwComments	-Коментарии "
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mvarOwSex As Boolean 'local copy
Private mvarOwSurName As String 'local copy
Private mvarOwName As String 'local copy
Private mvarOwMidName As String 'local copy
Private mvarOwBirthday As String 'local copy
Private mvarOwEvidences As String 'local copy
Private mvarOwPI As String 'local copy
Private mvarOwComments As String 'local copy
Public Property Let OwComments(ByVal vData As String)
Attribute OwComments.VB_Description = "Коментарии (Владелец ВД)"
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.OwComments = 5
    mvarOwComments = vData
End Property

Public Property Get OwComments() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.OwComments
    OwComments = mvarOwComments
End Property



Public Property Let OwPI(ByVal vData As String)
Attribute OwPI.VB_Description = "Контактные данные (Владелец ВД)"
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.OwPI = 5
    mvarOwPI = vData
End Property


Public Property Get OwPI() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.OwPI
    OwPI = mvarOwPI
End Property



Public Property Let OwEvidences(ByVal vData As String)
Attribute OwEvidences.VB_Description = "Принадлежащие ВД (Владелец ВД)"
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.OwEvidences = 5
    mvarOwEvidences = vData
End Property


Public Property Get OwEvidences() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.OwEvidences
    OwEvidences = mvarOwEvidences
End Property



Public Property Let OwBirthday(ByVal vData As String)
Attribute OwBirthday.VB_Description = "Год рождения (Владелец ВД)"
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.OwBirthday = 5
    mvarOwBirthday = vData
End Property


Public Property Get OwBirthday() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.OwBirthday
    OwBirthday = mvarOwBirthday
End Property



Public Property Let OwMidName(ByVal vData As String)
Attribute OwMidName.VB_Description = "Отчество (Владелец ВД)"
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.OwMidName = 5
    mvarOwMidName = vData
End Property


Public Property Get OwMidName() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.OwMidName
    OwMidName = mvarOwMidName
End Property



Public Property Let OwName(ByVal vData As String)
Attribute OwName.VB_Description = "Имя (Владелец ВД)"
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.OwName = 5
    mvarOwName = vData
End Property


Public Property Get OwName() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.OwName
    OwName = mvarOwName
End Property



Public Property Let OwSurName(ByVal vData As String)
Attribute OwSurName.VB_Description = "Фамилия (Владелец ВД)"
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.OwSurName = 5
    mvarOwSurName = vData
End Property


Public Property Get OwSurName() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.OwSurName
    OwSurName = mvarOwSurName
End Property



Public Property Let OwSex(ByVal vData As Boolean)
Attribute OwSex.VB_Description = "Пол (Владелец ВД)"
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.OwSex = 5
    mvarOwSex = vData
End Property


Public Property Get OwSex() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.OwSex
    OwSex = mvarOwSex
End Property



