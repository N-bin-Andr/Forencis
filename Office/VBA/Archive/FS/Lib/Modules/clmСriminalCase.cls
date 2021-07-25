VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "СriminalCase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = " Дело       -Case (Cs)\r\nCsID              -ключ (счетчик)\r\nCsPrvDate    -Дата вынесения постановления (ProvisionDate)\r\nCsCat          -Категория дела\r\nCsNum                -Номер дела\r\nCsCoroner        -Следователи\r\nCsIP            -Потерпевшие (InjuredPersons )\r\nCsSpP         -Подозреваемый (SuspectPerson)\r\nCsEFNum       -Номер СМЭ\r\nCsCatDamage       -Категория травмы\r\nCsComment  -Примечания\r\n"
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'local variable(s) to hold property value(s)
Private mvarCsPrvDate As Date 'local copy
Private mvarCsCat As String 'local copy
Private mvarCsNum As String 'local copy
Private mvarCsCoroner As String 'local copy
Private mvarCsIP As String 'local copy
Private mvarCsSpP As String 'local copy
Private mvarCsEFNum As String 'local copy
Private mvarCsCatDamage As String 'local copy
Private mvarCsComment As String 'local copy
'
Public Property Let CsComment(ByVal vData As String)
Attribute CsComment.VB_Description = "Примечания (Дело)"
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.CsComment = 5
    mvarCsComment = vData
End Property
'
Public Property Get CsComment() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.CsComment
    CsComment = mvarCsComment
End Property
'
Public Property Let CsCatDamage(ByVal vData As String)
Attribute CsCatDamage.VB_Description = "Категория травмы (Дело)"
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.CsCatDamage = 5
    mvarCsCatDamage = vData
End Property
'
Public Property Get CsCatDamage() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.CsCatDamage
    CsCatDamage = mvarCsCatDamage
End Property
'
Public Property Let CsEFNum(ByVal vData As String)
Attribute CsEFNum.VB_Description = "Номер экспертизы (Дело)"
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.CsEFNum = 5
    mvarCsEFNum = vData
End Property
'
Public Property Get CsEFNum() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.CsEFNum
    CsEFNum = mvarCsEFNum
End Property
'
Public Property Let CsSpP(ByVal vData As String)
Attribute CsSpP.VB_Description = "Подозреваемый(-ые) (Дело)"
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.CsSpP = 5
    mvarCsSpP = vData
End Property
'
Public Property Get CsSpP() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.CsSpP
    CsSpP = mvarCsSpP
End Property
'
Public Property Let CsIP(ByVal vData As String)
Attribute CsIP.VB_Description = "Потерпевший(-ие) (Дело)"
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.CsIP = 5
    mvarCsIP = vData
End Property
'
Public Property Get CsIP() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.CsIP
    CsIP = mvarCsIP
End Property
'
Public Property Let CsCoroner(ByVal vData As String)
Attribute CsCoroner.VB_Description = "Следователь (-ли) (Дело)"
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.CsCoroner = 5
    mvarCsCoroner = vData
End Property
'
Public Property Get CsCoroner() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.CsCoroner
    CsCoroner = mvarCsCoroner
End Property
'
Public Property Let CsNum(ByVal vData As String)
Attribute CsNum.VB_Description = "Номер (Дело)"
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.CsNum = 5
    mvarCsNum = vData
End Property
'
Public Property Get CsNum() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.CsNum
    CsNum = mvarCsNum
End Property
'
Public Property Let CsCat(ByVal vData As String)
Attribute CsCat.VB_Description = "Категория (Дело)"
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.CsCat = 5
    mvarCsCat = vData
End Property
'
Public Property Get CsCat() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.CsCat
    CsCat = mvarCsCat
End Property
'
Public Property Let CsPrvDate(ByVal vData As Date)
Attribute CsPrvDate.VB_Description = "Дата вынесения постановления (Дело)"
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.CsPrvDate = 5
    mvarCsPrvDate = vData
End Property
'
Public Property Get CsPrvDate() As Date
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.CsPrvDate
    CsPrvDate = mvarCsPrvDate
End Property
