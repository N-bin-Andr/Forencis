VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clmCase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "Класс ""Дело"" - общий класс для данных, относящихся к производству уголовного или административного дела."
'Класс "Дело"
'@author Andr.N@_bin
'@E-mail Andr.Nab.n@gmail.com
Option Explicit
Private tmpid As Long               'идентификатор (количество созданных объектов)
Private tmpcategory As String       'Категория Дела: уголовное, административное
Private tmpnumber As String         'Номер Дела
Private tmpcrimCondition As String  'Обстоятельства дела
Private tmpdamage As String         'Категория травмы
Private tmpсomments As String       'Примечания
Private tmpindex As String          'Индекс Дела
'
Private Sub Class_Initialize()
'конструктор
    id = id + 1
End Sub
'
Friend Function printCsData() As String
Attribute printCsData.VB_Description = "Печать данных ""основания назначения экспертизы"":"
'Функция печати данных по Делу
'строка = по материалам уголовного/административного дела ИНДЕКС № НОМЕР
    If index = "" Then
        If category = "" Then
            printCsData = "(данные по производству (делу) не указаны)"
        ElseIf number = "" Then
            printCsData = "(данные по производству (делу) не указаны)"
        Else: printCsData = " по материалам " & category & " №" & number
        End If
    Else
        If category = "" Then
            printCsData = "(данные по производству (делу) не указаны)"
        ElseIf number = "" Then
            printCsData = "(данные по производству (делу) не указаны)"
        Else: printCsData = " по материалам " & category & Chr(32) & index & " №" & number
        End If
    End If
End Function
'
'инкапсуляция
'Let
Private Property Let id(ByVal vData As Long)
'показания счетчика
    tmpid = vData
End Property
'
Friend Property Let number(ByVal vData As String)
Attribute number.VB_Description = "Номер дела"
'Номер Дела
    tmpnumber = vData
End Property
'
Friend Property Let index(ByVal vData As String)
'Индекс Дела
    tmpindex = vData
End Property
'
Friend Property Let category(ByVal vData As String)
'Категория Дела.
    tmpcategory = vData
End Property
'
Friend Property Let crimCondition(ByVal vData As String)
'Обстоятельства дела
     tmpcrimCondition = vData
End Property
'
Friend Property Let damage(ByVal vData As String)
'Категория травмы.
    tmpdamage = vData
End Property
'
Friend Property Let сomments(ByVal vData As String)
'Примечания.
    tmpсomments = vData
End Property
'
'Get
Friend Property Get id() As Long
'показания счетчика
    id = tmpid
'Debug.Print "количество счетчиков- " & id
End Property
'
Friend Property Get number() As String
'Номер Дела
    number = tmpnumber
'Debug.Print "Номер Дела= ", number
End Property
'
Friend Property Get index() As String
'Индекс Дела
    index = tmpindex
'Debug.Print "Индекс Дела- " & index
End Property
'
Friend Property Get category() As String
'Категория Дела.
    category = tmpcategory
'Debug.Print "Категория Дела= ", category
End Property
'
Friend Property Get crimCondition() As String
'Обстоятельства дела
     crimCondition = tmpcrimCondition
' Debug.Print "Категория травмы= " & crimCondition
End Property
'
Friend Property Get damage() As String
'Категория травмы.
    damage = tmpdamage
' Debug.Print "Обстоятельства дела= ", damage
End Property
'
Friend Property Get сomments() As String
Attribute сomments.VB_Description = "Примечания"
'Примечания.
    сomments = tmpсomments
'Debug.Print "Примечания= ", сomments
End Property

