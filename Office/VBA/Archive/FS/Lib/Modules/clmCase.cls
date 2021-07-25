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
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "Класс ""Дело"" - общий класс для данных, относящихся к производству уголовного или административного дела."
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'Класс "Дело"
'Цель: работа с программами VBA или базами данных Access
'Обработка данных при создании медико-криминалистических заключений
'Дата создания: 01.06.2016
'@version 0.0.1
'@author Andr.Nab.n@gmail.com
Option Explicit
Private mvarDtmCsPrvDate As Date        'Дата вынесения постановления
Private mvarstrCsCat As String          'Категория Дела: уголовное, административное
Private mvarstrCsIndex As String        'Индекс Дела
Private mvarstrCsNum As String          'Номер Дела
Private mvarstrCsCatDamage As String    'Категория травмы
Private mvarstrCsComments As String     'Примечания
Private mvarCsCounter As Integer        'Счетчик класса
'
Friend Function Print_CsData() As String
Attribute Print_CsData.VB_Description = "Печать данных ""основания назначения экспертизы"":"
'Функция печати данных по Делу
    If strCsIndex = "" Then
        If strCsCat = "" Then
            Print_CsData = "(данные по производству (делу) не указаны)"
        ElseIf Val(strCsNum) = 0 Then
            Print_CsData = "(данные по производству (делу) не указаны)"
        Else: Print_CsData = "по материалам " & strCsCat & " №" & strCsNum
        End If
    Else
        If strCsCat = "" Then
            Print_CsData = "(данные по производству (делу) не указаны)"
        ElseIf Val(strCsNum) = 0 Then
            Print_CsData = "(данные по производству (делу) не указаны)"
        Else: Print_CsData = "по материалам " & strCsCat & Chr(32) & strCsIndex & " №" & strCsNum
        End If
    End If
End Function
'
Friend Property Let strCsIndex(ByVal vData As String)
Attribute strCsIndex.VB_Description = "Индекс дела "
'Индекс Дела
    mvarstrCsIndex = vData
End Property
''
'Friend Property Set strCsIndex(ByVal vData As String)
''Индекс Дела
'    Set mvarstrCsIndex = vData
'End Property
'
Friend Property Get strCsIndex() As String
'Индекс Дела
'    If IsObject(mvarstrCsIndex) Then
'        Set strCsIndex = mvarstrCsIndex
'    Else
        strCsIndex = mvarstrCsIndex
'    End If
'Debug.Print "Индекс Дела=", strCsIndex
End Property
'
Friend Property Let strCsComments(ByVal vData As String)
Attribute strCsComments.VB_Description = "Примечания"
'Примечания.
    mvarstrCsComments = vData
End Property
'
Friend Property Get strCsComments() As String
'Примечания.
    strCsComments = mvarstrCsComments
'Debug.Print "Примечания= ", strCsComments
End Property
'
Friend Property Let strCsCatDamage(ByVal vData As String)
Attribute strCsCatDamage.VB_Description = "Категория травмы"
'Категория травмы.
    mvarstrCsCatDamage = vData
End Property
'
Friend Property Get strCsCatDamage() As String
'Категория травмы.
    strCsCatDamage = mvarstrCsCatDamage
' Debug.Print "Категория травмы= ", strCsCatDamage
End Property
'
Friend Property Let strCsNum(ByVal vData As String)
Attribute strCsNum.VB_Description = "Номер дела"
'Номер Дела
    mvarstrCsNum = vData
End Property
'
Friend Property Get strCsNum() As String
'Номер Дела
    strCsNum = mvarstrCsNum
'Debug.Print "Номер Дела= ", strCsNum
End Property
'
Friend Property Let strCsCat(ByVal vData As String)
Attribute strCsCat.VB_Description = "Категория дела"
'Категория Дела.
    mvarstrCsCat = vData
End Property
'
Friend Property Get strCsCat() As String
'Категория Дела.
    strCsCat = mvarstrCsCat
'Debug.Print "Категория Дела= ", strCsCat
End Property
'
Friend Property Let DtmCsPrvDate(ByVal vData As Date)
'Дата вынесения постановления.
    mvarDtmCsPrvDate = vData
End Property
'
Friend Property Get DtmCsPrvDate() As Date
'Дата вынесения постановления.
    DtmCsPrvDate = mvarDtmCsPrvDate
'Debug.Print "Дата вынесения постановления= ", DtmCsPrvDate
End Property
'
Friend Property Let CsCounter(ByVal vData As Integer)
'Счетчик класса.
    mvarCsCounter = vData
End Property
'
Friend Property Get CsCounter() As Integer
'Счетчик класса.
    CsCounter = mvarCsCounter
'Debug.Print "Счетчик класса Дело = ", CsCounter
End Property
'
Private Sub Class_Initialize()
CsCounter = CsCounter + 1
'Debug.Print "Счетчик класса Дело = ", CsCounter
End Sub
