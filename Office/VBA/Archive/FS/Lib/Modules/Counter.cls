VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Counter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "счетчик"
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'Класс "Counter"
'Цель: работа с программами VBA или базами данных Access
'Обработка данных при создании медико-криминалистических заключений
'Дата создания: 01.06.2016
'@version 0.0.1
'@author Andr.Nab.n@gmail.com
Option Explicit
Private mvarcount As Integer 'local copy
Private mvarstrName As String 'local copy
'
Friend Property Let strName(ByVal vData As String)
Attribute strName.VB_Description = "имя счетчика"
    mvarstrName = vData
End Property
'
Friend Property Get strName() As String
    strName = mvarstrName
'Debug.Print X.strName
End Property
'
'Friend Property Let count(ByVal vData As Integer)
'    mvarcount = vData
'End Property
''
'Friend Property Get count() As Integer
'    count = mvarcount
''Debug.Print X.count
'End Property
'
Friend Function toStirng() As String
Attribute toStirng.VB_Description = "строковое представление счетчика"
toString = CStr(count & Chr(32) & strName)
End Function
'
Friend Sub increment()
Attribute increment.VB_Description = "увеличение счетчика на 1"
count = count + 1
End Sub

