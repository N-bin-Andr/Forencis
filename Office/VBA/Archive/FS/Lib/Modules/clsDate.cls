VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "@auther Andr.Nab.n@gmail.com"
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'Класс "Date"
'Цель: работа с программами VBA или базами данных Access
'Обработка данных при создании медико-криминалистических заключений
'Дата создания: 01.06.2016
'@version 0.0.1
'@author Andr.Nab.n@gmail.com
Option Explicit

Friend Function compareTo(dtThat As Date) As Integer
Attribute compareTo.VB_Description = "сравнение даты с dtThat"
End Function
'
Friend Function equals(objThat As Object) As Boolean
Attribute equals.VB_Description = "совпадает ли дата с that?"
End Function
'
Friend Function toString() As String
Attribute toString.VB_Description = "строковое представление"
End Function
'
Friend Function year() As Integer
Attribute year.VB_Description = "Год"
End Function
'
Friend Function day() As Integer
Attribute day.VB_Description = "День"
End Function
'
Friend Function month() As Integer
Attribute month.VB_Description = "месяц"
End Function

