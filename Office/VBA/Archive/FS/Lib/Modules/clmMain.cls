VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
'Переменные класса
Public mvarlDocumentCount As Long 'переменная-счетчик количества новых документов
'
Public Property Let lDocumentCount(ByVal vData As Long)
Attribute lDocumentCount.VB_Description = "Переменная-счетчик количества создаваемых документов"
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.lDocumentCount = 5
    mvarlDocumentCount = vData
End Property
'
Public Property Get lDocumentCount() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.lDocumentCount
    lDocumentCount = mvarlDocumentCount
End Property



