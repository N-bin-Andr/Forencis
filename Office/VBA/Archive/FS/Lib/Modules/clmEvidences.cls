VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clmEvidences"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "Класс-модуль ""Вещественные доказательства"""
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'Класс "Вещественные доказательства"
'Работа с программами VBA или базами данных Access
'Обработка данных при создании медико-криминалистических заключений
'Дата создания: 01.06.2016
'@version 0.0.1
'@author Andr.Nab.n@gmail.com'
Option Explicit
'Свойства класса:
Private mvarstrEvName As String             'Название ВД
Private mvarintEvCount As Integer           'Количесво ВД, изъятых у владельца или из МП
Private mvarstrEvPlace As String            'Место изъятия ВД
Private mvarstrEvSend As String             'Куда отправил ВД
Private mvarBlnEvProvide As Boolean         'ВД предоставили (да/нет)
Private mvarDtmEvOutTake As Date            'Дата изъятия ВД
Private mvarDtmEvFirstDate As Date          'Дата предоставления ВД
Private mvarDtmEvDispatchDate As Date       'Дата отправки ВД (на биологию)
Private mvarDtmEvReturnDate As Date         'Дата возврата ВД
Private mvarDtmEvFinishedDate As Date       'Дата отправки ВД следователю
Private mvarstrEvCategory As String         'Категория ВД
Private mvarstrEvObjDescription As String   'Описание ВД
Private mvarstrComments As String           'Примечания
Private mvarstrEvPackage As String          'Упаковка
Private mvarstrEvStamp As String            'Печать
Private mvarcolEvidences As Collection      'коллекция "ВД" (для хранения значений массива)
'
Private Sub Class_Initialize()
'инициализация класса с объявлением нового экземпляра коллекции "ВД"
 Dim colEvidences As New Collection
' intEvGRCount = 0 'инициализация переменной  "Количесво групп ВД"
 intEvCount = 0
 BlnEvProvide = False
End Sub
'
Friend Property Set colEvidences(ByVal vData As Collection)
'Коллекция вещественых доказательств
'Syntax: Set x.colEvidences = Form1
    Set mvarcolEvidences = vData
End Property
'
Friend Property Get colEvidences() As Collection
'Коллекция вещественых доказательств
Set colEvidences = mvarcolEvidences
'Debug.Print "Коллекция ВД =  " & colEvidences
End Property
'
Friend Property Let strEvStamp(ByVal vData As String)
'Печать ВД
    mvarstrEvStamp = vData
End Property
'
Friend Property Get strEvStamp() As String
'Печать ВД
    strEvStamp = mvarstrEvStamp
'Debug.Print "'Печать ВД = ", strEvStamp
End Property
'
Friend Property Let strEvPackage(ByVal vData As String)
'упаковка
    mvarstrEvPackage = vData
End Property
'
Friend Property Get strEvPackage() As String
'упаковка
    strEvPackage = mvarstrEvPackage
'Debug.Print "упаковка= ", strEvPackage
End Property
'_____________________________________________________________________________________
'!!! Перенести в модуль mdPrint
Friend Static Function Print_OurStamp() As String
Dim strTx As String
    strTx = "мастичным оттиском синего цвета круглой печати " & Chr(171) & _
        "Государственный комитет судебных экспертиз Республики Беларусь. Главное управление судебно-медицинских экспертиз. Для пакетов" _
        & Chr(187)
Dim EvNumArr3(1, 1 To 2) 'создание массива
    EvNumArr3(1, 1) = "вещественное доказательство упаковано, опечатано " & strTx
    EvNumArr3(1, 2) = "вещественные доказательства упакованы, опечатаны " & strTx
'OurStamp = EvNumArr3(1, intEvGRCount)
End Function
'
'!!! Перенести в модуль mdPrint
Friend Static Function Print_EvNumArr2() As String
Dim strTx1 As String, strTx2 As String, strTx3 As String, strTx4 As String
    strTx1 = "Целостность упаковки не нарушена, извлечение"
    strTx2 = "без повреждения целостности упаковки невозможно. При вскрытии упаковки из нее"
    strTx3 = "для медико-криминалистического исследования,"
    strTx4 = "перечную, указанному в направлении и в сопроводительной надписи к вещественным доказательствам. "
Dim EvNumArr2(1, 1 To 2) 'создание массива
    EvNumArr2(1, 1) = strTx1 & " предоставленного объекта " & strTx2 & " был извлечен объект: " _
    & Chr(10) & _
    "Объект, представленный " & strTx3 & " соответствует " & strTx4
    EvNumArr2(1, 2) = strTx1 & " предоставленных объектов  " & strTx2 & " были извлечены следующие объекты: " & Chr(10) & _
    "Объекты, представленные " & strTx3 & " соответствуют " & strTx4
'Print_EvNumArr2 = EvNumArr2(1, intEvGRCount)
End Function
''
Friend Static Function Print_EvNumArr1() As String
'Функция печати описания упаковок и т.д. вещественных доказательств
Dim strTx1 As String, strTx2 As String
strTx2 = "мастичным оттиском синего цвета круглой печати"
    If StrComp(strEvPackage, "картонную коробку") = True Then
        strTx1 = "опечатанную"
    Else: strTx1 = "опечатанный"
    Debug.Print "strTx1 = ", strTx1
    End If
Dim EvNumArr1(1, 1 To 2) 'создание массива
    EvNumArr1(1, 1) = "Вещественное доказательство доставлено нарочным, упакованное в " & strEvPackage & _
           Chr(32) & strTx1 & Chr(32) & strTx2
    EvNumArr1(1, 2) = "Вещественные доказательства доставлены нарочным, упакованные в " & strEvPackage & _
           Chr(32) & strTx1 & Chr(32) & strTx2
'Print_EvNumArr1 = EvNumArr1(1, intEvGRCount)
End Function
'
Friend Property Let intEvCount(ByVal vData As Integer)
'Количесво ВД, изъятых у владельца или из МП
    mvarintEvCount = vData
'Debug.Print "intEvCount =", intEvCount
End Property
'
Friend Property Get intEvCount() As Integer
'Количесво ВД, изъятых у владельца или из МП
    intEvCount = mvarintEvCount
'Debug.Print "Количесво ВД, изъятых у владельца или из МП= ", intEvCount
End Property
'
Friend Property Let strComments(ByVal vData As String)
'Примечания
    mvarstrComments = vData
End Property
'
Friend Property Get strComments() As String
'Примечания
    strComments = mvarstrComments
'Debug.Print "Примечания= ", strComments
End Property
'
Friend Property Let strEvObjDescription(ByVal vData As String)
'Описание ВД
    mvarstrEvObjDescription = vData
End Property
'
Friend Property Get strEvObjDescription() As String
'Описание ВД
    strEvObjDescription = mvarstrEvObjDescription
'Debug.Print "Описание ВД= ", strEvObjDescription
End Property
'
Friend Property Let strEvCategory(ByVal vData As String)
'Категория ВД
    mvarstrEvCategory = vData
End Property
'
Friend Property Get strEvCategory() As String
'Категория ВД
    strEvCategory = mvarstrEvCategory
'Debug.Print "Категория ВД= ", strEvCategory
End Property
'
Friend Property Let DtmEvFinishedDate(ByVal vData As Date)
'Дата отправки ВД следователю
    mvarDtmEvFinishedDate = vData
End Property
'
Friend Property Get DtmEvFinishedDate() As Date
'Дата отправки ВД следователю
    DtmEvFinishedDate = mvarDtmEvFinishedDate
'Debug.Print "Дата отправки ВД следователю= ", DtmEvFinishedDate
End Property
'
Friend Property Let DtmEvReturnDate(ByVal vData As Date)
'Дата возврата ВД
    mvarDtmEvReturnDate = vData
End Property
'
Friend Property Get DtmEvReturnDate() As Date
'Дата возврата ВД
    DtmEvReturnDate = mvarDtmEvReturnDate
'Debug.Print "Дата возврата ВД= ", DtmEvReturnDate
End Property
'
Friend Property Let DtmEvDispatchDate(ByVal vData As Date)
'Дата отправки ВД
    mvarDtmEvDispatchDate = vData
End Property
'
Friend Property Get DtmEvDispatchDate() As Date
'Дата отправки ВД
    DtmEvDispatchDate = mvarDtmEvDispatchDate
'Debug.Print "Дата отправки ВД= ", DtmEvDispatchDate
End Property
'
Friend Property Let DtmEvFirstDate(ByVal vData As Date)
'Дата предоставления ВД
    mvarDtmEvFirstDate = vData
End Property
'
Friend Property Get DtmEvFirstDate() As Date
'Дата предоставления ВД
    DtmEvFirstDate = mvarDtmEvFirstDate
'Debug.Print "Дата предоставления ВД= ", DtmEvFirstDate
End Property
'
Friend Property Let DtmEvOutTake(ByVal vData As Date)
'Дата изъятия ВД
    mvarDtmEvOutTake = vData
End Property
'
Friend Property Get DtmEvOutTake() As Date
'Дата изъятия ВД
    DtmEvOutTake = mvarDtmEvOutTake
'Debug.Print "Дата изъятия ВД= ", DtmEvOutTake
End Property
'
Friend Property Let BlnEvProvide(ByVal vData As Boolean)
'ВД предоставили (да/нет)
    mvarBlnEvProvide = vData
End Property
'
Friend Property Get BlnEvProvide() As Boolean
'ВД предоставили (да/нет)
    BlnEvProvide = mvarBlnEvProvide
'Debug.Print "ВД предоставили (да/нет)= ", BlnEvProvide
End Property
'
Friend Property Let strEvSend(ByVal vData As String)
'Куда отправил ВД
    mvarstrEvSend = vData
End Property
'
Friend Property Get strEvSend() As String
'Куда отправил ВД
    strEvSend = mvarstrEvSend
'Debug.Print "Куда отправил ВД= ", strEvSend
End Property
'
Friend Property Let strEvPlace(ByVal vData As String)
'Место изъятия ВД
    mvarstrEvPlace = vData
End Property
'
Friend Property Get strEvPlace() As String
'Место изъятия ВД
    strEvPlace = mvarstrEvPlace
'Debug.Print "Место изъятия ВД= ", strEvPlace
End Property
'

Friend Property Let strEvName(ByVal vData As String)
'Название ВД
    mvarstrEvName = vData
End Property
'
Friend Property Get strEvName() As String
'Название ВД
    strEvName = mvarstrEvName
'Debug.Print "Название ВД= ", strEvName
End Property
'
