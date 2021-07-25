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
Private mvarstrEvPlace As String            'Место изъятия ВД
Private mvarstrEntrants As String           'принадлежность кому-то
Private mvarstrEvPackage As String          'Упаковка
Private mvarstrEvName As String             'Название ВД
Private mvarstrEvCategory As String         'Категория ВД
Private mvarstrEvObjDescription As String   'Описание ВД
Private mvarstrEvSend As String             'Куда отправил ВД'Количесво ВД в упаковке
Private mvarblnEvProvide As Boolean         'ВД предоставили (да/нет)
Private mvarDtmEvOutTake As Date            'Дата изъятия ВД
Private mvarDtmEvFirstDate As Date          'Дата предоставления ВД
Private mvarDtmEvDispatchDate As Date       'Дата отправки ВД (на биологию)
Private mvarDtmEvReturnDate As Date         'Дата возврата ВД
Private mvarDtmEvFinishedDate As Date       'Дата отправки ВД следователю
Private mvarstrEvStamp As String            'Печать
Private mvarstrComments As String           'Примечания
'коллекция ВД
Private colEvidences As New Collection 'коллекция "ВД" (в упаковке)
'Private mvarlngEvCount As Long              'Количесво ВД в упаковке
'
Private Sub Class_Initialize()
'Constructor
    Set colEvidences = New Collection   'Количесво ВД в упаковке
    blnEvProvide = False 'ВД не предоставили
End Sub
'
Private Sub Class_Terminate()
'Destructor
   Set colEvidences = Nothing 'коллекция "ВД" (в упаковке)
End Sub
'
'====================== М Е Т О Д Ы ===============================
'Строковые представления данных:
Friend Function evidsToStirng(Optional nameEv As String = "") As String
'строковое представление=количества упаковок/ВД
If nameEv = "" Then
'    Dim tmp As Long
        Select Case colEvidences.Count
            Case Is = 0, 5 To 20
                nameEv = "объектов"
            Case Is = 1
                nameEv = "объект"
            Case 2 To 4
                nameEv = "объекта"
'            Case 5 To 20
'                nameEv = "объектов"
            Case Else
            Dim tmp As Long
            tmp = Right(colEvidences.Count, 1)
                Select Case tmp
                    Case Is = 0, 5 To 9
                        nameEv = "объектов"
                    Case Is = 1
                        nameEv = "объект"
                    Case 2 To 4
                        nameEv = "объекта"
'                    Case 5 To 9
'                        nameEv = "объектов"
                    End Select
            End Select
   End If
evidsToStirng = " (" & colEvidences.Count & Chr(32) & nameEv & ")"
'Debug.Print "строковое представление счетчика- " & evidsToStirng
End Function
'
Friend Function print_LineEvid() As String
'печть содержимого коллекции ВД В одну линию через запятую
    Dim i As Long
    Dim str As String
    For i = 1 To colEvidences.Count
        If i < colEvidences.Count Then
            str = str & EVID & ", "
        Else
            str = str & EVID & "."
        End If
    Next i
print_LineEvid = str
Debug.Print "print_LineEvid = " & printLineEvid
End Function
'
Friend Function print_ColumnEvid() As String
'печть содержимого коллекции ВД В одну линию через запятую
    Dim i As Long
    Dim str As String
    For i = 1 To colEvidences.Count
        If i < colEvidences.Count Then
            str = str & EVID & ";" & Chr(10)
        Else
            str = str & EVID & "."
        End If
    Next i
print_ColumnEvid = str
Debug.Print "print_ColumnEvid = " & printColumnEvid
End Function
'
Friend Function print_NumericColumnEvid() As String
'печть содержимого коллекции ВД В одну линию через запятую
    Dim i As Long
    Dim str As String
    For i = 1 To colEvidences.Count
        If i < colEvidences.Count Then
            str = Chr(9) & i & ". " & str & EVID & ";" & Chr(10)
        Else
            str = Chr(9) & i & ". " & str & EVID & "."
        End If
    Next i
print_NumericColumnEvid = str
Debug.Print "print_NumericColumnEvid = " & print_NumericColumnEvid
End Function
'
'==================== И Н К А П С У Л Я Ц И Я ==================
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
Friend Property Let blnEvProvide(ByVal vData As Boolean)
'ВД предоставили (да/нет)
    mvarblnEvProvide = vData
End Property
'
Friend Property Get blnEvProvide() As Boolean
'ВД предоставили (да/нет)
    blnEvProvide = mvarblnEvProvide
'Debug.Print "ВД предоставили (да/нет)= ", blnEvProvide
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
Private Property Let strEntrants(ByVal vData As String)
'Принадлежность ВД
   mvarstrEntrants = vData
End Property
'
Friend Property Get strEntrants() As String
'Принадлежность ВД
    strEntrants = mvarstrEntrants
Debug.Print "переменная принадлежность ВД = " & strEntrants
End Property


'Friend Property Set colEvidences(ByVal vData As Collection)
''Коллекция вещественых доказательств
''Syntax: Set x.colEvidences = Form1
'    Set mvarcolEvidences = vData
'End Property
''
'Friend Property Get colEvidences() As Collection
''Коллекция вещественых доказательств
'Set colEvidences = mvarcolEvidences
''Debug.Print "Коллекция ВД =  " & colEvidences
'End Property
'
'Private Property Let lngEvCount(ByVal vData As Long)
''Количесво ВД в упаковке
'    mvarlngEvCount = vData
''Debug.Print "lngEvCount =", lngEvCount
'End Property
''
'Friend Property Get lngEvCount() As Long
''Количесво ВД в упаковке
'    lngEvCount = mvarlngEvCount
''Debug.Print "Количесво ВД, изъятых у владельца или из МП= ", lngEvCount
'End Property


