VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clmEvBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Класс "Коробка с вещественными доказательствами"
'Работа с программами VBA или базами данных Access
'Обработка данных при создании медико-криминалистических заключений
'Дата создания: 01.06.2016
'@author Andr.Nab.n@gmail.com'

Option Explicit
'Свойства класса:
Private mvarstrBxPlace As String            'Место изъятия ВД
Private mvarstrBxEntrants As String         'принадлежность кому-то
Private mvarstrBxPackage As String          'Упаковка
Private mvarstrBxName As String             'Название коробки
Private mvarstrBxCategory As String         'Категория
Private mvarstrBxObjDescription As String   'Описание коробки
Private mvarstrBxSend As String             'Куда отправил коробку
Private mvarblnBxProvide As Boolean         'коробку предоставили (да/нет)
Private mvarDtmBxOutTake As Date            'Дата изъятия ВД (коробки)
Private mvarDtmBxFirstDate As Date          'Дата предоставления ВД (коробки)
Private mvarDtmBxDispatchDate As Date       'Дата отправки ВД (коробки на биологию)
Private mvarDtmBxReturnDate As Date         'Дата возврата ВД (коробки)
Private mvarDtmBxFinishedDate As Date       'Дата отправки ВД (коробки) следователю
Private mvarstrBxStamp As String            'Печать на коробке
Private mvarstrBxComments As String         'Примечания
'коллекция ВД
Public colEvidences As New Collection 'коллекция "ВД" (в упаковке)
Const PROVIDE As String = "Для производства судебной медико-криминалистической экспертизы предоставили: "
'
Private Sub Class_Initialize()
'Constructor
    Set colEvidences = New Collection   'Количесво ВД в упаковке
    blnBxProvide = False 'ВД не предоставили
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
                    End Select
            End Select
   End If
evidsToStirng = colEvidences.Count & Chr(32) & nameEv & Chr(32) ' " (" & colEvidences.Count & Chr(32) & nameEv & ")"
'Debug.Print "строковое представление счетчика- " & evidsToStirng
End Function
'
Friend Function print_LineEvid() As String
''печть содержимого коллекции ВД в одну линию через запятую
    Dim i As Long
    Dim str As String
    With colEvidences
        If .Count > 0 Then
            If .Count = 1 Then
                str = .Item(1) & ". "
            Else
                For i = 1 To .Count
                    If i = .Count Then
                        str = str & .Item(i) & ". "
                    Else
                        str = str & .Item(i) & ", "
                    End If
                Next i
            End If
        End If
    End With
print_LineEvid = str
Debug.Print "print_LineEvid = " & print_LineEvid
End Function
'
Friend Function print_LineEvidBxEntrants() As String
''печть содержимого коллекции ВД в одну линию через запятую
    Dim i As Long
    Dim str As String
    With colEvidences
        If .Count > 0 Then
            If .Count = 1 Then
                str = .Item(1) & " (" & strBxEntrants & "); "
            Else
                For i = 1 To .Count
                    If i = .Count Then
                        str = str & .Item(i)
                    Else
                        str = str & .Item(i) & ", "
                    End If
                Next i
                str = str & " (" & strBxEntrants & ")"
            End If
        End If
    End With
print_LineEvidBxEntrants = str
'Debug.Print "print_LineEvidBxEntrants -> " & print_LineEvidBxEntrants
End Function
'
Friend Function print_ColumnEvid() As String
'печть содержимого коллекции ВД в столбик
Dim i As Long
Dim str As String
    With colEvidences
        If .Count > 0 Then
            If .Count > 1 Then
                For i = 1 To .Count
                    If i < .Count Then
                        str = str & .Item(i) & ";" & Chr(10)
                    Else
                        str = str & .Item(i) & ". " & Chr(10)
                    End If
                Next i
            Else
               str = .Item(1) & ". " & Chr(10)
            End If
        End If
    End With
print_ColumnEvid = str
'Debug.Print "print_ColumnEvid = " & print_ColumnEvid
End Function
'
Friend Function print_NumericColumnEvid() As String 'print_ColumnEvid() As String
''печть содержимого коллекции ВД в нумерованный столбик
Dim i As Long
Dim str As String
    With colEvidences
        If .Count > 0 Then
            If .Count > 1 Then
                For i = 1 To .Count
                    If i < .Count Then
                        str = str & i & "). " & .Item(i) & ";" & Chr(10)
                    Else
                        str = str & i & "). " & .Item(i) & ". " & Chr(10)
                    End If
                Next i
            Else
               str = "1). " & .Item(1) & ". " & Chr(10)
            End If
        End If
    End With
Debug.Print "print_NumericColumnEvid -> " & str
print_NumericColumnEvid = str
End Function
'
Friend Function print_NumericColumnBoxes() As String
'печать коробок в нумерованные столбцы: 1. Box / 1.1. Ev1 / 1.2. Ev2 ...
Dim i As Long
Dim str1 As String, str2 As String
    With colEvidences
        If .Count > 0 Then
            If .Count > 1 Then
                For i = 1 To .Count
                    If i < .Count Then
                        str1 = strBxEntrants & ":" & Chr(10)
'                        str1 = strBxEntrants & " (" & strBxPackage & "):" & Chr(10)
                        str2 = str2 & vbTab & i & "). " & .Item(i) & ";" & Chr(10)
                    Else
                        str2 = str2 & vbTab & i & "). " & .Item(i) & ". " & Chr(10)
                    End If
                Next i
            Else
                str1 = strBxEntrants & ":" & Chr(10)
'                str1 = strBxEntrants & " (" & strBxPackage & "):" & Chr(10)
                str2 = vbTab & "1). " & .Item(1) & ". " & Chr(10)
            End If '.Count > 1
        End If '.Count > 0
    End With
'Debug.Print "print_NumericColumnBoxes -> " & str1 & str2
print_NumericColumnBoxes = str1 & str2
End Function
'
Friend Function print_columnEvidEndFoto() As String
'печть содержимого коллекции ВД в столбик
Dim i As Long
Dim str As String
    With colEvidences
        If .Count > 0 Then
            If .Count > 1 Then
                For i = 1 To .Count
                    str = str & StrConv(.Item(i), vbProperCase) & " (Фото №№ ). " & Chr(10)
                Next i
            Else
               str = StrConv(.Item(1), vbProperCase) & " (Фото №№ ). " & Chr(10)
            End If
        End If
    End With
print_columnEvidEndFoto = str
'Debug.Print "print_columnEvidEndFoto -> " & print_columnEvidEndFoto
End Function
'
Public Function print_BxProvides() As String
'строка = Для производства СМЭ предоставили: ВД1....ВДn
    If blnBxProvide = True Then
        print_BxProvides = PROVIDE & print_LineEvid
    Else: print_BxProvides = "NB!!! ВД НЕ ПРЕДОСТАВЛЕНЫ" & Chr(10) & PROVIDE & print_LineEvid
    End If
End Function
'
Public Function print_DeliveryEv() As String
''Функция печати доставки упаковки/ок
''строка = "Вещественное/ые доказательство/а доставлено/ы нарочным, упакованное/ые в + упаковка + печать"
Dim arrStr(10) As String
    arrStr(0) = print_LineEvid
    arrStr(2) = "Целостность упаковки не нарушена, извлечение "
    arrStr(7) = print_columnEvidEndFoto
    
    If strBxPackage = "Картонная коробка" Then
        arrStr(8) = ", доставлена "
    Else
         arrStr(8) = ", доставлен "
    End If

    If DtmBxFirstDate <> #12:00:00 AM# Then
        arrStr(9) = arrStr(8) & DtmBxFirstDate
    Else: arrStr(9) = arrStr(8) & "(укажите дату)"
    End If
    
With colEvidences
    If strBxStamp <> "" Then
            arrStr(1) = ", опечатанная мастичным оттиском синего цвета круглой печати: " & strBxStamp & arrStr(9) & " (" & "Фото №№" & "). "
    Else
        arrStr(1) = arrStr(9) & " (" & "Фото №№" & "). "
    End If
    If .Count > 0 Then
        If .Count = 1 Then
            
            arrStr(3) = "предоставленного объекта "
            arrStr(4) = "был извлечен "
            arrStr(5) = "Объект, предоставленный на исследование, "
            arrStr(6) = "соответствует "
        Else
            arrStr(3) = "предоставленных объектов "
            arrStr(4) = "были извлечены: "
            arrStr(5) = "Объекты, предоставленные на исследование, "
            arrStr(6) = "соответствуют "
        End If
    End If
End With '& evidsToStirng & strBxEntrants &
print_DeliveryEv = strBxEntrants & arrStr(1) & arrStr(2) & arrStr(3) & _
    "без повреждения целостности упаковки невозможно. При вскрытии упаковки из нее " & arrStr(4) _
    & arrStr(0) & arrStr(5) & arrStr(6) & _
    "перечню, указанному в направлении и в сопроводительной надписи к вещественным доказательствам." & _
    Chr(10) & arrStr(7)
End Function
'
Friend Function print_BoxFotoList() As String
'строка = Фото №№ & ВД
Dim i As Long
Dim str1 As String, str2 As String
Dim arrStr(4) As String
    arrStr(0) = "Фото №№. "
'    arrStr(1) = "Упаковка с "
    With colEvidences
        If .Count > 0 Then
            If .Count = 1 Then
'                arrStr(2) = "объектом (" & strBxEntrants & "), общий вид. " & Chr(10)
                arrStr(2) = strBxEntrants & ", общий вид. " & Chr(10)
                arrStr(4) = arrStr(0) & StrConv(.Item(1), vbProperCase) & Chr(10)
            Else:
'                arrStr(2) = "объектами (" & strBxEntrants & "), общий вид. " & Chr(10)
                 arrStr(2) = strBxEntrants & ", общий вид. " & Chr(10)
                For i = 1 To .Count
                     arrStr(4) = arrStr(4) & arrStr(0) & StrConv(.Item(i), vbProperCase) & ", общий вид." & Chr(10)
                Next i
            End If
        End If
         arrStr(3) = arrStr(0) & arrStr(1) & arrStr(2) & Chr(10)
    End With
print_BoxFotoList = arrStr(3) & arrStr(4)
End Function
'
Friend Function print_Labels(fullNum As String, _
                            csData As String, _
                            firstDay As String, _
                            strExpert As String) As String
'печать этикеток на коробки с ВД
Dim i As Long
Dim str1 As String
Dim arrStr(6) As String
    arrStr(0) = "К судебной медико-криминалистической экспертизе " & fullNum & Chr(32)
    arrStr(1) = csData 'printCsData
    arrStr(2) = print_NumericColumnBoxes
    arrStr(3) = "Дата регистрации материалов экспертизы: " & firstDay & Chr(10)
    arrStr(4) = "Дата упаковки: " & Chr(10)
    arrStr(5) = "Эксперт(ы) " & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & strExpert & Chr(10)
     For i = 0 To UBound(arrStr)
        str1 = str1 & arrStr(i)
    Next i
print_Labels = str1
End Function
'
'==================== И Н К А П С У Л Я Ц И Я ==================
'
Public Property Let strBxStamp(ByVal vData As String)
'Печать ВД
    mvarstrBxStamp = vData
End Property
'
Friend Property Get strBxStamp() As String
'Печать ВД
    strBxStamp = mvarstrBxStamp
'Debug.Print "'Печать ВД = ", strBxStamp
End Property
'
Public Property Let strBxPackage(ByVal vData As String)
'упаковка
    mvarstrBxPackage = vData
End Property
'
Friend Property Get strBxPackage() As String
'упаковка
    strBxPackage = mvarstrBxPackage
'Debug.Print "упаковка= ", strBxPackage
End Property
'
Public Property Let strComments(ByVal vData As String)
'Примечания
    mvarstrBxComments = vData
End Property
'
Friend Property Get strBxComments() As String
'Примечания
    strBxComments = mvarstrBxComments
'Debug.Print "Примечания= ", strComments
End Property
'
Public Property Let strBxObjDescription(ByVal vData As String)
'Описание ВД
    mvarstrBxObjDescription = vData
End Property
'
Friend Property Get strBxObjDescription() As String
'Описание ВД
    strBxObjDescription = mvarstrBxObjDescription
'Debug.Print "Описание ВД= ", strBxObjDescription
End Property
'
Public Property Let strBxCategory(ByVal vData As String)
'Категория ВД
    mvarstrBxCategory = vData
End Property
'
Friend Property Get strBxCategory() As String
'Категория ВД
    strBxCategory = mvarstrBxCategory
'Debug.Print "Категория ВД= ", strBxCategory
End Property
'
Public Property Let DtmBxFinishedDate(ByVal vData As Date)
'Дата отправки ВД следователю
    mvarDtmBxFinishedDate = vData
End Property
'
Friend Property Get DtmBxFinishedDate() As Date
'Дата отправки ВД следователю
    DtmBxFinishedDate = mvarDtmBxFinishedDate
'Debug.Print "Дата отправки ВД следователю= ", DtmBxFinishedDate
End Property
'
Public Property Let DtmBxReturnDate(ByVal vData As Date)
'Дата возврата ВД
    mvarDtmBxReturnDate = vData
End Property
'
Friend Property Get DtmBxReturnDate() As Date
'Дата возврата ВД
    DtmBxReturnDate = mvarDtmBxReturnDate
'Debug.Print "Дата возврата ВД= ", DtmBxReturnDate
End Property
'
Public Property Let DtmBxDispatchDate(ByVal vData As Date)
'Дата отправки ВД
    mvarDtmBxDispatchDate = vData
End Property
'
Friend Property Get DtmBxDispatchDate() As Date
'Дата отправки ВД
    DtmBxDispatchDate = mvarDtmBxDispatchDate
'Debug.Print "Дата отправки ВД= ", DtmBxDispatchDate
End Property
'
Public Property Let DtmBxFirstDate(ByVal vData As Date)
'Дата предоставления ВД
    mvarDtmBxFirstDate = vData
End Property
'
Friend Property Get DtmBxFirstDate() As Date
'Дата предоставления ВД
    DtmBxFirstDate = mvarDtmBxFirstDate
'Debug.Print "Дата предоставления ВД= ", DtmBxFirstDate
End Property
'
Public Property Let DtmBxOutTake(ByVal vData As Date)
'Дата изъятия ВД
    mvarDtmBxOutTake = vData
End Property
'
Friend Property Get DtmBxOutTake() As Date
'Дата изъятия ВД
    DtmBxOutTake = mvarDtmBxOutTake
'Debug.Print "Дата изъятия ВД= ", DtmBxOutTake
End Property
'
Public Property Let blnBxProvide(ByVal vData As Boolean)
'ВД предоставили (да/нет)
    mvarblnBxProvide = vData
End Property
'
Friend Property Get blnBxProvide() As Boolean
'ВД предоставили (да/нет)
    blnBxProvide = mvarblnBxProvide
'Debug.Print "ВД предоставили (да/нет)= ", blnBxProvide
End Property
'
Public Property Let strBxSend(ByVal vData As String)
'Куда отправил ВД
    mvarstrBxSend = vData
End Property
'
Friend Property Get strBxSend() As String
'Куда отправил ВД
    strBxSend = mvarstrBxSend
'Debug.Print "Куда отправил ВД= ", strBxSend
End Property
'
Public Property Let strBxPlace(ByVal vData As String)
'Место изъятия ВД
    mvarstrBxPlace = vData
End Property
'
Friend Property Get strBxPlace() As String
'Место изъятия ВД
    strBxPlace = mvarstrBxPlace
'Debug.Print "Место изъятия ВД= ", strBxPlace
End Property
'

Friend Property Let strBxName(ByVal vData As String)
'Название ВД
    mvarstrBxName = vData
End Property
'
Public Property Get strBxName() As String
'Название ВД
    strBxName = mvarstrBxName
'Debug.Print "Название ВД= ", strBxName
End Property
'
Public Property Let strBxEntrants(ByVal vData As String)
'Принадлежность ВД
   mvarstrBxEntrants = vData
End Property
'
Friend Property Get strBxEntrants() As String
'Принадлежность ВД
    strBxEntrants = mvarstrBxEntrants
'Debug.Print "переменная принадлежность ВД = " & strBxEntrants
End Property

'================== L I B =======================
'Public Function print_DeliveryEv() As String
'''Функция печати доставки упаковки/ок
'''строка = "Вещественное/ые доказательство/а доставлено/ы нарочным, упакованное/ые в + упаковка + печать"
'Dim arrStr(8) As String
'    arrStr(0) = print_LineEvid
'    arrStr(7) = print_columnEvidEndFoto
'    If strBxPackage = "картонную коробку" Then 'If StrComp(strBxPackage, "картонную коробку") = True Then
'            arrStr(2) = "опечатанную "
'    Else: arrStr(2) = "опечатанный "
'    End If
'With colEvidences
'    If .Count > 0 Then
'        If .Count = 1 Then
'            arrStr(1) = "Вещественное доказательство доставлено нарочным, упакованное в "
'            arrStr(3) = "предоставленного объекта "
'            arrStr(4) = "был извлечен "
'            arrStr(5) = "Объект, предоставленный на исследование, "
'            arrStr(6) = "соответствует "
'        Else
'            arrStr(1) = "Вещественные доказательства доставлены нарочным, упакованные в "
'            arrStr(3) = "предоставленных объектов "
'            arrStr(4) = "были извлечены: "
'            arrStr(5) = "Объекты, предоставленные на исследование, "
'            arrStr(6) = "соответствуют "
'        End If
'    End If
'End With '& evidsToStirng & strBxEntrants &
'print_DeliveryEv = arrStr(1) & arrStr(2) & strBxPackage & " (" & "Фото №№" & "). " & _
'    "Целостность упаковки не нарушена, извлечение " & arrStr(3) & _
'    "без повреждения целостности упаковки невозможно. При вскрытии упаковки из нее " & arrStr(4) _
'    & print_LineEvidBxEntrants & ". " & arrStr(5) & arrStr(6) & _
'    "перечню, указанному в направлении и в сопроводительной надписи к вещественным доказательствам." & _
'    Chr(10) & arrStr(7)
'End Function


'Friend Function print_LineEvid() As String
''печть содержимого коллекции ВД в одну линию через запятую
'    Dim i As Long
'    Dim str As String
'    For i = 1 To colEvidences.Count
'        If i < colEvidences.Count Then
'            str = str & EVID & ", "
'        Else
'            If strBxEntrants <> "" Then
'                str = str & EVID & "(" & strBxEntrants & ")" & ";"
'            Else
'                str = str & EVID & ";"
'            End If
'        End If
'    Next i
'print_LineEvid = str
'Debug.Print "print_LineEvid = " & printLineEvid
'End Function
''
'Friend Function print_ColumnEvid() As String
''печть содержимого коллекции ВД в столбик
'    Dim i As Long
'    Dim str As String
'    For i = 1 To colEvidences.Count
'        If i < colEvidences.Count Then
'            str = str & EVID & ";" & Chr(10)
'        Else
'            str = str & EVID & "."
'        End If
'    Next i
'print_ColumnEvid = str
'Debug.Print "print_ColumnEvid = " & printColumnEvid
'End Function
''
'Friend Function print_NumericColumnEvid() As String
''печть содержимого коллекции ВД в нумерованный столбик
'    Dim i As Long
'    Dim str As String
'    For i = 1 To colEvidences.Count
'        If i < colEvidences.Count Then
'            str = Chr(9) & i & ". " & str & EVID & ";" & Chr(10)
'        Else
'            str = Chr(9) & i & ". " & str & EVID & "."
'        End If
'    Next i
'print_NumericColumnEvid = str
'Debug.Print "print_NumericColumnEvid = " & print_NumericColumnEvid
'End Function
''
'Friend Function print_boxLineEvid() As String
''строка = ВД в одну строку через запятую + (принадлежность упаковки )
'
'End Function
'
'




'Friend Property Set colBxidences(ByVal vData As Collection)
''Коллекция вещественых доказательств
''Syntax: Set x.colBxidences = Form1
'    Set mvarcolBxidences = vData
'End Property
''
'Friend Property Get colBxidences() As Collection
''Коллекция вещественых доказательств
'Set colBxidences = mvarcolBxidences
''Debug.Print "Коллекция ВД =  " & colBxidences
'End Property
'
'Private Property Let lngBxCount(ByVal vData As Long)
''Количесво ВД в упаковке
'    mvarlngBxCount = vData
''Debug.Print "lngBxCount =", lngBxCount
'End Property
''
'Friend Property Get lngBxCount() As Long
''Количесво ВД в упаковке
'    lngBxCount = mvarlngBxCount
''Debug.Print "Количесво ВД, изъятых у владельца или из МП= ", lngBxCount
'End Property



