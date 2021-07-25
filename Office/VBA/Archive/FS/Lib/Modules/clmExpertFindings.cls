VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clmExpertFindings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "Класс-модуль ""ExpertFindings"" - описывает свойства и методы, используемые в ""Заключении Эксперта"""
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'Класс "Заключение эксперта"
'Цель: работа с программами VBA или базами данных Access
'Обработка данных при создании медико-криминалистических заключений
'Дата создания: 01.06.2016
'@version 0.0.1
'@author Andr.Nab.n@gmail.com
Option Explicit
Public intTmpDt As Integer              'переменная для расчета периодов(3, 15 и 30 дней)
Private mvarstrEFNum As String          'журнальный номер заключения эксперта
Private Const CODE = "№5.2/"            'код экспертизы
Private mvarstrEFDefinition As String   'Основание назначения экспертизы: постановление/определение
Private mvarDtmEFReceptDate As Date     'Дата получения постановления
Private mvarDtmEFFirstDay As Date       'Дата начала экспертизы
Private mvarDtmEFSuspDate As Date       'Дата приостановления производства экспертизы
Private mvarDtmEFRenDate As Date        'Дата возобновления производства экспертизы
Private mvarDtmEFDueDate As Date        'Скрок окончания производства экспертизы
Private mvarCurEFReady As Currency      'Процент готовности экспертизы
Private mvarstrEFImportance As String   'Категория "Важность" экспертизы
Private mvarstrEFCategories As String   'Категория экспертизы (первичная, дополнительная и т.д.)
Private mvarstrEFEvCategories As String 'Категория экспертизы по исследуемым объктам
Private mvarstrComments As String       'Примечания
Private mvarstrEFCondition As String    'Состояние экспертизы (не начата, в работе и т.д.)
Private mvarDOC As String               'Шаблон документа (экспертизы)
'
Friend Property Let strEFDefinition(ByVal vData As String)
Attribute strEFDefinition.VB_Description = "Основание производства экспертизы - постановление, определение или на платной основе"
'Основание назначения экспертизы (постановление/определение).
    mvarstrEFDefinition = vData
End Property
'
Friend Property Get strEFDefinition() As String
'Основание назначения экспертизы (постановление/определение).
    strEFDefinition = mvarstrEFDefinition
Debug.Print "Основание назначения экспертизы= ", strEFDefinition
End Property
'
Friend Property Let DOC(ByVal vData As String)
Attribute DOC.VB_Description = "Директория для сохранения создаваемых документов"
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.DOC = 5
    mvarDOC = vData
End Property
'
Friend Property Get DOC() As String
'Шаблон документа (экспертизы)
    DOC = mvarDOC
'Syntax: Debug.Print X.DOC
End Property
'
Friend Function Print_EFAutopsyData() As String
'Функция печати в документ данных по вскрытию из направления общего эксперта
    Print_EFAutopsyData = " Заключение эксперта (Акт) № " & frmNewResearch.txtAEFNum.Text & _
                        " Краткие обстоятельства и предполагаемая причина смерти: ..." & """."
End Function
'
Friend Function Create_EFReminder(ByVal DtmEFFirstDay As Date) As String
'Функция создания напоминания (дата начала + intTmpDt) дней
Dim DtmDate As Date 'Объявление временной переменной:
    DtmDate = Calculate_EFPeriod(ByVal intTmpDt)
    Create_EFReminder = "Оформить сопроводительные документы и отправить объекты до " & CStr(DtmDate)
End Function
'
'Friend Function Create_EFReminder3(ByVal DtmEFFirstDay As Date) As String
''Функция создания напоминания (дата начала + 3) дня
'Dim DtmDate As Date 'Объявление временной переменной:
'    'Перенос даты с выходных на пятницу:
'    If Val(Weekday(DtmEFFirstDay + 3)) = 1 Then 'Если срок "+3" попадает на воскресение
'        DtmDate = DtmEFFirstDay + 1
'    ElseIf Val(Weekday(DtmEFFirstDay + 3)) = 7 Then 'Если срок "+3" попадает на субботу
'        DtmDate = DtmEFFirstDay + 2
'    Else: DtmDate = DtmEFFirstDay + 3 'Если срок "+3" выпадает на рабочие дни
'    End If
'    Create_EFReminder3 = "Оформить сопроводительные документы и отправить ВД до " & CStr(DtmDate)
'End Function
''
'Friend Function Calculate_EFPeriod15() As Date
''Функция расчета срока (дата начала + 15) дней
'    'Перенос даты с выходных на пятницу:
'    If Val(Weekday(DtmEFFirstDay + 15)) = 1 Then 'Если срок "+15" попадает на воскресение
'        Calculate_EFPeriod15 = DtmEFFirstDay + 13
'    ElseIf Val(Weekday(DtmEFFirstDay + 15)) = 7 Then 'Если срок "+15" попадает на субботу
'        Calculate_EFPeriod15 = DtmEFFirstDay + 14
'    Else: Calculate_EFPeriod15 = DtmEFFirstDay + 15 'Если срок "+15" выпадает на рабочие дни
'    End If
'End Function
'
Friend Function Calculate_EFPeriod(ByVal intTmpDt As Integer) As Date
'Функция расчета срока окончания экспертизы (дата начала + 30 дней)
    'Перенос даты с выходных на пятницу:
    If Val(Weekday(DtmEFFirstDay + intTmpDt)) = 1 Then 'Если срок попадает на воскресение
        Calculate_EFPeriod = DtmEFFirstDay + (intTmpDt + 1)
    ElseIf Val(Weekday(DtmEFFirstDay + 30)) = 7 Then 'Если срок попадает на субботу
        Calculate_EFPeriod = DtmEFFirstDay + (intTmpDt + 2)
    Else: Calculate_EFPeriod = DtmEFFirstDay + intTmpDt 'Если срок выпадает на рабочие дни
    End If
End Function
'
Friend Function Criate_FullEFNum() As String
'Функция создания полного номера документа (экспертизы)
    Criate_FullEFNum = CODE & strEFNum
End Function
'
Friend Function Criate_ArchName(ByVal strEFNum As String) As String
'Функция создания архивного номера документа (экспертизы)
    Criate_ArchName = "012_" & Format(strEFNum, "#00000") & "_28_" & Right(CStr(Year(Now)), 2) & "_Набебин_Комар"
End Function
'
Friend Property Let strEFCondition(ByVal vData As String)
'Состояние экспертизы
    mvarstrEFCondition = vData
End Property
'
Friend Property Get strEFCondition() As String
'Состояние экспертизы.
    strEFCondition = mvarstrEFCondition
'Debug.Print "Состояние экспертизы = ", strEFCondition
End Property
'
Friend Property Let strComments(ByVal vData As String)
'Примечания
    mvarstrComments = vData
End Property
'
Friend Property Get strComments() As String
'Примечания.
    strComments = mvarstrComment
'Debug.Print "Примечания = "; strComments
End Property
'
Friend Property Let strEFEvCategories(ByVal vData As String)
'Категория экспертизы по исследуемым объктам.
    mvarstrEFEvCategories = vData
End Property
'
Friend Property Get strEFEvCategories() As String
'Категория экспертизы по исследуемым объктам.
    strEFEvCategories = mvarstrEFEvCategories
'Debug.Print "Категория экспертизы по исследуемым объктам = ", strEFEvCategories
End Property
'
Friend Property Let strEFCategories(ByVal vData As String)
'Категория экспертизы (первичная и т.д.).
    mvarstrEFCategories = vData
End Property
'
Friend Property Get strEFCategories() As String
'Категория экспертизы (первичная и т.д.).
    strEFCategories = mvarstrEFCategories
'Debug.Print "Категория экспертизы = ", strEFCategories
End Property
'
Friend Property Let strEFImportance(ByVal vData As String)
'Категория "Важность" экспертизы.
    mvarstrEFImportance = vData
End Property
'
Friend Property Get strEFImportance() As String
'Категория "Важность" экспертизы.
    strEFImportance = mvarstrEFImportance
'Debug.Print "Категория Важность экспертизы = ", strEFImportance
End Property
'
Friend Property Let CurEFReady(ByVal vData As Currency)
'Процент готовности экспертизы.
    mvarCurEFReady = vData
End Property
'
Friend Property Get CurEFReady() As Currency
'Процент готовности экспертизы.
    CurEFReady = mvarCurEFReady
'Debug.Print "Процент готовности экспертизы = ", CurEFReady
End Property
'
Friend Property Let DtmEFDueDate(ByVal vData As Date)
'Срок окончания экспертизы.
    mvarDtmEFDueDate = vData
End Property
'
Friend Property Get DtmEFDueDate() As Date
'Срок окончания экспертизы.
    DtmEFDueDate = mvarDtmEFDueDate
'Debug.Print "Срок окончания экспертизы = ", DtmEFDueDate
End Property
'
Friend Property Let DtmEFFinDate(ByVal vData As Date)
'Дата окончания экспертизы.
    mvarDtmEFFinDate = vData
End Property
'
Friend Property Get DtmEFFinDate() As Date
'Дата окончания экспертизы.
    DtmEFFinDate = mvarDtmEFFinDate
'Debug.Print "Дата окончания экспертизы = ", DtmEFFinDate
End Property
'
Friend Property Let DtmEFRenDate(ByVal vData As Date)
'Дата возобновления производства экспертизы.
    mvarDtmEFRenDate = vData
End Property
'
Friend Property Get DtmEFRenDate() As Date
'Дата возобновления производства экспертизы.
    DtmEFRenDate = mvarDtmEFRenDate
'Debug.Print "'Дата возобновления производства экспертизы = ", DtmEFRenDate
End Property
'
Friend Property Let DtmEFSuspDate(ByVal vData As Date)
'Дата приостановления производства экспертизы.
    mvarDtmEFSuspDate = vData
End Property
'
Friend Property Get DtmEFSuspDate() As Date
'Дата приостановления производства экспертизы.
    DtmEFSuspDate = mvarDtmEFSuspDate
'Debug.Print "Дата возобновления производства экспертизы = ", DtmEFSuspDate
End Property
'
Friend Property Let DtmEFFirstDay(ByVal vData As Date)
'Дата начала производства экспертизы.
    mvarDtmEFFirstDay = vData
End Property
'
Friend Property Get DtmEFFirstDay() As Date
'Дата начала производства экспертизы.
    DtmEFFirstDay = mvarDtmEFFirstDay
'Debug.Print "Дата начала производства экспертизы = ", DtmEFFirstDay
End Property
'
Friend Property Let DtmEFReceptDate(ByVal vData As Date)
'Дата получения постановления.
    mvarDtmEFReceptDate = vData
End Property
'
Friend Property Get DtmEFReceptDate() As Date
'Дата получения постановления.
    DtmEFReceptDate = mvarDtmEFReceptDate
'Debug.Print "Дата получения постановления = ", DtmEFReceptDate
End Property
'
Friend Property Let strEFNum(ByVal vData As String)
'Журнальный номер документа (экспертизы).
    mvarstrEFNum = vData
End Property
'
Friend Property Get strEFNum() As String
'Журнальный номер документа (экспертизы).
    strEFNum = mvarstrEFNum
'Debug.Print "Журнальный номер документа (экспертизы) = ", strEFNum
End Property

