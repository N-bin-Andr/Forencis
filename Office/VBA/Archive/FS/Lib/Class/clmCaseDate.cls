VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clmCaseDate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Класс "Дата"
'@author Andr.N@_bin
'@E-mail Andr.Nab.n@gmail.com
Option Explicit
'Системная дата:
Private tmpDateNow As Date      'ТЕКУЩАЯ ДАТА
Private arrNow(2) As Integer    'массив с элементами текущей даты (!!!год(0), месяц(1),день(2))
Private tmpnumDayOfYear As Integer 'количество дней в году
Private arrDt(3) As Date '(0=fst, 1=susp, 2=ren, 3=fin)
'Пользовательская дата:
Private tmpuserDate As Date         'Дата, введенная пользователем
Private userDtArr(2, 2) As Integer  'массив с элементами пользовательской даты(!!!год(0), месяц(1),день(2))
Private Const Msg As String = "Ошибка ввода: "
Private tmpexperience As Date ' = #1/6/2005#
Const PERIOD As Long = 30 '30 дней - период по умолчанию
Const CORRECTION As Long = 1 'коррекция (=1 т.к. день начала тоже учитывается)
'
Private Sub Class_Initialize()
'конструктор без парамертов
    DateNow = DateTime.Date   'ТЕКУЩАЯ ДАТА
    arrNow(0) = DatePart("yyyy", DateNow, vbMonday, vbFirstJan1) 'ТЕКУЩИЙ ГОД
    arrNow(1) = DatePart("m", DateNow, vbMonday, vbFirstJan1)    'ТЕКУЩИЙ МЕСЯЦ
    arrNow(2) = DatePart("D", DateNow, vbMonday, vbFirstJan1)    'ТЕКУЩИЙ ДЕНЬ.
experience = #1/6/2005#
'Отладка:
'Debug.Print "Отладка массива даты"
'   Dim i As Integer
'    For i = 0 To 2
'      Debug.Print "arrNow" & i & " = " & arrNow(i)
'    Next i
End Sub
'
Public Sub initDtArray(fst As Date, Optional susp As Date = 0, _
        Optional ren As Date = 0, Optional fin As Date = 0)
'массив с параметром/параметрами пользовательских дат
Dim arrDt(3) As Date
        arrDt(0) = fst
        arrDt(1) = susp
        arrDt(2) = ren
        arrDt(3) = fin
 Dim i As Integer
        For i = 0 To 3
            If arrDt(i) <> 0 Then
                userDate = arrDt(i)
                userDtArr(i, 0) = DatePart("yyyy", userDate, vbMonday, vbFirstJan1) 'ГОД
                userDtArr(i, 1) = DatePart("m", userDate, vbMonday, vbFirstJan1)    'МЕСЯЦ
                userDtArr(i, 2) = DatePart("D", userDate, vbMonday, vbFirstJan1)    'ДЕНЬ.
            End If
        Next i
'Отладка:
'Debug.Print "Отладка массива пользовательской даты"
'   Dim i As Integer, j As Integer
'    For i = 0 To 3
'        For j = 0 To 2
'            Debug.Print "дата " & i & " = " & userDtArr(i, j)
'        Next j
'    Next i
End Sub
'
'Методы:
Public Function ExamDate(tmpStr As String) As Date
'проверка даты, введенной пользователем на валидность
    '1)исключение пустой строки
        Do While Len(tmpStr) = 0
            tmpStr = InputBox("Введите дату: ", "Ввод даты", DateTime.Date)
        Loop
    '2) исключение неправильно введенной даты
        Do While IsDate(tmpStr) = False
            MsgBox Msg + "неправильно введена дата!", vbCritical
            tmpStr = InputBox("Правильно введите дату: ", "Ввод даты", DateTime.Date)
        Loop
    ExamDate = DateValue(tmpStr)
'Debug.Print "ExamDate- " & ExamDate
End Function
'
Public Function validateDate(dt As Date) As Date
'проверка даты (не больше ли введенная дата текущей даты)
'dt = ExamDate(str)
Dim tmp As Double, str As String
'1)сравнение даты с сегодняшним днем,
    tmp = compareDt(dt, DateNow)
    If tmp < 0 Then
        MsgBox Msg + "Введенная дата больше текущей даты!"
        str = InputBox("Правильно введите дату: ", "Ввод даты", DateTime.Date)
        dt = ExamDate(str)
    End If
validateDate = dt
'Debug.Print "validateDate- " & dt
End Function
'
Public Function dateToString(dt As Date) As String
'представление даты в виде отформатированного текста
    Dim str As String
        If dt = 0 Then
            str = "дата не определена"
        Else
            str = VBA.Format(dt, "dd/mm/yyyy")
        End If
    dateToString = str
'Debug.Print "dateToString- " & dateToString
End Function
'
Public Function if_DayOff(dt As Date) As String
'строка = дата выпадает на выходные дни!
 Dim str1 As String, str2 As String ',str3 As String
        str1 = "субботу!"
        str2 = "воскресенье!"
    If dt <> 0 Then
        If Val(Weekday(dt)) = 1 Then 'Если срок попадает на воскресение
            if_DayOff = "выпадает на " & str2
        ElseIf Val(Weekday(dt)) = 7 Then 'Если срок попадает на субботу
            if_DayOff = "выпадает на " & str1
        Else: if_DayOff = "выпадает на рабочие дни!" 'Если срок выпадает на рабочие дни
        End If
    End If
End Function
'
Public Function getPeriod(fst As Date, DayTale As Long, Optional susp As Date = 0, _
        Optional ren As Date = 0, Optional fin As Date = 0) As Date
'расчет срока (через введенный интервал дней - DayTale от даты начала
'т.е. например: срок +(DayTale=15) дней
'1)проверка  переменной fst на 0
    Dim str As String
        Do While fst = 0
            str = InputBox("Введите дату начала: ", "Ввод даты", DateTime.Date)
            fst = ExamDate(str)
        Loop
'2)сравнение переменной fst с текущеей датой:
    fst = validateDate(fst)
'3)проверка переменной DayTale(количество дней) на 0
        Do While DayTale = 0 Or IsNumeric(DayTale) = False
            DayTale = CLng(InputBox("Введите количество дней: ", "Ввод длительности периода", 30))
        Loop
'4) 'расчет периодов с учетом даты преостановления и возобновления экспертизы.
Dim interval1 As Date, interval2 As Double 'временные переменные для расчетов
    If susp = 0 And ren = 0 And fin = 0 Then
        interval1 = fst + DayTale
    ElseIf susp <> 0 And ren = 0 And fin = 0 Then
        susp = validateDate(susp)
        interval1 = susp + DayTale
    ElseIf susp <> 0 And ren <> 0 And fin = 0 Then
        susp = validateDate(susp)
        ren = validateDate(ren)
        If compareDt(susp, ren) > 0 Or compareDt(susp, ren) = 0 Then
            interval2 = DateDiff("D", fst, susp, vbMonday, vbFirstJan1)
            interval1 = DateAdd("D", (DayTale - interval2), ren)
        End If
    Else 'т.е. susp<>0, ren<>0 and fin<>0
        susp = validateDate(susp)
        ren = validateDate(ren)
    
        If compareDt(susp, ren) > 0 Or compareDt(susp, ren) = 0 Then
            interval2 = DateDiff("D", fst, susp, vbMonday, vbFirstJan1)
            interval1 = DateAdd("D", (DayTale - interval2), ren)
        End If
'       сравнение расчетной даты с датой fin
        fin = validateDate(fin)
        interval2 = compareDt(fin, interval1)
    End If
getPeriod = interval1
'Debug.Print "getPeriod-" & DayTale & " - " & getPeriod 'dateToString(fst)
End Function
'
Friend Function periodToString(dt As Date) As String
'строковое представление "срока"
'расчет разницы (кол-во дней) между конечной датой и "сегодня"
Dim tmp As Double
    tmp = DateDiff("Y", DateNow, dt)
'проверка полученного количества дней
Dim str1 As String, str2 As String, str3 As String
    Select Case tmp
        Case Is < 0
            str1 = "Срок просрочен! "
        Case Is = 0
            str1 = "Срок истекает сегодня! "
        Case Else
            str1 = "Срок истекает через " & tmp & dayToString(tmp)
        End Select
periodToString = str1
End Function
'
Function compareDt(fst As Date, fin As Date) As Double
'сравнение двух дат
'1)проверка на 0
Dim str1 As String, str2  As String
    If fst = 0 Then
        MsgBox Msg + "Начальная дата не определена !"
        Do While fst = 0
            str1 = InputBox("Введите начальную дату: ", "Ввод даты", DateTime.Date)
            fst = ExamDate(str1)
        Loop
    ElseIf fin = 0 Then
        MsgBox Msg + "Конечная дата не определена !"
        Do While fin = 0
            str2 = InputBox("Введите конечную дату: ", "Ввод даты", DateTime.Date)
            fin = ExamDate(str2)
        Loop
    End If
'2)сравнение дат
Dim tmp As Double
    tmp = DateDiff("Y", fst, fin, vbMonday, vbFirstJan1)
Dim tmpStr As String
    If tmp > 0 Then
        tmpStr = " больше "
    ElseIf tmp = 0 Then
        tmpStr = " равны "
    Else
        tmpStr = " меньше "
    End If
compareDt = tmp
End Function
'
Friend Function getExperience(expr As Date, lst As Date) As String
'Функция расчета стажа работы (строка =  2005 года (11 лет))
'expr - дата начала стажа работы
'lst - дата на момент расчета
'1) проверка на 0
'для стажа
Dim str1 As String, str2 As String
    If expr = 0 Then
        MsgBox Msg + "Начальная дата не определена !"
        Do While lst = 0
            str1 = InputBox("Введите дату начала стажа работы: ", "Ввод даты", DateTime.Date)
            expr = ExamDate(str1)
        Loop
    End If
'дата на момент расчета
    If lst = 0 Then
        MsgBox Msg + "Начальная дата не определена !"
        Do While lst = 0
            str2 = InputBox("Введите дату для расчета: ", "Ввод даты", DateTime.Date)
            lst = ExamDate(str2)
        Loop
    End If
    
'Сравнение введенных дат
'    compareDt
    
'2)стаж работы:
Dim lngEx As Double
    lngEx = DateDiff("yyyy", experience, lst) 'расчет стажа работы
    Debug.Print "lngEx = " & lngEx
    
    If lngEx = 0 Then
        getExperience = "менее одного года"
    Else: getExperience = "c " & Year(experience) & " года " & "(" & yearToString(lngEx) & ")"
    End If
Debug.Print "getExperience = " & getExperience
End Function
'
Public Function dayOfWeek(dt As Date) As String
'строка = день недели (понедельник/вторник и т.д.) по его номеру
    Dim DOW(8) As String 'Day Of Week
    DOW(0) = "Текущий день недели -> "
    DOW(1) = "понедельник"
    DOW(2) = "вторник"
    DOW(3) = "среда"
    DOW(4) = "четверг"
    DOW(5) = "пятница"
    DOW(6) = "суббота"
    DOW(7) = "воскресенье"
    Dim i As Integer
        i = VBA.DatePart("w", dt, vbMonday, vbFirstJan1)
    dayOfWeek = DOW(i)
End Function
'
Public Function monthOfYear(dt As Date) As String
'строка = месяц года (Январь.../Декабрь) по его номеру
    Dim MOY(13) As String 'Month Of Year
    MOY(0) = "Текущий месяц года - > "
    MOY(1) = "Январь"
    MOY(2) = "Февраль"
    MOY(3) = "Март"
    MOY(4) = "Апрель"
    MOY(5) = "Май"
    MOY(6) = "Июнь"
    MOY(7) = "Июль"
    MOY(8) = "Август"
    MOY(9) = "Сентябрь"
    MOY(10) = "Октябрь"
    MOY(11) = "Ноябрь"
    MOY(12) = "Декабрь"
    Dim i As Integer
        i = VBA.Month(dt)
    monthOfYear = MOY(i)
Debug.Print MOY(0) & MOY(i)
End Function
'
Public Function dayToString(lngD As Double) As String
'строка = день/дня/дней/
    If lngD > 0 Then
        Select Case lngD
            Case Is = 1
                dayToString = lngD & " день"
            Case 2 To 4
                dayToString = lngD & " дня"
            Case 5 To 20
                dayToString = lngD & " дней"
            Case Else
            Dim str As String
                str = Right(lngD, 1)
                    Select Case str
                        Case Is = 0
                            dayToString = lngD & " дней"
                        Case Is = 1
                           dayToString = lngD & " день"
                        Case 2 To 4
                            dayToString = lngD & " дня"
                        Case 5 To 9
                            dayToString = lngD & " дней"
                    End Select
            End Select
    End If
End Function
'
Public Function monthToString(lngM As Double) As String
'строка = месяц/месяца/месяцев
    If lngM > 0 Then
        Select Case lngM
            Case Is = 1
                monthToString = lngM & " месяц"
            Case 2 To 4
                monthToString = lngM & " месяца"
            Case 5 To 20
                monthToString = lngM & " месяцев"
            Case Else
            Dim str As String
                str = Right(lngM, 1)
                    Select Case str
                        Case Is = 0
                           monthToString = lngM & " месяцев"
                        Case Is = 1
                           monthToString = lngM & " месяц"
                        Case 2 To 4
                            monthToString = lngM & " месяца"
                        Case 5 To 9
                            monthToString = lngM & " месяцев"
                    End Select
            End Select
    End If
End Function
'
Public Function yearToString(lngY As Double) As String
'строка = менее одного года/год/года/лет/
    If lngY > 0 Then
        Select Case lngY
           Case Is = 1
                yearToString = lngY & " год"
            Case 2 To 4
                yearToString = lngY & " года"
            Case 5 To 20
                yearToString = lngY & " лет"
            Case Else
            Dim str As String
            str = Right(lngY, 1)
                Select Case str
                    Case Is = 0
                        yearToString = lngY & " лет"
                    Case Is = 1
                        yearToString = lngY & " год"
                    Case 2 To 4
                        yearToString = lngY & " года"
                    Case 5 To 9
                        yearToString = lngY & " лет"
                End Select
        End Select
    End If
End Function
'
Public Function isLeapYear(fst As Date) As Boolean
'определение: является ли год високосным
'!!!побочный эффект - расчет количества дней в году.
'1) проверка на 0
Dim str As String
    If fst = 0 Then
        MsgBox Msg + "Начальная дата не определена !"
        Do While fst = 0
            str = InputBox("Введите начальную дату: ", "Ввод даты", DateTime.Date)
            fst = ExamDate(str)
        Loop
    End If
Dim tmpYear As Long
    tmpYear = DatePart("yyyy", fst, vbMonday, vbFirstJan1)
    numDayOfYear = 366 + (tmpYear Mod 4 <> 0 Or (tmpYear Mod 400 = 0 And tmpYear Mod 100 <> 0))
    If numDayOfYear <> 365 Then
        isLeapYear = True
    Else: isLeapYear = False
    End If
'Debug.Print "количество дней в году- " & numDayOfYear
'Debug.Print tmpYear & " год является високосным? - " & isLeapYear
End Function
'
Friend Function getReminder(fst As Date, DayTale As Long, Optional txt As String = "") As String
'строка = "срок в 3/15/30 дней истекает ДАТА! Оформить документы и отправить объекты."
'1) Проверка кол-ва дней на 0 и определение значения по умолчанию
        Do While DayTale = 0 Or IsNumeric(DayTale) = False
            DayTale = CLng(InputBox("Введите количество дней: ", "Ввод длительности периода", 30))
        Loop
'2)расчет конечной даты напоминания
Dim dt As Date
    dt = getPeriod(fst, DayTale)
'3)проверка строки текста напоминания на 0
    If txt = "" Then
        txt = InputBox("Введите текст напоминания: ", "Напоминание", "оформить документы и отправить объекты!")
    End If
getReminder = "До " & dateToString(dt) & Chr(32) & StrConv(txt, vbLowerCase) & "!"
'MsgBox getReminder, vbInformation 'Debug.Print
End Function
'
'Инкапсуляция
'Let++++++++++++++++++++++++++++++++++++++++++++++++
Private Property Let DateNow(ByVal vData As Date)
'текущая дата
    tmpDateNow = vData
End Property
'
Private Property Let numDayOfYear(ByVal vData As Integer)
'количество дней в году
   tmpnumDayOfYear = vData
End Property
'
Friend Property Let userDate(ByVal vData As Date)
'дата, введенная пользователем
    tmpuserDate = vData
End Property
'
Private Property Let experience(ByVal vData As Date)
'дата начала отсчета стажа
    tmpexperience = vData
End Property

'Get+++++++++++++++++++++++++++++++++++++++++++++++++
Friend Property Get DateNow() As Date
'текущая дата
    DateNow = tmpDateNow
'Debug.Print "текущая дата- " & DateNow
End Property
'
Private Property Get numDayOfYear() As Integer
'количество дней в году
    numDayOfYear = tmpnumDayOfYear
'Debug.Print "количество дней в году- " & numDayOfYear
End Property
'
Friend Property Get userDate() As Date
'дата, введенная пользователем
    userDate = tmpuserDate
'Debug.Print "дата, введенная пользователем- " & userDate
End Property
'
Friend Property Get experience() As Date
'дата начала отсчета стажа
    experience = tmpexperience
'Debug.Print "дата начала отсчета стажа: " & userDate
End Property
'
'LIB----------------------------------------
'является ли год високосным (вариант 2)
' If tmpYear Mod 4 <> 0 Then
'            daysInYear = 365
'            isLeapYear = False
'    Else
'        If (tmpYear Mod 400 = 0 Or tmpYear Mod 100 <> 0) Then
'            daysInYear = 366
'            isLeapYear = True
'        Else
'            daysInYear = 365
'            isLeapYear = False
'        End If
'    End If

