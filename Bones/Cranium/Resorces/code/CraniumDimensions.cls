VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CraniumDimensions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@author Andr.Nab.n@gmail.com
'@16/05/2017
Option Explicit
'Объявление пользовательского типа данных
'Объявление двумерного массива строковых переменных (надпись "метод исследования")
Private Массив(1, 24) As String
Private dimension(24) As Currency
Const IMGdir As String = "D:\Crime\Рабочие документы\По работе\Шаблоны разных документов\MasterForm\Bons\Cranium\VB\"
'

Private constMan(3, 24) As Currency
'min
'constMan(0, 0 To 24) = (160, 127, 121, 90, 86, 112, 99, 92, 476, 123, 99, 94, _
'                                30, 25, 120, 82, 59, 100, 93, 78, 44, 38, 105, 85, 27)
'max
'constMan(1, 0 To 24) = (200, 158, 153, 114, 115, 138, 126, 120, 560, 151, 125, _
'                                132, 42, 40, 147, 115, 84, 139, 117, 106, 60, 51.5, 135,119, 43)
'стандартное отклонение
'constMan(2, 0 To 24) = (6.9, 5.4, 5.5, 4.2, 4.4, 5.1, 4.6, 5.1, 15.4, 4.8, 5.0, _
'                                6.5, 2.3, 2.11, 5.4, 5.2, 4.5, 7.0, 4.1, 5.1, 3.2, 2.0, 5.6, 6.6, 3.1)
'среднее квадратическое отклонение
'constMan(3, 0 To 24)  = (0.35, 0.28, 0.28, 0.21, 0.22, 0.26, 0.25, 0.27, 0.8, 0.24, _
                                    0.25, 0.36, 0.12, 0.11, 0.27, 0.27, 0.23, 0.36, 0.22, 0.26, _
                                    0.16, 0.1, 0.29, 0.34, 0.19)
'
Private constWom(3, 24) As Currency
'min
'constWom(0, 0 To 24) = (151, 122, 111, 82, 84, 104, 94, 86, 465, 111, 90, 93, 29, _
'                            23.5, 107, 78, 55, 96, 87, 74, 41, 36, 100, 77, 24)

'max
'constWom(1, 0 To 24) = (187, 152, 140, 109, 108, 133, 120, 116, 540, 145, 121, 124, _
                            41, 35, 139, 107, 78, 132, 113, 104, 56, 48, 127, 112, 41)
'стандартное отклонение
'constWom(2, 0 To 24) = (6.8, 5.5, 5.3, 4.6, 4.6, 5.7, 4.7, 5.2, 15.3, 5.7, 5.2, 5.9, _
                            2.4, 2.3, 6.0, 5.5, 4.1, 6.9, 4.6, 5.5, 3.0, 2.1, 5.9, 6.3, 3.1)
'среднее квадратическое отклонение
'constWom(3, 0 To 24)  = (0.4, 0.33, 0.32, 0.28, 0.28, 0.34, 0.28, 0.32, 0.95, 0.35, _
                            0.31, 0.37, 0.14, 0.15, 0.36, 0.33, 0.25, 0.44, 0.29, 0.32, _
                            0.18, 0.12, 0.35, 0.37, 0.20)






'Архив:
'=====================================================================================
'Private Static Function getCrForm() As Currency
'    getCrForm = dimension(1) * 100 / dimension(0)
''Debug.Print "показатель формы =" & getCrForm
'End Function
''
'Private Static Function selectCrForm(ByVal getCrForm As Currency) As String
''Классификация черепов по форме(поперечный диаметр*100/продольный диаметр)
'    'широкие (короткие) - брахикранные - >80%
'    'средние - мезокранные - от 75% до 79,9%;
'    'узкие (длинные) - долихокранные - <75%;
'    Dim curTmp As Currency
'        curTmp = getCrForm
'        Select Case curTmp  'сравнение переменной с константами
'        Case Is < 75
'            selectCrForm = "череп долихокранной формы " & "(" & curTmp & ")" 'значение, возвращаемое функцией
'        Case 75 To 79.9
'            selectCrForm = "Исследуемый череп мезокранной формы " & "(" & curTmp & ")"
'        Case Is > 80
'            selectCrForm = "Исследуемый череп брахикранной формы " & "(" & curTmp & ")"
'    End Select
'End Function
''
'Private Static Function getCrHeight() As Currency
'    getCrHeight = dimension(2) * 100 / dimension(0)
''Debug.Print "высотный показатель = " & getCrHeight
'End Function
''
'Private Static Function selectCrHeight(ByVal getCrHeight As Currency) As String
'    Dim curTmp As Currency
'        curTmp = getCrHeight
'    Select Case curTmp  'сравнение переменной с константами
'        Case Is < 70
'            selectCrHeight = "низкий (хамекранный) " & "(" & curTmp & ")"
'        Case 74, 9 To 70
'            selectCrHeight = "средневысокий (ортокранный) " & "(" & curTmp & ")"
'        Case Is > 75
'            selectCrHeight = "высокий (гипсокранный) " & "(" & curTmp & ")"
'    End Select
''Debug.Print "по высоте -" & selectCrHeight
'End Function
''
'Private Static Function getSex() As String
'    If ДМ.count > 12 Or ДМ.count > ДЖ.count Then
'        getSex = "исследуемый череп принадлежал индивидууму мужского пола."
'    ElseIf ДЖ.count > 12 Or ДЖ.count > ДМ.count Then
'        getSex = "исследуемый череп принадлежал индивидууму женского пола."
'    Else
'        If (ДМ.count + ВМ.count) > (ДЖ.count + ВЖ.count) Then
'            getSex = "исследуемый череп мог принадлежать индивидууму мужского пола."
'        ElseIf (ДЖ.count + ВЖ.count) > (ДМ.count + ВМ.count) Then
'            getSex = "исследуемый череп мог принадлежать индивидууму женского пола."
'        Else: getSex = "определить пол исследуемого черепа не представляется возможным."
'        End If
'    End If
'End Function
''1
'Private Static Function Select_Продольный_Диаметр(ByVal CurDimensions As Currency) As String
''Сравнение переменной "Продольный диаметр: глабелла-опистокраниум" с табличными данными
'    If CurDimensions <= 0 Then 'Исключение нулевого значения
'        Select_Продольный_Диаметр = "Пол невозможно определить"
'    Else
'        Select Case CurDimensions 'сравниваемая переменная, с уже присвоенным значением
'            Case Is <= 160
'                Select_Продольный_Диаметр = "Достоверно женский"
'            Case 160.1 To 172
'                Select_Продольный_Диаметр = "Вероятно женский"
'            Case 172.1 To 178.5
'                Select_Продольный_Диаметр = "Неопределенный"
'            Case 178.6 To 187
'                Select_Продольный_Диаметр = "Вероятно мужской"
'            Case Is >= 187.1
'                Select_Продольный_Диаметр = "Достоверно мужской"
'            Case Else: MsgBox "Ошибка! Введенный параметр не является числом " _
'                & "или выходит за границы допустимых значений!" & Err.Description
'        End Select
'    End If
'End Function
''2
'Private Static Function Select_Поперечный_Диаметр(ByVal CurDimensions As Currency) As String
''Сравнение переменной "Поперечный_Диаметр" с табличными данными
'    If CurDimensions <= 0 Then 'Исключение нулевого значения
'        Select_Поперечный_Диаметр = "Пол невозможно определить"
'    Else
'        Select Case CurDimensions 'сравниваемая переменная, с уже присвоенным значением
'            Case Is <= 127
'                Select_Поперечный_Диаметр = "Достоверно женский"
'            Case 127.1 To 138
'                Select_Поперечный_Диаметр = "Вероятно женский"
'            Case 138.1 To 143
'                Select_Поперечный_Диаметр = "Неопределенный"
'            Case 143.1 To 152
'                Select_Поперечный_Диаметр = "Вероятно мужской"
'            Case Is >= 152.1
'                Select_Поперечный_Диаметр = "Достоверно мужской"
'            Case Else: MsgBox "Ошибка! Введенный параметр не является числом " _
'                & "или выходит за границы допустимых значений!" & Err.Description
'        End Select
'    End If
'End Function
''3
'Private Static Function Select_Высотный_Диаметр(ByVal CurDimensions As Currency) As String
''Сравнение переменной "Высотный_Диаметр" с табличными данными
'    If CurDimensions <= 0 Then 'Исключение нулевого значения
'        Select_Высотный_Диаметр = "Пол невозможно определить"
'    Else
'        Select Case CurDimensions 'сравниваемая переменная, с уже присвоенным значением
'            Case Is <= 121
'                Select_Высотный_Диаметр = "Достоверно женский"
'            Case 121.1 To 128
'                Select_Высотный_Диаметр = "Вероятно женский"
'            Case 128.1 To 134
'                Select_Высотный_Диаметр = "Неопределенный"
'            Case 134.1 To 140.5
'                Select_Высотный_Диаметр = "Вероятно мужской"
'            Case Is >= 140.6
'                Select_Высотный_Диаметр = "Достоверно мужской"
'           Case Else: MsgBox "Ошибка! Введенный параметр не является числом " _
'                & "или выходит за границы допустимых значений!" & Err.Description
'        End Select
'    End If
'End Function
''4
'Private Static Function Select_Длина_Основания(ByVal CurDimensions As Currency) As String
''Сравнение переменной "Длина_Основания" с табличными данными
'    If CurDimensions <= 0 Then 'Исключение нулевого значения
'        Select_Длина_Основания = "Пол невозможно определить"
'    Else
'        Select Case CurDimensions 'сравниваемая переменная, с уже присвоенным значением
'            Case Is <= 90
'                Select_Длина_Основания = "Достоверно женский"
'            Case 90.1 To 96
'                Select_Длина_Основания = "Вероятно женский"
'            Case 96.1 To 101
'                Select_Длина_Основания = "Неопределенный"
'            Case 101.1 To 109
'                Select_Длина_Основания = "Вероятно мужской"
'            Case Is >= 109.1
'                Select_Длина_Основания = "Достоверно мужской"
'            Case Else: MsgBox "Ошибка! Введенный параметр не является числом " _
'                & "или выходит за границы допустимых значений!" & Err.Description
'        End Select
'    End If
'End Function
''5
'Private Static Function Select_Min_ширина_Лба(ByVal CurDimensions As Currency) As String
''Сравнение переменной "Min_ширина_Лба" с табличными данными
'    If CurDimensions <= 0 Then 'Исключение нулевого значения
'        Select_Min_ширина_Лба = "Пол невозможно определить"
'    Else
'        Select Case CurDimensions 'сравниваемая переменная, с уже присвоенным значением
'            Case Is <= 86
'                Select_Min_ширина_Лба = "Достоверно женский"
'            Case 86.1 To 95
'                Select_Min_ширина_Лба = "Вероятно женский"
'            Case 95.1 To 98
'                Select_Min_ширина_Лба = "Неопределенный"
'            Case 98.1 To 108
'                Select_Min_ширина_Лба = "Вероятно мужской"
'            Case Is >= 108.1
'                Select_Min_ширина_Лба = "Достоверно мужской"
'           Case Else: MsgBox "Ошибка! Введенный параметр не является числом " _
'                & "или выходит за границы допустимых значений!" & Err.Description
'        End Select
'    End If
'End Function
''6
'Private Static Function Select_Ширина_Основания(ByVal CurDimensions As Currency) As String
''Сравнение переменной "Ширина_Основания" с табличными данными
'    If CurDimensions <= 0 Then 'Исключение нулевого значения
'        Select_Ширина_Основания = "Пол невозможно определить"
'    Else
'        Select Case CurDimensions 'сравниваемая переменная, с уже присвоенным значением
'            Case Is <= 112
'                Select_Ширина_Основания = "Достоверно женский"
'            Case 112.1 To 117
'                Select_Ширина_Основания = "Вероятно женский"
'            Case 117.1 To 123
'                Select_Ширина_Основания = "Неопределенный"
'            Case 123.1 To 133
'                Select_Ширина_Основания = "Вероятно мужской"
'            Case Is >= 133.1
'                Select_Ширина_Основания = "Достоверно мужской"
'            Case Else: MsgBox "Ошибка! Введенный параметр не является числом " _
'                & "или выходит за границы допустимых значений!" & Err.Description
'        End Select
'    End If
'End Function
''7
'Private Static Function Select_Ширина_Затылка(ByVal CurDimensions As Currency) As String
''Сравнение переменной "Ширина_Затылка" с табличными данными
'    If CurDimensions <= 0 Then 'Исключение нулевого значения
'        Select_Ширина_Затылка = "Пол невозможно определить"
'    Else
'        Select Case CurDimensions 'сравниваемая переменная, с уже присвоенным значением
'            Case Is <= 99
'                Select_Ширина_Затылка = "Достоверно женский"
'            Case 99.1 To 106.9
'                Select_Ширина_Затылка = "Вероятно женский"
'            Case 107 To 110.4
'                Select_Ширина_Затылка = "Неопределенный"
'            Case 110.5 To 120
'                Select_Ширина_Затылка = "Вероятно мужской"
'            Case Is >= 120.1
'                Select_Ширина_Затылка = "Достоверно мужской"
'            Case Else: MsgBox "Ошибка! Введенный параметр не является числом " _
'                & "или выходит за границы допустимых значений!" & Err.Description
'        End Select
'    End If
'End Function
''8
'Private Static Function Select_Сосцевидная_Ширина(ByVal CurDimensions As Currency) As String
''Сравнение переменной "Сосцевидная_Ширина" с табличными данными
'    If CurDimensions <= 0 Then 'Исключение нулевого значения
'        Select_Сосцевидная_Ширина = "Пол невозможно определить"
'    Else
'        Select Case CurDimensions 'сравниваемая переменная, с уже присвоенным значением
'            Case Is <= 92
'                Select_Сосцевидная_Ширина = "Достоверно женский"
'            Case 92.1 To 100
'                Select_Сосцевидная_Ширина = "Вероятно женский"
'            Case 100.1 To 105
'                Select_Сосцевидная_Ширина = "Неопределенный"
'            Case 105.1 To 116
'                Select_Сосцевидная_Ширина = "Вероятно мужской"
'            Case Is >= 116.1
'                Select_Сосцевидная_Ширина = "Достоверно мужской"
'            Case Else: MsgBox "Ошибка! Введенный параметр не является числом " _
'                & "или выходит за границы допустимых значений!" & Err.Description
'        End Select
'    End If
'End Function
''9
'Private Static Function Select_Окружность_Черепа(ByVal CurDimensions As Currency) As String
''Сравнение переменной "Окружность_Черепа" с табличными данными
'    If CurDimensions <= 0 Then 'Исключение нулевого значения
'        Select_Окружность_Черепа = "Пол невозможно определить"
'    Else
'        Select Case CurDimensions 'сравниваемая переменная, с уже присвоенным значением
'            Case Is <= 476
'                Select_Окружность_Черепа = "Достоверно женский"
'            Case 476.1 To 500.5
'                Select_Окружность_Черепа = "Вероятно женский"
'            Case 500.6 To 516.5
'                Select_Окружность_Черепа = "Неопределенный"
'            Case 516.6 To 540
'                Select_Окружность_Черепа = "Вероятно мужской"
'            Case Is >= 540.1
'                Select_Окружность_Черепа = "Достоверно мужской"
'           Case Else: MsgBox "Ошибка! Введенный параметр не является числом " _
'                & "или выходит за границы допустимых значений!" & Err.Description
'        End Select
'    End If
'End Function
''10
'Private Static Function Select_Сагиттальная_Хорда(ByVal CurDimensions As Currency) As String
''Сравнение переменной "Сагиттальная_Хорда" с табличными данными
'    If CurDimensions <= 0 Then 'Исключение нулевого значения
'        Select_Сагиттальная_Хорда = "Пол невозможно определить"
'    Else
'        Select Case CurDimensions 'сравниваемая переменная, с уже присвоенным значением
'            Case Is <= 123
'                Select_Сагиттальная_Хорда = "Достоверно женский"
'            Case 123.1 To 128.5
'                Select_Сагиттальная_Хорда = "Вероятно женский"
'            Case 128.6 To 134.5
'                Select_Сагиттальная_Хорда = "Неопределенный"
'            Case 134.6 To 145
'                Select_Сагиттальная_Хорда = "Вероятно мужской"
'            Case Is >= 145.1
'                Select_Сагиттальная_Хорда = "Достоверно мужской"
'            Case Else: MsgBox "Ошибка! Введенный параметр не является числом " _
'                & "или выходит за границы допустимых значений!" & Err.Description
'        End Select
'    End If
'End Function
''11
'Private Static Function Select_Лобная_Хорда(ByVal CurDimensions As Currency) As String
''Сравнение переменной "Лобная_хорда" с табличными данными
'    If CurDimensions <= 0 Then 'Исключение нулевого значения
'        Select_Лобная_Хорда = "Пол невозможно определить"
'    Else
'        Select Case CurDimensions  'сравниваемая переменная, с уже присвоенным значением
'            Case Is <= 99
'                Select_Лобная_Хорда = "Достоверно женский"
'            Case 99.1 To 107
'                Select_Лобная_Хорда = "Вероятно женский"
'            Case 107.1 To 111.5
'                Select_Лобная_Хорда = "Неопределенный"
'            Case 111.6 To 121
'                Select_Лобная_Хорда = "Вероятно мужской"
'            Case Is >= 121.1
'                Select_Лобная_Хорда = "Достоверно мужской"
'            Case Else: MsgBox "Ошибка! Введенный параметр не является числом " _
'                & "или выходит за границы допустимых значений!" & Err.Description
'        End Select
'    End If
'End Function
''12
'Private Static Function Select_Теменная_Хорда(ByVal CurDimensions As Currency) As String
''Сравнение переменной " Теменная_Хорда" с табличными данными
'    If CurDimensions <= 0 Then 'Исключение нулевого значения
'        Select_Теменная_Хорда = "Пол невозможно определить"
'    Else
'        Select Case CurDimensions  'сравниваемая переменная, с уже присвоенным значением
'            Case Is <= 94
'                Select_Теменная_Хорда = "Достоверно женский"
'            Case 94.1 To 107
'                Select_Теменная_Хорда = "Вероятно женский"
'            Case 107.1 To 110.5
'                Select_Теменная_Хорда = "Неопределенный"
'            Case 110.6 To 124
'                Select_Теменная_Хорда = "Вероятно мужской"
'            Case Is >= 124.1
'                Select_Теменная_Хорда = "Достоверно мужской"
'            Case Else: MsgBox "Ошибка! Введенный параметр не является числом " _
'                & "или выходит за границы допустимых значений!" & Err.Description
'        End Select
'    End If
'End Function
''13
'Private Static Function Select_Длина_БЗО(ByVal CurDimensions As Currency) As String
''Сравнение переменной "Длина_БЗО" с табличными данными
'    If CurDimensions <= 0 Then 'Исключение нулевого значения
'        Select_Длина_БЗО = "Пол невозможно определить"
'    Else
'        Select Case CurDimensions 'сравниваемая переменная, с уже присвоенным значением
'            Case Is <= 30
'                Select_Длина_БЗО = "Достоверно женский"
'            Case 30.1 To 34
'                Select_Длина_БЗО = "Вероятно женский"
'            Case 34.1 To 36
'                Select_Длина_БЗО = "Неопределенный"
'            Case 36.1 To 41
'                Select_Длина_БЗО = "Вероятно мужской"
'            Case Is >= 41.1
'                Select_Длина_БЗО = "Достоверно мужской"
'        Case Else: MsgBox "Ошибка! Введенный параметр не является числом " _
'                & "или выходит за границы допустимых значений!" & Err.Description
'        End Select
'    End If
'End Function
''14
'Private Static Function Select_Ширина_БЗО(ByVal CurDimensions As Currency) As String
''Сравнение переменной "Ширина_БЗО" с табличными данными
'    If CurDimensions <= 0 Then 'Исключение нулевого значения
'        Select_Ширина_БЗО = "Пол невозможно определить"
'    Else
'        Select Case CurDimensions 'сравниваемая переменная, с уже присвоенным значением
'            Case Is <= 25
'                Select_Ширина_БЗО = "Достоверно женский"
'            Case 25.1 To 28.5
'                Select_Ширина_БЗО = "Вероятно женский"
'            Case 28.6 To 30.5
'                Select_Ширина_БЗО = "Неопределенный"
'            Case 30.6 To 35
'                Select_Ширина_БЗО = "Вероятно мужской"
'            Case Is >= 35.1
'                Select_Ширина_БЗО = "Достоверно мужской"
'            Case Else: MsgBox "Ошибка! Введенный параметр не является числом " _
'                & "или выходит за границы допустимых значений!" & Err.Description
'        End Select
'    End If
'End Function
''15
'Private Static Function Select_Скуловой_Диаметр(ByVal CurDimensions As Currency) As String
''Сравнение переменной "Скуловой_Диаметр" с табличными данными
'    If CurDimensions <= 0 Then 'Исключение нулевого значения
'        Select_Скуловой_Диаметр = "Пол невозможно определить"
'    Else
'        Select Case CurDimensions 'сравниваемая переменная, с уже присвоенным значением
'            Case Is <= 120
'                Select_Скуловой_Диаметр = "Достоверно женский"
'            Case 120.1 To 124
'                Select_Скуловой_Диаметр = "Вероятно женский"
'            Case 124.1 To 132
'                Select_Скуловой_Диаметр = "Неопределенный"
'            Case 132.1 To 139
'                Select_Скуловой_Диаметр = "Вероятно мужской"
'            Case Is >= 139.1
'                Select_Скуловой_Диаметр = "Достоверно мужской"
'           Case Else: MsgBox "Ошибка! Введенный параметр не является числом " _
'                & "или выходит за границы допустимых значений!" & Err.Description
'        End Select
'    End If
'End Function
''16
'Private Static Function Select_Длина_Основания_Лица(ByVal CurDimensions As Currency) As String
''Сравнение переменной "Длина_Основания_Лица" с табличными данными
'    If CurDimensions <= 0 Then 'Исключение нулевого значения
'        Select_Длина_Основания_Лица = "Пол невозможно определить"
'    Else
'        Select Case CurDimensions 'сравниваемая переменная, с уже присвоенным значением
'            Case Is <= 82
'                Select_Длина_Основания_Лица = "Достоверно женский"
'            Case 82.1 To 93
'                Select_Длина_Основания_Лица = "Вероятно женский"
'            Case 93.1 To 97.5
'                Select_Длина_Основания_Лица = "Неопределенный"
'            Case 97.6 To 107
'                Select_Длина_Основания_Лица = "Вероятно мужской"
'            Case Is >= 107.1
'                Select_Длина_Основания_Лица = "Достоверно мужской"
'            Case Else: MsgBox "Ошибка! Введенный параметр не является числом " _
'                & "или выходит за границы допустимых значений!" & Err.Description
'        End Select
'    End If
'End Function
''17
'Private Static Function Select_Верхняя_Высота_Лица(ByVal CurDimensions As Currency) As String
''Сравнение переменной "Верхняя_Высота_Лица" с табличными данными
'    If CurDimensions <= 0 Then 'Исключение нулевого значения
'        Select_Верхняя_Высота_Лица = "Пол невозможно определить"
'    Else
'        Select Case CurDimensions 'сравниваемая переменная, с уже присвоенным значением
'            Case Is <= 59
'                Select_Верхняя_Высота_Лица = "Достоверно женский"
'            Case 59.1 To 66.5
'                Select_Верхняя_Высота_Лица = "Вероятно женский"
'            Case 66.6 To 71
'                Select_Верхняя_Высота_Лица = "Неопределенный"
'            Case 71.1 To 78
'                Select_Верхняя_Высота_Лица = "Вероятно мужской"
'            Case Is >= 78.1
'                Select_Верхняя_Высота_Лица = "Достоверно мужской"
'            Case Else: MsgBox "Ошибка! Введенный параметр не является числом " _
'                & "или выходит за границы допустимых значений!" & Err.Description
'        End Select
'    End If
'End Function
''18
'Private Static Function Select_Полная_Высота_Лица(ByVal CurDimensions As Currency) As String
''Сравнение переменной "Полная_Высота_Лица" с табличными данными
'    If CurDimensions <= 0 Then 'Исключение нулевого значения
'        Select_Полная_Высота_Лица = "Пол невозможно определить"
'    Else
'        Select Case CurDimensions 'сравниваемая переменная, с уже присвоенным значением
'            Case Is <= 100
'                Select_Полная_Высота_Лица = "Достоверно женский"
'            Case 100.1 To 111
'                Select_Полная_Высота_Лица = "Вероятно женский"
'            Case 111.1 To 119
'                Select_Полная_Высота_Лица = "Неопределенный"
'            Case 119.1 To 132
'                Select_Полная_Высота_Лица = "Вероятно мужской"
'            Case Is >= 132.1
'                Select_Полная_Высота_Лица = "Достоверно мужской"
'            Case Else: MsgBox "Ошибка! Введенный параметр не является числом " _
'                & "или выходит за границы допустимых значений!" & Err.Description
'        End Select
'    End If
'End Function
''19
'Private Static Function Select_Верхняя_Ширина_Лица(ByVal CurDimensions As Currency) As String
''Сравнение переменной "Верхняя_Ширина_Лица" с табличными данными
'    If CurDimensions <= 0 Then 'Исключение нулевого значения
'        Select_Верхняя_Ширина_Лица = "Пол невозможно определить"
'    Else
'        Select Case CurDimensions 'сравниваемая переменная, с уже присвоенным значением
'            Case Is <= 93
'                Select_Верхняя_Ширина_Лица = "Достоверно женский"
'            Case 93.1 To 101
'                Select_Верхняя_Ширина_Лица = "Вероятно женский"
'            Case 101.1 To 105
'                Select_Верхняя_Ширина_Лица = "Неопределенный"
'            Case 105.1 To 113
'                Select_Верхняя_Ширина_Лица = "Вероятно мужской"
'            Case Is >= 113.1
'                Select_Верхняя_Ширина_Лица = "Достоверно мужской"
'            Case Else: MsgBox "Ошибка! Введенный параметр не является числом " _
'                & "или выходит за границы допустимых значений!" & Err.Description
'        End Select
'    End If
'End Function
''20
'Private Static Function Select_Средняя_Ширина_Лица(ByVal CurDimensions As Currency) As String
''Сравнение переменной "Средняя_Ширина_Лица"с табличными данными
'    If CurDimensions <= 0 Then 'Исключение нулевого значения
'        Select_Средняя_Ширина_Лица = "Пол невозможно определить"
'    Else
'        Select Case CurDimensions 'сравниваемая переменная, с уже присвоенным значением
'            Case Is <= 78
'                Select_Средняя_Ширина_Лица = "Достоверно женский"
'            Case 78.1 To 89
'                Select_Средняя_Ширина_Лица = "Вероятно женский"
'            Case 89.1 To 93.5
'                Select_Средняя_Ширина_Лица = "Неопределенный"
'            Case 93.6 To 104
'                Select_Средняя_Ширина_Лица = "Вероятно мужской"
'            Case Is >= 104.1
'                Select_Средняя_Ширина_Лица = "Достоверно мужской"
'           Case Else: MsgBox "Ошибка! Введенный параметр не является числом " _
'                & "или выходит за границы допустимых значений!" & Err.Description
'        End Select
'    End If
'End Function
''21
'Private Static Function Select_Высота_Носа(ByVal CurDimensions As Currency) As String
''Сравнение переменной "Высота_Носа"с табличными данными
'    If CurDimensions <= 0 Then 'Исключение нулевого значения
'        Select_Высота_Носа = "Пол невозможно определить"
'    Else
'        Select Case CurDimensions 'сравниваемая переменная, с уже присвоенным значением
'            Case Is <= 44
'                Select_Высота_Носа = "Достоверно женский"
'            Case 44.1 To 48.5
'                Select_Высота_Носа = "Вероятно женский"
'            Case 48.6 To 52
'                Select_Высота_Носа = "Неопределенный"
'            Case 52.1 To 56
'                Select_Высота_Носа = "Вероятно мужской"
'            Case Is >= 56.1
'                Select_Высота_Носа = "Достоверно мужской"
'            Case Else: MsgBox "Ошибка! Введенный параметр не является числом " _
'                & "или выходит за границы допустимых значений!" & Err.Description
'        End Select
'    End If
'End Function
''22
'Private Static Function Select_Ширина_Орбиты(ByVal CurDimensions As Currency) As String
''Сравнение переменной "Ширина_Орбиты" с табличными данными
'    If CurDimensions <= 0 Then 'Исключение нулевого значения
'        Select_Ширина_Орбиты = "Пол невозможно определить"
'    Else
'        Select Case CurDimensions 'сравниваемая переменная, с уже присвоенным значением
'            Case Is <= 38
'                Select_Ширина_Орбиты = "Достоверно женский"
'            Case 38.1 To 42
'                Select_Ширина_Орбиты = "Вероятно женский"
'            Case 42.1 To 43.5
'                Select_Ширина_Орбиты = "Неопределенный"
'            Case 43.6 To 48
'                Select_Ширина_Орбиты = "Вероятно мужской"
'            Case Is >= 48.1
'                Select_Ширина_Орбиты = "Достоверно мужской"
'            Case Else: MsgBox "Ошибка! Введенный параметр не является числом " _
'                & "или выходит за границы допустимых значений!" & Err.Description
'        End Select
'    End If
'End Function
''23
'Private Static Function Select_Мыщелковая_Ширина(ByVal CurDimensions As Currency) As String
''Сравнение переменной "Мыщелковая_Ширина"с табличными данными
'    If CurDimensions <= 0 Then 'Исключение нулевого значения
'        Select_Мыщелковая_Ширина = "Пол невозможно определить"
'    Else
'        Select Case CurDimensions 'сравниваемая переменная, с уже присвоенным значением
'            Case Is <= 105
'            Select_Мыщелковая_Ширина = "Достоверно женский"
'            Case 105.1 To 113.5
'                Select_Мыщелковая_Ширина = "Вероятно женский"
'            Case 113.6 To 118.5
'                Select_Мыщелковая_Ширина = "Неопределенный"
'            Case 118.6 To 127
'                Select_Мыщелковая_Ширина = "Вероятно мужской"
'            Case Is >= 127.1
'                Select_Мыщелковая_Ширина = "Достоверно мужской"
'           Case Else: MsgBox "Ошибка! Введенный параметр не является числом " _
'                & "или выходит за границы допустимых значений!" & Err.Description
'        End Select
'    End If
'End Function
''24
'Private Static Function Select_Бигониальная_Ширина(ByVal CurDimensions As Currency) As String
''Сравнение переменной "Бигониальная_Ширина"с табличными данными
'    If CurDimensions <= 0 Then 'Исключение нулевого значения
'        Select_Бигониальная_Ширина = "Пол невозможно определить"
'    Else
'        Select Case CurDimensions 'сравниваемая переменная, с уже присвоенным значением
'            Case Is <= 85
'                Select_Бигониальная_Ширина = "Достоверно женский"
'            Case 85.1 To 95
'                Select_Бигониальная_Ширина = "Вероятно женский"
'            Case 95.1 To 102.5
'                Select_Бигониальная_Ширина = "Неопределенный"
'            Case 102.6 To 112
'                Select_Бигониальная_Ширина = "Вероятно мужской"
'            Case Is >= 112.1
'                Select_Бигониальная_Ширина = "Достоверно мужской"
'            Case Else: MsgBox "Ошибка! Введенный параметр не является числом " _
'                & "или выходит за границы допустимых значений!" & Err.Description
'        End Select
'    End If
'End Function
''25
'Private Static Function Select_Высота_ТелаНЧ(ByVal CurDimensions As Currency) As String
''Сравнение переменной "Высота_ТелаНЧ" с табличными данными
'    If CurDimensions <= 0 Then 'Исключение нулевого значения
'        Select_Высота_ТелаНЧ = "Пол невозможно определить"
'    Else
'        Select Case CurDimensions 'сравниваемая переменная, с уже присвоенным значением
'            Case Is <= 27
'                Select_Высота_ТелаНЧ = "Достоверно женский"
'            Case 27.1 To 31
'                Select_Высота_ТелаНЧ = "Вероятно женский"
'            Case 31.1 To 33.5
'                Select_Высота_ТелаНЧ = "Неопределенный"
'            Case 33.6 To 41
'                Select_Высота_ТелаНЧ = "Вероятно мужской"
'            Case Is >= 41.1
'                Select_Высота_ТелаНЧ = "Достоверно мужской"
'            Case Else: MsgBox "Ошибка! Введенный параметр не является числом " _
'                & "или выходит за границы допустимых значений!" & Err.Description
'        End Select
'    End If
'End Function
''
'Private Static Sub Select_CrParameter()
'    Select Case Массив(3, N.count) 'сравниваемая переменная, с уже присвоенным значением
'        Case Is = "Достоверно женский"
'            ДЖ.increment
'        Case Is = "Вероятно женский"
'            ВЖ.increment
'        Case Is = "Неопределенный"
'            НПВ.increment
'        Case Is = "Вероятно мужской"
'            ВМ.increment
'        Case Is = "Достоверно мужской"
'            ДМ.increment
'        Case Else: MsgBox "Ошибка!" & Err.Description
'    End Select
'End Sub
'
'инициализация массива
'надписи:
'    Массив(0, 0) = "Продольный диаметр"
'        Массив(1, 0) = "glabella(g.) - opistokranion(op.)"
'    Массив(0, 1) = "Поперечный диаметр"
'        Массив(1, 1) = "euryon(eu.) - euryon(eu.)"
'    Массив(0, 2) = "Высотный диаметр"
'        Массив(1, 2) = "basion(ba.) - bregma(ba.)"
'    Массив(0, 3) = "Длина основания черепа"
'        Массив(1, 3) = "basion(ba.) - basion(ba.)"
'    Массив(0, 4) = "Наименьшая ширина лба"
'        Массив(1, 4) = "fronto-temporale(ft.) - fronto-temporale(ft.)"
'    Массив(0, 5) = "Ширина основания черепа"
'        Массив(1, 5) = "auriculare(au.) - auriculare(au.)"
'    Массив(0, 6) = "Ширина затылка"
'        Массив(1, 6) = "asterion(ast.) - asterion(ast.)"
'    Массив(0, 7) = "Сосцевидная ширина"
'        Массив(1, 7) = "mastoidale(ms.) - mastoidale(ms.)"
'    Массив(0, 8) = "Окружность черепа"
'        Массив(1, 8) = ""
'    Массив(0, 9) = "Сагиттальная хорда"
'        Массив(1, 9) = "nasion(n.) - opistion(o.)"
'    Массив(0, 10) = "Лобная хорда"
'        Массив(1, 10) = "nasion(n.) - bregma(b.)"
'    Массив(0, 11) = "Теменная хорда"
'        Массив(1, 11) = "(bregma(b.) - lambda(l.)"
'    Массив(0, 12) = "Длина БЗО"
'        Массив(1, 12) = "basion(ba.) - opistion(o.)"
'    Массив(0, 13) = "Ширина БЗО"
'        Массив(1, 13) = ""
'    Массив(0, 14) = "Скуловой диаметр"
'        Массив(1, 14) = "zygion(zg.) - zygion(zg.)"
'    Массив(0, 15) = "Длина основания лица"
'        Массив(1, 15) = "basion(ba.) - prostion(pr.)"
'    Массив(0, 16) = "Верхняя высота лица"
'        Массив(1, 16) = "nasion(n.) - alveolare(al.)"
'    Массив(0, 17) = "Полная высота лица"
'        Массив(1, 17) = "nasion(n.) - gnation(gn.)"
'    Массив(0, 18) = "Верхняя ширина лица"
'        Массив(1, 18) = "fronto-malare-temporale(fmt.) - fronto-malare-temporale(fmt.)"
'    Массив(0, 19) = "Средняя ширина лица"
'        Массив(1, 19) = "zygomaxillare(zm.) - zygomaxillare(zm.)"
'    Массив(0, 20) = "Высота носа"
'        Массив(1, 20) = "nasion(n.) - nasospinale(ns.)"
'    Массив(0, 21) = "Ширина левой орбиты"
'        Массив(1, 21) = "maxillofrantale(mf.) - ektokonchion(ek.)"
'    Массив(0, 22) = "Мыщелковая ширина"
'        Массив(1, 22) = ""
'    Массив(0, 23) = "Бигониальная ширина"
'        Массив(1, 23) = "gonion(go.) - gonijon(go.)"
'    Массив(0, 24) = "Высота тела НЧ"
'        Массив(1, 24) = "gnation(gn.) - infradentale(id.)"
''картинки:
'    Массив(2, 0) = IMGdir & "25.jpg"
'    Массив(2, 1) = IMGdir & "1.jpg"
'    Массив(2, 2) = IMGdir & "2.jpg"
'    Массив(2, 3) = IMGdir & "3.jpg"
'    Массив(2, 4) = IMGdir & "4.jpg"
'    Массив(2, 5) = IMGdir & "5.jpg"
'    Массив(2, 6) = IMGdir & "6.jpg"
'    Массив(2, 7) = IMGdir & "7.jpg"
'    Массив(2, 8) = IMGdir & "8.jpg"
'    Массив(2, 9) = IMGdir & "9.jpg"
'    Массив(2, 10) = IMGdir & "10.jpg"
'    Массив(2, 11) = IMGdir & "11.jpg"
'    Массив(2, 12) = IMGdir & "12.jpg"
'    Массив(2, 13) = IMGdir & "13.jpg"
'    Массив(2, 14) = IMGdir & "14.jpg"
'    Массив(2, 15) = IMGdir & "15.jpg"
'    Массив(2, 16) = IMGdir & "16.jpg"
'    Массив(2, 17) = IMGdir & "17.jpg"
'    Массив(2, 18) = IMGdir & "18.jpg"
'    Массив(2, 19) = IMGdir & "19.jpg"
'    Массив(2, 20) = IMGdir & "20.jpg"
'    Массив(2, 21) = IMGdir & "21.jpg"
'    Массив(2, 22) = IMGdir & "22.jpg"
'    Массив(2, 23) = IMGdir & "23.jpg"
'    Массив(2, 24) = IMGdir & "24.jpg"

