VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmCraniumDimensions 
   Caption         =   "Краниометрический метод исследования черепа"
   ClientHeight    =   9165
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6630
   OleObjectBlob   =   "frmCraniumDimensions.frx":0000
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "frmCraniumDimensions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


'@author Andr.Nab.n@gmail.com
'@01/01/2015
'
Option Explicit
'
Public ДЖ As Counter
Public ВЖ As Counter
Public Неопр As Counter
Public ВМ As Counter
Public ДМ As Counter
Public НПВ As Counter
Public myForm As Object
Public N As Counter
'Объявление двумерного массива строковых переменных (надпись "метод исследования")
Private Массив(3, 24) As String
Private dimension(24) As Currency
'Объявление переменной директории, хранящей файлы
Const IMGdir As String = "D:\Crime\MasterForm\Bons\Cranium\VB\"
'
Private CurDimensions As Currency
'
Private Static Sub Select_Dimensions(ByVal CurDimensions As Currency)
'Сравнение переменной "Продольный диаметр: глабелла-опистокраниум" с табличными данными
    Dim tmp As Integer
        tmp = N.Count
    Select Case tmp 'сравниваемая переменная, с уже присвоенным значением
        Case Is = 0
            dimension(tmp) = CurDimensions
            lblSex1.Caption = Select_Продольный_Диаметр(ByVal CurDimensions)
        Case Is = 1
            dimension(tmp) = CurDimensions
            lblSex1.Caption = Select_Поперечный_Диаметр(ByVal CurDimensions)
        Case Is = 2
            dimension(tmp) = CurDimensions
            lblSex1.Caption = Select_Высотный_Диаметр(ByVal CurDimensions)
        Case Is = 3
            dimension(tmp) = CurDimensions
            lblSex1.Caption = Select_Длина_Основания(ByVal CurDimensions)
        Case Is = 4
            dimension(tmp) = CurDimensions
            lblSex1.Caption = Select_Min_ширина_Лба(ByVal CurDimensions)
        Case Is = 5
            dimension(tmp) = CurDimensions
            lblSex1.Caption = Select_Ширина_Основания(ByVal CurDimensions)
        Case Is = 6
            dimension(tmp) = CurDimensions
            lblSex1.Caption = Select_Ширина_Затылка(ByVal CurDimensions)
        Case Is = 7
            dimension(tmp) = CurDimensions
            lblSex1.Caption = Select_Сосцевидная_Ширина(ByVal CurDimensions)
        Case Is = 8
            dimension(tmp) = CurDimensions
            lblSex1.Caption = Select_Окружность_Черепа(ByVal CurDimensions)
        Case Is = 9
            dimension(tmp) = CurDimensions
            lblSex1.Caption = Select_Сагиттальная_Хорда(ByVal CurDimensions)
        Case Is = 10
            dimension(tmp) = CurDimensions
            lblSex1.Caption = Select_Лобная_Хорда(ByVal CurDimensions)
        Case Is = 11
            dimension(tmp) = CurDimensions
            lblSex1.Caption = Select_Теменная_Хорда(ByVal CurDimensions)
        Case Is = 12
            dimension(tmp) = CurDimensions
            lblSex1.Caption = Select_Длина_БЗО(ByVal CurDimensions)
        Case Is = 13
            dimension(tmp) = CurDimensions
            lblSex1.Caption = Select_Ширина_БЗО(ByVal CurDimensions)
        Case Is = 14
            dimension(tmp) = CurDimensions
            lblSex1.Caption = Select_Скуловой_Диаметр(ByVal CurDimensions)
        Case Is = 15
            dimension(tmp) = CurDimensions
            lblSex1.Caption = Select_Длина_Основания_Лица(ByVal CurDimensions)
        Case Is = 16
            dimension(tmp) = CurDimensions
            lblSex1.Caption = Select_Верхняя_Высота_Лица(ByVal CurDimensions)
        Case Is = 17
            dimension(tmp) = CurDimensions
            lblSex1.Caption = Select_Полная_Высота_Лица(ByVal CurDimensions)
        Case Is = 18
            dimension(tmp) = CurDimensions
            lblSex1.Caption = Select_Верхняя_Ширина_Лица(ByVal CurDimensions)
        Case Is = 19
            dimension(tmp) = CurDimensions
            lblSex1.Caption = Select_Средняя_Ширина_Лица(ByVal CurDimensions)
        Case Is = 20
            dimension(tmp) = CurDimensions
            lblSex1.Caption = Select_Высота_Носа(ByVal CurDimensions)
        Case Is = 21
            dimension(tmp) = CurDimensions
            lblSex1.Caption = Select_Ширина_Орбиты(ByVal CurDimensions)
        Case Is = 22
            dimension(tmp) = CurDimensions
            lblSex1.Caption = Select_Мыщелковая_Ширина(ByVal CurDimensions)
        Case Is = 23
            dimension(tmp) = CurDimensions
            lblSex1.Caption = Select_Бигониальная_Ширина(ByVal CurDimensions)
        Case Is = 24
            dimension(tmp) = CurDimensions
            lblSex1.Caption = Select_Высота_ТелаНЧ(ByVal CurDimensions)
        End Select
    Call lblVisible
End Sub
'
Private Sub cmdEraseData_Click()
    Call CrDimensions.Main
    Me.Hide
End Sub
'
Private Sub lblVisible()
    lblSex.Visible = True 'скрытие поля "определяемый пол"
    lblSex1.Visible = True 'скрытие поля "определяемый пол"
    cmdOK.Visible = True 'скрытие кнопки "ОК"
End Sub
'
Private Sub lblUnVisible()
    lblSex.Visible = False 'скрытие поля "определяемый пол"
    lblSex1.Visible = False 'скрытие поля "определяемый пол"
    cmdOK.Visible = False 'скрытие кнопки "ОК"
End Sub
'
Private Sub cmdOK_Click()
'Действия программы при нажатии кнопки "ОК"
    Call lblUnVisible
'Загрузка надписи "Метод исследования" и соответствующей картинки
    If N.Count < 24 Then
        Массив(3, N.Count) = lblSex1.Caption
        dimension(N.Count) = txtDimensions.Text
        Call Select_CrParameter
        N.increment
        lblMetod.Caption = Массив(0, N.Count)
        lblMetod1.Caption = Массив(1, N.Count)
        imgCraniumInv.Picture = LoadPicture(Массив(2, N.Count))  ' загрузка изображения
        txtDimensions.Text = "" '"измерение"
    Else
        Массив(3, N.Count) = lblSex1.Caption
        dimension(N.Count) = txtDimensions.Text
        Call Select_CrParameter
        With Me
'размер формы при окончании исследования
        .Height = 408
        .Width = 343.5
        .lblMetod.Height = 32 'размер надписи при окончнии исследования
        .lblMetod.Caption = "Результаты исследования"
'оформление надписи
        .lblMetod.BackColor = &H80FFFF
        .lblMetod.ForeColor = &HFF&
        .lblMetod.Font.name = "Arial"
        .lblMetod.Font.Size = 18
'скрытие ненужных элементов формы
        .imgCraniumInv.Visible = False
        Call lblUnVisible
        .lbltxtDimensions.Visible = False
        .txtDimensions.Visible = False
'Появление нужных элементов
        .lblCrForms.Visible = True
            .lblCrForms.Caption = selectCrForm(ByVal getCrForm)
        .lblCrHeight.Visible = True
            .lblCrHeight.Caption = selectCrHeight(ByVal getCrHeight)
        .lblParameter.Visible = True
        .lblДЖ.Visible = True
            .lblДЖ.Caption = ДЖ.toString
        .lblВЖ.Visible = True
            .lblВЖ.Caption = ВЖ.toString
        .lblНеопр.Visible = True
            .lblНеопр.Caption = Неопр.toString
        .lblВМ.Visible = True
            .lblВМ.Caption = ВМ.toString
        .lblДМ.Visible = True
            .lblДМ.Caption = ДМ.toString
        .lblНПВ.Visible = True
            .lblНПВ.Caption = НПВ.toString
        .cmdToWord.Visible = True
        .cmdEraseData.Visible = True
        .lblMetod1.Caption = getSex
        End With
    End If
End Sub
'
Private Sub cmdToWord_Click()
'Печать таблицы в активный документ Word
    With ActiveDocument
'Определение номера таблицы
        Dim TabN As Byte
'        TabN = .Tables.Count + 1 - 4
'Добавляем название таблицы
Selection.ParagraphFormat.Alignment = wdAlignParagraphRight
    Selection.TypeText Text:= _
        "Таблица №" & TabN & ". Определение пола индивидуума по черепу."
    Selection.TypeParagraph
    Selection.ParagraphFormat.Alignment = wdAlignParagraphJustify
    Selection.TypeText Text:="№ п/п." & vbTab & "Наименование размеров черепа" _
         & vbTab & "Диагностические размеры "
    Selection.TypeText Text:="(мм.)" & vbTab & "Показатели (м/ж)"
'Печать строк таблицы
 Dim i As Integer
        For i = 1 To 25
            Selection.TypeParagraph
            Selection.TypeText Text:=i & "." & vbTab & Массив(0, (i - 1)) & vbTab & _
                dimension(i - 1) & vbTab & Массив(3, (i - 1))
        Next i
'Оформление таблицы
    Selection.TypeParagraph
    Selection.MoveUp Unit:=wdLine, Count:=27, Extend:=wdExtend
    Selection.ConvertToTable Separator:=wdSeparateByTabs, NumColumns:=4, _
        NumRows:=26, AutoFitBehavior:=wdAutoFitContent
    With Selection.Tables(1)
        .Rows(1).Height = CentimetersToPoints(1.1)
        .Rows(2).Height = CentimetersToPoints(0.6)
        .Columns(1).Width = CentimetersToPoints(1.19)
        .Columns(2).Width = CentimetersToPoints(7.5)
        .Columns(3).Width = CentimetersToPoints(3.75)
        .Columns(4).Width = CentimetersToPoints(4.25)
        .Style = "Сетка таблицы"
        .ApplyStyleHeadingRows = True
        .ApplyStyleLastRow = False
        .ApplyStyleFirstColumn = True
        .ApplyStyleLastColumn = False
        .Columns(1).Select
            Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
            Selection.Cells.VerticalAlignment = wdCellAlignVerticalCenter
        .Columns(2).Select
            Selection.Cells.VerticalAlignment = wdCellAlignVerticalCenter
        .Columns(3).Select
            Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
            Selection.Cells.VerticalAlignment = wdCellAlignVerticalCenter
        .Columns(4).Select
            Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
            Selection.Cells.VerticalAlignment = wdCellAlignVerticalCenter
    End With
    Selection.MoveDown Unit:=wdLine, Count:=1
    Selection.TypeText Text:=selectCrForm(ByVal getCrForm) & ", " & selectCrHeight(ByVal getCrHeight) & ". " & _
    "Результаты проведенных измерений: " & ДМ.toString & ", " & ВМ.toString & ", " & Неопр.toString & ", " _
    & ВЖ.toString & ", " & ДЖ.toString & ", " & НПВ.toString & ". " & _
    "Следовательно, " & getSex
    End With
End Sub
'
Private Sub txtDimensions_Change()
'Функция при заполнении текстового поля "Измеренное значение в см"
    With txtDimensions
        If .Text <> "" Then
            Do
                If Not IsNumeric(.Text) Then ' Or Len(.Text) = 0
                    Beep
                    .BackColor = RGB(256, 0, 0)
                    MsgBox "Cледует вводить цифры", vbCritical, "Ошибка ввода"
                    .Text = InputBox("Введите правильно данные измерения!", "Исправление ошибки ввода")
                Else: CurDimensions = CCur(.Text)
                    .BackColor = RGB(200, 256, 200)
                    If CurDimensions = 0 Then
                        lblSex1.Caption = "Пол невозможно определить"
                        Call lblVisible
                    Else: Select_Dimensions ByVal CurDimensions
                    End If
            Exit Do
                End If
            Loop
        End If
    End With
'    CurDimensions = CCur(.Text)
'        If CurDimensions = 0 Then
'            lblSex1.Caption = "Пол невозможно определить"
'            Call lblVisible
'        Else: Select_Dimensions ByVal CurDimensions
'        End If
End Sub
'
Private Sub UserForm_Initialize()
'счетчиков определяемого пола:
'достоверно женский
    Set ДЖ = New Counter
        ДЖ.name = "достоверно женский"
'вероятно женский
    Set ВЖ = New Counter
        ВЖ.name = "вероятно женский"
'неопределенный
    Set Неопр = New Counter
        Неопр.name = "неопределенный"
'вероятно мужской
    Set ВМ = New Counter
        ВМ.name = "вероятно мужской"
'достоверно мужской
    Set ДМ = New Counter
        ДМ.name = "достоверно мужской"
'НПВ
    Set НПВ = New Counter
        НПВ.name = "НПВ"
'Счетчика
    Set N = New Counter
        N.name = "Ncounter"
'инициализация массива
'надписи:
    Массив(0, 0) = "Продольный диаметр"
        Массив(1, 0) = "glabella (g.) - opistokranion (op.)"
    Массив(0, 1) = "Поперечный диаметр"
        Массив(1, 1) = "euryon (eu.) - euryon (eu.)"
    Массив(0, 2) = "Высотный диаметр"
        Массив(1, 2) = "basion (ba.) - bregma (ba.)"
    Массив(0, 3) = "Длина основания черепа"
        Массив(1, 3) = "basion (ba.) - basion (ba.)"
    Массив(0, 4) = "Наименьшая ширина лба"
        Массив(1, 4) = "fronto-temporale (ft.) - fronto-temporale (ft.)"
    Массив(0, 5) = "Ширина основания черепа"
        Массив(1, 5) = "auriculare (au.) - auriculare (au.)"
    Массив(0, 6) = "Ширина затылка"
        Массив(1, 6) = "asterion (ast.) - asterion (ast.)"
    Массив(0, 7) = "Сосцевидная ширина"
        Массив(1, 7) = "mastoidale (ms.) - mastoidale (ms.)"
    Массив(0, 8) = "Окружность черепа"
        Массив(1, 8) = ""
    Массив(0, 9) = "Сагиттальная хорда"
        Массив(1, 9) = "nasion (n.) - opistion (o.)"
    Массив(0, 10) = "Лобная хорда"
        Массив(1, 10) = "nasion (n.) - bregma (b.)"
    Массив(0, 11) = "Теменная хорда"
        Массив(1, 11) = "bregma (b.) - lambda (l.)"
    Массив(0, 12) = "Длина БЗО"
        Массив(1, 12) = "basion (ba.) - opistion (o.)"
    Массив(0, 13) = "Ширина БЗО"
        Массив(1, 13) = ""
    Массив(0, 14) = "Скуловой диаметр"
        Массив(1, 14) = "zygion (zg.) - zygion (zg.)"
    Массив(0, 15) = "Длина основания лица"
        Массив(1, 15) = "basion (ba.) - prostion (pr.)"
    Массив(0, 16) = "Верхняя высота лица"
        Массив(1, 16) = "nasion (n.) - alveolare (al.)"
    Массив(0, 17) = "Полная высота лица"
        Массив(1, 17) = "nasion (n.) - gnation (gn.)"
    Массив(0, 18) = "Верхняя ширина лица"
        Массив(1, 18) = "fronto-malare-temporale (fmt.) - fronto-malare-temporale (fmt.)"
    Массив(0, 19) = "Средняя ширина лица"
        Массив(1, 19) = "zygomaxillare (zm.) - zygomaxillare (zm.)"
    Массив(0, 20) = "Высота носа"
        Массив(1, 20) = "nasion (n.) - nasospinale (ns.)"
    Массив(0, 21) = "Ширина левой орбиты"
        Массив(1, 21) = "maxillofrantale (mf.) - ektokonchion (ek.)"
    Массив(0, 22) = "Мыщелковая ширина"
        Массив(1, 22) = ""
    Массив(0, 23) = "Бигониальная ширина"
        Массив(1, 23) = "gonion (go.) - gonijon (go.)"
    Массив(0, 24) = "Высота тела НЧ"
        Массив(1, 24) = "gnation (gn.) - infradentale (id.)"
'картинки:
    Массив(2, 0) = IMGdir & "25.jpg"
    Массив(2, 1) = IMGdir & "1.jpg"
    Массив(2, 2) = IMGdir & "2.jpg"
    Массив(2, 3) = IMGdir & "3.jpg"
    Массив(2, 4) = IMGdir & "4.jpg"
    Массив(2, 5) = IMGdir & "5.jpg"
    Массив(2, 6) = IMGdir & "6.jpg"
    Массив(2, 7) = IMGdir & "7.jpg"
    Массив(2, 8) = IMGdir & "8.jpg"
    Массив(2, 9) = IMGdir & "9.jpg"
    Массив(2, 10) = IMGdir & "10.jpg"
    Массив(2, 11) = IMGdir & "11.jpg"
    Массив(2, 12) = IMGdir & "12.jpg"
    Массив(2, 13) = IMGdir & "13.jpg"
    Массив(2, 14) = IMGdir & "14.jpg"
    Массив(2, 15) = IMGdir & "15.jpg"
    Массив(2, 16) = IMGdir & "16.jpg"
    Массив(2, 17) = IMGdir & "17.jpg"
    Массив(2, 18) = IMGdir & "18.jpg"
    Массив(2, 19) = IMGdir & "19.jpg"
    Массив(2, 20) = IMGdir & "20.jpg"
    Массив(2, 21) = IMGdir & "21.jpg"
    Массив(2, 22) = IMGdir & "22.jpg"
    Массив(2, 23) = IMGdir & "23.jpg"
    Массив(2, 24) = IMGdir & "24.jpg"
'Открытие формы с первоначальными значениями
    lblMetod.Height = 38
    cmdToWord.Visible = False
    imgCraniumInv.BackColor = RGB(0, 164, 157)
'Вывод первой картинки при загрузке формы
    lblMetod.Caption = Массив(0, N.Count)
    lblMetod1.Caption = Массив(1, N.Count)
    imgCraniumInv.Picture = LoadPicture(Массив(2, N.Count))
End Sub
'
Private Static Function getCrForm() As Currency
    If dimension(1) > 0 And dimension(0) > 0 Then
        getCrForm = Format(dimension(1) * 100 / dimension(0), 0#)
    Else: getCrForm = 0
    End If
'Debug.Print "показатель формы =" & getCrForm
End Function
'
Private Static Function selectCrForm(ByVal getCrForm As Currency) As String
'Классификация черепов по форме(поперечный диаметр*100/продольный диаметр)
    'широкие (короткие) - брахикранные - >80%
    'средние - мезокранные - от 75% до 79,9%;
    'узкие (длинные) - долихокранные - <75%;
    Dim curTmp As Currency
        curTmp = getCrForm
        Select Case curTmp  'сравнение переменной с константами
            Case Is = 0
                selectCrForm = "Нельзя классифицировать череп по форме"
            Case 1 To 74, 9
                selectCrForm = "Исследуемый череп долихокранной формы " & "(" & curTmp & ")" 'значение, возвращаемое функцией
            Case 75 To 79.9
                selectCrForm = "Исследуемый череп мезокранной формы " & "(" & curTmp & ")"
            Case Is > 80
                selectCrForm = "Исследуемый череп брахикранной формы " & "(" & curTmp & ")"
        End Select
End Function
'
Private Static Function getCrHeight() As Currency
    If dimension(2) > 0 And dimension(0) > 0 Then
        getCrHeight = dimension(2) * 100 / dimension(0)
    Else: getCrHeight = 0
    End If
'Debug.Print "высотный показатель = " & getCrHeight
End Function
'
Private Static Function selectCrHeight(ByVal getCrHeight As Currency) As String
    Dim curTmp As Currency
        curTmp = getCrHeight
    Select Case curTmp  'сравнение переменной с константами
        Case Is = 0
            selectCrHeight = "нельзя классифицировать череп по высоте"
        Case 1 To 69, 9
            selectCrHeight = "низкий (хамекранный) " & "(" & curTmp & ")"
        Case 70 To 74, 9
            selectCrHeight = "средневысокий (ортокранный) " & "(" & curTmp & ")"
        Case Is > 75
            selectCrHeight = "высокий (гипсокранный) " & "(" & curTmp & ")"
    End Select
'Debug.Print "по высоте -" & selectCrHeight
End Function
'
Private Static Function getSex() As String
    If ДМ.Count > 12 Or ДМ.Count > ДЖ.Count Then
        getSex = "исследуемый череп принадлежал индивидууму мужского пола."
    ElseIf ДЖ.Count > 12 Or ДЖ.Count > ДМ.Count Then
        getSex = "исследуемый череп принадлежал индивидууму женского пола."
    Else
        If (ДМ.Count + ВМ.Count) > (ДЖ.Count + ВЖ.Count) Then
            getSex = "исследуемый череп мог принадлежать индивидууму мужского пола."
        ElseIf (ДЖ.Count + ВЖ.Count) > (ДМ.Count + ВМ.Count) Then
            getSex = "исследуемый череп мог принадлежать индивидууму женского пола."
        Else: getSex = "определить пол исследуемого черепа не представляется возможным."
        End If
    End If
End Function
'1
Private Static Function Select_Продольный_Диаметр(ByVal CurDimensions As Currency) As String
'Сравнение переменной "Продольный диаметр: глабелла-опистокраниум" с табличными данными
    If CurDimensions <= 0 Then 'Исключение нулевого значения
        Select_Продольный_Диаметр = "Пол невозможно определить"
    Else
        Select Case CurDimensions 'сравниваемая переменная, с уже присвоенным значением
            Case Is <= 160
                Select_Продольный_Диаметр = "Достоверно женский"
            Case 160.1 To 172
                Select_Продольный_Диаметр = "Вероятно женский"
            Case 172.1 To 178.5
                Select_Продольный_Диаметр = "Неопределенный"
            Case 178.6 To 187
                Select_Продольный_Диаметр = "Вероятно мужской"
            Case Is >= 187.1
                Select_Продольный_Диаметр = "Достоверно мужской"
            Case Else: MsgBox "Ошибка! Введенный параметр не является числом " _
                & "или выходит за границы допустимых значений!" & Err.Description
        End Select
    End If
End Function
'2
Private Static Function Select_Поперечный_Диаметр(ByVal CurDimensions As Currency) As String
'Сравнение переменной "Поперечный_Диаметр" с табличными данными
    If CurDimensions <= 0 Then 'Исключение нулевого значения
        Select_Поперечный_Диаметр = "Пол невозможно определить"
    Else
        Select Case CurDimensions 'сравниваемая переменная, с уже присвоенным значением
            Case Is <= 127
                Select_Поперечный_Диаметр = "Достоверно женский"
            Case 127.1 To 138
                Select_Поперечный_Диаметр = "Вероятно женский"
            Case 138.1 To 143
                Select_Поперечный_Диаметр = "Неопределенный"
            Case 143.1 To 152
                Select_Поперечный_Диаметр = "Вероятно мужской"
            Case Is >= 152.1
                Select_Поперечный_Диаметр = "Достоверно мужской"
            Case Else: MsgBox "Ошибка! Введенный параметр не является числом " _
                & "или выходит за границы допустимых значений!" & Err.Description
        End Select
    End If
End Function
'3
Private Static Function Select_Высотный_Диаметр(ByVal CurDimensions As Currency) As String
'Сравнение переменной "Высотный_Диаметр" с табличными данными
    If CurDimensions <= 0 Then 'Исключение нулевого значения
        Select_Высотный_Диаметр = "Пол невозможно определить"
    Else
        Select Case CurDimensions 'сравниваемая переменная, с уже присвоенным значением
            Case Is <= 121
                Select_Высотный_Диаметр = "Достоверно женский"
            Case 121.1 To 128
                Select_Высотный_Диаметр = "Вероятно женский"
            Case 128.1 To 134
                Select_Высотный_Диаметр = "Неопределенный"
            Case 134.1 To 140.5
                Select_Высотный_Диаметр = "Вероятно мужской"
            Case Is >= 140.6
                Select_Высотный_Диаметр = "Достоверно мужской"
           Case Else: MsgBox "Ошибка! Введенный параметр не является числом " _
                & "или выходит за границы допустимых значений!" & Err.Description
        End Select
    End If
End Function
'4
Private Static Function Select_Длина_Основания(ByVal CurDimensions As Currency) As String
'Сравнение переменной "Длина_Основания" с табличными данными
    If CurDimensions <= 0 Then 'Исключение нулевого значения
        Select_Длина_Основания = "Пол невозможно определить"
    Else
        Select Case CurDimensions 'сравниваемая переменная, с уже присвоенным значением
            Case Is <= 90
                Select_Длина_Основания = "Достоверно женский"
            Case 90.1 To 96
                Select_Длина_Основания = "Вероятно женский"
            Case 96.1 To 101
                Select_Длина_Основания = "Неопределенный"
            Case 101.1 To 109
                Select_Длина_Основания = "Вероятно мужской"
            Case Is >= 109.1
                Select_Длина_Основания = "Достоверно мужской"
            Case Else: MsgBox "Ошибка! Введенный параметр не является числом " _
                & "или выходит за границы допустимых значений!" & Err.Description
        End Select
    End If
End Function
'5
Private Static Function Select_Min_ширина_Лба(ByVal CurDimensions As Currency) As String
'Сравнение переменной "Min_ширина_Лба" с табличными данными
    If CurDimensions <= 0 Then 'Исключение нулевого значения
        Select_Min_ширина_Лба = "Пол невозможно определить"
    Else
        Select Case CurDimensions 'сравниваемая переменная, с уже присвоенным значением
            Case Is <= 86
                Select_Min_ширина_Лба = "Достоверно женский"
            Case 86.1 To 95
                Select_Min_ширина_Лба = "Вероятно женский"
            Case 95.1 To 98
                Select_Min_ширина_Лба = "Неопределенный"
            Case 98.1 To 108
                Select_Min_ширина_Лба = "Вероятно мужской"
            Case Is >= 108.1
                Select_Min_ширина_Лба = "Достоверно мужской"
           Case Else: MsgBox "Ошибка! Введенный параметр не является числом " _
                & "или выходит за границы допустимых значений!" & Err.Description
        End Select
    End If
End Function
'6
Private Static Function Select_Ширина_Основания(ByVal CurDimensions As Currency) As String
'Сравнение переменной "Ширина_Основания" с табличными данными
    If CurDimensions <= 0 Then 'Исключение нулевого значения
        Select_Ширина_Основания = "Пол невозможно определить"
    Else
        Select Case CurDimensions 'сравниваемая переменная, с уже присвоенным значением
            Case Is <= 112
                Select_Ширина_Основания = "Достоверно женский"
            Case 112.1 To 117
                Select_Ширина_Основания = "Вероятно женский"
            Case 117.1 To 123
                Select_Ширина_Основания = "Неопределенный"
            Case 123.1 To 133
                Select_Ширина_Основания = "Вероятно мужской"
            Case Is >= 133.1
                Select_Ширина_Основания = "Достоверно мужской"
            Case Else: MsgBox "Ошибка! Введенный параметр не является числом " _
                & "или выходит за границы допустимых значений!" & Err.Description
        End Select
    End If
End Function
'7
Private Static Function Select_Ширина_Затылка(ByVal CurDimensions As Currency) As String
'Сравнение переменной "Ширина_Затылка" с табличными данными
    If CurDimensions <= 0 Then 'Исключение нулевого значения
        Select_Ширина_Затылка = "Пол невозможно определить"
    Else
        Select Case CurDimensions 'сравниваемая переменная, с уже присвоенным значением
            Case Is <= 99
                Select_Ширина_Затылка = "Достоверно женский"
            Case 99.1 To 106.9
                Select_Ширина_Затылка = "Вероятно женский"
            Case 107 To 110.4
                Select_Ширина_Затылка = "Неопределенный"
            Case 110.5 To 120
                Select_Ширина_Затылка = "Вероятно мужской"
            Case Is >= 120.1
                Select_Ширина_Затылка = "Достоверно мужской"
            Case Else: MsgBox "Ошибка! Введенный параметр не является числом " _
                & "или выходит за границы допустимых значений!" & Err.Description
        End Select
    End If
End Function
'8
Private Static Function Select_Сосцевидная_Ширина(ByVal CurDimensions As Currency) As String
'Сравнение переменной "Сосцевидная_Ширина" с табличными данными
    If CurDimensions <= 0 Then 'Исключение нулевого значения
        Select_Сосцевидная_Ширина = "Пол невозможно определить"
    Else
        Select Case CurDimensions 'сравниваемая переменная, с уже присвоенным значением
            Case Is <= 92
                Select_Сосцевидная_Ширина = "Достоверно женский"
            Case 92.1 To 100
                Select_Сосцевидная_Ширина = "Вероятно женский"
            Case 100.1 To 105
                Select_Сосцевидная_Ширина = "Неопределенный"
            Case 105.1 To 116
                Select_Сосцевидная_Ширина = "Вероятно мужской"
            Case Is >= 116.1
                Select_Сосцевидная_Ширина = "Достоверно мужской"
            Case Else: MsgBox "Ошибка! Введенный параметр не является числом " _
                & "или выходит за границы допустимых значений!" & Err.Description
        End Select
    End If
End Function
'9
Private Static Function Select_Окружность_Черепа(ByVal CurDimensions As Currency) As String
'Сравнение переменной "Окружность_Черепа" с табличными данными
    If CurDimensions <= 0 Then 'Исключение нулевого значения
        Select_Окружность_Черепа = "Пол невозможно определить"
    Else
        Select Case CurDimensions 'сравниваемая переменная, с уже присвоенным значением
            Case Is <= 476
                Select_Окружность_Черепа = "Достоверно женский"
            Case 476.1 To 500.5
                Select_Окружность_Черепа = "Вероятно женский"
            Case 500.6 To 516.5
                Select_Окружность_Черепа = "Неопределенный"
            Case 516.6 To 540
                Select_Окружность_Черепа = "Вероятно мужской"
            Case Is >= 540.1
                Select_Окружность_Черепа = "Достоверно мужской"
           Case Else: MsgBox "Ошибка! Введенный параметр не является числом " _
                & "или выходит за границы допустимых значений!" & Err.Description
        End Select
    End If
End Function
'10
Private Static Function Select_Сагиттальная_Хорда(ByVal CurDimensions As Currency) As String
'Сравнение переменной "Сагиттальная_Хорда" с табличными данными
    If CurDimensions <= 0 Then 'Исключение нулевого значения
        Select_Сагиттальная_Хорда = "Пол невозможно определить"
    Else
        Select Case CurDimensions 'сравниваемая переменная, с уже присвоенным значением
            Case Is <= 123
                Select_Сагиттальная_Хорда = "Достоверно женский"
            Case 123.1 To 128.5
                Select_Сагиттальная_Хорда = "Вероятно женский"
            Case 128.6 To 134.5
                Select_Сагиттальная_Хорда = "Неопределенный"
            Case 134.6 To 145
                Select_Сагиттальная_Хорда = "Вероятно мужской"
            Case Is >= 145.1
                Select_Сагиттальная_Хорда = "Достоверно мужской"
            Case Else: MsgBox "Ошибка! Введенный параметр не является числом " _
                & "или выходит за границы допустимых значений!" & Err.Description
        End Select
    End If
End Function
'11
Private Static Function Select_Лобная_Хорда(ByVal CurDimensions As Currency) As String
'Сравнение переменной "Лобная_хорда" с табличными данными
    If CurDimensions <= 0 Then 'Исключение нулевого значения
        Select_Лобная_Хорда = "Пол невозможно определить"
    Else
        Select Case CurDimensions  'сравниваемая переменная, с уже присвоенным значением
            Case Is <= 99
                Select_Лобная_Хорда = "Достоверно женский"
            Case 99.1 To 107
                Select_Лобная_Хорда = "Вероятно женский"
            Case 107.1 To 111.5
                Select_Лобная_Хорда = "Неопределенный"
            Case 111.6 To 121
                Select_Лобная_Хорда = "Вероятно мужской"
            Case Is >= 121.1
                Select_Лобная_Хорда = "Достоверно мужской"
            Case Else: MsgBox "Ошибка! Введенный параметр не является числом " _
                & "или выходит за границы допустимых значений!" & Err.Description
        End Select
    End If
End Function
'12
Private Static Function Select_Теменная_Хорда(ByVal CurDimensions As Currency) As String
'Сравнение переменной " Теменная_Хорда" с табличными данными
    If CurDimensions <= 0 Then 'Исключение нулевого значения
        Select_Теменная_Хорда = "Пол невозможно определить"
    Else
        Select Case CurDimensions  'сравниваемая переменная, с уже присвоенным значением
            Case Is <= 94
                Select_Теменная_Хорда = "Достоверно женский"
            Case 94.1 To 107
                Select_Теменная_Хорда = "Вероятно женский"
            Case 107.1 To 110.5
                Select_Теменная_Хорда = "Неопределенный"
            Case 110.6 To 124
                Select_Теменная_Хорда = "Вероятно мужской"
            Case Is >= 124.1
                Select_Теменная_Хорда = "Достоверно мужской"
            Case Else: MsgBox "Ошибка! Введенный параметр не является числом " _
                & "или выходит за границы допустимых значений!" & Err.Description
        End Select
    End If
End Function
'13
Private Static Function Select_Длина_БЗО(ByVal CurDimensions As Currency) As String
'Сравнение переменной "Длина_БЗО" с табличными данными
    If CurDimensions <= 0 Then 'Исключение нулевого значения
        Select_Длина_БЗО = "Пол невозможно определить"
    Else
        Select Case CurDimensions 'сравниваемая переменная, с уже присвоенным значением
            Case Is <= 30
                Select_Длина_БЗО = "Достоверно женский"
            Case 30.1 To 34
                Select_Длина_БЗО = "Вероятно женский"
            Case 34.1 To 36
                Select_Длина_БЗО = "Неопределенный"
            Case 36.1 To 41
                Select_Длина_БЗО = "Вероятно мужской"
            Case Is >= 41.1
                Select_Длина_БЗО = "Достоверно мужской"
        Case Else: MsgBox "Ошибка! Введенный параметр не является числом " _
                & "или выходит за границы допустимых значений!" & Err.Description
        End Select
    End If
End Function
'14
Private Static Function Select_Ширина_БЗО(ByVal CurDimensions As Currency) As String
'Сравнение переменной "Ширина_БЗО" с табличными данными
    If CurDimensions <= 0 Then 'Исключение нулевого значения
        Select_Ширина_БЗО = "Пол невозможно определить"
    Else
        Select Case CurDimensions 'сравниваемая переменная, с уже присвоенным значением
            Case Is <= 25
                Select_Ширина_БЗО = "Достоверно женский"
            Case 25.1 To 28.5
                Select_Ширина_БЗО = "Вероятно женский"
            Case 28.6 To 30.5
                Select_Ширина_БЗО = "Неопределенный"
            Case 30.6 To 35
                Select_Ширина_БЗО = "Вероятно мужской"
            Case Is >= 35.1
                Select_Ширина_БЗО = "Достоверно мужской"
            Case Else: MsgBox "Ошибка! Введенный параметр не является числом " _
                & "или выходит за границы допустимых значений!" & Err.Description
        End Select
    End If
End Function
'15
Private Static Function Select_Скуловой_Диаметр(ByVal CurDimensions As Currency) As String
'Сравнение переменной "Скуловой_Диаметр" с табличными данными
    If CurDimensions <= 0 Then 'Исключение нулевого значения
        Select_Скуловой_Диаметр = "Пол невозможно определить"
    Else
        Select Case CurDimensions 'сравниваемая переменная, с уже присвоенным значением
            Case Is <= 120
                Select_Скуловой_Диаметр = "Достоверно женский"
            Case 120.1 To 124
                Select_Скуловой_Диаметр = "Вероятно женский"
            Case 124.1 To 132
                Select_Скуловой_Диаметр = "Неопределенный"
            Case 132.1 To 139
                Select_Скуловой_Диаметр = "Вероятно мужской"
            Case Is >= 139.1
                Select_Скуловой_Диаметр = "Достоверно мужской"
           Case Else: MsgBox "Ошибка! Введенный параметр не является числом " _
                & "или выходит за границы допустимых значений!" & Err.Description
        End Select
    End If
End Function
'16
Private Static Function Select_Длина_Основания_Лица(ByVal CurDimensions As Currency) As String
'Сравнение переменной "Длина_Основания_Лица" с табличными данными
    If CurDimensions <= 0 Then 'Исключение нулевого значения
        Select_Длина_Основания_Лица = "Пол невозможно определить"
    Else
        Select Case CurDimensions 'сравниваемая переменная, с уже присвоенным значением
            Case Is <= 82
                Select_Длина_Основания_Лица = "Достоверно женский"
            Case 82.1 To 93
                Select_Длина_Основания_Лица = "Вероятно женский"
            Case 93.1 To 97.5
                Select_Длина_Основания_Лица = "Неопределенный"
            Case 97.6 To 107
                Select_Длина_Основания_Лица = "Вероятно мужской"
            Case Is >= 107.1
                Select_Длина_Основания_Лица = "Достоверно мужской"
            Case Else: MsgBox "Ошибка! Введенный параметр не является числом " _
                & "или выходит за границы допустимых значений!" & Err.Description
        End Select
    End If
End Function
'17
Private Static Function Select_Верхняя_Высота_Лица(ByVal CurDimensions As Currency) As String
'Сравнение переменной "Верхняя_Высота_Лица" с табличными данными
    If CurDimensions <= 0 Then 'Исключение нулевого значения
        Select_Верхняя_Высота_Лица = "Пол невозможно определить"
    Else
        Select Case CurDimensions 'сравниваемая переменная, с уже присвоенным значением
            Case Is <= 59
                Select_Верхняя_Высота_Лица = "Достоверно женский"
            Case 59.1 To 66.5
                Select_Верхняя_Высота_Лица = "Вероятно женский"
            Case 66.6 To 71
                Select_Верхняя_Высота_Лица = "Неопределенный"
            Case 71.1 To 78
                Select_Верхняя_Высота_Лица = "Вероятно мужской"
            Case Is >= 78.1
                Select_Верхняя_Высота_Лица = "Достоверно мужской"
            Case Else: MsgBox "Ошибка! Введенный параметр не является числом " _
                & "или выходит за границы допустимых значений!" & Err.Description
        End Select
    End If
End Function
'18
Private Static Function Select_Полная_Высота_Лица(ByVal CurDimensions As Currency) As String
'Сравнение переменной "Полная_Высота_Лица" с табличными данными
    If CurDimensions <= 0 Then 'Исключение нулевого значения
        Select_Полная_Высота_Лица = "Пол невозможно определить"
    Else
        Select Case CurDimensions 'сравниваемая переменная, с уже присвоенным значением
            Case Is <= 100
                Select_Полная_Высота_Лица = "Достоверно женский"
            Case 100.1 To 111
                Select_Полная_Высота_Лица = "Вероятно женский"
            Case 111.1 To 119
                Select_Полная_Высота_Лица = "Неопределенный"
            Case 119.1 To 132
                Select_Полная_Высота_Лица = "Вероятно мужской"
            Case Is >= 132.1
                Select_Полная_Высота_Лица = "Достоверно мужской"
            Case Else: MsgBox "Ошибка! Введенный параметр не является числом " _
                & "или выходит за границы допустимых значений!" & Err.Description
        End Select
    End If
End Function
'19
Private Static Function Select_Верхняя_Ширина_Лица(ByVal CurDimensions As Currency) As String
'Сравнение переменной "Верхняя_Ширина_Лица" с табличными данными
    If CurDimensions <= 0 Then 'Исключение нулевого значения
        Select_Верхняя_Ширина_Лица = "Пол невозможно определить"
    Else
        Select Case CurDimensions 'сравниваемая переменная, с уже присвоенным значением
            Case Is <= 93
                Select_Верхняя_Ширина_Лица = "Достоверно женский"
            Case 93.1 To 101
                Select_Верхняя_Ширина_Лица = "Вероятно женский"
            Case 101.1 To 105
                Select_Верхняя_Ширина_Лица = "Неопределенный"
            Case 105.1 To 113
                Select_Верхняя_Ширина_Лица = "Вероятно мужской"
            Case Is >= 113.1
                Select_Верхняя_Ширина_Лица = "Достоверно мужской"
            Case Else: MsgBox "Ошибка! Введенный параметр не является числом " _
                & "или выходит за границы допустимых значений!" & Err.Description
        End Select
    End If
End Function
'20
Private Static Function Select_Средняя_Ширина_Лица(ByVal CurDimensions As Currency) As String
'Сравнение переменной "Средняя_Ширина_Лица"с табличными данными
    If CurDimensions <= 0 Then 'Исключение нулевого значения
        Select_Средняя_Ширина_Лица = "Пол невозможно определить"
    Else
        Select Case CurDimensions 'сравниваемая переменная, с уже присвоенным значением
            Case Is <= 78
                Select_Средняя_Ширина_Лица = "Достоверно женский"
            Case 78.1 To 89
                Select_Средняя_Ширина_Лица = "Вероятно женский"
            Case 89.1 To 93.5
                Select_Средняя_Ширина_Лица = "Неопределенный"
            Case 93.6 To 104
                Select_Средняя_Ширина_Лица = "Вероятно мужской"
            Case Is >= 104.1
                Select_Средняя_Ширина_Лица = "Достоверно мужской"
           Case Else: MsgBox "Ошибка! Введенный параметр не является числом " _
                & "или выходит за границы допустимых значений!" & Err.Description
        End Select
    End If
End Function
'21
Private Static Function Select_Высота_Носа(ByVal CurDimensions As Currency) As String
'Сравнение переменной "Высота_Носа"с табличными данными
    If CurDimensions <= 0 Then 'Исключение нулевого значения
        Select_Высота_Носа = "Пол невозможно определить"
    Else
        Select Case CurDimensions 'сравниваемая переменная, с уже присвоенным значением
            Case Is <= 44
                Select_Высота_Носа = "Достоверно женский"
            Case 44.1 To 48.5
                Select_Высота_Носа = "Вероятно женский"
            Case 48.6 To 52
                Select_Высота_Носа = "Неопределенный"
            Case 52.1 To 56
                Select_Высота_Носа = "Вероятно мужской"
            Case Is >= 56.1
                Select_Высота_Носа = "Достоверно мужской"
            Case Else: MsgBox "Ошибка! Введенный параметр не является числом " _
                & "или выходит за границы допустимых значений!" & Err.Description
        End Select
    End If
End Function
'22
Private Static Function Select_Ширина_Орбиты(ByVal CurDimensions As Currency) As String
'Сравнение переменной "Ширина_Орбиты" с табличными данными
    If CurDimensions <= 0 Then 'Исключение нулевого значения
        Select_Ширина_Орбиты = "Пол невозможно определить"
    Else
        Select Case CurDimensions 'сравниваемая переменная, с уже присвоенным значением
            Case Is <= 38
                Select_Ширина_Орбиты = "Достоверно женский"
            Case 38.1 To 42
                Select_Ширина_Орбиты = "Вероятно женский"
            Case 42.1 To 43.5
                Select_Ширина_Орбиты = "Неопределенный"
            Case 43.6 To 48
                Select_Ширина_Орбиты = "Вероятно мужской"
            Case Is >= 48.1
                Select_Ширина_Орбиты = "Достоверно мужской"
            Case Else: MsgBox "Ошибка! Введенный параметр не является числом " _
                & "или выходит за границы допустимых значений!" & Err.Description
        End Select
    End If
End Function
'23
Private Static Function Select_Мыщелковая_Ширина(ByVal CurDimensions As Currency) As String
'Сравнение переменной "Мыщелковая_Ширина"с табличными данными
    If CurDimensions <= 0 Then 'Исключение нулевого значения
        Select_Мыщелковая_Ширина = "Пол невозможно определить"
    Else
        Select Case CurDimensions 'сравниваемая переменная, с уже присвоенным значением
            Case Is <= 105
            Select_Мыщелковая_Ширина = "Достоверно женский"
            Case 105.1 To 113.5
                Select_Мыщелковая_Ширина = "Вероятно женский"
            Case 113.6 To 118.5
                Select_Мыщелковая_Ширина = "Неопределенный"
            Case 118.6 To 127
                Select_Мыщелковая_Ширина = "Вероятно мужской"
            Case Is >= 127.1
                Select_Мыщелковая_Ширина = "Достоверно мужской"
           Case Else: MsgBox "Ошибка! Введенный параметр не является числом " _
                & "или выходит за границы допустимых значений!" & Err.Description
        End Select
    End If
End Function
'24
Private Static Function Select_Бигониальная_Ширина(ByVal CurDimensions As Currency) As String
'Сравнение переменной "Бигониальная_Ширина"с табличными данными
    If CurDimensions <= 0 Then 'Исключение нулевого значения
        Select_Бигониальная_Ширина = "Пол невозможно определить"
    Else
        Select Case CurDimensions 'сравниваемая переменная, с уже присвоенным значением
            Case Is <= 85
                Select_Бигониальная_Ширина = "Достоверно женский"
            Case 85.1 To 95
                Select_Бигониальная_Ширина = "Вероятно женский"
            Case 95.1 To 102.5
                Select_Бигониальная_Ширина = "Неопределенный"
            Case 102.6 To 112
                Select_Бигониальная_Ширина = "Вероятно мужской"
            Case Is >= 112.1
                Select_Бигониальная_Ширина = "Достоверно мужской"
            Case Else: MsgBox "Ошибка! Введенный параметр не является числом " _
                & "или выходит за границы допустимых значений!" & Err.Description
        End Select
    End If
End Function
'25
Private Static Function Select_Высота_ТелаНЧ(ByVal CurDimensions As Currency) As String
'Сравнение переменной "Высота_ТелаНЧ" с табличными данными
    If CurDimensions <= 0 Then 'Исключение нулевого значения
        Select_Высота_ТелаНЧ = "Пол невозможно определить"
    Else
        Select Case CurDimensions 'сравниваемая переменная, с уже присвоенным значением
            Case Is <= 27
                Select_Высота_ТелаНЧ = "Достоверно женский"
            Case 27.1 To 31
                Select_Высота_ТелаНЧ = "Вероятно женский"
            Case 31.1 To 33.5
                Select_Высота_ТелаНЧ = "Неопределенный"
            Case 33.6 To 41
                Select_Высота_ТелаНЧ = "Вероятно мужской"
            Case Is >= 41.1
                Select_Высота_ТелаНЧ = "Достоверно мужской"
            Case Else: MsgBox "Ошибка! Введенный параметр не является числом " _
                & "или выходит за границы допустимых значений!" & Err.Description
        End Select
    End If
End Function
'
Private Static Sub Select_CrParameter()
    Select Case Массив(3, N.Count) 'сравниваемая переменная, с уже присвоенным значением
        Case Is = "Достоверно женский"
            ДЖ.increment
        Case Is = "Вероятно женский"
            ВЖ.increment
        Case Is = "Неопределенный"
            Неопр.increment
        Case Is = "Вероятно мужской"
            ВМ.increment
        Case Is = "Достоверно мужской"
            ДМ.increment
        Case Else: НПВ.increment
    End Select
End Sub
'
Private Sub UserForm_Terminate()
'счетчиков определяемого пола:
'достоверно женский
    Set ДЖ = Nothing
    Set ВЖ = Nothing
    Set Неопр = Nothing
    Set ВМ = Nothing
    Set ДМ = Nothing
    Set НПВ = Nothing
    Set N = Nothing
    Unload Me
End Sub
