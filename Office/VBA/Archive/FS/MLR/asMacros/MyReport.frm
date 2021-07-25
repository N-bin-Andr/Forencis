VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} MyReport 
   Caption         =   "Форма для расчета нагрузки за месяц"
   ClientHeight    =   12810
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10215
   OleObjectBlob   =   "MyReport.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "MyReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
'@author Andr.N@_bin
'@E-mail Andr.Nab.n@gmail.com
Private caseData() As String  ' массив с данными для печати
Private sum As Single 'сумма коэффициентов
Private W As Currency
Const LowBound As Integer = 0
Private HihtBound As Integer ' max значение массива (счетчик).
Const PATH As String = "D:\Crime\2019\Отпечатанные СМЭ\"
Private myPages As MultiPage
Private numEF As String
'
Private Sub UserForm_Initialize()
'конструктор
    HihtBound = LowBound
    Set myPages = MultiPage1
       With myPages
        .BackColor = &HCFC2AC
    End With
    With ActiveDocument
        numEF = "Заключение №" & .FormFields("WdEFNum").Result
        With Me.lblEFNum
            .Caption = numEF
            .ForeColor = &HFF0000
'            .BackColor = &HC0FFFF
        End With
    End With
    Debug.Print numEF
End Sub
'
Private Sub cmdSave_Click()
'подсчет, сохранение и печать данных
'1) обработка переключателей
Dim tmpObj As Object 'объявление объекта
    For Each tmpObj In Me.Controls
        If TypeName(tmpObj) = "OptionButton" Then
            With tmpObj
                If .Value = True Then 'если нажат
                    ReDim Preserve caseData(LowBound To HihtBound + 1)  'расширяем массив
                    caseData(HihtBound) = .Caption & " - " & .Tag 'присваиваем ему значение
                    HihtBound = HihtBound + 1  'увеличиваем счетчик
                    'c) подсчет суммы
                        sum = sum + CSng(.Tag)
                End If
            End With
        End If
    Next tmpObj
   'd) подсчет нагрузки
    W = (sum / 8 * 100)
Call writingFile
'Call cmbClear_Click
Me.Hide
End Sub
'
Private Sub writingFile()
'запись в файл
Dim fileDescr As Integer
    fileDescr = FreeFile
'1) определение текущей даты:
    Dim dt As Date
        dt = DateTime.Date
    Dim str1 As String
        str1 = VBA.Format(dt, "dd/mm/yyyy") ' текущая дата
    Dim str2 As String
        str2 = VBA.Format(dt, "mmmm/yyyy") 'текущий месяц/год
        Debug.Print str2
'2)печать в файл:
 Dim newFile As String, newNum As Integer
    newFile = str2 & "_Отчет" & ".txt"
    Open PATH & newFile For Append As #fileDescr
        Print #fileDescr, str1 'печать текущей даты;
        Print #fileDescr, numEF
'        печать массива данных
        Dim i As Integer
            For i = LowBound To HihtBound
                Print #fileDescr, caseData(i)
            Next i
        'печать суммы
        Print #fileDescr, "Сумма коэффициентов = "; sum
        Print #fileDescr, "Нагрузка = "; W; "%"
        Print #fileDescr, "-------------------------------------------------------------------------------------------------------"
    Close #fileDescr
End Sub
'
Private Sub cmbClear_Click()
'Очистка формы
    Dim tmpObj As Object
    For Each tmpObj In Me.Controls
        If TypeName(tmpObj) = "TextBox" Then
            With tmpObj
                .Text = ""
                .BackColor = &H80000005
            End With
        ElseIf TypeName(tmpObj) = "OptionButton" Then
            With tmpObj
                    .Value = 0
                    .BackColor = &H8000000F
            End With
        End If
    Next tmpObj
    sum = 0
    W = 0
   Erase caseData()
End Sub
'
Private Sub opbChange(tmpObj As Object)
'изменение переключателей
    With tmpObj
       If .Value = -1 Then
            .BackColor = &HDBF8F4
        Else: .BackColor = &H8000000F
        End If
    End With
End Sub
'
Private Sub OptionButton1_Change()
    Call opbChange(OptionButton1)
End Sub
'
Private Sub OptionButton10_Change()
    Call opbChange(OptionButton10)
End Sub
'
Private Sub OptionButton11_Change()
    Call opbChange(OptionButton11)
End Sub
'
Private Sub OptionButton12_Change()
    Call opbChange(OptionButton12)
End Sub
'
Private Sub OptionButton13_Change()
    Call opbChange(OptionButton13)
End Sub
'
Private Sub OptionButton14_Change()
    Call opbChange(OptionButton14)
End Sub
'
Private Sub OptionButton15_Change()
    Call opbChange(OptionButton15)
End Sub
'
Private Sub OptionButton16_Change()
    Call opbChange(OptionButton16)
End Sub
'
Private Sub OptionButton17_Change()
    Call opbChange(OptionButton17)
End Sub
'
Private Sub OptionButton18_Change()
    Call opbChange(OptionButton18)
End Sub
'
Private Sub OptionButton19_Change()
    Call opbChange(OptionButton19)
End Sub
'
Private Sub OptionButton2_Change()
    Call opbChange(OptionButton2)
End Sub
'
Private Sub OptionButton20_Change()
    Call opbChange(OptionButton20)
End Sub
'
Private Sub OptionButton21_Change()
    Call opbChange(OptionButton21)
End Sub

Private Sub OptionButton22_Change()
Call opbChange(OptionButton22)
End Sub
'
Private Sub OptionButton23_Change()
    Call opbChange(OptionButton23)
End Sub
'
Private Sub OptionButton24_Change()
    Call opbChange(OptionButton24)
End Sub
'
Private Sub OptionButton25_Change()
    Call opbChange(OptionButton25)
End Sub
'
Private Sub OptionButton26_Change()
    Call opbChange(OptionButton26)
End Sub
'
Private Sub OptionButton27_Change()
    Call opbChange(OptionButton27)
End Sub
'
Private Sub OptionButton28_Change()
    Call opbChange(OptionButton28)
End Sub
'
Private Sub OptionButton29_Change()
    Call opbChange(OptionButton29)
End Sub
'
Private Sub OptionButton3_Change()
    Call opbChange(OptionButton3)
End Sub
'
Private Sub OptionButton30_Change()
    Call opbChange(OptionButton30)
End Sub
'
Private Sub OptionButton31_Change()
    Call opbChange(OptionButton31)
End Sub
'
Private Sub OptionButton32_Change()
    Call opbChange(OptionButton32)
End Sub
'
Private Sub OptionButton33_Change()
    Call opbChange(OptionButton33)
End Sub
'
Private Sub OptionButton34_Change()
    Call opbChange(OptionButton34)
End Sub
'
Private Sub OptionButton35_Change()
    Call opbChange(OptionButton35)
End Sub
'
Private Sub OptionButton36_Change()
    Call opbChange(OptionButton36)
End Sub
'
Private Sub OptionButton37_Change()
    Call opbChange(OptionButton37)
End Sub
'
Private Sub OptionButton38_Change()
    Call opbChange(OptionButton38)
End Sub
'
Private Sub OptionButton39_Change()
    Call opbChange(OptionButton39)
End Sub
'
Private Sub OptionButton4_Change()
Call opbChange(OptionButton4)
End Sub
'
Private Sub OptionButton40_Change()
    Call opbChange(OptionButton40)
End Sub

Private Sub OptionButton41_Change()
    Call opbChange(OptionButton41)
End Sub
'
Private Sub OptionButton42_Change()
    Call opbChange(OptionButton42)
End Sub
'
Private Sub OptionButton43_Change()
    Call opbChange(OptionButton43)
End Sub
'
Private Sub OptionButton44_Change()
    Call opbChange(OptionButton44)
End Sub
'
Private Sub OptionButton45_Change()
    Call opbChange(OptionButton45)
End Sub

Private Sub OptionButton46_Change()
    Call opbChange(OptionButton46)
End Sub
'
Private Sub OptionButton47_Change()
    Call opbChange(OptionButton47)
End Sub
'
Private Sub OptionButton48_Change()
    Call opbChange(OptionButton48)
End Sub
'
Private Sub OptionButton49_Change()
    Call opbChange(OptionButton49)
End Sub
'
Private Sub OptionButton5_Change()
    Call opbChange(OptionButton5)
End Sub
'
Private Sub OptionButton50_Change()
    Call opbChange(OptionButton50)
End Sub
'
Private Sub OptionButton51_Change()
    Call opbChange(OptionButton51)
End Sub
'
Private Sub OptionButton52_Change()
    Call opbChange(OptionButton52)
End Sub
'
Private Sub OptionButton53_Change()
    Call opbChange(OptionButton53)
End Sub
'
Private Sub OptionButton54_Change()
    Call opbChange(OptionButton54)
End Sub
'
Private Sub OptionButton55_Change()
    Call opbChange(OptionButton55)
End Sub
'
Private Sub OptionButton56_Change()
    Call opbChange(OptionButton56)
End Sub
'
Private Sub OptionButton57_Change()
    Call opbChange(OptionButton57)
End Sub
'
Private Sub OptionButton58_Change()
    Call opbChange(OptionButton58)
End Sub
'
Private Sub OptionButton6_Change()
    Call opbChange(OptionButton6)
End Sub
'
Private Sub OptionButton7_Change()
    Call opbChange(OptionButton7)
End Sub

Private Sub OptionButton8_Change()
    Call opbChange(OptionButton8)
End Sub
'
Private Sub OptionButton9_Change()
    Call opbChange(OptionButton9)
End Sub
