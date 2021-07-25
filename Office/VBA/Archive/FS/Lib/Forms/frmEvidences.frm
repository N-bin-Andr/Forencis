VERSION 5.00
Begin VB.Form frmEvidences 
   BackColor       =   &H00CFC2AC&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Вещественные доказательства"
   ClientHeight    =   9540
   ClientLeft      =   45
   ClientTop       =   465
   ClientWidth     =   15795
   BeginProperty Font 
      Name            =   "Calibri"
      Size            =   11.25
      Charset         =   204
      Weight          =   400
      Underline       =   0   'False
      Italic          =   -1  'True
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9540
   ScaleWidth      =   15795
   Begin VB.CheckBox chbNoMedCrim 
      BackColor       =   &H00CFC2AC&
      Caption         =   "Упаковка с объектами не для мед.крим. иссл."
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   28
      Top             =   8280
      Width           =   4935
   End
   Begin VB.TextBox txtBxStamp 
      Alignment       =   2  'Центровка
      BackColor       =   &H00FFE8DF&
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   12
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1485
      Left            =   1440
      MultiLine       =   -1  'True
      TabIndex        =   4
      Top             =   3840
      Width           =   3855
   End
   Begin VB.ComboBox cmbBxPackage 
      BackColor       =   &H00FFE8DF&
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   12
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   1440
      TabIndex        =   3
      Top             =   3360
      Width           =   3855
   End
   Begin VB.CommandButton cmdBxClear 
      BackColor       =   &H008080FF&
      Caption         =   "&Очистить"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   8760
      Width           =   1455
   End
   Begin VB.TextBox txtBxFirstDate 
      Alignment       =   2  'Центровка
      BackColor       =   &H00FFE8DF&
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   12
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   406
      Left            =   3360
      MaxLength       =   10
      TabIndex        =   1
      Top             =   840
      Width           =   1935
   End
   Begin VB.CommandButton cmdSaveOK 
      BackColor       =   &H00C0FFC0&
      Caption         =   "&Сохранить"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3720
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   8760
      Width           =   1575
   End
   Begin VB.TextBox txtBxOutTake 
      Alignment       =   2  'Центровка
      BackColor       =   &H00FFE8DF&
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   12
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   406
      Left            =   2760
      MaxLength       =   10
      TabIndex        =   0
      Top             =   360
      Width           =   2535
   End
   Begin VB.ListBox lstEvList 
      Columns         =   1
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   12
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   8610
      Left            =   5640
      TabIndex        =   20
      Top             =   720
      Width           =   9615
   End
   Begin VB.TextBox txtBxPlace 
      Alignment       =   2  'Центровка
      BackColor       =   &H00FFE8DF&
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   12
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1365
      Left            =   360
      MultiLine       =   -1  'True
      TabIndex        =   2
      Top             =   1800
      Width           =   4935
   End
   Begin VB.CommandButton cmdBxCancel 
      BackColor       =   &H00C0C0FF&
      Caption         =   "&Отмена"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2040
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   8760
      Width           =   1455
   End
   Begin VB.Frame fraInjuredPerson1 
      BackColor       =   &H00C0CFC5&
      Caption         =   "Принадлежность ВД"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808000&
      Height          =   2655
      Left            =   360
      TabIndex        =   13
      Top             =   5520
      Width           =   4935
      Begin VB.TextBox txtOwSurName 
         BackColor       =   &H00DDEADB&
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   12
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   406
         Left            =   1560
         TabIndex        =   5
         Top             =   480
         Width           =   3135
      End
      Begin VB.TextBox txtOwName 
         BackColor       =   &H00DDEADB&
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   12
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   406
         Left            =   1560
         TabIndex        =   6
         Top             =   960
         Width           =   3135
      End
      Begin VB.TextBox txtOwMidName 
         BackColor       =   &H00DDEADB&
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   12
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   406
         Left            =   1560
         TabIndex        =   7
         Top             =   1440
         Width           =   3135
      End
      Begin VB.TextBox txtOwBirthday 
         BackColor       =   &H00DDEADB&
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   12
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   406
         Left            =   1920
         TabIndex        =   8
         Top             =   2010
         Width           =   1335
      End
      Begin VB.ListBox lstOwSex 
         BackColor       =   &H00DDEADB&
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   12
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   3840
         TabIndex        =   9
         Top             =   2040
         Width           =   855
      End
      Begin VB.Label lbl6 
         BackColor       =   &H00C0CFC5&
         Caption         =   "Год рождения"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12.75
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   240
         TabIndex        =   18
         Top             =   2050
         Width           =   1575
      End
      Begin VB.Label lbl5 
         BackColor       =   &H00C0CFC5&
         Caption         =   "Отчество"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12.75
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   240
         TabIndex        =   17
         Top             =   1440
         Width           =   1335
      End
      Begin VB.Label lbl4 
         BackColor       =   &H00C0CFC5&
         Caption         =   "Имя"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12.75
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   240
         TabIndex        =   16
         Top             =   960
         Width           =   1335
      End
      Begin VB.Label lbl3 
         BackColor       =   &H00C0CFC5&
         Caption         =   "Фамилия"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12.75
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   240
         TabIndex        =   15
         Top             =   480
         Width           =   1335
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Центровка
         BackColor       =   &H00C0CFC5&
         Caption         =   "Пол"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12.75
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   3240
         TabIndex        =   14
         Top             =   2040
         Width           =   615
      End
   End
   Begin VB.Label lblboxEvSum 
      BackColor       =   &H00CFC2AC&
      Caption         =   "Количество ВД в упаковке:"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   9480
      TabIndex        =   27
      Top             =   240
      Width           =   3255
   End
   Begin VB.Label lblBoxCounter 
      BackColor       =   &H00CFC2AC&
      Caption         =   "Количество упаковок:"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   375
      Left            =   5640
      TabIndex        =   26
      Top             =   240
      Width           =   2775
   End
   Begin VB.Label Label3 
      BackColor       =   &H00CFC2AC&
      Caption         =   "Печать"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12.75
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   360
      TabIndex        =   25
      Top             =   4080
      Width           =   1095
   End
   Begin VB.Label Label2 
      BackColor       =   &H00CFC2AC&
      Caption         =   "Упаковка"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12.75
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   360
      TabIndex        =   24
      Top             =   3415
      Width           =   1095
   End
   Begin VB.Label lbl7 
      BackColor       =   &H00CFC2AC&
      Caption         =   "Дата предоставления ВД"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12.75
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   360
      TabIndex        =   23
      Top             =   895
      Width           =   2895
   End
   Begin VB.Label lbl1 
      BackColor       =   &H00CFC2AC&
      Caption         =   "Дата изъятия ВД:"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12.75
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   360
      TabIndex        =   22
      Top             =   390
      Width           =   2055
   End
   Begin VB.Label lblAllEvSumCounter 
      BackColor       =   &H00CFC2AC&
      Caption         =   "Общая сумма ВД:"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004000&
      Height          =   375
      Left            =   12960
      TabIndex        =   21
      Top             =   240
      Width           =   2295
   End
   Begin VB.Label lbl2 
      BackColor       =   &H00CFC2AC&
      Caption         =   "Место изъятия ВД:"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12.75
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   405
      Left            =   360
      TabIndex        =   19
      Top             =   1440
      Width           =   2295
   End
End
Attribute VB_Name = "frmEvidences"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Форма "frmEvidences"
'Ввод данных "вещественные доказательства"
'Дата создания: 01.06.2016
'@author Andr.Nab.n@gmail.com
Option Explicit
'Collection (список названия коробки и название ВД  в ней)
'Public colEvidences As New Collection 'коллекция "ВД" (в упаковке)
'Fields
'Private mvarstrBoxKey As String 'ключ для упаковок с ВД ="BX****"
'Private mvarstrEvKey As String  'ключ для вещдоков (ВД)= "EV****"
'Private mvarstrKey As String    'общий ключ = "BX****" + "EV****" ("BX****EV****")
'Class
Public newEvOwner As clmEntrants 'Экземпляр класса "Владелец ВД"
Private newEvDate As clmCaseDate 'Экземпляр класса Даты
Private tmpevKey As String
'Errors:
Const Msg As String = "Ошибка ввода данных!"
'Colors
Private Enum frmColor
    LIGHT_YELLOW = &HC0FFFF 'RGB(102, 102, 153)
    CITRIC = &HFFFF& 'желтый
    RED = &HFF&
    Black = &H0&
    LIGHT_GREEN = &HC0FFC0  'Салатовый RGB(200, 256, 200)
    LIGHT_OLIVE = &HDDEADB 'оливковый светлый
    LIGHT_BLUE = &HFFE8DF   '&HFEF7F1
    PURPLE = &HFFC0C0
    BROWN = &H80FF&
End Enum
'Cancel
Dim Cancel As Integer
'
'=================== C O N S T R U C T O R ===================
'
Private Sub Form_Initialize()
'инициализация формы frmEvidences
     With Me
        .Caption = "Регистрация ВД" 'frmNewEF.newEF.categories
        .Width = 5580
        .cmdBxClear.Visible = False
    End With
    Call EvClass_Initialize
End Sub
'
Private Sub Form_Load()
'процедура загрузки формы
Move (Screen.Width - Width) / 2, (Screen.Height - Height) / 3.7 '/ 17.7
End Sub
'
Private Sub EvClass_Initialize()
'Процедура инициализации новых экземпляров классов "Evidences", "Owner":
    Set newEvOwner = New clmEntrants        'Экземпляр класса "Владелец ВД"
    Set newEvDate = New clmCaseDate         'Экземпляр класса Даты
    Call frmNewEF.addNewBox
End Sub
'
Private Sub EvClass_Terminate()
'Процедура уничтожения новых экземпляров классов "Evidences", "Owner":
    Set newEvOwner = Nothing   'Уничтожение экземпляра класса "Владелец ВД"
    Set newEvDate = Nothing 'Экземпляр класса Даты
End Sub
'
Public Sub Form_Terminate()
    Call EvClass_Terminate
End Sub


'================== I C A P S U L A T I O N =============================
'
'
'==================== К Н О П К И ======================================
'
Private Sub cmdEvClear_Click()
'Процедура нажатия кнопки "очистить"
'1) уничтожение экземпляров классов;
Call EvClass_Terminate
'2)Очистка полей формы:
    Dim x As Object
        For Each x In Me.Controls
            If TypeName(x) = "TextBox" Then
                x.Text = ""
            ElseIf TypeName(x) = "ComboBox" Then
                x.Clear
            End If
        Next x
Call Ev_Enabled
txtBxOutTake.SetFocus
End Sub
'
Private Sub Ev_Enabled()
'Процедура включения полей "Заключение эксперта"
    Dim x As Object
        For Each x In Me.Controls
            If TypeName(x) = "TextBox" Or TypeName(x) = "ComboBox" Then
                x.BackColor = &HF1F9FA
            End If
        Next x
    txtBxFirstDate.SetFocus
End Sub
''
Private Sub cmdEvCancel_Click()
'Процедура нажатия кнопки "Отмена"
MsgBox "Ввод завершен!", vbOKOnly   'вызов окна сообщения:
'    Set newEvDate = Nothing             'Уничтожение объекта "NewEvid"
    Set newEvOwner = Nothing            'Уничтожение объекта "newEvOwner"
'    Set mdCount.boxCounter = Nothing    'Уничтожение объекта счетчика "Количесво упаковок с вещ.доками"
'frmDocList.Show 'вывод следующей формы
'    Call Open_DocListForm ??????? - перенести в другое место.
    Call EvClass_Terminate
Me.Hide 'закрытие формы "Evidences"
End Sub
'
Private Sub cmdSaveOK_Click()
'Процедура нажатия кнопки "Сохранить"
'1)'разворачивание формы
    Me.Width = 15675
    With frmNewEF
        Dim str As String
'предыдущая версия ++++++++++++++++++++++++
'        .newBox.strBxEntrants = create_strOwner
'         str = "Упаковка с объектами " & .newBox.strBxEntrants & ":" 'Формирование значения владелец ВД:
'+++++++++++++++++++++++++++++++++++++++++
        Dim str1 As String
        str1 = create_strOwner(2, True)
        Me.Caption = str1
'        lstEvList.AddItem str1  'Добавление записи в лист формы frmEvidences:
        Me.lblBoxCounter.Caption = " Данные упаковки №" & .colBoxes.Count + 1
    End With
    Call InputEvid_MsBxShow 'Вызов диалогового окна создания новой записи объектов
End Sub
'
Private Static Sub InputEvid_MsBxShow()
'вызов диалогового окна "Добавить ВД?"
    If MsgBox("Добавть новый вещ.док. в список?", vbYesNo) = vbYes Then
        Call InputEvid
    Else:
        With frmNewEF
            .newBox.strBxEntrants = create_strOwner(.newBox.colEvidences.Count, True)
            .colBoxes.Add .newBox, .strBoxKey 'сохранение коробки с ВД  в коллекции коробок
            Call .show_MsgCreateNewBox 'вызов диалогового окна "Создать новую коробку с ВД"
'           Me.Hide
            Unload Me
        End With
    End If
End Sub
'
Private Static Sub InputEvid()
'Процедура создания списка ВД
'1) формирование ключа для ВД
    With frmNewEF
        .strEvKey = .newBox.colEvidences.Count + 1
'        Debug.Print "ключ для вещдоков = " & .strEvKey
    End With
'NB!!!! ПОЛНЫЙ КЛЮЧ strKey ФОРМИРУЕТСЯ В СВОЙСТВЕ Get_strKey
'3) ввод названия ВД
    With frmNewEF.newBox
        Dim strEvName As String
            strEvName = CStr(InputBox("Введите название вещественного доказательства", "Запись нового объекта", , 4000, 4300))
     '   Проверка на пустое значение
        If strEvName = "" Then '
            Call InputEvid_MsBxShow
        Else
            .colEvidences.Add strEvName, frmNewEF.strKey 'добавили ВД в коллекцию ВД frmNewEF.newBox
        End If
'4) работа со счетчиками и надписями на форме
        Me.lblboxEvSum.Caption = "Количество ВД в упаковке: " & .colEvidences.Count
        If chbNoMedCrim.Value = 0 Then
            frmNewEF.allEvSumCounter.increment 'увеличение счетчика общей суммы ВД
            Me.lblAllEvSumCounter.Caption = "Общая сумма ВД: " & frmNewEF.allEvSumCounter.getTale
        End If
'5)  оВД в окне формы:
        lstEvList.AddItem .colEvidences.Count & ". " & frmNewEF.newBox.colEvidences(frmNewEF.strKey)
    End With
 Call InputEvid_MsBxShow
End Sub
'
Private Static Function create_strOwner(Optional cnt As Long = 1, Optional version As Boolean = False) As String
'функция создания строки "принадлежность ВД"
'cnt - count количество ВД в коробке
'version - различние версии печати (False  - версия 1, True - версия 2)
Dim strTmp1 As String, strTmp2 As String, strTmp3 As String
    If newEvOwner.surName <> "" Then
        strTmp2 = newEvOwner.create_InitialslName  '" " &
    Else: strTmp2 = "(принадлежность не указана)"
    End If
    
    If cnt = 1 Then
        strTmp1 = "объектом "
    Else: strTmp1 = "объектами "
    End If
    
    If frmNewEF.newBox.strBxPlace <> "" Then
        If cnt = 1 Then
            If version = False Then
                strTmp3 = ", изъятый " & frmNewEF.newBox.strBxPlace
            Else: strTmp3 = ", изъятым " & frmNewEF.newBox.strBxPlace
            End If
        ElseIf cnt > 1 Then
            If version = False Then
                strTmp3 = ", изъятые " & frmNewEF.newBox.strBxPlace
            Else: strTmp3 = ", изъятыми " & frmNewEF.newBox.strBxPlace
            End If
        End If
    Else: strTmp3 = ", (место изъятия не указано)"
    End If
create_strOwner = frmNewEF.newBox.strBxPackage & " с " & strTmp1 & strTmp2 & strTmp3   '"Объекты" & strTmp1 & strTmp2
End Function
'
'==================== П О Л Я  Ф О Р М Ы ===============================
'
Private Sub txtEnter(tmpObj As Object)
'изменение текстоых полей при получении фокуса
    With tmpObj
        .Text = ""
        .BackColor = frmColor.LIGHT_YELLOW 'RGB(102, 102, 153) 'frmColor.Цвет1
        .ForeColor = frmColor.Black
    End With
End Sub
'
 Private Sub txt_Exit(tmpObj As Object)
'изменение текстоых полей при потере фокуса
    With tmpObj
        .BackColor = frmColor.BROWN
        If .name = "txtOwBirthday" Or _
                .name = "lstOwSex" _
            Then
                .Text = "не указан"
        ElseIf .name = "txtBxOutTake" _
                Or .name = "txtBxFirstDate" _
                Or .name = "cmbBxPackage" _
                Or .name = "txtBxStamp" _
                Or .name = "txtOwSurName" _
            Then
                .Text = "не указана"
        ElseIf .name = "txtBxPlace" _
                Or .name = "txtOwName" _
                Or .name = "txtOwMidName" _
            Then
                .Text = "не указано"
        Else
            .Text = "не указаны"
        End If
   End With
 End Sub
'
Private Sub txtBxOutTake_GotFocus()
'Дата изъятия ВД:
    Call txtEnter(txtBxOutTake)
   'Дата изъятия ВД = Дата вынесения постановления
    If frmNewEF.newEF.rulingDate <> 0 Then
         txtBxOutTake.Text = frmNewEF.newEF.rulingDate
    End If
End Sub
'
Private Sub txtBxOutTake_LostFocus()
'Дата изъятия ВД:
 Dim str As String, dt As Date
 With txtBxOutTake
    If .Text = "" Then
        Call txt_Exit(txtBxOutTake)
    Else
        With newEvDate
            dt = .validateDate(.ExamDate(txtBxOutTake.Text)) 'проверка введенной даты на валидность
            frmNewEF.newBox.DtmBxOutTake = dt
            str = .dateToString(dt)
        End With
        .BackColor = frmColor.LIGHT_GREEN 'RGB(200, 256, 200) 'frmColor.Салатовый
        .Text = str
    End If
End With
'Debug.Print "Дата изъятия ВД: " & frmNewEF.newBox.DtmBxOutTake
End Sub
'
Private Sub txtBxFirstDate_GotFocus()
'Дата предоставления ВД
      Call txtEnter(txtBxFirstDate)
      txtBxFirstDate.Text = frmNewEF.newBox.DtmBxOutTake + 1
End Sub
'
Private Sub txtBxFirstDate_LostFocus()
'Дата предоставления ВД
 Dim tmp As Double
 Dim dt As Date
 With txtBxFirstDate
    If .Text = "" Then
        Call txt_Exit(txtBxFirstDate)
    Else
        With newEvDate
            dt = .validateDate(.ExamDate(txtBxFirstDate.Text))
'            'сравнение дат: 'Дата предоставления ВД >= Дата изъятия ВД
            Do
                tmp = .compareDt(frmNewEF.newBox.DtmBxOutTake, dt)
                If tmp < 0 Then
                    MsgBox "Дата изъятия ВД больше даты предоставления ВД!", _
                            vbCritical, Msg
                    dt = .ExamDate(InputBox("Правильно введите дату предоставления ВД!", _
                                        "Ввод даты", frmNewEF.newBox.DtmBxOutTake + 1))
                Else
                    frmNewEF.newBox.DtmBxFirstDate = dt
                    Exit Do
                End If
            Loop
        End With
            With txtBxFirstDate
                .BackColor = RGB(200, 256, 200) 'frmColor.Салатовый
                .Text = newEvDate.dateToString(dt)
            End With
    End If
End With
Debug.Print "Дата предоставления ВД: " & frmNewEF.newBox.DtmBxFirstDate
End Sub
'
Private Sub txtBxPlace_GotFocus()
'Место изъятия ВД:
     Call txtEnter(txtBxPlace)
End Sub
'
Private Sub txtBxPlace_LostFocus()
'Место изъятия ВД:
      With txtBxPlace
        If .Text = "" Then
           Call txt_Exit(txtBxPlace)
        Else
            frmNewEF.newBox.strBxPlace = .Text
            .BackColor = frmColor.LIGHT_GREEN 'RGB(200, 256, 200) 'frmColor.Салатовый
        End If
    End With
Debug.Print "Место изъятия ВД: " & frmNewEF.newBox.strBxPlace
End Sub
'
Private Sub cmbBxPackage_GotFocus()
     Call txtEnter(cmbBxPackage)
 With cmbBxPackage
        .Clear
        .Text = "Картонная коробка"
        .AddItem "Картонная коробка"
        .AddItem "Пластиковый пакет"
        .AddItem "Спецпакет"
        .AddItem "Бумажный конверт"
        .AddItem "Бумажный сверток"
        .AddItem "Бумажный пакет"
    End With
End Sub
'
Private Sub cmbBxPackage_LostFocus()
'Упаковка
    With cmbBxPackage
        If .Text = "" Then
           Call txt_Exit(cmbBxPackage)
        Else
'            frmNewEF.colBoxes.Item(frmNewEF.strBoxKey).strBxPackage = .Text
        frmNewEF.newBox.strBxPackage = .Text
            .BackColor = frmColor.LIGHT_GREEN
        End If
    End With
Debug.Print "Упаковка - " & frmNewEF.newBox.strBxPackage
End Sub
'
Private Sub txtBxStamp_GotFocus()
'Печать
    Call txtEnter(txtBxStamp)
    With txtBxStamp
        .Text = "Следственного комитета Республики Беларусь " & "Для пакетов"
    End With
End Sub
'
Private Sub txtBxStamp_LostFocus()
'Печать
    With txtBxStamp
        If .Text = "" Then
           Call txt_Exit(txtBxStamp)
        Else
'            frmNewEF.colBoxes.Item(frmNewEF.strBoxKey).strBxStamp = .Text
        frmNewEF.newBox.strBxStamp = Chr(171) & .Text & Chr(187)
            .BackColor = frmColor.LIGHT_GREEN
        End If
    End With
Debug.Print "Печать - " & frmNewEF.newBox.strBxStamp
End Sub
'
Private Sub txtOwSurName_GotFocus()
'Фамилия владельца ВД (Owner)
    Call txtEnter(txtOwSurName)
'    Dim str As String
'    str = frmNewEF.newInjPr.surName
'        With txtOwSurName
'            If str <> "" Then
'                .Text = str
'            End If
'        End With
End Sub
'
Private Sub txtOwSurName_LostFocus()
'Фамилия владельца ВД (InjPrSurName)
    With txtOwSurName
        If .Text = "" Then
             Call txt_Exit(txtOwSurName)
        Else: newEvOwner.surName = StrConv(.Text, vbProperCase)
            .BackColor = RGB(200, 256, 200) 'frmColor.Салатовый
            .Text = newEvOwner.surName
        End If
    End With
Debug.Print "Фамилия владельца ВД: " & newEvOwner.surName
End Sub
'
Private Sub txtOwName_GotFocus()
'Имя владельца ВД (InjPrName)
    Call txtEnter(txtOwName)
'    Dim str As String
'    str = frmNewEF.newInjPr.name
'        With txtOwName
'            If str <> "" Then
'                .Text = str
'            End If
'        End With
End Sub
'
Private Sub txtOwName_LostFocus()
'Имя владельца ВД (InjPrName)
      With txtOwName
        If .Text = "" Then
             Call txt_Exit(txtOwName)
        Else: newEvOwner.name = StrConv(.Text, vbProperCase)
            .BackColor = RGB(200, 256, 200) 'frmColor.Салатовый
            .Text = newEvOwner.name
        End If
    End With
Debug.Print "Имя владельца ВД: " & newEvOwner.name
End Sub
''
Private Sub txtOwMidName_GotFocus()
'Отчество владельца ВД (Owner)
    Call txtEnter(txtOwMidName)
'    Dim str As String
'    str = frmNewEF.newInjPr.midName
'        With txtOwMidName
'            If str <> "" Then
'                .Text = str
'            End If
'        End With
End Sub
'
Private Sub txtOwMidName_LostFocus()
'Отчество владельца ВД (InjPrMidName)
     With txtOwMidName
        If .Text = "" Then
             Call txt_Exit(txtOwMidName)
        Else: newEvOwner.midName = StrConv(.Text, vbProperCase)
            .BackColor = RGB(200, 256, 200) 'frmColor.Салатовый
            .Text = newEvOwner.midName
        End If
    End With
Debug.Print "Отчество владельца ВД: " & newEvOwner.midName
End Sub
'
Private Sub txtOwBirthday_GotFocus()
'Дата рождения владельца ВД (InjPrBirthday):
    Call txtEnter(txtOwBirthday)
'    Dim str As String
'    str = frmNewEF.newInjPr.birthday
'        With txtOwBirthday
'            If str <> "" Then
'                .Text = str
'            End If
'        End With
End Sub
'
Private Sub txtOwBirthday_LostFocus()
'Дата рождения владельца ВД (InjPrBirthday):
    With txtOwBirthday
        If .Text = "" Then
             Call txt_Exit(txtOwBirthday)
        Else: newEvOwner.birthday = StrConv(.Text, vbProperCase)
            .BackColor = RGB(200, 256, 200) 'frmColor.Салатовый
            .Text = newEvOwner.birthday
        End If
    End With
Debug.Print "Дата рождения владельца ВД: " & newEvOwner.birthday
End Sub
'
Private Sub lstOwSex_GotFocus()
'Пол владельца ВД
    Dim str As String
    str = frmNewEF.newInjPr.sex
         With lstOwSex
            .Clear
            .AddItem "муж."
            .AddItem "жен."
            .AddItem "пол не указан"
            .BackColor = &HC0FFFF
            If str <> "" Then
                .Text = str
            End If
        End With
End Sub
'
Private Sub lstOwSex_LostFocus()
'Пол владельца ВД
    With lstOwSex
        newEvOwner.sex = .Text
        .BackColor = RGB(200, 256, 200)
    End With
Debug.Print "Пол владельца ВД: " & newEvOwner.sex
End Sub
'




'================= L I B ==========================.
'Private Sub cmbBxCategory_GotFocus()
''категория ВД
' Call txtEnter(cmbBxCategory)
' With cmbBxCategory
'        .Clear
'        .Text = "судебной медико-криминалистической"
''   работа с массивом названия экспертиз
'    Dim i As Long
'    If UBound(mdPrintDoc.arrEF) <> 0 Then
'            For i = LBound(mdPrintDoc.arrEF) To UBound(mdPrintDoc.arrEF)
'                .AddItem mdPrintDoc.arrEF(i)
'            Next i
'        End If
'    End With
'End Sub
''Private Sub cmbBxCategory_LostFocus()
''категория ВД
'  With cmbBxCategory
'        If .Text = "" Then
'           Call txt_Exit(cmbBxCategory)
'        Else
'            frmNewEF.newBox.strBxCategory = .Text
''            frmNewEF.colBoxes.Item(frmNewEF.strBoxKey).strBxCategory = .Text
'            .BackColor = frmColor.LIGHT_GREEN
'        End If
'    End With
'Debug.Print "cmbBxCategory - " & frmNewEF.newBox.strBxCategory
'End Sub

'


'=================== C O L L E C T I O N =================================
'1) работа с коллекцией
'
'Function keyExists(coll As Collection, _
'                key As String) As Boolean
''проверяет,существует ли ключ.
'On Error GoTo EH
'    coll.Item key
'    keyExistsExists = True
'EH: '"key is not exists"
'End Function
''
'Sub QuickSort(coll As Collection, _
'                first As Long, _
'                last As Long)
''Быстрая соотироака QuickSort
'Dim vCentreVal As Variant
'Dim vTemp As Variant
'Dim lTempLow As Long
'Dim lTempHi As Long
'    lTempLow = first
'    lTempHi = last
'    vCentreVal = coll((first + last) \ 2)
'    Do While lTempLow <= lTempHi
'
'        Do While coll(lTempLow) < vCentreVal And lTempLow < last
'            lTempLow = lTempLow + 1
'        Loop
'
'        Do While vCentreVal < coll(lTempHi) And lTempHi > first
'            lTempHi = lTempHi - 1
'        Loop
'
'        If lTempLow <= lTempHi Then
'    'Поменять значения
'            vTemp = coll(lTempLow)
'            coll.Add coll(lTempHi), After:=lTempLow
'            coll.Remove lTempLow
'            coll.Add vTemp, Before:=lTempHi
'            coll.Remove lTempHi + 1
'    'Перейти к следующим позициям
'            TempLow = lTempLow + 1
'            lTempHi = lTempHi - 1
'        End If
'    Loop
'    If first < lTempHi Then
'            QuickSort coll, first, lTempHi
'            If lTempLow < last Then
'                QuickSort coll, lTempLow, last
'            End If
'    End If
'End Sub
'

'Private Sub Form_Initialize()
''инициализация формы
'     With Me
'        .Caption = "Регистрация ВД" 'frmNewEF.newEF.categories
'        .Width = 5580
'        .cmdBxClear.Visible = False
'    End With
'    Call EvClass_Initialize
'    Call zeroControlStatus
'End Sub
'
'Private Sub zeroControlStatus()
''процедура возвращения полей формы в искходное состояние
'    With Me
'        Dim X As Object
'            For Each X In .Controls
'                If TypeName(X) = "TextBox" Then
'                    X.BackColor = frmColor.LIGHT_BLUE '&HFFE8DF   '&HF1F9FA
'                    X.Text = ""
'                        If X.name = "txtOwSurName" Or _
'                            X.name = "txtOwName" Or _
'                            X.name = "txtOwMidName" Or _
'                            X.name = "txtOwBirthday" _
'                            Then
'                                X.BackColor = frmColor.LIGHT_OLIVE
'                        End If
'                ElseIf TypeName(X) = "ListBox" Then
'                    X.BackColor = frmColor.LIGHT_OLIVE 'If TypeName(X) = "ListBox" Then
'                ElseIf TypeName(X) = "ComboBox" Then
'                    X.Clear
'                    X.BackColor = frmColor.LIGHT_BLUE
'                End If
'            Next X
'    End With
'End Sub
'
'Private Sub EvClass_Initialize()
''Процедура инициализации новых экземпляров классов "Evidences", "Owner":
'    Set newEvOwner = New clmEntrants        'Экземпляр класса "Владелец ВД"
'    Set newEvDate = New clmCaseDate         'Экземпляр класса Даты
'    Set frmNewEF.newBox = New clmEvBox 'Экземпляр класса "упаковка с ВД"
''   Счетчики:
''    Set mdCount.boxEvSumCounter = New clmCounter 'Экземпляр класса счетчика "сумма вещ.доков (ВД) в упаковке"
''        With mdCount.boxEvSumCounter
''            .name = "Количесво вещ.доков в упаковке "
''            Debug.Print .name & Chr(32) & .getTale
''        End With
''очистка полей формы (необходима при многократных заполнениях формы)
'Call zeroControlStatus
'End Sub
'
'=====================М Е Т О Д Ы  Ф О Р М Ы: ===========================
'Private Sub cmdBxCancel_Click()
''Процедура нажатия кнопки "Отмена"
'MsgBox "Ввод завершен!", vbOKOnly 'вызов окна сообщения:
'Set newBxid = Nothing 'Уничтожение объекта "NewEvid"
'Set newEvOwner = Nothing 'Уничтожение объекта "newEvOwner"
''frmDocList.Show 'вывод следующей формы
'Call Open_DocListForm
'Me.Hide 'закрытие формы "Evidences"
'End Sub
''
'Private Sub cmdEvClear_Click()
''Процедура нажатия кнопки "очистить"
''1) уничтожение экземпляров классов;
'Call EvClass_Terminate
''2)Очистка полей формы:
'    Dim X As Object
'        For Each X In Me.Controls
'            If TypeName(X) = "TextBox" Then
'                X.Text = ""
'            ElseIf TypeName(X) = "ComboBox" Then
'                X.Clear
'            End If
'        Next X
'Call Ev_Enabled
'txtEvOutTake.SetFocus
'End Sub
'

'Private Static Sub addNewEvBox_MsgShow()
''вызов диалогового окна "Создать новую группу ВД"
'    If MsgBox("Создать новую группу Объектов (ВД)?", vbYesNo) = vbYes Then
'        Call cmdEvClear_Click
'        Call EvClass_Initialize
'    Else:
''окрытие формы frmDocList
'        Dim tmGr As String
'            tmGr = CStr(Format(Me.lblEvidCount, "#000")) & "000" & "0000"
'            frmNewEF.colEvidences.Add Me.lblEvGRSum, "GrCount" & tmGr
'        MsgBox "Ввод объектов завершен!", vbOKOnly
'        Call EvClass_Terminate
'        Call Open_DocListForm
'        Me.Hide
'    End If
'''Debug.Print "Количство групп ВД= " & frmNewEF.EvGrCount
'End Sub

' Private Sub changeBoxCounter()
'   With mdCount.boxCounter
'        .increment 'увеличение (+1) значения счетчика
'            Debug.Print .name & Chr(32) & .getTale
''       изменение надписи "lblBoxCounter" на форме
'        lblBoxCounter.Caption = .name & " - " & .getTale
''       создание значения ключа для упаковок
'        strBoxKey = CStr(Format(.getTale, "#0000"))
'            Debug.Print "ключ для упаковок = " & strBoxKey
''       запись данных в коллекцию
'        Dim tmp As String
'            tmp = Create_strOwner
'        frmNewEF.colBoxes.Add tmp, strBoxKey
'Debug.Print "Название упаковки - " & frmNewEF.colBoxes.Item(strBoxKey)
'Debug.Print "Значение ключа = " & strBoxKey
''       Добавление записи в лист формы frmEvidences:
'        lstEvList.AddItem tmp
'    End With
' End Sub

'Private Static Sub InputEvid()
''Процедура создания списка ВД
''1) ввод названия ВД
'    With frmNewEF.newBox
'        .strBxName = CStr(InputBox("Введите название вещественного доказательства", "Запись нового объекта", , 4000, 4300))
'     '   Проверка на пустое значение
'        If .strBxName = "" Then '
'           Call InputEvid_MsBxShow
'        Else
''       1)счетчик "сумма ВД" в упаковке
'            With mdCount.boxEvSumCounter
'                .increment 'увеличение (+1) значения счетчика "сумма ВД" в упаковке
'                    Debug.Print .name & " = " & .getTale
''               создание значения ключа для вещдоков
'                strEvKey = CStr(Format(.getTale, "#0000"))
'                    Debug.Print "ключ для вещдоков = " & strEvKey
''               изменение надписи "lblBoxCounter.Caption" на форме
'                lblboxEvSum.Caption = .name & " - " & .getTale
'            End With
''        2)счетчик "Общая сумма ВД"
'            With mdCount.allEvSumCounter
'                .increment 'увеличение (+1) значения счетчика "сумма ВД" в упаковке
'                   Debug.Print .name & Chr(32) & .getTale
''               изменение надписи "lblAllEvSumCounter" на форме
'            lblAllEvSumCounter.Caption = .name & " - " & .getTale
'            End With
''        3)формирование полного ключа:
'            Dim tmpEv As String
'            tmpEv = strBoxKey & strEvKey
'                Debug.Print "полный ключ для вещдоков = " & tmpEv
''        4)запись ВД в коллекцию:
'            frmNewEF.newBox.colEvidences.Add .strBxName, tmpEv
'                Debug.Print "запись ВД в коллекции -" & frmNewEF.newBox.colEvidences.Item(tmpEv) & Chr(10) _
'                & "Количесво объектов в коллекции - " & frmNewEF.newBox.colEvidences.Count
''           отображение ВД в окне формы:
''        5)запись ВД в общую коллекцию Box & ВД
'             frmNewEF.colBoxes.Add .strBxName, tmpEv
'             Debug.Print "Содержание коробки = " & frmNewEF.colBoxes.Item(tmpEv)
'             Debug.Print "Значение ключа " & tmpEv
'            lstEvList.AddItem (mdCount.boxEvSumCounter.getTale) & ". " & frmNewEF.newBox.colEvidences(tmpEv)
'        End If
'    End With
' Call InputEvid_MsBxShow
'End Sub

'Объявление нового экземпляра класса "ВД":
'Private NewEvid As clmEvidences
'Private newEvOwner As Entrants
''
'Public Property Let fEvCount(ByVal vData As Integer)
''Счетчик экземпляров форм "ВД"
'mvarfEvCount = vData
'End Property
''
'Public Property Get fEvCount() As Integer
''Счетчик экземпляров форм "ВД"
'fEvCount = mvarfEvCount
''Debug.Print "Счетчик экземпляров форм ВД = ", fEvCount
'End Property
''

''
''
'Private Sub lstOwSex_LostFocus()
''Пол потерпевшего
'    With lstOwSex
'        newEvOwner.En_Sex = .Text
'        .BackColor = RGB(200, 256, 200)
'    End With
'End Sub
''

''
'Private Sub lstOwSex_GotFocus()
''Пол потерпевшего
'    With lstOwSex
'        .Clear
'        .AddItem "муж."
'        .AddItem "жен."
'        .AddItem "пол не указан"
'        .BackColor = &HC0FFFF
'    End With
'End Sub


'Private Sub txtEvOutTake_GotFocus()
''Дата изъятия ВД (EvOutTake)
'    With txtEvOutTake
'        .Text = ""
'        .BackColor = &HC0FFFF
'        .ForeColor = &H80000008
'    End With
'End Sub
''
'Private Sub txtEvOutTake_LostFocus()
''Дата изъятия ВД (EvOutTake)
'        With txtEvOutTake
'            If .Text = "" Then
'                .Text = "не указана"
'                .BackColor = &HC0E0FF
'            Else:
'                Do
'                    If Not IsDate(.Text) Then
'                        Beep
'                        .BackColor = RGB(256, 0, 0)
'                        MsgBox "Неправильрно введена дата!", vbCritical, "Ошибка ввода!"
'                        .Text = InputBox("Введите правильно дату!", _
'                            "Ввод корректной даты")
'                    Else: NewEvid.DtmEvOutTake = CDate(.Text)
'                        .BackColor = RGB(200, 256, 200)
'                Exit Do
'                    End If
'                Loop
'            End If
'    End With
'    cmdEvClear.Visible = True
''указание категории ВД
''NewEvid.strEvCategory = Me.Caption
'End Sub
''
'Private Sub txtEvPlace_GotFocus()
''Место изъятия ВД (EvPlace)
'    With txtEvPlace
'        .Text = ""
'        .BackColor = &HC0FFFF
'        .ForeColor = &H80000008
'    End With
'End Sub
''
'Private Sub txtEvPlace_LostFocus()
''Место изъятия ВД (EvPlace)
'    With txtEvPlace
'        If Len(.Text) = 0 Then
'            .Text = "(место изъятия объектов не указано)"
'            .BackColor = &HC0E0FF
'        Else: NewEvid.strEvPlace = .Text
'            .BackColor = RGB(200, 256, 200)
'        End If
'    End With
'End Sub

''
'Private Sub txtOwName_GotFocus()
''Имя владельца ВД (InjPrName)
'    With txtOwName
'        .BackColor = &HC0FFFF
'        .ForeColor = &H80000008
'        If .Text = "(имя не указано)" Then
'            .Text = ""
'        End If
'    End With
'End Sub
''
'Private Sub txtOwName_LostFocus()
''Имя владельца ВД (InjPrName)
'    With txtOwName
'        If Len(.Text) = 0 Then
'            .Text = "(имя не указано)"
'            .BackColor = &HC0E0FF
'        Else
'            If Len(.Text) = 1 Then
'                newEvOwner.En_Name = StrConv(.Text, vbProperCase) & "."
'            Else: newEvOwner.En_Name = StrConv(.Text, vbProperCase)
'            End If
'        .Text = newEvOwner.En_Name
'        .BackColor = RGB(200, 256, 200)
'        End If
'    End With
'End Sub
''
'Private Sub txtOwMidName_GotFocus()
''Отчество владельца ВД (Owner)
'    With txtOwMidName
'        .BackColor = &HC0FFFF
'        .ForeColor = &H80000008
'        .BackColor = &HC0FFFF
'        If .Text = "(отчество не указано)" Then
'            .Text = ""
'        End If
'    End With
'End Sub
''
'Private Sub txtOwMidName_LostFocus()
''Отчество владельца ВД (InjPrMidName)
'    With txtOwMidName
'        If Len(.Text) = 0 Then
'            .Text = "(отчество не указано)"
'            .BackColor = &HC0E0FF
'        Else
'            If Len(.Text) = 1 Then
'                newEvOwner.En_MidName = StrConv(.Text, vbProperCase) & "."
'            Else: newEvOwner.En_MidName = StrConv(.Text, vbProperCase)
'            End If
'        .Text = newEvOwner.En_MidName
'        .BackColor = RGB(200, 256, 200)
'        End If
'    End With
'End Sub
''
'Private Sub txtOwBirthday_GotFocus()
''Дата рождения владельца ВД (InjPrBirthday):
'    With txtOwBirthday
'        .BackColor = &HC0FFFF
'        .ForeColor = &H80000008
'    End With
'End Sub
''
'Private Sub txtOwBirthday_LostFocus()
''Дата рождения владельца ВД (InjPrBirthday):
'    With txtOwBirthday
'        If .Text = "" Then
'            .Text = "не указан"
'            .BackColor = &HC0E0FF
'        Else: newEvOwner.En_Birthday = .Text
'            .BackColor = RGB(200, 256, 200)
'        End If
'    End With
'End Sub
''
'Private Sub txtEvFirstDate_GotFocus()
''Дата предоставления ВД
'    With txtEvFirstDate
'        .BackColor = &HC0FFFF
'        .ForeColor = &H80000008
'    End With
'End Sub
''
'Private Sub txtEvFirstDate_LostFocus()
''Дата предоставления ВД
'    With txtEvFirstDate
'            If .Text = "" Then
'                .Text = "не указана"
'                .BackColor = &HC0E0FF
'            Else
'                Do
'                    If Not IsDate(.Text) Then  'Or (CDate(.Text) - NewEvid.DtmEvOutTake) < 0
'                        Beep
'                        .BackColor = RGB(256, 0, 0)
'                        MsgBox "Неправильно введена дата!", vbCritical, "Ошибка ввода!"
'                        .Text = InputBox("Введите правильно дату!", _
'                            "Ввод корректной даты")
'                    Else: NewEvid.DtmEvFirstDate = CDate(.Text)
'                        NewEvid.blnEvProvide = True 'предоставили = "Да"
'                        .BackColor = RGB(200, 256, 200)
'                Exit Do
'                    End If
'                Loop
'            End If
'    End With
'End Sub
''
'Private Static Sub addNewEvBox_MsgShow()
''вызов диалогового окна "Создать новую группу ВД"
'    If MsgBox("Создать новую группу Объектов (ВД)?", vbYesNo) = vbYes Then
'        Call cmdEvClear_Click
'        Call EvClass_Initialize
'    Else:
''окрытие формы frmDocList
'        Dim tmGr As String
'            tmGr = CStr(Format(Me.lblEvidCount, "#000")) & "000" & "0000"
'            frmNewEF.colEvidences.Add Me.lblEvGRSum, "GrCount" & tmGr
'        MsgBox "Ввод объектов завершен!", vbOKOnly
'        Call EvClass_Terminate
'        Call Open_DocListForm
'        Me.Hide
'    End If
'''Debug.Print "Количство групп ВД= " & frmNewEF.EvGrCount
'End Sub
''
'Private Sub Open_DocListForm() 'Процедура открытия формы "Перечень документов"
'    mdCount.fDocListCount = mdCount.fDocListCount + 1
'    Dim tmp As String 'Временная переменная-счетчик
'        tmp = Me.lblEvidCount 'получения номера-связки формы активной формы frmEvidences
'    Dim frmD As frmDocList 'обявление новой формы
'    Set frmD = New frmDocList
''добавление формы в список созданных форм
'        frmD.Caption = mdPrintDoc.DocCat & Chr(32) & tmp
'        frmD.lblDocListCount.Caption = tmp
'        mdCount.colForms.Add frmD, "DocListform" & tmp
'    frmD.Show
'    mdCount.colForms("Evidform" & tmp).Hide
'End Sub
''

'Private Sub Create_colOwner()
''Процедура записи даннбых полей формы в коллекцию ВД
''1)изменение счетчиков:
'newEvOwner.EnCount = newEvOwner.EnCount + 1
'Me.lblEvGRSum = EvGRCount
''формирование строки-ключа: 000_000_0000(СчетчикФормы_СчетчикГруппВД_СчетчикВД)
'    Dim tmp As String
'        tmp = CStr(Format(Me.lblEvidCount, "#000")) & CStr(Format(mdCount.EvGRCount, "#000")) & "0000"
'    With newEvOwner '"владелец"
'        frmNewEF.colEvidences.Add .Create_InitialsFullName, "newEvOwnerFullName" & tmp 'ФИО владельца ВД
'        frmNewEF.colEvidences.Add .En_Birthday, "newEvOwnerBirthday" & tmp 'Год рождения ВД
'        frmNewEF.colEvidences.Add .En_Sex, "newEvOwnerSex" & tmp 'Пол владельца ВД
'    End With
'    With NewEvid '"вещественные доказательства"
'        frmNewEF.colEvidences.Add .DtmEvOutTake, "EvOutTake" & tmp  'Дата изъятия
'        frmNewEF.colEvidences.Add .DtmEvFirstDate, "EvFirstDate" & tmp  'Дата предоставления
'        frmNewEF.colEvidences.Add .strEvPlace, "EvPlace" & tmp  'Место изъятия
'        frmNewEF.colEvidences.Add .strEvPackage, "EvPackage" & tmp  'Упаковка
'        frmNewEF.colEvidences.Add .strEvStamp, "EvStamp" & tmp  'Печать
'        NewEvid.Print_EvNumArr1
'    End With
''добавление записи "принадлежность объаетов"
'       frmNewEF.colEvidences.Add Create_strOwner, "Owner" & tmp
''Добавление записи в список:
'        lstEvList.AddItem Create_strOwner
''Debug.Print "Ключ =" & tmp
'Debug.Print "Владелец = " & frmNewEF.colEvidences("Owner" & tmp)
'End Sub
''
'Private Static Function Create_strOwner() As String
''функция создания строки "принадлежность ВД"
'Dim strTmp1 As String, strTmp2 As String
'    If newEvOwner.En_SurName <> "" Then
'        strTmp1 = " " & newEvOwner.Create_InitialsFullName
'    Else: strTmp1 = " (принадлежность объектов не указана)"
'    End If
'    If NewEvid.strEvPlace <> "" Then
'        strTmp2 = ", изъятые по адресу: " & NewEvid.strEvPlace & ":"
'    Else: strTmp2 = ", (место изъятия не указано):"
'    End If
'Create_strOwner = "Объекты" & strTmp1 & strTmp2
'End Function
''


''Private Sub Ev_Disabled()
'''Процедура исключения полей "Заключение эксперта"
''    Dim X As Object
''        For Each X In Me.Controls
''            If InStrRev(X.Name, "txtEE", 5) > 0 Or InStrRev(X.Name, "cboEF", 5) > 0 Then
''                X.Enabled = False
''                X.BackColor = &HFDEADB
''            End If
''        Next X
''End Sub
'
'
'
'
''
''Private Sub Erase_ListItem()
'''Проверка и очистка списка от пустых строк
''    For Item = lstEvList.ListCount - 1 To 0 Step -1
''        If List1.List(Item) = "" Then
''            List1.RemoveItem Item
''        End If
''    Next
''End Sub
'
''Вариант создания динамического массива
''Dim X As Object
''Dim EvCount As Integer
''EvCount = LowBound
''Debug.Print "Элементы массива frmFlArr: "
''    Select Case EvCount
''        Case Is <= 5
''            txtEvOutTake.SetFocus
''            For Each X In Me.Controls
''                If TypeName(X) = "TextBox" Then
''                    ReDim Preserve frmFlArr(EvCount)
''                    frmFlArr(EvCount) = X.Text
''                    Debug.Print EvCount & "." & frmFlArr(EvCount)
''                    EvCount = EvCount + 1
''                End If
''            Next X
''        Case Is = 6
''            frmFlArr(EvCount) = lstOwSex.Text
''            Debug.Print EvCount & "." & frmFlArr(EvCount)
''        Case Is = 7
''            frmFlArr(EvCount) = newEvOwner.Create_InitialsInjPrFullName
''            Debug.Print frmFlArr(EvCount)
''        End Select
'
'''заполнение списка формы
''lstEvList.AddItem Create_strOwner1
''lblEvGRSum.Caption = NewEvid.mdCount.EvGRCount
'''Отладка массива:
''Dim v As Integer, d As Integer
''    If d = 0 Then
''        For v = LowBound To n
''            Debug.Print "Ячейка №" & d & "_" & v & " массива = " & arrEvidences(d, v)
''        Next v
''    Else
''        For d = LowBound To NewEvid.mdCount.EvGRCount
''            For v = LowBound To n
''            Debug.Print "Ячейка №" & d & "_" & v & " массива = " & arrEvidences(d, v)
''        Next v
''    Next d
''    End If
''End Sub

''Function IsNotEmptyArray(parArray As Variant) As Boolean
'''Функция проверки инициализации массива
''  On Error Resume Next
''  IsNotEmptyArray = LBound(parArray) <= UBound(parArray)
''End Function
'
