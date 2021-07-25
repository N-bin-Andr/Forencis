VERSION 5.00
Begin VB.Form frmEvCategories 
   BackColor       =   &H00CFC2AC&
   Caption         =   "Мои шаблоны"
   ClientHeight    =   8085
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6240
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8085
   ScaleWidth      =   6240
   Visible         =   0   'False
   Begin VB.CommandButton cmdOK2 
      BackColor       =   &H00C0FFC0&
      Caption         =   "&OK"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   20.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4440
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   7200
      Width           =   1455
   End
   Begin VB.Frame fraWdShablons 
      BackColor       =   &H00DDEADB&
      Caption         =   "Выбирете шаблон заключения:"
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
      Height          =   6855
      Index           =   0
      Left            =   360
      TabIndex        =   1
      Top             =   120
      Width           =   5535
      Begin VB.OptionButton BlStainsAnalysis 
         BackColor       =   &H00DDEADB&
         Caption         =   "Ситуационный анализ (ОМП)"
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   12
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   15
         Left            =   240
         TabIndex        =   18
         Top             =   6120
         Width           =   5055
      End
      Begin VB.OptionButton Dactylography 
         BackColor       =   &H00DDEADB&
         Caption         =   "Кисти рук (восстановление папиллярных узоров)"
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   12
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   14
         Left            =   240
         TabIndex        =   17
         Top             =   5520
         Width           =   5055
      End
      Begin VB.OptionButton CrimCaseDoc 
         BackColor       =   &H00DDEADB&
         Caption         =   "Материалы дел"
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   12
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   13
         Left            =   240
         TabIndex        =   16
         Top             =   5160
         Width           =   5055
      End
      Begin VB.OptionButton PhSuperposition 
         BackColor       =   &H00DDEADB&
         Caption         =   "Препарат черепа (фотосовмещение)"
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   12
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   12
         Left            =   240
         TabIndex        =   15
         Top             =   4800
         Width           =   5055
      End
      Begin VB.OptionButton Pelvis 
         BackColor       =   &H00DDEADB&
         Caption         =   "Препараты костей таза"
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   12
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   11
         Left            =   240
         TabIndex        =   14
         Tag             =   "Pelvis"
         Top             =   4440
         Width           =   5055
      End
      Begin VB.OptionButton Scapula 
         BackColor       =   &H00DDEADB&
         Caption         =   "Препараты лопатки"
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   12
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   10
         Left            =   240
         TabIndex        =   13
         Tag             =   "Scapula"
         Top             =   4080
         Width           =   5055
      End
      Begin VB.OptionButton Cartilage 
         BackColor       =   &H00DDEADB&
         Caption         =   "Подъязычная кость и хрящи гортани"
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   12
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   9
         Left            =   240
         TabIndex        =   12
         Tag             =   "Cartilage"
         Top             =   3720
         Width           =   5055
      End
      Begin VB.OptionButton Identification 
         BackColor       =   &H00DDEADB&
         Caption         =   "Костные останки (идентификация)"
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   12
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   8
         Left            =   240
         TabIndex        =   11
         Tag             =   "Identification"
         Top             =   3360
         Width           =   5055
      End
      Begin VB.OptionButton Vertebra 
         BackColor       =   &H00DDEADB&
         Caption         =   "Препарат позвоночника"
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   12
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   7
         Left            =   240
         TabIndex        =   10
         Tag             =   "Vertebra"
         Top             =   3000
         Width           =   5055
      End
      Begin VB.OptionButton Costa 
         BackColor       =   &H00DDEADB&
         Caption         =   "Препараты ребер"
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   12
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   6
         Left            =   240
         TabIndex        =   9
         Tag             =   "Costa"
         Top             =   2640
         Width           =   5055
      End
      Begin VB.OptionButton Cranium 
         BackColor       =   &H00DDEADB&
         Caption         =   "Препарат черепа"
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   12
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   5
         Left            =   240
         TabIndex        =   8
         Tag             =   "Cranium"
         Top             =   2280
         Width           =   5055
      End
      Begin VB.OptionButton Bones 
         BackColor       =   &H00DDEADB&
         Caption         =   "Трубчатые кости"
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   12
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   4
         Left            =   240
         TabIndex        =   7
         Tag             =   "Bones"
         Top             =   1920
         Width           =   5055
      End
      Begin VB.OptionButton ShotGunInjury 
         BackColor       =   &H00DDEADB&
         Caption         =   "Огнестрельные повреждения"
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   12
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   3
         Left            =   240
         TabIndex        =   6
         Tag             =   "ShotGunInjury"
         Top             =   1560
         Width           =   5055
      End
      Begin VB.OptionButton Clothes 
         BackColor       =   &H00DDEADB&
         Caption         =   "Следы крови"
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   12
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   2
         Left            =   240
         TabIndex        =   5
         Tag             =   "Clothes"
         Top             =   1200
         Width           =   5055
      End
      Begin VB.OptionButton Skin 
         BackColor       =   &H00DDEADB&
         Caption         =   "Препараты кожи"
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   12
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   240
         TabIndex        =   4
         Tag             =   "Skin"
         Top             =   840
         Width           =   5055
      End
      Begin VB.OptionButton Simple 
         BackColor       =   &H00DDEADB&
         Caption         =   "Общий"
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   12
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   240
         TabIndex        =   3
         Tag             =   "Simple"
         Top             =   480
         Value           =   -1  'True
         Width           =   5055
      End
   End
   Begin VB.CommandButton cmdCancel2 
      BackColor       =   &H00C0C0FF&
      Caption         =   "&Отмена"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   15.75
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   7200
      Width           =   1455
   End
End
Attribute VB_Name = "frmEvCategories"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Форма "frmEvCategories"
'Обработка данных при создании медико-криминалистических заключений
'Дата создания: 01.06.2016
'@version 0.0.1
'@author Andr.Nab.n@gmail.com
Option Explicit
Dim Cancel As Integer 'переменная кнопок "Отмена"
'
Private Sub cmdCancel2_Click()
    Unload Me
End Sub
'
Public Static Sub Select_EFCategogy()
Dim x As Object
    For Each x In Me.Controls
        If TypeName(x) = "OptionButton" Then
            If x.Value = -1 Then
'!!!Вставить активацию и заполнение коллекций в модуле mdPrintDoc
'   например: mdPrintDoc.colExcelData.Add x.Caption, "B"
                mdPrintDoc.DocCat = x.Caption 'категория создваемого документа
                mdPrintDoc.strDOT = x.name & ".docm" ' ".dotm"
'Добавляем ссылку на переменную "выбранный шаблон" в коллекцию
                Set mdPrintDoc.colExcelData = New Collection
                 mdPrintDoc.colExcelData.Add x.Caption, "B"
                 Debug.Print "Запись категории в коллекции colExcelData -> " & mdPrintDoc.colExcelData.Item("B")
'           mdPrintDoc.colExcelData.Add strDOC, "Category" & mdCount.fEvCat_Count
'Debug.Print "Имя шаблона = ", mdPrintDoc.strDOT
            End If
        End If
    Next x
End Sub
'
Private Sub cmdOK2_Click()
'нажитие кнопки "ОК" на форме
Call Select_EFCategogy
    With frmNewEF
        .Caption = "Создание нового документа " & "(" & mdPrintDoc.DocCat & ")"
        .Show
    End With
'Me.Hide
Unload Me
'Lib
'mdCount.fEFCount = mdCount.fEFCount + 1
''обявление новой формы
'    Dim frmD  As frmNewEF
'    Set frmD = New frmNewEF
''добавление формы в список созданных форм
'    frmD.Caption = mdPrintDoc.DocCat & Chr(32) & Me.EvCat_ID
'    frmD.lblEFCount.Caption = Me.EvCat_ID
'    mdCount.colForms.Add frmD, "EFform" & Me.EvCat_ID
'Debug.Print "ID формы EF - ", frmD.lblEFCount.Caption
'    frmD.Show
'    Me.Hide
''    Unload Me
End Sub
'
Private Sub Form_Load()
    Move (Screen.Width - Width) / 0.5, (Screen.Height - Height) / 3.5
End Sub
Private Sub Form_Unload(Cancel As Integer)
'Процедура выгрузки формы
'    Dim str As String
'        str = "Закрыть форму " '& mdCount.colForms & " выбора категории документа?"
'    If MsgBox("Закрыть форму выбора категории документа?", vbYesNo, "Выход?") = vbYes Then
''        Set newFrm = Nothing
'    Else
'        Cancel = 1
'    End If
End Sub

