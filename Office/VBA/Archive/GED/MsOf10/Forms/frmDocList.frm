VERSION 5.00
Begin VB.Form frmDocList 
   BackColor       =   &H00D1CFA1&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Схемы"
   ClientHeight    =   2790
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5070
   BeginProperty Font 
      Name            =   "Cambria"
      Size            =   12
      Charset         =   204
      Weight          =   700
      Underline       =   0   'False
      Italic          =   -1  'True
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H008080FF&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2790
   ScaleWidth      =   5070
   ShowInTaskbar   =   0   'False
   Visible         =   0   'False
   Begin VB.CommandButton cmdAllCheck 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Выделить все"
      Height          =   375
      Left            =   2040
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   2280
      Width           =   1815
   End
   Begin VB.CommandButton cmdAllUnCheck 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Очистить все"
      Height          =   375
      Left            =   120
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   2280
      Width           =   1815
   End
   Begin VB.Frame fraMyDocuments 
      BackColor       =   &H00D1CFA1&
      Caption         =   "Какие документы необходимо создать:"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   1935
      Left            =   120
      TabIndex        =   1
      Top             =   240
      Width           =   4695
      Begin VB.CheckBox chkDocList 
         BackColor       =   &H00D1CFA1&
         Caption         =   "План проведения СМЭ"
         ForeColor       =   &H00404000&
         Height          =   375
         Index           =   6
         Left            =   120
         TabIndex        =   10
         Top             =   1440
         Width           =   2895
      End
      Begin VB.CheckBox chkDocList 
         BackColor       =   &H00D1CFA1&
         Caption         =   "Линейка"
         ForeColor       =   &H00404000&
         Height          =   375
         Index           =   1
         Left            =   120
         TabIndex        =   7
         Top             =   720
         Value           =   1  'Отмечено
         Width           =   1335
      End
      Begin VB.CheckBox chkDocList 
         BackColor       =   &H00D1CFA1&
         Caption         =   "Ростовая схема"
         ForeColor       =   &H00404000&
         Height          =   375
         Index           =   5
         Left            =   2280
         TabIndex        =   6
         Top             =   1080
         Width           =   2295
      End
      Begin VB.CheckBox chkDocList 
         BackColor       =   &H00D1CFA1&
         Caption         =   "Схема Голова"
         ForeColor       =   &H00404000&
         Height          =   375
         Index           =   4
         Left            =   2280
         TabIndex        =   5
         Top             =   720
         Width           =   2175
      End
      Begin VB.CheckBox chkDocList 
         BackColor       =   &H00D1CFA1&
         Caption         =   "Черновик"
         ForeColor       =   &H00404000&
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   4
         Top             =   360
         Value           =   1  'Отмечено
         Width           =   2055
      End
      Begin VB.CheckBox chkDocList 
         BackColor       =   &H00D1CFA1&
         Caption         =   "Схема ребер"
         ForeColor       =   &H00404000&
         Height          =   375
         Index           =   3
         Left            =   2280
         TabIndex        =   3
         Top             =   360
         Width           =   2055
      End
      Begin VB.CheckBox chkDocList 
         BackColor       =   &H00D1CFA1&
         Caption         =   "Фототаблица"
         ForeColor       =   &H00404000&
         Height          =   375
         Index           =   2
         Left            =   120
         MaskColor       =   &H00C0C0FF&
         TabIndex        =   2
         Top             =   1080
         Width           =   2055
      End
   End
   Begin VB.CommandButton cmdOK3 
      BackColor       =   &H00C0FFC0&
      Caption         =   "&OK"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3960
      MaskColor       =   &H00C0FFC0&
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   2280
      Width           =   855
   End
End
Attribute VB_Name = "frmDocList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Форма "Список создаваемых документов"
'@author Andr.N@_bin
'@E-mail Andr.Nab.n@gmail.com
Option Explicit
Dim Cancel As Integer 'переменная кнопок "Отмена"
'
Private Sub cmdAllCheck_Click()
Dim X As Object
For Each X In Me.Controls
    If TypeName(X) = "CheckBox" Then
        X.Value = 1
    End If
Next X
End Sub
'
Private Sub cmdAllUnCheck_Click()
Dim X As Object
For Each X In Me.Controls
    If TypeName(X) = "CheckBox" Then
        X.Value = 0
    End If
Next X
End Sub
'
Private Sub cmdOK3_Click()
    Call frmNewResearch.create_Blank
'Закрытие текущей формы и возврат к форме "Новое исследование"
    frmNewResearch.Visible = True
    Me.Visible = False
End Sub
'
Private Sub Form_Load()
 Move (Screen.Width - Width) / 2, (Screen.Height - Height) / 2
End Sub
