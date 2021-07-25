VERSION 5.00
Begin VB.Form frmAnalysis 
   BackColor       =   &H00CFBFAC&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Направления"
   ClientHeight    =   3945
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6000
   BeginProperty Font 
      Name            =   "Calibri"
      Size            =   12
      Charset         =   204
      Weight          =   700
      Underline       =   0   'False
      Italic          =   -1  'True
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3945
   ScaleWidth      =   6000
   ShowInTaskbar   =   0   'False
   Visible         =   0   'False
   Begin VB.CommandButton cmdAllUnCheck 
      BackColor       =   &H00C0E0FF&
      Caption         =   "&Очисть все"
      Height          =   375
      Left            =   120
      MaskColor       =   &H00808000&
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   3360
      UseMaskColor    =   -1  'True
      Width           =   1815
   End
   Begin VB.CommandButton cmdAllCheck 
      BackColor       =   &H00C0E0FF&
      Caption         =   "&Выделить все"
      Height          =   375
      Left            =   2280
      MaskColor       =   &H00C0C000&
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   3360
      Width           =   1815
   End
   Begin VB.CommandButton cmdCancel3 
      BackColor       =   &H00C0C0FF&
      Caption         =   "&Отмена"
      Height          =   375
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   120
      Width           =   1815
   End
   Begin VB.Frame fraMyDocuments 
      BackColor       =   &H00CFBFAC&
      Caption         =   "Какие бланки необходимо создать?"
      ForeColor       =   &H00800000&
      Height          =   2655
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   5655
      Begin VB.CheckBox chkAnalysis 
         BackColor       =   &H00CFBFAC&
         Caption         =   "На планктон"
         ForeColor       =   &H00404000&
         Height          =   375
         Index           =   2
         Left            =   2760
         TabIndex        =   14
         Top             =   1440
         Width           =   2175
      End
      Begin VB.CheckBox chkAnalysis 
         BackColor       =   &H00CFBFAC&
         Caption         =   "На определение СО"
         ForeColor       =   &H00404000&
         Height          =   375
         Index           =   11
         Left            =   2760
         TabIndex        =   13
         Top             =   1080
         Width           =   2415
      End
      Begin VB.CheckBox chkAnalysis 
         BackColor       =   &H00CFBFAC&
         Caption         =   "На криминалистику"
         ForeColor       =   &H00404000&
         Height          =   375
         Index           =   5
         Left            =   240
         TabIndex        =   12
         Top             =   2160
         Width           =   2535
      End
      Begin VB.CheckBox chkAnalysis 
         BackColor       =   &H00CFBFAC&
         Caption         =   "На алкоголь"
         ForeColor       =   &H00404000&
         Height          =   375
         Index           =   0
         Left            =   240
         TabIndex        =   11
         Top             =   720
         Value           =   1  'Отмечено
         Width           =   2175
      End
      Begin VB.CheckBox chkAnalysis 
         BackColor       =   &H00CFBFAC&
         Caption         =   "На гликоген"
         ForeColor       =   &H00404000&
         Height          =   375
         Index           =   7
         Left            =   2760
         TabIndex        =   10
         Top             =   720
         Width           =   2415
      End
      Begin VB.CheckBox chkAnalysis 
         BackColor       =   &H00CFBFAC&
         Caption         =   "На геном"
         ForeColor       =   &H00404000&
         Height          =   375
         Index           =   6
         Left            =   2760
         TabIndex        =   9
         Top             =   360
         Width           =   2415
      End
      Begin VB.CheckBox chkAnalysis 
         BackColor       =   &H00CFBFAC&
         Caption         =   "На биохимию"
         ForeColor       =   &H00404000&
         Height          =   375
         Index           =   3
         Left            =   240
         TabIndex        =   7
         Top             =   1440
         Width           =   2175
      End
      Begin VB.CheckBox chkAnalysis 
         BackColor       =   &H00CFBFAC&
         Caption         =   "На биологию"
         ForeColor       =   &H00404000&
         Height          =   375
         Index           =   8
         Left            =   240
         TabIndex        =   6
         Top             =   1080
         Width           =   2055
      End
      Begin VB.CheckBox chkAnalysis 
         BackColor       =   &H00CFBFAC&
         Caption         =   "На тяжелые металлы"
         ForeColor       =   &H00404000&
         Height          =   375
         Index           =   10
         Left            =   2760
         TabIndex        =   5
         Top             =   2160
         Width           =   2775
      End
      Begin VB.CheckBox chkAnalysis 
         BackColor       =   &H00CFBFAC&
         Caption         =   "На гистологию"
         ForeColor       =   &H00404000&
         Height          =   375
         Index           =   1
         Left            =   240
         TabIndex        =   4
         Top             =   360
         Value           =   1  'Отмечено
         Width           =   2175
      End
      Begin VB.CheckBox chkAnalysis 
         BackColor       =   &H00CFBFAC&
         Caption         =   "На холинэстеразу"
         ForeColor       =   &H00404000&
         Height          =   375
         Index           =   4
         Left            =   2760
         TabIndex        =   3
         Top             =   1800
         Width           =   2415
      End
      Begin VB.CheckBox chkAnalysis 
         BackColor       =   &H00CFBFAC&
         Caption         =   "На общую химию"
         ForeColor       =   &H00404000&
         Height          =   375
         Index           =   9
         Left            =   240
         TabIndex        =   2
         Top             =   1800
         Width           =   2535
      End
   End
   Begin VB.CommandButton cmdOK3 
      BackColor       =   &H00C0FFC0&
      Caption         =   "&OK"
      DownPicture     =   "frmAnalysis.frx":0000
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   15.75
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   4680
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   3360
      Width           =   1095
   End
End
Attribute VB_Name = "frmAnalysis"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Форма "Список создаваемых направлений"
'@author Andr.N@_bin
'@E-mail Andr.Nab.n@gmail.com
'&H00CFBFAC&
'&H00D1AFAC&
Option Explicit
'
Private Sub cmdAllCheck_Click()
'Процедура выбора всех исследований
Dim X As Object
For Each X In Me.Controls
    If TypeName(X) = "CheckBox" Then
        X.Value = 1
    End If
Next X
End Sub
'
Private Sub cmdAllUnCheck_Click()
'Процедура отмены выбора всех исследований
Dim X As Object
For Each X In Me.Controls
    If TypeName(X) = "CheckBox" Then
        X.Value = 0
    End If
Next X
End Sub
'
Private Sub cmdCancel3_Click()
    With frmNewResearch
        .Show
        Set .newDOC = Nothing  'уничтожение экземпляра класса "Создаваемые Документы"
    End With
frmAnalysis.Hide
End Sub
'
Private Sub cmdOK3_Click()
'Процедура нажатия кнопки "ОК"
    Call frmNewResearch.create_Analysis
frmDocList.Show
Me.Hide
End Sub
'
Private Sub Form_Load()
    Move (Screen.Width - Width) / 2, (Screen.Height - Height) / 3.7
End Sub
