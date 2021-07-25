VERSION 5.00
Begin VB.Form frmDocuments 
   BackColor       =   &H00CFBFAC&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Формирование пакета документов"
   ClientHeight    =   3240
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9240
   BeginProperty Font 
      Name            =   "Arial"
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
   ScaleHeight     =   3240
   ScaleWidth      =   9240
   ShowInTaskbar   =   0   'False
   Visible         =   0   'False
   Begin VB.CommandButton cmdAllUnCheck 
      BackColor       =   &H00C0E0FF&
      Caption         =   "&Очисть все"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2520
      MaskColor       =   &H00808000&
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   2640
      UseMaskColor    =   -1  'True
      Width           =   1935
   End
   Begin VB.CommandButton cmdAllCheck 
      BackColor       =   &H00C0E0FF&
      Caption         =   "&Выделить все"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4800
      MaskColor       =   &H00C0C000&
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   2640
      Width           =   1935
   End
   Begin VB.CommandButton cmdCancel3 
      BackColor       =   &H00C0C0FF&
      Caption         =   "&Отмена"
      Height          =   375
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   2640
      Width           =   1935
   End
   Begin VB.Frame fraMyDocuments 
      BackColor       =   &H00CFBFAC&
      Caption         =   "Какие документы необходимо создать?"
      ForeColor       =   &H00800000&
      Height          =   2295
      Left            =   240
      TabIndex        =   1
      Top             =   240
      Width           =   8655
      Begin VB.CheckBox chkAnalysis 
         BackColor       =   &H00CFBFAC&
         Caption         =   "Химия сопровод."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404000&
         Height          =   375
         Index           =   1
         Left            =   240
         TabIndex        =   13
         Tag             =   "Chemistry"
         Top             =   1800
         Width           =   3375
      End
      Begin VB.CheckBox chkAnalysis 
         BackColor       =   &H00CFBFAC&
         Caption         =   "Описание ВД для биологов"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404000&
         Height          =   375
         Index           =   2
         Left            =   3720
         TabIndex        =   12
         Tag             =   "Description"
         Top             =   360
         Width           =   3375
      End
      Begin VB.CheckBox chkAnalysis 
         BackColor       =   &H00CFBFAC&
         Caption         =   "Ходатайтво о продлении сроков СМЭ"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404000&
         Height          =   375
         Index           =   11
         Left            =   3720
         TabIndex        =   9
         Top             =   1800
         Width           =   4815
      End
      Begin VB.CheckBox chkAnalysis 
         BackColor       =   &H00CFBFAC&
         Caption         =   "Геном сопровод."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404000&
         Height          =   375
         Index           =   0
         Left            =   240
         TabIndex        =   8
         Tag             =   "Genome"
         Top             =   720
         Width           =   3375
      End
      Begin VB.CheckBox chkAnalysis 
         BackColor       =   &H00CFBFAC&
         Caption         =   "Акт несоответствия"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404000&
         Height          =   375
         Index           =   7
         Left            =   3720
         TabIndex        =   7
         Tag             =   "Nesootvetstvie"
         Top             =   1440
         Width           =   4815
      End
      Begin VB.CheckBox chkAnalysis 
         BackColor       =   &H00CFBFAC&
         Caption         =   "Ходатайство следователю"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404000&
         Height          =   375
         Index           =   6
         Left            =   3720
         TabIndex        =   6
         Tag             =   "Hodataistvo"
         Top             =   1080
         Width           =   4815
      End
      Begin VB.CheckBox chkAnalysis 
         BackColor       =   &H00CFBFAC&
         Caption         =   "Дактилоскопия сопровод."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404000&
         Height          =   375
         Index           =   3
         Left            =   240
         TabIndex        =   4
         Tag             =   "Dactylography"
         Top             =   1440
         Width           =   3375
      End
      Begin VB.CheckBox chkAnalysis 
         BackColor       =   &H00CFBFAC&
         Caption         =   "Трасология сопровод."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404000&
         Height          =   375
         Index           =   8
         Left            =   240
         TabIndex        =   3
         Tag             =   "Trasology"
         Top             =   1080
         Width           =   3375
      End
      Begin VB.CheckBox chkchkAnalysis 
         BackColor       =   &H00CFBFAC&
         Caption         =   "Биология сопровод."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404000&
         Height          =   375
         Index           =   1
         Left            =   240
         TabIndex        =   2
         Tag             =   "Biology"
         Top             =   360
         Width           =   3375
      End
   End
   Begin VB.CommandButton cmdOK3 
      BackColor       =   &H00C0FFC0&
      Caption         =   "&OK"
      DownPicture     =   "frmDocument.frx":0000
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   7080
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   2640
      Width           =   1815
   End
End
Attribute VB_Name = "frmDocuments"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Форма "Формирование пакета документов"
'@author Andr.N@_bin
'@E-mail Andr.Nab.n@gmail.com
'&H00CFBFAC&
'&H00D1AFAC&
Option Explicit
'
Private Sub cmdAllCheck_Click()
'Процедура выбора всех исследований
Dim x As Object
For Each x In Me.Controls
    If TypeName(x) = "CheckBox" Then
        x.Value = 1
    End If
Next x
End Sub
'
Private Sub cmdAllUnCheck_Click()
'Процедура отмены выбора всех исследований
Dim x As Object
For Each x In Me.Controls
    If TypeName(x) = "CheckBox" Then
        x.Value = 0
    End If
Next x
End Sub
'
Private Sub cmdCancel3_Click()
    frmNewEF.Show
'        Set .newDOC = Nothing  'уничтожение экземпляра класса "Создаваемые Документы"
    Me.Hide
End Sub
'
Private Sub cmdOK3_Click()
'Процедура нажатия кнопки "ОК"
    Call mdPrintDoc.withApplWD(mdMainFolders.arrDocDir(1), mdMainFolders.arrDocDir(0))
    Call mdPrintDoc.openEx
   frmDocuments.Hide
  If MsgBox("Документы успешно созданы!", vbYes, "Отчет") = vbYes Then
    Unload frmNewEF
    Unload Me
  End If
End Sub
'
Private Sub Form_Load()
    Move (Screen.Width - Width) / 2, (Screen.Height - Height) / 3.7
End Sub
'
