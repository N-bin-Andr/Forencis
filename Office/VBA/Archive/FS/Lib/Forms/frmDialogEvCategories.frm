VERSION 5.00
Begin VB.Form frmDialogEvCategories 
   BackColor       =   &H00C0CFC5&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Выбор экспертиз, входящих в комплексное исследование"
   ClientHeight    =   3795
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   7875
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3795
   ScaleWidth      =   7875
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command2 
      Caption         =   "<<"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3600
      TabIndex        =   5
      Top             =   1920
      Width           =   495
   End
   Begin VB.CommandButton Command1 
      Caption         =   ">>"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3600
      TabIndex        =   4
      Top             =   960
      Width           =   495
   End
   Begin VB.ListBox List2 
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   12
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2625
      Left            =   4320
      TabIndex        =   3
      Top             =   240
      Width           =   3135
   End
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   12
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2625
      Left            =   240
      TabIndex        =   2
      Top             =   240
      Width           =   3135
   End
   Begin VB.CommandButton CancelButton 
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1920
      TabIndex        =   1
      Top             =   3120
      Width           =   1215
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "OK"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4560
      TabIndex        =   0
      Top             =   3120
      Width           =   1215
   End
End
Attribute VB_Name = "frmDialogEvCategories"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Форма "frmDialogEvCategories"
'Ввод видов экспертиз, входящих в комплекс исследований (комплексная экспертиза)
'Дата создания: 01.06.2016
'@author Andr.Nab.n@gmail.com
Option Explicit

Private Sub Form_Initialize()
'инициализация формы
Dim myArray(1 To 10) As String
    myArray(1) = "("
    myArray(2) = "медико-криминалистической"
    myArray(3) = "дактилоскопической"
    myArray(4) = "судебно-биологической"
    myArray(5) = "судебно-генотипоскопической"
    myArray(6) = "трасологической"
    myArray(7) = "судебно-медицинской"
    myArray(8) = "судебно-химической"
    myArray(9) = "судебно-гистологической"
    myArray(10) = ")"
    
    List2.Columns = 1
    
    Wiht List1
        .Columns = 1
        .List = myArray
    End With
End Sub

