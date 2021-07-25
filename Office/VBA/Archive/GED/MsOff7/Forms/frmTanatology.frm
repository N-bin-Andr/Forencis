VERSION 5.00
Begin VB.Form frmNewResearch 
   BackColor       =   &H00BBC0AC&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   7860
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11820
   ForeColor       =   &H00000000&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5941.177
   ScaleLeft       =   100
   ScaleMode       =   0  'Пользовательское
   ScaleWidth      =   7201.652
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdSaveCreate 
      BackColor       =   &H00FFC0C0&
      Caption         =   "++"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9840
      Style           =   1  'Graphical
      TabIndex        =   50
      Top             =   7080
      Width           =   615
   End
   Begin VB.TextBox txtCsIndex 
      Alignment       =   2  'Центровка
      BackColor       =   &H00FEF7F1&
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
      Left            =   1680
      TabIndex        =   6
      Top             =   2400
      Width           =   855
   End
   Begin VB.ComboBox cboCsCat 
      BackColor       =   &H00FEF7F1&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   1680
      TabIndex        =   5
      Top             =   1920
      Width           =   4095
   End
   Begin VB.TextBox txtCsNum 
      Alignment       =   2  'Центровка
      BackColor       =   &H00FEF7F1&
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
      Left            =   3000
      MaxLength       =   11
      TabIndex        =   7
      Top             =   2400
      Width           =   2775
   End
   Begin VB.ComboBox cboCsDefinition 
      BackColor       =   &H00FEF7F1&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   1680
      TabIndex        =   3
      Top             =   1320
      Width           =   2295
   End
   Begin VB.Frame fraFactCase 
      BackColor       =   &H00BBC0AC&
      Caption         =   "Краткие обстоятельства дела"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   3375
      Left            =   6240
      TabIndex        =   45
      Top             =   3600
      Width           =   5295
      Begin VB.TextBox txtFactCase 
         BackColor       =   &H00DBF8F4&
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2895
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Вертикаль
         TabIndex        =   22
         Top             =   360
         Width           =   5055
      End
   End
   Begin VB.ComboBox cboEFCategories 
      BackColor       =   &H00FEF7F1&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   2160
      TabIndex        =   2
      Top             =   720
      Width           =   3615
   End
   Begin VB.CommandButton cmdEraseData 
      BackColor       =   &H008080FF&
      Caption         =   "&Очистить"
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
      Left            =   6240
      MaskColor       =   &H008080FF&
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   7080
      Width           =   1575
   End
   Begin VB.TextBox txtCsPrvDate 
      Alignment       =   2  'Центровка
      BackColor       =   &H00FEF7F1&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   4440
      TabIndex        =   4
      Top             =   1320
      Width           =   1335
   End
   Begin VB.TextBox txtDN 
      Alignment       =   2  'Центровка
      BackColor       =   &H00FEF7F1&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   405
      Left            =   600
      MaxLength       =   5
      TabIndex        =   0
      Top             =   240
      Width           =   975
   End
   Begin VB.Frame fraInjuredPerson 
      BackColor       =   &H00D1CFA1&
      Caption         =   "Потерпевший"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808000&
      Height          =   3375
      Left            =   6240
      TabIndex        =   32
      Top             =   120
      Width           =   5295
      Begin VB.TextBox txtInjPrAutopsy 
         Alignment       =   2  'Центровка
         BackColor       =   &H00F0FBE1&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   2160
         TabIndex        =   21
         Top             =   2880
         Width           =   1695
      End
      Begin VB.TextBox txtInjPrDecease 
         Alignment       =   2  'Центровка
         BackColor       =   &H00F0FBE1&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   2160
         TabIndex        =   20
         Top             =   2400
         Width           =   1695
      End
      Begin VB.ListBox lstInjPrSex 
         BackColor       =   &H00F0FBE1&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   330
         Left            =   3960
         TabIndex        =   19
         Top             =   2880
         Width           =   1095
      End
      Begin VB.TextBox txtInjPrSurName 
         BackColor       =   &H00F0FBE1&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   1920
         TabIndex        =   15
         Top             =   420
         Width           =   3135
      End
      Begin VB.TextBox txtInjPrName 
         BackColor       =   &H00F0FBE1&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   360
         Left            =   1920
         TabIndex        =   16
         Top             =   960
         Width           =   3135
      End
      Begin VB.TextBox txtInjPrMidName 
         BackColor       =   &H00F0FBE1&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   360
         Left            =   1920
         TabIndex        =   17
         Top             =   1440
         Width           =   3135
      End
      Begin VB.TextBox txtInjPrBirthday 
         Alignment       =   2  'Центровка
         BackColor       =   &H00F0FBE1&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   360
         Left            =   1920
         MaxLength       =   10
         TabIndex        =   18
         Top             =   1920
         Width           =   1815
      End
      Begin VB.Label lblInjPrAutopsy 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Прозрачно
         Caption         =   "Дата вскрытия"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404000&
         Height          =   375
         Left            =   120
         TabIndex        =   46
         Top             =   2880
         Width           =   1935
      End
      Begin VB.Label lblInjPrDecease 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Прозрачно
         Caption         =   "Дата смерти"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404000&
         Height          =   375
         Left            =   240
         TabIndex        =   44
         Top             =   2400
         Width           =   1695
      End
      Begin VB.Label lblInjPrSex 
         Alignment       =   2  'Центровка
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Прозрачно
         Caption         =   "Пол"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404000&
         Height          =   375
         Left            =   3960
         TabIndex        =   43
         Top             =   2520
         Width           =   495
      End
      Begin VB.Label lblInjPrSurName 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Прозрачно
         Caption         =   "Фамилия"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404000&
         Height          =   375
         Left            =   240
         TabIndex        =   36
         Top             =   480
         Width           =   1695
      End
      Begin VB.Label lblInjPrName 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Прозрачно
         Caption         =   "Имя"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404000&
         Height          =   375
         Left            =   240
         TabIndex        =   35
         Top             =   960
         Width           =   1695
      End
      Begin VB.Label lblInjPrMidName 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Прозрачно
         Caption         =   "Отчество"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404000&
         Height          =   375
         Left            =   240
         TabIndex        =   34
         Top             =   1440
         Width           =   1695
      End
      Begin VB.Label lblInjPrBirthday 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Прозрачно
         Caption         =   "Год рождения"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404000&
         Height          =   255
         Left            =   240
         TabIndex        =   33
         Top             =   1920
         Width           =   1575
      End
   End
   Begin VB.TextBox txtEFFirstDay 
      Alignment       =   2  'Центровка
      BackColor       =   &H00FEF7F1&
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "dd MMM yy"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1049
         SubFormatType   =   3
      EndProperty
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   4080
      MaxLength       =   10
      TabIndex        =   1
      Top             =   240
      Width           =   1695
   End
   Begin VB.Frame fraCoroner 
      BackColor       =   &H00CFC2AC&
      Caption         =   "Следователь"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808000&
      Height          =   4575
      Left            =   240
      TabIndex        =   26
      Top             =   3000
      Width           =   5535
      Begin VB.ListBox lstCrSex 
         BackColor       =   &H00FFE8DF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   4200
         TabIndex        =   14
         Top             =   3960
         Width           =   1095
      End
      Begin VB.TextBox txtCrPost 
         Alignment       =   2  'Центровка
         BackColor       =   &H00FFE8DF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   1560
         MultiLine       =   -1  'True
         TabIndex        =   8
         Top             =   360
         Width           =   3735
      End
      Begin VB.TextBox txtCrOprAr 
         Alignment       =   2  'Центровка
         BackColor       =   &H00FFE8DF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1080
         Left            =   1560
         MultiLine       =   -1  'True
         TabIndex        =   9
         Top             =   840
         Width           =   3735
      End
      Begin VB.TextBox txtCrRank 
         Alignment       =   2  'Центровка
         BackColor       =   &H00FFE8DF&
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
         Left            =   1560
         MultiLine       =   -1  'True
         TabIndex        =   10
         Top             =   2040
         Width           =   3735
      End
      Begin VB.TextBox txtCrSurName 
         BackColor       =   &H00FFE8DF&
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
         Left            =   1560
         TabIndex        =   11
         Top             =   2520
         Width           =   3735
      End
      Begin VB.TextBox txtCrName 
         BackColor       =   &H00FFE8DF&
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
         Left            =   1560
         TabIndex        =   12
         Top             =   3000
         Width           =   3735
      End
      Begin VB.TextBox txtCrMidName 
         BackColor       =   &H00FFE8DF&
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
         Left            =   1560
         TabIndex        =   13
         Top             =   3480
         Width           =   3735
      End
      Begin VB.Label Label1 
         BackColor       =   &H00CFC2AC&
         BackStyle       =   0  'Прозрачно
         Caption         =   "Должность"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   375
         Left            =   120
         TabIndex        =   49
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label lblCrSex 
         Alignment       =   2  'Центровка
         BackColor       =   &H00CFC2AC&
         BackStyle       =   0  'Прозрачно
         Caption         =   "Пол"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   375
         Left            =   3600
         TabIndex        =   42
         Top             =   3960
         Width           =   495
      End
      Begin VB.Label lblCrOprAr 
         BackColor       =   &H00CFC2AC&
         BackStyle       =   0  'Прозрачно
         Caption         =   "Отдел"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   375
         Left            =   120
         TabIndex        =   31
         Top             =   960
         Width           =   855
      End
      Begin VB.Label lblCrRank 
         BackColor       =   &H00CFC2AC&
         BackStyle       =   0  'Прозрачно
         Caption         =   "Звание"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   375
         Left            =   120
         TabIndex        =   30
         Top             =   2040
         Width           =   1335
      End
      Begin VB.Label lblCrSurName 
         BackColor       =   &H00CFC2AC&
         BackStyle       =   0  'Прозрачно
         Caption         =   "Фамилия"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   375
         Left            =   120
         TabIndex        =   29
         Top             =   2520
         Width           =   1335
      End
      Begin VB.Label lblCrMidName 
         BackColor       =   &H00CFC2AC&
         BackStyle       =   0  'Прозрачно
         Caption         =   "Отчество"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   375
         Left            =   120
         TabIndex        =   28
         Top             =   3480
         Width           =   1335
      End
      Begin VB.Label lblCrName 
         BackColor       =   &H00CFC2AC&
         BackStyle       =   0  'Прозрачно
         Caption         =   "Имя"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   375
         Left            =   120
         TabIndex        =   27
         Top             =   3000
         Width           =   1335
      End
   End
   Begin VB.CommandButton cmdOK1 
      BackColor       =   &H00B1CFAF&
      Caption         =   "&OK"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   10800
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   7080
      Width           =   735
   End
   Begin VB.CommandButton cmdCancel1 
      BackColor       =   &H00C0C0FF&
      Caption         =   "&Отмена"
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
      Left            =   7920
      Style           =   1  'Graphical
      TabIndex        =   25
      Top             =   7080
      Width           =   1575
   End
   Begin VB.Label lblCrimCase 
      BackColor       =   &H00BBC2AC&
      Caption         =   "Материалы"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404000&
      Height          =   255
      Left            =   120
      TabIndex        =   48
      Top             =   1980
      Width           =   1575
   End
   Begin VB.Label lblCsNum 
      BackColor       =   &H00BBC2AC&
      Caption         =   "№"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404000&
      Height          =   375
      Left            =   2640
      TabIndex        =   47
      Top             =   2400
      Width           =   375
   End
   Begin VB.Label lblCsDefinition 
      BackColor       =   &H00BBC2AC&
      Caption         =   "Основание:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404000&
      Height          =   375
      Left            =   120
      TabIndex        =   41
      Top             =   1320
      Width           =   1575
   End
   Begin VB.Label lblEFCategories 
      BackColor       =   &H00BBC2AC&
      Caption         =   "Вид экспертизы"
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
      Height          =   360
      Left            =   120
      TabIndex        =   40
      Top             =   735
      Width           =   2055
   End
   Begin VB.Label lblCsPrvDate 
      Alignment       =   2  'Центровка
      BackColor       =   &H00BBC2AC&
      Caption         =   "от"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404000&
      Height          =   405
      Left            =   4080
      TabIndex        =   39
      Top             =   1320
      Width           =   315
   End
   Begin VB.Label lblDN 
      BackColor       =   &H00BBC2AC&
      Caption         =   "№"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404000&
      Height          =   375
      Left            =   240
      TabIndex        =   38
      Top             =   255
      Width           =   375
   End
   Begin VB.Label lblEFFirstDay 
      Alignment       =   2  'Центровка
      BackColor       =   &H00BBC2AC&
      Caption         =   "Дата начала"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404000&
      Height          =   375
      Left            =   2280
      TabIndex        =   37
      Top             =   255
      Width           =   1695
   End
End
Attribute VB_Name = "frmNewResearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Форма "Новое исследование"
'@author Andr.N@_bin
'@E-mail Andr.Nab.n@gmail.com
'Форма: BackColor - &H00BBC2AC& / &H00FEF7F1&
Option Explicit
'Объявление Классов
Public newCase As clmCase          'Экземпляр класса "Дело"
Public newCor As clmHR_Official    'Экземпляр класса "Следователь" Coroner
Public SvcService As Object        'Экземпляр библиотеки svcsvc.dll
Public newEF As clmExpertFindings  'Экземпляр класса "Заключение эксперта"
Public newInjPr As clmEntrants     'Экземпляр класса "Потерпевший"InjuredPerson
Public newDate As clmCaseDate      'Экземпляр класса "Дата"
Public newDOC As clmCreateDoc      'Экземпляр класса "Создаваемые Документы"
'Переменные
Public newExpert As String       'переменная "Эксперт"
'Ошибки:
Const Msg As String = "Ошибка ввода данных!"
'Colors
Private Enum frmColor
    Светло_Желтый = &HC0FFFF 'RGB(102, 102, 153)
    Желтый = &HFFFF&
    Красный = &HFF&
    Черный = &H0&
    Салатовый = &HC0FFC0 'RGB(200, 256, 200)
    Светло_голубой = &HFEF7F1
    Фиолетовый = &HFFC0C0
    Коричневый = &H80FF&
End Enum
'Кнопка "Cancel"
Dim Cancel As Integer
'
'++++++++++++++++++++++++++++++++++++  И Н И Ц И А Л И З А Ц И Я +++++++++++++++++++++++++++++++++++
'
Public Static Sub Form_Initialize()
'Инициализация формы
End Sub
'
Private Sub Form_Load()
'Загрузка формы
'загрузка классов
Set newCase = New clmCase           'Экземпляр класса "Дело
Set newDate = New clmCaseDate       'Экземпляр класса "Дата"
Set newCor = New clmHR_Official     'Экземпляр класса "Следователь"
Call initClass
    Dim str As String
    With Me
        .txtCsIndex.Width = 520.932
        .txtCsIndex.Left = 1123.585
        .Caption = "Создание нового документа"  '" Создание нового документа " & "ГМСЭ: " & .newExpert &   .newEF.categories & ")"
    End With
 Move (Screen.Width - Width) / 2, (Screen.Height - Height) / 3.7 '/ 17.7
End Sub
'
Private Sub Form_Unload(Cancel As Integer)
'Процедура выгрузки формы
    If MsgBox("Уверены?", vbYesNo, "Выход?") = vbYes Then
        Set newCase = Nothing       'Экземпляр класса "Дело"
        Set newDate = Nothing       'Экземпляр класса "Дата"
        Set newCor = Nothing        'Экземпляр класса "Следователь"
        Call termClass
    Else
        Cancel = 1
    End If
End Sub
'
Private Sub initClass()
'инициализация отдельных классов
    Set newEF = New clmExpertFindings   'Экземпляр класса "Заключение эксперта"
        newEF.categories = frmMDI.category
    Set newInjPr = New clmEntrants      'экземпляр класса "Потерпевший"
End Sub
'
Private Sub termClass()
'уничтожение отдельных эксземпляров классов
    Set newEF = Nothing         'Экземпляр класса "Заключение эксперта"
    Set newInjPr = Nothing      'экземпляр класса "Потерпевший"
    Set newDOC = Nothing  'Экземпляр класса "Создаваемые Документы"
End Sub
'
'+++++++++++++++++++++++++++++++++++++++ М Е Т О Д Ы  Ф О Р М Ы: ++++++++++++++++++++++++++++++++++++
'
Public Sub create_Analysis()
'Создание бланков направлений на исследования
Set newDOC.MyWdApp = New Word.Application 'экземпляр приложения
'1)создание папки для создаваемых документов:
Dim nameFolder As String
    nameFolder = frmAnalysis.Caption & "_" & frmNewResearch.newEF.getNumber
Dim nameDOC As String 'имя документа ИмяДокумента_Номер_Год
'2) создание новых документов и сохранение их в созданной папке:
Dim X As Object 'переменная контролов
    For Each X In frmAnalysis.Controls
        If TypeName(X) = "CheckBox" Then
            If X.Value = 1 Then
                With newDOC
                    nameDOC = X.Caption & "_" & frmNewResearch.newEF.getNumber
                    .dirAnalysis = .Create_mainFolders(.dirExpert, nameFolder)
                    Call .print_Blank(X.Caption, nameDOC, .dirAnalysis)
                End With
            End If
        End If
    Next X
End Sub
'
Public Sub create_Blank()
'Создание документов "схемы"
'1)создание папки для создаваемых документов:
Dim nameFolder As String
    nameFolder = frmDocList.Caption & "_" & frmNewResearch.newEF.getNumber
Dim nameDOC As String 'имя документа ИмяДокумента_Номер_Год
'2) создание новых документов и сохранение их в созданной папке:
Dim X As Object 'переменная контролов
    For Each X In frmDocList.Controls
        If TypeName(X) = "CheckBox" Then
            If X.Value = 1 Then
                With newDOC
                    nameDOC = X.Caption & "_" & frmNewResearch.newEF.getNumber
                    .dirDoc = .Create_mainFolders(.dirExpert, nameFolder)
                    Call .print_Blank(X.Caption, nameDOC, .dirDoc)
                End With
            End If
        End If
    Next X
newDOC.MyWdApp.Quit
Set newDOC.MyWdApp = Nothing
End Sub
'
Public Function create_legalGround() As String
'юридическое основание проведения экспертизы = на основании постановления/определения
'должность следователя + звание + отдел + ФИО + вынесенного Дата по материалам уголовного дела
    create_legalGround = newEF.printDefinition & newCor.print_Cor & _
    ", вынесенного " & newEF.rulingDate & newCase.printCsData & "."
Debug.Print "Юр.Основание = " & create_legalGround
End Function
'
Public Static Sub addExpert()
'1)скрытие формы:
Me.Visible = False
'Функция выбора эксперта из списка
Set SvcService = CreateObject("Svcsvc.Service") 'объект библиотеки svcsvc.dll
        newExpert = SvcService.SelectValue("Дубовский В.В." & vbCrLf & _
                                        "Жданович Э.Н." & vbCrLf & _
                                        "Козлов А.А." & vbCrLf & _
                                        "Кузьмичев С.В." & vbCrLf & _
                                        "Мицкевич Ю.В." & vbCrLf & _
                                        "Муха А.И." & vbCrLf & _
                                        "Павлович Д.С." & vbCrLf & _
                                        "Савчина Е.В." & vbCrLf & _
                                        "Самойлович М.В." & vbCrLf & _
                                        "Сосновский А.А." & vbCrLf & _
                                        "Терешко В.В.", _
                                        "Выберите эксперта", True)
                                        'Debug.Print " addExpert: " &  addExpert
Set SvcService = Nothing
End Sub
'
Private Sub txtEnabled()
'выключение отдельных полей
    Dim X As Object
    For Each X In Me.Controls
        If TypeName(X) = "TextBox" Then
            If X.name = "txtCsIndex" Or _
                X.name = "txtCsNum" Or _
                X.name = "txtCrPost" Or _
                X.name = "txtCrOprAr" Or _
                X.name = "txtCrRank" Or _
                X.name = "txtCrSurName" Or _
                X.name = "txtCrName" Or _
                X.name = "txtCrMidName" Then
                X.Enabled = True
            End If
        ElseIf TypeName(X) = "ComboBox" Then
            If X.name = "lstCrSex" Then
                X.Enabled = True
            End If
        End If
     Next X
End Sub
'
'+++++++++++++++++++++++++++++++++++++++ П О Л Я  Ф О Р М Ы: ++++++++++++++++++++++++++++++++++++++++++++++++++
'
Private Sub txtEnter(tmpObj As Object)
'изменение текстоых полей при получении фокуса
    With tmpObj
        .Text = ""
        .BackColor = frmColor.Светло_Желтый 'RGB(102, 102, 153) 'frmColor.Цвет1
        .ForeColor = frmColor.Черный
    End With
End Sub
'
 Private Sub txt_Exit(tmpObj As Object)
'изменение текстоых полей при потере фокуса
    With tmpObj
        .BackColor = frmColor.Коричневый
        If tmpObj.name = "txtCsIndex" Or _
            tmpObj.name = "txtCsNum" Or _
            tmpObj.name = "txtCrOprAr" Or _
            tmpObj.name = "txtInjPrBirthday" _
            Then
                .Text = "не указан"
        ElseIf tmpObj.name = "txtCsPrvDate" Or _
            tmpObj.name = "txtCrPost" Or _
            tmpObj.name = "txtCrSurName" Or _
            tmpObj.name = "txtInjPrSurName" Or _
            tmpObj.name = "txtInjPrDecease" Or _
            tmpObj.name = "txtInjPrAutopsy" _
            Then
                .Text = "не указана"
        ElseIf tmpObj.name = "txtCrName" Or _
            tmpObj.name = "txtCrMidName" Or _
            tmpObj.name = "txtInjPrName" Or _
            tmpObj.name = "txtInjPrMidName" _
            Then
                .Text = "не указано"
        Else
            .Text = "не указаны"
        End If
   End With
 End Sub
'
'1) Номер экспертизы
Private Static Sub txtDN_GotFocus()
'Номер экспертизы
    Call txtEnter(txtDN)
 End Sub
'
Private Static Sub txtDN_LostFocus()
'Номер экспертизы
    With txtDN
        Do
            If Not IsNumeric(.Text) Or Len(.Text) = 0 Then 'если не цифры или пустое поле
                Beep
                .BackColor = RGB(255, 0, 0) 'frmColor.Красный
                MsgBox "Cледует вводить цифры", vbCritical, Msg
                .Text = InputBox("Введите правильно номер экспертизы!", "Исправление ошибки ввода")
                 If VarType(.Text) = vbBoolean Then Exit Sub    ' нажата кнопка ОТМЕНА
            ElseIf CInt(.Text) <= 0 Then 'если меньше нуля или ноль
                 Beep
                .BackColor = RGB(255, 0, 0) 'frmColor.Красный
                MsgBox "Номер экспертизы должен быть больше нуля!", vbCritical, Msg
                .Text = InputBox("Введите правильно номер экспертизы!", "Исправление ошибки ввода")
'                 If VarType(.Text) = vbBoolean Then Exit Sub    ' нажата кнопка ОТМЕНА
            Else
                newEF.number = .Text
                .BackColor = RGB(200, 256, 200) 'frmColor.Салатовый
        Exit Do
            End If
        Loop
    End With
    
    With Me
        .Caption = " Создание нового документа №" & newEF.getFullNumber '"ГМСЭ: " & Me.newExpert &
        With .txtEFFirstDay
            .SetFocus 'передача фокуса
            Call txtEnter(Me.txtEFFirstDay)
            .Text = newDate.DateNow
        End With
    End With
Debug.Print "Номер экспертизы: №" & newEF.number
End Sub
'2) Дата начала экспертизы
Private Sub txtEFFirstDay_GotFocus()
'Дата начала экспертизы
    Call txtEnter(txtEFFirstDay)
    'Дата начала = текущая дата
    txtEFFirstDay.Text = newDate.DateNow
End Sub
'
Private Sub txtEFFirstDay_LostFocus()
'Дата начала экспертизы
    Dim tmp As Double, dt As Date, str As String
    If txtEFFirstDay.Text <> "" Then
        With newDate
            dt = .validateDate(.ExamDate(txtEFFirstDay.Text))
        End With
    Else 'если дата не введена:
            With txtEFFirstDay
                Do 'условие обязательного ввода даты:
                    If .Text = "" Then 'если  пустое поле
                        Beep
                    .BackColor = RGB(255, 0, 0) 'frmColor.Красный
                        MsgBox "Cледует ввести дату начала экспертизы!", vbCritical, Msg
                        .Text = InputBox("Введите правильно дату начала экспертизы!", "Исправление ошибки ввода", newDate.DateNow)
'                 If VarType(.Text) = vbBoolean Then Exit Sub    ' нажата кнопка ОТМЕНА
                    Else
                        dt = newDate.validateDate(newDate.ExamDate(.Text))
                        .BackColor = RGB(200, 256, 200) 'frmColor.Салатовый
                        Exit Do
                    End If
                Loop
            End With
    End If
    'с классом:
    With newEF
        'Дата начала
        .firstDay = dt
        'Срок
        .dueDate = newDate.getPeriod(.firstDay, 29)
    End With
    'с текстовым полем:
    With txtEFFirstDay
        .BackColor = frmColor.Салатовый
        .Text = newDate.dateToString(newEF.firstDay)
    End With
'Debug.Print "Дата начала: " & newEF.firstDay
'Debug.Print "Срок: " & newEF.dueDate
End Sub
'
'3)Категория экспертизы
Private Sub cboEFCategories_GotFocus()
'Категория экспертизы (первичная, дополнительная и т.д.)
    With cboEFCategories
        .Clear
        .BackColor = frmColor.Светло_Желтый
        .AddItem "первичной"
        .AddItem "комплексной"
        .AddItem "дополнительной"
        .AddItem "повторной"
        .AddItem "комиссионной"
    End With
End Sub
'
Private Sub cboEFCategories_LostFocus()
'Категория экспертизы (первичная, дополнительная и т.д.)
    With cboEFCategories
        If .Text = "" Then
            .Text = "первичной"
        End If
            newEF.categories = .Text
            .BackColor = frmColor.Салатовый 'RGB(200, 256, 200)
    End With
'Debug.Print "Категория экспертизы= ", newEF.categories
End Sub
'
'4)Основание проведения экспертизы
Private Sub cboCsDefinition_GotFocus()
'Основание проведения экспертизы
    With cboCsDefinition
        .Clear
        .BackColor = frmColor.Светло_Желтый
        .AddItem "постановления"
        .AddItem "определения"
        .AddItem "на платной основе"
    End With
End Sub
'
Private Sub cboCsDefinition_LostFocus()
'Основание проведения экспертизы
    With cboCsDefinition
        If .Text = "" Then
            .Text = "постановления"
        End If
        newEF.definition = .Text 'Присвоение переменым значений
        .BackColor = frmColor.Салатовый 'Прекращение посвечивания поля при потере им фокуса
    End With
'Debug.Print "Основание: " & newEF.definition
End Sub
'
'5) Дата вынесения постановления
Private Static Sub txtCsPrvDate_GotFocus()
'Дата вынесения постановления
    Call txtEnter(txtCsPrvDate)
    'Дата вынесения постановления = (текущая дата - 1):
    txtCsPrvDate.Text = CStr(newEF.firstDay - 1)
End Sub
'
Private Static Sub txtCsPrvDate_LostFocus()
'Дата вынесения постановления
 Dim tmp As Double, dt As Date
 With txtCsPrvDate
    If .Text = "" Then
        Call txt_Exit(txtCsPrvDate)
    Else
        With newDate
            dt = .validateDate(.ExamDate(txtCsPrvDate.Text)) 'newEF.rulingDate
            'сравнение дат: Дата начала > Дата вынесения постановления
            Do
                tmp = .compareDt(newEF.firstDay, dt)
                If tmp > 0 Then
                    MsgBox "Дата вынесения постановления больше даты начала экспертизы!", _
                            vbCritical, Msg
                    dt = .ExamDate(InputBox("Правильно введите дату вынесения постановления!", _
                                        "Ввод даты", newEF.firstDay - 1))
                Else
                    newEF.rulingDate = dt
                    Exit Do
                End If
            Loop
        End With
            'With txtCsPrvDate
            .BackColor = RGB(200, 256, 200) 'frmColor.Салатовый
            .Text = newDate.dateToString(newEF.rulingDate)
    End If
End With
'Debug.Print "Дата вынесения постановления: " & newEF.rulingDate
End Sub
'
'6) Категория дела
Private Sub cboCsCat_GotFocus()
'Категория дела
With cboCsCat
    .Clear
        .BackColor = frmColor.Светло_Желтый  '&HC0FFFF
'Добавить строки в комбинированное поле
        .AddItem "уголовного дела"
        .AddItem "проверки"
        .AddItem "административного дела"
    End With
End Sub
'
Private Sub cboCsCat_LostFocus()
'Категория дела
With cboCsCat
    If Len(.Text) = 0 Then
            .Text = "(категория дела не указана)"
           .BackColor = frmColor.Коричневый ' &HC0E0FF
    Else:
        If .Text = "проверки" Then
            txtCsIndex.Visible = True
            txtCsIndex.SetFocus
        Else:
            txtCsIndex.Visible = False
            txtCsNum.SetFocus
        End If
            txtCsNum.Visible = True
             newCase.category = .Text
            .BackColor = frmColor.Салатовый 'RGB(200, 256, 200)
    End If
    End With
'Debug.Print "Категория дела: " & newCase.category
End Sub
'
'7)Индекс Дела
Private Sub txtCsIndex_GotFocus()
'Индекс Дела
 Call txtEnter(txtCsIndex)
    With txtCsIndex
        .Width = 520.932
        .Left = 1123.585
    End With
End Sub
'
Private Sub txtCsIndex_LostFocus()
'Индекс Дела
    With txtCsIndex
        If Len(.Text) = 0 Then
            .Width = 1398.29
            .Left = 246.226
            Call txt_Exit(txtCsIndex)
        Else: newCase.index = StrConv(.Text, vbProperCase)
            .BackColor = frmColor.Салатовый 'RGB(200, 256, 200)
            .Text = newCase.index
        End If
    End With
'Debug.Print "Индекс Дела: ", newCase.index
End Sub
'
'8) Номер дела
Private Static Sub txtCsNum_GotFocus()
'Номер дела
    Call txtEnter(txtCsNum)
    txtCsNum.MaxLength = 11
End Sub
'
Private Static Sub txtCsNum_LostFocus()
'Номер дела CsNum
With txtCsNum
    If Len(.Text) = 0 Then
            Call txt_Exit(txtCsNum)
    Else
    '1) если не уголовное дело
        If newCase.category <> "уголовного дела" Then
            newCase.number = .Text
            .BackColor = RGB(200, 256, 200)
        Else '2) если уголовное дело:
             Do
                If Not IsNumeric(.Text) Or Len(.Text) = 0 Then 'anee ia oeo?u eee ionoia iiea
                   Beep
                    .BackColor = RGB(255, 0, 0)
                    MsgBox "Следует вводить цифры!", vbCritical, Msg
                        .Text = InputBox("Правильно введите номер уголовного дела!", "Исправление ошибки ввода данных!")
'                    If VarType(.Text) = vbBoolean Then Exit Do    'нажата кнопка "Отмена"
                    ElseIf Len(.Text) <> 11 Then '
                        MsgBox "Номер дела должен состоять из 11 цифр!", vbCritical, Msg
                        .Text = InputBox("Правильно введите  номер дела!", "Исправление ошибки ввода")
'                    If VarType(.Text) = vbBoolean Then Exit Do    'нажата кнопка "Отмена"
                   
                    End If
                   Exit Do
                Loop
        End If
    End If
    If .Text = "не указан" Then
        .BackColor = frmColor.Коричневый
    Else
        newCase.number = .Text
        .BackColor = RGB(200, 256, 200) 'frmColor.Салатовый
    End If
End With
'Debug.Print "Номер дела: №" & newCase.number
End Sub
'
'9)Должность следователя
Private Static Sub txtCrPost_GotFocus()
'Должность следователя
     Call txtEnter(txtCrPost)
     txtCrPost.Text = "старшего следователя"
End Sub
'
Private Static Sub txtCrPost_LostFocus()
'Должность следователя
    With txtCrPost
        If .Text = "" Then
           Call txt_Exit(txtCrPost)
        Else: newCor.post = .Text
            .BackColor = RGB(200, 256, 200) 'frmColor.Салатовый
        End If
    End With
'Debug.Print "Должность следователя= " & newCor.post
End Sub
'
'10) Отдел (следователя)
Private Static Sub txtCrOprAr_GotFocus()
'Отдел (следователя)
    Call txtEnter(txtCrOprAr)
    txtCrOprAr.Text = "(г. Минска) районного отдела Следственного комитета Республики Беларусь"
End Sub
'
Private Static Sub txtCrOprAr_LostFocus()
'Отдел (следователя)
    With txtCrOprAr
        If .Text = "" Then
           Call txt_Exit(txtCrOprAr)
        Else: newCor.conformation = .Text
            .BackColor = RGB(200, 256, 200) 'frmColor.Салатовый 'RGB(200, 256, 200)
        End If
    End With
'Debug.Print "Отдел (следователя): " & newCor.conformation
End Sub
'
'11) Звание следователя
Private Static Sub txtCrRank_GotFocus()
'Звание следователя
    Call txtEnter(txtCrRank)
    txtCrRank.Text = "старшего лейтенанта юстиции"
End Sub
'
Private Static Sub txtCrRank_LostFocus()
'Звание следователя
    With txtCrRank
        If .Text = "" Then
            Call txt_Exit(txtCrRank)
        Else: newCor.rank = .Text
            .BackColor = RGB(200, 256, 200) 'frmColor.Салатовый 'RGB(200, 256, 200)
        End If
    End With
'Debug.Print "Звание следователя: " & newCor.rank
End Sub
'
'12)Фамилия следователя
Private Static Sub txtCrSurName_GotFocus()
'Фамилия следователя
    Call txtEnter(txtCrSurName)
End Sub
'
Private Static Sub txtCrSurName_LostFocus()
'Фамилия следователя
    With txtCrSurName
        If .Text = "" Then
            Call txt_Exit(txtCrSurName)
        Else: newCor.surName = StrConv(.Text, vbProperCase)
            .BackColor = RGB(200, 256, 200)
            .Text = newCor.surName
        End If
    End With
'Debug.Print "Фамилия следователя: " & newCor.surName
End Sub
'
'13) Имя следователя
Private Static Sub txtCrName_GotFocus()
'Имя следователя
    Call txtEnter(txtCrName)
End Sub
'
Private Static Sub txtCrName_LostFocus()
'Имя следователя
    With txtCrName
        If .Text = "" Then
            Call txt_Exit(txtCrName)
        Else
    'постановка точки после инициала
            If Len(txtCrName.Text) = 1 Then
                newCor.name = newCor.create_Initials(.Text)
            Else: newCor.name = StrConv(.Text, vbProperCase)
            End If
        .Text = newCor.name
        .BackColor = RGB(200, 256, 200) 'frmColor.Салатовый 'Прекращение посвечивания поля при потере им фокуса
        End If
    End With
'Debug.Print "Имя следователя: " & newCor.name
End Sub
'
'14)Отчество следователя
Private Static Sub txtCrMidName_GotFocus()
'Отчество следователя
   Call txtEnter(txtCrMidName)
End Sub
'
Private Static Sub txtCrMidName_LostFocus()
'Отчество следователя
    With txtCrMidName
        If .Text = "" Then
           Call txt_Exit(txtCrMidName)
        Else
            If Len(txtCrMidName.Text) = 1 Then
                newCor.midName = newCor.create_Initials(.Text)
            Else: newCor.midName = StrConv(.Text, vbProperCase)
            End If
            .BackColor = RGB(200, 256, 200) 'frmColor.Салатовый
            .Text = newCor.midName
        End If
    End With
'Debug.Print "Отчество следователя: " & newCor.midName
End Sub
'
'15) Пол следователя
Private Sub lstCrSex_GotFocus()
'Пол следователя
    With lstCrSex
        .Clear
        .AddItem "муж."
        .AddItem "жен."
        .AddItem "пол не указан"
        .BackColor = &HC0FFFF
    End With
End Sub
'
Private Sub lstCrSex_LostFocus()
'Пол следователя
    With lstCrSex
        newCor.sex = .Text
        .BackColor = frmColor.Салатовый
    End With
Debug.Print "Пол следователя: " & newCor.sex
End Sub

'16)Фамилия потерпевшего
Private Static Sub txtInjPrSurName_GotFocus()
'Фамилия потерпевшего
   Call txtEnter(txtInjPrSurName)
End Sub
'
Private Static Sub txtInjPrSurName_LostFocus()
'Фамилия потерпевшего
    With txtInjPrSurName
        If .Text = "" Then
             Call txt_Exit(txtInjPrSurName)
        Else: newInjPr.surName = StrConv(.Text, vbProperCase)
            .BackColor = RGB(200, 256, 200) 'frmColor.Салатовый
            .Text = newInjPr.surName
        End If
    End With
'Debug.Print "Фамилия потерпевшего: " & newInjPr.surName
End Sub
'
'17)
Private Static Sub txtInjPrName_GotFocus()
'Имя потерпевшего
   Call txtEnter(txtInjPrName)
End Sub
'
Private Static Sub txtInjPrName_LostFocus()
'Имя потерпевшего
    With txtInjPrName
        If .Text = "" Then
             Call txt_Exit(txtInjPrName)
        Else
            If Len(.Text) = 1 Then
                newInjPr.name = newInjPr.create_Initials(.Text)
            Else: newInjPr.name = StrConv(.Text, vbProperCase)
            End If
        .Text = newInjPr.name
        .BackColor = RGB(200, 256, 200) 'frmColor.Салатовый
        End If
    End With
'Debug.Print "Имя потерпевшего: " & newInjPr.name
End Sub
'
'18)
Private Static Sub txtInjPrMidName_GotFocus()
'Отчество потерпевшего
    Call txtEnter(txtInjPrMidName)
End Sub
'
Private Static Sub txtInjPrMidName_LostFocus()
'Отчество потерпевшего
    With txtInjPrMidName
        If .Text = "" Then
             Call txt_Exit(txtInjPrMidName)
        Else
            If Len(.Text) = 1 Then
                newInjPr.midName = newInjPr.create_Initials(.Text)
            Else: newInjPr.midName = StrConv(.Text, vbProperCase)
            End If
        .BackColor = RGB(200, 256, 200)
        .Text = newInjPr.midName
        End If
    End With
'Debug.Print "Отчество потерпевшего: " & newInjPr.midName
End Sub
'
'19)
Private Sub lstInjPrSex_GotFocus()
 'Пол потерпевшего
    With lstInjPrSex
        .Clear
        .AddItem "муж."
        .AddItem "жен."
        .AddItem "пол не указан"
        .BackColor = &HC0FFFF
    End With
End Sub
'
Private Sub lstInjPrSex_LostFocus()
'Пол потерпевшего
    With lstInjPrSex
        newInjPr.sex = .Text
        .BackColor = RGB(200, 256, 200)
    End With
'Debug.Print "Пол потерпевшего: " & newInjPr.sex
End Sub
'
'20)'Год рождения потерпевшего
Private Static Sub txtInjPrBirthday_GotFocus()
'Год рождения потерпевшего
   Call txtEnter(txtInjPrBirthday)
End Sub
'
Private Static Sub txtInjPrBirthday_LostFocus()
'Год рождения потерпевшего
    With txtInjPrBirthday
        If .Text = "" Then
             Call txt_Exit(txtInjPrBirthday)
        Else: newInjPr.birthday = .Text
            .BackColor = RGB(200, 256, 200)
        End If
    End With
'Debug.Print "Год рождения потерпевшего: " & newInjPr.birthday
End Sub
'
'21) Дата смерти
Private Sub txtInjPrDecease_GotFocus()
'Дата смерти
   Call txtEnter(txtInjPrDecease)
   'дата смерти = дата вынесения постановления
   txtInjPrDecease.Text = newEF.rulingDate
End Sub

Private Sub txtInjPrDecease_LostFocus()
'Дата смерти
 Dim tmp As Double, dt As Date ' str As String
    If txtInjPrDecease.Text = "" Then
        Call txt_Exit(txtInjPrDecease)
    Else
        With newDate
            dt = .validateDate(.ExamDate(txtInjPrDecease.Text))
            'сравнение дат: Дата смерти <= Дата вынесения постановления
                '                       <= Дата начала экспертизы
                '                       <= Дата вскрытия
            Do
                tmp = .compareDt(newEF.rulingDate, dt)
                If tmp > 0 Then
                    MsgBox "Дата смерти больше даты вынесения постановления!", _
                            vbCritical, Msg
                    dt = .ExamDate(InputBox("Правильно введите дату смерти!", _
                                        "Ввод даты", newEF.rulingDate))
                Else
                    newInjPr.decease = dt
                    Exit Do
                End If
            Loop
        End With
        'с текстовым полем:
        With txtInjPrDecease
            .BackColor = RGB(200, 256, 200)
            .Text = newDate.dateToString(newInjPr.decease)
        End With
    End If
    'передача фокуса следующему полю
    With Me.txtInjPrAutopsy
        .SetFocus
        Call txtEnter(Me.txtInjPrAutopsy)
        .Text = newEF.firstDay
    End With
'Debug.Print "Дата смерти: " & newInjPr.decease
End Sub
''
'22)Дата вскрытия
Private Sub txtInjPrAutopsy_GotFocus()
'Дата вскрытия
   Call txtEnter(txtInjPrAutopsy)
   'Дата вскрытия = Дата начала
   txtInjPrAutopsy.Text = newEF.firstDay
End Sub
'
Private Sub txtInjPrAutopsy_LostFocus()
'Дата вскрытия
 Dim tmp As Double, dt As Date, str As String
    If txtInjPrAutopsy.Text = "" Then
        Call txt_Exit(txtInjPrAutopsy)
    Else
        With newDate
            dt = .validateDate(.ExamDate(txtInjPrAutopsy.Text))
            'сравнение дат: Дата вскрытия >= Дата вынесения постановления
                '                         >= Дата начала экспертизы
                '                         >= Дата смерти
            Do
                tmp = .compareDt(newEF.firstDay, dt)
                If tmp < 0 Then
                    MsgBox "Дата вскрытия меньше даты начала экспертизы!", _
                            vbCritical, Msg
                    dt = .ExamDate(InputBox("Правильно введите дату вскрытия!", _
                                        "Ввод даты", newEF.firstDay))
                Else
                    newInjPr.autopsyDate = dt
                    Exit Do
                End If
            Loop
        End With
        With txtInjPrAutopsy
            .BackColor = RGB(200, 256, 200)
            .Text = newDate.dateToString(newInjPr.autopsyDate)
        End With
    End If
    'передача фокуса полю
    With Me.txtFactCase
        Call txtEnter(txtFactCase)
        .Text = newDate.dateToString(newEF.rulingDate) & " труп " & newInjPr.create_InitialslName
    End With
'Debug.Print "Дата вскрытия: " & newInjPr.autopsyDate
'On Error Resume Next ' Отключаем ошибки
End Sub
'
'23)
Private Sub txtFactCase_GotFocus()
'Краткие обстоятельства дела
    Call txtEnter(txtFactCase)
    txtFactCase.Text = newDate.dateToString(newEF.rulingDate) & _
                        " труп " & newInjPr.print_InjPersons
'&H00FFE8DF&
'&H00CDEEEF&
End Sub
'
Private Sub txtFactCase_LostFocus()
 'Краткие обстоятельства дела
    If txtInjPrDecease.Text = "" Then
        Call txt_Exit(txtFactCase)
    Else
        With txtInjPrDecease
            newCase.crimCondition = .Text
            .BackColor = RGB(200, 256, 200)
        End With
    End If
'Debug.Print "Краткие обстоятельства дела: " & newCase.crimCondition
End Sub
'
'+++++++++++++++++++++++++++++++++++++++ К Н О П К И: ++++++++++++++++++++++++++++++++++++++++++++++++++
Private Sub cmdOK1_Click()
'кнопка "ОК"
'1)Добавляем эксперта:
 Call addExpert
 Me.Caption = "ГМСЭ: " & Me.newExpert & Chr(32) & " Создание нового документа №" & newEF.getFullNumber
'2) изменение директории на выбранную
Set newDOC = New clmCreateDoc 'Экземпляр класса "Создаваемые Документы"
Dim tmpname As String
    tmpname = Me.newExpert & "_" & Me.newEF.getNumber
Dim tmp As String
    With newDOC
        .dirExpert = .Create_mainFolders(frmMDI.newRoot, tmpname)
        tmpname = "Фото_" & Me.newEF.getNumber
        tmp = .Create_mainFolders(.dirExpert, tmpname)
    End With
    frmAnalysis.Show
    frmNewResearch.Visible = False  'скрытие формы "Новое исследование"
End Sub
'
Private Sub cmdCancel1_Click()
'кнопка "Отмена"
    Unload Me
End Sub
'
Private Sub txtDisabled()
'выключение отдельных полей
    Dim X As Object
    For Each X In Me.Controls
        If TypeName(X) = "TextBox" Then
            If X.name = "txtCsIndex" Or _
                X.name = "txtCsNum" Or _
                X.name = "txtCrPost" Or _
                X.name = "txtCrOprAr" Or _
                X.name = "txtCrRank" Or _
                X.name = "txtCrSurName" Or _
                X.name = "txtCrName" Or _
                X.name = "txtCrMidName" Then
                X.Enabled = False
            Else: X.Text = ""
            End If
        ElseIf TypeName(X) = "ComboBox" Then
            If X.name = "cboCsCat" Or _
                X.name = "lstCrSex" Then
                X.Enabled = False
            Else: X.Clear
            End If
        End If
     Next X
End Sub

'
Private Sub cmdEraseData_Click()
'кнопка "Очистить"
Dim X As Object
    For Each X In Me.Controls
        If TypeName(X) = "TextBox" Then
            X.Text = ""
            If X.name = "txtDN" Or _
                X.name = "txtEFFirstDay" Or _
                X.name = "txtCsPrvDate" Or _
                X.name = "txtCsIndex" Or _
                X.name = "txtCsNum" Then
                X.BackColor = &HFEF7F1
            ElseIf X.name = "txtCrPost" Or _
                X.name = "txtCrOprAr" Or _
                X.name = "txtCrRank" Or _
                X.name = "txtCrSurName" Or _
                X.name = "txtCrName" Or _
                X.name = "txtCrMidName" Then
                X.BackColor = &HFFE8DF
            ElseIf X.name = "txtInjPrSurName" Or _
                X.name = "txtInjPrName" Or _
                X.name = "txtInjPrMidName" Or _
                X.name = "txtInjPrBirthday" Or _
                X.name = "txtInjPrDecease" Or _
                X.name = "txtInjPrAutopsy" Then
                X.BackColor = &HF0FBE1
            End If
        ElseIf TypeName(X) = "ComboBox" Then
            X.Clear
            If X.name = "cboEFCategories" Or _
                X.name = "cboCsDefinition" Or _
                X.name = "cboCsCat" Then
                X.BackColor = &HFEF7F1
            ElseIf X.name = "lstCrSex" Then
                X.BackColor = &HFFE8DF
            ElseIf X.name = "lstInjPrSex" Then
                X.BackColor = &HF0FBE1
            End If
        End If
         X.Enabled = True
    Next X
Call txtEnabled
With Me
    .Caption = "Создание нового документа"
    .txtDN.SetFocus
End With
End Sub
'
Private Sub cmdSaveCreate_Click()
'процедура нажатия кнопки "Cоздать с таким же основанием"
    Call txtDisabled 'выключение нужных и заполенных полей
    Call termClass 'Уничтожение существующих экземпляров классов и освобождение памяти
    Call initClass 'Создание новых экземпляров классов
'Изменение надписи формы
With Me
    .Caption = "ГМСЭ: " & .newExpert & " Создание нового документа"
    .txtDN.SetFocus
End With
End Sub
'
