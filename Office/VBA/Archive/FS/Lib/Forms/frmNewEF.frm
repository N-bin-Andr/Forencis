VERSION 5.00
Begin VB.Form frmNewEF 
   BackColor       =   &H00BBC0AC&
   ClientHeight    =   8460
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12075
   LinkTopic       =   "Form1"
   ScaleHeight     =   11067.75
   ScaleLeft       =   100
   ScaleMode       =   0  'Пользовательское
   ScaleWidth      =   11567.52
   Begin VB.CommandButton cmdSaveCreate 
      BackColor       =   &H00FFC0C0&
      Caption         =   "&+ +"
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
      Left            =   9720
      Style           =   1  'Graphical
      TabIndex        =   31
      Top             =   7560
      Width           =   615
   End
   Begin VB.ComboBox cboCsCat 
      BackColor       =   &H00C0CFC5&
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
      Left            =   1680
      TabIndex        =   5
      Top             =   1920
      Width           =   4095
   End
   Begin VB.TextBox txtCsNum 
      Alignment       =   2  'Центровка
      BackColor       =   &H00C0CFC5&
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
      Left            =   3120
      TabIndex        =   7
      Top             =   2640
      Width           =   2655
   End
   Begin VB.TextBox txtCsIndex 
      Alignment       =   2  'Центровка
      BackColor       =   &H00C0CFC5&
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
      Left            =   1680
      TabIndex        =   6
      Top             =   2640
      Width           =   975
   End
   Begin VB.TextBox txtCsPrvDate 
      Alignment       =   2  'Центровка
      BackColor       =   &H00C0CFC5&
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
      Left            =   4440
      TabIndex        =   4
      Top             =   1380
      Width           =   1335
   End
   Begin VB.ComboBox cboCsDefinition 
      BackColor       =   &H00C0CFC5&
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
      Left            =   1680
      TabIndex        =   3
      Top             =   1380
      Width           =   2295
   End
   Begin VB.ComboBox cboEFCategories 
      BackColor       =   &H00C0CFC5&
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
      Left            =   2160
      TabIndex        =   2
      Top             =   795
      Width           =   3615
   End
   Begin VB.TextBox txtEFFirstDay 
      Alignment       =   2  'Центровка
      BackColor       =   &H00C0CFC5&
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
      Left            =   4080
      TabIndex        =   1
      Top             =   180
      Width           =   1695
   End
   Begin VB.TextBox txtDN 
      Alignment       =   2  'Центровка
      BackColor       =   &H00C0CFC5&
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
      Left            =   1200
      TabIndex        =   0
      Top             =   180
      Width           =   1215
   End
   Begin VB.CommandButton cmdEraseData 
      BackColor       =   &H008080FF&
      Caption         =   "&Очистить"
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
      Left            =   6120
      MaskColor       =   &H8000000F&
      Style           =   1  'Graphical
      TabIndex        =   30
      Top             =   7560
      Width           =   1575
   End
   Begin VB.Frame fraExpert 
      BackColor       =   &H00DDEADB&
      Caption         =   "Направление общего  эксперта"
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
      Height          =   4065
      Left            =   6120
      TabIndex        =   47
      Top             =   3240
      Width           =   5535
      Begin VB.TextBox txtInjPrAutopsy 
         Alignment       =   2  'Центровка
         BackColor       =   &H00EBFFEA&
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   406
         Left            =   3420
         TabIndex        =   24
         Top             =   1920
         Width           =   1935
      End
      Begin VB.TextBox txtInjPrDecease 
         Alignment       =   2  'Центровка
         BackColor       =   &H00EBFFEA&
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
         Left            =   3420
         TabIndex        =   23
         Top             =   1440
         Width           =   1935
      End
      Begin VB.TextBox txtAEFNum 
         Alignment       =   2  'Центровка
         BackColor       =   &H00EBFFEA&
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
         Left            =   2640
         TabIndex        =   21
         Top             =   960
         Width           =   735
      End
      Begin VB.TextBox txtAutopsyDate 
         Alignment       =   2  'Центровка
         BackColor       =   &H00EBFFEA&
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
         Left            =   3780
         TabIndex        =   22
         Top             =   960
         Width           =   1575
      End
      Begin VB.TextBox txtAEFcode 
         BackColor       =   &H00EBFFEA&
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   12
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   960
         TabIndex        =   20
         Top             =   960
         Width           =   1575
      End
      Begin VB.TextBox txtAEMidName 
         BackColor       =   &H00EBFFEA&
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
         Left            =   1500
         TabIndex        =   27
         Top             =   3360
         Width           =   3855
      End
      Begin VB.TextBox txtAEName 
         BackColor       =   &H00EBFFEA&
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
         Left            =   1500
         TabIndex        =   26
         Top             =   2880
         Width           =   3855
      End
      Begin VB.TextBox txtAESurName 
         BackColor       =   &H00EBFFEA&
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
         Left            =   1500
         TabIndex        =   25
         Top             =   2400
         Width           =   3855
      End
      Begin VB.CheckBox chkAddExData 
         BackColor       =   &H00DDEADB&
         Caption         =   "&Добавить данные"
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   12
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   400
         Left            =   2830
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   360
         Width           =   2535
      End
      Begin VB.Label lblInjPrAutopsy 
         BackColor       =   &H00DDEADB&
         Caption         =   "Дата вскрытия"
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   12
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   1440
         TabIndex        =   62
         Top             =   1920
         Width           =   1935
      End
      Begin VB.Label lblInjPDeceaseDate 
         BackColor       =   &H00DDEADB&
         Caption         =   "Дата смерти"
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   12
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   1440
         TabIndex        =   61
         Top             =   1560
         Width           =   1935
      End
      Begin VB.Label lblAENum 
         BackColor       =   &H00DDEADB&
         Caption         =   "СМЭ №"
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   12
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   120
         TabIndex        =   60
         Top             =   1020
         Width           =   855
      End
      Begin VB.Label lblAutopsyDate 
         BackColor       =   &H00DDEADB&
         Caption         =   "от"
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   12
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   3360
         TabIndex        =   59
         Top             =   1020
         Width           =   345
      End
      Begin VB.Label Label1 
         BackColor       =   &H00DDEADB&
         Caption         =   "/"
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   12
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2520
         TabIndex        =   58
         Top             =   960
         Width           =   135
      End
      Begin VB.Label lblAEMidName 
         BackColor       =   &H00DDEADB&
         Caption         =   "Отчество"
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   12
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   120
         TabIndex        =   50
         Top             =   3420
         Width           =   1215
      End
      Begin VB.Label lblAEName 
         BackColor       =   &H00DDEADB&
         Caption         =   "Имя"
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   12
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   120
         TabIndex        =   49
         Top             =   2940
         Width           =   495
      End
      Begin VB.Label lblAESurName 
         BackColor       =   &H00DDEADB&
         Caption         =   "Фамилия"
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   12
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   120
         TabIndex        =   48
         Top             =   2460
         Width           =   1215
      End
   End
   Begin VB.Frame fraInjuredPerson 
      BackColor       =   &H00D1CFA1&
      Caption         =   "Потерпевший"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12.75
         Charset         =   204
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808000&
      Height          =   3015
      Left            =   6120
      TabIndex        =   39
      Top             =   120
      Width           =   5535
      Begin VB.ListBox lstInjPrSex 
         BackColor       =   &H00F0FBE1&
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
         Left            =   3960
         TabIndex        =   18
         Top             =   2520
         Width           =   1335
      End
      Begin VB.TextBox txtInjPrBirthday 
         BackColor       =   &H00F0FBE1&
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
         Left            =   2160
         TabIndex        =   17
         Top             =   1920
         Width           =   3135
      End
      Begin VB.TextBox txtInjPrMidName 
         BackColor       =   &H00F0FBE1&
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
         TabIndex        =   16
         Top             =   1440
         Width           =   3380
      End
      Begin VB.TextBox txtInjPrName 
         BackColor       =   &H00F0FBE1&
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
         TabIndex        =   15
         Top             =   960
         Width           =   3380
      End
      Begin VB.TextBox txtInjPrSurName 
         BackColor       =   &H00F0FBE1&
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
         TabIndex        =   14
         Top             =   480
         Width           =   3380
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Центровка
         BackColor       =   &H00D1CFA1&
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
         ForeColor       =   &H00004000&
         Height          =   285
         Left            =   3360
         TabIndex        =   54
         Top             =   2550
         Width           =   495
      End
      Begin VB.Label lblInjPrSurName 
         BackColor       =   &H00D1CFA1&
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
         ForeColor       =   &H00004000&
         Height          =   285
         Left            =   240
         TabIndex        =   43
         Top             =   541
         Width           =   1695
      End
      Begin VB.Label lblInjPrName 
         BackColor       =   &H00D1CFA1&
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
         ForeColor       =   &H00004000&
         Height          =   285
         Left            =   240
         TabIndex        =   42
         Top             =   1021
         Width           =   1695
      End
      Begin VB.Label lblInjPrMidName 
         BackColor       =   &H00D1CFA1&
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
         ForeColor       =   &H00004000&
         Height          =   285
         Left            =   240
         TabIndex        =   41
         Top             =   1501
         Width           =   1695
      End
      Begin VB.Label lblInjPrBirthday 
         BackColor       =   &H00D1CFA1&
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
         ForeColor       =   &H00004000&
         Height          =   285
         Left            =   240
         TabIndex        =   40
         Top             =   2040
         Width           =   1935
      End
   End
   Begin VB.Frame fraCoroner 
      BackColor       =   &H00CFC2AC&
      Caption         =   "Следователь"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12.75
         Charset         =   204
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808000&
      Height          =   4695
      Left            =   240
      TabIndex        =   32
      Top             =   3360
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
         TabIndex        =   57
         Top             =   4080
         Width           =   1095
      End
      Begin VB.TextBox txtCrMidName 
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
         Left            =   1680
         TabIndex        =   13
         Top             =   3480
         Width           =   3615
      End
      Begin VB.TextBox txtCrName 
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
         Left            =   1680
         TabIndex        =   12
         Top             =   3000
         Width           =   3615
      End
      Begin VB.TextBox txtCrSurName 
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
         Left            =   1680
         TabIndex        =   11
         Top             =   2520
         Width           =   3615
      End
      Begin VB.TextBox txtCrRank 
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
         Left            =   1680
         MultiLine       =   -1  'True
         TabIndex        =   10
         Top             =   2040
         Width           =   3615
      End
      Begin VB.TextBox txtCrOprAr 
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
         Height          =   1080
         Left            =   1680
         MultiLine       =   -1  'True
         TabIndex        =   9
         Top             =   840
         Width           =   3615
      End
      Begin VB.TextBox txtCrPost 
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
         Left            =   1680
         TabIndex        =   8
         Top             =   360
         Width           =   3615
      End
      Begin VB.Label lblCrSex 
         Alignment       =   2  'Центровка
         BackColor       =   &H00CFC2AC&
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
         Left            =   3480
         TabIndex        =   56
         Top             =   4100
         Width           =   735
      End
      Begin VB.Label lblCrPost 
         BackColor       =   &H00CFC2AC&
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
         Height          =   285
         Left            =   240
         TabIndex        =   38
         Top             =   420
         Width           =   1455
      End
      Begin VB.Label lblPIOprAr 
         BackColor       =   &H00CFC2AC&
         Caption         =   "                        Отдел"
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
         Height          =   915
         Left            =   240
         TabIndex        =   37
         Top             =   960
         Width           =   1455
      End
      Begin VB.Label lblCrRank 
         BackColor       =   &H00CFC2AC&
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
         Height          =   285
         Left            =   240
         TabIndex        =   36
         Top             =   2101
         Width           =   855
      End
      Begin VB.Label lblCrSurName 
         BackColor       =   &H00CFC2AC&
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
         Height          =   285
         Left            =   240
         TabIndex        =   35
         Top             =   2580
         Width           =   1215
      End
      Begin VB.Label lblCrMidName 
         BackColor       =   &H00CFC2AC&
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
         Height          =   285
         Left            =   240
         TabIndex        =   34
         Top             =   3540
         Width           =   1455
      End
      Begin VB.Label lblCrName 
         BackColor       =   &H00CFC2AC&
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
         Height          =   285
         Left            =   240
         TabIndex        =   33
         Top             =   3061
         Width           =   975
      End
   End
   Begin VB.CommandButton cmdOK1 
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
      Left            =   10680
      Style           =   1  'Graphical
      TabIndex        =   28
      Top             =   7560
      Width           =   975
   End
   Begin VB.CommandButton cmdCancel1 
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
      Left            =   8040
      Style           =   1  'Graphical
      TabIndex        =   29
      Top             =   7560
      Width           =   1335
   End
   Begin VB.Label lblCsNum 
      Alignment       =   2  'Центровка
      BackColor       =   &H00BBC0AC&
      Caption         =   "№"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   2760
      TabIndex        =   55
      Top             =   2700
      Width           =   375
   End
   Begin VB.Label Label3 
      BackColor       =   &H00BBC0AC&
      Caption         =   "Основание:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   240
      TabIndex        =   53
      Top             =   1440
      Width           =   1215
   End
   Begin VB.Label lblEFCategories 
      BackColor       =   &H00BBC0AC&
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
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   240
      TabIndex        =   52
      Top             =   840
      Width           =   1935
   End
   Begin VB.Label lblProvDate 
      Alignment       =   2  'Центровка
      BackColor       =   &H00BBC0AC&
      Caption         =   "от"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   3960
      TabIndex        =   51
      Top             =   1440
      Width           =   435
   End
   Begin VB.Label lblCrResearchN 
      BackColor       =   &H00BBC0AC&
      Caption         =   "№ 5.2/"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   240
      TabIndex        =   46
      Top             =   240
      Width           =   855
   End
   Begin VB.Label lblCrimCase 
      BackColor       =   &H00BBC0AC&
      Caption         =   "Материалы"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   240
      TabIndex        =   45
      Top             =   2040
      Width           =   1455
   End
   Begin VB.Label lblCrResFirstDay 
      Alignment       =   2  'Центровка
      BackColor       =   &H00BBC0AC&
      Caption         =   "Дата начала"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   2520
      TabIndex        =   44
      Top             =   225
      Width           =   1575
   End
End
Attribute VB_Name = "frmNewEF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Форма "frmNewEF"
'Ввод основных данных для записи в базу данных и для создания новых документов
'Дата создания: 01.06.2016
'@version 1.0.0
'@author Andr.Nab.n@gmail.com
Option Explicit
'Объявление Классов
Public newDate As clmCaseDate      'Экземпляр класса "Дата"
Public newDOC As clmCreateDoc      'Экземпляр класса "Создаваемые Документы"
Public SvcService As Object        'Экземпляр библиотеки svcsvc.dll
'Формы
Public newEvForm As Form           'Экземпляр frmEvidences

'Clases
Public newCase As clmCase          'Экземпляр класса "Дело"
Public newCor As clmHR_Official    'Экземпляр класса "Следователь" Coroner
Public newExpert As String          'переменная "Эксперт"
Public newExperts As clmHR_Official  'экземпляр класса "Эксперты" (криминалисты/биологи и т.д.)
Public newAExperts As clmHR_Official  'экземпляр класса "общие эксперты"
Public newEF As clmExpertFindings  'Экземпляр класса "Заключение эксперта"
Public newAEF As clmExpertFindings 'Экземпляр класса "Заключение общего эксперта"
Public newInjPr As clmEntrants     'Экземпляр класса "Потерпевший"InjuredPerson
Public newBox As clmEvBox          'Экземпляр класса "коробка с ВД"
Public allEvSumCounter As clmCounter 'Экземпляр класса счетчик общей суммы ВД

'Key
Private mvarstrBoxKey As String 'ключ для упаковок с ВД ="BX****"
Private mvarstrEvKey As String  'ключ для вещдоков (ВД)= "EV****"
Private mvarstrKey As String    'общий ключ = "BX****" + "EV****" ("BX****EV****")
'коллекции
Public colBoxes As New Collection      'коллекция упаковок с вещдоками
Public colExperts As New Collection    ' коллекция экспертов
'Ошибки:
Const Msg As String = "Ошибка ввода данных!"
'Constants
Private Const BOX As String = "BX"  'префикс ключа для коробок
Private Const EVID As String = "EV" 'префикс ключа для вещдоков.
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
        .Width = 12000
        .Height = 9000
        .txtCsIndex.Width = 520.932
        .txtCsIndex.Left = 1123.585
        .Caption = "Создание нового документа " & "(" & mdPrintDoc.DocCat & ")"
        '& "ГМСЭ: " & .newExpert '" Создание нового документа " & "ГМСЭ: " & .newExpert & .newEF.categories & ")"
        'блокировка полей "Направление эксперта"
        .chkAddExData.Value = 0
        Call AE_Disabled
    End With
 Move (Screen.Width - Width) / 2, (Screen.Height - Height) / 3.7 '/ 17.7
End Sub
'
Private Sub Form_Unload(Cancel As Integer)
'Процедура выгрузки формы
'    If MsgBox("Уверены?", vbYesNo, "Выход?") = vbYes Then
        Set newCase = Nothing       'Экземпляр класса "Дело"
        Set newDate = Nothing       'Экземпляр класса "Дата"
        Set newCor = Nothing        'Экземпляр класса "Следователь"
        Call termClass
'    Else
'        Cancel = 1
'    End If
End Sub
'
Private Sub initClass()
'инициализация отдельных классов
    Set newEF = New clmExpertFindings   'Экземпляр класса "Заключение эксперта"
        newEF.evidensCategories = mdPrintDoc.DocCat 'категория экспертизы по ВД
    Set newAEF = New clmExpertFindings   'Экземпляр класса "Заключение общего эксперта"
'   Esperts
    Set newExperts = New clmHR_Official  'экземпляр класса "эксперты"
    Set newAExperts = New clmHR_Official  'экземпляр класса "общие эксперты"
    Set colExperts = New Collection      'экзепляр коллекции "Эксперты"
'   InjPr
    Set newInjPr = New clmEntrants    'экземпляр класса "Потерпевший"
'   ВД
    Set colBoxes = New Collection     'коллекция упаковок с вещдоками
    Set newBox = New clmEvBox         'Экземпляр класса "коробка с ВД"
'Counter
    Set allEvSumCounter = New clmCounter  'Экземпляр класса счетчик общей суммы ВД
         With allEvSumCounter
            .name = "общая сумма ВД"
'            Debug.Print .name & Chr(32) & .getTale
          End With
    Set mdPrintDoc.colTextDoc = Nothing
    Call addExpert
End Sub
'
Private Sub termClass()
'уничтожение отдельных эксземпляров классов
    Set newEF = Nothing         'Экземпляр класса "Заключение эксперта"
    Set newAEF = Nothing        'Экземпляр класса "Заключение общего эксперта"
    Set newExperts = Nothing    'экземпляр класса "эксперты"
    Set newAExperts = Nothing   'экземпляр класса "общие эксперты"
    Set colExperts = Nothing    'экзепляр коллекции "Эксперты"
    Set newInjPr = Nothing      'экземпляр класса "Потерпевший"
'   ВД
    Set colBoxes = Nothing      'коллекция упаковок с вещдоками
    Set newBox = Nothing        'Экземпляр класса "коробка с ВД"
    Set allEvSumCounter = Nothing  'Экземпляр класса счетчик общей суммы ВД
'    Set newDOC = Nothing       'Экземпляр класса "Создаваемые Документы"
End Sub
'
'+++++++++++++++++++++++++++++++++++++++ М Е Т О Д Ы  Ф О Р М Ы: ++++++++++++++++++++++++++++++++++++
'1) вызов диалогового окна
Public Sub show_MsgCreateNewBox()
'вызов диалогового окна "Создать новую коробку с ВД?"
    If MsgBox("Создать новую упаковку с объектами (ВД)?", vbYesNo) = vbYes Then
        Call Open_EvidForm  'Процедура открытия новой формы "Вещественные доказательства"
'       открытие формы frmEvidences
        
    Else
        MsgBox "Ввод упаковок с объектами (ВД) завершен!", vbOKOnly
         mdPrintDoc.colExcelData.Add Me.allEvSumCounter.getTale, "J" 'для работы с Excel
            Debug.Print "Запись Сумма ВД в коллекции colExcelData -> " & mdPrintDoc.colExcelData.Item("J")
        Call Open_DocListForm
    End If
End Sub
'
'2) создание новой коробки
Public Sub addNewBox()
  'процедура создания новой коробки с ВД
    With Me
        strBoxKey = BOX & CStr(Format(.colBoxes.Count + 1, "#0000"))
        Set .newBox = New clmEvBox
        .newBox.strBxName = .strBoxKey
    End With
End Sub
'
Public Sub Open_EvidForm()
'Процедура открытия новой формы "Вещественные доказательства"
Set newEvForm = New frmEvidences
    With newEvForm
        .Show
    End With
Me.Visible = False
End Sub
'
Public Sub Open_DocListForm()
'Процедура открытия формы frmDocuments "Перечень документов"
    Dim frmD As Form
'    Set frmD = New frmDocuments
        With frmDocuments
            .Show
        End With
Me.Visible = False
End Sub
'
Private Sub chkAddExData_Click()
'нажатие на кнопу "Добавить данные"
    With chkAddExData
        If .Value = 1 Then
            Call AE_Enabled
            .BackColor = &H80FFFF
            'добавляем класс "Заключение общего эксперта"
            Set newAEF = New clmExpertFindings
                With newAEF
                    .name = "судебно-медицинской"
                    .definition = newEF.definition
                    .tanatology = True
                    .condition = "в работе"
                    .evidensCategories = "экспертизы трупа"
                End With
        Else
            Call AE_Disabled
            .BackColor = &HDDEADB
            'уничтожение класса "Заключение общего эксперта"
            Set newAEF = Nothing
        End If
    End With
End Sub
'
Private Sub AE_Disabled()
'Процедура исключения полей Общий эксперт
    Dim x As Object
        For Each x In Me.Controls
            If InStrRev(x.name, "txtAE", 5) > 0 _
                Or x.name = "txtAutopsyDate" Or _
                x.name = "txtInjPrDecease" Or _
                x.name = "txtInjPrAutopsy" Then
                x.Text = ""
                x.Enabled = False
                x.BackColor = &HEBFFEA
            End If
        Next x
'уничтожение класса "Заключение общего эксперта"
Set newAEF = Nothing
End Sub
'
Private Sub AE_Enabled()
'Процедура включения полей Общий эксперт
    Dim x As Object
        For Each x In Me.Controls
            If InStrRev(x.name, "txtAE", 5) > 0 _
                Or x.name = "txtAutopsyDate" Or _
                x.name = "txtInjPrDecease" Or _
                x.name = "txtInjPrAutopsy" Then
                x.Enabled = True
                x.BackColor = &HEBFFEA
            End If
        Next x
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
'Процедура добавления эксперта
'1)скрытие формы:
Me.Visible = False
'Функция выбора эксперта из списка
Set SvcService = CreateObject("Svcsvc.Service") 'объект библиотеки svcsvc.dll
    Do
        newExpert = SvcService.SelectValue("Абраменко М.С." & vbCrLf & _
                                        "Бубенко Н.И." & vbCrLf & _
                                        "Дегтярев Н.Е." & vbCrLf & _
                                        "Комар В.Е." & vbCrLf & _
                                        "Куль О.А." & vbCrLf & _
                                        "Набебин А.В." & vbCrLf & _
                                        "Санкович В.В." & vbCrLf & _
                                        "Филиппов Ю.А." & vbCrLf, _
                                        "Выберите эксперта", True)
'Debug.Print " addExpert: " & newExpert
        If newExpert <> "" Then
            'формирование массива фамилий экспертов:
            mdPrintDoc.arrExpert = Split(newExpert, vbCrLf)
            Exit Do
        Else
        'если эксперт не выбран:
            MsgBox "Cледует выбрать фамилию эксперта!", vbCritical, Msg
        End If
    Loop
'отладка массива
'Dim i As Byte
'For i = LBound(arrExpert) To UBound(arrExpert)
'Debug.Print "Массив: " & i & " - " & arrExpert(i)
'Next i
Set SvcService = Nothing
End Sub
'
Private Sub txtEnabled()
'выключение отдельных полей
    Dim x As Object
    For Each x In Me.Controls
        If TypeName(x) = "TextBox" Then
            If x.name = "txtCsIndex" Or _
                x.name = "txtCsNum" Or _
                x.name = "txtCrPost" Or _
                x.name = "txtCrOprAr" Or _
                x.name = "txtCrRank" Or _
                x.name = "txtCrSurName" Or _
                x.name = "txtCrName" Or _
                x.name = "txtCrMidName" Or _
                x.name = "txtInjPrDecease" Or _
                x.name = "txtInjPrAutopsy" Then
                x.Enabled = True
            End If
        ElseIf TypeName(x) = "ComboBox" Then
            If x.name = "lstCrSex" Then
                x.Enabled = True
            End If
        End If
     Next x
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
Private Sub lblInjPDeceaseDate_Click()

End Sub
'
Private Sub txtAEFcode_GotFocus()
     Call txtEnter(txtAEFcode)
End Sub
'
Private Sub txtAEFcode_LostFocus()
     With txtAEFcode
        If .Text = "" Then
             Call txt_Exit(txtAEFcode)
        Else:
        .BackColor = RGB(200, 256, 200) 'frmColor.Салатовый
        newAEF.EFcode = .Text
        End If
    End With
End Sub
'
Private Sub txtAEMidName_GotFocus()
'отчество общего эксперта
Call txtEnter(txtAEMidName)
End Sub

Private Sub txtAEMidName_LostFocus()
'отчество общего эксперта
 With txtAEMidName
        If .Text = "" Then
             Call txt_Exit(txtAEMidName)
        Else
             If Len(txtAEMidName.Text) = 1 Then
                newAExperts.midName = newCor.create_Initials(.Text)
            Else
                newAExperts.midName = StrConv(.Text, vbProperCase)
            End If
            .BackColor = RGB(200, 256, 200) 'frmColor.Салатовый
            .Text = newAExperts.midName
        End If
    End With
End Sub
'
Private Sub txtAEName_GotFocus()
'имя общего эксперта
    Call txtEnter(txtAEName)
End Sub
'
Private Sub txtAEName_LostFocus()
'имя общего эксперта
 With txtAEName
        If .Text = "" Then
             Call txt_Exit(txtAEName)
        Else: newAExperts.name = StrConv(.Text, vbProperCase)
         'постановка точки после инициала
            If Len(txtAEName.Text) = 1 Then
                newAExperts.name = newCor.create_Initials(.Text)
            Else: newAExperts.name = StrConv(.Text, vbProperCase)
            End If
            .BackColor = RGB(200, 256, 200) 'frmColor.Салатовый
            .Text = newAExperts.name
        End If
    End With
End Sub

Private Sub txtAESurName_GotFocus()
'Фамилия общего эксперта
 Call txtEnter(txtAESurName)
End Sub

Private Sub txtAESurName_LostFocus()
'Фамилия общего эксперта
    With txtAESurName
        If .Text = "" Then
             Call txt_Exit(txtAESurName)
        Else: newAExperts.surName = StrConv(.Text, vbProperCase)
            .BackColor = RGB(200, 256, 200) 'frmColor.Салатовый
            .Text = newAExperts.surName
        End If
    End With
End Sub
'
Private Sub txtAEFNum_GotFocus()
'номер заключения  общего эксперта
    Call txtEnter(txtAEFNum)
End Sub
'
Private Sub txtAEFNum_LostFocus()
'номер заключения  общего эксперта
With txtAEFNum
        If .Text = "" Then
             Call txt_Exit(txtAEFNum)
        Else: newAEF.number = StrConv(txtAEFNum.Text, vbProperCase)
            .BackColor = RGB(200, 256, 200) 'frmColor.Салатовый
'            .Text = newAExperts.surName
        End If
    End With
End Sub
'
Private Sub txtAutopsyDate_GotFocus()
'Дата вскрытия
   Call txtEnter(txtAutopsyDate)
   'Дата вскрытия = Дата начала
    txtAutopsyDate.Text = newDate.dateToString(newEF.rulingDate)
End Sub
'
Private Sub txtAutopsyDate_LostFocus()
'Дата начала общей экспертизы <=  дате вскрытия
 Dim tmp As Double, dt As Date, str As String
    If txtAutopsyDate.Text = "" Then
        Call txt_Exit(txtAutopsyDate)
    Else
        With newDate
            dt = .validateDate(.ExamDate(txtAutopsyDate.Text))
'            'сравнение дат: Дата вскрытия >= Дата начала общей экспертизы
            Do
                tmp = .compareDt(newEF.rulingDate, dt)
                If tmp < 0 Then
                    MsgBox "Дата начала общей экспертизы меньше даты вынесения постановления!", _
                            vbCritical, Msg
                    dt = .ExamDate(InputBox("Правильно введите дату начала экспертизы трупа!", _
                                        "Ввод даты", newEF.rulingDate))
                Else
                    newAEF.firstDay = dt
                    Exit Do
                End If
            Loop
        End With
        With txtAutopsyDate
            .BackColor = RGB(200, 256, 200)
            .Text = newDate.dateToString(newAEF.firstDay)
        End With
    End If
'    'передача фокуса полю
Me.cmdOK1.SetFocus
'    With Me.txtFactCase
''        Call txtEnter(txtFactCase)
''        .Text = newDate.dateToString(newEF.rulingDate) & " труп " & newInjPr.create_InitialslName
''    End With
''Debug.Print "Дата вскрытия: " & newInjPr.autopsyDate
''On Error Resume Next ' Отключаем ошибки
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
        .Caption = "Заключение " & newEF.getFullNumber & Chr(32) & mdPrintDoc.DocCat
'        mdPrintDoc.DocCat & " №" & newEF.getFullNumber '" Создание нового документа №" & newEF.getFullNumber "ГМСЭ: " & Me.newExpert &
        With .txtEFFirstDay
            .SetFocus 'передача фокуса
            Call txtEnter(Me.txtEFFirstDay)
            .Text = newDate.DateNow
        End With
    End With
'Debug.Print "Номер экспертизы: №" & newEF.number
End Sub
'2) Дата начала экспертизы
Private Sub txtEFFirstDay_GotFocus()
''Дата начала экспертизы
End Sub
'
Private Sub txtEFFirstDay_LostFocus()
'Дата начала мед-крим экспертизы(5), больше или равна другим датам:
'вводится первой, поэтому, не сравнивается с другими датами (т.к. они еще не введены)!!!
'зависимости остальных дат:
'                              Дата смерти(1)
'                           <= Дата вынесения постановления(2)
'                           <= Дата начала экспертизы трупа(3)
'                           <= Дата вскрытия трупа(4)
'                           <= Дата начала мед-крим экспертизы(5)
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
         mdPrintDoc.strMonth = Me.newDate.monthOfYear(.firstDay) 'название текущего месяца для работы с Excel
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
        .AddItem "судебно-медицинской"
        .AddItem "судебной медико-криминалистической"
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
'добавление экспертиз, входящих в комплекс:
        If .Text = "комплексной" Then
            'вставка открытия формы для ввода названия экспертиз
           'Функция выбора эксперта из списка
            Set SvcService = CreateObject("Svcsvc.Service") 'объект библиотеки svcsvc.dll
            Do
                newEF.name = SvcService.SelectValue("судебно-биологической" & vbCrLf & _
                                        "судебно-генетической" & vbCrLf & _
                                        "медико-криминалистической" & vbCrLf & _
                                        "дактилоскопической" & vbCrLf & _
                                        "трасологической" & vbCrLf & _
                                        "экспертизы холодного оружия" & vbCrLf & _
                                        "судебной экспертизы волокнистых материалов и изделий из них" & vbCrLf & _
                                        "судебно-медицинской" & vbCrLf & _
                                        "судебно-гистологической" & vbCrLf & _
                                        "судебно-химической", _
                                        "Выберите экспертизы:", True)
                                        Debug.Print " addnewEFname: " & newEF.name
                If newEF.name <> "" Then
                'формирование массива названий экспертиз
                    mdPrintDoc.arrEF = Split(newEF.name, vbCrLf)
                    Exit Do:
                Else 'если экспертизы не выбраны:
                    MsgBox "Cледует выбрать экспертизы!", vbCritical, Msg
                End If
            Loop
'отладка массива
'Dim i As Byte
'For i = LBound(arrEF) To UBound(arrEF)
'Debug.Print "Массив: " & i & " - " & arrEF(i)
'Next i
            Set SvcService = Nothing
        End If
        'проверка на значение поля
        If .Text = "" Then
            .Text = "судебной медико-криминалистической"
            mdPrintDoc.arrEF = Split(.Text)
        End If
'       создание строки с перечислением экспертиз:
        If newEF.name <> "" Then
            newEF.categories = .Text & " (" & Join(mdPrintDoc.arrEF, ", ") & ") "
        Else: newEF.categories = .Text
        End If
            .BackColor = frmColor.Салатовый 'RGB(200, 256, 200)
    End With
Debug.Print "Категория экспертизы= ", newEF.categories
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
'Дата вынесения постановления(2)
'              <= Дата начала экспертизы трупа(3)
'              <= Дата вскрытия трупа(4)
'              <= Дата начала мед-крим экспертизы(5)
'

 Dim tmp As Double, dt As Date
 With txtCsPrvDate
    If .Text = "" Then
        Call txt_Exit(txtCsPrvDate)
    Else
        With newDate
            dt = .validateDate(.ExamDate(txtCsPrvDate.Text)) 'newEF.rulingDate
            'сравнение дат: Дата начала мед-крим экспертизы(5)>= Дата вынесения постановления(2)
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
    newCase.index = ""
End Sub
'
Private Sub txtCsIndex_Change()
'Индекс Дела
If txtCsIndex <> "не указан" Then
    Dim str As String
    With txtCsIndex
        str = Right(.Text, 1)
        newCase.index = newCase.index + StrConv(str, vbProperCase)
    End With
End If
End Sub
'
Private Sub txtCsIndex_LostFocus()
'Индекс Дела
    With txtCsIndex
        If Len(.Text) = 0 Then
            .Width = 1398.29
            .Left = 246.226
            Call txt_Exit(txtCsIndex)
        Else: 'newCase.index = StrConv(.Text, vbProperCase)
            .BackColor = frmColor.Салатовый 'RGB(200, 256, 200)
            .Text = newCase.index
        End If
    End With
Debug.Print "Индекс Дела: ", newCase.index
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
   txtInjPrDecease.Text = newAEF.firstDay - 1
End Sub

Private Sub txtInjPrDecease_LostFocus()
'Дата смерти(1)<= Дата вынесения постановления(2)
'              <= Дата начала экспертизы трупа(3)
'              <= Дата вскрытия трупа(4)
'              <= Дата начала мед-крим экспертизы(5)
'
 Dim tmp As Double, dt As Date ' str As String
    If txtInjPrDecease.Text = "" Then
        Call txt_Exit(txtInjPrDecease)
    Else
        With newDate
            dt = .validateDate(.ExamDate(txtInjPrDecease.Text))
            'сравнение дат: Дата смерти(1) <= Дата вынесения постановления(2)
            Do
                tmp = .compareDt(newAEF.firstDay, dt)
                If tmp > 0 Then
                    MsgBox "Дата смерти больше даты вынесения постановления!", _
                            vbCritical, Msg
                    dt = .ExamDate(InputBox("Правильно введите дату смерти!", _
                                        "Ввод даты", newAEF.firstDay - 1))
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
'        Call txtEnter(Me.txtInjPrAutopsy)
'        .Text = newAEF.firstDay
    End With
'Debug.Print "Дата смерти: " & newInjPr.decease
End Sub
''
'22)Дата вскрытия
Private Sub txtInjPrAutopsy_GotFocus()
'Дата вскрытия
   Call txtEnter(txtInjPrAutopsy)
   'Дата вскрытия = Дата начала
   txtInjPrAutopsy.Text = newAEF.firstDay 'newInjPr.autopsyDate 'newEF.rulingDate
End Sub
'
    Private Sub txtInjPrAutopsy_LostFocus()
'Дата вскрытия трупа(4)
 Dim tmp As Double, dt As Date, str As String
    If txtInjPrAutopsy.Text = "" Then
        Call txt_Exit(txtInjPrAutopsy)
    Else
        With newDate
            dt = .validateDate(.ExamDate(txtInjPrAutopsy.Text))
            'сравнение дат: Дата вскрытия трупа(4) >= Дата вынесения постановления
            Do
                tmp = .compareDt(newAEF.firstDay, dt)
                If tmp < 0 Then
                    MsgBox "Дата вскрытия меньше даты начала экспертизы трупа!", _
                            vbCritical, Msg
                    dt = .ExamDate(InputBox("Правильно введите дату вскрытия!", _
                                        "Ввод даты", newAEF.firstDay))
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
'    With Me.txtFactCase
'        Call txtEnter(txtFactCase)
'        .Text = newDate.dateToString(newEF.rulingDate) & " труп " & newInjPr.create_InitialslName
'    End With
'Debug.Print "Дата вскрытия: " & newInjPr.autopsyDate
'On Error Resume Next ' Отключаем ошибки
End Sub
'
'+++++++++++++++++++++++++++++++++++++++ К Н О П К И: ++++++++++++++++++++++++++++++++++++++++++++++++++
Private Sub cmdOK1_Click()
'событие нажатия кнопки "ОК" формы "frmNewEF"
'1) заполнение коллекции для работы с Excel:
    With mdPrintDoc.colExcelData
        .Add txtDN.Text, "A" 'номера СМЭ
            Debug.Print "Запись номера СМЭ в коллекции colExcelData -> " & .Item("A")
        .Add txtEFFirstDay.Text, "E"   'Дата Начала
            Debug.Print "Запись Дата Начала в коллекции colExcelData -> " & .Item("E")
        .Add newDate.dateToString(newEF.dueDate), "G" 'Срок
            Debug.Print "Запись Срок в коллекции colExcelData -> " & .Item("G")
        .Add cboCsCat.Text, "C" 'Категория дела
            Debug.Print "Запись Категория дела в коллекции colExcelData -> " & .Item("C")
        .Add txtCsNum.Text, "D" 'номера дела
            Debug.Print "Запись номера дела в коллекции colExcelData -> " & .Item("D")
    End With
'2) создание папок, заключения эксперта и шаблонов направлений
    Call mdMainFolders.makeDocDir(Me.newEF.number, mdPrintDoc.DocCat)
    Call show_MsgCreateNewBox
End Sub
'
Public Sub cmdCancel1_Click()
'кнопка "Отмена"
    Unload Me
End Sub
'
Private Sub txtDisabled()
'выключение отдельных полей
    Dim x As Object
    For Each x In Me.Controls
        If TypeName(x) = "TextBox" Then
            If x.name = "txtCsIndex" Or _
                x.name = "txtCsNum" Or _
                x.name = "txtCrPost" Or _
                x.name = "txtCrOprAr" Or _
                x.name = "txtCrRank" Or _
                x.name = "txtCrSurName" Or _
                x.name = "txtCrName" Or _
                x.name = "txtCrMidName" Then
                x.Enabled = False
            Else: x.Text = ""
            End If
        ElseIf TypeName(x) = "ComboBox" Then
            If x.name = "cboCsCat" Or _
                x.name = "lstCrSex" Then
                x.Enabled = False
            Else: x.Clear
            End If
        End If
     Next x
End Sub
'
Private Sub cmdEraseData_Click()
'кнопка "Очистить"
Dim x As Object
    For Each x In Me.Controls
        If TypeName(x) = "TextBox" Then
            x.Text = ""
            If x.name = "txtDN" Or _
                x.name = "txtEFFirstDay" Or _
                x.name = "txtCsPrvDate" Or _
                x.name = "txtCsIndex" Or _
                x.name = "txtCsNum" Then
                x.BackColor = &HFEF7F1
            ElseIf x.name = "txtCrPost" Or _
                x.name = "txtCrOprAr" Or _
                x.name = "txtCrRank" Or _
                x.name = "txtCrSurName" Or _
                x.name = "txtCrName" Or _
                x.name = "txtCrMidName" Then
                x.BackColor = &HFFE8DF
            ElseIf x.name = "txtInjPrSurName" Or _
                x.name = "txtInjPrName" Or _
                x.name = "txtInjPrMidName" Or _
                x.name = "txtInjPrBirthday" Then
                x.BackColor = &HF0FBE1
            End If
        ElseIf TypeName(x) = "ComboBox" Then
            x.Clear
            If x.name = "cboEFCategories" Or _
                x.name = "cboCsDefinition" Or _
                x.name = "cboCsCat" Then
                x.BackColor = &HFEF7F1
            ElseIf x.name = "lstCrSex" Then
                x.BackColor = &HFFE8DF
            ElseIf x.name = "lstInjPrSex" Then
                x.BackColor = &HF0FBE1
            End If
        End If
         x.Enabled = True
    Next x
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
'================== I C A P S U L A T I O N =============================
'
Private Property Let strBoxKey(ByVal vData As String)
'ключ для упаковок
     mvarstrBoxKey = vData
End Property
'
Friend Property Get strBoxKey() As String
'ключ для упаковок
strBoxKey = mvarstrBoxKey
'     Dim str As String
'        If mvarstrBoxKey <> "" Then
'            str = BOX & mvarstrBoxKey
'        End If
'     strBoxKey = str
'Debug.Print "ключ для упаковок = " & strBoxKey
End Property
'
Friend Property Let strEvKey(ByVal vData As String)
'ключ для вещдоков в упаковке
    mvarstrEvKey = CStr(Format(vData, "#0000"))
'    Debug.Print "ключ для вещдоков = " & frmNewEF.strEvKey
End Property
'
Friend Property Get strEvKey() As String
'ключ для вещдоков в упаковке
     Dim str As String
        If mvarstrEvKey <> "" Then
            str = EVID & mvarstrEvKey
        End If
     strEvKey = str
Debug.Print "ключ для вещдоков в упаковке = " & strEvKey
End Property
'
Friend Property Get strKey() As String
'общий ключ = "BX****" + "EV****" ("BX****EV****")
     Dim str As String
        If strBoxKey <> "" And strEvKey <> "" Then
            str = strBoxKey & strEvKey
        End If
     strKey = str
Debug.Print "ключ для вещдоков в упаковке = " & strKey
End Property



'Private Sub debugPrint_colBoxes()
''отладочная печать данных коллекции коробок
'Dim tmpBx As clmEvBox
'Dim i As Long, bxKey As String
' With frmNewEF.colBoxes
'    For i = 1 To .Count
'        bxKey = BOX & CStr(Format(i, "#0000"))
'        Set tmpBx = .Item(bxKey)
'            Debug.Print "Отладочная печать коробки_" & bxKey
'            Debug.Print "Название коробки ->" & tmpBx.strBxName
'            Debug.Print "Место изъятия ВД -> " & tmpBx.strBxPlace
'            Debug.Print "принадлежность ВД -> " & tmpBx.strBxEntrants
'            Debug.Print "Упаковка -> " & tmpBx.strBxPackage
'            Debug.Print "Категория -> " & tmpBx.strBxCategory
'            Debug.Print "Описание коробки -> " & tmpBx.strBxObjDescription
''           ВД
'            Dim ev As Variant
'            For Each ev In tmpBx.colEvidences
'                Debug.Print "Содержание коробки: "
'                Debug.Print "ВД - > " & ev
'            Next ev
'         Set tmpBx = Nothing
'        Next i
'    End With
'End Sub

'Public Sub create_Analysis()
''Создание бланков направлений на исследования
'Set newDOC.MyWdApp = New Word.Application 'экземпляр приложения
''1)создание папки для создаваемых документов:
'Dim nameFolder As String
'    nameFolder = frmAnalysis.Caption & "_" & frmNewResearch.newEF.getNumber
'Dim nameDOC As String 'имя документа ИмяДокумента_Номер_Год
''2) создание новых документов и сохранение их в созданной папке:
'Dim X As Object 'переменная контролов
'    For Each X In frmAnalysis.Controls
'        If TypeName(X) = "CheckBox" Then
'            If X.Value = 1 Then
'                With newDOC
'                    nameDOC = X.Caption & "_" & frmNewResearch.newEF.getNumber
'                    .dirAnalysis = .Create_mainFolders(.dirExpert, nameFolder)
'                    Call .print_Blank(X.Caption, nameDOC, .dirAnalysis)
'                End With
'            End If
'        End If
'    Next X
'End Sub
'

''++++++++++++++++++ М Е Т О Д Ы +++++++++++++++++++++++++++
''
'Private Sub chkAddExData_Click()
''нажатие на кнопу "Добавить данные"
'    With chkAddExData
'        If .Value = 1 Then
'            Call AE_Enabled
'            .BackColor = &H80FFFF
'            'добавляем класс "Заключение общего эксперта"
'            Set newAEF = New clmExpertFindings
'                With newAEF
'                    .name = "судебно-медицинской"
'                    .definition = newEF.definition
'                    .tanatology = True
'                    .condition = "в работе"
'                    .evidensCategories = "экспертизы трупа"
'                End With
'        Else
'            Call AE_Disabled
'            .BackColor = &HDDEADB
'            'уничтожение класса "Заключение общего эксперта"
'            Set newAEF = Nothing
'        End If
'    End With
'End Sub
''
'Private Sub AE_Disabled()
''Процедура исключения полей Общий эксперт
'    Dim X As Object
'        For Each X In Me.Controls
'            If InStrRev(X.name, "txtAE", 5) > 0 Or X.name = "txtAutopsyDate" Then
'                X.Text = ""
'                X.Enabled = False
'                X.BackColor = &HEBFFEA
'            End If
'        Next X
''уничтожение класса "Заключение общего эксперта"
'Set newAEF = Nothing
'End Sub
''
'Private Sub AE_Enabled()
''Процедура включения полей Общий эксперт
'    Dim X As Object
'        For Each X In Me.Controls
'            If InStrRev(X.name, "txtAE", 5) > 0 Or X.name = "txtAutopsyDate" Then
'                X.Enabled = True
'                X.BackColor = &HEBFFEA
'            End If
'        Next X
'End Sub
''

'Lib++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'Private NewCase As clmCase                      'Экземпляр класса "Дело"
'Private NewCoroner As HR_Official               'экземпляр класса "Следователь"
'Private NewInjuredPerson As Entrants            'экземпляр класса "Потерпевший"
'Private NewExperts As HR_Official               'экземпляр класса "эксперты"
'Private NewAutopsyExperts As HR_Official        'экземпляр класса "общие эксперты"
'Private NewExpertFindings As clmExpertFindings  'эксземпляр класса "Заключение эксперта"
'Private NewAutopsyEF As clmExpertFindings       '"Заключение общего эксперта", эксземпляр класса "Заключение эксперта"
'Private mvarcolEvidences As Collection          'коллекция "ВД" (для хранения значений массива)
'Private newDate As clmCaseDate                  'экземпляр класса "Даты"
'Private fst As Date, susp As Date, ren As Date, fin As Date, dueDate As Date 'переменные дат: начала, приостановки, возобновления и окончания

''раннее связывание приложений "MsOffice"
'Private MyWdApp As Word.Application         'экземпляр приложения MsOffice Word
'Private MyWdDoc As Word.Document                'экземпляр документа MsOffice Word
'Private MyExApp As Excel.Application        'экземпляр приложения MsOffice Excel
''Private MyExDoc As Excel.Workbook               'экземпляр документа MsOffice Excel
'Private MyOlApp As Outlook.Application      'экземпляfр приложения MsOffice Outlook
''Private MyOlTask As Outlook.TaskItem            'экземпляр задачи MsOffice Outlook
'Private fs As New FileSystemObject              'Объявление нового объекта (папки)
'Private fll As Variant
''Объявление объекта приложения ActiveX
'Private SvcService As Object, Sel As String
''директория расположения шаблонов документов по работе
''Private DirDOT As String
'Const DirDOT = "D:\Crime\MasterForm\Word\"
''"D:\Андрей\Програмирование\В разработке\Окончание проекта\Проект FM_Doc\MasterDoc\" 'путь шаблонов
''объявление переменной "Отмена"
'Dim Cancel As Integer
''
'Public DocNum As String 'Объявление переменной-счетчика номер документа
''Объявление массива
'Private EvidArray() As String
'Private CellArr(1 To 26) As String
''Объявление счетчика
'Public n As Byte, i As Byte, bytEvid As Byte, WdEvName As String

''
'Public Property Set colEvidences(ByVal vData As Collection)
''Коллекция вещественых доказательств
''Syntax: Set x.colEvidences = Form1
'    Set mvarcolEvidences = vData
'End Property
''
'Public Property Get colEvidences() As Collection
''Коллекция вещественых доказательств
'Set colEvidences = mvarcolEvidences
''Debug.Print "Коллекция ВД =  " & colEvidences
'End Property
''
'Private Sub cboCsCat_GotFocus()
''Категория дела
'    With cboCsCat
'        .Clear
'        .BackColor = &HC0FFFF 'Подсвечивание активного поля при получении им фокуса
''Добавить строки в комбинированное поле
'        .AddItem "уголовного дела"
'        .AddItem "проверки"
'        .AddItem "административного дела"
'    End With
'End Sub
''
'Private Sub cboCsCat_LostFocus()
''Категория дела
'    With cboCsCat
'        If Len(.Text) = 0 Then
'            .Text = "(категория дела не указана)"
'            .BackColor = &HC0E0FF
'        Else: newCase.strCsCat = .Text 'Присвоение переменым значений
'            If newCase.strCsCat = "проверки" Then
'                txtCsIndex.Visible = True
'                txtCsIndex.SetFocus
'            Else: txtCsIndex.Visible = False
'            End If
'        .BackColor = RGB(200, 256, 200) 'Прекращение посвечивания поля при потере им фокуса
'        End If
'    End With
'End Sub
''
'Private Sub cboCsDefinition_GotFocus()
''Основание проведения экспертизы
'    With cboCsDefinition
'        .Clear
'        .BackColor = &HC0FFFF
'        .AddItem "постановления"
'        .AddItem "определения"
'    End With
'End Sub
''
'Private Sub cboCsDefinition_LostFocus()
''Основание проведения экспертизы
'    With cboCsDefinition
'        If Len(.Text) = 0 Then
'            .Text = "постановления"
'        End If
'        NewExpertFindings.strEFDefinition = .Text 'Присвоение переменым значений
'        .BackColor = RGB(200, 256, 200) 'Прекращение посвечивания поля при потере им фокуса
'    End With
'End Sub
''
'    Private Sub cboEFCategories_GotFocus()
''Категория экспертизы (первичная, дополнительная и т.д.)
'    With cboEFCategories
'        .Clear
'        .BackColor = &HC0FFFF
'        .AddItem "медико-криминалистической"
'        .AddItem "комплексной"
'        .AddItem "дополнительной"
'        .AddItem "повторной"
'        .AddItem "комиссионной"
'    End With
'End Sub
''
'Private Sub cboEFCategories_LostFocus()
''Категория экспертизы (первичная, дополнительная и т.д.)
'    With cboEFCategories
'        If Len(.Text) = 0 Then
'            .Text = "медико-криминалистической"
'        End If
'            NewExpertFindings.strEFCategories = .Text
'            .BackColor = RGB(200, 256, 200)
'    End With
''Debug.Print "Категория экспертизы= ", NewExpertFindings.strEFCategories
'End Sub
''

'Private Sub cmdCancel1_Click()
''Нажатие кнопки "Отмена"
''Уничтожение новых экземпляров классов и освобождение памяти
'Set newCase = Nothing           'Экземпляр класса "Дело"
'Set NewInjuredPerson = Nothing  'экземпляр класса "Потерпевший"
'Set NewCoroner = Nothing        'экземпляр класса "Следователь"
'Set NewExperts = Nothing        'Экземпляр класса "Эксперт"
'Set NewAutopsyExperts = Nothing 'экземпляр класса "Эксперты"
'Set NewExpertFindings = Nothing 'эксземпляр класса "Заключение эксперта"
'Set NewAutopsyEF = Nothing      '"Заключение общего эксперта", эксземпляр класса "Заключение эксперта"
'Unload Me
'End Sub
'
'Private Sub cmdEraseData_Click()
'Dim X As Object
'For Each X In Me.Controls
'    If TypeName(X) = "TextBox" Then
'        X.Text = ""
'        X.BackColor = &H80000005
'    ElseIf TypeName(X) = "ComboBox" Then
'        X.Clear
'        X.BackColor = &H80000005
'    End If
'Next X
'Call EF_Enabled
'Call Cs_Enabled
'Call Cor_Enabled
'Call InjPr_Enabled
'Call AE_Disabled
'Set MyWdApp = Nothing
''Set MyExApp = Nothing
'Set MyWdApp = New Word.Application
''Set MyExApp = New Excel.Application
'chkAddExData.Value = 0
'txtEFnum.SetFocus
'End Sub
''
'Private Sub cmdSaveCreate_Click()
''процедура нажатия кнопки "Cоздать с таким же основанием"
''выключение нужных и заполенных полей
'Call Cs_Disabled
'Call Cor_Disabled
'Call InjPr_Disabled
'Call AE_ErText
'Call AE_Disabled
''Уничтожение существующих экземпляров классов и освобождение памяти
'Set NewExpertFindings = Nothing 'эксземпляр класса "Заключение эксперта"
'Set NewAutopsyEF = Nothing      '"Заключение общего эксперта", эксземпляр класса "Заключение эксперта"
'Set NewAutopsyExperts = Nothing 'экземпляр класса "Общий эксперт"
'Set MyWdApp = Nothing
'Set MyExApp = Nothing
''Создание новых экземпляров классов
'Set NewExpertFindings = New clmExpertFindings  'эксземпляр класса "Заключение эксперта"
'Set NewAutopsyEF = New clmExpertFindings       '"Заключение общего эксперта", эксземпляр класса "Заключение эксперта"
'Set NewAutopsyExperts = New HR_Official        'экземпляр класса "общие эксперты"
'Set MyWdApp = New Word.Application
'Set MyExApp = New Excel.Application
''Изменение счетчиков
'mdCount.fEFCount = mdCount.fEFCount + 1
'Me.lblEFCount = mdCount.fEFCount 'Изменение номера формы
'mdCount.colForms.Add Me, "EFform" & mdCount.fEFCount
'    Dim strTemp As String
'        strTemp = Me.Caption
''дописать: длина надписи - номер
''        Me.Caption = Right(strTemp, 2) & " " & mdCount.fEFCount
'txtEFnum.SetFocus
'End Sub
''
'Private Sub lstAESex_GotFocus()
''Пол общего эксперта
'    With lstAESex
'        .Clear
'        .AddItem "муж."
'        .AddItem "жен."
'        .AddItem "пол не указан"
'        .BackColor = &HC0FFFF
'    End With
'End Sub
''
'Private Sub lstAESex_LostFocus()
''Пол общего эксперта
'    With lstAESex
'        NewAutopsyExperts.HR_sex = .Text
'        .BackColor = RGB(200, 256, 200)
'    End With
''Debug.Print "Пол общего эксперта", NewAutopsyExperts.strAESex
'End Sub
''
'Private Sub lstCrSex_GotFocus()
''Пол следователя
'    With lstCrSex
'        .Clear
'        .AddItem "муж."
'        .AddItem "жен."
'        .AddItem "пол не указан"
'        .BackColor = &HC0FFFF
'    End With
'End Sub
''
'Private Sub lstCrSex_LostFocus()
''Пол следователя
'    With lstCrSex
'        NewCoroner.HR_sex = .Text
'        .BackColor = RGB(200, 256, 200)
'    End With
''Debug.Print "Пол следователя= ", NewCoroner.strCrSex
'End Sub
''
'Private Sub lstInjPrSex_GotFocus()
' 'Пол потерпевшего
'    With lstInjPrSex
'        .Clear
'        .AddItem "муж."
'        .AddItem "жен."
'        .AddItem "пол не указан"
'        .BackColor = &HC0FFFF
'    End With
'End Sub
''
'Private Sub lstInjPrSex_LostFocus()
''Пол потерпевшего
'    With lstInjPrSex
'        NewInjuredPerson.En_Sex = .Text
'        .BackColor = RGB(200, 256, 200)
'    End With
''Debug.Print "Пол потерпевшего = ", NewInjuredPerson.strInjPrSex
'End Sub
''
'Private Sub Erase_Data()
''Процедура удаления (очистки) даных текстовых полей формы
'   Dim X As Object
'For Each X In Me.Controls
'    If TypeName(X) = "TextBox" Then
'        X.Text = ""
'        X.BackColor = &H80000005
'    ElseIf TypeName(X) = "ComboBox" Then
'        X.Clear
'        X.BackColor = &H80000005
'    End If
'Next X
'End Sub
''

''
'Private Sub Open_DocListForm()
''Процедура открытия формы "Перечень документов"
'    Dim frmD As frmDocList
'    Set frmD = New frmDocList
'    frmD.Caption = Me.Caption
'    frmD.lblDocListCount = Me.lblEFCount
'    frmD.Show
'End Sub

'Public Static Sub Form_Initialize()
''Создание новых экземпляров
'Set newCase = New clmCase                      'Экземпляр класса "Дело"
'Set NewInjuredPerson = New Entrants            'экземпляр класса "Потерпевший"
'Set NewCoroner = New HR_Official               'экземпляр класса "Следователь"
'Set NewExperts = New HR_Official               'экземпляр класса "эксперты"
'Set NewAutopsyExperts = New HR_Official        'экземпляр класса "общие эксперты"
'Set NewExpertFindings = New clmExpertFindings  'эксземпляр класса "Заключение эксперта"
'Set NewAutopsyEF = New clmExpertFindings       '"Заключение общего эксперта" - эксземпляр класса "Заключение эксперта"
'Set newDate = New clmCaseDate                  'экземпляр класса "Даты"
''Инициализация коллекции массива данных
'Set colEvidences = New Collection
''Инициализация докуметнов:
'Set MyWdApp = New Word.Application
''Set MyExApp = New Excel.Application
''Set MyOlApp = New Outlook.Application
'End Sub
''
'Private Sub Form_Load()
'    Me.Height = 8730
'    Me.Width = 12165
'    Move (Screen.Width - Width) / 2, (Screen.Height - Height) / 17.7
'chkAddExData.Value = 0
'Call AE_Disabled
'End Sub
''
'Private Sub Form_Unload(Cancel As Integer)
''Процедура выгрузки формы
'    If MsgBox("Уверены?", vbYesNo, "Выход?") = vbYes Then
'        Set frmEvCategories = Nothing
'    'Удаление объектов классов
'        Set newCase = Nothing           'Экземпляр класса "Дело"
'        Set NewInjuredPerson = Nothing  'экземпляр класса "Потерпевший"
'        Set NewCoroner = Nothing        'экземпляр класса "Следователь"
'        Set NewExperts = Nothing        'Экземпляр класса "Эксперт"
'        Set NewAutopsyExperts = Nothing 'экземпляр класса "Общий эксперт"
'        Set NewExpertFindings = Nothing 'эксземпляр класса "Заключение эксперта"
'        Set NewAutopsyEF = Nothing      '"Заключение общего эксперта", эксземпляр класса "Заключение эксперта"
'        Set newDate = Nothing
'        Set MyWdApp = Nothing
''        Set MyExApp = Nothing
''        Set MyOlApp = Nothing
'Dim tmp As String
'        tmp = Me.lblEFCount
''Выгрузка скрытой формы frmEvidences
'    Unload mdCount.colForms("Evidform" & tmp)
''Очистка коллекции форм colForms:
'    mdCount.colForms.Remove ("Evidform" & tmp)
'    mdCount.colForms.Remove ("DocListform" & tmp)
'    mdCount.colForms.Remove ("EFform" & tmp)
'    Else
'        Cancel = 1
'    End If
'End Sub
''
'Private Static Sub txtAEMidName_GotFocus()
''Отчество общего эксперта
'    With txtAEMidName
'        .Text = ""
'        .BackColor = &HC0FFFF
'    End With
'End Sub
''
'Private Static Sub txtAEMidName_LostFocus()
''Отчество общего эксперта
'    With txtAEMidName
'        If Len(.Text) = 0 Then
'            .Text = "(не указано)"
'            .BackColor = &HC0E0FF
'        Else
'            If Len(.Text) = 1 Then
'                NewAutopsyExperts.HR_MidName = StrConv(.Text, vbProperCase) & "."
'            Else: NewAutopsyExperts.HR_MidName = StrConv(.Text, vbProperCase)
'            End If
'        .BackColor = RGB(200, 256, 200)
'        .Text = NewAutopsyExperts.HR_MidName
'        End If
'    End With
''Debug.Print "Отчество общего эксперта= ", NewAutopsyExperts.HR_MidName
'End Sub
''
'Private Static Sub txtAEName_GotFocus()
''Имя общего эксперта
'    With txtAEName
'        .Text = ""
'        .BackColor = &HC0FFFF
'    End With
'End Sub
''
'Private Static Sub txtAEName_LostFocus()
''Имя общего эксперта
'    With txtAEName
'        If Len(.Text) = 0 Then
'            .Text = "(имя не указано)"
'            .BackColor = &HC0E0FF
'        Else
'            If Len(.Text) = 1 Then
'                NewAutopsyExperts.HR_name = StrConv(.Text, vbProperCase) & "."
'            Else: NewAutopsyExperts.HR_name = StrConv(.Text, vbProperCase)
'            End If
'        .BackColor = RGB(200, 256, 200)
'        .Text = NewAutopsyExperts.HR_name
'        End If
'    End With
''Debug.Print "Имя общего эксперта= ", NewAutopsyExperts.HR_name
'End Sub
''
'Private Static Sub txtAEFNum_GotFocus()
''журнальный номер заключения общего эксперта
'    With txtAEFNum
'        .Text = ""
'        .BackColor = &HC0FFFF
'    End With
'End Sub
''
'Private Static Sub txtAEFNum_LostFocus()
''журнальный номер заключения общего эксперта
'    With txtAEFNum
'        If Len(.Text) = 0 Then
'            .Text = "(не указан)"
'            .BackColor = &HC0E0FF
'        Else: NewAutopsyEF.strEFNum = .Text
'            .BackColor = RGB(200, 256, 200)
'        End If
'    End With
''Debug.Print "номер заключения общего эксперта= ", NewAutopsyEF.strEFNum
'End Sub
''
'Private Static Sub txtAESurName_GotFocus()
''Фамилия общего эксперта
'    With txtAESurName
'        .Text = ""
'        .BackColor = &HC0FFFF
'    End With
'End Sub
''
'Private Static Sub txtAESurName_LostFocus()
''Фамилия общего эксперта
'    With txtAESurName
'        If Len(.Text) = 0 Then
'            .Text = "(не указана)"
'            .BackColor = &HC0E0FF
'        Else: NewAutopsyExperts.HR_SurName = StrConv(.Text, vbProperCase) 'написание с заглавной буквы
'            .Text = NewAutopsyExperts.HR_SurName
'            .BackColor = RGB(200, 256, 200)
'        End If
'    End With
''Debug.Print "Фамилия общего эксперта= ", NewAutopsyExperts.HR_SurName
'End Sub
''
'Private Static Sub txtAutopsyDate_GotFocus()
''Дата начала общей СМЭ
'    With txtAutopsyDate
'        .Text = ""
'        .BackColor = &HC0FFFF
'        .Text = CStr(NewExpertFindings.DtmEFFirstDay - 2)
'    End With
'End Sub
''
'Private Static Sub txtAutopsyDate_LostFocus()
''Дата начала общей СМЭ
'    With txtAutopsyDate
'        If Len(.Text) = 0 Then
'            .Text = "(не указана)"
'            .BackColor = &HC0E0FF
'        Else: NewInjuredPerson.En_Autopsy = CDate(.Text)
'            .BackColor = RGB(200, 256, 200)
'        End If
'    End With
''Debug.Print "Дата начала общей СМЭ= ", NewInjuredPerson.En_Autopsy
'End Sub
''
'Private Static Sub txtCorOprAr_GotFocus()
''Район обслуживания
'    With txtCorOprAr
'        .Text = ""
'        .BackColor = &HC0FFFF
'        .Text = "(г. Минска) районного отдела Следственного комитета Республики Беларусь"
'    End With
'End Sub
''
'Private Static Sub txtCorOprAr_LostFocus()
''Район обслуживания
'    With txtCorOprAr
'        If Len(.Text) = 0 Then
'            .Text = "(отдел не указан)"
'            .BackColor = &HC0E0FF
'        Else: NewCoroner.HR_OprAr = .Text
'            .BackColor = RGB(200, 256, 200)
'        End If
'    End With
''Debug.Print "Отдел = ", NewCoroner.HR_OprAr
'End Sub
''
'Private Static Sub txtCorPost_GotFocus()
''Должность следователя
'    With txtCorPost
'    .BackColor = &HC0FFFF
'    .Text = "старшего следователя"
'    End With
'End Sub
''
'Private Static Sub txtCorPost_LostFocus()
''Должность следователя
'    With txtCorPost
'        If Len(.Text) = 0 Then
'            .Text = "(должость не указана)"
'            .BackColor = &HC0E0FF
'        Else: NewCoroner.HR_Post = .Text
'            .BackColor = RGB(200, 256, 200)
'        End If
'    End With
''Debug.Print "Должность следователя= ", NewCoroner.HR_Post
'End Sub
''
'Private Sub txtCsIndex_GotFocus()
''Индекс Дела
'    With txtCsIndex
'        .Text = ""
'        .BackColor = &HC0FFFF
'    End With
'End Sub
''
'Private Sub txtCsIndex_LostFocus()
''Индекс Дела
'    With txtCsIndex
'        If Len(.Text) = 0 Then
'            .Text = "(не указан)"
'            .BackColor = &HC0E0FF
'        Else: newCase.strCsIndex = .Text
'            .BackColor = RGB(200, 256, 200)
'        End If
'    End With
''Debug.Print "Индекс Дела= ", NewCase.strCsIndex
'End Sub
''
'Private Static Sub txtEFnum_GotFocus()
''Номер крим экспертизы
'    With txtEFnum
'        .Text = ""
'        .BackColor = &HC0FFFF
'        .ForeColor = &H80000008
'    End With
'End Sub
''
'Private Static Sub txtEFnum_LostFocus()
''Номер крим экспертизы
'    With txtEFnum
'        Do
'            If Not IsNumeric(.Text) Or Len(.Text) = 0 Then
'                Beep
'                .BackColor = RGB(256, 0, 0)
'                MsgBox "Cледует вводить цифры", vbCritical, "Ошибка ввода"
'                .Text = InputBox("Введите правильно номер экспертизы!", "Исправление ошибки ввода")
'            Else: NewExpertFindings.strEFNum = .Text
'                .BackColor = RGB(200, 256, 200)
'        Exit Do
'            End If
'        Loop
'    End With
'''заполнение категории и шаблона документа:
''NewExpertFindings.DOC = mdPrintDoc.strDOC 'шаблон
'NewExpertFindings.strEFEvCategories = mdPrintDoc.DocCat 'категория
'End Sub
''
'Private Static Sub txtCsNum_GotFocus()
''Номер дела
'    With txtCsNum
'        .BackColor = &HC0FFFF
'        .Text = ""
'        .MaxLength = 11
'    End With
'End Sub
''
'Private Static Sub txtCsNum_LostFocus()
''Номер дела CsNum
'    With txtCsNum
'        If Len(.Text) = 0 Then
'            .MaxLength = 22
'            .Text = "(номер дела не указан)"
'            .BackColor = &HC0E0FF
'        Else
'            Do
'                If Not IsNumeric(.Text) Then
'                    Beep
'                    MsgBox "Cледует вводить цифры", vbCritical, "Ошибка ввода"
'                    .Text = InputBox("Введите правильно номер дела!", "Исправление ошибки ввода")
'                Else: newCase.strCsNum = .Text
'                    .BackColor = RGB(200, 256, 200)
'            Exit Do
'                End If
'            Loop
'        End If
'    End With
'' Debug.Print "Номер дела= ", NewCase.strCsNum
'End Sub
''
'Private Static Sub txtCorMidName_GotFocus()
''Отчество следователя
'    With txtCorMidName
'        .Text = ""
'        .BackColor = &HC0FFFF 'Подсвечивание активного поля при получении им фокуса
'    End With
'End Sub
''
'Private Static Sub txtCorMidName_LostFocus()
''Отчество следователя
'    With txtCorMidName
'        If Len(.Text) = 0 Then
'            .Text = "(отчество не указано)"
'            .BackColor = &HC0E0FF
'        Else
'            If Len(txtCorMidName.Text) = 1 Then
'                NewCoroner.HR_MidName = StrConv(txtCorMidName.Text, vbProperCase) & "."
'            Else: NewCoroner.HR_MidName = StrConv(.Text, vbProperCase)
'            End If
'            .BackColor = RGB(200, 256, 200) 'Прекращение посвечивания поля при потере им фокуса
'            .Text = NewCoroner.HR_MidName
'        End If
'    End With
''Debug.Print "Отчество следователя= ", NewCoroner.HR_MidName
'End Sub
''
'Private Static Sub txtCorName_GotFocus()
''Имя следователя
'    With txtCorName
'        .Text = ""
'        .BackColor = &HC0FFFF 'Подсвечивание активного поля при получении им фокуса
'    End With
'End Sub
''
'Private Static Sub txtCorName_LostFocus()
''Имя следователя
'    With txtCorName
'        If Len(.Text) = 0 Then
'            .Text = "(имя не указано)"
'            .BackColor = &HC0E0FF
'        Else
''постановка точки после инициала
'            If Len(txtCorName.Text) = 1 Then
'                NewCoroner.HR_name = StrConv(.Text, vbProperCase) & "."
'            Else: NewCoroner.HR_name = StrConv(txtCorName.Text, vbProperCase)
'            End If
'        .Text = NewCoroner.HR_name
'        .BackColor = RGB(200, 256, 200) 'Прекращение посвечивания поля при потере им фокуса
'        End If
'    End With
''Debug.Print "Имя следователя= ", NewCoroner.HR_name
'End Sub
''
'Private Static Sub txtCorRank_GotFocus()
''Звание следователя
'    With txtCorRank
'        .BackColor = &HC0FFFF 'Подсвечивание активного поля при получении им фокуса
'        .Text = "старшего лейтенанта юстиции"
'    End With
'End Sub
''
'Private Static Sub txtCorRank_LostFocus()
''Звание следователя
'    With txtCorRank
'        If Len(.Text) = 0 Then
'            .Text = "(звание не указано)"
'            .BackColor = &HC0E0FF
'        Else: NewCoroner.HR_Rank = .Text
'            .BackColor = RGB(200, 256, 200)
'        End If
'    End With
''Debug.Print "Звание следователя= ", NewCoroner.HR_Rank
'End Sub
''
'Private Static Sub txtEFFirstDay_GotFocus()
''Дата начала
'    With txtEFFirstDay
'        .Text = ""
'        .BackColor = &HC0FFFF
'        .ForeColor = &H80000008
'        .Text = newDate.DateNow 'отображение текущей даты в текстовом поле
'    End With
'End Sub
''
'Private Static Sub txtEFFirstDay_LostFocus()
''Дата начала
'With txtEFFirstDay
'    fst = newDate.ExamDate(.Text)
'    NewExpertFindings.DtmEFFirstDay = fst
'    .Text = newDate.dateToString(fst)
'    .BackColor = RGB(200, 256, 200)
'
'
''    If Len(.Text) = 0 Then
''        .Text = InputBox("Введите дату начала экспертизы", "Ввод даты", DateTime.Date)
''            If Len(.Text) = 0 Then
''                .Text = "не указана"
''                .BackColor = RGB(256, 0, 0)
''                .ForeColor = &HFFFF&
''            Else
''                Dim tmpDt As Date
''                    tmpDt = newDate.ExamDate(.Text)
''                    .Text = newDate.dateToString(tmpDt)
''                NewExpertFindings.DtmEFFirstDay = tmpDt
''                .ForeColor = &H80000008
''                On Error Resume Next ' Отключаем ошибки
''            End If
''    ElseIf .Text = "не указана" Then
''            .BackColor = RGB(256, 0, 0)
''            .ForeColor = &HFFFF&
''    Else
''        tmpDt = newDate.ExamDate(.Text)
''        NewExpertFindings.DtmEFFirstDay = tmpDt
''        .ForeColor = &H80000008
''        .Text = newDate.dateToString(tmpDt)
''    'вызов процедуры расчета срока окончания экспертизы
'    NewExpertFindings.DtmEFDueDate = newDate.getPeriod(fst)
''    End If
'End With
'Debug.Print "Срок =", NewExpertFindings.DtmEFDueDate
'End Sub
''
'Private Static Sub txtCorSurName_GotFocus()
''Фамилия следователя
'    With txtCorSurName
'        .Text = ""
'        .BackColor = &HC0FFFF
'    End With
'End Sub
''
'Private Static Sub txtCorSurName_LostFocus()
''Фамилия следователя
'    With txtCorSurName
'        If Len(.Text) = 0 Then
'            .Text = "(фамилия не указана)"
'            .BackColor = &HC0E0FF
'        Else: NewCoroner.HR_SurName = StrConv(.Text, vbProperCase)
'            .BackColor = RGB(200, 256, 200)
'            .Text = NewCoroner.HR_SurName
'        End If
'    End With
''Debug.Print "Фамилия следователя= ", NewCoroner.HR_SurName
'End Sub
''
'Private Static Sub txtInjPrDeceaseDate_GotFocus()
''Дата смерти потерпевшего
'    With txtInjPrDeceaseDate
'        .Text = ""
'        .BackColor = &HC0FFFF
'        .Text = CStr(NewExpertFindings.DtmEFFirstDay - 3)
'    End With
'End Sub
''
'Private Static Sub txtInjPrDeceaseDate_LostFocus()
''Дата смерти потерпевшего
'    With txtInjPrDeceaseDate
'        If Len(.Text) = 0 Then
'            .Text = "(не указана)"
'            .BackColor = &HC0E0FF
'        Else: NewInjuredPerson.En_Decease = .Text
'            .BackColor = RGB(200, 256, 200)
'        End If
'    End With
''Debug.Print "Дата смерти потерпевшего= ", NewInjuredPerson.En_Decease
'End Sub
''
'Private Static Sub txtCsPrvDate_GotFocus()
''Дата вынесения постановления
'    With txtCsPrvDate
'        .Text = ""
'        .BackColor = &HC0FFFF
'        .ForeColor = &H80000008
'        .Text = CStr(NewExpertFindings.DtmEFFirstDay - 3)
'    End With
'End Sub
''
'Private Static Sub txtCsPrvDate_LostFocus()
''Дата вынесения постановления
'    With txtCsPrvDate
'        If Len(.Text) = 0 Then
'            .Text = "(не указана)"
'            .BackColor = RGB(256, 0, 0)
'            .ForeColor = &HFFFF&
'       ElseIf .Text = "не указана" Then
'            .BackColor = RGB(256, 0, 0)
'            .ForeColor = &HFFFF&
'        Else: newCase.DtmCsPrvDate = CDate(.Text)
'            .BackColor = RGB(200, 256, 200)
'        End If
'    End With
''Debug.Print "Дата вынесения постановления= ", NewCase.DtmCsPrvDate
'End Sub
''
'Private Static Sub txtInjPrBirthday_GotFocus()
''Год рождения потерпевшего
'    With txtInjPrBirthday
'        .Text = ""
'        .BackColor = &HC0FFFF
'    End With
'End Sub
''
'Private Static Sub txtInjPrBirthday_LostFocus()
''Год рождения потерпевшего
'    With txtInjPrBirthday
'        If Len(.Text) = 0 Then
'            .Text = "(не указан)"
'            .BackColor = &HC0E0FF
'        Else: NewInjuredPerson.En_Birthday = .Text
'            .BackColor = RGB(200, 256, 200)
'        End If
''!!! Вставить подсчет количества символов или длину текста
'    End With
''Debug.Print "Год рождения потерпевшего", NewInjuredPerson.En_Birthday
'End Sub
''
'Private Static Sub txtInjPrMidName_GotFocus()
''Отчество потерпевшего
'    With txtInjPrMidName
'        .Text = ""
'        .BackColor = &HC0FFFF
'    End With
'End Sub
''
'Private Static Sub txtInjPrMidName_LostFocus()
''Отчество потерпевшего
'    With txtInjPrMidName
'        If Len(.Text) = 0 Then
'            .Text = "(отчество не указано)"
'            .BackColor = &HC0E0FF
'        Else
'            If Len(.Text) = 1 Then
'                NewInjuredPerson.En_MidName = StrConv(.Text, vbProperCase) & "."
'            Else: NewInjuredPerson.En_MidName = StrConv(.Text, vbProperCase)
'            End If
'        .BackColor = RGB(200, 256, 200)
'        .Text = NewInjuredPerson.En_MidName
'        End If
'    End With
''Debug.Print "Отчество потерпевшего= ", NewInjuredPerson.En_MidName
'End Sub
''
'Private Static Sub txtInjPrName_GotFocus()
''Имя потерпевшего
'    With txtInjPrName
'        .Text = ""
'        .BackColor = &HC0FFFF
'    End With
'End Sub
''
'Private Static Sub txtInjPrName_LostFocus()
''Имя потерпевшего
'    With txtInjPrName
'        If Len(.Text) = 0 Then
'            .Text = "(имя не указано)"
'            .BackColor = &HC0E0FF
'        Else
'            If Len(.Text) = 1 Then
'                NewInjuredPerson.En_Name = StrConv(.Text, vbProperCase) & "."
'            Else: NewInjuredPerson.En_Name = StrConv(.Text, vbProperCase)
'            End If
'        .Text = NewInjuredPerson.En_Name
'        .BackColor = RGB(200, 256, 200)
'        End If
'    End With
''Debug.Print "Имя потерпевшего= ", NewInjuredPerson.En_Name
'End Sub
''
'Private Static Sub txtInjPrSurName_GotFocus()
''Фамилия потерпевшего
'    With txtInjPrSurName
'        .Text = ""
'        .BackColor = &HC0FFFF
'    End With
'End Sub
''
'Private Static Sub txtInjPrSurName_LostFocus()
''Фамилия потерпевшего
'    With txtInjPrSurName
'        If Len(.Text) = 0 Then
'            .Text = "(фамилия не указана)"
'            .BackColor = &HC0E0FF
'        Else: NewInjuredPerson.En_SurName = StrConv(.Text, vbProperCase)
'            .BackColor = RGB(200, 256, 200)
'            .Text = NewInjuredPerson.En_SurName
'        End If
'    End With
''Debug.Print "Фамилия потерпевшего= ", NewInjuredPerson.En_SurName
'End Sub
''
'Public Static Sub Criate_NewFolders()
''Создание нового каталога папок в заданной директории
'ChDrive "D" 'выбор необходимого диска
''NewExpertFindings.strEFEvCategories = frmMDI.strCrEFCat
'NewExpertFindings.DOC = "D:\Crime\" & Year(Now) & "\" & NewExpertFindings.strEFNum & "_" & Right(Year(Now), 2) & "_" & _
'            NewExpertFindings.strEFEvCategories & "\"
'If Not fs.FolderExists(NewExpertFindings.DOC) Then  ' если такая папка не существует в этом месте
'        Set fll = fs.CreateFolder(NewExpertFindings.DOC) 'то создаем ее
''создаем вложенную папку для фото
'        fs.CreateFolder (fll & "\" & "Img_" & NewExpertFindings.strEFNum & "_" & Right(Year(Now), 2))
'        fs.CreateFolder (fll & "\" & "Img_" & NewExpertFindings.strEFNum & "_" & Right(Year(Now), 2) & _
'                        "\" & "ImgEOD_" & NewExpertFindings.strEFNum & "_" & Right(Year(Now), 2))
'    Else
'        fll = NewExpertFindings.DOC ' в противном случае переменная  fll будет содеждать ссылку на эту (уже существующую) папку
'    End If
''Debug.Print "DirDOC=", NewExpertFindings.DOC
'End Sub
''
'Public Static Sub Create_NewDoc()
''Процедура создания новых документов
''DirDOT = "D:\Андрей\Програмирование\В разработке\Окончание проекта\Проект FM_Doc\MasterDoc\" 'Дом разработка
'''1)Процедура открытия и заполнения отчета
''    If frmDocList.chkReportEF.Value = 1 Then
''        Create_ReportEF
''    End If
''''Процедура создания задачи Outlook
'''    If frmDocList.chkTaskItem.Value = 1 Then
'''        Create_TaskItem
'''    End If
''______________________________________________________________
''Работа с документами Word
'    With MyWdApp
'        .Visible = False
''mdCount.colForms.Add frmD, "EFform" & Me.EvCat_ID
'        Set MyWdDoc = .Documents.Add(DirDOT & mdPrintDoc.strDOC)
'            Create_EF
''временная переменная-счетчик
'Dim tmp As String
'tmp = Me.lblEFCount
'''Процедура создания фототаблицы
''    If mdCount.colForms("DocListform" & tmp).chkFotoListEF.Value = 1 Then
''    Set MyWdDoc = .Documents.Add(DirDOT & "FotoList.dotm")
''        Create_FotoList
''    End If
'''Процедура создания направления на биологию
''    If mdCount.colForms("DocListform" & tmp).chkCNBiology.Value = 1 Then
''        Set MyWdDoc = .Documents.Add(DirDOT & "Biology.dotm")  'Добавляем шаблон"
''        Create_Biology
''    End If
'''Процедура создания направления геном
''    If mdCount.colForms("DocListform" & tmp).chkCNGenom.Value = 1 Then
''        Set MyWdDoc = .Documents.Add(DirDOT & "Genome.dotm")  'Добавляем шаблон"
''        Create_Genom
''    End If
'''Процедура создания направления на химию
''    If mdCount.colForms("DocListform" & tmp).chkCrimMet.Value = 1 Then
''       Set MyWdDoc = .Documents.Add(DirDOT & "FocusOnMet.dotm")  'Добавляем шаблон"
''       Create_FocusOnMet
''    End If
'''Процедура создания этикеток
''    If mdCount.colForms("DocListform" & tmp).chkCoverNote.Value = 1 Then
''       Set MyWdDoc = .Documents.Add(DirDOT & "Labels.dotm")  'Добавляем шаблон"
''        Create_CovNoteLabel
''    End If
'''Процедура создания запроса
''    If mdCount.colForms("DocListform" & tmp).chkAEFInquiry.Value = 1 Then
''       Set MyWdDoc = .Documents.Add(DirDOT & ".dotm")  'Добавляем шаблон"
'''       Create_AEFInquiry
''    End If
'    .Quit 'Закрываем приложение
'    End With
''Set MyWdDoc = Nothing
'End Sub
''
'Public Static Sub Create_EF()
''Процедура создания нового документа Ecspert Findings
'    MyWdDoc.SaveAs2 FileName:=NewExpertFindings.DOC & "Заключение_" & NewExpertFindings.strEFNum _
'            & "_" & Right(Year(Now), 2), FileFormat:=wdFormatXMLDocument, _
'            LockComments:=False, Password:="", AddToRecentFiles:=True, WritePassword:="", _
'            ReadOnlyRecommended:=False, EmbedTrueTypeFonts:=False, SaveNativePictureFormat:=False, _
'            SaveFormsData:=False, SaveAsAOCELetter:=False
'    With MyWdApp.ActiveDocument  'Выбираем активный документ
'        On Error Resume Next
''Поля:
'       .FormFields("WdDirDOC").Result = NewExpertFindings.DOC
'       .FormFields("WdEFNum").Result = NewExpertFindings.strEFNum
'       .FormFields("WdFirstDay").Result = NewExpertFindings.DtmEFFirstDay
''Закладки:
'        .Bookmarks("WdEFfullNum").Range.Text = NewExpertFindings.Criate_FullEFNum 'Код экспертизы
'        .Bookmarks("WdExperience").Range.Text = NewExperts.Calculate_Experience 'расчет стажа работы
'        .Bookmarks("WdRsnComplain").Range.Text = Print_RsnComplain 'юр.основание
'        .Bookmarks("FactCase").Range.Text = Print_FactCase 'Обстоятельства дела
'    If chkAddExData.Value = 1 Then
'        .Bookmarks("WdAEDirection").Range.Text = Print_AEDirection 'Направление общего эксперта
'    End If
'    Dim tmp As String
'        tmp = Me.lblEFCount
'    If mdCount.colForms("DocListform" & tmp).chkAEFInquiry.Value = 1 Then
'        NewExpertFindings.DtmEFSuspDate = DateTime.Date
'        .Bookmarks("WdEFSuspDate").Range.Text = NewExpertFindings.DtmEFSuspDate
'        .Bookmarks("WdAEFInquiry").Range.Text = Print_AEFInquiry 'направленные ходатайства
'    End If
'
''Печать описания упаковок и т.д. вещественных доказательств
''        .Bookmarks("WdEvNumArr1").Range.Text = NewEvid.Print_EvNumArr1
''        .Bookmarks("WdEvNumArr2").Range.Text = NewEvid.Print_EvNumArr2
''        .Bookmarks("WdEvNumArr3").Range.Text = NewEvid.Print_EvNumArr3
'''Печать ВД:
'    .Bookmarks("WdEvColumnArr").Range.Select
'            Call Print_NumColumn
''         For i = 1 To n - 1
''            .Bookmarks("WdEvColumnArr").Range.Text = EvidArray(i) & Chr(10)
''            .Bookmarks("WdEvLineArr").Range.Text = EvidArray(i) & ", "
''            .Bookmarks("WdEvLineArr1").Range.Text = EvidArray(i) & ", "
''            .Bookmarks("WdEvLineArr3").Range.Text = EvidArray(i) & ", "
''        Next i
''    On Error GoTo 999 ' Включаем обработку ошибки
'    .Close SaveChanges:=wdSaveChanges
'    End With
'Set MyWdDoc = Nothing
''999:
''    MsgBox Err.Description  'Ошибка
''    Err.Clear
'End Sub
''

''
'Private Static Sub Create_CovNoteLabel()
''Создание этикеток
''    MyWdDoc.SaveAs2 FileName:=NewExpertFindings.DOC & "Этикетки_" & NewExpertFindings.strEFNum _
''            & "_" & Right(Year(Now), 2), FileFormat:=wdFormatXMLDocument, _
''            LockComments:=False, Password:="", AddToRecentFiles:=True, WritePassword:="", _
''            ReadOnlyRecommended:=False, EmbedTrueTypeFonts:=False, SaveNativePictureFormat:=False, _
''            SaveFormsData:=False, SaveAsAOCELetter:=False
''    With MyWdApp.ActiveDocument  'Выбираем активный документ
''        On Error Resume Next ' Отключаем ошибки
'''работа с полями и закладками документа:
'''Поля:
'''работа с полем документа:
''    .Bookmarks("WdCovNoteLable").Range.Text = Print_CovNoteLabel
'''Завершение работы с документом
''        .Save
''        .Close
''     End With
'End Sub
''
'Private Static Function Print_CovNoteLabel() As String
''  'Функция печати этикеток
'''    вставка массива:
''Dim strCN As String
''    If NewEvid.bytEvUnit = 1 Then
''        Print_CovNoteLabel = "Управление медико-криминалистических экспертиз." & Chr(10) _
''                & "Заключение " & NewExpertFindings.Criate_FullEFNum & NewCase.Print_CsData & "." & Chr(10) _
''                & Chr(10) & "Объект: " & EvidArray(1) & "." & Chr(10) & Chr(10) & "Дата регистрации объектов: " & _
''                NewExpertFindings.DtmEFFirstDay & "." & Chr(10) & "Дата упаковки: " & DateTime.Date & Chr(10) _
''                & Chr(10) & "ГМСЭ" & vbTab & Chr(32) & vbTab & Chr(32) & vbTab & Chr(32) & vbTab & Chr(32) _
''                & vbTab & Chr(32) & vbTab & Chr(32) & vbTab & Chr(32) & vbTab & Chr(32) & vbTab _
''                & Chr(32) & "А.В. Набебин."
''    Else:
''            For i = 1 To n - 1
''                strCN = strCN & "-" & EvidArray(i) & ";" & Chr(10)
''                Print_CovNoteLabel = "Управление медико-криминалистических экспертиз." & Chr(10) _
''                & "Заключение " & NewExpertFindings.Criate_FullEFNum & NewCase.Print_CsData & "." & Chr(10) _
''                & Chr(10) & "Объекты:" & Chr(10) & strCN & Chr(10) & "Дата регистрации объектов: " & _
''                NewExpertFindings.DtmEFFirstDay & "." & Chr(10) & "Дата упаковки: " & DateTime.Date & "." & Chr(10) _
''                & Chr(10) & "ГМСЭ" & vbTab & Chr(32) & vbTab & Chr(32) & vbTab & Chr(32) & vbTab & Chr(32) _
''                & vbTab & Chr(32) & vbTab & Chr(32) & vbTab & Chr(32) & vbTab & Chr(32) & vbTab _
''                & Chr(32) & "А.В. Набебин."
''            Next i
''   End If
'End Function
''
'Private Static Sub Print_CovNoteLabel_1()
''''Функция печати этикеток
''''    вставка массива:
'''N = i + 1
'''    For i = 1 To N - 1
'''        Selection.TypeParagraph
'''        Selection.TypeText Text:="Управление медико-криминалистических экспертиз" & Chr(10) _
'''                & "Заключение " & NewExpertFindings.Criate_FullEFNum & NewCase.Print_CsData & Chr(10) _
'''                & "Объект:"
'''        Selection.TypeParagraph
'''        Selection.TypeText Text:=EvidArray(i) & Chr(10)
'''        Selection.TypeParagraph
'''        Selection.TypeText Text:="Дата регистрации объектов: " & _
'''                NewExpertFindings.DtmEFFirstDay & Chr(10) & "Дата упаковки: "
''''          вставка ссылки на поле "текущая дата"
'''        Selection.InsertCrossReference ReferenceType:="Закладка", ReferenceKind:= _
'''            wdContentText, ReferenceItem:="ActualDate", InsertAsHyperlink:=True, _
'''            IncludePosition:=False, SeparateNumbers:=False, SeparatorString:=" "
'''        Selection.MoveLeft Unit:=wdCharacter, Count:=1, Extend:=wdExtend
'''        Selection.Font.Name = "Times New Roman"
'''        Selection.Font.Size = 14
'''        Selection.MoveRight Unit:=wdCharacter, Count:=1
'''        Selection.TypeParagraph
'''        Selection.TypeText Text:="Эксперт" & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab _
'''            & vbTab & "А.В. Набебин" & Chr(10)
''''            & "Лаборант" & vbTab & vbTab & vbTab & vbTab _
''''            & vbTab & vbTab & vbTab & Chr(10) & "И.В. Астрейко"
'''        Selection.TypeParagraph
'''    Next i
'''End Sub
'''Private Static Sub AEFdirection()
'''    If chkAddExData.Value = 0 Then
'''        Selection.TypeText Text:="Из направления " & WithDoc.Print_AEFData & " следует: ... " & WithDoc.Print_AEFData & _
'''            " Краткие обстоятельства и предполагаемая причина смерти: ..."
'''    End If
'End Sub
''
'Private Static Sub Print_PetitionCor()
''Функция набора текста "Запрос следователю"
''    If frmDocList.chkCoverNote.Value = 1 Then
''        Selection.TypeText Text:=WithDoc.dtFirstDay & "на имя " & WithDoc.Print_Coroner _
''           & " выслано ходатайство о предоставлении заключения (копии заключения)" & _
''        " по определению характера и степени тяжести телесных повреждений у " & WithDoc.Print_InjPr & " (" & WithDoc.strAEFNum & ") "
''    End If
'End Sub
''
'Private Static Sub Create_Biology()
''Процедура создания нового документа "Направление на биологию"
'    MyWdDoc.SaveAs2 FileName:=NewExpertFindings.DOC & "Биология_" & NewExpertFindings.strEFNum _
'            & "_" & Right(Year(Now), 2), FileFormat:=wdFormatXMLDocument, _
'            LockComments:=False, Password:="", AddToRecentFiles:=True, WritePassword:="", _
'            ReadOnlyRecommended:=False, EmbedTrueTypeFonts:=False, SaveNativePictureFormat:=False, _
'            SaveFormsData:=False, SaveAsAOCELetter:=False
'    With MyWdApp.ActiveDocument  'Выбираем активный документ
'        On Error Resume Next ' Отключаем ошибки
''Закладки:
''        .Bookmarks("WdRsnComplain").Range.Text = NewCoroner.Print_Coroner
''Массивы:
''         For i = 1 To n - 1
''          .Bookmarks("WdEvLineArr").Range.Text = EvidArray(i) & ", "
''          .Bookmarks("WdEvLineArr1").Range.Text = EvidArray(i) & ", "
''          .Bookmarks("WdEvColumnArr").Range.Text = "-" & EvidArray(i) & Chr(10)
''        Next i
' .Close SaveChanges:=wdSaveChanges
'    End With
'Set MyWdDoc = Nothing
'End Sub '
''
'Private Static Sub Create_Genom()
''Процедура создания нового документа "Направление на геном"
'    MyWdDoc.SaveAs2 FileName:=NewExpertFindings.DOC & "Геном_" & NewExpertFindings.strEFNum _
'            & "_" & Right(Year(Now), 2), FileFormat:=wdFormatXMLDocument, _
'            LockComments:=False, Password:="", AddToRecentFiles:=True, WritePassword:="", _
'            ReadOnlyRecommended:=False, EmbedTrueTypeFonts:=False, SaveNativePictureFormat:=False, _
'            SaveFormsData:=False, SaveAsAOCELetter:=False
'    With MyWdApp.ActiveDocument  'Выбираем активный документ
'        On Error Resume Next ' Отключаем ошибки
'        'Закладки:
''        .Bookmarks("WdRsnComplain").Range.Text = NewCoroner.Print_Coroner
''        'Массивы:
''         For i = 1 To n - 1
''          .Bookmarks("WdEvLineArr").Range.Text = EvidArray(i) & ", "
''          .Bookmarks("WdEvLineArr1").Range.Text = EvidArray(i) & ", "
''          .Bookmarks("WdEvColumnArr").Range.Text = "-" & EvidArray(i) & Chr(10)
''        Next i
'    .Close SaveChanges:=wdSaveChanges
'    End With
'Set MyWdDoc = Nothing
'End Sub
''
'Private Static Sub Create_FocusOnMet()
''Процедура создания нового документа "Направление на химию, металлизация"
'    MyWdDoc.SaveAs2 FileName:=NewExpertFindings.DOC & "Направление на Me_" & NewExpertFindings.strEFNum _
'            & "_" & Right(Year(Now), 2), FileFormat:=wdFormatXMLDocument, _
'            LockComments:=False, Password:="", AddToRecentFiles:=True, WritePassword:="", _
'            ReadOnlyRecommended:=False, EmbedTrueTypeFonts:=False, SaveNativePictureFormat:=False, _
'            SaveFormsData:=False, SaveAsAOCELetter:=False
'    With MyWdApp.ActiveDocument  'Выбираем активный документ
'        On Error Resume Next ' Отключаем ошибки
'        'Поля:
''.FormFields("WdInjPr").Result = NewInjuredPerson.Create_InitialsInjPrFullName
''.FormFields("WdInjPrBirthday").Result = NewInjuredPerson.strInjPrBirthday & "г.р."
''.FormFields("WdInjPrSex").Result = NewInjuredPerson.strInjPrSex
''     n = i + 1
''    For i = 1 To n - 1
''        If i <= 8 Then
''            .FormFields("WdEvidences" & i).Result = EvidArray(i)
''        End If
''    Next i
''.FormFields("WdDeceaseDate").Result = NewInjuredPerson.DtmInjPrDecease
''.FormFields("WdAutopsyDate").Result = NewInjuredPerson.DtmInjPrAutopsy
''.FormFields("WdEFfullNum").Result = NewExpertFindings.Criate_FullEFNum
''.FormFields("WdAEFNum").Result = NewAutopsyEF.strEFNum
''.FormFields("WdLegalGround").Result = Print_LegalGround
'''Закладки:
''.Bookmarks("WdEvFirstDate").Range.Text = NewExpertFindings.DtmEFFirstDay
''.Bookmarks("WdCsData").Range.Text = NewCase.Print_CsData
''.Bookmarks("WdEFfullNum1").Range.Text = NewExpertFindings.Criate_FullEFNum
''    For i = 1 To n - 1
''        .Bookmarks("PrintCovNoteLable").Range.Text = "-" & EvidArray(i) & Chr(10)
''    Next i
' .Close SaveChanges:=wdSaveChanges
'    End With
'Set MyWdDoc = Nothing
'End Sub '
''
'Private Static Function Print_AEDirection() As String
''Печать шаблона направления эксперта
'Print_AEDirection = "Из направления " & NewAutopsyExperts.Print_Expert & " следует: " & _
'        Chr(34) & NewInjuredPerson.Print_Autopsy & " Заключение (Акт) №" & NewAutopsyEF.strEFNum _
'        & " Краткие обстоятельства и предполагаемая причина смерти Судебно-медицинский диагноз …" _
'        & Chr(34) & "." & Chr(10)
'End Function
''
'Private Static Function Print_AEFInquiry() As String
''  Функция печати раздела "Заявленные ходатайства"
'Dim tmp As String
'        tmp = Me.lblEFCount
'    If mdCount.colForms("DocListform" & tmp).chkAEFInquiry.Value = 1 Then
'        Print_AEFInquiry = DateTime.Date & " на имя " & NewCoroner.Print_Coroner _
'            & " выслано ходатайство о предоставлении заключения (копии заключения) эксперта по определению характера и степени тяжести телесных повреждений у " _
'            & NewInjuredPerson.Print_InjPersons & Chr(10) & _
'            vbTab & "Запрашиваемые материалы предоставлены/не предоставлены." & Chr(10)
'    Else: Print_AEFInquiry = "Ходатайства не заявлялись." & Chr(10)
'    End If
'End Function
''
'Private Static Function Print_FactCase() As String
''Функция печати "обстоятельства дела"
'  Print_FactCase = newCase.DtmCsPrvDate & Chr(32) & NewInjuredPerson.Print_InjPersons & Chr(10)
'End Function
''
'Public Static Function Print_LegalGround() As String
''Функция набора текста "Юридическое основание проведения экспертизы"
''        Print_LegalGround = "на основании " & NewCase.strCsDefinition & Chr(32) & NewCoroner.Print_Coroner _
''        & Chr(32) & NewCase.Print_CsData
'End Function
''
'Public Static Function Print_RsnComplain() As String
''Функция набора текста "Основание проведения экспертизы"
'    If txtAESurName.Text = "" Then
'        Print_RsnComplain = " на основании " & NewExpertFindings.strEFDefinition & Chr(32) _
'        & "о назначении " & NewExpertFindings.strEFCategories & " экспертизы " & NewCoroner.Print_Coroner _
'        & ", вынесенного " & newCase.DtmCsPrvDate & Chr(32) & newCase.Print_CsData & ", провёл медико-криминалистическую экспертизу."
'    Else: Print_RsnComplain = " на основании " & NewExpertFindings.strEFDefinition & Chr(32) _
'    & "о назначении " & NewExpertFindings.strEFCategories & " экспертизы " & NewCoroner.Print_Coroner _
'        & ", вынесенного " & newCase.DtmCsPrvDate & Chr(32) & newCase.Print_CsData & " и направления " & NewAutopsyExperts.Print_Expert & " от " & _
'        NewInjuredPerson.En_Autopsy & ", провёл медико-криминалистическую экспертизу."
'    End If
'End Function
''
'Private Static Sub Create_FotoList()
''Создание фототаблицы
'    MyWdDoc.SaveAs2 FileName:=NewExpertFindings.DOC & "Фототаблица_" & NewExpertFindings.strEFNum _
'            & "_" & Right(Year(Now), 2), FileFormat:=wdFormatXMLDocument, _
'            LockComments:=False, Password:="", AddToRecentFiles:=True, WritePassword:="", _
'            ReadOnlyRecommended:=False, EmbedTrueTypeFonts:=False, SaveNativePictureFormat:=False, _
'            SaveFormsData:=False, SaveAsAOCELetter:=False
'    With MyWdApp.ActiveDocument  'Выбираем активный документ
'        On Error Resume Next ' Отключаем ошибки
''Работа с верхним колонтитулом:
'    If .ActiveWindow.View.SplitSpecial <> wdPaneNone Then
'        .ActiveWindow.Panes(2).Close
'    End If
'    If .ActiveWindow.ActivePane.View.Type = wdNormalView Or ActiveWindow. _
'        ActivePane.View.Type = wdOutlineView Then
'        .ActiveWindow.ActivePane.View.Type = wdPrintView
'    End If
'    .ActiveWindow.ActivePane.View.SeekView = wdSeekCurrentPageHeader
'    .Bookmarks("WdEFfullNum").Range.Text = NewExpertFindings.Criate_FullEFNum
'    .ActiveWindow.ActivePane.View.SeekView = wdSeekMainDocument
''''работа с закладкой основного документа:
'''i = 1
''    For i = 1 To n - 1 Step 1
''        .Bookmarks("WdEvColumnArr").Range.Text = "Фото № " & "." & Chr(32) & EvidArray(i) & Chr(10) & Chr(10)
''    Next i
''Завершение работы с документом
'        .Close SaveChanges:=wdSaveChanges
'    End With
'Set MyWdDoc = Nothing
'End Sub
''
'Private Sub AE_ErText()
''Процедура очистки полей Общий эксперт
'    Dim X As Object
'        For Each X In Me.Controls
'            If InStrRev(X.name, "txtAE", 5) > 0 Or X.name = "txtAutopsyDate" Then
'                X.Text = ""
'                X.BackColor = &HC0FFC0
'                X.Enabled = False
'            End If
'        Next X
'    chkAddExData.Value = 0
'End Sub
''
'Private Sub Cor_Disabled()
''Процедура исключения полей "Следователь"
'    Dim X As Object
'        For Each X In Me.Controls
'            If InStrRev(X.name, "txtCor", 6) > 0 Or X.name = "lstCrSex" Then
'                X.Enabled = False
'                X.BackColor = &HFFC2AC
'            End If
'        Next X
'End Sub
''
'Private Sub Cor_Enabled()
''Процедура включения полей "Следователь"
'    Dim X As Object
'        For Each X In Me.Controls
'            If InStrRev(X.name, "txtCor", 6) > 0 Or X.name = "lstCrSex" Then
'                X.Enabled = True
'                X.BackColor = &HFFE8DF
'            End If
'        Next X
'End Sub
''
'Private Sub InjPr_Disabled()
''Процедура исключения полей "Потерпевший"
'    Dim X As Object
'        For Each X In Me.Controls
'            If InStrRev(X.name, "txtInjPr", 9) > 0 Or X.name = "lstInjPrSex" Then
'                X.Enabled = False
'                X.BackColor = &HFCE4FB
'            End If
'        Next X
'End Sub
''
'Private Sub InjPr_Enabled()
''Процедура включения полей "Потерпевший"
'    Dim X As Object
'        For Each X In Me.Controls
'            If InStrRev(X.name, "txtInjPr", 9) > 0 Or X.name = "lstInjPrSex" Then
'                X.Enabled = True
'                X.BackColor = &HFDEEFC
'            End If
'        Next X
'End Sub
''
'Private Sub Cs_Disabled()
''Процедура исключения полей "Дело"
'    Dim X As Object
'        For Each X In Me.Controls
'            If InStrRev(X.name, "txtCs", 5) > 0 Or InStrRev(X.name, "cboCs", 5) > 0 Then
'                X.Enabled = False
'                X.BackColor = &HFDEADB
'            End If
'        Next X
'End Sub
'
'Private Sub Cs_Enabled()
''Процедура включения полей "Дело"
'    Dim X As Object
'        For Each X In Me.Controls
'            If InStrRev(X.name, "txtCs", 5) > 0 Or InStrRev(X.name, "cboCs", 5) > 0 Then
'                X.Enabled = True
'                X.BackColor = &HFEF7F1
'            End If
'        Next X
'End Sub
''
'Private Sub EF_Disabled()
''Процедура исключения полей "Заключение эксперта"
'    Dim X As Object
'        For Each X In Me.Controls
'            If InStrRev(X.name, "txtEE", 5) > 0 Or InStrRev(X.name, "cboEF", 5) > 0 Then
'                X.Enabled = False
'                X.BackColor = &HFDEADB
'            End If
'        Next X
'End Sub
''
'Private Sub EF_Enabled()
''Процедура включения полей "Заключение эксперта"
'    Dim X As Object
'        For Each X In Me.Controls
'            If InStrRev(X.name, "txtEF", 5) > 0 Or InStrRev(X.name, "cboEF", 5) > 0 Then
'                X.Enabled = True
'                X.BackColor = &HFEF7F1
'            End If
'        Next X
'    txtEFnum.SetFocus
'End Sub
''
'''Private Static Sub Print_CovNoteLable1(ByVal i As Byte, N As Byte)
''''Функция печати этикеток
''''    вставка массива:
'''N = i + 1
'''    For i = 1 To N - 1
'''        Selection.TypeParagraph
'''        Selection.TypeText Text:="Мед.крим №5.2/" & WithDoc.strDN & " от " & WithDoc.dtFirstDay & Chr(10) & _
'''            "по материалам " & WithDoc.strCsCat & " №" & WithDoc.strCsNum & Chr(10)
'''        Selection.TypeParagraph
'''        Selection.TypeText Text:=EvidArray(i) & Chr(10)
'''        Selection.TypeParagraph
''''          вставка ссылки на поле "текущая дата"
'''        Selection.InsertCrossReference ReferenceType:="Закладка", ReferenceKind:= _
'''            wdContentText, ReferenceItem:="ActualDate", InsertAsHyperlink:=True, _
'''            IncludePosition:=False, SeparateNumbers:=False, SeparatorString:=" "
'''        Selection.MoveLeft Unit:=wdCharacter, Count:=1, Extend:=wdExtend
'''        Selection.Font.Name = "Times New Roman"
'''        Selection.Font.Size = 14
'''        Selection.MoveRight Unit:=wdCharacter, Count:=1
'''        Selection.TypeParagraph
'''        Selection.TypeText Text:="Эксперт" & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab _
'''            & vbTab & "А.В. Набебин" & Chr(10) & "Лаборант" & vbTab & vbTab & vbTab & vbTab _
'''            & vbTab & vbTab & vbTab & Chr(10) '& "И.В. Астрейко"
'''        Selection.TypeParagraph
'''    Next i
'''End Sub
''
'
''
'Static Function Print_EvPackage() As String
''Функция печати описания упаковок и т.д. вещественных доказательств
''Dim strTx1 As String, strTx2 As String
''strTx2 = "мастичным оттиском синего цвета круглой печати"
''    If StrComp(frmNewEF.colEvidences(), "картонную коробку") = True Then
''        strTx1 = "опечатанную"
''    Else: strTx1 = "опечатанный"
''    Debug.Print "strTx1 = ", strTx1
''    End If
''Dim EvNumArr1(1, 1 To 2) 'создание массива
''    EvNumArr1(1, 1) = "Вещественное доказательство доставлено нарочным, упакованное в " & strEvPackage & _
''           Chr(32) & strTx1 & Chr(32) & strTx2
''    EvNumArr1(1, 2) = "Вещественные доказательства доставлены нарочным, упакованные в " & strEvPackage & _
''           Chr(32) & strTx1 & Chr(32) & strTx2
'''Print_EvNumArr1 = EvNumArr1(1, intEvGRCount)
'End Function
''
''Static Function Print_PacIntegrity() As String
'''целостность упаковки
'''Dim strTx1 As String, strTx2 As String, strTx3 As String, strTx4 As String
'''    strTx1 = "Целостность упаковки не нарушена, извлечение"
'''    strTx2 = "без повреждения целостности упаковки невозможно. При вскрытии упаковки из нее"
'''    strTx3 = "для медико-криминалистического исследования,"
'''    strTx4 = "перечную, указанному в направлении и в сопроводительной надписи к вещественным доказательствам. "
'''Dim EvNumArr2(1, 1 To 2) 'создание массива
'''    EvNumArr2(1, 1) = strTx1 & " предоставленного объекта " & strTx2 & " был извлечен объект: " _
'''    & Chr(10) & _
'''    "Объект, представленный " & strTx3 & " соответствует " & strTx4
'''    EvNumArr2(1, 2) = strTx1 & " предоставленных объектов  " & strTx2 & " были извлечены следующие объекты: " & Chr(10) & _
'''    "Объекты, представленные " & strTx3 & " соответствуют " & strTx4
''''Print_EvNumArr2 = EvNumArr2(1, intEvGRCount)
''End Function
''''
''Static Function Print_EvStamp() As String
'''печать нашего управления
'''Dim strTx As String
'''    strTx = "мастичным оттиском синего цвета круглой печати " & Chr(171) & _
'''        "Государственный комитет судебных экспертиз Республики Беларусь. Главное управление судебно-медицинских экспертиз. Для пакетов" _
'''        & Chr(187)
'''Dim EvNumArr3(1, 1 To 2) 'создание массива
'''    EvNumArr3(1, 1) = "вещественное доказательство упаковано, опечатано " & strTx
'''    EvNumArr3(1, 2) = "вещественные доказательства упакованы, опечатаны " & strTx
''''Print_EvNumArr3 = EvNumArr3(1, intEvGRCount)
''End Function
'''
'Private Static Sub Print_NumColumn()
''нумерованный список в столбик:
'Dim i As Integer, n As Integer
'Dim tmpEF As String, tmpGr As String, tmpEv As String
'    For i = 1 To CInt(colEvidences("GrCount" & tmpEF & "000" & "0000"))
'        tmpGr = CStr(Format(i, "#000"))
'Debug.Print "Список групп ВД - " & colEvidences("Owner" & tmpEF & tmpGr & "0000")
'        Selection.TypeText Text:=colEvidences("Owner" & tmpEF & tmpGr & "0000")
'            For n = 1 To CInt(colEvidences("EvCount" & tmpEF & tmpGr & "0000"))
'                tmpEv = CStr(Format(n, "#0000"))
'Debug.Print "Список ВД - " & colEvidences("EvName" & tmpEF & tmpGr & tmpEv)
'                Selection.TypeText Text:=n & "." & Chr(32) _
'                    & colEvidences("EvName" & tmpEF & tmpGr & tmpEv) & Chr(10)
'            Next n
'    Next i
'End Sub

'Private Sub Open_EvidForm()
''Процедура открытия формы "Вещественные доказательства"
'    Dim tmp As String 'Временная переменная-счетчик
'        tmp = Me.lblEFCount 'получения номера-связки формы активной формы frmNewEF
'    Dim frmD As frmEvidences 'обявление новой формы
'        mdCount.fEvidCount = mdCount.fEvidCount + 1 'изменение счетчика форм "Вещественные доказательства"
'    Set frmD = New frmEvidences
''добавление формы в коллекцию созданных форм
'        frmD.Caption = mdPrintDoc.DocCat & Chr(32) & tmp 'mdCount.fEvidCount заполнение названия формы
'        frmD.lblEvidCount.Caption = tmp
'        mdCount.colForms.Add frmD, "Evidform" & tmp 'mdCount.fEvidCount
'        With frmD
'            .Caption = NewExpertFindings.strEFEvCategories
'            .txtEvFirstDate = NewExpertFindings.DtmEFFirstDay
'            .txtOwSurName = NewInjuredPerson.En_SurName
'            .txtOwName = NewInjuredPerson.En_Name
'            .txtOwMidName = NewInjuredPerson.En_MidName
'            .txtOwBirthday = NewInjuredPerson.En_Birthday
'            Dim X As Object
'                For Each X In .Controls
'                    If TypeName(X) = "TextBox" Then
'                        If X.Text <> "" Then
'                            X.BackColor = &H9DB1E6
'                        Else: X.BackColor = &HF1F9FA
'                        End If
'                    End If
'                Next X
'        End With
'    frmD.Show
'    Me.Visible = False
'End Sub
















'_____________________________________________________________________
'Public Static Sub Print_NumColumnArr(ByVal i As Byte, n As Byte)
''нумерованный список в столбик:
''n = i + 1
''    For i = 1 To n - 1
'''        Selection.TypeText Text:=i & "." & Chr(32) & EvidArray(i) & Chr(10)
''    Next i
'End Sub
''
'Public Static Sub Print_ColumnArr(ByVal i As Byte, ByVal n As Byte)
''список в столбик:
''n = i + 1
''    For i = 1 To n - 1
'''        Selection.TypeText Text:=EvidArray(i) & Chr(10)
''        Next i
'End Sub
'''
'Public Static Sub Print_LineArr(ByVal i As Byte, ByVal n As Byte)
''список в строку
''n = i + 1
''    For i = 1 To n - 1
'''       Selection.TypeText Text:=EvidArray(i) & ", "
''        Next i
'End Sub
'
