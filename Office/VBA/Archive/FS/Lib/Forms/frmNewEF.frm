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
   ScaleMode       =   0  '����������������
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
      Alignment       =   2  '���������
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
      Alignment       =   2  '���������
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
      Alignment       =   2  '���������
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
      Alignment       =   2  '���������
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
      Alignment       =   2  '���������
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
      Caption         =   "&��������"
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
      Caption         =   "����������� ������  ��������"
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
         Alignment       =   2  '���������
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
         Alignment       =   2  '���������
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
         Alignment       =   2  '���������
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
         Alignment       =   2  '���������
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
         Caption         =   "&�������� ������"
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
         Caption         =   "���� ��������"
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
         Caption         =   "���� ������"
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
         Caption         =   "��� �"
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
         Caption         =   "��"
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
         Caption         =   "��������"
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
         Caption         =   "���"
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
         Caption         =   "�������"
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
      Caption         =   "�����������"
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
         Alignment       =   2  '���������
         BackColor       =   &H00D1CFA1&
         Caption         =   "���"
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
         Caption         =   "�������"
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
         Caption         =   "���"
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
         Caption         =   "��������"
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
         Caption         =   "��� ��������"
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
      Caption         =   "�����������"
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
         Alignment       =   2  '���������
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
         Alignment       =   2  '���������
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
         Alignment       =   2  '���������
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
         Alignment       =   2  '���������
         BackColor       =   &H00CFC2AC&
         Caption         =   "���"
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
         Caption         =   "���������"
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
         Caption         =   "                        �����"
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
         Caption         =   "������"
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
         Caption         =   "�������"
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
         Caption         =   "��������"
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
         Caption         =   "���"
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
      Caption         =   "&������"
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
      Alignment       =   2  '���������
      BackColor       =   &H00BBC0AC&
      Caption         =   "�"
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
      Caption         =   "���������:"
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
      Caption         =   "��� ����������"
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
      Alignment       =   2  '���������
      BackColor       =   &H00BBC0AC&
      Caption         =   "��"
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
      Caption         =   "� 5.2/"
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
      Caption         =   "���������"
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
      Alignment       =   2  '���������
      BackColor       =   &H00BBC0AC&
      Caption         =   "���� ������"
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
'����� "frmNewEF"
'���� �������� ������ ��� ������ � ���� ������ � ��� �������� ����� ����������
'���� ��������: 01.06.2016
'@version 1.0.0
'@author Andr.Nab.n@gmail.com
Option Explicit
'���������� �������
Public newDate As clmCaseDate      '��������� ������ "����"
Public newDOC As clmCreateDoc      '��������� ������ "����������� ���������"
Public SvcService As Object        '��������� ���������� svcsvc.dll
'�����
Public newEvForm As Form           '��������� frmEvidences

'Clases
Public newCase As clmCase          '��������� ������ "����"
Public newCor As clmHR_Official    '��������� ������ "�����������" Coroner
Public newExpert As String          '���������� "�������"
Public newExperts As clmHR_Official  '��������� ������ "��������" (������������/������� � �.�.)
Public newAExperts As clmHR_Official  '��������� ������ "����� ��������"
Public newEF As clmExpertFindings  '��������� ������ "���������� ��������"
Public newAEF As clmExpertFindings '��������� ������ "���������� ������ ��������"
Public newInjPr As clmEntrants     '��������� ������ "�����������"InjuredPerson
Public newBox As clmEvBox          '��������� ������ "������� � ��"
Public allEvSumCounter As clmCounter '��������� ������ ������� ����� ����� ��

'Key
Private mvarstrBoxKey As String '���� ��� �������� � �� ="BX****"
Private mvarstrEvKey As String  '���� ��� �������� (��)= "EV****"
Private mvarstrKey As String    '����� ���� = "BX****" + "EV****" ("BX****EV****")
'���������
Public colBoxes As New Collection      '��������� �������� � ���������
Public colExperts As New Collection    ' ��������� ���������
'������:
Const Msg As String = "������ ����� ������!"
'Constants
Private Const BOX As String = "BX"  '������� ����� ��� �������
Private Const EVID As String = "EV" '������� ����� ��� ��������.
'Colors
Private Enum frmColor
    ������_������ = &HC0FFFF 'RGB(102, 102, 153)
    ������ = &HFFFF&
    ������� = &HFF&
    ������ = &H0&
    ��������� = &HC0FFC0 'RGB(200, 256, 200)
    ������_������� = &HFEF7F1
    ���������� = &HFFC0C0
    ���������� = &H80FF&
End Enum

'������ "Cancel"
Dim Cancel As Integer
'
'++++++++++++++++++++++++++++++++++++  � � � � � � � � � � � � � +++++++++++++++++++++++++++++++++++
'
Public Static Sub Form_Initialize()
'������������� �����
End Sub
'
Private Sub Form_Load()
'�������� �����
'�������� �������
Set newCase = New clmCase           '��������� ������ "����
Set newDate = New clmCaseDate       '��������� ������ "����"
Set newCor = New clmHR_Official     '��������� ������ "�����������"
Call initClass
    Dim str As String
    With Me
        .Width = 12000
        .Height = 9000
        .txtCsIndex.Width = 520.932
        .txtCsIndex.Left = 1123.585
        .Caption = "�������� ������ ��������� " & "(" & mdPrintDoc.DocCat & ")"
        '& "����: " & .newExpert '" �������� ������ ��������� " & "����: " & .newExpert & .newEF.categories & ")"
        '���������� ����� "����������� ��������"
        .chkAddExData.Value = 0
        Call AE_Disabled
    End With
 Move (Screen.Width - Width) / 2, (Screen.Height - Height) / 3.7 '/ 17.7
End Sub
'
Private Sub Form_Unload(Cancel As Integer)
'��������� �������� �����
'    If MsgBox("�������?", vbYesNo, "�����?") = vbYes Then
        Set newCase = Nothing       '��������� ������ "����"
        Set newDate = Nothing       '��������� ������ "����"
        Set newCor = Nothing        '��������� ������ "�����������"
        Call termClass
'    Else
'        Cancel = 1
'    End If
End Sub
'
Private Sub initClass()
'������������� ��������� �������
    Set newEF = New clmExpertFindings   '��������� ������ "���������� ��������"
        newEF.evidensCategories = mdPrintDoc.DocCat '��������� ���������� �� ��
    Set newAEF = New clmExpertFindings   '��������� ������ "���������� ������ ��������"
'   Esperts
    Set newExperts = New clmHR_Official  '��������� ������ "��������"
    Set newAExperts = New clmHR_Official  '��������� ������ "����� ��������"
    Set colExperts = New Collection      '�������� ��������� "��������"
'   InjPr
    Set newInjPr = New clmEntrants    '��������� ������ "�����������"
'   ��
    Set colBoxes = New Collection     '��������� �������� � ���������
    Set newBox = New clmEvBox         '��������� ������ "������� � ��"
'Counter
    Set allEvSumCounter = New clmCounter  '��������� ������ ������� ����� ����� ��
         With allEvSumCounter
            .name = "����� ����� ��"
'            Debug.Print .name & Chr(32) & .getTale
          End With
    Set mdPrintDoc.colTextDoc = Nothing
    Call addExpert
End Sub
'
Private Sub termClass()
'����������� ��������� ������������ �������
    Set newEF = Nothing         '��������� ������ "���������� ��������"
    Set newAEF = Nothing        '��������� ������ "���������� ������ ��������"
    Set newExperts = Nothing    '��������� ������ "��������"
    Set newAExperts = Nothing   '��������� ������ "����� ��������"
    Set colExperts = Nothing    '�������� ��������� "��������"
    Set newInjPr = Nothing      '��������� ������ "�����������"
'   ��
    Set colBoxes = Nothing      '��������� �������� � ���������
    Set newBox = Nothing        '��������� ������ "������� � ��"
    Set allEvSumCounter = Nothing  '��������� ������ ������� ����� ����� ��
'    Set newDOC = Nothing       '��������� ������ "����������� ���������"
End Sub
'
'+++++++++++++++++++++++++++++++++++++++ � � � � � �  � � � � �: ++++++++++++++++++++++++++++++++++++
'1) ����� ����������� ����
Public Sub show_MsgCreateNewBox()
'����� ����������� ���� "������� ����� ������� � ��?"
    If MsgBox("������� ����� �������� � ��������� (��)?", vbYesNo) = vbYes Then
        Call Open_EvidForm  '��������� �������� ����� ����� "������������ ��������������"
'       �������� ����� frmEvidences
        
    Else
        MsgBox "���� �������� � ��������� (��) ��������!", vbOKOnly
         mdPrintDoc.colExcelData.Add Me.allEvSumCounter.getTale, "J" '��� ������ � Excel
            Debug.Print "������ ����� �� � ��������� colExcelData -> " & mdPrintDoc.colExcelData.Item("J")
        Call Open_DocListForm
    End If
End Sub
'
'2) �������� ����� �������
Public Sub addNewBox()
  '��������� �������� ����� ������� � ��
    With Me
        strBoxKey = BOX & CStr(Format(.colBoxes.Count + 1, "#0000"))
        Set .newBox = New clmEvBox
        .newBox.strBxName = .strBoxKey
    End With
End Sub
'
Public Sub Open_EvidForm()
'��������� �������� ����� ����� "������������ ��������������"
Set newEvForm = New frmEvidences
    With newEvForm
        .Show
    End With
Me.Visible = False
End Sub
'
Public Sub Open_DocListForm()
'��������� �������� ����� frmDocuments "�������� ����������"
    Dim frmD As Form
'    Set frmD = New frmDocuments
        With frmDocuments
            .Show
        End With
Me.Visible = False
End Sub
'
Private Sub chkAddExData_Click()
'������� �� ����� "�������� ������"
    With chkAddExData
        If .Value = 1 Then
            Call AE_Enabled
            .BackColor = &H80FFFF
            '��������� ����� "���������� ������ ��������"
            Set newAEF = New clmExpertFindings
                With newAEF
                    .name = "�������-�����������"
                    .definition = newEF.definition
                    .tanatology = True
                    .condition = "� ������"
                    .evidensCategories = "���������� �����"
                End With
        Else
            Call AE_Disabled
            .BackColor = &HDDEADB
            '����������� ������ "���������� ������ ��������"
            Set newAEF = Nothing
        End If
    End With
End Sub
'
Private Sub AE_Disabled()
'��������� ���������� ����� ����� �������
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
'����������� ������ "���������� ������ ��������"
Set newAEF = Nothing
End Sub
'
Private Sub AE_Enabled()
'��������� ��������� ����� ����� �������
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
'����������� ��������� ���������� ���������� = �� ��������� �������������/�����������
'��������� ����������� + ������ + ����� + ��� + ����������� ���� �� ���������� ���������� ����
    create_legalGround = newEF.printDefinition & newCor.print_Cor & _
    ", ����������� " & newEF.rulingDate & newCase.printCsData & "."
Debug.Print "��.��������� = " & create_legalGround
End Function
'
Public Static Sub addExpert()
'��������� ���������� ��������
'1)������� �����:
Me.Visible = False
'������� ������ �������� �� ������
Set SvcService = CreateObject("Svcsvc.Service") '������ ���������� svcsvc.dll
    Do
        newExpert = SvcService.SelectValue("��������� �.�." & vbCrLf & _
                                        "������� �.�." & vbCrLf & _
                                        "�������� �.�." & vbCrLf & _
                                        "����� �.�." & vbCrLf & _
                                        "���� �.�." & vbCrLf & _
                                        "������� �.�." & vbCrLf & _
                                        "�������� �.�." & vbCrLf & _
                                        "�������� �.�." & vbCrLf, _
                                        "�������� ��������", True)
'Debug.Print " addExpert: " & newExpert
        If newExpert <> "" Then
            '������������ ������� ������� ���������:
            mdPrintDoc.arrExpert = Split(newExpert, vbCrLf)
            Exit Do
        Else
        '���� ������� �� ������:
            MsgBox "C������ ������� ������� ��������!", vbCritical, Msg
        End If
    Loop
'������� �������
'Dim i As Byte
'For i = LBound(arrExpert) To UBound(arrExpert)
'Debug.Print "������: " & i & " - " & arrExpert(i)
'Next i
Set SvcService = Nothing
End Sub
'
Private Sub txtEnabled()
'���������� ��������� �����
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
'+++++++++++++++++++++++++++++++++++++++ � � � �  � � � � �: ++++++++++++++++++++++++++++++++++++++++++++++++++
'
Private Sub txtEnter(tmpObj As Object)
'��������� �������� ����� ��� ��������� ������
    With tmpObj
        .Text = ""
        .BackColor = frmColor.������_������ 'RGB(102, 102, 153) 'frmColor.����1
        .ForeColor = frmColor.������
    End With
End Sub
'
 Private Sub txt_Exit(tmpObj As Object)
'��������� �������� ����� ��� ������ ������
    With tmpObj
        .BackColor = frmColor.����������
        If tmpObj.name = "txtCsIndex" Or _
            tmpObj.name = "txtCsNum" Or _
            tmpObj.name = "txtCrOprAr" Or _
            tmpObj.name = "txtInjPrBirthday" _
            Then
                .Text = "�� ������"
        ElseIf tmpObj.name = "txtCsPrvDate" Or _
            tmpObj.name = "txtCrPost" Or _
            tmpObj.name = "txtCrSurName" Or _
            tmpObj.name = "txtInjPrSurName" Or _
            tmpObj.name = "txtInjPrDecease" Or _
            tmpObj.name = "txtInjPrAutopsy" _
            Then
                .Text = "�� �������"
        ElseIf tmpObj.name = "txtCrName" Or _
            tmpObj.name = "txtCrMidName" Or _
            tmpObj.name = "txtInjPrName" Or _
            tmpObj.name = "txtInjPrMidName" _
            Then
                .Text = "�� �������"
        Else
            .Text = "�� �������"
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
        .BackColor = RGB(200, 256, 200) 'frmColor.���������
        newAEF.EFcode = .Text
        End If
    End With
End Sub
'
Private Sub txtAEMidName_GotFocus()
'�������� ������ ��������
Call txtEnter(txtAEMidName)
End Sub

Private Sub txtAEMidName_LostFocus()
'�������� ������ ��������
 With txtAEMidName
        If .Text = "" Then
             Call txt_Exit(txtAEMidName)
        Else
             If Len(txtAEMidName.Text) = 1 Then
                newAExperts.midName = newCor.create_Initials(.Text)
            Else
                newAExperts.midName = StrConv(.Text, vbProperCase)
            End If
            .BackColor = RGB(200, 256, 200) 'frmColor.���������
            .Text = newAExperts.midName
        End If
    End With
End Sub
'
Private Sub txtAEName_GotFocus()
'��� ������ ��������
    Call txtEnter(txtAEName)
End Sub
'
Private Sub txtAEName_LostFocus()
'��� ������ ��������
 With txtAEName
        If .Text = "" Then
             Call txt_Exit(txtAEName)
        Else: newAExperts.name = StrConv(.Text, vbProperCase)
         '���������� ����� ����� ��������
            If Len(txtAEName.Text) = 1 Then
                newAExperts.name = newCor.create_Initials(.Text)
            Else: newAExperts.name = StrConv(.Text, vbProperCase)
            End If
            .BackColor = RGB(200, 256, 200) 'frmColor.���������
            .Text = newAExperts.name
        End If
    End With
End Sub

Private Sub txtAESurName_GotFocus()
'������� ������ ��������
 Call txtEnter(txtAESurName)
End Sub

Private Sub txtAESurName_LostFocus()
'������� ������ ��������
    With txtAESurName
        If .Text = "" Then
             Call txt_Exit(txtAESurName)
        Else: newAExperts.surName = StrConv(.Text, vbProperCase)
            .BackColor = RGB(200, 256, 200) 'frmColor.���������
            .Text = newAExperts.surName
        End If
    End With
End Sub
'
Private Sub txtAEFNum_GotFocus()
'����� ����������  ������ ��������
    Call txtEnter(txtAEFNum)
End Sub
'
Private Sub txtAEFNum_LostFocus()
'����� ����������  ������ ��������
With txtAEFNum
        If .Text = "" Then
             Call txt_Exit(txtAEFNum)
        Else: newAEF.number = StrConv(txtAEFNum.Text, vbProperCase)
            .BackColor = RGB(200, 256, 200) 'frmColor.���������
'            .Text = newAExperts.surName
        End If
    End With
End Sub
'
Private Sub txtAutopsyDate_GotFocus()
'���� ��������
   Call txtEnter(txtAutopsyDate)
   '���� �������� = ���� ������
    txtAutopsyDate.Text = newDate.dateToString(newEF.rulingDate)
End Sub
'
Private Sub txtAutopsyDate_LostFocus()
'���� ������ ����� ���������� <=  ���� ��������
 Dim tmp As Double, dt As Date, str As String
    If txtAutopsyDate.Text = "" Then
        Call txt_Exit(txtAutopsyDate)
    Else
        With newDate
            dt = .validateDate(.ExamDate(txtAutopsyDate.Text))
'            '��������� ���: ���� �������� >= ���� ������ ����� ����������
            Do
                tmp = .compareDt(newEF.rulingDate, dt)
                If tmp < 0 Then
                    MsgBox "���� ������ ����� ���������� ������ ���� ��������� �������������!", _
                            vbCritical, Msg
                    dt = .ExamDate(InputBox("��������� ������� ���� ������ ���������� �����!", _
                                        "���� ����", newEF.rulingDate))
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
'    '�������� ������ ����
Me.cmdOK1.SetFocus
'    With Me.txtFactCase
''        Call txtEnter(txtFactCase)
''        .Text = newDate.dateToString(newEF.rulingDate) & " ���� " & newInjPr.create_InitialslName
''    End With
''Debug.Print "���� ��������: " & newInjPr.autopsyDate
''On Error Resume Next ' ��������� ������
End Sub
'
'1) ����� ����������
Private Static Sub txtDN_GotFocus()
'����� ����������
    Call txtEnter(txtDN)
 End Sub
'
Private Static Sub txtDN_LostFocus()
'����� ����������
    With txtDN
        Do
            If Not IsNumeric(.Text) Or Len(.Text) = 0 Then '���� �� ����� ��� ������ ����
                Beep
                .BackColor = RGB(255, 0, 0) 'frmColor.�������
                MsgBox "C������ ������� �����", vbCritical, Msg
                .Text = InputBox("������� ��������� ����� ����������!", "����������� ������ �����")
                 If VarType(.Text) = vbBoolean Then Exit Sub    ' ������ ������ ������
                ElseIf CInt(.Text) <= 0 Then '���� ������ ���� ��� ����
                 Beep
                .BackColor = RGB(255, 0, 0) 'frmColor.�������
                MsgBox "����� ���������� ������ ���� ������ ����!", vbCritical, Msg
                .Text = InputBox("������� ��������� ����� ����������!", "����������� ������ �����")
'                 If VarType(.Text) = vbBoolean Then Exit Sub    ' ������ ������ ������
            Else
                newEF.number = .Text
                .BackColor = RGB(200, 256, 200) 'frmColor.���������
        Exit Do
            End If
        Loop
    End With
    
    With Me
        .Caption = "���������� " & newEF.getFullNumber & Chr(32) & mdPrintDoc.DocCat
'        mdPrintDoc.DocCat & " �" & newEF.getFullNumber '" �������� ������ ��������� �" & newEF.getFullNumber "����: " & Me.newExpert &
        With .txtEFFirstDay
            .SetFocus '�������� ������
            Call txtEnter(Me.txtEFFirstDay)
            .Text = newDate.DateNow
        End With
    End With
'Debug.Print "����� ����������: �" & newEF.number
End Sub
'2) ���� ������ ����������
Private Sub txtEFFirstDay_GotFocus()
''���� ������ ����������
End Sub
'
Private Sub txtEFFirstDay_LostFocus()
'���� ������ ���-���� ����������(5), ������ ��� ����� ������ �����:
'�������� ������, �������, �� ������������ � ������� ������ (�.�. ��� ��� �� �������)!!!
'����������� ��������� ���:
'                              ���� ������(1)
'                           <= ���� ��������� �������������(2)
'                           <= ���� ������ ���������� �����(3)
'                           <= ���� �������� �����(4)
'                           <= ���� ������ ���-���� ����������(5)
    Dim tmp As Double, dt As Date, str As String
    If txtEFFirstDay.Text <> "" Then
        With newDate
            dt = .validateDate(.ExamDate(txtEFFirstDay.Text))
        End With
    Else '���� ���� �� �������:
            With txtEFFirstDay
                Do '������� ������������� ����� ����:
                    If .Text = "" Then '����  ������ ����
                        Beep
                    .BackColor = RGB(255, 0, 0) 'frmColor.�������
                        MsgBox "C������ ������ ���� ������ ����������!", vbCritical, Msg
                        .Text = InputBox("������� ��������� ���� ������ ����������!", "����������� ������ �����", newDate.DateNow)
'                 If VarType(.Text) = vbBoolean Then Exit Sub    ' ������ ������ ������
                    Else
                        dt = newDate.validateDate(newDate.ExamDate(.Text))
                        .BackColor = RGB(200, 256, 200) 'frmColor.���������
                        Exit Do
                    End If
                Loop
            End With
    End If
    '� �������:
    With newEF
        '���� ������
        .firstDay = dt
        '����
        .dueDate = newDate.getPeriod(.firstDay, 29)
         mdPrintDoc.strMonth = Me.newDate.monthOfYear(.firstDay) '�������� �������� ������ ��� ������ � Excel
    End With
    '� ��������� �����:
    With txtEFFirstDay
        .BackColor = frmColor.���������
        .Text = newDate.dateToString(newEF.firstDay)
    End With
'Debug.Print "���� ������: " & newEF.firstDay
'Debug.Print "����: " & newEF.dueDate
End Sub
'
'3)��������� ����������
Private Sub cboEFCategories_GotFocus()
'��������� ���������� (���������, �������������� � �.�.)
    With cboEFCategories
        .Clear
        .BackColor = frmColor.������_������
        .AddItem "�������-�����������"
        .AddItem "�������� ������-������������������"
        .AddItem "�����������"
        .AddItem "��������������"
        .AddItem "���������"
        .AddItem "������������"
    End With
End Sub
'
Private Sub cboEFCategories_LostFocus()
'��������� ���������� (���������, �������������� � �.�.)
    With cboEFCategories
'���������� ���������, �������� � ��������:
        If .Text = "�����������" Then
            '������� �������� ����� ��� ����� �������� ���������
           '������� ������ �������� �� ������
            Set SvcService = CreateObject("Svcsvc.Service") '������ ���������� svcsvc.dll
            Do
                newEF.name = SvcService.SelectValue("�������-�������������" & vbCrLf & _
                                        "�������-������������" & vbCrLf & _
                                        "������-������������������" & vbCrLf & _
                                        "������������������" & vbCrLf & _
                                        "���������������" & vbCrLf & _
                                        "���������� ��������� ������" & vbCrLf & _
                                        "�������� ���������� ����������� ���������� � ������� �� ���" & vbCrLf & _
                                        "�������-�����������" & vbCrLf & _
                                        "�������-���������������" & vbCrLf & _
                                        "�������-����������", _
                                        "�������� ����������:", True)
                                        Debug.Print " addnewEFname: " & newEF.name
                If newEF.name <> "" Then
                '������������ ������� �������� ���������
                    mdPrintDoc.arrEF = Split(newEF.name, vbCrLf)
                    Exit Do:
                Else '���� ���������� �� �������:
                    MsgBox "C������ ������� ����������!", vbCritical, Msg
                End If
            Loop
'������� �������
'Dim i As Byte
'For i = LBound(arrEF) To UBound(arrEF)
'Debug.Print "������: " & i & " - " & arrEF(i)
'Next i
            Set SvcService = Nothing
        End If
        '�������� �� �������� ����
        If .Text = "" Then
            .Text = "�������� ������-������������������"
            mdPrintDoc.arrEF = Split(.Text)
        End If
'       �������� ������ � ������������� ���������:
        If newEF.name <> "" Then
            newEF.categories = .Text & " (" & Join(mdPrintDoc.arrEF, ", ") & ") "
        Else: newEF.categories = .Text
        End If
            .BackColor = frmColor.��������� 'RGB(200, 256, 200)
    End With
Debug.Print "��������� ����������= ", newEF.categories
End Sub
'
'4)��������� ���������� ����������
Private Sub cboCsDefinition_GotFocus()
'��������� ���������� ����������
    With cboCsDefinition
        .Clear
        .BackColor = frmColor.������_������
        .AddItem "�������������"
        .AddItem "�����������"
        .AddItem "�� ������� ������"
    End With
End Sub
'
Private Sub cboCsDefinition_LostFocus()
'��������� ���������� ����������
    With cboCsDefinition
        If .Text = "" Then
            .Text = "�������������"
        End If
        newEF.definition = .Text '���������� ��������� ��������
        .BackColor = frmColor.��������� '����������� ������������ ���� ��� ������ �� ������
    End With
'Debug.Print "���������: " & newEF.definition
End Sub
'
'5) ���� ��������� �������������
Private Static Sub txtCsPrvDate_GotFocus()
'���� ��������� �������������
    Call txtEnter(txtCsPrvDate)
    '���� ��������� ������������� = (������� ���� - 1):
    txtCsPrvDate.Text = CStr(newEF.firstDay - 1)
End Sub
'
Private Static Sub txtCsPrvDate_LostFocus()
'���� ��������� �������������(2)
'              <= ���� ������ ���������� �����(3)
'              <= ���� �������� �����(4)
'              <= ���� ������ ���-���� ����������(5)
'

 Dim tmp As Double, dt As Date
 With txtCsPrvDate
    If .Text = "" Then
        Call txt_Exit(txtCsPrvDate)
    Else
        With newDate
            dt = .validateDate(.ExamDate(txtCsPrvDate.Text)) 'newEF.rulingDate
            '��������� ���: ���� ������ ���-���� ����������(5)>= ���� ��������� �������������(2)
            Do
                tmp = .compareDt(newEF.firstDay, dt)
                If tmp > 0 Then
                    MsgBox "���� ��������� ������������� ������ ���� ������ ����������!", _
                            vbCritical, Msg
                    dt = .ExamDate(InputBox("��������� ������� ���� ��������� �������������!", _
                                        "���� ����", newEF.firstDay - 1))
                Else
                    newEF.rulingDate = dt
                    Exit Do
                End If
            Loop
        End With
            'With txtCsPrvDate
            .BackColor = RGB(200, 256, 200) 'frmColor.���������
            .Text = newDate.dateToString(newEF.rulingDate)
    End If
End With
'Debug.Print "���� ��������� �������������: " & newEF.rulingDate
End Sub
'
'6) ��������� ����
Private Sub cboCsCat_GotFocus()
'��������� ����
With cboCsCat
    .Clear
        .BackColor = frmColor.������_������  '&HC0FFFF
'�������� ������ � ��������������� ����
        .AddItem "���������� ����"
        .AddItem "��������"
        .AddItem "����������������� ����"
    End With
End Sub
'
Private Sub cboCsCat_LostFocus()
'��������� ����
With cboCsCat
    If Len(.Text) = 0 Then
            .Text = "(��������� ���� �� �������)"
           .BackColor = frmColor.���������� ' &HC0E0FF
    Else:
        If .Text = "��������" Then
            txtCsIndex.Visible = True
            txtCsIndex.SetFocus
        Else:
            txtCsIndex.Visible = False
            txtCsNum.SetFocus
        End If
            txtCsNum.Visible = True
             newCase.category = .Text
            .BackColor = frmColor.��������� 'RGB(200, 256, 200)
    End If
    End With
'Debug.Print "��������� ����: " & newCase.category
End Sub
'
'7)������ ����
Private Sub txtCsIndex_GotFocus()
'������ ����
 Call txtEnter(txtCsIndex)
    With txtCsIndex
        .Width = 520.932
        .Left = 1123.585
    End With
    newCase.index = ""
End Sub
'
Private Sub txtCsIndex_Change()
'������ ����
If txtCsIndex <> "�� ������" Then
    Dim str As String
    With txtCsIndex
        str = Right(.Text, 1)
        newCase.index = newCase.index + StrConv(str, vbProperCase)
    End With
End If
End Sub
'
Private Sub txtCsIndex_LostFocus()
'������ ����
    With txtCsIndex
        If Len(.Text) = 0 Then
            .Width = 1398.29
            .Left = 246.226
            Call txt_Exit(txtCsIndex)
        Else: 'newCase.index = StrConv(.Text, vbProperCase)
            .BackColor = frmColor.��������� 'RGB(200, 256, 200)
            .Text = newCase.index
        End If
    End With
Debug.Print "������ ����: ", newCase.index
End Sub
'
'8) ����� ����
Private Static Sub txtCsNum_GotFocus()
'����� ����
    Call txtEnter(txtCsNum)
    txtCsNum.MaxLength = 11
End Sub
'
Private Static Sub txtCsNum_LostFocus()
'����� ���� CsNum
With txtCsNum
    If Len(.Text) = 0 Then
            Call txt_Exit(txtCsNum)
    Else
    '1) ���� �� ��������� ����
        If newCase.category <> "���������� ����" Then
            newCase.number = .Text
            .BackColor = RGB(200, 256, 200)
        Else '2) ���� ��������� ����:
             Do
                If Not IsNumeric(.Text) Or Len(.Text) = 0 Then 'anee ia oeo?u eee ionoia iiea
                   Beep
                    .BackColor = RGB(255, 0, 0)
                    MsgBox "������� ������� �����!", vbCritical, Msg
                        .Text = InputBox("��������� ������� ����� ���������� ����!", "����������� ������ ����� ������!")
'                    If VarType(.Text) = vbBoolean Then Exit Do    '������ ������ "������"
                    ElseIf Len(.Text) <> 11 Then '
                        MsgBox "����� ���� ������ �������� �� 11 ����!", vbCritical, Msg
                        .Text = InputBox("��������� �������  ����� ����!", "����������� ������ �����")
'                    If VarType(.Text) = vbBoolean Then Exit Do    '������ ������ "������"
                   
                    End If
                   Exit Do
                Loop
        End If
    End If
    If .Text = "�� ������" Then
        .BackColor = frmColor.����������
    Else
        newCase.number = .Text
        .BackColor = RGB(200, 256, 200) 'frmColor.���������
    End If
End With
'Debug.Print "����� ����: �" & newCase.number
End Sub
'
'9)��������� �����������
Private Static Sub txtCrPost_GotFocus()
'��������� �����������
     Call txtEnter(txtCrPost)
     txtCrPost.Text = "�������� �����������"
End Sub
'
Private Static Sub txtCrPost_LostFocus()
'��������� �����������
    With txtCrPost
        If .Text = "" Then
           Call txt_Exit(txtCrPost)
        Else: newCor.post = .Text
            .BackColor = RGB(200, 256, 200) 'frmColor.���������
        End If
    End With
'Debug.Print "��������� �����������= " & newCor.post
End Sub
'
'10) ����� (�����������)
Private Static Sub txtCrOprAr_GotFocus()
'����� (�����������)
    Call txtEnter(txtCrOprAr)
    txtCrOprAr.Text = "(�. ������) ��������� ������ ������������� �������� ���������� ��������"
End Sub
'
Private Static Sub txtCrOprAr_LostFocus()
'����� (�����������)
    With txtCrOprAr
        If .Text = "" Then
           Call txt_Exit(txtCrOprAr)
        Else: newCor.conformation = .Text
            .BackColor = RGB(200, 256, 200) 'frmColor.��������� 'RGB(200, 256, 200)
        End If
    End With
'Debug.Print "����� (�����������): " & newCor.conformation
End Sub
'
'11) ������ �����������
Private Static Sub txtCrRank_GotFocus()
'������ �����������
    Call txtEnter(txtCrRank)
    txtCrRank.Text = "�������� ���������� �������"
End Sub
'
Private Static Sub txtCrRank_LostFocus()
'������ �����������
    With txtCrRank
        If .Text = "" Then
            Call txt_Exit(txtCrRank)
        Else: newCor.rank = .Text
            .BackColor = RGB(200, 256, 200) 'frmColor.��������� 'RGB(200, 256, 200)
        End If
    End With
'Debug.Print "������ �����������: " & newCor.rank
End Sub
'
'12)������� �����������
Private Static Sub txtCrSurName_GotFocus()
'������� �����������
    Call txtEnter(txtCrSurName)
End Sub
'
Private Static Sub txtCrSurName_LostFocus()
'������� �����������
    With txtCrSurName
        If .Text = "" Then
            Call txt_Exit(txtCrSurName)
        Else: newCor.surName = StrConv(.Text, vbProperCase)
            .BackColor = RGB(200, 256, 200)
            .Text = newCor.surName
        End If
    End With
'Debug.Print "������� �����������: " & newCor.surName
End Sub
'
'13) ��� �����������
Private Static Sub txtCrName_GotFocus()
'��� �����������
    Call txtEnter(txtCrName)
End Sub
'
Private Static Sub txtCrName_LostFocus()
'��� �����������
    With txtCrName
        If .Text = "" Then
            Call txt_Exit(txtCrName)
        Else
    '���������� ����� ����� ��������
            If Len(txtCrName.Text) = 1 Then
                newCor.name = newCor.create_Initials(.Text)
            Else: newCor.name = StrConv(.Text, vbProperCase)
            End If
        .Text = newCor.name
        .BackColor = RGB(200, 256, 200) 'frmColor.��������� '����������� ������������ ���� ��� ������ �� ������
        End If
    End With
'Debug.Print "��� �����������: " & newCor.name
End Sub
'
'14)�������� �����������
Private Static Sub txtCrMidName_GotFocus()
'�������� �����������
   Call txtEnter(txtCrMidName)
End Sub
'
Private Static Sub txtCrMidName_LostFocus()
'�������� �����������
    With txtCrMidName
        If .Text = "" Then
           Call txt_Exit(txtCrMidName)
        Else
            If Len(txtCrMidName.Text) = 1 Then
                newCor.midName = newCor.create_Initials(.Text)
            Else: newCor.midName = StrConv(.Text, vbProperCase)
            End If
            .BackColor = RGB(200, 256, 200) 'frmColor.���������
            .Text = newCor.midName
        End If
    End With
'Debug.Print "�������� �����������: " & newCor.midName
End Sub
'
'15) ��� �����������
Private Sub lstCrSex_GotFocus()
'��� �����������
    With lstCrSex
        .Clear
        .AddItem "���."
        .AddItem "���."
        .AddItem "��� �� ������"
        .BackColor = &HC0FFFF
    End With
End Sub
'
Private Sub lstCrSex_LostFocus()
'��� �����������
    With lstCrSex
        newCor.sex = .Text
        .BackColor = frmColor.���������
    End With
Debug.Print "��� �����������: " & newCor.sex
End Sub

'16)������� ������������
Private Static Sub txtInjPrSurName_GotFocus()
'������� ������������
   Call txtEnter(txtInjPrSurName)
End Sub
'
Private Static Sub txtInjPrSurName_LostFocus()
'������� ������������
    With txtInjPrSurName
        If .Text = "" Then
             Call txt_Exit(txtInjPrSurName)
        Else: newInjPr.surName = StrConv(.Text, vbProperCase)
            .BackColor = RGB(200, 256, 200) 'frmColor.���������
            .Text = newInjPr.surName
        End If
    End With
'Debug.Print "������� ������������: " & newInjPr.surName
End Sub
'
'17)
Private Static Sub txtInjPrName_GotFocus()
'��� ������������
   Call txtEnter(txtInjPrName)
End Sub
'
Private Static Sub txtInjPrName_LostFocus()
'��� ������������
    With txtInjPrName
        If .Text = "" Then
             Call txt_Exit(txtInjPrName)
        Else
            If Len(.Text) = 1 Then
                newInjPr.name = newInjPr.create_Initials(.Text)
            Else: newInjPr.name = StrConv(.Text, vbProperCase)
            End If
        .Text = newInjPr.name
        .BackColor = RGB(200, 256, 200) 'frmColor.���������
        End If
    End With
'Debug.Print "��� ������������: " & newInjPr.name
End Sub
'
'18)
Private Static Sub txtInjPrMidName_GotFocus()
'�������� ������������
    Call txtEnter(txtInjPrMidName)
End Sub
'
Private Static Sub txtInjPrMidName_LostFocus()
'�������� ������������
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
'Debug.Print "�������� ������������: " & newInjPr.midName
End Sub
'
'19)
Private Sub lstInjPrSex_GotFocus()
 '��� ������������
    With lstInjPrSex
        .Clear
        .AddItem "���."
        .AddItem "���."
        .AddItem "��� �� ������"
        .BackColor = &HC0FFFF
    End With
End Sub
'
Private Sub lstInjPrSex_LostFocus()
'��� ������������
    With lstInjPrSex
        newInjPr.sex = .Text
        .BackColor = RGB(200, 256, 200)
    End With
'Debug.Print "��� ������������: " & newInjPr.sex
End Sub
'
'20)'��� �������� ������������
Private Static Sub txtInjPrBirthday_GotFocus()
'��� �������� ������������
   Call txtEnter(txtInjPrBirthday)
End Sub
'
Private Static Sub txtInjPrBirthday_LostFocus()
'��� �������� ������������
    With txtInjPrBirthday
        If .Text = "" Then
             Call txt_Exit(txtInjPrBirthday)
        Else: newInjPr.birthday = .Text
            .BackColor = RGB(200, 256, 200)
        End If
    End With
'Debug.Print "��� �������� ������������: " & newInjPr.birthday
End Sub
'
'21) ���� ������
Private Sub txtInjPrDecease_GotFocus()
'���� ������
   Call txtEnter(txtInjPrDecease)
   '���� ������ = ���� ��������� �������������
   txtInjPrDecease.Text = newAEF.firstDay - 1
End Sub

Private Sub txtInjPrDecease_LostFocus()
'���� ������(1)<= ���� ��������� �������������(2)
'              <= ���� ������ ���������� �����(3)
'              <= ���� �������� �����(4)
'              <= ���� ������ ���-���� ����������(5)
'
 Dim tmp As Double, dt As Date ' str As String
    If txtInjPrDecease.Text = "" Then
        Call txt_Exit(txtInjPrDecease)
    Else
        With newDate
            dt = .validateDate(.ExamDate(txtInjPrDecease.Text))
            '��������� ���: ���� ������(1) <= ���� ��������� �������������(2)
            Do
                tmp = .compareDt(newAEF.firstDay, dt)
                If tmp > 0 Then
                    MsgBox "���� ������ ������ ���� ��������� �������������!", _
                            vbCritical, Msg
                    dt = .ExamDate(InputBox("��������� ������� ���� ������!", _
                                        "���� ����", newAEF.firstDay - 1))
                Else
                    newInjPr.decease = dt
                    Exit Do
                End If
            Loop
        End With
        '� ��������� �����:
        With txtInjPrDecease
            .BackColor = RGB(200, 256, 200)
            .Text = newDate.dateToString(newInjPr.decease)
        End With
    End If
    '�������� ������ ���������� ����
    With Me.txtInjPrAutopsy
        .SetFocus
'        Call txtEnter(Me.txtInjPrAutopsy)
'        .Text = newAEF.firstDay
    End With
'Debug.Print "���� ������: " & newInjPr.decease
End Sub
''
'22)���� ��������
Private Sub txtInjPrAutopsy_GotFocus()
'���� ��������
   Call txtEnter(txtInjPrAutopsy)
   '���� �������� = ���� ������
   txtInjPrAutopsy.Text = newAEF.firstDay 'newInjPr.autopsyDate 'newEF.rulingDate
End Sub
'
    Private Sub txtInjPrAutopsy_LostFocus()
'���� �������� �����(4)
 Dim tmp As Double, dt As Date, str As String
    If txtInjPrAutopsy.Text = "" Then
        Call txt_Exit(txtInjPrAutopsy)
    Else
        With newDate
            dt = .validateDate(.ExamDate(txtInjPrAutopsy.Text))
            '��������� ���: ���� �������� �����(4) >= ���� ��������� �������������
            Do
                tmp = .compareDt(newAEF.firstDay, dt)
                If tmp < 0 Then
                    MsgBox "���� �������� ������ ���� ������ ���������� �����!", _
                            vbCritical, Msg
                    dt = .ExamDate(InputBox("��������� ������� ���� ��������!", _
                                        "���� ����", newAEF.firstDay))
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
    '�������� ������ ����
'    With Me.txtFactCase
'        Call txtEnter(txtFactCase)
'        .Text = newDate.dateToString(newEF.rulingDate) & " ���� " & newInjPr.create_InitialslName
'    End With
'Debug.Print "���� ��������: " & newInjPr.autopsyDate
'On Error Resume Next ' ��������� ������
End Sub
'
'+++++++++++++++++++++++++++++++++++++++ � � � � � �: ++++++++++++++++++++++++++++++++++++++++++++++++++
Private Sub cmdOK1_Click()
'������� ������� ������ "��" ����� "frmNewEF"
'1) ���������� ��������� ��� ������ � Excel:
    With mdPrintDoc.colExcelData
        .Add txtDN.Text, "A" '������ ���
            Debug.Print "������ ������ ��� � ��������� colExcelData -> " & .Item("A")
        .Add txtEFFirstDay.Text, "E"   '���� ������
            Debug.Print "������ ���� ������ � ��������� colExcelData -> " & .Item("E")
        .Add newDate.dateToString(newEF.dueDate), "G" '����
            Debug.Print "������ ���� � ��������� colExcelData -> " & .Item("G")
        .Add cboCsCat.Text, "C" '��������� ����
            Debug.Print "������ ��������� ���� � ��������� colExcelData -> " & .Item("C")
        .Add txtCsNum.Text, "D" '������ ����
            Debug.Print "������ ������ ���� � ��������� colExcelData -> " & .Item("D")
    End With
'2) �������� �����, ���������� �������� � �������� �����������
    Call mdMainFolders.makeDocDir(Me.newEF.number, mdPrintDoc.DocCat)
    Call show_MsgCreateNewBox
End Sub
'
Public Sub cmdCancel1_Click()
'������ "������"
    Unload Me
End Sub
'
Private Sub txtDisabled()
'���������� ��������� �����
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
'������ "��������"
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
    .Caption = "�������� ������ ���������"
    .txtDN.SetFocus
End With
End Sub
'
Private Sub cmdSaveCreate_Click()
'��������� ������� ������ "C������ � ����� �� ����������"
    Call txtDisabled '���������� ������ � ���������� �����
    Call termClass '����������� ������������ ����������� ������� � ������������ ������
    Call initClass '�������� ����� ����������� �������
'��������� ������� �����
With Me
    .Caption = "����: " & .newExpert & " �������� ������ ���������"
    .txtDN.SetFocus
End With
End Sub
'
'================== I C A P S U L A T I O N =============================
'
Private Property Let strBoxKey(ByVal vData As String)
'���� ��� ��������
     mvarstrBoxKey = vData
End Property
'
Friend Property Get strBoxKey() As String
'���� ��� ��������
strBoxKey = mvarstrBoxKey
'     Dim str As String
'        If mvarstrBoxKey <> "" Then
'            str = BOX & mvarstrBoxKey
'        End If
'     strBoxKey = str
'Debug.Print "���� ��� �������� = " & strBoxKey
End Property
'
Friend Property Let strEvKey(ByVal vData As String)
'���� ��� �������� � ��������
    mvarstrEvKey = CStr(Format(vData, "#0000"))
'    Debug.Print "���� ��� �������� = " & frmNewEF.strEvKey
End Property
'
Friend Property Get strEvKey() As String
'���� ��� �������� � ��������
     Dim str As String
        If mvarstrEvKey <> "" Then
            str = EVID & mvarstrEvKey
        End If
     strEvKey = str
Debug.Print "���� ��� �������� � �������� = " & strEvKey
End Property
'
Friend Property Get strKey() As String
'����� ���� = "BX****" + "EV****" ("BX****EV****")
     Dim str As String
        If strBoxKey <> "" And strEvKey <> "" Then
            str = strBoxKey & strEvKey
        End If
     strKey = str
Debug.Print "���� ��� �������� � �������� = " & strKey
End Property



'Private Sub debugPrint_colBoxes()
''���������� ������ ������ ��������� �������
'Dim tmpBx As clmEvBox
'Dim i As Long, bxKey As String
' With frmNewEF.colBoxes
'    For i = 1 To .Count
'        bxKey = BOX & CStr(Format(i, "#0000"))
'        Set tmpBx = .Item(bxKey)
'            Debug.Print "���������� ������ �������_" & bxKey
'            Debug.Print "�������� ������� ->" & tmpBx.strBxName
'            Debug.Print "����� ������� �� -> " & tmpBx.strBxPlace
'            Debug.Print "�������������� �� -> " & tmpBx.strBxEntrants
'            Debug.Print "�������� -> " & tmpBx.strBxPackage
'            Debug.Print "��������� -> " & tmpBx.strBxCategory
'            Debug.Print "�������� ������� -> " & tmpBx.strBxObjDescription
''           ��
'            Dim ev As Variant
'            For Each ev In tmpBx.colEvidences
'                Debug.Print "���������� �������: "
'                Debug.Print "�� - > " & ev
'            Next ev
'         Set tmpBx = Nothing
'        Next i
'    End With
'End Sub

'Public Sub create_Analysis()
''�������� ������� ����������� �� ������������
'Set newDOC.MyWdApp = New Word.Application '��������� ����������
''1)�������� ����� ��� ����������� ����������:
'Dim nameFolder As String
'    nameFolder = frmAnalysis.Caption & "_" & frmNewResearch.newEF.getNumber
'Dim nameDOC As String '��� ��������� ������������_�����_���
''2) �������� ����� ���������� � ���������� �� � ��������� �����:
'Dim X As Object '���������� ���������
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

''++++++++++++++++++ � � � � � � +++++++++++++++++++++++++++
''
'Private Sub chkAddExData_Click()
''������� �� ����� "�������� ������"
'    With chkAddExData
'        If .Value = 1 Then
'            Call AE_Enabled
'            .BackColor = &H80FFFF
'            '��������� ����� "���������� ������ ��������"
'            Set newAEF = New clmExpertFindings
'                With newAEF
'                    .name = "�������-�����������"
'                    .definition = newEF.definition
'                    .tanatology = True
'                    .condition = "� ������"
'                    .evidensCategories = "���������� �����"
'                End With
'        Else
'            Call AE_Disabled
'            .BackColor = &HDDEADB
'            '����������� ������ "���������� ������ ��������"
'            Set newAEF = Nothing
'        End If
'    End With
'End Sub
''
'Private Sub AE_Disabled()
''��������� ���������� ����� ����� �������
'    Dim X As Object
'        For Each X In Me.Controls
'            If InStrRev(X.name, "txtAE", 5) > 0 Or X.name = "txtAutopsyDate" Then
'                X.Text = ""
'                X.Enabled = False
'                X.BackColor = &HEBFFEA
'            End If
'        Next X
''����������� ������ "���������� ������ ��������"
'Set newAEF = Nothing
'End Sub
''
'Private Sub AE_Enabled()
''��������� ��������� ����� ����� �������
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
'Private NewCase As clmCase                      '��������� ������ "����"
'Private NewCoroner As HR_Official               '��������� ������ "�����������"
'Private NewInjuredPerson As Entrants            '��������� ������ "�����������"
'Private NewExperts As HR_Official               '��������� ������ "��������"
'Private NewAutopsyExperts As HR_Official        '��������� ������ "����� ��������"
'Private NewExpertFindings As clmExpertFindings  '���������� ������ "���������� ��������"
'Private NewAutopsyEF As clmExpertFindings       '"���������� ������ ��������", ���������� ������ "���������� ��������"
'Private mvarcolEvidences As Collection          '��������� "��" (��� �������� �������� �������)
'Private newDate As clmCaseDate                  '��������� ������ "����"
'Private fst As Date, susp As Date, ren As Date, fin As Date, dueDate As Date '���������� ���: ������, ������������, ������������� � ���������

''������ ���������� ���������� "MsOffice"
'Private MyWdApp As Word.Application         '��������� ���������� MsOffice Word
'Private MyWdDoc As Word.Document                '��������� ��������� MsOffice Word
'Private MyExApp As Excel.Application        '��������� ���������� MsOffice Excel
''Private MyExDoc As Excel.Workbook               '��������� ��������� MsOffice Excel
'Private MyOlApp As Outlook.Application      '��������f� ���������� MsOffice Outlook
''Private MyOlTask As Outlook.TaskItem            '��������� ������ MsOffice Outlook
'Private fs As New FileSystemObject              '���������� ������ ������� (�����)
'Private fll As Variant
''���������� ������� ���������� ActiveX
'Private SvcService As Object, Sel As String
''���������� ������������ �������� ���������� �� ������
''Private DirDOT As String
'Const DirDOT = "D:\Crime\MasterForm\Word\"
''"D:\������\���������������\� ����������\��������� �������\������ FM_Doc\MasterDoc\" '���� ��������
''���������� ���������� "������"
'Dim Cancel As Integer
''
'Public DocNum As String '���������� ����������-�������� ����� ���������
''���������� �������
'Private EvidArray() As String
'Private CellArr(1 To 26) As String
''���������� ��������
'Public n As Byte, i As Byte, bytEvid As Byte, WdEvName As String

''
'Public Property Set colEvidences(ByVal vData As Collection)
''��������� ����������� �������������
''Syntax: Set x.colEvidences = Form1
'    Set mvarcolEvidences = vData
'End Property
''
'Public Property Get colEvidences() As Collection
''��������� ����������� �������������
'Set colEvidences = mvarcolEvidences
''Debug.Print "��������� �� =  " & colEvidences
'End Property
''
'Private Sub cboCsCat_GotFocus()
''��������� ����
'    With cboCsCat
'        .Clear
'        .BackColor = &HC0FFFF '������������� ��������� ���� ��� ��������� �� ������
''�������� ������ � ��������������� ����
'        .AddItem "���������� ����"
'        .AddItem "��������"
'        .AddItem "����������������� ����"
'    End With
'End Sub
''
'Private Sub cboCsCat_LostFocus()
''��������� ����
'    With cboCsCat
'        If Len(.Text) = 0 Then
'            .Text = "(��������� ���� �� �������)"
'            .BackColor = &HC0E0FF
'        Else: newCase.strCsCat = .Text '���������� ��������� ��������
'            If newCase.strCsCat = "��������" Then
'                txtCsIndex.Visible = True
'                txtCsIndex.SetFocus
'            Else: txtCsIndex.Visible = False
'            End If
'        .BackColor = RGB(200, 256, 200) '����������� ������������ ���� ��� ������ �� ������
'        End If
'    End With
'End Sub
''
'Private Sub cboCsDefinition_GotFocus()
''��������� ���������� ����������
'    With cboCsDefinition
'        .Clear
'        .BackColor = &HC0FFFF
'        .AddItem "�������������"
'        .AddItem "�����������"
'    End With
'End Sub
''
'Private Sub cboCsDefinition_LostFocus()
''��������� ���������� ����������
'    With cboCsDefinition
'        If Len(.Text) = 0 Then
'            .Text = "�������������"
'        End If
'        NewExpertFindings.strEFDefinition = .Text '���������� ��������� ��������
'        .BackColor = RGB(200, 256, 200) '����������� ������������ ���� ��� ������ �� ������
'    End With
'End Sub
''
'    Private Sub cboEFCategories_GotFocus()
''��������� ���������� (���������, �������������� � �.�.)
'    With cboEFCategories
'        .Clear
'        .BackColor = &HC0FFFF
'        .AddItem "������-������������������"
'        .AddItem "�����������"
'        .AddItem "��������������"
'        .AddItem "���������"
'        .AddItem "������������"
'    End With
'End Sub
''
'Private Sub cboEFCategories_LostFocus()
''��������� ���������� (���������, �������������� � �.�.)
'    With cboEFCategories
'        If Len(.Text) = 0 Then
'            .Text = "������-������������������"
'        End If
'            NewExpertFindings.strEFCategories = .Text
'            .BackColor = RGB(200, 256, 200)
'    End With
''Debug.Print "��������� ����������= ", NewExpertFindings.strEFCategories
'End Sub
''

'Private Sub cmdCancel1_Click()
''������� ������ "������"
''����������� ����� ����������� ������� � ������������ ������
'Set newCase = Nothing           '��������� ������ "����"
'Set NewInjuredPerson = Nothing  '��������� ������ "�����������"
'Set NewCoroner = Nothing        '��������� ������ "�����������"
'Set NewExperts = Nothing        '��������� ������ "�������"
'Set NewAutopsyExperts = Nothing '��������� ������ "��������"
'Set NewExpertFindings = Nothing '���������� ������ "���������� ��������"
'Set NewAutopsyEF = Nothing      '"���������� ������ ��������", ���������� ������ "���������� ��������"
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
''��������� ������� ������ "C������ � ����� �� ����������"
''���������� ������ � ���������� �����
'Call Cs_Disabled
'Call Cor_Disabled
'Call InjPr_Disabled
'Call AE_ErText
'Call AE_Disabled
''����������� ������������ ����������� ������� � ������������ ������
'Set NewExpertFindings = Nothing '���������� ������ "���������� ��������"
'Set NewAutopsyEF = Nothing      '"���������� ������ ��������", ���������� ������ "���������� ��������"
'Set NewAutopsyExperts = Nothing '��������� ������ "����� �������"
'Set MyWdApp = Nothing
'Set MyExApp = Nothing
''�������� ����� ����������� �������
'Set NewExpertFindings = New clmExpertFindings  '���������� ������ "���������� ��������"
'Set NewAutopsyEF = New clmExpertFindings       '"���������� ������ ��������", ���������� ������ "���������� ��������"
'Set NewAutopsyExperts = New HR_Official        '��������� ������ "����� ��������"
'Set MyWdApp = New Word.Application
'Set MyExApp = New Excel.Application
''��������� ���������
'mdCount.fEFCount = mdCount.fEFCount + 1
'Me.lblEFCount = mdCount.fEFCount '��������� ������ �����
'mdCount.colForms.Add Me, "EFform" & mdCount.fEFCount
'    Dim strTemp As String
'        strTemp = Me.Caption
''��������: ����� ������� - �����
''        Me.Caption = Right(strTemp, 2) & " " & mdCount.fEFCount
'txtEFnum.SetFocus
'End Sub
''
'Private Sub lstAESex_GotFocus()
''��� ������ ��������
'    With lstAESex
'        .Clear
'        .AddItem "���."
'        .AddItem "���."
'        .AddItem "��� �� ������"
'        .BackColor = &HC0FFFF
'    End With
'End Sub
''
'Private Sub lstAESex_LostFocus()
''��� ������ ��������
'    With lstAESex
'        NewAutopsyExperts.HR_sex = .Text
'        .BackColor = RGB(200, 256, 200)
'    End With
''Debug.Print "��� ������ ��������", NewAutopsyExperts.strAESex
'End Sub
''
'Private Sub lstCrSex_GotFocus()
''��� �����������
'    With lstCrSex
'        .Clear
'        .AddItem "���."
'        .AddItem "���."
'        .AddItem "��� �� ������"
'        .BackColor = &HC0FFFF
'    End With
'End Sub
''
'Private Sub lstCrSex_LostFocus()
''��� �����������
'    With lstCrSex
'        NewCoroner.HR_sex = .Text
'        .BackColor = RGB(200, 256, 200)
'    End With
''Debug.Print "��� �����������= ", NewCoroner.strCrSex
'End Sub
''
'Private Sub lstInjPrSex_GotFocus()
' '��� ������������
'    With lstInjPrSex
'        .Clear
'        .AddItem "���."
'        .AddItem "���."
'        .AddItem "��� �� ������"
'        .BackColor = &HC0FFFF
'    End With
'End Sub
''
'Private Sub lstInjPrSex_LostFocus()
''��� ������������
'    With lstInjPrSex
'        NewInjuredPerson.En_Sex = .Text
'        .BackColor = RGB(200, 256, 200)
'    End With
''Debug.Print "��� ������������ = ", NewInjuredPerson.strInjPrSex
'End Sub
''
'Private Sub Erase_Data()
''��������� �������� (�������) ����� ��������� ����� �����
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
''��������� �������� ����� "�������� ����������"
'    Dim frmD As frmDocList
'    Set frmD = New frmDocList
'    frmD.Caption = Me.Caption
'    frmD.lblDocListCount = Me.lblEFCount
'    frmD.Show
'End Sub

'Public Static Sub Form_Initialize()
''�������� ����� �����������
'Set newCase = New clmCase                      '��������� ������ "����"
'Set NewInjuredPerson = New Entrants            '��������� ������ "�����������"
'Set NewCoroner = New HR_Official               '��������� ������ "�����������"
'Set NewExperts = New HR_Official               '��������� ������ "��������"
'Set NewAutopsyExperts = New HR_Official        '��������� ������ "����� ��������"
'Set NewExpertFindings = New clmExpertFindings  '���������� ������ "���������� ��������"
'Set NewAutopsyEF = New clmExpertFindings       '"���������� ������ ��������" - ���������� ������ "���������� ��������"
'Set newDate = New clmCaseDate                  '��������� ������ "����"
''������������� ��������� ������� ������
'Set colEvidences = New Collection
''������������� ����������:
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
''��������� �������� �����
'    If MsgBox("�������?", vbYesNo, "�����?") = vbYes Then
'        Set frmEvCategories = Nothing
'    '�������� �������� �������
'        Set newCase = Nothing           '��������� ������ "����"
'        Set NewInjuredPerson = Nothing  '��������� ������ "�����������"
'        Set NewCoroner = Nothing        '��������� ������ "�����������"
'        Set NewExperts = Nothing        '��������� ������ "�������"
'        Set NewAutopsyExperts = Nothing '��������� ������ "����� �������"
'        Set NewExpertFindings = Nothing '���������� ������ "���������� ��������"
'        Set NewAutopsyEF = Nothing      '"���������� ������ ��������", ���������� ������ "���������� ��������"
'        Set newDate = Nothing
'        Set MyWdApp = Nothing
''        Set MyExApp = Nothing
''        Set MyOlApp = Nothing
'Dim tmp As String
'        tmp = Me.lblEFCount
''�������� ������� ����� frmEvidences
'    Unload mdCount.colForms("Evidform" & tmp)
''������� ��������� ���� colForms:
'    mdCount.colForms.Remove ("Evidform" & tmp)
'    mdCount.colForms.Remove ("DocListform" & tmp)
'    mdCount.colForms.Remove ("EFform" & tmp)
'    Else
'        Cancel = 1
'    End If
'End Sub
''
'Private Static Sub txtAEMidName_GotFocus()
''�������� ������ ��������
'    With txtAEMidName
'        .Text = ""
'        .BackColor = &HC0FFFF
'    End With
'End Sub
''
'Private Static Sub txtAEMidName_LostFocus()
''�������� ������ ��������
'    With txtAEMidName
'        If Len(.Text) = 0 Then
'            .Text = "(�� �������)"
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
''Debug.Print "�������� ������ ��������= ", NewAutopsyExperts.HR_MidName
'End Sub
''
'Private Static Sub txtAEName_GotFocus()
''��� ������ ��������
'    With txtAEName
'        .Text = ""
'        .BackColor = &HC0FFFF
'    End With
'End Sub
''
'Private Static Sub txtAEName_LostFocus()
''��� ������ ��������
'    With txtAEName
'        If Len(.Text) = 0 Then
'            .Text = "(��� �� �������)"
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
''Debug.Print "��� ������ ��������= ", NewAutopsyExperts.HR_name
'End Sub
''
'Private Static Sub txtAEFNum_GotFocus()
''���������� ����� ���������� ������ ��������
'    With txtAEFNum
'        .Text = ""
'        .BackColor = &HC0FFFF
'    End With
'End Sub
''
'Private Static Sub txtAEFNum_LostFocus()
''���������� ����� ���������� ������ ��������
'    With txtAEFNum
'        If Len(.Text) = 0 Then
'            .Text = "(�� ������)"
'            .BackColor = &HC0E0FF
'        Else: NewAutopsyEF.strEFNum = .Text
'            .BackColor = RGB(200, 256, 200)
'        End If
'    End With
''Debug.Print "����� ���������� ������ ��������= ", NewAutopsyEF.strEFNum
'End Sub
''
'Private Static Sub txtAESurName_GotFocus()
''������� ������ ��������
'    With txtAESurName
'        .Text = ""
'        .BackColor = &HC0FFFF
'    End With
'End Sub
''
'Private Static Sub txtAESurName_LostFocus()
''������� ������ ��������
'    With txtAESurName
'        If Len(.Text) = 0 Then
'            .Text = "(�� �������)"
'            .BackColor = &HC0E0FF
'        Else: NewAutopsyExperts.HR_SurName = StrConv(.Text, vbProperCase) '��������� � ��������� �����
'            .Text = NewAutopsyExperts.HR_SurName
'            .BackColor = RGB(200, 256, 200)
'        End If
'    End With
''Debug.Print "������� ������ ��������= ", NewAutopsyExperts.HR_SurName
'End Sub
''
'Private Static Sub txtAutopsyDate_GotFocus()
''���� ������ ����� ���
'    With txtAutopsyDate
'        .Text = ""
'        .BackColor = &HC0FFFF
'        .Text = CStr(NewExpertFindings.DtmEFFirstDay - 2)
'    End With
'End Sub
''
'Private Static Sub txtAutopsyDate_LostFocus()
''���� ������ ����� ���
'    With txtAutopsyDate
'        If Len(.Text) = 0 Then
'            .Text = "(�� �������)"
'            .BackColor = &HC0E0FF
'        Else: NewInjuredPerson.En_Autopsy = CDate(.Text)
'            .BackColor = RGB(200, 256, 200)
'        End If
'    End With
''Debug.Print "���� ������ ����� ���= ", NewInjuredPerson.En_Autopsy
'End Sub
''
'Private Static Sub txtCorOprAr_GotFocus()
''����� ������������
'    With txtCorOprAr
'        .Text = ""
'        .BackColor = &HC0FFFF
'        .Text = "(�. ������) ��������� ������ ������������� �������� ���������� ��������"
'    End With
'End Sub
''
'Private Static Sub txtCorOprAr_LostFocus()
''����� ������������
'    With txtCorOprAr
'        If Len(.Text) = 0 Then
'            .Text = "(����� �� ������)"
'            .BackColor = &HC0E0FF
'        Else: NewCoroner.HR_OprAr = .Text
'            .BackColor = RGB(200, 256, 200)
'        End If
'    End With
''Debug.Print "����� = ", NewCoroner.HR_OprAr
'End Sub
''
'Private Static Sub txtCorPost_GotFocus()
''��������� �����������
'    With txtCorPost
'    .BackColor = &HC0FFFF
'    .Text = "�������� �����������"
'    End With
'End Sub
''
'Private Static Sub txtCorPost_LostFocus()
''��������� �����������
'    With txtCorPost
'        If Len(.Text) = 0 Then
'            .Text = "(�������� �� �������)"
'            .BackColor = &HC0E0FF
'        Else: NewCoroner.HR_Post = .Text
'            .BackColor = RGB(200, 256, 200)
'        End If
'    End With
''Debug.Print "��������� �����������= ", NewCoroner.HR_Post
'End Sub
''
'Private Sub txtCsIndex_GotFocus()
''������ ����
'    With txtCsIndex
'        .Text = ""
'        .BackColor = &HC0FFFF
'    End With
'End Sub
''
'Private Sub txtCsIndex_LostFocus()
''������ ����
'    With txtCsIndex
'        If Len(.Text) = 0 Then
'            .Text = "(�� ������)"
'            .BackColor = &HC0E0FF
'        Else: newCase.strCsIndex = .Text
'            .BackColor = RGB(200, 256, 200)
'        End If
'    End With
''Debug.Print "������ ����= ", NewCase.strCsIndex
'End Sub
''
'Private Static Sub txtEFnum_GotFocus()
''����� ���� ����������
'    With txtEFnum
'        .Text = ""
'        .BackColor = &HC0FFFF
'        .ForeColor = &H80000008
'    End With
'End Sub
''
'Private Static Sub txtEFnum_LostFocus()
''����� ���� ����������
'    With txtEFnum
'        Do
'            If Not IsNumeric(.Text) Or Len(.Text) = 0 Then
'                Beep
'                .BackColor = RGB(256, 0, 0)
'                MsgBox "C������ ������� �����", vbCritical, "������ �����"
'                .Text = InputBox("������� ��������� ����� ����������!", "����������� ������ �����")
'            Else: NewExpertFindings.strEFNum = .Text
'                .BackColor = RGB(200, 256, 200)
'        Exit Do
'            End If
'        Loop
'    End With
'''���������� ��������� � ������� ���������:
''NewExpertFindings.DOC = mdPrintDoc.strDOC '������
'NewExpertFindings.strEFEvCategories = mdPrintDoc.DocCat '���������
'End Sub
''
'Private Static Sub txtCsNum_GotFocus()
''����� ����
'    With txtCsNum
'        .BackColor = &HC0FFFF
'        .Text = ""
'        .MaxLength = 11
'    End With
'End Sub
''
'Private Static Sub txtCsNum_LostFocus()
''����� ���� CsNum
'    With txtCsNum
'        If Len(.Text) = 0 Then
'            .MaxLength = 22
'            .Text = "(����� ���� �� ������)"
'            .BackColor = &HC0E0FF
'        Else
'            Do
'                If Not IsNumeric(.Text) Then
'                    Beep
'                    MsgBox "C������ ������� �����", vbCritical, "������ �����"
'                    .Text = InputBox("������� ��������� ����� ����!", "����������� ������ �����")
'                Else: newCase.strCsNum = .Text
'                    .BackColor = RGB(200, 256, 200)
'            Exit Do
'                End If
'            Loop
'        End If
'    End With
'' Debug.Print "����� ����= ", NewCase.strCsNum
'End Sub
''
'Private Static Sub txtCorMidName_GotFocus()
''�������� �����������
'    With txtCorMidName
'        .Text = ""
'        .BackColor = &HC0FFFF '������������� ��������� ���� ��� ��������� �� ������
'    End With
'End Sub
''
'Private Static Sub txtCorMidName_LostFocus()
''�������� �����������
'    With txtCorMidName
'        If Len(.Text) = 0 Then
'            .Text = "(�������� �� �������)"
'            .BackColor = &HC0E0FF
'        Else
'            If Len(txtCorMidName.Text) = 1 Then
'                NewCoroner.HR_MidName = StrConv(txtCorMidName.Text, vbProperCase) & "."
'            Else: NewCoroner.HR_MidName = StrConv(.Text, vbProperCase)
'            End If
'            .BackColor = RGB(200, 256, 200) '����������� ������������ ���� ��� ������ �� ������
'            .Text = NewCoroner.HR_MidName
'        End If
'    End With
''Debug.Print "�������� �����������= ", NewCoroner.HR_MidName
'End Sub
''
'Private Static Sub txtCorName_GotFocus()
''��� �����������
'    With txtCorName
'        .Text = ""
'        .BackColor = &HC0FFFF '������������� ��������� ���� ��� ��������� �� ������
'    End With
'End Sub
''
'Private Static Sub txtCorName_LostFocus()
''��� �����������
'    With txtCorName
'        If Len(.Text) = 0 Then
'            .Text = "(��� �� �������)"
'            .BackColor = &HC0E0FF
'        Else
''���������� ����� ����� ��������
'            If Len(txtCorName.Text) = 1 Then
'                NewCoroner.HR_name = StrConv(.Text, vbProperCase) & "."
'            Else: NewCoroner.HR_name = StrConv(txtCorName.Text, vbProperCase)
'            End If
'        .Text = NewCoroner.HR_name
'        .BackColor = RGB(200, 256, 200) '����������� ������������ ���� ��� ������ �� ������
'        End If
'    End With
''Debug.Print "��� �����������= ", NewCoroner.HR_name
'End Sub
''
'Private Static Sub txtCorRank_GotFocus()
''������ �����������
'    With txtCorRank
'        .BackColor = &HC0FFFF '������������� ��������� ���� ��� ��������� �� ������
'        .Text = "�������� ���������� �������"
'    End With
'End Sub
''
'Private Static Sub txtCorRank_LostFocus()
''������ �����������
'    With txtCorRank
'        If Len(.Text) = 0 Then
'            .Text = "(������ �� �������)"
'            .BackColor = &HC0E0FF
'        Else: NewCoroner.HR_Rank = .Text
'            .BackColor = RGB(200, 256, 200)
'        End If
'    End With
''Debug.Print "������ �����������= ", NewCoroner.HR_Rank
'End Sub
''
'Private Static Sub txtEFFirstDay_GotFocus()
''���� ������
'    With txtEFFirstDay
'        .Text = ""
'        .BackColor = &HC0FFFF
'        .ForeColor = &H80000008
'        .Text = newDate.DateNow '����������� ������� ���� � ��������� ����
'    End With
'End Sub
''
'Private Static Sub txtEFFirstDay_LostFocus()
''���� ������
'With txtEFFirstDay
'    fst = newDate.ExamDate(.Text)
'    NewExpertFindings.DtmEFFirstDay = fst
'    .Text = newDate.dateToString(fst)
'    .BackColor = RGB(200, 256, 200)
'
'
''    If Len(.Text) = 0 Then
''        .Text = InputBox("������� ���� ������ ����������", "���� ����", DateTime.Date)
''            If Len(.Text) = 0 Then
''                .Text = "�� �������"
''                .BackColor = RGB(256, 0, 0)
''                .ForeColor = &HFFFF&
''            Else
''                Dim tmpDt As Date
''                    tmpDt = newDate.ExamDate(.Text)
''                    .Text = newDate.dateToString(tmpDt)
''                NewExpertFindings.DtmEFFirstDay = tmpDt
''                .ForeColor = &H80000008
''                On Error Resume Next ' ��������� ������
''            End If
''    ElseIf .Text = "�� �������" Then
''            .BackColor = RGB(256, 0, 0)
''            .ForeColor = &HFFFF&
''    Else
''        tmpDt = newDate.ExamDate(.Text)
''        NewExpertFindings.DtmEFFirstDay = tmpDt
''        .ForeColor = &H80000008
''        .Text = newDate.dateToString(tmpDt)
''    '����� ��������� ������� ����� ��������� ����������
'    NewExpertFindings.DtmEFDueDate = newDate.getPeriod(fst)
''    End If
'End With
'Debug.Print "���� =", NewExpertFindings.DtmEFDueDate
'End Sub
''
'Private Static Sub txtCorSurName_GotFocus()
''������� �����������
'    With txtCorSurName
'        .Text = ""
'        .BackColor = &HC0FFFF
'    End With
'End Sub
''
'Private Static Sub txtCorSurName_LostFocus()
''������� �����������
'    With txtCorSurName
'        If Len(.Text) = 0 Then
'            .Text = "(������� �� �������)"
'            .BackColor = &HC0E0FF
'        Else: NewCoroner.HR_SurName = StrConv(.Text, vbProperCase)
'            .BackColor = RGB(200, 256, 200)
'            .Text = NewCoroner.HR_SurName
'        End If
'    End With
''Debug.Print "������� �����������= ", NewCoroner.HR_SurName
'End Sub
''
'Private Static Sub txtInjPrDeceaseDate_GotFocus()
''���� ������ ������������
'    With txtInjPrDeceaseDate
'        .Text = ""
'        .BackColor = &HC0FFFF
'        .Text = CStr(NewExpertFindings.DtmEFFirstDay - 3)
'    End With
'End Sub
''
'Private Static Sub txtInjPrDeceaseDate_LostFocus()
''���� ������ ������������
'    With txtInjPrDeceaseDate
'        If Len(.Text) = 0 Then
'            .Text = "(�� �������)"
'            .BackColor = &HC0E0FF
'        Else: NewInjuredPerson.En_Decease = .Text
'            .BackColor = RGB(200, 256, 200)
'        End If
'    End With
''Debug.Print "���� ������ ������������= ", NewInjuredPerson.En_Decease
'End Sub
''
'Private Static Sub txtCsPrvDate_GotFocus()
''���� ��������� �������������
'    With txtCsPrvDate
'        .Text = ""
'        .BackColor = &HC0FFFF
'        .ForeColor = &H80000008
'        .Text = CStr(NewExpertFindings.DtmEFFirstDay - 3)
'    End With
'End Sub
''
'Private Static Sub txtCsPrvDate_LostFocus()
''���� ��������� �������������
'    With txtCsPrvDate
'        If Len(.Text) = 0 Then
'            .Text = "(�� �������)"
'            .BackColor = RGB(256, 0, 0)
'            .ForeColor = &HFFFF&
'       ElseIf .Text = "�� �������" Then
'            .BackColor = RGB(256, 0, 0)
'            .ForeColor = &HFFFF&
'        Else: newCase.DtmCsPrvDate = CDate(.Text)
'            .BackColor = RGB(200, 256, 200)
'        End If
'    End With
''Debug.Print "���� ��������� �������������= ", NewCase.DtmCsPrvDate
'End Sub
''
'Private Static Sub txtInjPrBirthday_GotFocus()
''��� �������� ������������
'    With txtInjPrBirthday
'        .Text = ""
'        .BackColor = &HC0FFFF
'    End With
'End Sub
''
'Private Static Sub txtInjPrBirthday_LostFocus()
''��� �������� ������������
'    With txtInjPrBirthday
'        If Len(.Text) = 0 Then
'            .Text = "(�� ������)"
'            .BackColor = &HC0E0FF
'        Else: NewInjuredPerson.En_Birthday = .Text
'            .BackColor = RGB(200, 256, 200)
'        End If
''!!! �������� ������� ���������� �������� ��� ����� ������
'    End With
''Debug.Print "��� �������� ������������", NewInjuredPerson.En_Birthday
'End Sub
''
'Private Static Sub txtInjPrMidName_GotFocus()
''�������� ������������
'    With txtInjPrMidName
'        .Text = ""
'        .BackColor = &HC0FFFF
'    End With
'End Sub
''
'Private Static Sub txtInjPrMidName_LostFocus()
''�������� ������������
'    With txtInjPrMidName
'        If Len(.Text) = 0 Then
'            .Text = "(�������� �� �������)"
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
''Debug.Print "�������� ������������= ", NewInjuredPerson.En_MidName
'End Sub
''
'Private Static Sub txtInjPrName_GotFocus()
''��� ������������
'    With txtInjPrName
'        .Text = ""
'        .BackColor = &HC0FFFF
'    End With
'End Sub
''
'Private Static Sub txtInjPrName_LostFocus()
''��� ������������
'    With txtInjPrName
'        If Len(.Text) = 0 Then
'            .Text = "(��� �� �������)"
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
''Debug.Print "��� ������������= ", NewInjuredPerson.En_Name
'End Sub
''
'Private Static Sub txtInjPrSurName_GotFocus()
''������� ������������
'    With txtInjPrSurName
'        .Text = ""
'        .BackColor = &HC0FFFF
'    End With
'End Sub
''
'Private Static Sub txtInjPrSurName_LostFocus()
''������� ������������
'    With txtInjPrSurName
'        If Len(.Text) = 0 Then
'            .Text = "(������� �� �������)"
'            .BackColor = &HC0E0FF
'        Else: NewInjuredPerson.En_SurName = StrConv(.Text, vbProperCase)
'            .BackColor = RGB(200, 256, 200)
'            .Text = NewInjuredPerson.En_SurName
'        End If
'    End With
''Debug.Print "������� ������������= ", NewInjuredPerson.En_SurName
'End Sub
''
'Public Static Sub Criate_NewFolders()
''�������� ������ �������� ����� � �������� ����������
'ChDrive "D" '����� ������������ �����
''NewExpertFindings.strEFEvCategories = frmMDI.strCrEFCat
'NewExpertFindings.DOC = "D:\Crime\" & Year(Now) & "\" & NewExpertFindings.strEFNum & "_" & Right(Year(Now), 2) & "_" & _
'            NewExpertFindings.strEFEvCategories & "\"
'If Not fs.FolderExists(NewExpertFindings.DOC) Then  ' ���� ����� ����� �� ���������� � ���� �����
'        Set fll = fs.CreateFolder(NewExpertFindings.DOC) '�� ������� ��
''������� ��������� ����� ��� ����
'        fs.CreateFolder (fll & "\" & "Img_" & NewExpertFindings.strEFNum & "_" & Right(Year(Now), 2))
'        fs.CreateFolder (fll & "\" & "Img_" & NewExpertFindings.strEFNum & "_" & Right(Year(Now), 2) & _
'                        "\" & "ImgEOD_" & NewExpertFindings.strEFNum & "_" & Right(Year(Now), 2))
'    Else
'        fll = NewExpertFindings.DOC ' � ��������� ������ ����������  fll ����� ��������� ������ �� ��� (��� ������������) �����
'    End If
''Debug.Print "DirDOC=", NewExpertFindings.DOC
'End Sub
''
'Public Static Sub Create_NewDoc()
''��������� �������� ����� ����������
''DirDOT = "D:\������\���������������\� ����������\��������� �������\������ FM_Doc\MasterDoc\" '��� ����������
'''1)��������� �������� � ���������� ������
''    If frmDocList.chkReportEF.Value = 1 Then
''        Create_ReportEF
''    End If
''''��������� �������� ������ Outlook
'''    If frmDocList.chkTaskItem.Value = 1 Then
'''        Create_TaskItem
'''    End If
''______________________________________________________________
''������ � ����������� Word
'    With MyWdApp
'        .Visible = False
''mdCount.colForms.Add frmD, "EFform" & Me.EvCat_ID
'        Set MyWdDoc = .Documents.Add(DirDOT & mdPrintDoc.strDOC)
'            Create_EF
''��������� ����������-�������
'Dim tmp As String
'tmp = Me.lblEFCount
'''��������� �������� �����������
''    If mdCount.colForms("DocListform" & tmp).chkFotoListEF.Value = 1 Then
''    Set MyWdDoc = .Documents.Add(DirDOT & "FotoList.dotm")
''        Create_FotoList
''    End If
'''��������� �������� ����������� �� ��������
''    If mdCount.colForms("DocListform" & tmp).chkCNBiology.Value = 1 Then
''        Set MyWdDoc = .Documents.Add(DirDOT & "Biology.dotm")  '��������� ������"
''        Create_Biology
''    End If
'''��������� �������� ����������� �����
''    If mdCount.colForms("DocListform" & tmp).chkCNGenom.Value = 1 Then
''        Set MyWdDoc = .Documents.Add(DirDOT & "Genome.dotm")  '��������� ������"
''        Create_Genom
''    End If
'''��������� �������� ����������� �� �����
''    If mdCount.colForms("DocListform" & tmp).chkCrimMet.Value = 1 Then
''       Set MyWdDoc = .Documents.Add(DirDOT & "FocusOnMet.dotm")  '��������� ������"
''       Create_FocusOnMet
''    End If
'''��������� �������� ��������
''    If mdCount.colForms("DocListform" & tmp).chkCoverNote.Value = 1 Then
''       Set MyWdDoc = .Documents.Add(DirDOT & "Labels.dotm")  '��������� ������"
''        Create_CovNoteLabel
''    End If
'''��������� �������� �������
''    If mdCount.colForms("DocListform" & tmp).chkAEFInquiry.Value = 1 Then
''       Set MyWdDoc = .Documents.Add(DirDOT & ".dotm")  '��������� ������"
'''       Create_AEFInquiry
''    End If
'    .Quit '��������� ����������
'    End With
''Set MyWdDoc = Nothing
'End Sub
''
'Public Static Sub Create_EF()
''��������� �������� ������ ��������� Ecspert Findings
'    MyWdDoc.SaveAs2 FileName:=NewExpertFindings.DOC & "����������_" & NewExpertFindings.strEFNum _
'            & "_" & Right(Year(Now), 2), FileFormat:=wdFormatXMLDocument, _
'            LockComments:=False, Password:="", AddToRecentFiles:=True, WritePassword:="", _
'            ReadOnlyRecommended:=False, EmbedTrueTypeFonts:=False, SaveNativePictureFormat:=False, _
'            SaveFormsData:=False, SaveAsAOCELetter:=False
'    With MyWdApp.ActiveDocument  '�������� �������� ��������
'        On Error Resume Next
''����:
'       .FormFields("WdDirDOC").Result = NewExpertFindings.DOC
'       .FormFields("WdEFNum").Result = NewExpertFindings.strEFNum
'       .FormFields("WdFirstDay").Result = NewExpertFindings.DtmEFFirstDay
''��������:
'        .Bookmarks("WdEFfullNum").Range.Text = NewExpertFindings.Criate_FullEFNum '��� ����������
'        .Bookmarks("WdExperience").Range.Text = NewExperts.Calculate_Experience '������ ����� ������
'        .Bookmarks("WdRsnComplain").Range.Text = Print_RsnComplain '��.���������
'        .Bookmarks("FactCase").Range.Text = Print_FactCase '�������������� ����
'    If chkAddExData.Value = 1 Then
'        .Bookmarks("WdAEDirection").Range.Text = Print_AEDirection '����������� ������ ��������
'    End If
'    Dim tmp As String
'        tmp = Me.lblEFCount
'    If mdCount.colForms("DocListform" & tmp).chkAEFInquiry.Value = 1 Then
'        NewExpertFindings.DtmEFSuspDate = DateTime.Date
'        .Bookmarks("WdEFSuspDate").Range.Text = NewExpertFindings.DtmEFSuspDate
'        .Bookmarks("WdAEFInquiry").Range.Text = Print_AEFInquiry '������������ �����������
'    End If
'
''������ �������� �������� � �.�. ������������ �������������
''        .Bookmarks("WdEvNumArr1").Range.Text = NewEvid.Print_EvNumArr1
''        .Bookmarks("WdEvNumArr2").Range.Text = NewEvid.Print_EvNumArr2
''        .Bookmarks("WdEvNumArr3").Range.Text = NewEvid.Print_EvNumArr3
'''������ ��:
'    .Bookmarks("WdEvColumnArr").Range.Select
'            Call Print_NumColumn
''         For i = 1 To n - 1
''            .Bookmarks("WdEvColumnArr").Range.Text = EvidArray(i) & Chr(10)
''            .Bookmarks("WdEvLineArr").Range.Text = EvidArray(i) & ", "
''            .Bookmarks("WdEvLineArr1").Range.Text = EvidArray(i) & ", "
''            .Bookmarks("WdEvLineArr3").Range.Text = EvidArray(i) & ", "
''        Next i
''    On Error GoTo 999 ' �������� ��������� ������
'    .Close SaveChanges:=wdSaveChanges
'    End With
'Set MyWdDoc = Nothing
''999:
''    MsgBox Err.Description  '������
''    Err.Clear
'End Sub
''

''
'Private Static Sub Create_CovNoteLabel()
''�������� ��������
''    MyWdDoc.SaveAs2 FileName:=NewExpertFindings.DOC & "��������_" & NewExpertFindings.strEFNum _
''            & "_" & Right(Year(Now), 2), FileFormat:=wdFormatXMLDocument, _
''            LockComments:=False, Password:="", AddToRecentFiles:=True, WritePassword:="", _
''            ReadOnlyRecommended:=False, EmbedTrueTypeFonts:=False, SaveNativePictureFormat:=False, _
''            SaveFormsData:=False, SaveAsAOCELetter:=False
''    With MyWdApp.ActiveDocument  '�������� �������� ��������
''        On Error Resume Next ' ��������� ������
'''������ � ������ � ���������� ���������:
'''����:
'''������ � ����� ���������:
''    .Bookmarks("WdCovNoteLable").Range.Text = Print_CovNoteLabel
'''���������� ������ � ����������
''        .Save
''        .Close
''     End With
'End Sub
''
'Private Static Function Print_CovNoteLabel() As String
''  '������� ������ ��������
'''    ������� �������:
''Dim strCN As String
''    If NewEvid.bytEvUnit = 1 Then
''        Print_CovNoteLabel = "���������� ������-������������������ ���������." & Chr(10) _
''                & "���������� " & NewExpertFindings.Criate_FullEFNum & NewCase.Print_CsData & "." & Chr(10) _
''                & Chr(10) & "������: " & EvidArray(1) & "." & Chr(10) & Chr(10) & "���� ����������� ��������: " & _
''                NewExpertFindings.DtmEFFirstDay & "." & Chr(10) & "���� ��������: " & DateTime.Date & Chr(10) _
''                & Chr(10) & "����" & vbTab & Chr(32) & vbTab & Chr(32) & vbTab & Chr(32) & vbTab & Chr(32) _
''                & vbTab & Chr(32) & vbTab & Chr(32) & vbTab & Chr(32) & vbTab & Chr(32) & vbTab _
''                & Chr(32) & "�.�. �������."
''    Else:
''            For i = 1 To n - 1
''                strCN = strCN & "-" & EvidArray(i) & ";" & Chr(10)
''                Print_CovNoteLabel = "���������� ������-������������������ ���������." & Chr(10) _
''                & "���������� " & NewExpertFindings.Criate_FullEFNum & NewCase.Print_CsData & "." & Chr(10) _
''                & Chr(10) & "�������:" & Chr(10) & strCN & Chr(10) & "���� ����������� ��������: " & _
''                NewExpertFindings.DtmEFFirstDay & "." & Chr(10) & "���� ��������: " & DateTime.Date & "." & Chr(10) _
''                & Chr(10) & "����" & vbTab & Chr(32) & vbTab & Chr(32) & vbTab & Chr(32) & vbTab & Chr(32) _
''                & vbTab & Chr(32) & vbTab & Chr(32) & vbTab & Chr(32) & vbTab & Chr(32) & vbTab _
''                & Chr(32) & "�.�. �������."
''            Next i
''   End If
'End Function
''
'Private Static Sub Print_CovNoteLabel_1()
''''������� ������ ��������
''''    ������� �������:
'''N = i + 1
'''    For i = 1 To N - 1
'''        Selection.TypeParagraph
'''        Selection.TypeText Text:="���������� ������-������������������ ���������" & Chr(10) _
'''                & "���������� " & NewExpertFindings.Criate_FullEFNum & NewCase.Print_CsData & Chr(10) _
'''                & "������:"
'''        Selection.TypeParagraph
'''        Selection.TypeText Text:=EvidArray(i) & Chr(10)
'''        Selection.TypeParagraph
'''        Selection.TypeText Text:="���� ����������� ��������: " & _
'''                NewExpertFindings.DtmEFFirstDay & Chr(10) & "���� ��������: "
''''          ������� ������ �� ���� "������� ����"
'''        Selection.InsertCrossReference ReferenceType:="��������", ReferenceKind:= _
'''            wdContentText, ReferenceItem:="ActualDate", InsertAsHyperlink:=True, _
'''            IncludePosition:=False, SeparateNumbers:=False, SeparatorString:=" "
'''        Selection.MoveLeft Unit:=wdCharacter, Count:=1, Extend:=wdExtend
'''        Selection.Font.Name = "Times New Roman"
'''        Selection.Font.Size = 14
'''        Selection.MoveRight Unit:=wdCharacter, Count:=1
'''        Selection.TypeParagraph
'''        Selection.TypeText Text:="�������" & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab _
'''            & vbTab & "�.�. �������" & Chr(10)
''''            & "��������" & vbTab & vbTab & vbTab & vbTab _
''''            & vbTab & vbTab & vbTab & Chr(10) & "�.�. ��������"
'''        Selection.TypeParagraph
'''    Next i
'''End Sub
'''Private Static Sub AEFdirection()
'''    If chkAddExData.Value = 0 Then
'''        Selection.TypeText Text:="�� ����������� " & WithDoc.Print_AEFData & " �������: ... " & WithDoc.Print_AEFData & _
'''            " ������� �������������� � �������������� ������� ������: ..."
'''    End If
'End Sub
''
'Private Static Sub Print_PetitionCor()
''������� ������ ������ "������ �����������"
''    If frmDocList.chkCoverNote.Value = 1 Then
''        Selection.TypeText Text:=WithDoc.dtFirstDay & "�� ��� " & WithDoc.Print_Coroner _
''           & " ������� ����������� � �������������� ���������� (����� ����������)" & _
''        " �� ����������� ��������� � ������� ������� �������� ����������� � " & WithDoc.Print_InjPr & " (" & WithDoc.strAEFNum & ") "
''    End If
'End Sub
''
'Private Static Sub Create_Biology()
''��������� �������� ������ ��������� "����������� �� ��������"
'    MyWdDoc.SaveAs2 FileName:=NewExpertFindings.DOC & "��������_" & NewExpertFindings.strEFNum _
'            & "_" & Right(Year(Now), 2), FileFormat:=wdFormatXMLDocument, _
'            LockComments:=False, Password:="", AddToRecentFiles:=True, WritePassword:="", _
'            ReadOnlyRecommended:=False, EmbedTrueTypeFonts:=False, SaveNativePictureFormat:=False, _
'            SaveFormsData:=False, SaveAsAOCELetter:=False
'    With MyWdApp.ActiveDocument  '�������� �������� ��������
'        On Error Resume Next ' ��������� ������
''��������:
''        .Bookmarks("WdRsnComplain").Range.Text = NewCoroner.Print_Coroner
''�������:
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
''��������� �������� ������ ��������� "����������� �� �����"
'    MyWdDoc.SaveAs2 FileName:=NewExpertFindings.DOC & "�����_" & NewExpertFindings.strEFNum _
'            & "_" & Right(Year(Now), 2), FileFormat:=wdFormatXMLDocument, _
'            LockComments:=False, Password:="", AddToRecentFiles:=True, WritePassword:="", _
'            ReadOnlyRecommended:=False, EmbedTrueTypeFonts:=False, SaveNativePictureFormat:=False, _
'            SaveFormsData:=False, SaveAsAOCELetter:=False
'    With MyWdApp.ActiveDocument  '�������� �������� ��������
'        On Error Resume Next ' ��������� ������
'        '��������:
''        .Bookmarks("WdRsnComplain").Range.Text = NewCoroner.Print_Coroner
''        '�������:
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
''��������� �������� ������ ��������� "����������� �� �����, ������������"
'    MyWdDoc.SaveAs2 FileName:=NewExpertFindings.DOC & "����������� �� Me_" & NewExpertFindings.strEFNum _
'            & "_" & Right(Year(Now), 2), FileFormat:=wdFormatXMLDocument, _
'            LockComments:=False, Password:="", AddToRecentFiles:=True, WritePassword:="", _
'            ReadOnlyRecommended:=False, EmbedTrueTypeFonts:=False, SaveNativePictureFormat:=False, _
'            SaveFormsData:=False, SaveAsAOCELetter:=False
'    With MyWdApp.ActiveDocument  '�������� �������� ��������
'        On Error Resume Next ' ��������� ������
'        '����:
''.FormFields("WdInjPr").Result = NewInjuredPerson.Create_InitialsInjPrFullName
''.FormFields("WdInjPrBirthday").Result = NewInjuredPerson.strInjPrBirthday & "�.�."
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
'''��������:
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
''������ ������� ����������� ��������
'Print_AEDirection = "�� ����������� " & NewAutopsyExperts.Print_Expert & " �������: " & _
'        Chr(34) & NewInjuredPerson.Print_Autopsy & " ���������� (���) �" & NewAutopsyEF.strEFNum _
'        & " ������� �������������� � �������������� ������� ������ �������-����������� ������� �" _
'        & Chr(34) & "." & Chr(10)
'End Function
''
'Private Static Function Print_AEFInquiry() As String
''  ������� ������ ������� "���������� �����������"
'Dim tmp As String
'        tmp = Me.lblEFCount
'    If mdCount.colForms("DocListform" & tmp).chkAEFInquiry.Value = 1 Then
'        Print_AEFInquiry = DateTime.Date & " �� ��� " & NewCoroner.Print_Coroner _
'            & " ������� ����������� � �������������� ���������� (����� ����������) �������� �� ����������� ��������� � ������� ������� �������� ����������� � " _
'            & NewInjuredPerson.Print_InjPersons & Chr(10) & _
'            vbTab & "������������� ��������� �������������/�� �������������." & Chr(10)
'    Else: Print_AEFInquiry = "����������� �� ����������." & Chr(10)
'    End If
'End Function
''
'Private Static Function Print_FactCase() As String
''������� ������ "�������������� ����"
'  Print_FactCase = newCase.DtmCsPrvDate & Chr(32) & NewInjuredPerson.Print_InjPersons & Chr(10)
'End Function
''
'Public Static Function Print_LegalGround() As String
''������� ������ ������ "����������� ��������� ���������� ����������"
''        Print_LegalGround = "�� ��������� " & NewCase.strCsDefinition & Chr(32) & NewCoroner.Print_Coroner _
''        & Chr(32) & NewCase.Print_CsData
'End Function
''
'Public Static Function Print_RsnComplain() As String
''������� ������ ������ "��������� ���������� ����������"
'    If txtAESurName.Text = "" Then
'        Print_RsnComplain = " �� ��������� " & NewExpertFindings.strEFDefinition & Chr(32) _
'        & "� ���������� " & NewExpertFindings.strEFCategories & " ���������� " & NewCoroner.Print_Coroner _
'        & ", ����������� " & newCase.DtmCsPrvDate & Chr(32) & newCase.Print_CsData & ", ����� ������-������������������ ����������."
'    Else: Print_RsnComplain = " �� ��������� " & NewExpertFindings.strEFDefinition & Chr(32) _
'    & "� ���������� " & NewExpertFindings.strEFCategories & " ���������� " & NewCoroner.Print_Coroner _
'        & ", ����������� " & newCase.DtmCsPrvDate & Chr(32) & newCase.Print_CsData & " � ����������� " & NewAutopsyExperts.Print_Expert & " �� " & _
'        NewInjuredPerson.En_Autopsy & ", ����� ������-������������������ ����������."
'    End If
'End Function
''
'Private Static Sub Create_FotoList()
''�������� �����������
'    MyWdDoc.SaveAs2 FileName:=NewExpertFindings.DOC & "�����������_" & NewExpertFindings.strEFNum _
'            & "_" & Right(Year(Now), 2), FileFormat:=wdFormatXMLDocument, _
'            LockComments:=False, Password:="", AddToRecentFiles:=True, WritePassword:="", _
'            ReadOnlyRecommended:=False, EmbedTrueTypeFonts:=False, SaveNativePictureFormat:=False, _
'            SaveFormsData:=False, SaveAsAOCELetter:=False
'    With MyWdApp.ActiveDocument  '�������� �������� ��������
'        On Error Resume Next ' ��������� ������
''������ � ������� ������������:
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
''''������ � ��������� ��������� ���������:
'''i = 1
''    For i = 1 To n - 1 Step 1
''        .Bookmarks("WdEvColumnArr").Range.Text = "���� � " & "." & Chr(32) & EvidArray(i) & Chr(10) & Chr(10)
''    Next i
''���������� ������ � ����������
'        .Close SaveChanges:=wdSaveChanges
'    End With
'Set MyWdDoc = Nothing
'End Sub
''
'Private Sub AE_ErText()
''��������� ������� ����� ����� �������
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
''��������� ���������� ����� "�����������"
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
''��������� ��������� ����� "�����������"
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
''��������� ���������� ����� "�����������"
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
''��������� ��������� ����� "�����������"
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
''��������� ���������� ����� "����"
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
''��������� ��������� ����� "����"
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
''��������� ���������� ����� "���������� ��������"
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
''��������� ��������� ����� "���������� ��������"
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
''''������� ������ ��������
''''    ������� �������:
'''N = i + 1
'''    For i = 1 To N - 1
'''        Selection.TypeParagraph
'''        Selection.TypeText Text:="���.���� �5.2/" & WithDoc.strDN & " �� " & WithDoc.dtFirstDay & Chr(10) & _
'''            "�� ���������� " & WithDoc.strCsCat & " �" & WithDoc.strCsNum & Chr(10)
'''        Selection.TypeParagraph
'''        Selection.TypeText Text:=EvidArray(i) & Chr(10)
'''        Selection.TypeParagraph
''''          ������� ������ �� ���� "������� ����"
'''        Selection.InsertCrossReference ReferenceType:="��������", ReferenceKind:= _
'''            wdContentText, ReferenceItem:="ActualDate", InsertAsHyperlink:=True, _
'''            IncludePosition:=False, SeparateNumbers:=False, SeparatorString:=" "
'''        Selection.MoveLeft Unit:=wdCharacter, Count:=1, Extend:=wdExtend
'''        Selection.Font.Name = "Times New Roman"
'''        Selection.Font.Size = 14
'''        Selection.MoveRight Unit:=wdCharacter, Count:=1
'''        Selection.TypeParagraph
'''        Selection.TypeText Text:="�������" & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab _
'''            & vbTab & "�.�. �������" & Chr(10) & "��������" & vbTab & vbTab & vbTab & vbTab _
'''            & vbTab & vbTab & vbTab & Chr(10) '& "�.�. ��������"
'''        Selection.TypeParagraph
'''    Next i
'''End Sub
''
'
''
'Static Function Print_EvPackage() As String
''������� ������ �������� �������� � �.�. ������������ �������������
''Dim strTx1 As String, strTx2 As String
''strTx2 = "��������� �������� ������ ����� ������� ������"
''    If StrComp(frmNewEF.colEvidences(), "��������� �������") = True Then
''        strTx1 = "�����������"
''    Else: strTx1 = "�����������"
''    Debug.Print "strTx1 = ", strTx1
''    End If
''Dim EvNumArr1(1, 1 To 2) '�������� �������
''    EvNumArr1(1, 1) = "������������ �������������� ���������� ��������, ����������� � " & strEvPackage & _
''           Chr(32) & strTx1 & Chr(32) & strTx2
''    EvNumArr1(1, 2) = "������������ �������������� ���������� ��������, ����������� � " & strEvPackage & _
''           Chr(32) & strTx1 & Chr(32) & strTx2
'''Print_EvNumArr1 = EvNumArr1(1, intEvGRCount)
'End Function
''
''Static Function Print_PacIntegrity() As String
'''����������� ��������
'''Dim strTx1 As String, strTx2 As String, strTx3 As String, strTx4 As String
'''    strTx1 = "����������� �������� �� ��������, ����������"
'''    strTx2 = "��� ����������� ����������� �������� ����������. ��� �������� �������� �� ���"
'''    strTx3 = "��� ������-������������������� ������������,"
'''    strTx4 = "��������, ���������� � ����������� � � ���������������� ������� � ������������ ���������������. "
'''Dim EvNumArr2(1, 1 To 2) '�������� �������
'''    EvNumArr2(1, 1) = strTx1 & " ���������������� ������� " & strTx2 & " ��� �������� ������: " _
'''    & Chr(10) & _
'''    "������, �������������� " & strTx3 & " ������������� " & strTx4
'''    EvNumArr2(1, 2) = strTx1 & " ��������������� ��������  " & strTx2 & " ���� ��������� ��������� �������: " & Chr(10) & _
'''    "�������, �������������� " & strTx3 & " ������������� " & strTx4
''''Print_EvNumArr2 = EvNumArr2(1, intEvGRCount)
''End Function
''''
''Static Function Print_EvStamp() As String
'''������ ������ ����������
'''Dim strTx As String
'''    strTx = "��������� �������� ������ ����� ������� ������ " & Chr(171) & _
'''        "��������������� ������� �������� ��������� ���������� ��������. ������� ���������� �������-����������� ���������. ��� �������" _
'''        & Chr(187)
'''Dim EvNumArr3(1, 1 To 2) '�������� �������
'''    EvNumArr3(1, 1) = "������������ �������������� ���������, ��������� " & strTx
'''    EvNumArr3(1, 2) = "������������ �������������� ���������, ��������� " & strTx
''''Print_EvNumArr3 = EvNumArr3(1, intEvGRCount)
''End Function
'''
'Private Static Sub Print_NumColumn()
''������������ ������ � �������:
'Dim i As Integer, n As Integer
'Dim tmpEF As String, tmpGr As String, tmpEv As String
'    For i = 1 To CInt(colEvidences("GrCount" & tmpEF & "000" & "0000"))
'        tmpGr = CStr(Format(i, "#000"))
'Debug.Print "������ ����� �� - " & colEvidences("Owner" & tmpEF & tmpGr & "0000")
'        Selection.TypeText Text:=colEvidences("Owner" & tmpEF & tmpGr & "0000")
'            For n = 1 To CInt(colEvidences("EvCount" & tmpEF & tmpGr & "0000"))
'                tmpEv = CStr(Format(n, "#0000"))
'Debug.Print "������ �� - " & colEvidences("EvName" & tmpEF & tmpGr & tmpEv)
'                Selection.TypeText Text:=n & "." & Chr(32) _
'                    & colEvidences("EvName" & tmpEF & tmpGr & tmpEv) & Chr(10)
'            Next n
'    Next i
'End Sub

'Private Sub Open_EvidForm()
''��������� �������� ����� "������������ ��������������"
'    Dim tmp As String '��������� ����������-�������
'        tmp = Me.lblEFCount '��������� ������-������ ����� �������� ����� frmNewEF
'    Dim frmD As frmEvidences '��������� ����� �����
'        mdCount.fEvidCount = mdCount.fEvidCount + 1 '��������� �������� ���� "������������ ��������������"
'    Set frmD = New frmEvidences
''���������� ����� � ��������� ��������� ����
'        frmD.Caption = mdPrintDoc.DocCat & Chr(32) & tmp 'mdCount.fEvidCount ���������� �������� �����
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
''������������ ������ � �������:
''n = i + 1
''    For i = 1 To n - 1
'''        Selection.TypeText Text:=i & "." & Chr(32) & EvidArray(i) & Chr(10)
''    Next i
'End Sub
''
'Public Static Sub Print_ColumnArr(ByVal i As Byte, ByVal n As Byte)
''������ � �������:
''n = i + 1
''    For i = 1 To n - 1
'''        Selection.TypeText Text:=EvidArray(i) & Chr(10)
''        Next i
'End Sub
'''
'Public Static Sub Print_LineArr(ByVal i As Byte, ByVal n As Byte)
''������ � ������
''n = i + 1
''    For i = 1 To n - 1
'''       Selection.TypeText Text:=EvidArray(i) & ", "
''        Next i
'End Sub
'
