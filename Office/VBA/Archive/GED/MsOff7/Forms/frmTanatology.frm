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
   ScaleMode       =   0  '����������������
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
      Alignment       =   2  '���������
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
      Alignment       =   2  '���������
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
      Caption         =   "������� �������������� ����"
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
         ScrollBars      =   2  '���������
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
      Caption         =   "&��������"
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
      Alignment       =   2  '���������
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
      Alignment       =   2  '���������
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
      Caption         =   "�����������"
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
         Alignment       =   2  '���������
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
         Alignment       =   2  '���������
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
         Alignment       =   2  '���������
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
         BackStyle       =   0  '���������
         Caption         =   "���� ��������"
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
         BackStyle       =   0  '���������
         Caption         =   "���� ������"
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
         Alignment       =   2  '���������
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  '���������
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
         ForeColor       =   &H00404000&
         Height          =   375
         Left            =   3960
         TabIndex        =   43
         Top             =   2520
         Width           =   495
      End
      Begin VB.Label lblInjPrSurName 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  '���������
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
         ForeColor       =   &H00404000&
         Height          =   375
         Left            =   240
         TabIndex        =   36
         Top             =   480
         Width           =   1695
      End
      Begin VB.Label lblInjPrName 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  '���������
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
         ForeColor       =   &H00404000&
         Height          =   375
         Left            =   240
         TabIndex        =   35
         Top             =   960
         Width           =   1695
      End
      Begin VB.Label lblInjPrMidName 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  '���������
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
         ForeColor       =   &H00404000&
         Height          =   375
         Left            =   240
         TabIndex        =   34
         Top             =   1440
         Width           =   1695
      End
      Begin VB.Label lblInjPrBirthday 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  '���������
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
         ForeColor       =   &H00404000&
         Height          =   255
         Left            =   240
         TabIndex        =   33
         Top             =   1920
         Width           =   1575
      End
   End
   Begin VB.TextBox txtEFFirstDay 
      Alignment       =   2  '���������
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
      Caption         =   "�����������"
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
         Alignment       =   2  '���������
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
         Alignment       =   2  '���������
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
         Alignment       =   2  '���������
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
         BackStyle       =   0  '���������
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
         Height          =   375
         Left            =   120
         TabIndex        =   49
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label lblCrSex 
         Alignment       =   2  '���������
         BackColor       =   &H00CFC2AC&
         BackStyle       =   0  '���������
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
         Left            =   3600
         TabIndex        =   42
         Top             =   3960
         Width           =   495
      End
      Begin VB.Label lblCrOprAr 
         BackColor       =   &H00CFC2AC&
         BackStyle       =   0  '���������
         Caption         =   "�����"
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
         BackStyle       =   0  '���������
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
         Height          =   375
         Left            =   120
         TabIndex        =   30
         Top             =   2040
         Width           =   1335
      End
      Begin VB.Label lblCrSurName 
         BackColor       =   &H00CFC2AC&
         BackStyle       =   0  '���������
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
         Height          =   375
         Left            =   120
         TabIndex        =   29
         Top             =   2520
         Width           =   1335
      End
      Begin VB.Label lblCrMidName 
         BackColor       =   &H00CFC2AC&
         BackStyle       =   0  '���������
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
         Height          =   375
         Left            =   120
         TabIndex        =   28
         Top             =   3480
         Width           =   1335
      End
      Begin VB.Label lblCrName 
         BackColor       =   &H00CFC2AC&
         BackStyle       =   0  '���������
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
      Caption         =   "&������"
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
      ForeColor       =   &H00404000&
      Height          =   255
      Left            =   120
      TabIndex        =   48
      Top             =   1980
      Width           =   1575
   End
   Begin VB.Label lblCsNum 
      BackColor       =   &H00BBC2AC&
      Caption         =   "�"
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
      Caption         =   "���������:"
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
      ForeColor       =   &H00404000&
      Height          =   360
      Left            =   120
      TabIndex        =   40
      Top             =   735
      Width           =   2055
   End
   Begin VB.Label lblCsPrvDate 
      Alignment       =   2  '���������
      BackColor       =   &H00BBC2AC&
      Caption         =   "��"
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
      Caption         =   "�"
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
      Alignment       =   2  '���������
      BackColor       =   &H00BBC2AC&
      Caption         =   "���� ������"
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
'����� "����� ������������"
'@author Andr.N@_bin
'@E-mail Andr.Nab.n@gmail.com
'�����: BackColor - &H00BBC2AC& / &H00FEF7F1&
Option Explicit
'���������� �������
Public newCase As clmCase          '��������� ������ "����"
Public newCor As clmHR_Official    '��������� ������ "�����������" Coroner
Public SvcService As Object        '��������� ���������� svcsvc.dll
Public newEF As clmExpertFindings  '��������� ������ "���������� ��������"
Public newInjPr As clmEntrants     '��������� ������ "�����������"InjuredPerson
Public newDate As clmCaseDate      '��������� ������ "����"
Public newDOC As clmCreateDoc      '��������� ������ "����������� ���������"
'����������
Public newExpert As String       '���������� "�������"
'������:
Const Msg As String = "������ ����� ������!"
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
        .txtCsIndex.Width = 520.932
        .txtCsIndex.Left = 1123.585
        .Caption = "�������� ������ ���������"  '" �������� ������ ��������� " & "����: " & .newExpert &   .newEF.categories & ")"
    End With
 Move (Screen.Width - Width) / 2, (Screen.Height - Height) / 3.7 '/ 17.7
End Sub
'
Private Sub Form_Unload(Cancel As Integer)
'��������� �������� �����
    If MsgBox("�������?", vbYesNo, "�����?") = vbYes Then
        Set newCase = Nothing       '��������� ������ "����"
        Set newDate = Nothing       '��������� ������ "����"
        Set newCor = Nothing        '��������� ������ "�����������"
        Call termClass
    Else
        Cancel = 1
    End If
End Sub
'
Private Sub initClass()
'������������� ��������� �������
    Set newEF = New clmExpertFindings   '��������� ������ "���������� ��������"
        newEF.categories = frmMDI.category
    Set newInjPr = New clmEntrants      '��������� ������ "�����������"
End Sub
'
Private Sub termClass()
'����������� ��������� ������������ �������
    Set newEF = Nothing         '��������� ������ "���������� ��������"
    Set newInjPr = Nothing      '��������� ������ "�����������"
    Set newDOC = Nothing  '��������� ������ "����������� ���������"
End Sub
'
'+++++++++++++++++++++++++++++++++++++++ � � � � � �  � � � � �: ++++++++++++++++++++++++++++++++++++
'
Public Sub create_Analysis()
'�������� ������� ����������� �� ������������
Set newDOC.MyWdApp = New Word.Application '��������� ����������
'1)�������� ����� ��� ����������� ����������:
Dim nameFolder As String
    nameFolder = frmAnalysis.Caption & "_" & frmNewResearch.newEF.getNumber
Dim nameDOC As String '��� ��������� ������������_�����_���
'2) �������� ����� ���������� � ���������� �� � ��������� �����:
Dim X As Object '���������� ���������
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
'�������� ���������� "�����"
'1)�������� ����� ��� ����������� ����������:
Dim nameFolder As String
    nameFolder = frmDocList.Caption & "_" & frmNewResearch.newEF.getNumber
Dim nameDOC As String '��� ��������� ������������_�����_���
'2) �������� ����� ���������� � ���������� �� � ��������� �����:
Dim X As Object '���������� ���������
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
'����������� ��������� ���������� ���������� = �� ��������� �������������/�����������
'��������� ����������� + ������ + ����� + ��� + ����������� ���� �� ���������� ���������� ����
    create_legalGround = newEF.printDefinition & newCor.print_Cor & _
    ", ����������� " & newEF.rulingDate & newCase.printCsData & "."
Debug.Print "��.��������� = " & create_legalGround
End Function
'
Public Static Sub addExpert()
'1)������� �����:
Me.Visible = False
'������� ������ �������� �� ������
Set SvcService = CreateObject("Svcsvc.Service") '������ ���������� svcsvc.dll
        newExpert = SvcService.SelectValue("��������� �.�." & vbCrLf & _
                                        "�������� �.�." & vbCrLf & _
                                        "������ �.�." & vbCrLf & _
                                        "��������� �.�." & vbCrLf & _
                                        "�������� �.�." & vbCrLf & _
                                        "���� �.�." & vbCrLf & _
                                        "�������� �.�." & vbCrLf & _
                                        "������� �.�." & vbCrLf & _
                                        "���������� �.�." & vbCrLf & _
                                        "���������� �.�." & vbCrLf & _
                                        "������� �.�.", _
                                        "�������� ��������", True)
                                        'Debug.Print " addExpert: " &  addExpert
Set SvcService = Nothing
End Sub
'
Private Sub txtEnabled()
'���������� ��������� �����
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
        .Caption = " �������� ������ ��������� �" & newEF.getFullNumber '"����: " & Me.newExpert &
        With .txtEFFirstDay
            .SetFocus '�������� ������
            Call txtEnter(Me.txtEFFirstDay)
            .Text = newDate.DateNow
        End With
    End With
Debug.Print "����� ����������: �" & newEF.number
End Sub
'2) ���� ������ ����������
Private Sub txtEFFirstDay_GotFocus()
'���� ������ ����������
    Call txtEnter(txtEFFirstDay)
    '���� ������ = ������� ����
    txtEFFirstDay.Text = newDate.DateNow
End Sub
'
Private Sub txtEFFirstDay_LostFocus()
'���� ������ ����������
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
        .AddItem "���������"
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
        If .Text = "" Then
            .Text = "���������"
        End If
            newEF.categories = .Text
            .BackColor = frmColor.��������� 'RGB(200, 256, 200)
    End With
'Debug.Print "��������� ����������= ", newEF.categories
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
'���� ��������� �������������
 Dim tmp As Double, dt As Date
 With txtCsPrvDate
    If .Text = "" Then
        Call txt_Exit(txtCsPrvDate)
    Else
        With newDate
            dt = .validateDate(.ExamDate(txtCsPrvDate.Text)) 'newEF.rulingDate
            '��������� ���: ���� ������ > ���� ��������� �������������
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
End Sub
'
Private Sub txtCsIndex_LostFocus()
'������ ����
    With txtCsIndex
        If Len(.Text) = 0 Then
            .Width = 1398.29
            .Left = 246.226
            Call txt_Exit(txtCsIndex)
        Else: newCase.index = StrConv(.Text, vbProperCase)
            .BackColor = frmColor.��������� 'RGB(200, 256, 200)
            .Text = newCase.index
        End If
    End With
'Debug.Print "������ ����: ", newCase.index
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
   txtInjPrDecease.Text = newEF.rulingDate
End Sub

Private Sub txtInjPrDecease_LostFocus()
'���� ������
 Dim tmp As Double, dt As Date ' str As String
    If txtInjPrDecease.Text = "" Then
        Call txt_Exit(txtInjPrDecease)
    Else
        With newDate
            dt = .validateDate(.ExamDate(txtInjPrDecease.Text))
            '��������� ���: ���� ������ <= ���� ��������� �������������
                '                       <= ���� ������ ����������
                '                       <= ���� ��������
            Do
                tmp = .compareDt(newEF.rulingDate, dt)
                If tmp > 0 Then
                    MsgBox "���� ������ ������ ���� ��������� �������������!", _
                            vbCritical, Msg
                    dt = .ExamDate(InputBox("��������� ������� ���� ������!", _
                                        "���� ����", newEF.rulingDate))
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
        Call txtEnter(Me.txtInjPrAutopsy)
        .Text = newEF.firstDay
    End With
'Debug.Print "���� ������: " & newInjPr.decease
End Sub
''
'22)���� ��������
Private Sub txtInjPrAutopsy_GotFocus()
'���� ��������
   Call txtEnter(txtInjPrAutopsy)
   '���� �������� = ���� ������
   txtInjPrAutopsy.Text = newEF.firstDay
End Sub
'
Private Sub txtInjPrAutopsy_LostFocus()
'���� ��������
 Dim tmp As Double, dt As Date, str As String
    If txtInjPrAutopsy.Text = "" Then
        Call txt_Exit(txtInjPrAutopsy)
    Else
        With newDate
            dt = .validateDate(.ExamDate(txtInjPrAutopsy.Text))
            '��������� ���: ���� �������� >= ���� ��������� �������������
                '                         >= ���� ������ ����������
                '                         >= ���� ������
            Do
                tmp = .compareDt(newEF.firstDay, dt)
                If tmp < 0 Then
                    MsgBox "���� �������� ������ ���� ������ ����������!", _
                            vbCritical, Msg
                    dt = .ExamDate(InputBox("��������� ������� ���� ��������!", _
                                        "���� ����", newEF.firstDay))
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
    With Me.txtFactCase
        Call txtEnter(txtFactCase)
        .Text = newDate.dateToString(newEF.rulingDate) & " ���� " & newInjPr.create_InitialslName
    End With
'Debug.Print "���� ��������: " & newInjPr.autopsyDate
'On Error Resume Next ' ��������� ������
End Sub
'
'23)
Private Sub txtFactCase_GotFocus()
'������� �������������� ����
    Call txtEnter(txtFactCase)
    txtFactCase.Text = newDate.dateToString(newEF.rulingDate) & _
                        " ���� " & newInjPr.print_InjPersons
'&H00FFE8DF&
'&H00CDEEEF&
End Sub
'
Private Sub txtFactCase_LostFocus()
 '������� �������������� ����
    If txtInjPrDecease.Text = "" Then
        Call txt_Exit(txtFactCase)
    Else
        With txtInjPrDecease
            newCase.crimCondition = .Text
            .BackColor = RGB(200, 256, 200)
        End With
    End If
'Debug.Print "������� �������������� ����: " & newCase.crimCondition
End Sub
'
'+++++++++++++++++++++++++++++++++++++++ � � � � � �: ++++++++++++++++++++++++++++++++++++++++++++++++++
Private Sub cmdOK1_Click()
'������ "��"
'1)��������� ��������:
 Call addExpert
 Me.Caption = "����: " & Me.newExpert & Chr(32) & " �������� ������ ��������� �" & newEF.getFullNumber
'2) ��������� ���������� �� ���������
Set newDOC = New clmCreateDoc '��������� ������ "����������� ���������"
Dim tmpname As String
    tmpname = Me.newExpert & "_" & Me.newEF.getNumber
Dim tmp As String
    With newDOC
        .dirExpert = .Create_mainFolders(frmMDI.newRoot, tmpname)
        tmpname = "����_" & Me.newEF.getNumber
        tmp = .Create_mainFolders(.dirExpert, tmpname)
    End With
    frmAnalysis.Show
    frmNewResearch.Visible = False  '������� ����� "����� ������������"
End Sub
'
Private Sub cmdCancel1_Click()
'������ "������"
    Unload Me
End Sub
'
Private Sub txtDisabled()
'���������� ��������� �����
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
'������ "��������"
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
