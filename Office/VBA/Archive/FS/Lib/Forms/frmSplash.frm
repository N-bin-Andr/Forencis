VERSION 5.00
Begin VB.Form frmSplash 
   BackColor       =   &H00000000&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   9990
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   8010
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9990
   ScaleWidth      =   8010
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.Frame fraMainFrame 
      BackColor       =   &H00FF8080&
      Height          =   9495
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   7380
      Begin VB.PictureBox picLogo1 
         DrawMode        =   1  'Blackness
         Height          =   4755
         Left            =   120
         Picture         =   "frmSplash.frx":0000
         ScaleHeight     =   4695
         ScaleWidth      =   7035
         TabIndex        =   9
         Top             =   3960
         Width           =   7095
      End
      Begin VB.Label lblCopyright 
         BackColor       =   &H00FF8080&
         Caption         =   "Copyright"
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   9.75
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2640
         TabIndex        =   8
         Tag             =   "1051"
         Top             =   9120
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Label lblCompany 
         Alignment       =   2  'Центровка
         BackColor       =   &H00FF8080&
         Caption         =   "Andr.Nab.n@gmail.com"
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   12
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404080&
         Height          =   375
         Left            =   4560
         TabIndex        =   7
         Tag             =   "1052"
         Top             =   8760
         Width           =   2655
      End
      Begin VB.Label lblVersion 
         Alignment       =   1  'Правая привязка
         AutoSize        =   -1  'True
         BackColor       =   &H00FF8080&
         Caption         =   "Version 1.0.0"
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   9.75
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   225
         Left            =   6000
         TabIndex        =   6
         Tag             =   "1054"
         Top             =   9120
         Visible         =   0   'False
         Width           =   1125
      End
      Begin VB.Label lblPlatform 
         Alignment       =   1  'Правая привязка
         AutoSize        =   -1  'True
         BackColor       =   &H00FF8080&
         Caption         =   "VBA Platform "
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   9.75
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   225
         Left            =   1320
         TabIndex        =   5
         Tag             =   "1055"
         Top             =   9120
         Visible         =   0   'False
         Width           =   1230
      End
      Begin VB.Label lblCompanyProduct 
         Alignment       =   2  'Центровка
         AutoSize        =   -1  'True
         BackColor       =   &H00FF8080&
         Caption         =   "CompanyProduct"
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   9.75
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   225
         Left            =   240
         TabIndex        =   4
         Tag             =   "1056"
         Top             =   8760
         Width           =   1785
      End
      Begin VB.Label lblProductName 
         AutoSize        =   -1  'True
         BackColor       =   &H00FF8080&
         Caption         =   "Product Name"
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   9.75
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   225
         Left            =   2040
         TabIndex        =   3
         Tag             =   "1057"
         Top             =   8760
         Width           =   1245
      End
      Begin VB.Label lblLicenseTo 
         Alignment       =   2  'Центровка
         BackColor       =   &H00FF8080&
         Caption         =   "LicenseTo"
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   9.75
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   240
         TabIndex        =   2
         Tag             =   "1058"
         Top             =   9120
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Label lblWarning 
         BackColor       =   &H00FF8080&
         Caption         =   "Warning"
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   9.75
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   3360
         TabIndex        =   1
         Tag             =   "1053"
         Top             =   8760
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Image picLogo2 
         BorderStyle     =   1  'Фиксировано один
         Height          =   3720
         Left            =   0
         Picture         =   "frmSplash.frx":220E5
         Top             =   120
         Width           =   9840
      End
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Форма "frmSplash"
'Обработка данных при создании медико-криминалистических заключений
'Дата создания: 01.06.2016
'@version 0.0.1
'@author Andr.Nab.n@gmail.com
Option Explicit

Private Sub Form_Load()
    LoadResStrings Me
    lblVersion.Caption = "Version " & App.Major & "." & App.Minor & "." & App.Revision
    lblProductName.Caption = App.Title
End Sub

