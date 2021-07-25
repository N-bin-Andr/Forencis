VERSION 5.00
Begin VB.Form frmLogin 
   BackColor       =   &H00BBC0AC&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Авторизация"
   ClientHeight    =   1485
   ClientLeft      =   45
   ClientTop       =   465
   ClientWidth     =   4200
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1485
   ScaleWidth      =   4200
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Tag             =   "1059"
   Begin VB.CommandButton cmdCancel 
      BackColor       =   &H00C0C0FF&
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   360
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   5
      Tag             =   "1063"
      Top             =   960
      Width           =   1140
   End
   Begin VB.CommandButton cmdOK 
      BackColor       =   &H00FFC0C0&
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   360
      Left            =   2880
      Style           =   1  'Graphical
      TabIndex        =   4
      Tag             =   "1062"
      Top             =   960
      Width           =   1140
   End
   Begin VB.TextBox txtPassword 
      Alignment       =   2  'Центровка
      BackColor       =   &H00F0FBE1&
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1425
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   525
      Width           =   2550
   End
   Begin VB.TextBox txtUserName 
      Alignment       =   2  'Центровка
      BackColor       =   &H00F0FBE1&
      Height          =   285
      Left            =   1425
      TabIndex        =   3
      Top             =   135
      Width           =   2550
   End
   Begin VB.Label lblLabels 
      BackColor       =   &H00BBC0AC&
      Caption         =   "&Пароль:"
      BeginProperty Font 
         Name            =   "MS Reference Sans Serif"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   248
      Index           =   1
      Left            =   105
      TabIndex        =   0
      Tag             =   "1061"
      Top             =   540
      Width           =   1080
   End
   Begin VB.Label lblLabels 
      BackColor       =   &H00BBC0AC&
      Caption         =   "&Пользователь:"
      BeginProperty Font 
         Name            =   "MS Reference Sans Serif"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   105
      TabIndex        =   2
      Tag             =   "1060"
      Top             =   150
      Width           =   1320
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'форма "frmLogin"
'Обработка данных при создании медико-криминалистических заключений
'Дата создания: 01.06.2016
'@version 0.0.1
'@author Andr.Nab.n@gmail.com
Option Explicit
Private Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpbuffer As String, nSize As Long) As Long
Public OK As Boolean
'
Private Sub Form_Load()
'Загрузка формы
    Dim sBuffer As String
    Dim lSize As Long
    
    LoadResStrings Me

    sBuffer = Space$(255)
    lSize = Len(sBuffer)
    Call GetUserName(sBuffer, lSize)
    If lSize > 0 Then
        txtUserName.Text = Left$(sBuffer, lSize)
    Else
        txtUserName.Text = vbNullString
    End If
End Sub
'
Private Sub cmdCancel_Click()
    OK = False
    Me.Hide
End Sub
'
Private Sub cmdOK_Click()
'Процедура нажатия кнопки "ОК"
    'ToDo: create test for correct password
    'check for correct password
    If txtPassword.Text = "" Then
        OK = True
        Me.Hide
    Else
        MsgBox "Неравильный пароль, введите заново!", , "Login"
        txtPassword.SetFocus
        txtPassword.SelStart = 0
        txtPassword.SelLength = Len(txtPassword.Text)
    End If
End Sub

