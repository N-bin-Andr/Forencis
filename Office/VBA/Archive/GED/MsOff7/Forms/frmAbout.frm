VERSION 5.00
Begin VB.Form frmAbout 
   BackColor       =   &H00FDF0EC&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "О приложении:"
   ClientHeight    =   4605
   ClientLeft      =   2340
   ClientTop       =   1935
   ClientWidth     =   5355
   ClipControls    =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3178.453
   ScaleMode       =   0  'Пользовательское
   ScaleWidth      =   5028.622
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox picIcon 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'Нет
      ClipControls    =   0   'False
      Height          =   3540
      Left            =   240
      Picture         =   "frmAbout.frx":0000
      ScaleHeight     =   2486.26
      ScaleMode       =   0  'Пользовательское
      ScaleWidth      =   1264.2
      TabIndex        =   1
      Top             =   240
      Width           =   1800
   End
   Begin VB.CommandButton cmdOK 
      BackColor       =   &H00F0FBE1&
      Cancel          =   -1  'True
      Caption         =   "&OK"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   9.75
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   3840
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   4080
      Width           =   1260
   End
   Begin VB.CommandButton cmdSysInfo 
      BackColor       =   &H00FCF2E4&
      Caption         =   "&Информация о системе"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   9.75
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   4080
      Width           =   2565
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Index           =   3
      X1              =   2028.352
      X2              =   4732.821
      Y1              =   1242.392
      Y2              =   1242.392
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Index           =   2
      X1              =   0
      X2              =   3042.528
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Центровка
      BackColor       =   &H00FCF2E4&
      BorderStyle     =   1  'Фиксировано один
      Caption         =   "                                                 Ты, если что, заходи... Andr.Nab.n@gmail.com"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   12
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   1335
      Left            =   2160
      TabIndex        =   7
      Top             =   240
      Width           =   2895
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderStyle     =   6  'Внутренняя заливка
      Index           =   1
      X1              =   225.372
      X2              =   4732.821
      Y1              =   2733.262
      Y2              =   2733.262
   End
   Begin VB.Label lblDescription 
      Alignment       =   2  'Центровка
      BackColor       =   &H00FCF2E4&
      BackStyle       =   0  'Прозрачно
      Caption         =   "Для создания бланков."
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   12
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808000&
      Height          =   330
      Left            =   2160
      TabIndex        =   3
      Top             =   2280
      Width           =   2895
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Центровка
      BackColor       =   &H00FCF2E4&
      BackStyle       =   0  'Прозрачно
      Caption         =   "Application Title"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   12
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808000&
      Height          =   360
      Left            =   2160
      TabIndex        =   5
      Top             =   1920
      Width           =   2895
   End
   Begin VB.Label lblVersion 
      Alignment       =   2  'Центровка
      BackColor       =   &H00FCF2E4&
      BackStyle       =   0  'Прозрачно
      Caption         =   "Version"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   12
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808000&
      Height          =   315
      Left            =   2160
      TabIndex        =   6
      Top             =   2640
      Width           =   2895
   End
   Begin VB.Label lblDisclaimer 
      BackColor       =   &H00FCF2E4&
      BackStyle       =   0  'Прозрачно
      Caption         =   "Все права защищены."
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   8.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   225
      Left            =   3360
      TabIndex        =   4
      Top             =   3720
      Width           =   1695
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Информационная форма
'@author Andr.N@_bin
'@E-mail Andr.Nab.n@gmail.com
Option Explicit
'Reg Key Security Options...
Const READ_CONTROL = &H20000
Const KEY_QUERY_VALUE = &H1
Const KEY_SET_VALUE = &H2
Const KEY_CREATE_SUB_KEY = &H4
Const KEY_ENUMERATE_SUB_KEYS = &H8
Const KEY_NOTIFY = &H10
Const KEY_CREATE_LINK = &H20
Const KEY_ALL_ACCESS = KEY_QUERY_VALUE + KEY_SET_VALUE + _
                       KEY_CREATE_SUB_KEY + KEY_ENUMERATE_SUB_KEYS + _
                       KEY_NOTIFY + KEY_CREATE_LINK + READ_CONTROL
'Reg Key ROOT Types...
Const HKEY_LOCAL_MACHINE = &H80000002
Const ERROR_SUCCESS = 0
Const REG_SZ = 1    'Unicode nul terminated string
Const REG_DWORD = 4 '32-bit number
'
Const gREGKEYSYSINFOLOC = "SOFTWARE\Microsoft\Shared Tools Location"
Const gREGVALSYSINFOLOC = "MSINFO"
Const gREGKEYSYSINFO = "SOFTWARE\Microsoft\Shared Tools\MSINFO"
Const gREGVALSYSINFO = "PATH"
'
Private Declare Function RegOpenKeyEx Lib "advapi32" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, ByRef phkResult As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, ByRef lpType As Long, ByVal lpData As String, ByRef lpcbData As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32" (ByVal hKey As Long) As Long
'
Private Sub cmdSysInfo_Click()
  Call StartSysInfo
End Sub
'
Private Sub cmdOK_Click()
  Unload Me
End Sub
'
Private Sub Form_Load()
    Me.Caption = "About " & App.Title
    lblVersion.Caption = "Version " & App.Major & "." & App.Minor & "." & App.Revision
    lblTitle.Caption = App.Title
    Me.BackColor = RGB(201, 222, 225)
End Sub
'
Public Sub StartSysInfo()
    On Error GoTo SysInfoErr
'
    Dim rc As Long
    Dim SysInfoPath As String
    'Try To Get System Info Program Path\Name From Registry...
    If GetKeyValue(HKEY_LOCAL_MACHINE, gREGKEYSYSINFO, gREGVALSYSINFO, SysInfoPath) Then
    'Try To Get System Info Program Path Only From Registry...
    ElseIf GetKeyValue(HKEY_LOCAL_MACHINE, gREGKEYSYSINFOLOC, gREGVALSYSINFOLOC, SysInfoPath) Then
    ' Validate Existance Of Known 32 Bit File Version
        If (Dir(SysInfoPath & "\MSINFO32.EXE") <> "") Then
            SysInfoPath = SysInfoPath & "\MSINFO32.EXE"
        'Ошибка: "файл не найден..." (Error - File Can Not Be Found...)
        Else
            GoTo SysInfoErr
        End If
    'Error - Registry Entry Can Not Be Found...
    Else
        GoTo SysInfoErr
    End If
    Call Shell(SysInfoPath, vbNormalFocus)
    Exit Sub
SysInfoErr:
    MsgBox "В данный момент системная информация не доступна!", vbOKOnly '"System Information Is Unavailable At This Time"
End Sub
'
Public Function GetKeyValue(KeyRoot As Long, KeyName As String, SubKeyRef As String, ByRef KeyVal As String) As Boolean
    Dim i As Long               'Loop Counter
    Dim rc As Long              'Return Code
    Dim hKey As Long            'Handle To An Open Registry Key
    Dim hDepth As Long          '
    Dim KeyValType As Long      'Data Type Of A Registry Key
    Dim tmpVal As String        'Tempory Storage For A Registry Key Value
    Dim KeyValSize As Long      'Size Of Registry Key Variable
    '------------------------------------------------------------
    ' Open RegKey Under KeyRoot {HKEY_LOCAL_MACHINE...}
    '------------------------------------------------------------
    rc = RegOpenKeyEx(KeyRoot, KeyName, 0, KEY_ALL_ACCESS, hKey) ' Open Registry Key
    If (rc <> ERROR_SUCCESS) Then GoTo GetKeyError  'Handle Error...
    tmpVal = String$(1024, 0)   'Allocate Variable Space
    KeyValSize = 1024           'Mark Variable Size
    '------------------------------------------------------------
    ' Retrieve Registry Key Value...
    '------------------------------------------------------------
    rc = RegQueryValueEx(hKey, SubKeyRef, 0, _
                         KeyValType, tmpVal, KeyValSize)    ' Get/Create Key Value
    If (rc <> ERROR_SUCCESS) Then GoTo GetKeyError          ' Handle Errors
    If (Asc(Mid(tmpVal, KeyValSize, 1)) = 0) Then           ' Win95 Adds Null Terminated String...
        tmpVal = Left(tmpVal, KeyValSize - 1)               ' Null Found, Extract From String
    Else                                                    ' WinNT Does NOT Null Terminate String...
        tmpVal = Left(tmpVal, KeyValSize)                   ' Null Not Found, Extract String Only
    End If
    '------------------------------------------------------------
    ' Determine Key Value Type For Conversion...
    '------------------------------------------------------------
    Select Case KeyValType                                  ' Search Data Types...
    Case REG_SZ                                             ' String Registry Key Data Type
        KeyVal = tmpVal                                     ' Copy String Value
    Case REG_DWORD                                          ' Double Word Registry Key Data Type
        For i = Len(tmpVal) To 1 Step -1                    ' Convert Each Bit
            KeyVal = KeyVal + Hex(Asc(Mid(tmpVal, i, 1)))   ' Build Value Char. By Char.
        Next
        KeyVal = Format$("&h" + KeyVal)                     ' Convert Double Word To String
    End Select
'
    GetKeyValue = True                                      ' Return Success
    rc = RegCloseKey(hKey)                                  ' Close Registry Key
    Exit Function                                           ' Exit
'
GetKeyError:      ' Cleanup After An Error Has Occured...
    KeyVal = ""                                             ' Set Return Val To Empty String
    GetKeyValue = False                                     ' Return Failure
    rc = RegCloseKey(hKey)                                  ' Close Registry Key
End Function

