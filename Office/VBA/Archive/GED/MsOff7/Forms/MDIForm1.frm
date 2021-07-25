VERSION 5.00
Begin VB.MDIForm frmMDI 
   BackColor       =   &H00808000&
   ClientHeight    =   9450
   ClientLeft      =   60
   ClientTop       =   855
   ClientWidth     =   12405
   LinkTopic       =   "MDIForm1"
   WindowState     =   2  'Развернуто
   Begin VB.Menu mnuFile 
      Caption         =   "&Создать"
      Begin VB.Menu mnuTanatology 
         Caption         =   "&Заключение по вскрытию"
      End
   End
   Begin VB.Menu mnuChange 
      Caption         =   "&Изменить"
      Begin VB.Menu mnuFolder 
         Caption         =   "&Папку для сохранения документов"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Справка"
   End
End
Attribute VB_Name = "frmMDI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Main Form
'@author Andr.N@_bin
'@E-mail Andr.Nab.n@gmail.com
Option Explicit
Private newFrm As Form
Private tmpcategory As String
Private SvcService As Object        'объект библиотеки svcsvc.dll
Private tmpnewRoot As String
'
'++++++++++++++++++  И Н И Ц И А Л И З А Ц И Я ++++++++++++++
'
Private Sub MDIForm_Initialize()
    Me.Caption = "ООЭ №3 Создание новых документов."
End Sub
'
Private Sub MDIForm_Unload(Cancel As Integer)
'Процедура выгрузки формы
    If MsgBox("Закрыть приложение?", vbYesNo, "Выход?") = vbYes Then
'        Set newFrm = Nothing
    Else
        Cancel = 1
    End If
End Sub
'
'++++++++++++++++++ И Н К А П С У Л Я Ц И Я ++++++++++++++++++
Private Property Let category(ByVal vData As String)
'категория создаваемого документа
    tmpcategory = vData
End Property
'
Friend Property Get category() As String
'категория создаваемого документа
    category = tmpcategory
'Debug.Print "category: " & category
End Property
'
Private Property Let newRoot(ByVal vData As String)
'новая директория, введенная пользователем
    tmpnewRoot = vData
End Property
'
Friend Property Get newRoot() As String
'новая директория, введенная пользователем
    newRoot = tmpnewRoot
'Debug.Print "новая директория, введенная пользователем: " & newRoot
End Property
'
'++++++++++++++++++ М Е Т О Д Ы ++++++++++++++++++++++++++++++
Private Sub mnuHelp_Click()
frmAbout.Show
End Sub
'
Public Sub mnuTanatology_Click()
    category = "Танатология"
'    Set newFrm = frmNewResearch
        With frmNewResearch
'            Call .addExpert
            .Show
        End With
End Sub
'
Private Sub mnuFolder_Click()
    Set SvcService = CreateObject("Svcsvc.Service") 'объект библиотеки  svcsvc.dll
        newRoot = SvcService.SelectFolder("Выбираем папку для сохранения документов", "", &H10 + &H4000, "")
    Set SvcService = Nothing    'объект библиотеки  svcsvc.dll
'Debug.Print "новая директория, введенная пользователем: " & newRoot
End Sub

