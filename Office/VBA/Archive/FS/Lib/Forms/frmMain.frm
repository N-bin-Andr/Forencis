VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.MDIForm frmMain 
   BackColor       =   &H00808000&
   Caption         =   "��������������  �������� ������ ���������"
   ClientHeight    =   7095
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   9195
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.Toolbar tbToolBar 
      Align           =   1  '��������� �����
      Height          =   420
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   9195
      _ExtentX        =   16219
      _ExtentY        =   741
      ButtonWidth     =   609
      ButtonHeight    =   582
      Appearance      =   1
      ImageList       =   "imlToolbarIcons"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   3
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "New"
            Object.ToolTipText     =   "������� ����� ��������"
            ImageKey        =   "New"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "DocDir"
            Object.ToolTipText     =   "������� ����� ��� ���������� ����������"
            ImageKey        =   "Open"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar sbStatusBar 
      Align           =   2  '��������� ����
      Height          =   270
      Left            =   0
      TabIndex        =   0
      Top             =   6825
      Width           =   9195
      _ExtentX        =   16219
      _ExtentY        =   476
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   10557
            Text            =   "Status"
            TextSave        =   "Status"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            AutoSize        =   2
            TextSave        =   "17.06.2020"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            AutoSize        =   2
            TextSave        =   "13:38"
         EndProperty
      EndProperty
   End
   Begin MSComDlg.CommonDialog dlgCommonDialog 
      Left            =   1740
      Top             =   1290
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList imlToolbarIcons 
      Left            =   1740
      Top             =   1290
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   14
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0000
            Key             =   "New"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0112
            Key             =   "Open"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0224
            Key             =   "Save"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0336
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0448
            Key             =   "Cut"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":055A
            Key             =   "Copy"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":066C
            Key             =   "Paste"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":077E
            Key             =   "Bold"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0890
            Key             =   "Italic"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":09A2
            Key             =   "Underline"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0AB4
            Key             =   "Justify"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0BC6
            Key             =   "Align Left"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0CD8
            Key             =   "Center"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0DEA
            Key             =   "Align Right"
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuFile 
      Caption         =   "����"
      Begin VB.Menu mnuFileNew 
         Caption         =   "������� ����� ��������"
         Shortcut        =   ^N
      End
   End
   Begin VB.Menu mnuWindow 
      Caption         =   "���"
      WindowList      =   -1  'True
      Begin VB.Menu mnuWindowNewWindow 
         Caption         =   "������� ����� ����"
      End
      Begin VB.Menu mnuWindowBar0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuWindowCascade 
         Caption         =   "����������� ���� ��������"
      End
      Begin VB.Menu mnuWindowTileHorizontal 
         Caption         =   "����������� ���� �����"
      End
      Begin VB.Menu mnuWindowTileVertical 
         Caption         =   "����������� ���� �����������"
      End
      Begin VB.Menu mnuWindowArrangeIcons 
         Caption         =   "����������� ����"
      End
   End
   Begin VB.Menu mnuOption 
      Caption         =   "������"
      Begin VB.Menu mnuNewDocFolder 
         Caption         =   "�������� ����� ��� ���������� ����� ����������"
      End
      Begin VB.Menu mnuNewDotFolder 
         Caption         =   "�������� �����, ���������� ������� ����������"
      End
   End
   Begin VB.Menu mnuHelpAbout 
      Caption         =   "� ���������"
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'����� "frmMain"
'���� ��������: 01.06.2016
'@author Andr.Nab.n@gmail.com
Option Explicit
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, _
                                                                        ByVal wMsg As Long, _
                                                                        ByVal wParam As Long, _
                                                                        ByVal lParam As Any) As Long
Const EM_UNDO = &HC7
Private Declare Function OSWinHelp% Lib "user32" Alias "WinHelpA" (ByVal hwnd&, _
                                                                ByVal HelpFile$, _
                                                                ByVal wCommand%, _
                                                                dwData As Any)
Public newFrm As Form
Private tmpcategory As String       '��������� ���������
Private SvcService As Object        '������ ���������� svcsvc.dll
Private tmpnewRoot As String
'
Private Sub MDIForm_Initialize()
'������������� ����� Main
    Me.Caption = "�������� ����� ���-����. ����������."
End Sub
'
Private Sub MDIForm_Load()
'�������� Main_�����
    LoadResStrings Me
    Me.Left = GetSetting(App.Title, "Settings", "MainLeft", 1000)
    Me.Top = GetSetting(App.Title, "Settings", "MainTop", 1000)
    Me.Width = GetSetting(App.Title, "Settings", "MainWidth", 6500)
    Me.Height = GetSetting(App.Title, "Settings", "MainHeight", 6500)
    LoadNewDoc
'�������� ������ ������ ����
'Set mdCount.colForms = New Collection       '����� ��������� ��������� ����
'Set mdPrintDoc.colDocCat = New Collection
End Sub
'
Private Sub MDIForm_Unload(Cancel As Integer)
'�������� Main_�����
    If Me.WindowState <> vbMinimized Then
        SaveSetting App.Title, "Settings", "MainLeft", Me.Left
        SaveSetting App.Title, "Settings", "MainTop", Me.Top
        SaveSetting App.Title, "Settings", "MainWidth", Me.Width
        SaveSetting App.Title, "Settings", "MainHeight", Me.Height
    End If
'
    If MsgBox("������� ����������?", vbYesNo, "�����?") = vbYes Then
    Else
        Cancel = 1
    End If
End Sub
'
Private Sub LoadNewDoc()
'�������� ������ ���������
'1)�������� ���� �� ��������� � ����� � ��������� ����������:
  Call mdMainFolders.getMainDir(mdMainFolders.USER_DOT_DIR, mdMainFolders.DOT_INDEX)
  Call mdMainFolders.getMainDir(mdMainFolders.USER_DOC_DIR, mdMainFolders.DOC_INDEX)
'2) ������ � ������ frmEvCategories
    Set newFrm = New frmEvCategories
    With newFrm
        .Caption = "�������� ������ ��������� "
        .Height = 8625
        .Width = 6360
'        .Tag = ""
        .Show 'vbModal, Me
    End With
'Lib
'        mdCount.fEvCat_Count = mdCount.fEvCat_Count + 1 '������� ���� "��������� ����������"
'        frmD.Caption = "�������� ������ ��������� " & mdCount.fEvCat_Count
'        frmD.Height = 8625
'        frmD.Width = 6360
'        frmD.EvCat_ID.Caption = mdCount.fEvCat_Count
'        mdCount.colForms.Add frmD, "fEvCat" & mdCount.fEvCat_Count 'frmD.EvCat_ID.Caption
'        frmD.Show 'vbModal, Me
End Sub

Private Sub mnuHelpAbout_Click()
'����� ������� "� ���������"
    frmAbout.Show vbModal, Me
End Sub

Private Sub getFolderDir()
    Set SvcService = CreateObject("Svcsvc.Service") '������ ����������  svcsvc.dll
        newRoot = SvcService.SelectFolder("�������� ����� ��� ���������� ����������", "", &H10 + &H4000, "")
    Set SvcService = Nothing    '������ ����������  svcsvc.dll
'Debug.Print "����� ����������, ��������� �������������: " & newRoot
End Sub

Private Sub mnuNewDocFolder_Click()
'������� ����� ��� ���������� ����� ����������
    Call mdMainFolders.inputUserDir(DOC_INDEX)
End Sub

Private Sub mnuNewDotFolder_Click()
'������� ����� � ��������� ����� ����������
    Call mdMainFolders.inputUserDir(DOT_INDEX)
End Sub
'
Private Sub mnuWindowArrangeIcons_Click()
    Me.Arrange vbArrangeIcons
End Sub
'
Private Sub mnuWindowTileVertical_Click()
    Me.Arrange vbTileVertical
End Sub
'
Private Sub mnuWindowTileHorizontal_Click()
    Me.Arrange vbTileHorizontal
End Sub
'
Private Sub mnuWindowCascade_Click()
    Me.Arrange vbCascade
End Sub
'
Private Sub mnuWindowNewWindow_Click()
    LoadNewDoc
End Sub
'
'Private Sub mnuViewWebBrowser_Click()
'    'ToDo: Add 'mnuViewWebBrowser_Click' code.
'    MsgBox "Add 'mnuViewWebBrowser_Click' code."
'End Sub
'
'Private Sub mnuViewOptions_Click()
'    frmOptions.Show vbModal, Me
'End Sub
''
'Private Sub mnuViewRefresh_Click()
'    'ToDo: Add 'mnuViewRefresh_Click' code.
'    MsgBox "Add 'mnuViewRefresh_Click' code."
'End Sub
''
'Private Sub mnuViewStatusBar_Click()
'    mnuViewStatusBar.Checked = Not mnuViewStatusBar.Checked
'    sbStatusBar.Visible = mnuViewStatusBar.Checked
'End Sub
'

'
Private Sub tbToolBar_ButtonClick(ByVal Button As MSComCtlLib.Button)
    On Error Resume Next
    Select Case Button.Key
        Case "New"
            LoadNewDoc
'            MsgBox "Add 'New' button code."
        Case "DocDir"
            MsgBox "Add 'DocDir' button code."
        Case "DocDir"
            Call getFolderDir
    End Select
'Lib
'  Case "Open"
'            mnuFileOpen_Click
'        Case "Save"
'            mnuFileSave_Click
'        Case "Print"
'            mnuFilePrint_Click
'        Case "Cut"
'            mnuEditCut_Click
'        Case "Copy"
'            mnuEditCopy_Click
'        Case "Paste"
'            mnuEditPaste_Click
'        Case "Bold"
'            ActiveForm.rtfText.SelBold = Not ActiveForm.rtfText.SelBold
'            Button.Value = IIf(ActiveForm.rtfText.SelBold, tbrPressed, tbrUnpressed)
'        Case "Italic"
'            ActiveForm.rtfText.SelItalic = Not ActiveForm.rtfText.SelItalic
'            Button.Value = IIf(ActiveForm.rtfText.SelItalic, tbrPressed, tbrUnpressed)
'        Case "Underline"
'            ActiveForm.rtfText.SelUnderline = Not ActiveForm.rtfText.SelUnderline
'            Button.Value = IIf(ActiveForm.rtfText.SelUnderline, tbrPressed, tbrUnpressed)
'        Case "Justify"
'            'ToDo: Add 'Justify' button code.
'            MsgBox "Add 'Justify' button code."
'        Case "Align Left"
'            ActiveForm.rtfText.SelAlignment = rtfLeft
'        Case "Center"
'            ActiveForm.rtfText.SelAlignment = rtfCenter
'        Case "Align Right"
'            ActiveForm.rtfText.SelAlignment = rtfRight
End Sub
''
'Private Sub mnuToolsOptions_Click()
'    frmOptions.Show vbModal, Me
'End Sub
'
'
'Private Sub mnuHelpSearchForHelpOn_Click()
'    Dim nRet As Integer
'    'if there is no helpfile for this project display a message to the user
'    'you can set the HelpFile for your application in the
'    'Project Properties dialog
'    If Len(App.HelpFile) = 0 Then
'        MsgBox "������� �������� ����������. There is no Help associated with this project.", vbInformation, Me.Caption
'    Else
'        On Error Resume Next
'        nRet = OSWinHelp(Me.hwnd, App.HelpFile, 261, 0)
'        If Err Then
'            MsgBox Err.Description
'        End If
'    End If
'End Sub
'
'Private Sub mnuHelpContents_Click()
''��������� ������� ������ ���� "�������"
'    Dim nRet As Integer
'    'if there is no helpfile for this project display a message to the user
'    'you can set the HelpFile for your application in the
'    'Project Properties dialog
'    If Len(App.HelpFile) = 0 Then
'        MsgBox "Unable to display Help Contents. There is no Help associated with this project.", vbInformation, Me.Caption
'    Else
'        On Error Resume Next
'        nRet = OSWinHelp(Me.hwnd, App.HelpFile, 3, 0)
'        If Err Then
'            MsgBox Err.Description
'        End If
'    End If
'End Sub
'Private Sub mnuViewToolbar_Click()
'    mnuViewToolbar.Checked = Not mnuViewToolbar.Checked
'    tbToolBar.Visible = mnuViewToolbar.Checked
'End Sub
''
'Private Sub mnuEditPasteSpecial_Click()
'    'ToDo: Add 'mnuEditPasteSpecial_Click' code.
'    MsgBox "Add 'mnuEditPasteSpecial_Click' code."
'End Sub
'
'Private Sub mnuEditPaste_Click()
'    On Error Resume Next
'    ActiveForm.rtfText.SelRTF = Clipboard.GetText
'End Sub
'
'Private Sub mnuEditCopy_Click()
'    On Error Resume Next
'    Clipboard.SetText ActiveForm.rtfText.SelRTF
'End Sub
'
'Private Sub mnuEditCut_Click()
'    On Error Resume Next
'    Clipboard.SetText ActiveForm.rtfText.SelRTF
'    ActiveForm.rtfText.SelText = vbNullString
'End Sub
'
'Private Sub mnuEditUndo_Click()
''��������� ������� "�������"
'    'ToDo: Add 'mnuEditUndo_Click' code.
'    MsgBox "Add 'mnuEditUndo_Click' code."
'End Sub
''
'Private Sub mnuFileExit_Click()
''��������� �������� ������� �����
'    'unload the form
'    Unload Me
'End Sub
'
'Private Sub mnuFileSend_Click()
''��������� ������� "�������� ��������"
'    'ToDo: Add 'mnuFileSend_Click' code.
'    MsgBox "Add 'mnuFileSend_Click' code."
'End Sub
'
'Private Sub mnuFilePrint_Click()
''��������� ������� "������ �����"
'    On Error Resume Next
'    If ActiveForm Is Nothing Then Exit Sub
'     With dlgCommonDialog
'        .DialogTitle = "Print"
'        .CancelError = True
'        .Flags = cdlPDReturnDC + cdlPDNoPageNums
'        If ActiveForm.rtfText.SelLength = 0 Then
'            .Flags = .Flags + cdlPDAllPages
'        Else
'            .Flags = .Flags + cdlPDSelection
'        End If
'        .ShowPrinter
'        If Err <> MSComDlg.cdlCancel Then
'            ActiveForm.rtfText.SelPrint .hDC
'        End If
'    End With
'End Sub
'
'Private Sub mnuFilePrintPreview_Click()
''��������� ������� "��������� ������ �����"
'    'ToDo: Add 'mnuFilePrintPreview_Click' code.
'    MsgBox "Add 'mnuFilePrintPreview_Click' code."
'End Sub
'
'Private Sub mnuFilePageSetup_Click()
''��������� ������� "��������� ��������"
'    On Error Resume Next
'    With dlgCommonDialog
'        .DialogTitle = "Page Setup"
'        .CancelError = True
'        .ShowPrinter
'    End With
'End Sub
''
'Private Sub mnuFileProperties_Click()
''��������� ������� "�������� �����"
'    'ToDo: Add 'mnuFileProperties_Click' code.
'    MsgBox "Add 'mnuFileProperties_Click' code."
'End Sub
''
'Private Sub mnuFileSaveAll_Click()
''��������� ������� "��������� ��� �������� �����"
'    'ToDo: Add 'mnuFileSaveAll_Click' code.
'    MsgBox "Add 'mnuFileSaveAll_Click' code."
'End Sub
'
'Private Sub mnuFileSaveAs_Click()
''��������� ������� "��������� ���� ���"
'    Dim sFile As String
'    If ActiveForm Is Nothing Then Exit Sub
'    With dlgCommonDialog
'        .DialogTitle = "Save As"
'        .CancelError = False
'        'ToDo: set the flags and attributes of the common dialog control
'        .Filter = "All Files (*.*)|*.*"
'        .ShowSave
'        If Len(.FileName) = 0 Then
'            Exit Sub
'        End If
'        sFile = .FileName
'    End With
'    ActiveForm.Caption = sFile
'    ActiveForm.rtfText.SaveFile sFile
'End Sub
''
'Private Sub mnuFileSave_Click()
''��������� ������� "��������� ����"
'    Dim sFile As String
'    If Left$(ActiveForm.Caption, 8) = "Document" Then
'        With dlgCommonDialog
'            .DialogTitle = "Save"
'            .CancelError = False
'            'ToDo: set the flags and attributes of the common dialog control
'            .Filter = "All Files (*.*)|*.*"
'            .ShowSave
'            If Len(.FileName) = 0 Then
'                Exit Sub
'            End If
'            sFile = .FileName
'        End With
'        ActiveForm.rtfText.SaveFile sFile
'    Else
'        sFile = ActiveForm.Caption
'        ActiveForm.rtfText.SaveFile sFile
'    End If
'End Sub
''
'Private Sub mnuFileClose_Click()
''��������� ������� "������� ����"
'    'ToDo: Add 'mnuFileClose_Click' code.
'    MsgBox "Add 'mnuFileClose_Click' code."
'End Sub
'
Private Sub mnuFileOpen_Click()
'��������� ������� "������� ����"
    Dim sFile As String
    If ActiveForm Is Nothing Then LoadNewDoc
    With dlgCommonDialog
        .DialogTitle = "Open"
        .CancelError = False
        'ToDo: set the flags and attributes of the common dialog control
        .Filter = "All Files (*.*)|*.*"
        .ShowOpen
        If Len(.FileName) = 0 Then
            Exit Sub
        End If
        sFile = .FileName
    End With
    ActiveForm.rtfText.LoadFile sFile
    ActiveForm.Caption = sFile
End Sub
'
Private Sub mnuFileNew_Click()
'��������� ������� "������� ����� ��������"
    LoadNewDoc
End Sub
'
'++++++++++++++++++ � � � � � � � � � � � � ++++++++++++++++++
Friend Property Let category(ByVal vData As String)
'��������� ������������ ���������
    tmpcategory = vData
End Property
'
Friend Property Get category() As String
'��������� ������������ ���������
    category = tmpcategory
'Debug.Print "category: " & category
End Property
'
Private Property Let newRoot(ByVal vData As String)
'����� ����������, ��������� �������������
    tmpnewRoot = vData
End Property
'
Friend Property Get newRoot() As String
'����� ����������, ��������� �������������
    newRoot = tmpnewRoot
'Debug.Print "����� ����������, ��������� �������������: " & newRoot
End Property
