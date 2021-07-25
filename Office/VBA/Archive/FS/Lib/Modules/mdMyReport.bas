Attribute VB_Name = "mdMyReport"
Option Explicit
Public myForm As frmMyReport

Public Sub printMyReport()
'расчет и сохранение данных о нагрузке
    Set myForm = New frmMyReport
        myForm.Show
    Set myForm = Nothing
End Sub
