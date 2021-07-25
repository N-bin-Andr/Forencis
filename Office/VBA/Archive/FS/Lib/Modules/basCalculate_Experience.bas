Private Static Function Calculate_Experience() As String
'‘ункци€ расчета стажа работы
'¬ проекте заменить Text1.Text на дату начала трудовой де€тельности
Dim bytEx1 As Byte, strEx2 As String
strEx2 = "01.08.2005"
    bytEx1 = DateDiff("yyyy", CDate(strEx2), Now) 'расчет стажа работы
    Select Case bytEx1
        Case Is = 1
            Calculate_Experience = bytEx1 & " год,"
        Case 2 To 4
            Calculate_Experience = bytEx1 & " года,"
        Case 5 To 20
            Calculate_Experience = bytEx1 & " лет,"
          Case Else
            strEx2 = Right(bytEx1, 1)
                Select Case CByte(strEx2)
                    Case Is = 0
                        Calculate_Experience = bytEx1 & " лет,"
                    Case Is = 1
                        Calculate_Experience = bytEx1 & " год,"
                    Case 2 To 4
                        Calculate_Experience = bytEx1 & " года,"
                    Case 5 To 9
                        Calculate_Experience = bytEx1 & " лет,"
                End Select
    End Select
End Function
