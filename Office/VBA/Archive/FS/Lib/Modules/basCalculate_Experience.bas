Private Static Function Calculate_Experience() As String
'������� ������� ����� ������
'� ������� �������� Text1.Text �� ���� ������ �������� ������������
Dim bytEx1 As Byte, strEx2 As String
strEx2 = "01.08.2005"
    bytEx1 = DateDiff("yyyy", CDate(strEx2), Now) '������ ����� ������
    Select Case bytEx1
        Case Is = 1
            Calculate_Experience = bytEx1 & " ���,"
        Case 2 To 4
            Calculate_Experience = bytEx1 & " ����,"
        Case 5 To 20
            Calculate_Experience = bytEx1 & " ���,"
          Case Else
            strEx2 = Right(bytEx1, 1)
                Select Case CByte(strEx2)
                    Case Is = 0
                        Calculate_Experience = bytEx1 & " ���,"
                    Case Is = 1
                        Calculate_Experience = bytEx1 & " ���,"
                    Case 2 To 4
                        Calculate_Experience = bytEx1 & " ����,"
                    Case 5 To 9
                        Calculate_Experience = bytEx1 & " ���,"
                End Select
    End Select
End Function
