Attribute VB_Name = "mdSklonenie"
Option Explicit

  Private fi As String
    Private im As String
    Private ot As String
    Private pol As String
    
    Private fi_s As String
    Private im_s As String
    Private ot_s As String
    Private fi_str As String
    Private im_str As String
    Private ot_str As String
'
Function FIO(fi As String, im As String, ot As String, pol As String) As String
 'Склонение в родительном падеже (declension)
    fi = StrConv(Trim(fi), vbLowerCase)
    im = StrConv(Trim(im), vbLowerCase)
    ot = StrConv(Trim(ot), vbLowerCase)
    pol = pol
'
    fi_s = Len(fi)
    im_s = Len(im)
    ot_s = Len(ot)
    
    Dim i As Integer, k As String
'
    If pol = "ж" Then
        For i = fi_s To fi_s - (fi_s + 2) Step -1
            If fi_s = 0 Then GoTo fi
                k = Mid(fi, i, 1)
            If k = "а" Or k = "у" Or k = "е" Or k = "э" Or k = "о" Or _
                k = "я" Or k = "и" Or k = "ю" Then
                k = Mid(fi, i - 1, 1)
            If k = "а" Or k = "у" Or k = "е" Or k = "э" Or k = "о" Or _
                k = "я" Or k = "и" Or k = "ю" Then
            Else
              k = Mid(fi, i - 2, 1)
                If k = "а" Or k = "у" Or k = "е" Or k = "э" Or k = "о" Or _
                    k = "я" Or k = "и" Or k = "ю" Then
                    fi_str = Left(fi, i - 1) & "ой"
                Exit For
                Else
                    fi_str = fi
                Exit For
                End If
            End If
        Else
          fi_str = fi
          Exit For
        End If
      Next
fi:
      For i = im_s To im_s - (im_s + 2) Step -1
         If im_s = 0 Then GoTo im
         k = Mid(im, i, 1)
        'MsgBox k
        If k = "а" Or k = "у" Or k = "е" Or k = "э" Or k = "о" Or _
          k = "я" Or k = "и" Or k = "ю" Then
          'If i > 1 Then
            k = Mid(im, i - 1, 1)
            If k = "а" Or k = "у" Or k = "е" Or k = "э" Or k = "о" Or _
                k = "я" Or k = "и" Or k = "ю" Then
              im_str = Left(im, i - 2) & "ии"
              Exit For
              
            Else
              im_str = Left(im, i - 1) & "е"
              Exit For
            End If
          'End If
        ElseIf k = "ь" Then
          im_str = Left(im, i - 1) & "и"
          Exit For
        Else
          im_str = im
          Exit For
        End If
      Next
im:
      If ot_s <> 0 Then
        ot_str = Left(ot, Len(ot) - 1) & "е"
      End If
    ElseIf pol = "м" Then
      'MsgBox "м - пол не готов"
        
      For i = fi_s To fi_s - (fi_s + 2) Step -1
      If fi_s = 0 Then GoTo fi1
      k = Mid(fi, i, 1)
        If k = "а" Or k = "у" Or k = "е" Or k = "э" Or k = "о" Or _
          k = "я" Or k = "и" Or k = "ю" Then
          fi_str = fi
          Exit For
        Else
          fi_str = fi & "у"
          Exit For
            
        End If
      Next
fi1:
      For i = im_s To im_s - (im_s + 2) Step -1
      If im_s = 0 Then GoTo im1
        k = Mid(im, i, 1)
        If k = "й" Or k = "ь" Then
          im_str = Left(im, Len(im) - 1) & "ю"
          Exit For
        Else
          im_str = im & "у"
          Exit For
        End If
      Next
im1:
      If ot_s <> 0 Then
        ot_str = ot & "у"
      End If
    End If
    FIO = StrConv(fi_str, vbProperCase) & " " & StrConv(im_str, vbProperCase) & " " & StrConv(ot_str, vbProperCase)
    End Function
'
Private Sub cmdOK_Click()
    lblResult.Caption = FIO(fi, im, ot, pol)
End Sub

