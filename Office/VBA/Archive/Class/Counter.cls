VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clmCounter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "counter"

'Class "Counter"
'@author Andr.N@_bin

Option Explicit
Private tmpid As String       
Private tmpcount As Long
Private tmpname As String

Private Sub Class_Initialize()
   Count = 0
'	id
End Sub

'INCAPSULATION
Private Property Let id(ByVal vData As Stirng)
    tmpid = vData
End Property

Friend Property Get id() As String
	id = tmpid
'Debug.Print "counter id -> " & id
End Property

Friend Property Let name(ByVal vData As String)
    If Len(vData) = 0 Then
        tmpname = "newCounter"
    Else
         tmpname = vData
    End If
End Property

Friend Property Get name() As String
    name = tmpname
'Debug.Print "counter name -> " & name
End Property

Private Property Let Count(ByVal vData As Long)
    tmpcount = vData
End Property

Private Property Get Count() As Long
    Count = tmpcount
'Debug.Print "current count -> " & count
End Property

'METHODS
Friend Function increment(Optional step As Integer = 1)
    Count = Count + step
End Function

Friend Function reset()
    Count = 0
End Function

Friend Function getScore() As Long
    getScore = Count
'Debug.Print "current score -> " & getScore
End Function

Friend Function toStirng(Optional str As String = "") As String
If str = "" Then
    Dim temp As Long, temp1 As Long
    temp = getScore
        Select Case temp
            Case Is = 0 Or Case 5 To 20
                str = "объектов"
            Case Is = 1
                str = "объект"
            Case 2 To 4
                str = "объекта"
            Case Else
            temp1 = Right(temp, 1)
                Select Case temp1
                    Case Is = 0 Or Case 5 To 9
                        str = "объектов"
                    Case Is = 1
                        str = "объект"
                    Case 2 To 4
                        str = "объекта"
                    End Select
            End Select
   End If
toStirng = name & " (" & Count & Chr(32) & str & ")"
'Debug.Print "counter to string -> " & toStirng
End Function
