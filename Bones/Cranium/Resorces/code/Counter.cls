VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Counter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@author Andr.Nab.n@gmail.com
'@01/01/2016
Option Explicit
'
Private mvarcount As Integer
Private mvarname As String

Private Property Let count(vData As Integer)
    mvarcount = vData
End Property
'
Public Property Get count() As Integer
    count = mvarcount
End Property
'
Public Property Let name(vData As String)
    mvarname = vData
End Property
'
Public Property Get name() As String
    name = mvarname
End Property
'
Public Static Sub increment()
    count = count + 1
End Sub
'
Public Static Function toString() As String
    toString = name & " - " & count
End Function

Private Sub Class_Initialize()
    count = 0
End Sub
