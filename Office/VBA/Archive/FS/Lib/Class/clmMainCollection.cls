VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clmCollection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"Collection1"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
' ласс "MainCollection"
'ќпределение пола и расчет прижизненного роста человека по данной кости.
'@author
Option Explicit
Private mCol As Collection 'создание коллекции
'
'инкапсул€ци€:
Public Property Get Item(vntIndexKey As Variant) As Object
'используетс€ при ссылке на элемент в коллекции.
'vntIndexKey содержит либо индекс, либо ключ к коллекции,
'Ђпоэтому он объ€влен как вариант
'Syntax: Set foo = x.Item(xyz) or Set foo = x.Item(5)
  Set Item = mCol(vntIndexKey)
End Property
'
Public Property Get Count() As Long
'используетс€ дл€ извлечени€ числа элементов в коллекции
'Syntax: Debug.Print x.Count
    Count = mCol.Count
End Property
'
Public Property Get NewEnum() As IUnknown
'это свойство позвол€ет обрабатывать коллекцию с помощью цикла For...Each
    Set NewEnum = mCol.[_NewEnum]
End Property
'
Public Sub Remove(vntIndexKey As Variant)
'используетс€ при удалении элемента из коллекции
'vntIndexKey содержит либо Index, либо Key, поэтому
'Ђон объ€влен как вариант
'Syntax: x.Remove(xyz)
    mCol.Remove vntIndexKey
End Sub
'
Public Function keyExists(sKey As String) As Boolean
'функци€ провер€ет, существует ли ключ
On Error GoTo EH
     mCol.Item key
    keyExists = True
EH: MsgBox " люч не существует!", vbCritical, "ќшибка!"
End Function
'
Public Sub quickSort(first As Long, last As Long)
'быстра€ сортировка коллекции
'пример использовани€:
'quickSort  mCol, 1,  mCol.Count
'   Dim v As Variant
'   For Each  v In  mCol
'       Debug.Print v
'   Next
Dim vCentreVal As Variant
Dim vTemp As Variant
Dim ITempLow As Long
Dim ITempHi As Long
    ITempLow = first
    ITempHi = last
    vCentreVal = mCol(first + last / 2)
    Do While ITempLow <= ITempHi
        Do While mCol(ITempLow) < vCentreVal And ITempLow < last
            ITempLow = ITempLow + 1
        Loop
        Do While vCentreVal < mCol(ITempHi) And ITempHi > first
          ITempHi = ITempHi - 1
        Loop
        If ITempLow <= ITempHi Then
          'помен€ть значени€
            vTemp = mCol(ITempLow)
             mCol.Add mCol(ITempHi), After:=ITempLow
             mCol.Revove ITempLow
             mCol.Add vTemp, Before:=ITempHi
             mCol.Remove ITempHi + 1
'           перейти к следующим  позици€м
            ITempLow = ITempLow + 1
            ITempHi = ITempHi - 1
        End If
    Loop
'
    If first < ITempHi Then
        quickSort mCol, first, ITempHi
        If ITempLow < last Then
            quickSort mCol, ITempLow, last
        End If
    End If
End Sub
'
Public Function addNewClass(cls As Object, sKey As Variant) As Object
''создает новый объект класса (cls) и добавл€ет его в коллекцию
    Dim objNewMember As Object
    Set objNewMember = New cls
'установливает свойства, переданные в метод (set the properties passed into the method)
    If Len(sKey) = 0 Then
        mCol.Add objNewMember
    Else
        mCol.Add objNewMember, sKey
    End If
'возвращает созданный объект (return the object created)
    Set Add = objNewMember
    Set objNewMember = Nothing
End Function


''ћетоды (процедуры/функции):
'Public Function Add(Optional sKey As String) As Object
''создает новый объект
'    Dim objNewMember As '!!!исправить
'    Set objNewMember = New '!!!исправить
''установливает свойства, переданные в метод (set the properties passed into the method)
'    If Len(sKey) = 0 Then
'        mCol.Add objNewMember
'    Else
'        mCol.Add objNewMember, sKey
'    End If
''возвращает созданный объект (return the object created)
'    Set Add = objNewMember
'    Set objNewMember = Nothing
'End Function



                    

