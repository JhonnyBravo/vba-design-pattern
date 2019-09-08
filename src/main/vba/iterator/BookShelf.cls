VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "BookShelf"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Implements Aggregate

Private objBooks As Collection
Private intLast As Integer
Private intMax As Integer

Private Sub Class_Initialize()
    Set objBooks = New Collection
    intLast = 1
End Sub

'''
'@param intValue 本棚に格納できる本の最大冊数を指定する。
'''
Public Sub setMax(intValue As Integer)
    intMax = intValue
End Sub

'''
'@return Iterator
'''
Private Function Aggregate_getIterator() As Iterator
    Set Aggregate_getIterator = NewBookShelfIterator(Me)
End Function

'''
'@param intIndex 本棚から取得したい本のインデックス番号を指定する。
'@return Book
'''
Public Function getBookAt(intIndex As Integer) As Book
    Set getBookAt = objBooks(intIndex)
End Function

'''
'@param objBook 本棚へ格納する本を指定する。
'''
Public Sub appendBook(objBook As Book)
    If intLast > intMax Then
        Err.Raise vbObjectError + 1, "BookShelf.getBookAt", "本棚が満杯です。"
    End If

    objBooks.Add objBook
    intLast = intLast + 1
End Sub

'''
'@return length 現在本棚に格納されている本の冊数を返す。
'''
Public Function getLength() As Integer
    getLength = intLast
End Function
