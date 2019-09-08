VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "BookShelfIterator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Implements Iterator

Private objBs As BookShelf
Private intIndex As Integer

Private Sub Class_Initialize()
    intIndex = 1
End Sub

'''
'@param objBookShelf 操作対象とする BookShelf オブジェクトを指定する。
'''
Public Sub setBookShelf(objBookShelf As BookShelf)
    Set objBs = objBookShelf
End Sub

'''
'@return result 本棚に次の本が存在するかどうかを真偽値で返す。
'   True: 本棚に本が存在することを表す。
'   False: 本棚に本が存在しないことを表す。
'''
Private Function Iterator_hasNext() As Boolean
    Dim boolResult As Boolean

    If intIndex < objBs.getLength Then
        boolResult = True
    Else
        boolResult = False
    End If

    Iterator_hasNext = boolResult
End Function

'''
'@return book 本棚から本を取り出す。
'''
Private Function Iterator_nextItem() As Object
    Dim objBook As Object
    Set objBook = objBs.getBookAt(intIndex)
    intIndex = intIndex + 1

    Set Iterator_nextItem = objBook
End Function
