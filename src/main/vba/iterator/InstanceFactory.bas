Attribute VB_Name = "InstanceFactory"
Option Compare Database
Option Explicit

'''
'@param strName 書名を指定する。
'@return Book
'''
Public Function NewBook(strName As String) As Book
    Dim objBook As New Book
    objBook.bookName = strName

    Set NewBook = objBook
End Function

'''
'@param intMax 本棚に格納できる本の最大冊数を指定する。
'@return BookShelf
'''
Public Function NewBookShelf(intMax As Integer) As BookShelf
    Dim objBs As New BookShelf
    objBs.setMax intMax

    Set NewBookShelf = objBs
End Function

'''
'@param objBs BookShelf オブジェクトを指定する。
'@return BookShelfIterator
'''
Public Function NewBookShelfIterator(objBs As BookShelf) As BookShelfIterator
    Dim objBsi As New BookShelfIterator
    objBsi.setBookShelf objBs

    Set NewBookShelfIterator = objBsi
End Function
