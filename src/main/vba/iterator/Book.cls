VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Book"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private propName As String

'''
'@return bookName 書名を返す。
'''
Public Property Get bookName() As String
    bookName = propName
End Property

'''
'@param strName 書名を指定する。
'''
Public Property Let bookName(strName As String)
    propName = strName
End Property
