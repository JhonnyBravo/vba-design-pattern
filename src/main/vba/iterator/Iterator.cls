VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Iterator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Compare Database
Option Explicit

'''
'@return result 次の要素が存在するかどうかを真偽値で返す。
'   True: 次の要素が存在することを表す。
'   False: 次の要素が存在しないことを表す。
'''
Public Function hasNext() As Boolean

End Function

'''
'@return item 集合体から要素を一つ返す。
'''
Public Function nextItem() As Object

End Function
