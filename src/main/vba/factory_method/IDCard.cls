VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IDCard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Implements Product

Private propOwner As String

'''
'@return owner カードの所有者名を返す。
'''
Public Property Get owner() As String
    owner = propOwner
End Property

'''
'@param strOwner カードの所有者として設定する名前を指定する。
'''
Public Property Let owner(strOwner As String)
    Debug.Print strOwner & " のカードを作ります。"
    propOwner = strOwner
End Property

'''
'カードを使用する。
'''
Private Sub Product_use()
    Debug.Print owner & " のカードを使います。"
End Sub
