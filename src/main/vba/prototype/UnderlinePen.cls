VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "UnderlinePen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Implements Product

Private strChar As String

'''
'@param strValue 飾り文字として設定する文字を指定する。
'''
Public Sub setChar(strValue As String)
    strChar = strValue
End Sub

'''
'@return Product 自分自身のコピーを返す。
'''
Private Function Product_createClone() As Product
    Dim objProto As Product
    Set objProto = Me

    Set Product_createClone = objProto
End Function

'''
'文字列に飾り文字を付与して出力する。
'
'@param strValue 出力メッセージとして設定する文字列を指定する。
'''
Private Sub Product_use(strValue As String)
    Dim intLength As Integer
    Dim intIndex As Integer

    intLength = Len(strValue)

    Debug.Print """" & strValue & """"
    Debug.Print " ";

    For intIndex = 1 To intLength
        Debug.Print strChar;
    Next

    Debug.Print ""
End Sub
