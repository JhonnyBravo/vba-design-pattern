VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Banner"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private strMessage As String

'''
'@param strValue メッセージとして設定する文字列を指定する。
'''
Public Sub setMessage(strValue As String)
    strMessage = strValue
End Sub

'''
'下記の形式でメッセージを出力する。
'   (メッセージ)
'''
Public Sub showWithParen()
    Debug.Print "(" & strMessage & ")"
End Sub

'''
'下記の形式でメッセージを出力する。
'   *メッセージ*
'''
Public Sub showWithAster()
    Debug.Print "*" & strMessage & "*"
End Sub
