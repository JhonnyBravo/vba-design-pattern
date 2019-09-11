VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "StringDisplayImpl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Implements DisplayImpl

Private strMessage As String
Private intWidth As Integer

'''
'@param message 出力メッセージとして設定する文字列を指定する。
'''
Public Sub setMessage(message As String)
    strMessage = message
    intWidth = Len(strMessage)
End Sub

'''
'飾り枠を出力する。
'''
Private Sub printLine()
    Dim intIndex As Integer

    Debug.Print "+";

    For intIndex = 1 To intWidth
        Debug.Print "-";
    Next

    Debug.Print "+"
End Sub

'''
'飾り枠(終了枠)を出力する。
'''
Private Sub DisplayImpl_rawClose()
    printLine
End Sub

'''
'飾り枠(開始枠)を出力する。
'''
Private Sub DisplayImpl_rawOpen()
    printLine
End Sub

'''
'メッセージ本体を出力する。
'''
Private Sub DisplayImpl_rawPrint()
    Debug.Print "|" & strMessage & "|"
End Sub
