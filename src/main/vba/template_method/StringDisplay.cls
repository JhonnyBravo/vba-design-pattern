VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "StringDisplay"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Implements AbstractDisplay

Private strMessage As String
Private intWidth As Integer

'''
'@param strValue メッセージとして設定する文字列を指定する。
'''
Public Sub setMessage(strValue As String)
    strMessage = strValue
    intWidth = Len(strValue)
End Sub

'''
'修飾文字列を出力する。
'''
Private Sub printLine()
    Debug.Print "+";

    Dim intIndex As Integer

    For intIndex = 1 To intWidth
        Debug.Print "-";
    Next

    Debug.Print "+"
End Sub

'''
'@see AbstractDisplay#closeMessage
'''
Private Sub AbstractDisplay_closeMessage()
    printLine
End Sub

'''
'@see AbstractDisplay#openMessage
'''
Private Sub AbstractDisplay_openMessage()
    printLine
End Sub

'''
'@see AbstractDisplay#printMessage
'''
Private Sub AbstractDisplay_printMessage()
    Debug.Print "|" & strMessage & "|"
End Sub
