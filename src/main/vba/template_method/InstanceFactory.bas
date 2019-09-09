Attribute VB_Name = "InstanceFactory"
Option Compare Database
Option Explicit

'''
'@param strMessage メッセージとして設定する文字列を指定する。
'@return CharDisplay
'''
Public Function NewCharDisplay(strMessage As String) As CharDisplay
    Dim objCd As New CharDisplay
    objCd.setMessage strMessage

    Set NewCharDisplay = objCd
End Function

'''
'@param strMessage メッセージとして設定する文字列を指定する。
'@return StringDisplay
'''
Public Function NewStringDisplay(strMessage As String) As StringDisplay
    Dim objSd As New StringDisplay
    objSd.setMessage strMessage

    Set NewStringDisplay = objSd
End Function

'''
'メッセージを出力する。
'@param objDisplay 操作対象とする AbstractDisplay の実装クラスを指定する。
'''
Public Sub display(objDisplay As AbstractDisplay)
    Dim intIndex As Integer

    With objDisplay
        .openMessage

        For intIndex = 1 To 5
            .printMessage
        Next

        .closeMessage
    End With
End Sub
