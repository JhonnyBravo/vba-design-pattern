Attribute VB_Name = "InstanceFactory"
Option Compare Database
Option Explicit

'''
'@param objImpl 操作対象とする DisplayImpl オブジェクトを指定する。
'@return Display
'''
Public Function NewDisplay(ByRef objImpl As DisplayImpl) As display
    Dim objDisplay As New display
    objDisplay.setImpl objImpl

    Set NewDisplay = objDisplay
End Function

'''
'@param objImpl 操作対象とする DisplayImpl オブジェクトを指定する。
'@return CountDisplay
'''
Public Function NewCountDisplay(ByRef objImpl As DisplayImpl) As CountDisplay
    Dim objCd As New CountDisplay
    objCd.setDisplay NewDisplay(objImpl)

    Set NewCountDisplay = objCd
End Function

'''
'@param strMessage 出力メッセージとして設定する文字列を指定する。
'@return StringDisplayImpl
'''
Public Function NewStringDisplayImpl(strMessage As String) As StringDisplayImpl
    Dim objSdi As New StringDisplayImpl
    objSdi.setMessage strMessage

    Set NewStringDisplayImpl = objSdi
End Function
