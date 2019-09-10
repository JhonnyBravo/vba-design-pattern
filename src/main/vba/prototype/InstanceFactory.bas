Attribute VB_Name = "InstanceFactory"
Option Compare Database
Option Explicit

'''
'@param strChar 飾り文字として設定する文字を指定する。
'@return MessageBox
'''
Public Function NewMessageBox(strChar As String) As MessageBox
    Dim objMb As New MessageBox
    objMb.setChar strChar

    Set NewMessageBox = objMb
End Function

'''
'@param strChar 飾り文字として設定する文字を指定する。
'@return UnderlinePen
'''
Public Function NewUnderlinePen(strChar As String) As UnderlinePen
    Dim objUp As New UnderlinePen
    objUp.setChar strChar

    Set NewUnderlinePen = objUp
End Function
