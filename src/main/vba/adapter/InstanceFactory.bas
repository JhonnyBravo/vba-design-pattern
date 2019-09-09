Attribute VB_Name = "InstanceFactory"
Option Compare Database
Option Explicit

'''
'@param strMessage メッセージとして設定する文字列を指定する。
'@return Banner
'''
Public Function NewBanner(strMessage As String) As Banner
    Dim objBanner As New Banner
    objBanner.setMessage strMessage

    Set NewBanner = objBanner
End Function

'''
'@param strMessage メッセージとして設定する文字列を指定する。
'@return PrintBanner
'''
Public Function NewPrintBanner(strMessage As String) As PrintBanner
    Dim objPb As New PrintBanner
    objPb.setBanner NewBanner(strMessage)

    Set NewPrintBanner = objPb
End Function
