Attribute VB_Name = "InstanceFactory"
Option Compare Database
Option Explicit

'''
'@param objBuilder 操作対象とする Builder オブジェクトを指定する。
'@return Director
'''
Public Function NewDirector(ByRef objBuilder As Builder) As Director
    Dim objDirector As New Director
    objDirector.setBuilder objBuilder

    Set NewDirector = objDirector
End Function
