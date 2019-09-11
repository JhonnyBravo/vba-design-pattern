VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "HTMLBuilder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Implements Builder

Private fileName As String

'''
'@see Builder#closeBuilder
'''
Private Sub Builder_closeBuilder()
    Debug.Print "</body></html>"
End Sub

'''
'@see Builder#makeItems
'''
Private Sub Builder_makeItems(items As Variant)
    Dim intIndex As Integer

    Debug.Print "<ul>"

    For intIndex = 0 To UBound(items)
        Debug.Print "<li>" & items(intIndex) & "</li>"
    Next

    Debug.Print "</ul>"
End Sub

'''
'@see Builder#makeString
'''
Private Sub Builder_makeString(str As String)
    Debug.Print "<p>" & str & "</p>"
End Sub

'''
'@see Builder#makeTitle
'''
Private Sub Builder_makeTitle(title As String)
    fileName = title & ".html"
    Debug.Print fileName
    Debug.Print "<html><head><title>" & title & "</title></head><body>"
    Debug.Print "<h1>" & title & "</h1>"
End Sub

'''
'@return fileName HTML ファイル名を返す。
'''
Public Function getResult() As String
    getResult = fileName
End Function
