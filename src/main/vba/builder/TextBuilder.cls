VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TextBuilder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Implements Builder

Private objBuffer As Collection
Private strDeco As String

Private Sub Class_Initialize()
    Set objBuffer = New Collection
End Sub

'''
'@see Builder#closeBuilder
'''
Private Sub Builder_closeBuilder()
    objBuffer.Add strDeco & vbCrLf
End Sub

'''
'@see Builder#makeItems
'''
Private Sub Builder_makeItems(items As Variant)
    Dim intIndex As Integer

    For intIndex = 0 To UBound(items)
        objBuffer.Add " * " & items(intIndex) & vbCrLf
    Next

    objBuffer.Add vbCrLf
End Sub

'''
'@see Builder#makeString
'''
Private Sub Builder_makeString(str As String)
    With objBuffer
        .Add "■" & str & vbCrLf
        .Add vbCrLf
    End With
End Sub

'''
'@see Builder#makeTitle
'''
Private Sub Builder_makeTitle(title As String)
    Dim intIndex As Integer

    For intIndex = 1 To Len(title) + 4
        strDeco = strDeco & "="
    Next

    With objBuffer
        .Add strDeco & vbCrLf
        .Add "『" & title & "』" & vbCrLf
        .Add vbCrLf
    End With
End Sub

'''
'@return result Buffer コレクションに格納された文字列をすべて連結して返す。
'''
Public Function getResult() As String
    Dim varItem As Variant
    Dim strResult As String

    For Each varItem In objBuffer
        strResult = strResult & varItem
    Next

    getResult = strResult
End Function
