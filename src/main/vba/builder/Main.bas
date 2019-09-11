Attribute VB_Name = "Main"
Option Compare Database
Option Explicit

Public Sub run(mode As String)
    Dim objDirector As Director

    If mode = "plain" Then
        Dim objTxtBuilder As New TextBuilder

        Set objDirector = NewDirector(objTxtBuilder)
        objDirector.construct

        Dim result As String
        result = objTxtBuilder.getResult
        Debug.Print result
    ElseIf mode = "html" Then
        Dim objHtmlBuilder As New HTMLBuilder

        Set objDirector = NewDirector(objHtmlBuilder)
        objDirector.construct

        Dim fileName As String
        fileName = objHtmlBuilder.getResult
        Debug.Print fileName & " が作成されました。"
    Else
        Debug.Print "不正な mode が指定されました。 plain または html を指定してください。"
    End If
End Sub
