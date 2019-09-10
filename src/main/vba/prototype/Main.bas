Attribute VB_Name = "Main"
Option Compare Database
Option Explicit

Public Sub run()
    Dim objUpen As UnderlinePen
    Dim objMBox As MessageBox
    Dim objSBox As MessageBox

    '飾り文字の設定とオブジェクトの生成。
    Set objUpen = NewUnderlinePen("~")
    Set objMBox = NewMessageBox("*")
    Set objSBox = NewMessageBox("/")

    'オブジェクトを Manager へ格納する。
    Dim objManager As New Manager

    With objManager
        .register "strong message", objUpen
        .register "warning box", objMBox
        .register "slash box", objSBox
    End With

    'Manager に格納されているオブジェクトの取り出し。
    Dim objP1 As Product
    Dim objP2 As Product
    Dim objP3 As Product

    With objManager
        Set objP1 = .createProduct("strong message")
        Set objP2 = .createProduct("warning box")
        Set objP3 = .createProduct("slash box")
    End With

    '実処理の実行。
    objP1.use "Hello, world."
    objP2.use "Hello, world."
    objP3.use "Hello, world."
End Sub
