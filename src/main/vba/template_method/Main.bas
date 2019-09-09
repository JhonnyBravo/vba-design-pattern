Attribute VB_Name = "Main"
Option Compare Database
Option Explicit

Public Sub run()
    Dim d1 As AbstractDisplay
    Set d1 = NewCharDisplay("H")

    Dim d2 As AbstractDisplay
    Set d2 = NewStringDisplay("Hello, world.")

    Dim d3 As AbstractDisplay
    Set d3 = NewStringDisplay("こんにちは。")

    display d1
    display d2
    display d3
End Sub
