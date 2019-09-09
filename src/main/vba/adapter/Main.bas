Attribute VB_Name = "Main"
Option Compare Database
Option Explicit

Public Sub run()
    Dim objPrint As IPrint
    Set objPrint = NewPrintBanner("Hello")

    With objPrint
        .printWeak
        .printStrong
    End With
End Sub
