Attribute VB_Name = "Main"
Option Compare Database
Option Explicit

Public Sub run()
    Dim objD1 As display
    Dim objD2 As display
    Dim objD3 As CountDisplay

    Set objD1 = NewDisplay(NewStringDisplayImpl("Hello, Japan."))
    Set objD2 = NewCountDisplay(NewStringDisplayImpl("Hello, World.")).getDisplay
    Set objD3 = NewCountDisplay(NewStringDisplayImpl("Hello, Universe."))

    objD1.display
    objD2.display
    objD3.getDisplay.display
    objD3.multiDisplay 5
End Sub
