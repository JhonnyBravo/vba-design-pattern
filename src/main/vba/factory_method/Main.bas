Attribute VB_Name = "Main"
Option Compare Database
Option Explicit

Public Sub run()
    Dim objFactory As New IDCardFactory

    Dim objCard1 As Product
    Dim objCard2 As Product
    Dim objCard3 As Product

    With objFactory
        Set objCard1 = .createCard("結城浩")
        Set objCard2 = .createCard("とむら")
        Set objCard3 = .createCard("佐藤花子")
    End With

    objCard1.use
    objCard2.use
    objCard3.use
End Sub
