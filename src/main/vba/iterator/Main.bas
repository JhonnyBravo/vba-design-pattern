Attribute VB_Name = "Main"
Option Compare Database
Option Explicit

Public Sub run()
    Dim objBs As BookShelf
    Set objBs = NewBookShelf(4)

    With objBs
        .appendBook NewBook("Around the World in 80 Days")
        .appendBook NewBook("Bible")
        .appendBook NewBook("Cinderella")
        .appendBook NewBook("Daddy-Long-Legs")
    End With

    Dim objAggregate As Aggregate
    Dim objIterator As Iterator
    Dim objBook As Book

    Set objAggregate = objBs
    Set objIterator = objAggregate.getIterator

    While objIterator.hasNext
        Set objBook = objIterator.nextItem
        Debug.Print objBook.bookName
    Wend
End Sub
