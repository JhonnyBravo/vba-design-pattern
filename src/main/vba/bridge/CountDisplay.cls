VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CountDisplay"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private objSuper As display

'''
'@param objDisplay 継承の対象とする Display オブジェクトを指定する。
'''
Public Sub setDisplay(ByRef objDisplay As display)
    Set objSuper = objDisplay
End Sub

'''
'@return Display 継承した Display オブジェクトを呼び出す。
'''
Public Function getDisplay() As display
    Set getDisplay = objSuper
End Function

'''
'飾り枠付きメッセージを繰り返し出力する。
'
'@param intTimes 繰返し回数を指定する。
'''
Public Sub multiDisplay(intTimes As Integer)
    Dim intIndex As Integer

    With objSuper
        .openDisplay

        For intIndex = 1 To intTimes
            .printDisplay
        Next

        .closeDisplay
    End With
End Sub
