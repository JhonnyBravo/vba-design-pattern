VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Director"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private objBuilder As Builder

'''
'@param obj 操作対象とする Builder オブジェクトを指定する。
'''
Public Sub setBuilder(ByRef obj As Builder)
    Set objBuilder = obj
End Sub

'''
'処理を実行する。
'''
Public Sub construct()
    With objBuilder
        .makeTitle "Greeting"
        .makeString "朝から昼にかけて"
        .makeItems Array("おはようございます。", "こんにちは。")

        .makeString "夜に"
        .makeItems Array("こんばんは。", "おやすみなさい。", "さようなら。")
        .closeBuilder
    End With
End Sub
