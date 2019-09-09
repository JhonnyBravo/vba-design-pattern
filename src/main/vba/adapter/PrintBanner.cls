VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PrintBanner"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Implements IPrint

Private objBanner As Banner

'''
'@param objInstance 操作対象とする Banner オブジェクトを指定する。
'''
Public Sub setBanner(objInstance As Banner)
    Set objBanner = objInstance
End Sub

'''
'@see IPrint#printStrong
'''
Private Sub IPrint_printStrong()
    objBanner.showWithAster
End Sub

'''
'@see IPrint#printWeak
'''
Private Sub IPrint_printWeak()
    objBanner.showWithParen
End Sub
