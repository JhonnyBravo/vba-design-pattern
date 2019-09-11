VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Display"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private objImpl As DisplayImpl

'''
'@param impl 操作対象とする DisplayImpl オブジェクトを指定する。
'''
Public Sub setImpl(ByRef impl As DisplayImpl)
    Set objImpl = impl
End Sub

'''
'@see DisplayImpl#rawOpen
'''
Public Sub openDisplay()
    objImpl.rawOpen
End Sub

'''
'@see DisplayImpl#rawPrint
'''
Public Sub printDisplay()
    objImpl.rawPrint
End Sub

'''
'@see DisplayImpl#rawClose
'''
Public Sub closeDisplay()
    objImpl.rawClose
End Sub

'''
'メッセージに飾り枠を付与して出力する。
'''
Public Sub display()
    openDisplay
    printDisplay
    closeDisplay
End Sub
