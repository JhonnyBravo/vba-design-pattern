VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Manager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private objDict As Dictionary

Private Sub Class_Initialize()
    Set objDict = New Dictionary
End Sub

'''
'Product オブジェクトを格納する。
'
'@param strName 登録するオブジェクトの名前を指定する。
'@param objProto 登録対象とする Product オブジェクトを指定する。
'''
Public Sub register(strName As String, objProto As Product)
    objDict.Add strName, objProto
End Sub

'''
'@param strName 取得対象とする Product オブジェクトの名前を指定する。
'@return Product 該当する名前に紐づいて格納されている Product オブジェクトを返す。
'''
Public Function createProduct(strName As String) As Product
    Dim objProto As Product
    Set objProto = objDict.Item(strName)

    Set createProduct = objProto.createClone
End Function
