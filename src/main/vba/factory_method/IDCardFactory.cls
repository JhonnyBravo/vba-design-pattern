VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IDCardFactory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Implements Factory

Private objOwners As Collection

Private Sub Class_Initialize()
    Set objOwners = New Collection
End Sub

'''
'@param strOwner ID カードの所有者として設定する名前を指定する。
'@return Product ID カードを作成して返す。
'''
Private Function Factory_createProduct(strOwner As String) As Product
    Dim objCard As New IDCard
    objCard.owner = strOwner

    Set Factory_createProduct = objCard
End Function

'''
'ID カードを製品登録する。
'
'@param objProduct 登録対象とする IDCard オブジェクトを指定する。
'''
Private Sub Factory_registerProduct(objProduct As Product)
    Dim objCard As IDCard
    Set objCard = objProduct

    objOwners.Add objCard.owner
End Sub

'''
'ID カードの所有者一覧を返す。
'''
Public Function getOwners() As Collection
    Set getOwners = objOwners
End Function

'''
'@param strOwner ID カードの所有者として設定する名前を指定する。
'@return Product ID カードを作成し、製品登録したうえで ID カードを返す。
'''
Public Function createCard(strOwner As String) As Product
    Dim objSelf As Factory
    Dim objProduct As Product

    Set objSelf = Me

    With objSelf
        Set objProduct = .createProduct(strOwner)
        .registerProduct objProduct
    End With

    Set createCard = objProduct
End Function
