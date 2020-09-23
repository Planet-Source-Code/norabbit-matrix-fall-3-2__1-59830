VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsCoordonnees"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'******************************************************************
'* CLASSE PERMETTANT L'ARCHIVAGE D'INFORMATIONS SUIVANT LE MODELE *
'* D'UNE LISTE.                                                   *
'*                                                                *
'* par Thomas John                                                *
'******************************************************************
'
'index des classes
Private IndexSz() As Integer

Private Sub Class_Initialize()
    '
    ReDim v(0 To 0)
    '
End Sub
'
Public Sub Ajout(Index As Integer)
    '
    On Error Resume Next
    '
    ReDim Preserve IndexSz(LBound(IndexSz) To UBound(IndexSz) + 1)
    '
    IndexSz(UBound(IndexSz)) = Index
    '
End Sub
'
Public Sub Enlever(Index As Integer)
    '
    On Error Resume Next
    '
    Dim IndexSz2 As Integer
    '
    If Index < 1 Then Exit Sub
    If Index > UBound(IndexSz) Then Exit Sub
    '
    For IndexSz2 = Index To UBound(IndexSz) - 1
        '
        IndexSz(IndexSz2) = IndexSz(IndexSz2 + 1)
        '
    Next
    '
    ReDim Preserve IndexSz(LBound(IndexSz) To UBound(IndexSz) - 1)
    '
End Sub
'
Public Sub ToutEffacer()
    '
    ReDim IndexSz(0 To 0)
    '
End Sub
'
Public Function TotalElements() As Long
    '
    TotalElements = UBound(IndexSz)
    '
End Function
'
Public Function ElementIndex(Index As Integer) As Integer
    '
    On Error Resume Next
    '
    If Index <= 0 Then Exit Function
    If Index > UBound(IndexSz) Then Exit Function
    '
    ElementIndex = IndexSz(Index)
    '
End Function
'
Public Sub ChangerIndex(Index As Integer, IndexT As Integer)
    '
    On Error Resume Next
    '
    If Index <= 0 Then Exit Sub
    If Index > UBound(IndexSz) Then Exit Sub
    '
    IndexSz(Index) = IndexT
    '
End Sub



