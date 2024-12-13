VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassLote"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Option Base 1

Const MAXCOLS = 3

Private QtdElementos
Private ArrayLote() As Variant
Public Function AddNew(ByVal C�digo As String, ByVal M�ltiplo As Single, ByVal Quantidade As Single)
    QtdElementos = QtdElementos + 1
    ReDim Preserve ArrayLote(MAXCOLS, QtdElementos)
    ArrayLote(1, QtdElementos) = C�digo
    ArrayLote(2, QtdElementos) = M�ltiplo
    ArrayLote(3, QtdElementos) = Quantidade
End Function
Public Function Ascan(ByVal C�digo As String) As Integer
    Dim Cont As Integer
    
    Ascan = 0
    
    For Cont = 1 To QtdElementos
        If ArrayLote(1, Cont) = C�digo Then
            Ascan = Cont
            Exit For
        End If
    Next
End Function
Public Function Count() As Integer
    Count = QtdElementos
End Function
Public Function GetC�digoDoLote(ByVal Elemento As Integer) As String
    GetC�digoDoLote = ArrayLote(1, Elemento)
End Function
Public Function GetM�ltiplo(ByVal Elemento As Integer) As Single
    GetM�ltiplo = ArrayLote(2, Elemento)
End Function
Public Function GetQuantidade(ByVal Elemento As Integer) As Single
    GetQuantidade = ArrayLote(3, Elemento)
End Function
Public Sub RemoveItem(ByVal Elemento As Integer)
    Dim Cont As Integer
    
    If QtdElementos < 1 Then
        Exit Sub
    End If
    
    For Cont = Elemento To UBound(ArrayLote, 2) - 1
        ArrayLote(1, Cont) = ArrayLote(1, Cont + 1)
        ArrayLote(2, Cont) = ArrayLote(2, Cont + 1)
        ArrayLote(3, Cont) = ArrayLote(3, Cont + 1)
    Next
    
    QtdElementos = QtdElementos - 1
    ReDim Preserve ArrayLote(MAXCOLS, IIf(QtdElementos = 0, 1, QtdElementos))
End Sub
Public Sub RemoveAllItens()
    QtdElementos = 0
    ReDim ArrayLote(MAXCOLS, 1)
End Sub
Public Sub SetQuantidade(ByVal Elemento As Integer, ByVal Quantidade As Single)
    ArrayLote(3, Elemento) = Quantidade
End Sub
Private Sub Class_Initialize()
    QtdElementos = 0
    ReDim ArrayLote(MAXCOLS, 1)
End Sub
Private Sub Class_Terminate()
    ReDim ArrayLote(MAXCOLS, 1)
    QtdElementos = 0
End Sub
