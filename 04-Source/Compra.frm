VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form frmCompra 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Compra"
   ClientHeight    =   5535
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9540
   Icon            =   "Compra.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5535
   ScaleWidth      =   9540
   Begin VB.Frame frDadosCadastrais 
      Height          =   1140
      Left            =   0
      TabIndex        =   16
      Top             =   0
      Width           =   9540
      Begin VB.TextBox txtFornecedor 
         Height          =   300
         Left            =   1230
         Locked          =   -1  'True
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   690
         Width           =   4905
      End
      Begin VB.TextBox txtDataDaNotaFiscal 
         Height          =   285
         Left            =   8100
         TabIndex        =   1
         Text            =   "  /  /"
         Top             =   300
         Width           =   990
      End
      Begin VB.TextBox txtNotaFiscal 
         Height          =   285
         Left            =   1230
         TabIndex        =   0
         Top             =   300
         Width           =   2475
      End
      Begin VB.CommandButton cmdTabelaCliente 
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   6150
         TabIndex        =   3
         Top             =   660
         Width           =   375
      End
      Begin VB.Label lblFornecedor 
         Caption         =   "Fornecedor"
         Height          =   180
         Left            =   150
         TabIndex        =   19
         Top             =   720
         Width           =   885
      End
      Begin VB.Label lblData 
         Caption         =   "Data"
         Height          =   210
         Left            =   7530
         TabIndex        =   18
         Top             =   330
         Width           =   465
      End
      Begin VB.Label lblNotaFiscal 
         Caption         =   "Nota Fiscal"
         Height          =   180
         Left            =   150
         TabIndex        =   17
         Top             =   330
         Width           =   825
      End
   End
   Begin VB.Frame frItens 
      Caption         =   " Itens "
      Height          =   2595
      Left            =   0
      TabIndex        =   15
      Top             =   1140
      Width           =   9540
      Begin MSDBGrid.DBGrid dbgrdItens 
         Height          =   2325
         Left            =   60
         OleObjectBlob   =   "Compra.frx":030A
         TabIndex        =   4
         Top             =   210
         Width           =   9405
      End
   End
   Begin VB.CommandButton cmdGravar 
      Caption         =   "&Gravar"
      Height          =   345
      Left            =   6960
      TabIndex        =   7
      Top             =   5175
      Width           =   1245
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   345
      Left            =   8280
      TabIndex        =   8
      Top             =   5175
      Width           =   1245
   End
   Begin VB.CommandButton cmdFormaDePagamento 
      Caption         =   "&Forma de Pagemanto"
      Height          =   345
      Left            =   45
      TabIndex        =   6
      Top             =   5175
      Width           =   1980
   End
   Begin VB.Frame frTotais 
      Caption         =   "Totais"
      Height          =   1365
      Left            =   0
      TabIndex        =   9
      Top             =   3750
      Width           =   9525
      Begin VB.TextBox txtValor 
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   7740
         Locked          =   -1  'True
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   150
         Width           =   1665
      End
      Begin VB.TextBox txtDesconto 
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   7740
         TabIndex        =   5
         Top             =   540
         Width           =   1665
      End
      Begin VB.TextBox txtValorTotal 
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   6750
         Locked          =   -1  'True
         TabIndex        =   10
         TabStop         =   0   'False
         Text            =   "R$"
         Top             =   930
         Width           =   2655
      End
      Begin VB.Label lblDesconto 
         Caption         =   "Desconto"
         Height          =   255
         Left            =   6930
         TabIndex        =   14
         Top             =   630
         Width           =   1065
      End
      Begin VB.Label lblTotalGeral 
         Caption         =   "Total do Or�amento"
         Height          =   225
         Left            =   5280
         TabIndex        =   13
         Top             =   990
         Width           =   1425
      End
      Begin VB.Label lblSubTotal 
         Caption         =   "Sub Total"
         Height          =   255
         Left            =   6930
         TabIndex        =   12
         Top             =   240
         Width           =   705
      End
   End
End
Attribute VB_Name = "frmCompra"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const MAXCOLS = 5
Const MAXCOLSPG = 3

Dim lFirstColumnEdited As Boolean
Dim lInserir           As Boolean
Dim lAlterar           As Boolean
Dim lAlterarGrid       As Boolean
Dim lAlterarGridPg     As Boolean
Dim lInicio            As Boolean
    
Dim lAllowInsert  As Boolean
Dim lAllowEdit    As Boolean
Dim lAllowDelete  As Boolean
Dim lAllowConsult As Boolean

Dim mTotalRows               As Integer
Dim mTotalRowsAntigos        As Integer
Dim dbgrdItensArray()        As String
Dim dbgrdItensAntigosArray() As String
    
Dim mTotalPagamentos        As Integer
Dim mValorAVista            As String
Dim mValorAPrazo            As String
Dim mTipoDePagamento        As Long
Dim FormaDePagamentoArray() As String

Dim lPula As Boolean
Dim mDigitBonus As Boolean
Dim mDigitPorcent As Boolean
Dim mlRefazDesconto As Boolean
Dim Row As Integer
Dim mFechar As Boolean

Dim mC�digo As Integer
Dim mOldValue As String

Dim mCGCCPF As String

Dim TBLCompra         As Table
Dim CompraAberto      As Boolean
Dim IndiceCompraAtivo As String

Dim TBLCompraItens         As Table
Dim CompraItensAberto      As Boolean
Dim IndiceCompraItensAtivo As String

Dim TBLPar�metros    As Table
Dim Par�metrosAberto As Boolean

Dim TBLFormaDePagamento         As Table
Dim FormaDePagamentoAberto      As Boolean
Dim IndiceFormaDePagamentoAtivo As String

Dim TBLPlanoDePagamento         As Table
Dim PlanoDePagamentoAberto      As Boolean
Dim IndicePlanoDePagamentoAtivo As String

Dim TBLProduto         As Table
Dim ProdutoAberto      As Boolean
Dim IndiceProdutoAtivo As String

Dim StatusBarAviso$

Dim DataBaseName(1 To 1) As String
Public Relat�rio$
Public TotalDatabaseName%

Public lAtualizar As Boolean
Private Sub AtivaCampos()
    frDadosCadastrais.Enabled = True
    frItens.Enabled = True
    frTotais.Enabled = True
    cmdFormaDePagamento.Enabled = True
    Bot�oGravar (lInserir Or lAllowEdit)
    cmdCancelar.Enabled = (lInserir Or lAllowEdit)
    cmdGravar.Enabled = (lInserir Or lAllowEdit)
End Sub
Private Function AtualizaProduto(ByVal lExclusao As Boolean) As Boolean
    Dim Cont As Integer
    
    For Cont = 0 To mTotalRowsAntigos - 1
        If dbgrdItensAntigosArray(4, Cont) <> Empty Then
            TBLProduto.Seek "=", dbgrdItensAntigosArray(4, Cont)
            
            If TBLProduto.NoMatch Then
                AtualizaProduto = False
                Exit Function
            End If
            
            TBLProduto.Edit
            TBLProduto("QUANTIDADE") = TBLProduto("QUANTIDADE") - dbgrdItensAntigosArray(1, Cont)
            TBLProduto.Update
        End If
    Next
    
    If Not lExclusao Then
        For Cont = 0 To mTotalRows - 1
            TBLProduto.Seek "=", dbgrdItensArray(4, Cont)
            
            If TBLProduto.NoMatch Then
                AtualizaProduto = False
                Exit Function
            End If
            
            TBLProduto.Edit
            TBLProduto("QUANTIDADE") = TBLProduto("QUANTIDADE") + dbgrdItensArray(1, Cont)
            TBLProduto.Update
        Next
    End If
    
    AtualizaProduto = True
End Function
Private Function Cancelamento()
    Dim Confirma��o%, Espa�os%, Msg1$, Msg2$
    
    Msg1 = "Voc� est� preste a cancelar a opera��o que esta realizando !"
    Msg2 = "Tem certeza?"
    Espa�os = ((Len(Msg1) - Len(Msg2)) / 2) + 4
    Msg2 = String(Espa�os, " ") + Msg2
    Confirma��o = MsgBox(Msg1 + vbCr + Msg2, vbYesNo + vbQuestion + vbDefaultButton2, "Confirma��o")
    
    If Confirma��o = vbNo Then
        Cancelamento = False
        Exit Function
    End If
    
    If lInserir Then
        StatusBarAviso = "Inclus�o cancelada"
    End If
    If lAlterar Then
        StatusBarAviso = "Altera��o cancelada"
    End If
    
    BarraDeStatus StatusBarAviso
    
    Bot�oIncluir lAllowInsert
    
    'Limpa todos os campos
    If TBLCompra.RecordCount = 0 Then
        Navega��oInferior False
        Navega��oSuperior False
        Bot�oGravar False
        cmdGravar.Enabled = False
        cmdCancelar.Enabled = False
        DesativaCampos
        ZeraCampos
        Cancelamento = True
        lInserir = False
        lAlterar = False
        lAlterarGrid = False
        lAlterarGridPg = False
        Exit Function
    End If
    
    lInserir = False
    lAlterar = False
    lAlterarGrid = False
    lAlterarGridPg = False
    
    Cancelamento = True
    
    TestaInferior TBLCompra, lAllowEdit, lAllowDelete, lAllowConsult
    TestaSuperior TBLCompra, lAllowEdit, lAllowDelete, lAllowConsult
    
    GetRecords
End Function
Private Sub DesativaCampos()
    frDadosCadastrais.Enabled = False
    frItens.Enabled = False
    frTotais.Enabled = False
    cmdFormaDePagamento.Enabled = False
    cmdGravar.Enabled = False
    cmdCancelar.Enabled = False
    Bot�oGravar False
End Sub
Public Sub Encontrar()
    If Not lAllowConsult Then
        Exit Sub
    End If

End Sub
Public Sub Excluir()
    Dim Confirma��o As Integer, Msg1$, Msg2$, C�digoDoProduto As Variant
    Dim SQL As String

    If lAlterar Then
       If Not Cancelamento Then
            Exit Sub
        End If
    End If

    StatusBarAviso = "Exclus�o"
    BarraDeStatus StatusBarAviso
    
    Msg1 = "Voc� est� preste a apagar um registro !"
    Msg2 = "Tem certeza?"
    Msg2 = String(((Len(Msg1) - Len(Msg2)) / 2), " ") + Msg2
    Confirma��o = MsgBox(Msg1 + vbCr + Msg2, vbYesNo + vbQuestion + vbDefaultButton2, "Confirma��o")
    
    If Confirma��o = vbNo Then
        Exit Sub
    End If
    
    WS.BeginTrans
        
    If Not AtualizaProduto(True) Then
        GoTo ErroAtualiza
    End If
        
    SQL = "DELETE * FROM [COMPRA - ITENS] WHERE [C�DIGO DE COMPRA] = " & TBLCompra("C�DIGO")
    DBFinanceiro.Execute SQL
    
    SQL = "DELETE * FROM [COMPRA - FORMA DE PAGAMENTO] WHERE [C�DIGO DE COMPRA] = " & TBLCompra("C�DIGO")
    DBFinanceiro.Execute SQL
    
    TBLCompra.Delete
            
    If Err <> 0 Then
        GeraMensagemDeErro "Compra - Excluir - " & mC�digo, True
        StatusBarAviso = "Falha na exclus�o"
        BarraDeStatus StatusBarAviso
        Exit Sub
    End If
    
    WS.CommitTrans
    
    Log gUsu�rio, "Exclus�o - Compra: " & txtNotaFiscal & " - " & txtFornecedor
    
    StatusBarAviso = "Exclus�o bem sucedida"
    BarraDeStatus StatusBarAviso
    
    If TBLCompra.RecordCount = 0 Then
        Navega��oInferior False
        Navega��oSuperior False
        Bot�oExcluir False
        Bot�oGravar False
        cmdGravar.Enabled = False
        cmdCancelar.Enabled = False
        DesativaCampos
        ZeraCampos
        Exit Sub
    End If
    
    If TBLCompra.BOF Then
        TBLCompra.MoveFirst
    ElseIf TBLCompra.EOF Then
        TBLCompra.MoveLast
    Else
        TBLCompra.MovePrevious
        If TBLCompra.BOF Then
            TBLCompra.MoveNext
        End If
    End If
    
    GetRecords
    
    TestaInferior TBLCompra, lAllowEdit, lAllowDelete, lAllowConsult
    TestaSuperior TBLCompra, lAllowEdit, lAllowDelete, lAllowConsult
    
    Exit Sub
    
ErroAtualiza:
    GeraMensagemDeErro "COMPRA - Excluir - ErroAtualiza - " & mC�digo, True
End Sub
Public Function ExcluirPagamento() As Boolean
    On Error GoTo Erro
    
    lAlterarGridPg = True
    If Not lInserir Then
        lAlterar = True
    End If
    
    ReDim FormaDePagamentoArray(MAXCOLSPG, 0)
    
    mTotalPagamentos = 0
    mTipoDePagamento = 0
    
    ExcluirPagamento = True
    
    Exit Function
    
Erro:
    ExcluirPagamento = False
End Function
Private Sub FillGrid(ByVal Chave As Long)
    dbgrdItens.ReBind
    
    ReDim dbgrdItensArray(MAXCOLS - 1, 0)
    ReDim dbgrdItensAntigosArray(MAXCOLS - 1, 0)
    
    mTotalRows = 0
    mTotalRowsAntigos = 0
    
    TBLCompraItens.Seek "=", Chave
    If Not TBLCompraItens.NoMatch Then
        Do While Not TBLCompraItens.EOF And TBLCompraItens("C�DIGO DE COMPRA") = Chave
            mTotalRows = mTotalRows + 1
            mTotalRowsAntigos = mTotalRowsAntigos + 1
            ReDim Preserve dbgrdItensArray(MAXCOLS - 1, mTotalRows - 1)
            ReDim Preserve dbgrdItensAntigosArray(MAXCOLS - 1, mTotalRows - 1)
            
            dbgrdItensArray(0, mTotalRows - 1) = SearchProduto(TBLCompraItens("C�DIGO DO PRODUTO")) 'Nome do Produto
            dbgrdItensArray(1, mTotalRows - 1) = FormatStringMask("@V ######0", StrVal(TBLCompraItens("QUANTIDADE"))) 'Quantidade
            dbgrdItensArray(2, mTotalRows - 1) = FormatStringMask("@V ##.###.##0,00", StrVal(TBLCompraItens("VALOR UNIT�RIO"))) 'Pre�o Unit�rio
            dbgrdItensArray(3, mTotalRows - 1) = FormatStringMask("@V ##.###.##0,00", StrVal((TBLCompraItens("VALOR UNIT�RIO") * TBLCompraItens("QUANTIDADE")))) 'Pre�o de Venda
            dbgrdItensArray(4, mTotalRows - 1) = TBLCompraItens("C�DIGO DO PRODUTO") 'C�digo do Produto
            
            dbgrdItensAntigosArray(0, mTotalRows - 1) = SearchProduto(TBLCompraItens("C�DIGO DO PRODUTO")) 'Nome do Produto
            dbgrdItensAntigosArray(1, mTotalRows - 1) = FormatStringMask("@V ######0", StrVal(TBLCompraItens("QUANTIDADE"))) 'Quantidade
            dbgrdItensAntigosArray(2, mTotalRows - 1) = FormatStringMask("@V ##.###.##0,00", StrVal(TBLCompraItens("VALOR UNIT�RIO"))) 'Pre�o Unit�rio
            dbgrdItensAntigosArray(3, mTotalRows - 1) = FormatStringMask("@V ##.###.##0,00", StrVal((TBLCompraItens("VALOR UNIT�RIO") * TBLCompraItens("QUANTIDADE")))) 'Pre�o de Venda
            dbgrdItensAntigosArray(4, mTotalRows - 1) = TBLCompraItens("C�DIGO DO PRODUTO") 'C�digo do Produto
            
            TBLCompraItens.MoveNext

            If TBLCompraItens.EOF Then
                Exit Do
            End If
        Loop
    End If
    
    dbgrdItens.Refresh
End Sub
Private Sub FillGridPg(ByVal Chave As Long)
    ReDim FormaDePagamentoArray(MAXCOLSPG - 1, 0)
    
    mTotalPagamentos = 0
    
    TBLFormaDePagamento.Seek "=", Chave
    If Not TBLFormaDePagamento.NoMatch Then
        Do While Not TBLFormaDePagamento.EOF And TBLFormaDePagamento("C�DIGO DE COMPRA") = Chave
            mTotalPagamentos = mTotalPagamentos + 1
            ReDim Preserve FormaDePagamentoArray(MAXCOLSPG - 1, mTotalPagamentos - 1)
            
            FormaDePagamentoArray(0, mTotalPagamentos - 1) = TBLFormaDePagamento("DOCUMENTO")
            FormaDePagamentoArray(1, mTotalPagamentos - 1) = TBLFormaDePagamento("VENCIMENTO")
            FormaDePagamentoArray(2, mTotalPagamentos - 1) = FormatStringMask("@V ##.###.##0,00", StrVal(TBLFormaDePagamento("VALOR")))

            TBLFormaDePagamento.MoveNext

            If TBLFormaDePagamento.EOF Then
                Exit Do
            End If
        Loop
    End If
End Sub
Public Sub Gravar()
    ReDim Preserve dbgrdItensAntigosArray(MAXCOLS - 1, mTotalRows - 1)
    
    If lInserir Then
        'Pega o novo c�digo interno do produto e atualiza na Tabela Par�metros
        TBLPar�metros.Edit
        If IsNull(TBLPar�metros("COMPRA")) Then
            mC�digo = 1
        Else
            mC�digo = TBLPar�metros("COMPRA") + 1
        End If
        TBLPar�metros("COMPRA") = mC�digo
        TBLPar�metros.Update
        
        If SetRecords Then
            PosRecords
            lInserir = False
            StatusBarAviso = "Inclus�o bem sucedida"
        Else
            StatusBarAviso = "Falha na inclus�o"
        End If
    ElseIf lAlterar Then
        If TBLCompra.RecordCount > 0 And Not TBLCompra.BOF And Not TBLCompra.EOF Then
            mC�digo = TBLCompra("C�DIGO")
            If SetRecords Then
                PosRecords
                lAlterar = False
                lAlterarGrid = False
                lAlterarGridPg = False
                StatusBarAviso = "Altera��o bem sucedida"
            Else
                StatusBarAviso = "Falha na altera��o"
            End If
        End If
    End If
    
    BarraDeStatus StatusBarAviso
    
    TestaInferior TBLCompra, lAllowEdit, lAllowDelete, lAllowConsult
    TestaSuperior TBLCompra, lAllowEdit, lAllowDelete, lAllowConsult
    
    If TBLCompra.RecordCount = 0 Then
        If Not lInserir And Not lAlterar Then
            Bot�oExcluir False
            Bot�oGravar False
            cmdGravar.Enabled = False
            cmdCancelar.Enabled = False
        End If
    Else
        Bot�oExcluir lAllowDelete
    End If
    
    Bot�oIncluir lAllowInsert
    
    If txtNotaFiscal.Enabled Then
        txtNotaFiscal.SetFocus
    End If
End Sub
Private Sub GetRecords()
    On Error GoTo Erro
    
    Dim pValor1 As Currency
    Dim pValor2 As Currency
    Dim pValor3 As Currency
    Dim pValor4 As Currency
    
    lPula = True
    
    If Not lAllowConsult Then
        ZeraCampos
        DesativaCampos
        lPula = False
        Exit Sub
    End If
    
    mC�digo = TBLCompra("C�DIGO")
    mCGCCPF = TBLCompra("FORNECEDOR")
    txtFornecedor = SearchFornecedor(mCGCCPF)
    txtValor = FormatStringMask("@V ##.###.##0,00", StrVal(TBLCompra("VALOR TOTAL DA COMPRA")))
    txtDesconto = FormatStringMask("@V ##.###.##0,00", StrVal(TBLCompra("DESCONTO TOTAL DA COMPRA")))
    txtNotaFiscal = TBLCompra("NOTA FISCAL")
    
    If TBLCompra("DATA DA NOTA FISCAL") <> vbNull Then
        txtDataDaNotaFiscal = FormatStringMask(CheckDataMask, TBLCompra("DATA DA NOTA FISCAL"))
        CorrigeData DataMask, txtDataDaNotaFiscal, TBLCompra("DATA DA NOTA FISCAL")
    Else
        txtDataDaNotaFiscal = DataNula
    End If
    
    CorrigeValorTotal
    
    mTotalPagamentos = TBLCompra("QUANTIDADE DE VENCIMENTOS")
    mValorAVista = FormatStringMask("@V ##.###.##0,00", StrVal(pValor4))
    mValorAPrazo = FormatStringMask("@V ##.###.##0,00", StrVal(TBLCompra("VALOR � PRAZO")))
    mTipoDePagamento = TBLCompra("TIPO DE PAGAMENTO")
    
    FillGrid TBLCompra("C�DIGO")
    FillGridPg TBLCompra("C�DIGO")
    
    lPula = False
    Exit Sub
    
Erro:
    GeraMensagemDeErro "COMPRA - GetRecords"
    lPula = False
    ZeraCampos
    If Not lAllowEdit Then
        DesativaCampos
    End If
End Sub
Public Function GravaPagamento(ByRef Matriz() As String) As Boolean
    On Error GoTo Erro
    
    Dim Cont As Integer, Cont1 As Integer
    
    lAlterarGridPg = True
    If Not lInserir Then
        lAlterar = True
    End If
    
    ReDim FormaDePagamentoArray(MAXCOLSPG, UBound(Matriz, 2) + 1)
    
    'mTotalPagamentos = UBound(Matriz, 2) + 1
    mTotalPagamentos = frmFormaDePagamento.mTotalPagamentos
    mTipoDePagamento = frmFormaDePagamento.mTipoDePagamento
    
    For Cont = 0 To UBound(Matriz, 2)
        For Cont1 = 0 To MAXCOLSPG - 1
            FormaDePagamentoArray(Cont1, Cont) = Matriz(Cont1, Cont)
        Next
    Next
    
    GravaPagamento = True
    
    Exit Function
    
Erro:
    GravaPagamento = False
End Function
Public Function GetPagamentos(ByVal Coluna As Integer, ByVal Linha As Integer) As String
    GetPagamentos = FormaDePagamentoArray(Coluna, Linha)
End Function
Public Sub Imprimir()

End Sub
Public Sub Incluir()
    If lAlterar Then
       If Not Cancelamento Then
            Exit Sub
        End If
    End If
    
    lInserir = True
    
    ZeraCampos
    AtivaCampos
    
    Bot�oGravar (lInserir Or lAllowEdit)
    Bot�oIncluir False
    cmdGravar.Enabled = (lInserir Or lAllowEdit)
    cmdCancelar.Enabled = (lInserir Or lAllowEdit)
    
    Navega��oInferior False
    Navega��oSuperior False
    
    StatusBarAviso = "Inclus�o"
    BarraDeStatus StatusBarAviso
    
    txtNotaFiscal.SetFocus
End Sub
Public Sub MoveFirst()
    If lAlterar Then
       If Not Cancelamento Then
            Exit Sub
        End If
    End If
    
    StatusBarAviso = "Pronto"
    BarraDeStatus StatusBarAviso
    
    TBLCompra.MoveFirst
    
    Navega��oInferior False
    Navega��oSuperior lAllowConsult
    
    GetRecords
End Sub
Public Sub MoveLast()
    If lAlterar Then
       If Not Cancelamento Then
            Exit Sub
        End If
    End If
    
    StatusBarAviso = "Pronto"
    BarraDeStatus StatusBarAviso
    
    TBLCompra.MoveLast
    
    Navega��oInferior lAllowConsult
    Navega��oSuperior False
    
    GetRecords
End Sub
Public Sub MoveNext()
    If lAlterar Then
       If Not Cancelamento Then
            Exit Sub
        End If
    End If
    
    TBLCompra.MoveNext
    If TBLCompra.EOF Then
        TBLCompra.MovePrevious
        Exit Sub
    End If
    
    StatusBarAviso = "Pronto"
    BarraDeStatus StatusBarAviso
    
    Navega��oInferior lAllowConsult
    TestaSuperior TBLCompra, lAllowEdit, lAllowDelete, lAllowConsult
    
    GetRecords
End Sub
Public Sub MovePrevious()
    If lAlterar Then
       If Not Cancelamento Then
            Exit Sub
        End If
    End If
    
    TBLCompra.MovePrevious
    If TBLCompra.BOF Then
        TBLCompra.MoveNext
        Exit Sub
    End If
    
    StatusBarAviso = "Pronto"
    BarraDeStatus StatusBarAviso
    
    Navega��oSuperior lAllowConsult
    TestaInferior TBLCompra, lAllowEdit, lAllowDelete, lAllowConsult
    
    GetRecords
End Sub
Sub PosRecords()
    If TBLCompra.RecordCount = 0 Then
        Exit Sub
    End If
    
    TBLCompra.Seek "=", Val(mC�digo)
    If TBLCompra.NoMatch Then
        'MsgBox "N�o consegui encontrar o cliente com CGC/CPF " + txtCgcCpf, vbExclamation, "Erro"
        TBLCompra.MoveFirst
        Navega��oInferior False
        Navega��oInferior lAllowConsult
    Else
        TestaInferior TBLCompra, lAllowEdit, lAllowDelete, lAllowConsult
        TestaSuperior TBLCompra, lAllowEdit, lAllowDelete, lAllowConsult
    End If
    GetRecords
End Sub
Private Function SetRecords()
    On Error GoTo ErroCompra
    
    Dim Cont As Integer
    Dim Msg As String
    Dim SQL As String
    Dim Confirma��o As Integer, Msg1 As String, Msg2 As String
    Dim pAlterar As Boolean, pInserir As Boolean
    
    WS.BeginTrans 'Inicia uma Transa��o
        
    If lInserir Then
        TBLCompra.AddNew
    Else
        TBLCompra.Edit
    End If
    
    'Inclus�o do cabe�alho do produto
    TBLCompra("C�DIGO") = Val(mC�digo)
    TBLCompra("FORNECEDOR") = mCGCCPF
    TBLCompra("VALOR TOTAL DA COMPRA") = ValStr(txtValor)
    TBLCompra("DESCONTO TOTAL DA COMPRA") = ValStr(txtDesconto)
    TBLCompra("NOTA FISCAL") = txtNotaFiscal
    TBLCompra("DATA DA NOTA FISCAL") = IIf(Trim(StrTran(txtDataDaNotaFiscal, "/")) <> Empty, txtDataDaNotaFiscal, vbNull)
    TBLCompra("BAIXADO") = False
    TBLCompra("TIPO DE PAGAMENTO") = mTipoDePagamento
    TBLCompra("QUANTIDADE DE VENCIMENTOS") = mTotalPagamentos
    TBLCompra("VALOR � PRAZO") = ValStr(mValorAPrazo)
    If lInserir Then
        TBLCompra("USERNAME - CRIA") = gUsu�rio
        TBLCompra("DATA - CRIA") = Date
        TBLCompra("HORA - CRIA") = Time
        TBLCompra("USERNAME - ALTERA") = "VAZIO"
        TBLCompra("DATA - ALTERA") = vbNull
        TBLCompra("HORA - ALTERA") = vbNull
    End If
    If lAlterar Then
        TBLCompra("USERNAME - ALTERA") = gUsu�rio
        TBLCompra("DATA - ALTERA") = Date
        TBLCompra("HORA - ALTERA") = Time
    End If
    TBLCompra.Update
            
    On Error GoTo ErroCompraItens
    'Inclus�o de itens
    If lAlterarGrid Or lInserir Then
        SQL = "DELETE * FROM [COMPRA - ITENS] WHERE [C�DIGO DE COMPRA] = " & Val(mC�digo)
        DBFinanceiro.Execute SQL
        For Cont = 0 To mTotalRows - 1
            TBLCompraItens.AddNew
            TBLCompraItens("C�DIGO DE COMPRA") = Val(mC�digo)
            TBLCompraItens("C�DIGO DO PRODUTO") = dbgrdItensArray(4, Cont)
            TBLCompraItens("QUANTIDADE") = StrVal(dbgrdItensArray(1, Cont))
            TBLCompraItens("VALOR UNIT�RIO") = StrVal(dbgrdItensArray(2, Cont))
            TBLCompraItens.Update
        Next
    End If
    
    On Error GoTo ErroFormaDePagamento
    'Inclus�o de Forma de Pagamento
    If lAlterarGridPg Or lInserir Then
        SQL = "DELETE * FROM [COMPRA - FORMA DE PAGAMENTO] WHERE [C�DIGO DE COMPRA] = " & Val(mC�digo)
        DBFinanceiro.Execute SQL
        For Cont = 0 To mTotalPagamentos - 1
            TBLFormaDePagamento.AddNew
            TBLFormaDePagamento("C�DIGO DE COMPRA") = Val(mC�digo)
            TBLFormaDePagamento("DOCUMENTO") = FormaDePagamentoArray(0, Cont)
            TBLFormaDePagamento("VENCIMENTO") = IIf(Trim(StrTran(FormaDePagamentoArray(1, Cont), "/")) <> Empty, FormaDePagamentoArray(1, Cont), vbNull)
            TBLFormaDePagamento("VALOR") = StrVal(FormaDePagamentoArray(2, Cont))
            TBLFormaDePagamento.Update
        Next
    End If
    
    If Not AtualizaProduto(False) Then
        GoTo ErroAtualiza
    End If
   
    WS.CommitTrans 'Grava as altera��es ou inclus�es se n�o houverem erros
    
    If lInserir Then
        Log gUsu�rio, "Inclus�o - Compra: " & txtNotaFiscal & vbCr & "Fornecedor: " & txtFornecedor
    Else
        Log gUsu�rio, "Altera��o - Compra: " & txtNotaFiscal & vbCr & "Fornecedor: " & txtFornecedor
    End If
    
    lAlterar = False
    lInserir = False
    lAlterarGrid = False
    lAlterarGridPg = False
    
    SetRecords = True
    
    Exit Function
    
ErroCompra:
    TBLCompra.CancelUpdate
    GeraMensagemDeErro "COMPRA - SetRecords - ErroCompra - " & mC�digo, True
    SetRecords = False
    Exit Function
    
ErroCompraItens:
    TBLCompraItens.CancelUpdate
    GeraMensagemDeErro "COMPRA - SetRecords - ErroCompraItens - " & mC�digo, True
    SetRecords = False
    Exit Function
    
ErroFormaDePagamento:
    TBLFormaDePagamento.CancelUpdate
    GeraMensagemDeErro "COMPRA - SetRecords - ErroFormaDePagamento - " & mC�digo, True
    SetRecords = False
    Exit Function
    
ErroAtualiza:
    GeraMensagemDeErro "COMPRA - SetRecords - ErroAtualiza - " & mC�digo, True
    SetRecords = False
    Exit Function
    
End Function
Private Sub ZeraCampos()
    On Error Resume Next
    
    lPula = True
    mC�digo = 0
    txtNotaFiscal = Empty
    txtDataDaNotaFiscal = "  /  /  "
    txtValor = FormatStringMask("@V ##.###.##0,00", "0,00")
    txtValorTotal = "R$" & String(6, " ") & FormatStringMask("@V ##.###.##0,00", "0,00")
    txtDesconto = FormatStringMask("@V ##.###.##0,00", "0,00")
    txtFornecedor = Empty
    ReDim dbgrdItensArray(MAXCOLS - 1, 0)
    ReDim dbgrdItensAntigosArray(MAXCOLS - 1, 0)
    ReDim FormaDePagamentoArray(MAXCOLSPG - 1, 0)
    mTotalRows = 0
    mTotalPagamentos = 0
    dbgrdItens.ReBind
    lPula = False
    mTotalPagamentos = Empty
    mValorAVista = Empty
    mValorAPrazo = Empty
    mTipoDePagamento = 0
End Sub
Private Sub cmdCancelar_Click()
    Cancelamento
End Sub
Private Sub cmdFormaDePagamento_Click()
    If ValStr(Trim(StrTran(txtValorTotal, "R$"))) = 0 Then
        MsgBox "N�o � poss�vel cadastrar uma Forma de Pagamento" & Chr(13) & "Pois o Valor Total � igual a 0", vbInformation, "Aviso"
    Else
        frmFormaDePagamento.mValorAVista = Trim(StrTran(txtValorTotal, "R$"))
        frmFormaDePagamento.mTotalPagamentos = mTotalPagamentos
        frmFormaDePagamento.mTipoDePagamento = mTipoDePagamento
        frmFormaDePagamento.lEdit = True
        frmFormaDePagamento.lCompra = True
        Set frmFormaDePagamento.ptrForm = Me
        Set frmFormaDePagamento.TBLPlanoDePagamento = TBLPlanoDePagamento
        frmFormaDePagamento.Show 1
    End If
End Sub
Private Sub cmdGravar_Click()
    If mCGCCPF = Empty Then
        MsgBox "O campo FORNECEDOR n�o est� preenchido !", vbInformation, "Aviso"
        Exit Sub
    End If
    Gravar
End Sub
Private Sub cmdTabelaCliente_Click()
    If Not lInserir Then
        lAlterar = True
        StatusBarAviso = "Altera��o"
        BarraDeStatus StatusBarAviso
    End If
    Set frmEncontrar.DBBancoDeDados = DBCadastro
    frmEncontrar.NomeDaJanela = "Fornecedor"
    frmEncontrar.Mensagem = "Nenhum fornecedor foi selecionado!"
    frmEncontrar.BancoDeDados = "CADASTRO"
    frmEncontrar.Tabela = "FORNECEDOR"
    frmEncontrar.Indice = "2"
    frmEncontrar.CampoChave = "CGC - CPF"
    frmEncontrar.CampoPreencheLista = "RAZ�O SOCIAL"
    frmEncontrar.Show vbModal
    mCGCCPF = frmEncontrar.Chave
    txtFornecedor = frmEncontrar.Nome
    txtFornecedor.ForeColor = &H80000008
End Sub
Private Sub dbgrdItens_AfterColEdit(ByVal ColIndex As Integer)
    If ColIndex = 0 Then 'Produto
    ElseIf ColIndex = 1 Then 'Quantidade
    ElseIf ColIndex = 2 Then 'Valor Unit�rio
        FormatMask "@V ##.###.##0,00", dbgrdItens
    ElseIf ColIndex = 3 Then 'Valor Total
    End If
End Sub
Private Sub dbgrdItens_AfterUpdate()
    dbgrdItens.Refresh
    lFirstColumnEdited = False
End Sub
Private Sub dbgrdItens_BeforeColEdit(ByVal ColIndex As Integer, ByVal KeyAscii As Integer, Cancel As Integer)
    Dim oldColIndex As Integer
    Dim Valor As String

    If ColIndex = 0 Then 'Produto
        lFirstColumnEdited = True
    ElseIf ColIndex = 1 Then 'Quantidade
    ElseIf ColIndex = 2 Then 'Valor Unit�rio
        'Cancel = 1
    ElseIf ColIndex = 3 Then 'Valor Total
        Cancel = 1
    ElseIf ColIndex = 4 Then 'C�digo do Produto
        Cancel = 1
        dbgrdItens.Col = 3
    End If
End Sub
Private Sub dbgrdItens_BeforeColUpdate(ByVal ColIndex As Integer, oldvalue As Variant, Cancel As Integer)
    If ColIndex = 0 Then 'Produto
        If Not DoProduto(ColIndex) Then
            Cancel = 1
            dbgrdItens.ReBind
        End If
    ElseIf ColIndex = 1 Then 'Quantidade
        DoQuantidade ColIndex
    ElseIf ColIndex = 2 Then 'Valor Unit�rio
        DoValorUnit�rio ColIndex
    ElseIf ColIndex = 3 Then 'Valor Total
    End If
End Sub
Private Sub dbgrdItens_BeforeDelete(Cancel As Integer)
    Dim pValor As Currency, pDesconto As Currency
    
    If Not lInserir Then
        lAlterar = True
        lAlterarGrid = True
        StatusBarAviso = "Altera��o da Compra"
        BarraDeStatus StatusBarAviso
    End If
    
    dbgrdItens.Col = 3
    pValor = ValStr(dbgrdItens.Text)
       
    pValor = ValStr(txtValor) - pValor
    
    lPula = True
    txtValor = FormatStringMask("@V ##.###.##0,00", StrVal(pValor))
    CorrigeValorTotal
    lPula = False
End Sub
Private Sub dbgrdItens_Change()
    If Not lInserir Then
        lAlterar = True
        lAlterarGrid = True
        StatusBarAviso = "Altera��o da Compra"
        BarraDeStatus StatusBarAviso
    End If
    If dbgrdItens.Col = 0 Then
    ElseIf dbgrdItens.Col = 1 Then
    ElseIf dbgrdItens.Col = 2 Then
        FormatMask "@K 99.999.999,99", dbgrdItens
        lAlterar = True
        lAlterarGrid = True
        StatusBarAviso = "Altera��o da Compra"
        BarraDeStatus StatusBarAviso
    ElseIf dbgrdItens.Col = 3 Then
    End If
End Sub
Private Sub dbgrdItens_ColResize(ByVal ColIndex As Integer, Cancel As Integer)
    Cancel = 1
End Sub
Private Sub dbgrdItens_RowResize(Cancel As Integer)
    Cancel = 1
End Sub
Private Sub dbgrdItens_UnboundAddData(ByVal RowBuf As MSDBGrid.RowBuffer, NewRowBookmark As Variant)
    Dim Col%
        
    mTotalRows = mTotalRows + 1
    ReDim Preserve dbgrdItensArray(MAXCOLS - 1, mTotalRows - 1)
    
    'Sets the bookmark to the last row.
    NewRowBookmark = mTotalRows - 1
    
    ' The following loop adds a new record to the database.
    For Col = 0 To UBound(dbgrdItensArray, 1)
        If Not IsNull(RowBuf.Value(0, Col)) Then
            dbgrdItensArray(Col, mTotalRows - 1) = RowBuf.Value(0, Col)
        Else
            ' If no value set for column, then use the
            ' DefaultValue
            dbgrdItensArray(Col, mTotalRows - 1) = dbgrdItens.Columns(Col).DefaultValue
        End If
    Next
End Sub
Private Sub dbgrdItens_UnboundDeleteRow(Bookmark As Variant)
    Dim iCol As Integer, iRow As Integer
    
    ' Move all rows above the deleted row down in the
    ' array.
    
    For iRow = Bookmark + 1 To mTotalRows - 1
        For iCol = 0 To MAXCOLS - 1
            dbgrdItensArray(iCol, iRow - 1) = dbgrdItensArray(iCol, iRow)
        Next iCol
    Next iRow
    
    mTotalRows = mTotalRows - 1
End Sub
Private Sub dbgrdItens_UnboundReadData(ByVal RowBuf As MSDBGrid.RowBuffer, StartLocation As Variant, ByVal ReadPriorRows As Boolean)
    Dim CurRow&, iRow As Integer, iCol As Integer, iRowsFetched As Integer, iIncr As Integer
    ' DBGrid is requesting rows so give them to it
    
    If mTotalRows = 0 Then Exit Sub
    
    If ReadPriorRows Then
        iIncr = -1
    Else
        iIncr = 1
    End If
    
    ' If StartLocation is Null then start reading at the end
    ' or beginning of the data set.
    If IsNull(StartLocation) Then
        If ReadPriorRows Then
            CurRow = RowBuf.RowCount - 1
        Else
            CurRow = 0
        End If
    Else
        ' Find the position to start reading based on the
        ' StartLocation bookmark and the iIncr variable
        CurRow = CLng(StartLocation) + iIncr
    End If
    
    ' Transfer data from our data set array to the RowBuf object
    ' which DBGrid uses to display the data
    For iRow = 0 To RowBuf.RowCount - 1
        If CurRow < 0 Or CurRow >= mTotalRows Then Exit For
        For iCol = 0 To UBound(dbgrdItensArray, 1)
            RowBuf.Value(iRow, iCol) = dbgrdItensArray(iCol, CurRow&)
        Next iCol
        ' Set bookmark using CurRow& which is also our
        ' array index
        RowBuf.Bookmark(iRow) = CStr(CurRow)
        CurRow = CurRow + iIncr
        iRowsFetched = iRowsFetched + 1
    Next iRow
    RowBuf.RowCount = iRowsFetched
End Sub
Private Sub dbgrdItens_UnboundWriteData(ByVal RowBuf As MSDBGrid.RowBuffer, WriteLocation As Variant)
    Dim iCol As Integer
    ' Data is being updated
    'MsgBox WriteLocation
    ' Update each column in the data set array
    For iCol = 0 To MAXCOLS - 1
        If Not IsNull(RowBuf.Value(0, iCol)) Then
            dbgrdItensArray(iCol, WriteLocation) = RowBuf.Value(0, iCol)
        End If
    Next iCol
End Sub
Private Sub Form_Activate()
    If mFechar Then
        Unload Me
        Exit Sub
    End If
    If Not CompraAberto Then
        Unload Me
        Exit Sub
    End If
    If Not CompraItensAberto Then
        Unload Me
        Exit Sub
    End If
    If Not Par�metrosAberto Then
        Unload Me
        Exit Sub
    End If
    
    TestaInferior TBLCompra, lAllowEdit, lAllowDelete, lAllowConsult
    TestaSuperior TBLCompra, lAllowEdit, lAllowDelete, lAllowConsult
    
    If TBLCompra.RecordCount = 0 Then
        Bot�oGravar False
        cmdGravar.Enabled = False
        cmdCancelar.Enabled = False
        Bot�oImprimir False
    Else
        Bot�oGravar (lInserir Or lAllowEdit)
        cmdGravar.Enabled = (lInserir Or lAllowEdit)
        cmdCancelar.Enabled = (lInserir Or lAllowEdit)
        Bot�oImprimir True
        If lInicio Then
            txtNotaFiscal.SetFocus
            lInicio = False
        End If
    End If
    
    If lInserir Then
        Bot�oGravar (lInserir Or lAllowEdit)
        cmdGravar.Enabled = (lInserir Or lAllowEdit)
        cmdCancelar.Enabled = (lInserir Or lAllowEdit)
        Navega��oInferior False
        Navega��oSuperior False
        Bot�oExcluir False
        Bot�oIncluir False
    ElseIf lAlterar Then
        Bot�oIncluir lAllowInsert
    Else
        Bot�oIncluir lAllowInsert
        StatusBarAviso = "Pronto"
    End If
    
    BarraDeStatus StatusBarAviso
    dbgrdItens.Refresh

    If lAtualizar Then
        Bot�oAtualizar True
    Else
        Bot�oAtualizar False
    End If
    
    dbgrdItens.Refresh
End Sub
Private Sub Form_Deactivate()
    cmdGravar.Enabled = False
    cmdCancelar.Enabled = False
    Bot�oImprimir False
End Sub
Private Sub Form_Load()
    On Error GoTo Erro
    
    Dim Cont%
    
    ZeraCampos
    
    lAllowInsert = Allow("COMPRA", "I")
    lAllowEdit = Allow("COMPRA", "A")
    lAllowDelete = Allow("COMPRA", "E")
    lAllowConsult = Allow("COMPRA", "C")
    
    lFirstColumnEdited = False
    lInserir = False
    lAlterar = False
    lAlterarGrid = False
    lAlterarGridPg = False
    lInicio = True
    
    CompraAberto = AbreTabela(Dicion�rio, "FINANCEIRO", "COMPRA", DBFinanceiro, TBLCompra, TBLTabela, dbOpenTable)
    
    If CompraAberto Then
        IndiceCompraAtivo = "COMPRA1"
        TBLCompra.Index = IndiceCompraAtivo
    Else
        MsgBox "N�o consegui abrir a tabela 'Compra ' !", vbCritical, "Erro"
        Exit Sub
    End If
    
    CompraItensAberto = AbreTabela(Dicion�rio, "FINANCEIRO", "COMPRA - ITENS", DBFinanceiro, TBLCompraItens, TBLTabela, dbOpenTable)
    
    If CompraItensAberto Then
        IndiceCompraItensAtivo = "COMPRAITENS1"
        TBLCompraItens.Index = IndiceCompraItensAtivo
    Else
        MsgBox "N�o consegui abrir a tabela 'Itens de Compra' !", vbCritical, "Erro"
        Exit Sub
    End If
    
    FormaDePagamentoAberto = AbreTabela(Dicion�rio, "FINANCEIRO", "COMPRA - FORMA DE PAGAMENTO", DBFinanceiro, TBLFormaDePagamento, TBLTabela, dbOpenTable)
    
    If FormaDePagamentoAberto Then
        IndiceFormaDePagamentoAtivo = "COMPRAFORMADEPAGAMENTO2"
        TBLFormaDePagamento.Index = IndiceFormaDePagamentoAtivo
    Else
        MsgBox "N�o consegui abrir a tabela 'Forma de Pagamento - Compra' !", vbCritical, "Erro"
        Exit Sub
    End If
    
    PlanoDePagamentoAberto = AbreTabela(Dicion�rio, "FINANCEIRO", "PLANO DE PAGAMENTO", DBFinanceiro, TBLPlanoDePagamento, TBLTabela, dbOpenTable)
    
    If PlanoDePagamentoAberto Then
        IndicePlanoDePagamentoAtivo = "PLANODEPAGAMENTO1"
        TBLPlanoDePagamento.Index = IndicePlanoDePagamentoAtivo
    Else
        MsgBox "N�o consegui abrir a tabela 'Forma de Pagamento' !", vbCritical, "Erro"
        Exit Sub
    End If
    
    Par�metrosAberto = AbreTabela(Dicion�rio, "SISTEMA", "PAR�METROS", DBSistema, TBLPar�metros, TBLTabela, dbOpenTable)
    
    If Par�metrosAberto Then
    Else
        MsgBox "N�o consegui abrir a tabela 'Par�metros' !", vbCritical, "Erro"
        Exit Sub
    End If
    
    ProdutoAberto = AbreTabela(Dicion�rio, "CADASTRO", "PRODUTO", DBCadastro, TBLProduto, TBLTabela, dbOpenTable)
    
    If ProdutoAberto Then
        IndiceProdutoAtivo = "PRODUTO1"
        TBLProduto.Index = IndiceProdutoAtivo
    Else
        MsgBox "N�o consegui abrir a tabela 'Produto' !", vbCritical, "Erro"
        Exit Sub
    End If
    
    dbgrdItens.Columns.Add 1
    dbgrdItens.Columns.Add 1
    dbgrdItens.Columns.Add 1
    
    For Cont = 0 To dbgrdItens.Columns.Count - 1
        dbgrdItens.Columns(Cont).Visible = True
    Next
       
    dbgrdItens.Columns(0).Caption = "Produto"
    dbgrdItens.Columns(0).Width = 3245
    dbgrdItens.Columns(0).DefaultValue = " "
    dbgrdItens.Columns(0).Alignment = dbgLeft
    
    dbgrdItens.Columns(1).Caption = "Quantidade"
    dbgrdItens.Columns(1).Width = 1000
    dbgrdItens.Columns(1).DefaultValue = "0"
    dbgrdItens.Columns(1).Alignment = dbgRight
    
    dbgrdItens.Columns(2).Caption = "Valor Unit�rio"
    dbgrdItens.Columns(2).Width = 2310
    dbgrdItens.Columns(2).DefaultValue = "0,00"
    dbgrdItens.Columns(2).Alignment = dbgRight
    
    dbgrdItens.Columns(3).Caption = "Valor Total"
    dbgrdItens.Columns(3).Width = 2310
    dbgrdItens.Columns(3).DefaultValue = "0,00"
    dbgrdItens.Columns(3).Alignment = dbgRight
    
    dbgrdItens.Columns(4).Caption = "" 'C�digo do Produto
    dbgrdItens.Columns(4).Width = 1
    dbgrdItens.Columns(4).DefaultValue = "0"
    
    dbgrdItens.ReBind
    
    Bot�oIncluir lAllowInsert
 
    If TBLCompra.RecordCount = 0 Then
        DesativaCampos
        Bot�oExcluir False
        Bot�oGravar False
    Else
        AtivaCampos
        Bot�oExcluir lAllowDelete
        Bot�oGravar (lInserir Or lAllowEdit)
        GetRecords
    End If
    
    Navega��oInferior False
        
    If TBLCompra.RecordCount = 0 Or TBLCompra.RecordCount = 1 Then
        Navega��oSuperior False
    Else
        Navega��oInferior lAllowConsult
    End If

    StatusBarAviso = "Pronto"
    
    TotalDatabaseName = 0
    mFechar = False
    Exit Sub
    
Erro:
    GeraMensagemDeErro "COMPRA - Load"
    mFechar = True
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If lInserir Then
        MsgBox "Voc� est� em uma inclus�o!", vbExclamation, Caption
        StatusBarAviso = "Finalize a inclus�o"
        BarraDeStatus StatusBarAviso
        Cancel = 1
        SetaFocus Me
        mdiGeal.Mostrar
        Exit Sub
    End If
    If lAlterar Then
        MsgBox "Voc� est� em uma altera��o!", vbExclamation, Caption
        StatusBarAviso = "Finalize a altera��o"
        BarraDeStatus StatusBarAviso
        Cancel = 1
        SetaFocus Me
        mdiGeal.Mostrar
        Exit Sub
    End If
    
    mdiGeal.StatusBar.Panels("Posi��o").Visible = False
    ResizeStatusBar
    
    Set frmVenda = Nothing
End Sub
Private Sub Form_Unload(Cancel As Integer)
    If CompraAberto Then
        TBLCompra.Close
    End If
    If CompraItensAberto Then
        TBLCompraItens.Close
    End If
    If PlanoDePagamentoAberto Then
        TBLPlanoDePagamento.Close
    End If
    If FormaDePagamentoAberto Then
        TBLFormaDePagamento.Close
    End If
    If Forms.Count = 2 Then
        AllBot�es False
    End If
End Sub
Private Sub txtDataDaNotaFiscal_Change()
    If Not lPula Then
        lPula = True
        FormatMask DataMask, txtDataDaNotaFiscal
        lPula = False
    End If
End Sub
Private Sub txtDataDaNotaFiscal_KeyPress(KeyAscii As Integer)
    If Not lInserir Then
        lAlterar = True
        StatusBarAviso = "Altera��o"
        BarraDeStatus StatusBarAviso
    End If
End Sub
Private Sub txtDataDaNotaFiscal_LostFocus()
    If StrTran(txtDataDaNotaFiscal.Text, "/") <> Space(8) Then
        lPula = True
        CorrigeData DataMask, txtDataDaNotaFiscal, Date
        lPula = False
        If Not FormatMask(CheckDataMask, txtDataDaNotaFiscal) Then
            Beep
            MsgBox "Data inv�lida !", vbCritical, "Erro"
            txtDataDaNotaFiscal.SelStart = 0
            txtDataDaNotaFiscal.SetFocus
        End If
    End If
End Sub
Private Sub txtDesconto_Change()
    If Not lPula Then
        FormatMask "@K 99.999.999,99", txtDesconto
    End If
End Sub
Private Sub txtDesconto_KeyPress(KeyAscii As Integer)
    If Not lInserir Then
        lAlterar = True
        StatusBarAviso = "Altera��o da Compra"
        BarraDeStatus StatusBarAviso
    End If
End Sub
Private Sub txtDesconto_LostFocus()
    Dim pValor1 As Currency
    Dim pValor2 As Currency
    
    If lPula Then
        Exit Sub
    End If
    
    lPula = True
    FormatMask "@V ##.###.##0,00", txtDesconto
    lPula = False
    
    pValor1 = ValStr(txtValor)
    pValor2 = ValStr(txtDesconto)
    
    If pValor2 > pValor1 Then
        MsgBox "Valor de desconto n�o pode ser maior que o valor da compra!", vbCritical, "Aviso"
        txtDesconto.SetFocus
        Exit Sub
    End If
    
    pValor1 = pValor1 - pValor2
    
    txtValorTotal = "R$" + String(6, " ") + FormatStringMask("@V ##.###.##0,00", StrVal(pValor1))
End Sub
Private Sub txtNotaFiscal_KeyPress(KeyAscii As Integer)
    If Not lInserir Then
        lAlterar = True
        StatusBarAviso = "Altera��o da Compra"
        BarraDeStatus StatusBarAviso
    End If
End Sub
Private Sub txtValor_Change()
    If Not lPula Then
        FormatMask "@K 99.999.999,99", txtValor
    End If
End Sub
Private Sub txtValor_KeyPress(KeyAscii As Integer)
    If Not lInserir Then
        lAlterar = True
        StatusBarAviso = "Altera��o da Compra"
        BarraDeStatus StatusBarAviso
    End If
    mDigitPorcent = True
End Sub
Private Sub txtValor_LostFocus()
    If lPula Then
        Exit Sub
    End If
    
    lPula = True
    FormatMask "@V ##.###.##0,00", txtValor
    lPula = False
End Sub
Private Sub DoQuantidade(ByVal ColIndex As Integer)
    Dim pC�digo As String
    Dim pQuantidade As Integer
    Dim pszValor As String
    Dim pValor As Currency
    Dim pValor1 As Currency
    Dim pValor2 As Currency
    Dim pOldValor As Currency
    Dim pNewValor As Currency
    
    pQuantidade = Val(dbgrdItens.Text)
    
    dbgrdItens.Col = 4
    pC�digo = Val(dbgrdItens.Text)
    
    dbgrdItens.Col = 2
    pValor = ValStr(dbgrdItens.Text)
    
    dbgrdItens.Col = 3
    pOldValor = ValStr(dbgrdItens.Text)
    pValor1 = ValStr(txtValor) - ValStr(dbgrdItens.Text)
    
    pValor2 = pValor
    
    pValor2 = pValor2 * pQuantidade
    dbgrdItens.Text = FormatStringMask("@V ##.###.##0,00", StrVal(pValor2))
    pNewValor = StrVal(dbgrdItens.Text)
    
    pValor2 = pValor1 + pValor2
    pszValor = StrVal(pValor2)
    
    'Atualiza campo Valor
    lPula = True
    txtValor = FormatStringMask("@V ##.###.##0,00", pszValor)
    lPula = False
           
    'Atualiza campo Valor Total
    lPula = True
    CorrigeValorTotal
    lPula = False
    
    dbgrdItens.Col = ColIndex
    dbgrdItens.Text = pQuantidade
End Sub
Private Function DoProduto(ByVal ColIndex As Integer) As Boolean
    Dim pC�digo As String
    Dim plgC�digo As Long
    Dim pQuantidade As Integer
    Dim pszValor As String
    Dim pValor As Currency
    Dim pValor1 As Currency
    Dim pValor2 As Currency
    Dim pOldValor As Currency
    Dim pNewValor As Currency
    Dim Cont As Byte
    
    pC�digo = UCase(dbgrdItens.Text)
    plgC�digo = Val(SearchAdvancedProduto(pC�digo, vbC�digo))
               
    dbgrdItens.Col = 4
    dbgrdItens.Text = plgC�digo
        
    If plgC�digo = 0 Then
        frmEncontraProduto.Show 1
        dbgrdItens.Col = 4
        dbgrdItens.Text = frmEncontraProduto.C�digo
        plgC�digo = Val(frmEncontraProduto.C�digo)
    End If

    If mTotalRows > 0 Then
        For Cont = 0 To mTotalRows - 1
            If dbgrdItensArray(MAXCOLS - 1, Cont) = plgC�digo Then
                MsgBox "O item j� foi inclu�do na tabela!", vbInformation, "Aviso"
                DoProduto = False
                Exit Function
            End If
        Next
    End If
    
    dbgrdItens.Col = 2
    dbgrdItens.Text = "0,00" 'SearchAdvancedProduto(plgC�digo, vbValorUnit�rio, vbIndice2)
    
    dbgrdItens.Col = 1
    If dbgrdItens.Text = "" Then
        dbgrdItens.Text = "1"
        dbgrdItens.Col = 3
        pValor1 = 0#     'SearchAdvancedProduto(plgC�digo, vbValValorUnit�rio, vbIndice2)
        pValor2 = ValStr(txtValor)
        pValor2 = pValor1 + pValor2
        pszValor = StrVal(pValor2)
        lPula = True
        txtValor = FormatStringMask("@V ##.###.##0,00", pszValor)
        CorrigeValorTotal
        lPula = False
        dbgrdItens.Text = "0,00" 'SearchAdvancedProduto(plgC�digo, vbValorUnit�rio, vbIndice2)
    Else
        'Corrige o valor total
        pQuantidade = Val(dbgrdItens.Text)
        dbgrdItens.Col = 3
        pOldValor = ValStr(dbgrdItens.Text)
        pValor1 = ValStr(txtValor) - ValStr(dbgrdItens.Text)
        pValor2 = 0 * pQuantidade 'SearchAdvancedProduto(plgC�digo, vbValValorUnit�rio, vbIndice2) * pQuantidade
        pValor2 = pValor1 + pValor2
        pszValor = StrVal(pValor2)
        lPula = True
        txtValor = FormatStringMask("@V ##.###.##0,00", pszValor)
        CorrigeValorTotal
        lPula = False
        dbgrdItens.Text = FormatStringMask("@V ##.###.##0,00", (SearchAdvancedProduto(plgC�digo, vbValValorUnit�rio, vbIndice2) * pQuantidade))
        pNewValor = ValStr(dbgrdItens.Text)
    End If
    
    'Retorna a descri��o do produto na primeira coluna
    dbgrdItens.Col = ColIndex
    dbgrdItens.Text = SearchAdvancedProduto(plgC�digo, vbDescri��o)
    
    DoProduto = True
End Function
Private Function DoValorUnit�rio(ByVal ColIndex As Integer) As Boolean
    Dim pC�digo As String
    Dim pQuantidade As Integer
    Dim pszValor As String
    Dim pValor As Currency
    Dim pValor1 As Currency
    Dim pValor2 As Currency
    Dim pOldValor As Currency
    Dim pNewValor As Currency
    
    pValor = ValStr(dbgrdItens.Text)
    
    dbgrdItens.Col = 4
    pC�digo = Val(dbgrdItens.Text)
    
    dbgrdItens.Col = 1
    pQuantidade = Val(dbgrdItens.Text)
    
    dbgrdItens.Col = 3
    pOldValor = ValStr(dbgrdItens.Text)
    pszValor = StrVal(pValor * pQuantidade)
    dbgrdItens.Text = FormatStringMask("@V ##.###.##0,00", pszValor)
    
    pszValor = StrVal((ValStr(txtValor) + (pValor * pQuantidade) - pOldValor))
    
    'Atualiza campo Valor
    lPula = True
    txtValor = FormatStringMask("@V ##.###.##0,00", pszValor)
    lPula = False
    
    pNewValor = StrVal(dbgrdItens.Text)
    
    'Atualiza campo Valor Total
    lPula = True
    CorrigeValorTotal
    lPula = False
    
    dbgrdItens.Col = ColIndex
    dbgrdItens.Text = pValor
End Function
Private Sub CorrigeValorTotal()
    Dim pszValor As String
    Dim pValor1 As Currency
    Dim pValor2 As Currency
    
    pValor1 = ValStr(txtValor)
    pValor2 = ValStr(txtDesconto)
    
    pszValor = StrVal(pValor1 - pValor2)
    
    txtValorTotal = "R$" + String(6, " ") + FormatStringMask("@V ##.###.##0,00", pszValor)
End Sub
