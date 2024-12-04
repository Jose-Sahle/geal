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
         Caption         =   "Total do Orçamento"
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

Dim mCódigo As Integer
Dim mOldValue As String

Dim mCGCCPF As String

Dim TBLCompra         As Table
Dim CompraAberto      As Boolean
Dim IndiceCompraAtivo As String

Dim TBLCompraItens         As Table
Dim CompraItensAberto      As Boolean
Dim IndiceCompraItensAtivo As String

Dim TBLParâmetros    As Table
Dim ParâmetrosAberto As Boolean

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
Public Relatório$
Public TotalDatabaseName%

Public lAtualizar As Boolean
Private Sub AtivaCampos()
    frDadosCadastrais.Enabled = True
    frItens.Enabled = True
    frTotais.Enabled = True
    cmdFormaDePagamento.Enabled = True
    BotãoGravar (lInserir Or lAllowEdit)
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
    Dim Confirmação%, Espaços%, Msg1$, Msg2$
    
    Msg1 = "Você está preste a cancelar a operação que esta realizando !"
    Msg2 = "Tem certeza?"
    Espaços = ((Len(Msg1) - Len(Msg2)) / 2) + 4
    Msg2 = String(Espaços, " ") + Msg2
    Confirmação = MsgBox(Msg1 + vbCr + Msg2, vbYesNo + vbQuestion + vbDefaultButton2, "Confirmação")
    
    If Confirmação = vbNo Then
        Cancelamento = False
        Exit Function
    End If
    
    If lInserir Then
        StatusBarAviso = "Inclusão cancelada"
    End If
    If lAlterar Then
        StatusBarAviso = "Alteração cancelada"
    End If
    
    BarraDeStatus StatusBarAviso
    
    BotãoIncluir lAllowInsert
    
    'Limpa todos os campos
    If TBLCompra.RecordCount = 0 Then
        NavegaçãoInferior False
        NavegaçãoSuperior False
        BotãoGravar False
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
    BotãoGravar False
End Sub
Public Sub Encontrar()
    If Not lAllowConsult Then
        Exit Sub
    End If

End Sub
Public Sub Excluir()
    Dim Confirmação As Integer, Msg1$, Msg2$, CódigoDoProduto As Variant
    Dim SQL As String

    If lAlterar Then
       If Not Cancelamento Then
            Exit Sub
        End If
    End If

    StatusBarAviso = "Exclusão"
    BarraDeStatus StatusBarAviso
    
    Msg1 = "Você está preste a apagar um registro !"
    Msg2 = "Tem certeza?"
    Msg2 = String(((Len(Msg1) - Len(Msg2)) / 2), " ") + Msg2
    Confirmação = MsgBox(Msg1 + vbCr + Msg2, vbYesNo + vbQuestion + vbDefaultButton2, "Confirmação")
    
    If Confirmação = vbNo Then
        Exit Sub
    End If
    
    WS.BeginTrans
        
    If Not AtualizaProduto(True) Then
        GoTo ErroAtualiza
    End If
        
    SQL = "DELETE * FROM [COMPRA - ITENS] WHERE [CÓDIGO DE COMPRA] = " & TBLCompra("CÓDIGO")
    DBFinanceiro.Execute SQL
    
    SQL = "DELETE * FROM [COMPRA - FORMA DE PAGAMENTO] WHERE [CÓDIGO DE COMPRA] = " & TBLCompra("CÓDIGO")
    DBFinanceiro.Execute SQL
    
    TBLCompra.Delete
            
    If Err <> 0 Then
        GeraMensagemDeErro "Compra - Excluir - " & mCódigo, True
        StatusBarAviso = "Falha na exclusão"
        BarraDeStatus StatusBarAviso
        Exit Sub
    End If
    
    WS.CommitTrans
    
    Log gUsuário, "Exclusão - Compra: " & txtNotaFiscal & " - " & txtFornecedor
    
    StatusBarAviso = "Exclusão bem sucedida"
    BarraDeStatus StatusBarAviso
    
    If TBLCompra.RecordCount = 0 Then
        NavegaçãoInferior False
        NavegaçãoSuperior False
        BotãoExcluir False
        BotãoGravar False
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
    GeraMensagemDeErro "COMPRA - Excluir - ErroAtualiza - " & mCódigo, True
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
        Do While Not TBLCompraItens.EOF And TBLCompraItens("CÓDIGO DE COMPRA") = Chave
            mTotalRows = mTotalRows + 1
            mTotalRowsAntigos = mTotalRowsAntigos + 1
            ReDim Preserve dbgrdItensArray(MAXCOLS - 1, mTotalRows - 1)
            ReDim Preserve dbgrdItensAntigosArray(MAXCOLS - 1, mTotalRows - 1)
            
            dbgrdItensArray(0, mTotalRows - 1) = SearchProduto(TBLCompraItens("CÓDIGO DO PRODUTO")) 'Nome do Produto
            dbgrdItensArray(1, mTotalRows - 1) = FormatStringMask("@V ######0", StrVal(TBLCompraItens("QUANTIDADE"))) 'Quantidade
            dbgrdItensArray(2, mTotalRows - 1) = FormatStringMask("@V ##.###.##0,00", StrVal(TBLCompraItens("VALOR UNITÁRIO"))) 'Preço Unitário
            dbgrdItensArray(3, mTotalRows - 1) = FormatStringMask("@V ##.###.##0,00", StrVal((TBLCompraItens("VALOR UNITÁRIO") * TBLCompraItens("QUANTIDADE")))) 'Preço de Venda
            dbgrdItensArray(4, mTotalRows - 1) = TBLCompraItens("CÓDIGO DO PRODUTO") 'Código do Produto
            
            dbgrdItensAntigosArray(0, mTotalRows - 1) = SearchProduto(TBLCompraItens("CÓDIGO DO PRODUTO")) 'Nome do Produto
            dbgrdItensAntigosArray(1, mTotalRows - 1) = FormatStringMask("@V ######0", StrVal(TBLCompraItens("QUANTIDADE"))) 'Quantidade
            dbgrdItensAntigosArray(2, mTotalRows - 1) = FormatStringMask("@V ##.###.##0,00", StrVal(TBLCompraItens("VALOR UNITÁRIO"))) 'Preço Unitário
            dbgrdItensAntigosArray(3, mTotalRows - 1) = FormatStringMask("@V ##.###.##0,00", StrVal((TBLCompraItens("VALOR UNITÁRIO") * TBLCompraItens("QUANTIDADE")))) 'Preço de Venda
            dbgrdItensAntigosArray(4, mTotalRows - 1) = TBLCompraItens("CÓDIGO DO PRODUTO") 'Código do Produto
            
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
        Do While Not TBLFormaDePagamento.EOF And TBLFormaDePagamento("CÓDIGO DE COMPRA") = Chave
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
        'Pega o novo código interno do produto e atualiza na Tabela Parâmetros
        TBLParâmetros.Edit
        If IsNull(TBLParâmetros("COMPRA")) Then
            mCódigo = 1
        Else
            mCódigo = TBLParâmetros("COMPRA") + 1
        End If
        TBLParâmetros("COMPRA") = mCódigo
        TBLParâmetros.Update
        
        If SetRecords Then
            PosRecords
            lInserir = False
            StatusBarAviso = "Inclusão bem sucedida"
        Else
            StatusBarAviso = "Falha na inclusão"
        End If
    ElseIf lAlterar Then
        If TBLCompra.RecordCount > 0 And Not TBLCompra.BOF And Not TBLCompra.EOF Then
            mCódigo = TBLCompra("CÓDIGO")
            If SetRecords Then
                PosRecords
                lAlterar = False
                lAlterarGrid = False
                lAlterarGridPg = False
                StatusBarAviso = "Alteração bem sucedida"
            Else
                StatusBarAviso = "Falha na alteração"
            End If
        End If
    End If
    
    BarraDeStatus StatusBarAviso
    
    TestaInferior TBLCompra, lAllowEdit, lAllowDelete, lAllowConsult
    TestaSuperior TBLCompra, lAllowEdit, lAllowDelete, lAllowConsult
    
    If TBLCompra.RecordCount = 0 Then
        If Not lInserir And Not lAlterar Then
            BotãoExcluir False
            BotãoGravar False
            cmdGravar.Enabled = False
            cmdCancelar.Enabled = False
        End If
    Else
        BotãoExcluir lAllowDelete
    End If
    
    BotãoIncluir lAllowInsert
    
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
    
    mCódigo = TBLCompra("CÓDIGO")
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
    mValorAPrazo = FormatStringMask("@V ##.###.##0,00", StrVal(TBLCompra("VALOR À PRAZO")))
    mTipoDePagamento = TBLCompra("TIPO DE PAGAMENTO")
    
    FillGrid TBLCompra("CÓDIGO")
    FillGridPg TBLCompra("CÓDIGO")
    
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
    
    BotãoGravar (lInserir Or lAllowEdit)
    BotãoIncluir False
    cmdGravar.Enabled = (lInserir Or lAllowEdit)
    cmdCancelar.Enabled = (lInserir Or lAllowEdit)
    
    NavegaçãoInferior False
    NavegaçãoSuperior False
    
    StatusBarAviso = "Inclusão"
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
    
    NavegaçãoInferior False
    NavegaçãoSuperior lAllowConsult
    
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
    
    NavegaçãoInferior lAllowConsult
    NavegaçãoSuperior False
    
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
    
    NavegaçãoInferior lAllowConsult
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
    
    NavegaçãoSuperior lAllowConsult
    TestaInferior TBLCompra, lAllowEdit, lAllowDelete, lAllowConsult
    
    GetRecords
End Sub
Sub PosRecords()
    If TBLCompra.RecordCount = 0 Then
        Exit Sub
    End If
    
    TBLCompra.Seek "=", Val(mCódigo)
    If TBLCompra.NoMatch Then
        'MsgBox "Não consegui encontrar o cliente com CGC/CPF " + txtCgcCpf, vbExclamation, "Erro"
        TBLCompra.MoveFirst
        NavegaçãoInferior False
        NavegaçãoInferior lAllowConsult
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
    Dim Confirmação As Integer, Msg1 As String, Msg2 As String
    Dim pAlterar As Boolean, pInserir As Boolean
    
    WS.BeginTrans 'Inicia uma Transação
        
    If lInserir Then
        TBLCompra.AddNew
    Else
        TBLCompra.Edit
    End If
    
    'Inclusão do cabeçalho do produto
    TBLCompra("CÓDIGO") = Val(mCódigo)
    TBLCompra("FORNECEDOR") = mCGCCPF
    TBLCompra("VALOR TOTAL DA COMPRA") = ValStr(txtValor)
    TBLCompra("DESCONTO TOTAL DA COMPRA") = ValStr(txtDesconto)
    TBLCompra("NOTA FISCAL") = txtNotaFiscal
    TBLCompra("DATA DA NOTA FISCAL") = IIf(Trim(StrTran(txtDataDaNotaFiscal, "/")) <> Empty, txtDataDaNotaFiscal, vbNull)
    TBLCompra("BAIXADO") = False
    TBLCompra("TIPO DE PAGAMENTO") = mTipoDePagamento
    TBLCompra("QUANTIDADE DE VENCIMENTOS") = mTotalPagamentos
    TBLCompra("VALOR À PRAZO") = ValStr(mValorAPrazo)
    If lInserir Then
        TBLCompra("USERNAME - CRIA") = gUsuário
        TBLCompra("DATA - CRIA") = Date
        TBLCompra("HORA - CRIA") = Time
        TBLCompra("USERNAME - ALTERA") = "VAZIO"
        TBLCompra("DATA - ALTERA") = vbNull
        TBLCompra("HORA - ALTERA") = vbNull
    End If
    If lAlterar Then
        TBLCompra("USERNAME - ALTERA") = gUsuário
        TBLCompra("DATA - ALTERA") = Date
        TBLCompra("HORA - ALTERA") = Time
    End If
    TBLCompra.Update
            
    On Error GoTo ErroCompraItens
    'Inclusão de itens
    If lAlterarGrid Or lInserir Then
        SQL = "DELETE * FROM [COMPRA - ITENS] WHERE [CÓDIGO DE COMPRA] = " & Val(mCódigo)
        DBFinanceiro.Execute SQL
        For Cont = 0 To mTotalRows - 1
            TBLCompraItens.AddNew
            TBLCompraItens("CÓDIGO DE COMPRA") = Val(mCódigo)
            TBLCompraItens("CÓDIGO DO PRODUTO") = dbgrdItensArray(4, Cont)
            TBLCompraItens("QUANTIDADE") = StrVal(dbgrdItensArray(1, Cont))
            TBLCompraItens("VALOR UNITÁRIO") = StrVal(dbgrdItensArray(2, Cont))
            TBLCompraItens.Update
        Next
    End If
    
    On Error GoTo ErroFormaDePagamento
    'Inclusão de Forma de Pagamento
    If lAlterarGridPg Or lInserir Then
        SQL = "DELETE * FROM [COMPRA - FORMA DE PAGAMENTO] WHERE [CÓDIGO DE COMPRA] = " & Val(mCódigo)
        DBFinanceiro.Execute SQL
        For Cont = 0 To mTotalPagamentos - 1
            TBLFormaDePagamento.AddNew
            TBLFormaDePagamento("CÓDIGO DE COMPRA") = Val(mCódigo)
            TBLFormaDePagamento("DOCUMENTO") = FormaDePagamentoArray(0, Cont)
            TBLFormaDePagamento("VENCIMENTO") = IIf(Trim(StrTran(FormaDePagamentoArray(1, Cont), "/")) <> Empty, FormaDePagamentoArray(1, Cont), vbNull)
            TBLFormaDePagamento("VALOR") = StrVal(FormaDePagamentoArray(2, Cont))
            TBLFormaDePagamento.Update
        Next
    End If
    
    If Not AtualizaProduto(False) Then
        GoTo ErroAtualiza
    End If
   
    WS.CommitTrans 'Grava as alterações ou inclusões se não houverem erros
    
    If lInserir Then
        Log gUsuário, "Inclusão - Compra: " & txtNotaFiscal & vbCr & "Fornecedor: " & txtFornecedor
    Else
        Log gUsuário, "Alteração - Compra: " & txtNotaFiscal & vbCr & "Fornecedor: " & txtFornecedor
    End If
    
    lAlterar = False
    lInserir = False
    lAlterarGrid = False
    lAlterarGridPg = False
    
    SetRecords = True
    
    Exit Function
    
ErroCompra:
    TBLCompra.CancelUpdate
    GeraMensagemDeErro "COMPRA - SetRecords - ErroCompra - " & mCódigo, True
    SetRecords = False
    Exit Function
    
ErroCompraItens:
    TBLCompraItens.CancelUpdate
    GeraMensagemDeErro "COMPRA - SetRecords - ErroCompraItens - " & mCódigo, True
    SetRecords = False
    Exit Function
    
ErroFormaDePagamento:
    TBLFormaDePagamento.CancelUpdate
    GeraMensagemDeErro "COMPRA - SetRecords - ErroFormaDePagamento - " & mCódigo, True
    SetRecords = False
    Exit Function
    
ErroAtualiza:
    GeraMensagemDeErro "COMPRA - SetRecords - ErroAtualiza - " & mCódigo, True
    SetRecords = False
    Exit Function
    
End Function
Private Sub ZeraCampos()
    On Error Resume Next
    
    lPula = True
    mCódigo = 0
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
        MsgBox "Não é possível cadastrar uma Forma de Pagamento" & Chr(13) & "Pois o Valor Total é igual a 0", vbInformation, "Aviso"
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
        MsgBox "O campo FORNECEDOR não está preenchido !", vbInformation, "Aviso"
        Exit Sub
    End If
    Gravar
End Sub
Private Sub cmdTabelaCliente_Click()
    If Not lInserir Then
        lAlterar = True
        StatusBarAviso = "Alteração"
        BarraDeStatus StatusBarAviso
    End If
    Set frmEncontrar.DBBancoDeDados = DBCadastro
    frmEncontrar.NomeDaJanela = "Fornecedor"
    frmEncontrar.Mensagem = "Nenhum fornecedor foi selecionado!"
    frmEncontrar.BancoDeDados = "CADASTRO"
    frmEncontrar.Tabela = "FORNECEDOR"
    frmEncontrar.Indice = "2"
    frmEncontrar.CampoChave = "CGC - CPF"
    frmEncontrar.CampoPreencheLista = "RAZÃO SOCIAL"
    frmEncontrar.Show vbModal
    mCGCCPF = frmEncontrar.Chave
    txtFornecedor = frmEncontrar.Nome
    txtFornecedor.ForeColor = &H80000008
End Sub
Private Sub dbgrdItens_AfterColEdit(ByVal ColIndex As Integer)
    If ColIndex = 0 Then 'Produto
    ElseIf ColIndex = 1 Then 'Quantidade
    ElseIf ColIndex = 2 Then 'Valor Unitário
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
    ElseIf ColIndex = 2 Then 'Valor Unitário
        'Cancel = 1
    ElseIf ColIndex = 3 Then 'Valor Total
        Cancel = 1
    ElseIf ColIndex = 4 Then 'Código do Produto
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
    ElseIf ColIndex = 2 Then 'Valor Unitário
        DoValorUnitário ColIndex
    ElseIf ColIndex = 3 Then 'Valor Total
    End If
End Sub
Private Sub dbgrdItens_BeforeDelete(Cancel As Integer)
    Dim pValor As Currency, pDesconto As Currency
    
    If Not lInserir Then
        lAlterar = True
        lAlterarGrid = True
        StatusBarAviso = "Alteração da Compra"
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
        StatusBarAviso = "Alteração da Compra"
        BarraDeStatus StatusBarAviso
    End If
    If dbgrdItens.Col = 0 Then
    ElseIf dbgrdItens.Col = 1 Then
    ElseIf dbgrdItens.Col = 2 Then
        FormatMask "@K 99.999.999,99", dbgrdItens
        lAlterar = True
        lAlterarGrid = True
        StatusBarAviso = "Alteração da Compra"
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
    If Not ParâmetrosAberto Then
        Unload Me
        Exit Sub
    End If
    
    TestaInferior TBLCompra, lAllowEdit, lAllowDelete, lAllowConsult
    TestaSuperior TBLCompra, lAllowEdit, lAllowDelete, lAllowConsult
    
    If TBLCompra.RecordCount = 0 Then
        BotãoGravar False
        cmdGravar.Enabled = False
        cmdCancelar.Enabled = False
        BotãoImprimir False
    Else
        BotãoGravar (lInserir Or lAllowEdit)
        cmdGravar.Enabled = (lInserir Or lAllowEdit)
        cmdCancelar.Enabled = (lInserir Or lAllowEdit)
        BotãoImprimir True
        If lInicio Then
            txtNotaFiscal.SetFocus
            lInicio = False
        End If
    End If
    
    If lInserir Then
        BotãoGravar (lInserir Or lAllowEdit)
        cmdGravar.Enabled = (lInserir Or lAllowEdit)
        cmdCancelar.Enabled = (lInserir Or lAllowEdit)
        NavegaçãoInferior False
        NavegaçãoSuperior False
        BotãoExcluir False
        BotãoIncluir False
    ElseIf lAlterar Then
        BotãoIncluir lAllowInsert
    Else
        BotãoIncluir lAllowInsert
        StatusBarAviso = "Pronto"
    End If
    
    BarraDeStatus StatusBarAviso
    dbgrdItens.Refresh

    If lAtualizar Then
        BotãoAtualizar True
    Else
        BotãoAtualizar False
    End If
    
    dbgrdItens.Refresh
End Sub
Private Sub Form_Deactivate()
    cmdGravar.Enabled = False
    cmdCancelar.Enabled = False
    BotãoImprimir False
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
    
    CompraAberto = AbreTabela(Dicionário, "FINANCEIRO", "COMPRA", DBFinanceiro, TBLCompra, TBLTabela, dbOpenTable)
    
    If CompraAberto Then
        IndiceCompraAtivo = "COMPRA1"
        TBLCompra.Index = IndiceCompraAtivo
    Else
        MsgBox "Não consegui abrir a tabela 'Compra ' !", vbCritical, "Erro"
        Exit Sub
    End If
    
    CompraItensAberto = AbreTabela(Dicionário, "FINANCEIRO", "COMPRA - ITENS", DBFinanceiro, TBLCompraItens, TBLTabela, dbOpenTable)
    
    If CompraItensAberto Then
        IndiceCompraItensAtivo = "COMPRAITENS1"
        TBLCompraItens.Index = IndiceCompraItensAtivo
    Else
        MsgBox "Não consegui abrir a tabela 'Itens de Compra' !", vbCritical, "Erro"
        Exit Sub
    End If
    
    FormaDePagamentoAberto = AbreTabela(Dicionário, "FINANCEIRO", "COMPRA - FORMA DE PAGAMENTO", DBFinanceiro, TBLFormaDePagamento, TBLTabela, dbOpenTable)
    
    If FormaDePagamentoAberto Then
        IndiceFormaDePagamentoAtivo = "COMPRAFORMADEPAGAMENTO2"
        TBLFormaDePagamento.Index = IndiceFormaDePagamentoAtivo
    Else
        MsgBox "Não consegui abrir a tabela 'Forma de Pagamento - Compra' !", vbCritical, "Erro"
        Exit Sub
    End If
    
    PlanoDePagamentoAberto = AbreTabela(Dicionário, "FINANCEIRO", "PLANO DE PAGAMENTO", DBFinanceiro, TBLPlanoDePagamento, TBLTabela, dbOpenTable)
    
    If PlanoDePagamentoAberto Then
        IndicePlanoDePagamentoAtivo = "PLANODEPAGAMENTO1"
        TBLPlanoDePagamento.Index = IndicePlanoDePagamentoAtivo
    Else
        MsgBox "Não consegui abrir a tabela 'Forma de Pagamento' !", vbCritical, "Erro"
        Exit Sub
    End If
    
    ParâmetrosAberto = AbreTabela(Dicionário, "SISTEMA", "PARÂMETROS", DBSistema, TBLParâmetros, TBLTabela, dbOpenTable)
    
    If ParâmetrosAberto Then
    Else
        MsgBox "Não consegui abrir a tabela 'Parâmetros' !", vbCritical, "Erro"
        Exit Sub
    End If
    
    ProdutoAberto = AbreTabela(Dicionário, "CADASTRO", "PRODUTO", DBCadastro, TBLProduto, TBLTabela, dbOpenTable)
    
    If ProdutoAberto Then
        IndiceProdutoAtivo = "PRODUTO1"
        TBLProduto.Index = IndiceProdutoAtivo
    Else
        MsgBox "Não consegui abrir a tabela 'Produto' !", vbCritical, "Erro"
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
    
    dbgrdItens.Columns(2).Caption = "Valor Unitário"
    dbgrdItens.Columns(2).Width = 2310
    dbgrdItens.Columns(2).DefaultValue = "0,00"
    dbgrdItens.Columns(2).Alignment = dbgRight
    
    dbgrdItens.Columns(3).Caption = "Valor Total"
    dbgrdItens.Columns(3).Width = 2310
    dbgrdItens.Columns(3).DefaultValue = "0,00"
    dbgrdItens.Columns(3).Alignment = dbgRight
    
    dbgrdItens.Columns(4).Caption = "" 'Código do Produto
    dbgrdItens.Columns(4).Width = 1
    dbgrdItens.Columns(4).DefaultValue = "0"
    
    dbgrdItens.ReBind
    
    BotãoIncluir lAllowInsert
 
    If TBLCompra.RecordCount = 0 Then
        DesativaCampos
        BotãoExcluir False
        BotãoGravar False
    Else
        AtivaCampos
        BotãoExcluir lAllowDelete
        BotãoGravar (lInserir Or lAllowEdit)
        GetRecords
    End If
    
    NavegaçãoInferior False
        
    If TBLCompra.RecordCount = 0 Or TBLCompra.RecordCount = 1 Then
        NavegaçãoSuperior False
    Else
        NavegaçãoInferior lAllowConsult
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
        MsgBox "Você está em uma inclusão!", vbExclamation, Caption
        StatusBarAviso = "Finalize a inclusão"
        BarraDeStatus StatusBarAviso
        Cancel = 1
        SetaFocus Me
        mdiGeal.Mostrar
        Exit Sub
    End If
    If lAlterar Then
        MsgBox "Você está em uma alteração!", vbExclamation, Caption
        StatusBarAviso = "Finalize a alteração"
        BarraDeStatus StatusBarAviso
        Cancel = 1
        SetaFocus Me
        mdiGeal.Mostrar
        Exit Sub
    End If
    
    mdiGeal.StatusBar.Panels("Posição").Visible = False
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
        AllBotões False
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
        StatusBarAviso = "Alteração"
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
            MsgBox "Data inválida !", vbCritical, "Erro"
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
        StatusBarAviso = "Alteração da Compra"
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
        MsgBox "Valor de desconto não pode ser maior que o valor da compra!", vbCritical, "Aviso"
        txtDesconto.SetFocus
        Exit Sub
    End If
    
    pValor1 = pValor1 - pValor2
    
    txtValorTotal = "R$" + String(6, " ") + FormatStringMask("@V ##.###.##0,00", StrVal(pValor1))
End Sub
Private Sub txtNotaFiscal_KeyPress(KeyAscii As Integer)
    If Not lInserir Then
        lAlterar = True
        StatusBarAviso = "Alteração da Compra"
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
        StatusBarAviso = "Alteração da Compra"
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
    Dim pCódigo As String
    Dim pQuantidade As Integer
    Dim pszValor As String
    Dim pValor As Currency
    Dim pValor1 As Currency
    Dim pValor2 As Currency
    Dim pOldValor As Currency
    Dim pNewValor As Currency
    
    pQuantidade = Val(dbgrdItens.Text)
    
    dbgrdItens.Col = 4
    pCódigo = Val(dbgrdItens.Text)
    
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
    Dim pCódigo As String
    Dim plgCódigo As Long
    Dim pQuantidade As Integer
    Dim pszValor As String
    Dim pValor As Currency
    Dim pValor1 As Currency
    Dim pValor2 As Currency
    Dim pOldValor As Currency
    Dim pNewValor As Currency
    Dim Cont As Byte
    
    pCódigo = UCase(dbgrdItens.Text)
    plgCódigo = Val(SearchAdvancedProduto(pCódigo, vbCódigo))
               
    dbgrdItens.Col = 4
    dbgrdItens.Text = plgCódigo
        
    If plgCódigo = 0 Then
        frmEncontraProduto.Show 1
        dbgrdItens.Col = 4
        dbgrdItens.Text = frmEncontraProduto.Código
        plgCódigo = Val(frmEncontraProduto.Código)
    End If

    If mTotalRows > 0 Then
        For Cont = 0 To mTotalRows - 1
            If dbgrdItensArray(MAXCOLS - 1, Cont) = plgCódigo Then
                MsgBox "O item já foi incluído na tabela!", vbInformation, "Aviso"
                DoProduto = False
                Exit Function
            End If
        Next
    End If
    
    dbgrdItens.Col = 2
    dbgrdItens.Text = "0,00" 'SearchAdvancedProduto(plgCódigo, vbValorUnitário, vbIndice2)
    
    dbgrdItens.Col = 1
    If dbgrdItens.Text = "" Then
        dbgrdItens.Text = "1"
        dbgrdItens.Col = 3
        pValor1 = 0#     'SearchAdvancedProduto(plgCódigo, vbValValorUnitário, vbIndice2)
        pValor2 = ValStr(txtValor)
        pValor2 = pValor1 + pValor2
        pszValor = StrVal(pValor2)
        lPula = True
        txtValor = FormatStringMask("@V ##.###.##0,00", pszValor)
        CorrigeValorTotal
        lPula = False
        dbgrdItens.Text = "0,00" 'SearchAdvancedProduto(plgCódigo, vbValorUnitário, vbIndice2)
    Else
        'Corrige o valor total
        pQuantidade = Val(dbgrdItens.Text)
        dbgrdItens.Col = 3
        pOldValor = ValStr(dbgrdItens.Text)
        pValor1 = ValStr(txtValor) - ValStr(dbgrdItens.Text)
        pValor2 = 0 * pQuantidade 'SearchAdvancedProduto(plgCódigo, vbValValorUnitário, vbIndice2) * pQuantidade
        pValor2 = pValor1 + pValor2
        pszValor = StrVal(pValor2)
        lPula = True
        txtValor = FormatStringMask("@V ##.###.##0,00", pszValor)
        CorrigeValorTotal
        lPula = False
        dbgrdItens.Text = FormatStringMask("@V ##.###.##0,00", (SearchAdvancedProduto(plgCódigo, vbValValorUnitário, vbIndice2) * pQuantidade))
        pNewValor = ValStr(dbgrdItens.Text)
    End If
    
    'Retorna a descrição do produto na primeira coluna
    dbgrdItens.Col = ColIndex
    dbgrdItens.Text = SearchAdvancedProduto(plgCódigo, vbDescrição)
    
    DoProduto = True
End Function
Private Function DoValorUnitário(ByVal ColIndex As Integer) As Boolean
    Dim pCódigo As String
    Dim pQuantidade As Integer
    Dim pszValor As String
    Dim pValor As Currency
    Dim pValor1 As Currency
    Dim pValor2 As Currency
    Dim pOldValor As Currency
    Dim pNewValor As Currency
    
    pValor = ValStr(dbgrdItens.Text)
    
    dbgrdItens.Col = 4
    pCódigo = Val(dbgrdItens.Text)
    
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
