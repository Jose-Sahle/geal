VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form frmVenda 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Venda"
   ClientHeight    =   6975
   ClientLeft      =   1215
   ClientTop       =   1245
   ClientWidth     =   9540
   Icon            =   "Venda.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6975
   ScaleWidth      =   9540
   Begin VB.Frame frObserva��o 
      Caption         =   "Observa��o"
      Height          =   1095
      Left            =   0
      TabIndex        =   24
      Top             =   3750
      Width           =   9525
      Begin VB.TextBox txtObserva��o 
         Height          =   765
         Left            =   90
         MultiLine       =   -1  'True
         TabIndex        =   5
         Top             =   210
         Width           =   9315
      End
   End
   Begin VB.Frame frTotais 
      Caption         =   "Totais"
      Height          =   1695
      Left            =   0
      TabIndex        =   19
      Top             =   4860
      Width           =   9525
      Begin VB.TextBox txtPorcentagemBonus 
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
         TabIndex        =   8
         Text            =   "  0,00"
         Top             =   930
         Width           =   855
      End
      Begin VB.TextBox txtValorBonus 
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
         TabIndex        =   9
         Text            =   "         0,00"
         Top             =   930
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
         Top             =   1320
         Width           =   2655
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
         Locked          =   -1  'True
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   540
         Width           =   1665
      End
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
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   150
         Width           =   1665
      End
      Begin VB.Label lblSubTotal 
         Caption         =   "Sub Total"
         Height          =   255
         Left            =   6930
         TabIndex        =   23
         Top             =   240
         Width           =   705
      End
      Begin VB.Label lblTotalGeral 
         Caption         =   "Total do Or�amento"
         Height          =   225
         Left            =   5280
         TabIndex        =   22
         Top             =   1350
         Width           =   1425
      End
      Begin VB.Label lblDesconto 
         Caption         =   "Desconto"
         Height          =   255
         Left            =   6930
         TabIndex        =   21
         Top             =   630
         Width           =   1065
      End
      Begin VB.Label lblBonus 
         Caption         =   "Bonus"
         Height          =   195
         Left            =   6180
         TabIndex        =   20
         Top             =   990
         Width           =   495
      End
   End
   Begin VB.CommandButton cmdFormaDePagamento 
      Caption         =   "&Forma de Pagemanto"
      Height          =   345
      Left            =   30
      TabIndex        =   11
      Top             =   6630
      Width           =   1980
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   345
      Left            =   8280
      TabIndex        =   13
      Top             =   6615
      Width           =   1245
   End
   Begin VB.CommandButton cmdGravar 
      Caption         =   "&Gravar"
      Height          =   345
      Left            =   6960
      TabIndex        =   12
      Top             =   6615
      Width           =   1245
   End
   Begin VB.Frame frItens 
      Caption         =   " Itens "
      Height          =   2595
      Left            =   0
      TabIndex        =   18
      Top             =   1140
      Width           =   9540
      Begin MSDBGrid.DBGrid dbgrdItens 
         Height          =   2325
         Left            =   60
         OleObjectBlob   =   "Venda.frx":030A
         TabIndex        =   4
         Top             =   210
         Width           =   9405
      End
   End
   Begin VB.Frame frDadosCadastrais 
      Height          =   1140
      Left            =   0
      TabIndex        =   14
      Top             =   0
      Width           =   9540
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
         Left            =   6480
         TabIndex        =   2
         Top             =   660
         Width           =   375
      End
      Begin VB.TextBox txtOr�amento 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1230
         TabIndex        =   0
         Top             =   300
         Width           =   765
      End
      Begin VB.TextBox txtData 
         Height          =   285
         Left            =   8250
         TabIndex        =   3
         Text            =   "  /  /"
         Top             =   690
         Width           =   990
      End
      Begin VB.TextBox txtCliente 
         Height          =   300
         Left            =   1230
         Locked          =   -1  'True
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   690
         Width           =   5235
      End
      Begin VB.Label lblOr�amento 
         Caption         =   "Or�amento"
         Height          =   180
         Left            =   150
         TabIndex        =   17
         Top             =   330
         Width           =   825
      End
      Begin VB.Label lblData 
         Caption         =   "Data"
         Height          =   210
         Left            =   7680
         TabIndex        =   16
         Top             =   720
         Width           =   465
      End
      Begin VB.Label lblCliente 
         Caption         =   "Cliente"
         Height          =   180
         Left            =   150
         TabIndex        =   15
         Top             =   720
         Width           =   645
      End
   End
End
Attribute VB_Name = "frmVenda"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const MAXCOLS = 6
Const MAXCOLSPG = 3

Dim mUsu�rio As String

Dim lAllowInsert  As Boolean
Dim lAllowEdit    As Boolean
Dim lAllowDelete  As Boolean
Dim lAllowConsult As Boolean
Dim lGoToInsert   As Boolean

Dim lFirstColumnEdited As Boolean
Dim lInserir As Boolean
Dim lAlterar As Boolean
Dim lAlterarGrid As Boolean
Dim lAlterarGridPg As Boolean
Dim lInicio As Boolean
    
Dim mTotalRows%
Dim dbgrdItensArray() As String
    
Dim mTotalPagamentos As Integer
Dim mValorAVista As String
Dim mValorAPrazo As String
Dim mTipoDePagamento As Long
Dim FormaDePagamentoArray() As String

Dim HasLote As String
Dim mValorLote As String
Dim mQuantidadeDeLotes As Integer
Dim mLotesArray() As Variant

Dim lPula As Boolean
Dim mDigitPorcent As Boolean
Dim mlRefazDesconto As Boolean
Dim Row As Integer
Dim mFechar As Boolean

Dim mUsu�rioDescontoM�ximo As String

Dim mC�digo As Long
Dim mC�digoProduto As String
Dim mOldValue As String

Dim mC�digoDoCliente As String

Dim TBLVendas As Table
Dim VendasAberto As Boolean
Dim IndiceVendasAtivo$

Dim TBLVendasItens As Table
Dim VendasItensAberto As Boolean
Dim IndiceVendasItensAtivo$

Dim TBLVendaLotes As Table
Dim VendaLotesAberto As Boolean
Dim IndiceVendaLotesAtivo$

Dim ArrayLotes() As Variant

Dim TBLPar�metros As Table
Dim Par�metrosAberto As Boolean

Dim TBLFormaDePagamento As Table
Dim FormaDePagamentoAberto As Boolean
Dim IndiceFormaDePagamentoAtivo$

Dim TBLPlanoDePagamento As Table
Dim PlanoDePagamentoAberto As Boolean
Dim IndicePlanoDePagamentoAtivo$

Dim TBLLoteDoProduto As Table
Dim LoteDoProdutoAberto As Boolean
Dim IndiceLoteDoProdutoAtivo$

Dim mTipoDeBusca As Byte
Dim mCrit�rioDeBusca As Byte
Dim mCondi��oSQL As String

Dim StatusBarAviso$

Dim DataBaseName(1 To 1) As String
Public Relat�rio$
Public TotalDatabaseName%

Public lAtualizar As Boolean
Private Sub AcertaValores()
    Dim Cont             As Long
    Dim ValorTotal       As Currency
    Dim ValorBonus       As Currency
    Dim ValorPorcentagem As Currency
    Dim ValorDesconto    As Currency
    
    lPula = True
    
    For Cont = 0 To mTotalRows - 1
        ValorTotal = ValorTotal + ValStr(dbgrdItensArray(4, Cont))
        ValorDesconto = ValorDesconto + (ValStr(dbgrdItensArray(4, Cont)) * (ValStr(dbgrdItensArray(3, Cont)) / 100))
    Next
    
    ValorBonus = ValStr(txtValorBonus)
    
    lPula = True
    
    'Atualiza Porcentagem do Bonus
    If (ValorTotal - ValorDesconto) = 0 Then
        ValorPorcentagem = 0
    Else
        ValorPorcentagem = ValorBonus * 100 / (ValorTotal - ValorDesconto)
    End If
    txtPorcentagemBonus = FormatStringMask("@V ##0,00", StrVal(ValorPorcentagem))
    
    txtDesconto = FormatStringMask("@V ##.###.##0,00", ValStr(ValorDesconto))
    
    txtValor = FormatStringMask("@V ##.###.##0,00", StrVal(ValorTotal))
    
    ValorTotal = ValorTotal - ValorDesconto - ValorBonus
    txtValorTotal = "R$" + String(6, " ") + FormatStringMask("@V ##.###.##0,00", StrVal(ValorTotal))
    lPula = False
End Sub
Private Sub AcertaDesconto(ByVal oldvalue As Currency, ByVal Desconto As Currency, ByVal NewValor)
    Dim pValorDescontoTotal As Currency
    Dim pValorDesconto As Currency
    
'    pValorDescontoTotal = StrVal(txtDesconto)
'    pValorDesconto = oldvalue * Desconto
'    pValorDesconto = StrVal(FormatStringMask("@V ##.###.##0,00", ValStr(pValorDesconto)))
'    pValorDescontoTotal = pValorDescontoTotal - pValorDesconto
'    pValorDesconto = NewValor * Desconto
'    pValorDescontoTotal = pValorDescontoTotal + pValorDesconto
'
'    lPula = True
'    txtDesconto = FormatStringMask("@V ##.###.##0,00", ValStr(pValorDescontoTotal))
'    lPula = False
End Sub
Private Sub AdelLote(ByVal Elemento As Integer)
    Dim Cont As Integer
    Dim Tamanho As Integer
    
    For Cont = Elemento To UBound(ArrayLotes) - 1
        Set ArrayLotes(Cont) = ArrayLotes(Cont + 1)
    Next
    Set ArrayLotes(UBound(ArrayLotes)) = Nothing
    
    If UBound(ArrayLotes) - 1 < 0 Then
        Tamanho = 0
    Else
        Tamanho = UBound(ArrayLotes) - 1
    End If
    ReDim Preserve ArrayLotes(0 To Tamanho)
End Sub
Private Sub AtivaCampos()
    frDadosCadastrais.Enabled = True
    frItens.Enabled = True
    frObserva��o.Enabled = True
    frTotais.Enabled = True
    cmdFormaDePagamento.Enabled = True
    Bot�oGravar (lInserir Or lAllowEdit)
    cmdCancelar.Enabled = (lInserir Or lAllowEdit)
    cmdGravar.Enabled = (lInserir Or lAllowEdit)
End Sub
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
    If TBLVendas.RecordCount = 0 Then
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
    
    TestaInferior TBLVendas, lAllowEdit, lAllowDelete, lAllowConsult
    TestaSuperior TBLVendas, lAllowEdit, lAllowDelete, lAllowConsult
    
    GetRecords
End Function
Private Sub DesativaCampos()
    frDadosCadastrais.Enabled = False
    frItens.Enabled = False
    frObserva��o.Enabled = False
    frTotais.Enabled = False
    cmdFormaDePagamento.Enabled = False
    cmdGravar.Enabled = False
    cmdCancelar.Enabled = False
    Bot�oGravar False
End Sub
Private Function DescontoM�ximoOk() As Boolean
    Dim Cont As Integer
    Dim Desconto As Currency
    Dim ValorTotal As Currency
    Dim Usu�rio As String
    Dim AllowDesconto As Boolean
    Dim ValorDesconto As Currency
    
    Desconto = 0
    For Cont = 0 To mTotalRows - 1
        ValorDesconto = SearchAdvancedProduto(dbgrdItensArray(5, Cont), vbDescontoM�ximo, vbIndice2)
        Desconto = Desconto + Round(((1 - (ValorDesconto / 100)) * (ValStr(dbgrdItensArray(2, Cont) * ValStr(dbgrdItensArray(1, Cont))))), 2)
    Next
    
    If Desconto > Round(ValStr(Trim(StrTran(txtValorTotal, "R$")))) Then
        If MsgBox("O valor m�nimo permitido para a venda � de: " & Trim(FormatStringMask("@V ###.###.##0,00", Desconto)) & vbCr & vbCr & "Deseja autorizar a venda!", vbCritical + vbYesNo, "Aviso") = vbNo Then
            DescontoM�ximoOk = False
            mUsu�rioDescontoM�ximo = Empty
            Exit Function
        End If
    Else
        DescontoM�ximoOk = True
        mUsu�rioDescontoM�ximo = Empty
        Exit Function
    End If
    
    'Valida Usu�rio
    frmValidaUsu�rio.Show 1
    
    Usu�rio = frmValidaUsu�rio.Usu�rio
    
    Set frmValidaUsu�rio = Nothing
    
    If Usu�rio = Empty Then
        DescontoM�ximoOk = False
        Exit Function
    End If
    
    AllowDesconto = Allow("VENDA", "D", Usu�rio)
    
    If AllowDesconto Then
        mUsu�rioDescontoM�ximo = Usu�rio
        DescontoM�ximoOk = True
    Else
        MsgBox "Usu�rio " & Usu�rio & " n�o possui autoriza��o " & vbCr & "para validar desconto m�ximo!", vbCritical, "Aviso"
        DescontoM�ximoOk = False
    End If
End Function
Public Sub Encontrar()
    If Not lAllowConsult Then
        Exit Sub
    End If
    Set frmEncontrar.DBBancoDeDados = DBFinanceiro
    frmEncontrar.NomeDaJanela = "Or�amento"
    frmEncontrar.LabelDescription = "C�digo"
    frmEncontrar.Mensagem = "Nenhum or�amento foi selecionado!"
    frmEncontrar.BancoDeDados = "FINANCEIRO"
    frmEncontrar.Tabela = "VENDA"
    frmEncontrar.Indice = "1"
    frmEncontrar.CampoChave = "C�DIGO"
    frmEncontrar.CampoPreencheLista = "TIPO,DATA DO OR�AMENTO,C�DIGO"
    frmEncontrar.Show vbModal
    lPula = True
    txtOr�amento = frmEncontrar.Chave
    lPula = False
    PosRecords
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
        
    SQL = "DELETE * FROM [VENDA - ITENS] WHERE [OR�AMENTO] = " & TBLVendas("C�DIGO")
    DBFinanceiro.Execute SQL
    
    SQL = "DELETE * FROM [VENDA - FORMA DE PAGAMENTO] WHERE [OR�AMENTO] = " & TBLVendas("C�DIGO")
    DBFinanceiro.Execute SQL
    
    SQL = "DELETE * FROM [VENDA - LOTES] WHERE [OR�AMENTO] = " & TBLVendas("C�DIGO")
    DBFinanceiro.Execute SQL
    
    TBLVendas.Delete
            
    If Err <> 0 Then
        GeraMensagemDeErro "Sa�daVenda - Excluir - " & txtOr�amento, True
        StatusBarAviso = "Falha na exclus�o"
        BarraDeStatus StatusBarAviso
        Exit Sub
    End If
    
    WS.CommitTrans
    
    Log gUsu�rio, "Exclus�o - Or�amento: " & txtOr�amento
    
    StatusBarAviso = "Exclus�o bem sucedida"
    BarraDeStatus StatusBarAviso
    
    If TBLVendas.RecordCount = 0 Then
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
    
    If TBLVendas.BOF Then
        TBLVendas.MoveFirst
    ElseIf TBLVendas.EOF Then
        TBLVendas.MoveLast
    Else
        TBLVendas.MovePrevious
        If TBLVendas.BOF Then
            TBLVendas.MoveNext
        End If
    End If
    
    GetRecords
    
    TestaInferior TBLVendas, lAllowEdit, lAllowDelete, lAllowConsult
    TestaSuperior TBLVendas, lAllowEdit, lAllowDelete, lAllowConsult
End Sub
Public Function ExcluirPagamento() As Boolean

    lAlterarGridPg = True
    If Not lInserir Then
        lAlterar = True
    End If
    
    ReDim FormaDePagamentoArray(MAXCOLSPG, 0)
    
    mTotalPagamentos = 0
    mTipoDePagamento = 0
    
    ExcluirPagamento = True
End Function
Private Sub FillGrid(ByVal Chave As Long)
    dbgrdItens.ReBind
    
    ReDim dbgrdItensArray(MAXCOLS - 1, 0)
    
    mTotalRows = 0
    
    TBLVendasItens.Seek "=", Chave
    If Not TBLVendasItens.NoMatch Then
        Do While Not TBLVendasItens.EOF And TBLVendasItens("OR�AMENTO") = Chave
            mTotalRows = mTotalRows + 1
            ReDim Preserve dbgrdItensArray(MAXCOLS - 1, mTotalRows - 1)
            
            FillLote Chave, TBLVendasItens("C�DIGO DO PRODUTO")
            
            dbgrdItensArray(0, mTotalRows - 1) = SearchProduto(TBLVendasItens("C�DIGO DO PRODUTO")) 'Nome do Produto
            dbgrdItensArray(1, mTotalRows - 1) = FormatStringMask("@V ######0,00", StrVal(TBLVendasItens("QUANTIDADE"))) 'Quantidade
            dbgrdItensArray(2, mTotalRows - 1) = FormatStringMask("@V ##.###.##0,00", StrVal(TBLVendasItens("VALOR UNIT�RIO"))) 'Pre�o Unit�rio
            dbgrdItensArray(3, mTotalRows - 1) = FormatStringMask("@V ##.###.##0,00", StrVal(TBLVendasItens("DESCONTO"))) 'Desconto no valor do produto
            dbgrdItensArray(4, mTotalRows - 1) = FormatStringMask("@V ##.###.##0,00", StrVal((TBLVendasItens("VALOR UNIT�RIO") * TBLVendasItens("QUANTIDADE")))) 'Pre�o de Venda
            dbgrdItensArray(5, mTotalRows - 1) = TBLVendasItens("C�DIGO DO PRODUTO") 'C�digo do Produto

            TBLVendasItens.MoveNext

            If TBLVendasItens.EOF Then
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
        Do While Not TBLFormaDePagamento.EOF And TBLFormaDePagamento("OR�AMENTO") = Chave
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
Private Sub FillLote(ByVal Chave As Long, ByVal ChaveProduto As Long)
    On Error GoTo Erro
    
    ReDim Preserve ArrayLotes(0 To mTotalRows - 1)
    
    TBLVendaLotes.Seek "=", Chave, ChaveProduto
    If Not TBLVendaLotes.NoMatch Then
        Set ArrayLotes(mTotalRows - 1) = New ClassLote
        Do While Not TBLVendaLotes.EOF And TBLVendaLotes("OR�AMENTO") = Chave And TBLVendaLotes("C�DIGO DO PRODUTO") = ChaveProduto
            ArrayLotes(mTotalRows - 1).AddNew TBLVendaLotes("C�DIGO DO LOTE"), TBLVendaLotes("M�LTIPLO"), TBLVendaLotes("QUANTIDADE")
            
            TBLVendaLotes.MoveNext

            If TBLVendaLotes.EOF Then
                Exit Do
            End If
        Loop
    Else
        Set ArrayLotes(mTotalRows - 1) = New ClassLote
    End If
    
    Exit Sub
    
Erro:
    GeraMensagemDeErro "Sa�daVenda - FillLote"
End Sub
Public Sub Gravar()
    If lInserir Then
        
        'Pega o novo c�digo interno do produto e atualiza na Tabela Par�metros
        TBLPar�metros.Edit
        mC�digo = TBLPar�metros("OR�AMENTO") + 1
        TBLPar�metros("OR�AMENTO") = mC�digo
        TBLPar�metros.Update
        txtOr�amento = mC�digo
        
        If SetRecords Then
            PosRecords
            lInserir = False
            StatusBarAviso = "Inclus�o bem sucedida"
        Else
            StatusBarAviso = "Falha na inclus�o"
        End If
    ElseIf lAlterar Then
        If TBLVendas("TIPO") <> "O" Then
            MsgBox "Este or�amento n�o pode ser alterado!", vbCritical, "Aviso"
            StatusBarAviso = "Altera��o negada"
            PosRecords
            lAlterar = False
            lAlterarGrid = False
            lAlterarGridPg = False
        Else
            If TBLVendas.RecordCount > 0 And Not TBLVendas.BOF And Not TBLVendas.EOF Then
                mC�digo = TBLVendas("C�DIGO")
                txtOr�amento = mC�digo
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
    End If
    
    BarraDeStatus StatusBarAviso
    
    TestaInferior TBLVendas, lAllowEdit, lAllowDelete, lAllowConsult
    TestaSuperior TBLVendas, lAllowEdit, lAllowDelete, lAllowConsult
    
    If TBLVendas.RecordCount = 0 Then
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
    
    If txtCliente.Enabled Then
        txtCliente.SetFocus
    End If
End Sub
Public Function GravaPagamento(ByRef Matriz() As String) As Boolean
    Dim Cont As Integer, Cont1 As Integer
    
    lAlterarGridPg = True
    If Not lInserir Then
        lAlterar = True
    End If
    
    ReDim FormaDePagamentoArray(MAXCOLSPG, UBound(Matriz, 2) + 1)
    
    mTotalPagamentos = frmFormaDePagamento.mTotalPagamentos
    mTipoDePagamento = frmFormaDePagamento.mTipoDePagamento
    
    For Cont = 0 To UBound(Matriz, 2)
        For Cont1 = 0 To MAXCOLSPG - 1
            FormaDePagamentoArray(Cont1, Cont) = Matriz(Cont1, Cont)
        Next
    Next
    
    GravaPagamento = True
    
End Function
Public Function GetPagamentos(ByVal Coluna As Integer, ByVal Linha As Integer) As String
    GetPagamentos = FormaDePagamentoArray(Coluna, Linha)
End Function
Public Sub Imprimir()
    On Error GoTo Erro
    
    Dim TotalDeLinha As Byte, LinhaCorrente As Byte, ColunaCorrente As Byte
    Dim Empresa As String, TotalColPrinter As Integer, Espa�o As Byte, Papel As Single
    Dim TamanhoLinha As Single, TamanhoColuna As Single, XPosCorrente As Single, YPosCorrente As Single
    Dim XPosC�digo As Single, XPosQuantidade As Single, XPosDescri��o As Single, XPosValorUnit�rio As Single, XPosValorTotal As Single
    Dim YPosInicial As Single, YPosFinal As Single, YPosInicialImpress�o As Single
    Dim XPosQuadro As Single, YPosQuadro As Single
    Dim Cont As Integer, ValorTotal As Currency
    Dim ValorDesconto As Currency, SubValor As Currency
    
    Empresa = TBLPar�metros("EMPRESA")
    
    On Error Resume Next
    'Configura��o da impressora
    Printer.ScaleMode = vbPixels
    Printer.FontBold = False
    Printer.FontItalic = False
    Printer.ColorMode = Printer.ColorMode
    
    On Error GoTo Erro
    
    GoSub Cabe�alho
    
    For Cont = 0 To mTotalRows - 1
        'C�digo
        Printer.CurrentX = XPosC�digo
        Printer.CurrentY = LinhaCorrente * TamanhoLinha
        Printer.Print SearchAdvancedProduto(dbgrdItensArray(5, Cont), vbC�digoDoFornecedor, vbIndice2)

        'Descri��o
        Printer.CurrentX = XPosDescri��o
        Printer.CurrentY = LinhaCorrente * TamanhoLinha
        Printer.Print dbgrdItensArray(0, Cont)

        'Quantidade
        Printer.CurrentX = XPosQuantidade
        Printer.CurrentY = LinhaCorrente * TamanhoLinha
        Printer.Print dbgrdItensArray(1, Cont)

        'Valor Unit�rio
        Printer.CurrentY = LinhaCorrente * TamanhoLinha
        Printer.CurrentX = XPosValorUnit�rio - Printer.TextWidth(FormatStringMask("@V ##.###.##0,00", dbgrdItensArray(2, Cont)))
        Printer.Print FormatStringMask("@V ##.###.##0,00", dbgrdItensArray(2, Cont))

        'Valor Total
        ValorTotal = ValStr(dbgrdItensArray(2, Cont)) * Val(dbgrdItensArray(1, Cont))
        Printer.CurrentY = LinhaCorrente * TamanhoLinha
        Printer.CurrentX = XPosValorTotal - Printer.TextWidth(FormatStringMask("@V ##.###.##0,00", StrVal(ValorTotal)))
        Printer.Print FormatStringMask("@V ##.###.##0,00", StrVal(ValorTotal))

        LinhaCorrente = LinhaCorrente + 1
        
        If LinhaCorrente > TotalDeLinha + 42 Then
            Printer.EndDoc
            GoSub Cabe�alho
        End If
    Next
    
    Printer.FontBold = True
    Printer.FontSize = 12
    
    SubValor = ValStr(txtValor)
    ValorDesconto = ValStr(txtDesconto) + ValStr(txtValorBonus)
    ValorTotal = SubValor - ValorDesconto
    
    If ValorDesconto > 0 Then 'Se houver desconto
        'Sub Total
        Printer.CurrentY = YPosFinal + (2 * Printer.TextHeight("W"))
        Printer.CurrentX = XPosValorUnit�rio - Printer.TextWidth("Sub-Total") - 20
        Printer.Print "Sub-Total"
        
        Printer.CurrentY = YPosFinal + (2 * Printer.TextHeight("W"))
        Printer.CurrentX = XPosValorTotal - Printer.TextWidth(FormatStringMask("@V ##.###.##0,00", StrVal(SubValor))) - 10
        Printer.Print FormatStringMask("@V ##.###.##0,00", StrVal(SubValor))
        
        'Desconto
        Printer.CurrentY = YPosFinal + (3.5 * Printer.TextHeight("W"))
        Printer.CurrentX = XPosValorUnit�rio - Printer.TextWidth("Desconto") - 20
        Printer.Print "Desconto"
        
        Printer.CurrentY = YPosFinal + (3.5 * Printer.TextHeight("W"))
        Printer.CurrentX = XPosValorTotal - Printer.TextWidth(FormatStringMask("@V ##.###.##0,00", StrVal(ValorDesconto))) - 10
        Printer.Print FormatStringMask("@V ##.###.##0,00", StrVal(ValorDesconto))
        
        'Total
        Printer.CurrentY = YPosFinal + (5 * Printer.TextHeight("W"))
        Printer.CurrentX = XPosValorUnit�rio - Printer.TextWidth("Valor Total do Or�amento") - 20
        Printer.Print "Valor Total do Or�amento"
        
        Printer.CurrentY = YPosFinal + (5 * Printer.TextHeight("W"))
        Printer.CurrentX = XPosValorTotal - Printer.TextWidth(FormatStringMask("@V ##.###.##0,00", StrVal(ValorTotal))) - 10
        Printer.Print FormatStringMask("@V ##.###.##0,00", StrVal(ValorTotal))
    Else
        'Total
        Printer.CurrentY = YPosFinal + (5 * Printer.TextHeight("W"))
        Printer.CurrentX = XPosValorUnit�rio - Printer.TextWidth("Valor Total do Or�amento") - 20
        Printer.Print "Valor Total do Or�amento"
        
        Printer.CurrentY = YPosFinal + (5 * Printer.TextHeight("W"))
        Printer.CurrentX = XPosValorTotal - Printer.TextWidth(FormatStringMask("@V ##.###.##0,00", StrVal(ValorTotal))) - 10
        Printer.Print FormatStringMask("@V ##.###.##0,00", StrVal(ValorTotal))
    End If
    
    Printer.EndDoc
    
    Exit Sub
    
Cabe�alho:
    LinhaCorrente = 1
    ColunaCorrente = 1
    
    Printer.FontSize = 28
    Printer.FontBold = True
    
    Papel = Printer.ScaleHeight
    
    TotalColPrinter = Int(Printer.ScaleWidth - Printer.TextWidth(Empresa))
    TotalColPrinter = Int(TotalColPrinter / Printer.TextWidth(" "))
    
    Espa�o = Int(TotalColPrinter / 2)
    
    Printer.Print String(Espa�o, " ") & Empresa & String(TotalColPrinter - Espa�o, " ")
    
    TamanhoColuna = Printer.TextWidth("W")
    
    Printer.CurrentY = Printer.TextHeight(Empresa) + 32
    Printer.CurrentX = 10
    
    Printer.DrawWidth = 10
    Printer.Line -((45 * TamanhoColuna), Printer.CurrentY)
    
    'Volta para o tamanho normal de letras
    Printer.FontSize = 12
    
    TamanhoLinha = Printer.TextHeight("W")
    TamanhoColuna = Printer.TextWidth("W")
    
    Papel = Papel - Printer.TextHeight(Empresa)
    
    ColunaCorrente = 3
    LinhaCorrente = 4
    
    Printer.CurrentX = ColunaCorrente * TamanhoColuna
    Printer.CurrentY = LinhaCorrente * TamanhoLinha
        
    Printer.FontBold = True
    Printer.FontItalic = True
    
    Printer.Print "Or�amento:"
    
    ColunaCorrente = 10
    Printer.CurrentX = ColunaCorrente * TamanhoColuna
    Printer.CurrentY = LinhaCorrente * TamanhoLinha
    
    Printer.Print txtOr�amento
    
    ColunaCorrente = 40
    
    Printer.CurrentX = ColunaCorrente * TamanhoColuna
    Printer.CurrentY = LinhaCorrente * TamanhoLinha
    
    Printer.Print "Data: " & txtData
    
    LinhaCorrente = LinhaCorrente + 1
    ColunaCorrente = 3
    
    Printer.CurrentX = ColunaCorrente * TamanhoColuna
    Printer.CurrentY = LinhaCorrente * TamanhoLinha
    
    Printer.Print "Cliente: "
    XPosCorrente = ColunaCorrente * TamanhoColuna + Printer.TextWidth("Cliente: ")
    
    Printer.DrawWidth = 3
    
    Printer.CurrentX = XPosCorrente
    Printer.CurrentY = LinhaCorrente * TamanhoLinha + Printer.TextHeight("C") + 5
    
    'Posi��es das finais das linhas
    XPosQuadro = 48 * TamanhoColuna
    YPosQuadro = 40 * TamanhoLinha
    
    Printer.Line -(XPosQuadro, Printer.CurrentY)
    
    Printer.FontBold = False
    Printer.CurrentX = XPosCorrente
    Printer.CurrentY = LinhaCorrente * TamanhoLinha
    
    Printer.Print txtCliente
    
    LinhaCorrente = LinhaCorrente + 3
    ColunaCorrente = 3
    
    YPosInicial = LinhaCorrente * TamanhoLinha - 7
    YPosFinal = YPosQuadro
    
    Printer.CurrentY = LinhaCorrente * TamanhoLinha - 7
    Printer.CurrentX = ColunaCorrente * TamanhoColuna
    'Printer.DrawWidth = 7
    
    'Quadro aonde ser�o preenchidos com o produtos
    Printer.Line -((XPosQuadro), YPosQuadro), , B
    
    Printer.CurrentY = LinhaCorrente * TamanhoLinha
    Printer.CurrentX = ColunaCorrente * TamanhoColuna + 10
    
    Printer.FontSize = 10
    Printer.FontBold = True
    
    'Impress�o dos t�tulos do quadro de itens
    YPosInicialImpress�o = Printer.CurrentY
    
    'C�digo
    XPosC�digo = Printer.CurrentX
    Printer.Print "C�digo"
    
    'Descri��o
    ColunaCorrente = 10
    Printer.CurrentY = LinhaCorrente * TamanhoLinha
    Printer.CurrentX = ColunaCorrente * TamanhoColuna
    XPosDescri��o = ColunaCorrente * TamanhoColuna - 5
    Printer.Print "Descri��o"
    
    'Quantidade
    ColunaCorrente = 30
    Printer.CurrentY = LinhaCorrente * TamanhoLinha
    Printer.CurrentX = ColunaCorrente * TamanhoColuna
    XPosQuantidade = ColunaCorrente * TamanhoColuna - 5
    Printer.Print "QT"
    
    'Valor Unit�rio
    ColunaCorrente = 35
    Printer.CurrentY = LinhaCorrente * TamanhoLinha
    Printer.CurrentX = ColunaCorrente * TamanhoColuna
    XPosValorUnit�rio = ColunaCorrente * TamanhoColuna - 5
    Printer.Print "Valor Unit�rio"
    
    'Valor Total
    ColunaCorrente = 41
    Printer.CurrentY = LinhaCorrente * TamanhoLinha
    Printer.CurrentX = ColunaCorrente * TamanhoColuna + 6
    XPosValorTotal = ((ColunaCorrente * TamanhoColuna) + 6) - 5
    Printer.Print "Valor Total"
    
    'Tra�o que divide C�DIGO de DESCRI��O
    Printer.Line (XPosDescri��o, YPosInicial)-(XPosDescri��o, YPosFinal)
    
    'Tra�o que divide DESCRI��O de QUANTIDADE
    Printer.Line (XPosQuantidade, YPosInicial)-(XPosQuantidade, YPosFinal)
    
    'Tra�o que divide QUANTIDADE de VALOR UNIT�RIO
    Printer.Line (XPosValorUnit�rio, YPosInicial)-(XPosValorUnit�rio, YPosFinal)
    
    'Tra�o que divide VALOR UNIT�RIO de VALOR TOTAL
    Printer.Line (XPosValorTotal, YPosInicial)-(XPosValorTotal, YPosFinal)
    
    Printer.FontBold = False
    Printer.FontSize = 8
    
    TamanhoLinha = Printer.TextWidth("W") + 10
    TamanhoColuna = Printer.TextHeight("W")
        
    LinhaCorrente = Int(YPosInicialImpress�o / TamanhoLinha)
    LinhaCorrente = LinhaCorrente + 2
    
    TotalDeLinha = LinhaCorrente
    
    XPosQuantidade = XPosQuantidade + 5
    XPosDescri��o = XPosDescri��o + 5
    
    'Alinhado � direita
    XPosValorUnit�rio = XPosValorTotal - 10
    XPosValorTotal = XPosQuadro - 10
    
    Return
    
Erro:
    GeraMensagemDeErro "Venda - Imprimir"
End Sub
Public Sub Incluir()
    If lAlterar Then
       If Not Cancelamento Then
            Exit Sub
        End If
    End If
    
    Caption = "Venda"
    
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
    
    txtCliente.SetFocus
End Sub
Public Sub MoveFirst()
    If lAlterar Then
       If Not Cancelamento Then
            Exit Sub
        End If
    End If
    
    StatusBarAviso = "Pronto"
    BarraDeStatus StatusBarAviso
    
    TBLVendas.MoveFirst
    
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
    
    TBLVendas.MoveLast
    
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
    
    TBLVendas.MoveNext
    If TBLVendas.EOF Then
        TBLVendas.MovePrevious
        Exit Sub
    End If
    
    StatusBarAviso = "Pronto"
    BarraDeStatus StatusBarAviso
    
    Navega��oInferior lAllowConsult
    TestaSuperior TBLVendas, lAllowEdit, lAllowDelete, lAllowConsult
    
    GetRecords
End Sub
Public Sub MovePrevious()
    If lAlterar Then
       If Not Cancelamento Then
            Exit Sub
        End If
    End If
    
    TBLVendas.MovePrevious
    If TBLVendas.BOF Then
        TBLVendas.MoveNext
        Exit Sub
    End If
    
    StatusBarAviso = "Pronto"
    BarraDeStatus StatusBarAviso
    
    Navega��oSuperior lAllowConsult
    TestaInferior TBLVendas, lAllowEdit, lAllowDelete, lAllowConsult
    
    GetRecords
End Sub
Sub PosRecords()
    If TBLVendas.RecordCount = 0 Then
        Exit Sub
    End If
    
    TBLVendas.Seek "=", Val(txtOr�amento)
    If TBLVendas.NoMatch Then
        'MsgBox "N�o consegui encontrar o cliente com CGC/CPF " + txtCgcCpf, vbExclamation, "Erro"
        TBLVendas.MoveFirst
        Navega��oInferior False
        Navega��oInferior lAllowConsult
    Else
        TestaInferior TBLVendas, lAllowEdit, lAllowDelete, lAllowConsult
        TestaSuperior TBLVendas, lAllowEdit, lAllowDelete, lAllowConsult
    End If
    GetRecords
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
    
    txtOr�amento = TBLVendas("C�DIGO")
    mC�digoDoCliente = TBLVendas("C�DIGO DO CLIENTE")
    txtCliente = SearchCliente(mC�digoDoCliente, byCodigo)
    txtValor = FormatStringMask("@V ##.###.##0,00", StrVal(TBLVendas("VALOR TOTAL DA VENDA")))
    txtDesconto = FormatStringMask("@V ##.###.##0,00", StrVal(TBLVendas("DESCONTO TOTAL DA VENDA")))
    mUsu�rioDescontoM�ximo = TBLVendas("AUTORIZOU DESCONTO")
    txtValorBonus = FormatStringMask("@V ##.###.##0,00", StrVal(TBLVendas("VALOR DO BONUS")))
    
    If IsNull(TBLVendas("OBSERVA��O")) Then
        txtObserva��o = Empty
    Else
        txtObserva��o = TBLVendas("OBSERVA��O")
    End If
    
    If TBLVendas("DATA DO OR�AMENTO") <> vbNull Then
        txtData = FormatStringMask(CheckDataMask, TBLVendas("DATA DO OR�AMENTO"))
        CorrigeData DataMask, txtData, TBLVendas("DATA DO OR�AMENTO")
    Else
        txtData = DataNula
    End If
    
    'Calcula valor total
    pValor1 = ValStr(txtValor)
    pValor2 = ValStr(txtDesconto)
    pValor3 = ValStr(txtValorBonus)
    pValor4 = pValor1 - pValor2 - pValor3
    
    txtValorTotal = "R$" + String(6, " ") + FormatStringMask("@V ##.###.##0,00", StrVal(pValor4))
    
    mTotalPagamentos = TBLVendas("QUANTIDADE DE VENCIMENTOS")
    mValorAVista = FormatStringMask("@V ##.###.##0,00", StrVal(pValor4))
    mValorAPrazo = FormatStringMask("@V ##.###.##0,00", StrVal(TBLVendas("VALOR A PRAZO")))
    mTipoDePagamento = TBLVendas("TIPO DE PAGAMENTO")
    
    'Calcula porcentagem de bonus
    If (pValor1 - pValor2) = 0 Then
        pValor4 = 0
    Else
        pValor4 = pValor3 * 100 / (pValor1 - pValor2)
    End If
    txtPorcentagemBonus = FormatStringMask("@V ##0,00", StrVal(pValor4))
    
    FillGrid TBLVendas("C�DIGO")
    FillGridPg TBLVendas("C�DIGO")
    
    lPula = False
    Exit Sub
    
Erro:
    GeraMensagemDeErro "Sa�daVenda - GetRecords"
    lPula = False
    ZeraCampos
    If Not lAllowEdit Then
        DesativaCampos
    End If
    If Not lAllowEdit Then
        DesativaCampos
    End If
End Sub
Private Function SetRecords()
    On Error GoTo ErroVendas
    
    Dim Cont%, Cont1%
    Dim Msg$
    Dim Confirma��o As Integer, Msg1$, Msg2$
    Dim pAlterar As Boolean, pInserir As Boolean
    
    WS.BeginTrans 'Inicia uma Transa��o
        
    If lInserir Then
        TBLVendas.AddNew
    Else
        TBLVendas.Edit
    End If
    
    TBLVendas("C�DIGO") = Val(txtOr�amento)
    TBLVendas("C�DIGO DO CLIENTE") = mC�digoDoCliente
    TBLVendas("TIPO") = "O"
    TBLVendas("VALOR TOTAL DA VENDA") = ValStr(txtValor)
    TBLVendas("DESCONTO TOTAL DA VENDA") = ValStr(txtDesconto)
    TBLVendas("AUTORIZOU DESCONTO") = mUsu�rioDescontoM�ximo
    TBLVendas("VALOR DO BONUS") = ValStr(txtValorBonus)
    TBLVendas("DATA DA VENDA") = vbNull
    TBLVendas("DATA DO OR�AMENTO") = IIf(Trim(StrTran(txtData, "/")) <> Empty, txtData, vbNull)
    TBLVendas("BAIXADO") = False
    TBLVendas("TIPO DE PAGAMENTO") = mTipoDePagamento
    TBLVendas("QUANTIDADE DE VENCIMENTOS") = mTotalPagamentos
    TBLVendas("VALOR A PRAZO") = ValStr(mValorAPrazo)
    TBLVendas("OBSERVA��O") = txtObserva��o
    If lInserir Then
        TBLVendas("USERNAME - CRIA") = mUsu�rio
        TBLVendas("DATA - CRIA") = Date
        TBLVendas("HORA - CRIA") = Time
        TBLVendas("USERNAME - ALTERA") = "VAZIO"
        TBLVendas("DATA - ALTERA") = vbNull
        TBLVendas("HORA - ALTERA") = vbNull
    End If
    If lAlterar Then
        TBLVendas("USERNAME - ALTERA") = mUsu�rio
        TBLVendas("DATA - ALTERA") = Date
        TBLVendas("HORA - ALTERA") = Time
    End If
    TBLVendas.Update
            
    On Error GoTo ErroVendaItens
    
    If lAlterarGrid Or lInserir Then
        DBFinanceiro.Execute "DELETE * FROM [VENDA - ITENS] WHERE [OR�AMENTO] = " & Val(txtOr�amento)

        For Cont = 0 To mTotalRows - 1
            TBLVendasItens.AddNew
            TBLVendasItens("OR�AMENTO") = Val(txtOr�amento)
            TBLVendasItens("C�DIGO DO PRODUTO") = dbgrdItensArray(5, Cont)
            TBLVendasItens("QUANTIDADE") = ValStr(dbgrdItensArray(1, Cont))
            TBLVendasItens("VALOR UNIT�RIO") = ValStr(dbgrdItensArray(2, Cont))
            TBLVendasItens("DESCONTO") = ValStr(dbgrdItensArray(3, Cont))
            TBLVendasItens.Update
        Next
    End If
    
    On Error GoTo ErroVendaLotes
    
    If lAlterarGrid Or lInserir Then
        DBFinanceiro.Execute "DELETE * FROM [VENDA - LOTES] WHERE [OR�AMENTO] = " & Val(txtOr�amento)
        For Cont = 0 To mTotalRows - 1
            For Cont1 = 1 To ArrayLotes(Cont).Count
                TBLVendaLotes.AddNew
                TBLVendaLotes("OR�AMENTO") = Val(txtOr�amento)
                TBLVendaLotes("C�DIGO DO PRODUTO") = dbgrdItensArray(5, Cont)
                TBLVendaLotes("C�DIGO DO LOTE") = ArrayLotes(Cont).GetC�digoDoLote(Cont1)
                TBLVendaLotes("QUANTIDADE") = ArrayLotes(Cont).GetQuantidade(Cont1)
                TBLVendaLotes("M�LTIPLO") = ArrayLotes(Cont).GetM�ltiplo(Cont1)
                TBLVendaLotes.Update
            Next
        Next
    End If
    
    On Error GoTo ErroFormaDePagamento
    
    If lAlterarGridPg Or lInserir Then
        DBFinanceiro.Execute "DELETE * FROM [VENDA - FORMA DE PAGAMENTO] WHERE [OR�AMENTO] = " & Val(txtOr�amento)
        For Cont = 0 To mTotalPagamentos - 1
            TBLFormaDePagamento.AddNew
            TBLFormaDePagamento("OR�AMENTO") = Val(txtOr�amento)
            TBLFormaDePagamento("DOCUMENTO") = FormaDePagamentoArray(0, Cont)
            TBLFormaDePagamento("VENCIMENTO") = IIf(Trim(StrTran(FormaDePagamentoArray(1, Cont), "/")) <> Empty, FormaDePagamentoArray(1, Cont), vbNull)
            TBLFormaDePagamento("VALOR") = StrVal(FormaDePagamentoArray(2, Cont))
            TBLFormaDePagamento.Update
        Next
    End If
       
    WS.CommitTrans 'Grava as altera��es ou inclus�es se n�o houverem erros
    
    If lInserir Then
        Log gUsu�rio, "Inclus�o - Or�amento: " & txtOr�amento
    Else
        Log gUsu�rio, "Altera��o - Or�amento: " & txtOr�amento
    End If
    
    lAlterar = False
    lInserir = False
    lAlterarGrid = False
    lAlterarGridPg = False
    
    SetRecords = True
    
    Exit Function
    
ErroVendas:
    TBLVendas.CancelUpdate
    GeraMensagemDeErro "Sa�daVendas - SetRecords - ErroVendas - " & txtOr�amento, True
    SetRecords = False
    Exit Function
    
ErroVendaItens:
    TBLVendasItens.CancelUpdate
    GeraMensagemDeErro "Sa�daVenda - SetRecords - ErroVendaItens - " & txtOr�amento, True
    SetRecords = False
    Exit Function
    
ErroVendaLotes:
    TBLVendaLotes.CancelUpdate
    GeraMensagemDeErro "Sa�daVenda - SetRecords - ErroVendaLotes - " & txtOr�amento, True
    SetRecords = False
    Exit Function
    
ErroFormaDePagamento:
    TBLFormaDePagamento.CancelUpdate
    GeraMensagemDeErro "Sa�daVenda - SetRecords - ErroFormaDePagamento - " & txtOr�amento, True
    SetRecords = False
    Exit Function
    
End Function
Private Sub ZeraCampos()
    On Error Resume Next
    
    lPula = True
    txtOr�amento = Empty
    txtData = FormatStringMask(CheckDataMask, Date)
    txtValor = FormatStringMask("@V ##.###.##0,00", "0,00")
    txtValorTotal = "R$" & String(6, " ") & FormatStringMask("@V ##.###.##0,00", "0,00")
    txtDesconto = FormatStringMask("@V ##.###.##0,00", "0,00")
    txtValorBonus = FormatStringMask("@V ##.###.##0,00", "0,00")
    txtPorcentagemBonus = FormatStringMask("@V ##0,00", "  0,00")
    txtCliente = Empty
    txtObserva��o = Empty
    mUsu�rioDescontoM�ximo = Empty
    mC�digo = 0
    ReDim dbgrdItensArray(MAXCOLS - 1, 0)
    ReDim FormaDePagamentoArray(MAXCOLSPG - 1, 0)
    ReDim ArrayLotes(0)
    Set ArrayLotes(0) = Nothing
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
        frmFormaDePagamento.lCompra = False
        Set frmFormaDePagamento.ptrForm = Me
        Set frmFormaDePagamento.TBLPlanoDePagamento = TBLPlanoDePagamento
        frmFormaDePagamento.Show 1
    End If
End Sub
Private Sub cmdGravar_Click()
    If mC�digoDoCliente = Empty Then
        MsgBox "O campo CLIENTE n�o est� preenchido !", vbInformation, "Aviso"
        Exit Sub
    End If
    
    'Verifica desconto m�ximo
    If Not DescontoM�ximoOk Then
        Exit Sub
    End If
    
    'Valida Usu�rio
    frmValidaUsu�rio.Show 1
    
    mUsu�rio = frmValidaUsu�rio.Usu�rio
    
    Set frmValidaUsu�rio = Nothing
    
    If mUsu�rio = Empty Then
        Exit Sub
    End If
    
    Caption = "Venda - " & mUsu�rio
    
    Gravar
End Sub
Private Sub cmdTabelaCliente_Click()
    If Not lInserir Then
        lAlterar = True
        StatusBarAviso = "Altera��o"
        BarraDeStatus StatusBarAviso
    End If
    Set frmEncontrar.DBBancoDeDados = DBCadastro
    frmEncontrar.NomeDaJanela = "Cliente"
    frmEncontrar.Mensagem = "Nenhum cliente foi selecionado!"
    frmEncontrar.BancoDeDados = "CADASTRO"
    frmEncontrar.Tabela = "CLIENTE"
    frmEncontrar.Indice = "2"
    frmEncontrar.CampoChave = "C�DIGO"
    frmEncontrar.CampoPreencheLista = "NOME - RAZ�O SOCIAL"
    frmEncontrar.Show vbModal
    mC�digoDoCliente = frmEncontrar.Chave
    txtCliente = frmEncontrar.Nome
    txtCliente.ForeColor = &H80000008
End Sub
Private Sub dbgrdItens_AfterColEdit(ByVal ColIndex As Integer)
    If ColIndex = 0 Then 'Produto
    ElseIf ColIndex = 1 Then 'Quantidade
        If lPula Then
            Exit Sub
        End If
        lPula = True
        FormatMask "@V ######0,00", dbgrdItens
        lPula = False
    ElseIf ColIndex = 2 Then 'Valor Unit�rio
    ElseIf ColIndex = 3 Then 'Desconto
        If lPula Then
            Exit Sub
        End If
        lPula = True
        FormatMask "@V ##0,00", dbgrdItens
        lPula = False
    ElseIf ColIndex = 4 Then 'Valor Total
    End If
End Sub
Private Sub dbgrdItens_AfterDelete()
    AcertaValores
End Sub
Private Sub dbgrdItens_AfterUpdate()
    lFirstColumnEdited = False
    dbgrdItens.Col = 0
    HasLote = Empty
    mC�digoProduto = Empty
    dbgrdItens.Refresh
    AcertaValores
End Sub
Private Sub dbgrdItens_BeforeColEdit(ByVal ColIndex As Integer, ByVal KeyAscii As Integer, Cancel As Integer)
    Dim oldColIndex As Integer
    Dim Valor As String

    If ColIndex = 0 Then 'Produto
        lFirstColumnEdited = True
    ElseIf ColIndex = 1 Then 'Quantidade
        If mC�digoProduto = Empty Then
            mC�digoProduto = dbgrdItensArray(5, dbgrdItens.Row)
        End If
        If HasLote = Empty Then
            HasLote = IIf(SearchAdvancedProduto(mC�digoProduto, vbLote), "T", "F")
        End If
        If HasLote = "T" Then
            Set frmSelecionaLote.MatrizLote = ArrayLotes(dbgrdItens.Row)
            frmSelecionaLote.mC�digoProduto = mC�digoProduto
            frmSelecionaLote.mQuantidade = ValStr(dbgrdItens.Text)
            frmSelecionaLote.Show 1
            mValorLote = FormatStringMask("@V ######0,00", StrVal(frmSelecionaLote.mQuantidade))
            Set frmSelecionaLote = Nothing
        End If
    ElseIf ColIndex = 2 Then 'Valor Unit�rio
        Cancel = 1
    ElseIf ColIndex = 3 Then 'Desconto
        mOldValue = dbgrdItens.Text
        dbgrdItens.Text = Valor
    ElseIf ColIndex = 4 Then 'Valor Total
        Cancel = 1
    ElseIf ColIndex = 5 Then 'C�digo do Produto
        Cancel = 1
        dbgrdItens.Col = 4
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
    ElseIf ColIndex = 3 Then 'Desconto
        DoDesconto ColIndex
    ElseIf ColIndex = 4 Then 'Valor Total
    End If
End Sub
Private Sub dbgrdItens_BeforeDelete(Cancel As Integer)
    Dim pValor As Currency, pDesconto As Currency
    
    If Not lInserir Then
        lAlterar = True
        lAlterarGrid = True
        StatusBarAviso = "Altera��o do Or�amento"
        BarraDeStatus StatusBarAviso
    End If
    
    dbgrdItens.Col = 4
    pValor = ValStr(dbgrdItens.Text)
    
    dbgrdItens.Col = 3
    pDesconto = ValStr(dbgrdItens.Text)
    
    pDesconto = pValor * (pDesconto / 100)
    pDesconto = StrVal(FormatStringMask("@V ##.###.##0,00", StrVal(pDesconto)))
    pDesconto = ValStr(txtDesconto) - pDesconto
    
    pValor = ValStr(txtValor) - pValor
    
    lPula = True
'    txtDesconto = FormatStringMask("@V ##.###.##0,00", StrVal(pDesconto))
'    txtValor = FormatStringMask("@V ##.###.##0,00", StrVal(pValor))
    lPula = False
    
    AdelLote dbgrdItens.Row
End Sub
Private Sub dbgrdItens_BeforeInsert(Cancel As Integer)
    ReDim Preserve ArrayLotes(0 To mTotalRows)
    
    Set ArrayLotes(mTotalRows) = New ClassLote
End Sub
Private Sub dbgrdItens_Change()
    If Not lInserir Then
        lAlterar = True
        lAlterarGrid = True
        StatusBarAviso = "Altera��o do Or�amento"
        BarraDeStatus StatusBarAviso
    End If
    If dbgrdItens.Col = 0 Then      'Produto
    ElseIf dbgrdItens.Col = 1 Then  'Quantidade
        If lPula Then
            Exit Sub
        End If
        
        If HasLote = "T" Then
            dbgrdItens = mValorLote
        Else
            FormatMask "@K 9999999,99", dbgrdItens
        End If
    ElseIf dbgrdItens.Col = 2 Then  'Valor Unit�rio
    ElseIf dbgrdItens.Col = 3 Then  'Desconto
        FormatMask "@K ##0,00", dbgrdItens
    ElseIf dbgrdItens.Col = 4 Then   'Valor total
    End If
End Sub
Private Sub dbgrdItens_ColResize(ByVal ColIndex As Integer, Cancel As Integer)
    Cancel = 1
End Sub
Private Sub dbgrdItens_GotFocus()
    dbgrdItens.Col = 0
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
    If Not VendasAberto Then
        Unload Me
        Exit Sub
    End If
    If Not VendasItensAberto Then
        Unload Me
        Exit Sub
    End If
    If Not VendaLotesAberto Then
        Unload Me
        Exit Sub
    End If
    If Not Par�metrosAberto Then
        Unload Me
        Exit Sub
    End If
    If Not LoteDoProdutoAberto Then
        Unload Me
        Exit Sub
    End If
    
    TestaInferior TBLVendas, lAllowEdit, lAllowDelete, lAllowConsult
    TestaSuperior TBLVendas, lAllowEdit, lAllowDelete, lAllowConsult
    
    If TBLVendas.RecordCount = 0 Then
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
            txtCliente.SetFocus
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
    dbgrdItens.ReBind
    dbgrdItens.Refresh

    If lAtualizar Then
        Bot�oAtualizar True
    Else
        Bot�oAtualizar False
    End If
    
    If lGoToInsert Then
        Incluir
    End If
End Sub
Private Sub Form_Deactivate()
    cmdGravar.Enabled = False
    cmdCancelar.Enabled = False
    Bot�oImprimir False
End Sub
Private Sub Form_Load()
    On Error GoTo Erro
    
    Dim Cont%
    
    Top = 0
    Left = 0
    
    lAllowInsert = Allow("VENDA", "I")
    lAllowEdit = Allow("VENDA", "A")
    lAllowDelete = Allow("VENDA", "E")
    lAllowConsult = Allow("VENDA", "C")
    lGoToInsert = IIf(gUsu�rio <> "ADMIN", Allow("VENDA", "U"), False)

    ZeraCampos
    
    lFirstColumnEdited = False
    lInserir = False
    lAlterar = False
    lAlterarGrid = False
    lAlterarGridPg = False
    lInicio = True
    
    VendasAberto = AbreTabela(Dicion�rio, "FINANCEIRO", "VENDA", DBFinanceiro, TBLVendas, TBLTabela, dbOpenTable)
    
    If VendasAberto Then
        IndiceVendasAtivo = "VENDA1"
        TBLVendas.Index = IndiceVendasAtivo
    Else
        MsgBox "N�o consegui abrir a tabela 'Vendas' !", vbCritical, "Erro"
        Exit Sub
    End If
    
    VendasItensAberto = AbreTabela(Dicion�rio, "FINANCEIRO", "VENDA - ITENS", DBFinanceiro, TBLVendasItens, TBLTabela, dbOpenTable)
    
    If VendasItensAberto Then
        IndiceVendasItensAtivo = "VENDAITENS1"
        TBLVendasItens.Index = IndiceVendasItensAtivo
    Else
        MsgBox "N�o consegui abrir a tabela 'Itens de Venda' !", vbCritical, "Erro"
        Exit Sub
    End If
    
    VendaLotesAberto = AbreTabela(Dicion�rio, "FINANCEIRO", "VENDA - LOTES", DBFinanceiro, TBLVendaLotes, TBLTabela, dbOpenTable)
    
    If VendaLotesAberto Then
        IndiceVendaLotesAtivo = "VENDALOTES3"
        TBLVendaLotes.Index = IndiceVendaLotesAtivo
    Else
        MsgBox "N�o consegui abrir a tabela 'Venda - Lotes' !", vbCritical, "Erro"
        Exit Sub
    End If
    
    FormaDePagamentoAberto = AbreTabela(Dicion�rio, "FINANCEIRO", "VENDA - FORMA DE PAGAMENTO", DBFinanceiro, TBLFormaDePagamento, TBLTabela, dbOpenTable)
    
    If FormaDePagamentoAberto Then
        IndiceFormaDePagamentoAtivo = "VENDAFORMADEPAGAMENTO2"
        TBLFormaDePagamento.Index = IndiceFormaDePagamentoAtivo
    Else
        MsgBox "N�o consegui abrir a tabela 'Forma de Pagamento - Venda' !", vbCritical, "Erro"
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
    
    LoteDoProdutoAberto = AbreTabela(Dicion�rio, "CADASTRO", "LOTE DO PRODUTO", DBCadastro, TBLLoteDoProduto, TBLTabela, dbOpenTable)
    
    If LoteDoProdutoAberto Then
        IndiceLoteDoProdutoAtivo = "LOTEDOPRODUTO2"
        TBLLoteDoProduto.Index = IndiceLoteDoProdutoAtivo
    Else
        MsgBox "N�o consegui abrir a tabela 'C�digo do Produto' !", vbCritical, "Erro"
        Exit Sub
    End If
    
    dbgrdItens.Columns.Add 1
    dbgrdItens.Columns.Add 1
    dbgrdItens.Columns.Add 1
    dbgrdItens.Columns.Add 1
    
    For Cont = 0 To dbgrdItens.Columns.Count - 1
        dbgrdItens.Columns(Cont).Visible = True
    Next
       
    dbgrdItens.Columns(0).Caption = "Produto"
    dbgrdItens.Columns(0).Width = 3045
    dbgrdItens.Columns(0).DefaultValue = " "
    dbgrdItens.Columns(0).Alignment = dbgLeft
    
    dbgrdItens.Columns(1).Caption = "Quantidade"
    dbgrdItens.Columns(1).Width = 1000
    dbgrdItens.Columns(1).DefaultValue = "0"
    dbgrdItens.Columns(1).Alignment = dbgRight
    
    dbgrdItens.Columns(2).Caption = "Valor Unit�rio"
    dbgrdItens.Columns(2).Width = 1910
    dbgrdItens.Columns(2).DefaultValue = "0,00"
    dbgrdItens.Columns(2).Alignment = dbgRight
    
    dbgrdItens.Columns(3).Caption = "Desconto"
    dbgrdItens.Columns(3).Width = 1000
    dbgrdItens.Columns(3).DefaultValue = "0,00"
    dbgrdItens.Columns(3).Alignment = dbgRight
    
    dbgrdItens.Columns(4).Caption = "Valor Total"
    dbgrdItens.Columns(4).Width = 1910
    dbgrdItens.Columns(4).DefaultValue = "0,00"
    dbgrdItens.Columns(4).Alignment = dbgRight
    
    dbgrdItens.Columns(5).Caption = "" 'C�digo do Produto
    dbgrdItens.Columns(5).Width = 1
    dbgrdItens.Columns(5).DefaultValue = "0"
    
    dbgrdItens.ReBind
    
    Bot�oIncluir lAllowInsert
 
    If TBLVendas.RecordCount = 0 Then
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
        
    If TBLVendas.RecordCount = 0 Or TBLVendas.RecordCount = 1 Then
        Navega��oSuperior False
    Else
        Navega��oInferior lAllowConsult
    End If

    StatusBarAviso = "Pronto"
    
    Relat�rio = AddPath(Aplica��oPath, "REPORT\VENDAS.RPT")
    TotalDatabaseName = 1
    DataBaseName(1) = AddPath(Aplica��oPath, "DATABASE\FINANCEIRO.MDB")
    mFechar = False
    
    Exit Sub
    
Erro:
    GeraMensagemDeErro "Sa�daVenda - Load"
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
    If VendasAberto Then
        TBLVendas.Close
    End If
    If VendasItensAberto Then
        TBLVendasItens.Close
    End If
    If VendaLotesAberto Then
        TBLVendaLotes.Close
    End If
    If PlanoDePagamentoAberto Then
        TBLPlanoDePagamento.Close
    End If
    If FormaDePagamentoAberto Then
        TBLFormaDePagamento.Close
    End If
    If LoteDoProdutoAberto Then
        TBLLoteDoProduto.Close
    End If
    If Forms.Count = 2 Then
        AllBot�es False
    End If
End Sub
Private Sub txtCliente_Change()
    FormatMask "@!S40", txtCliente
End Sub
Private Sub txtCliente_KeyPress(KeyAscii As Integer)
    txtCliente.ForeColor = &HFF&
    mC�digoDoCliente = Empty
    If Not lInserir Then
        lAlterar = True
        StatusBarAviso = "Altera��o"
        BarraDeStatus StatusBarAviso
    End If
End Sub
Private Sub txtData_Change()
    If Not lPula Then
        lPula = True
        FormatMask DataMask, txtData
        lPula = False
    End If
End Sub
Private Sub txtData_KeyPress(KeyAscii As Integer)
    If Not lInserir Then
        lAlterar = True
        StatusBarAviso = "Altera��o"
        BarraDeStatus StatusBarAviso
    End If
End Sub
Private Sub txtData_LostFocus()
    If StrTran(txtData.Text, "/") <> Space(8) Then
        lPula = True
        CorrigeData DataMask, txtData, Date
        lPula = False
        If Not FormatMask(CheckDataMask, txtData) Then
            Beep
            MsgBox "Data inv�lida !", vbCritical, "Erro"
            txtData.SelStart = 0
            txtData.SetFocus
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
        StatusBarAviso = "Altera��o do Or�amento"
        BarraDeStatus StatusBarAviso
    End If
End Sub
Private Sub txtDesconto_LostFocus()
    If lPula Then
        Exit Sub
    End If
    
    lPula = True
    FormatMask "@V ##.###.##0,00", txtDesconto
    lPula = False
End Sub
Private Sub txtObserva��o_KeyPress(KeyAscii As Integer)
    If Not lInserir Then
        lAlterar = True
        StatusBarAviso = "Altera��o do Or�amento"
        BarraDeStatus StatusBarAviso
    End If
End Sub
Private Sub txtPorcentagemBonus_Change()
    If Not lPula Then
        FormatMask "@K 999,99", txtPorcentagemBonus
        mDigitPorcent = True
    End If
End Sub
Private Sub txtPorcentagemBonus_KeyPress(KeyAscii As Integer)
    If Not lInserir Then
        lAlterar = True
        StatusBarAviso = "Altera��o do Or�amento"
        BarraDeStatus StatusBarAviso
    End If
End Sub
Private Sub txtPorcentagemBonus_LostFocus()
    Dim pValor1 As Currency
    Dim pValor2 As Currency
    Dim pValor3 As Currency
    
    If Not mDigitPorcent Then
        Exit Sub
    End If
    
    If lPula Then
        Exit Sub
    End If
    
    mDigitPorcent = False
    
    lPula = True
    FormatMask "@V ##0,00", txtPorcentagemBonus
    lPula = False
    
    'Atualiza Valor do Bonus
    pValor1 = ValStr(txtValor)
    pValor2 = ValStr(txtDesconto)
    pValor2 = pValor1 - pValor2
    pValor3 = ValStr(txtPorcentagemBonus)
    pValor3 = pValor2 * (pValor3 / 100)
    pValor2 = pValor2 - pValor3
    
    lPula = True
    txtValorBonus = FormatStringMask("@V ##.###.##0,00", StrVal(pValor3))
    txtValorTotal = "R$" + String(6, " ") & FormatStringMask("@V ##.###.##0,00", StrVal(pValor2))
    lPula = False
End Sub
Private Sub txtValor_Change()
    If Not lPula Then
        FormatMask "@K 99.999.999,99", txtValor
    End If
End Sub
Private Sub txtValor_KeyPress(KeyAscii As Integer)
    If Not lInserir Then
        lAlterar = True
        StatusBarAviso = "Altera��o do Or�amento"
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
Private Sub DoDesconto(ByVal ColIndex As Integer)
    Dim pC�digo  As String
    Dim pszValor As String
    Dim pValor   As Currency
    Dim pValor1  As Currency
    Dim pValor2  As Currency
    Dim pValor3  As Currency
    
    pszValor = FormatStringMask("@V ##0,00", dbgrdItens.Text)
    
    dbgrdItens.Col = 5
    pC�digo = Val(dbgrdItens.Text)
    
    pValor = SearchAdvancedProduto(pC�digo, vbDescontoM�ximo, vbIndice2)
    
    dbgrdItens.Col = 3
    If ValStr(pszValor) > pValor Then
        pszValor = FormatStringMask("@V ##0,00", StrVal(pValor))
        MsgBox "Valor do desconto � maior que o" & vbCr & "valor permitido para este produto!" & vbCr & vbCr & "O valor m�ximo �: " & pszValor, vbInformation, "Desconto Negado !"
    End If
    
    dbgrdItens.Text = pszValor
    
    dbgrdItens.Col = 4
    pValor = ValStr(dbgrdItens.Text)
    dbgrdItens.Col = 3
    
    If mOldValue = Empty Then
        pValor1 = 0
    Else
        pValor1 = ValStr(mOldValue)
        pValor1 = (pValor1 / 100) * pValor
        pValor1 = StrVal(FormatStringMask("@V ##.###.##0,00", StrVal(pValor1)))
    End If
    
    'Atualiza desconto
    dbgrdItens.Text = pszValor
    pValor1 = ValStr(txtDesconto) - pValor1
    pValor2 = ValStr(dbgrdItens.Text)
    pValor2 = (pValor2 / 100) * pValor
    pValor2 = pValor1 + pValor2
    pszValor = StrVal(pValor2)
    lPula = True
    'txtDesconto = FormatStringMask("@V ##.###.##0,00", pszValor)
    lPula = False
    
    'Atualiza Porcentagem do Bonus
    pValor1 = ValStr(txtValor)
    pValor2 = ValStr(txtDesconto)
    pValor3 = ValStr(txtValorBonus)
    If (pValor1 - pValor2) = 0 Then
        pValor3 = 0
    Else
        pValor3 = pValor3 * 100 / (pValor1 - pValor2)
    End If
    lPula = True
    'txtPorcentagemBonus = FormatStringMask("@V ##0,00", StrVal(pValor3))
    lPula = False
    
    'Atualiza Valor Total
    pValor1 = ValStr(txtValor)
    pValor2 = ValStr(txtValorBonus)
    pValor3 = ValStr(txtDesconto)
    pValor2 = pValor1 - (pValor2 + pValor3)
    pszValor = StrVal(pValor2)
    lPula = True
    'txtValorTotal = "R$" + String(6, " ") + FormatStringMask("@V ##.###.##0,00", pszValor)
    lPula = False
End Sub
Private Sub DoQuantidade(ByVal ColIndex As Integer)
    Dim pC�digo As String
    Dim pQuantidade As Single
    Dim pszValor As String
    Dim pValor As Currency
    Dim pValor1 As Currency
    Dim pValor2 As Currency
    Dim pValor3 As Currency
    Dim pOldValor As Currency
    Dim pNewValor As Currency
    
    pQuantidade = ValStr(dbgrdItens.Text)
    
    dbgrdItens.Col = 5
    pC�digo = Val(dbgrdItens.Text)
    
    dbgrdItens.Col = 2
    pValor = ValStr(dbgrdItens.Text) * pQuantidade
    
    dbgrdItens.Col = 4
    pOldValor = ValStr(dbgrdItens.Text)
    pValor1 = ValStr(txtValor) - ValStr(dbgrdItens.Text)
    pValor2 = SearchAdvancedProduto(pC�digo, vbValValorUnit�rio, vbIndice2)
    pValor2 = pValor2 * pQuantidade
    pValor2 = pValor1 + pValor2
    pszValor = StrVal(pValor2)
    
    'Atualiza campo Valor
    lPula = True
    'txtValor = FormatStringMask("@V ##.###.##0,00", pszValor)
    lPula = False
    
    dbgrdItens.Text = FormatStringMask("@V ##.###.##0,00", pValor)
    pNewValor = StrVal(dbgrdItens.Text)
       
    'Atualiza valor de desconto
    dbgrdItens.Col = 3
    mOldValue = dbgrdItens.Text
'    If mOldValue = "" Then mOldValue = 0
'    AcertaDesconto pOldValor, (ValStr(mOldValue) / 100), pNewValor
    
'    'Atualiza campo Valor Total
'    pValor1 = ValStr(txtValor)
'    pValor2 = ValStr(txtValorBonus)
'    pValor3 = ValStr(txtDesconto)
'    pValor2 = pValor1 - (pValor2 + pValor3)
'    pszValor = StrVal(pValor2)
'    lPula = True
'    txtValorTotal = "R$" + String(6, " ") + FormatStringMask("@V ##.###.##0,00", pszValor)
'    lPula = False
    
    dbgrdItens.Col = ColIndex
    dbgrdItens.Text = FormatStringMask("@V ######0,00", StrVal(pQuantidade))
    
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
    Dim pfrmEncontraProduto As New frmEncontraProduto
    Dim Cont As Byte
    
    pC�digo = UCase(dbgrdItens.Text) 'C�digo digitado pelo usu�rio
    plgC�digo = Val(SearchAdvancedProduto(pC�digo, vbC�digo)) 'Retorna o c�digo do produto, se encontrar
               
    dbgrdItens.Col = 5
    dbgrdItens.Text = plgC�digo
        
    If plgC�digo = 0 Then 'Se o c�digo for igual a zero, significa que o produto n�o existe. Abre uma janela de consulta
        pfrmEncontraProduto.TipoDeBusca = mTipoDeBusca
        pfrmEncontraProduto.Crit�rioDeBusca = mCrit�rioDeBusca
        pfrmEncontraProduto.Condi��oSQL = mCondi��oSQL
        pfrmEncontraProduto.Show 1
        mTipoDeBusca = pfrmEncontraProduto.TipoDeBusca
        mCrit�rioDeBusca = pfrmEncontraProduto.Crit�rioDeBusca
        mCondi��oSQL = pfrmEncontraProduto.Condi��oSQL
        
        dbgrdItens.Col = 5
        dbgrdItens.Text = pfrmEncontraProduto.C�digo
        plgC�digo = Val(pfrmEncontraProduto.C�digo)
    End If

    'Verifica se o item j� est� cadastrado na venda, para n�o repit�-lo
    If mTotalRows > 0 Then
        For Cont = 0 To mTotalRows - 1
            If dbgrdItensArray(MAXCOLS - 1, Cont) = plgC�digo Then
                MsgBox "O item j� foi inclu�do na tabela!", vbInformation, "Aviso"
                DoProduto = False
                Exit Function
            End If
        Next
    End If
    
    'Verifica se est� dividido por Lote
    If SearchAdvancedProduto(plgC�digo, vbLote) Then
        HasLote = "T"
    Else
        HasLote = "F"
    End If
    
    dbgrdItens.Col = 2 'Valor Unit�rio
    dbgrdItens.Text = SearchAdvancedProduto(plgC�digo, vbValorUnit�rio, vbIndice2)
    
    dbgrdItens.Col = 1 'Quantidade do Produto
    dbgrdItens.Text = ""
    If dbgrdItens.Text = "" Then
        If HasLote = "F" Then
            dbgrdItens.Text = "1,00"
            pValor1 = SearchAdvancedProduto(plgC�digo, vbValValorUnit�rio, vbIndice2)
        Else
            dbgrdItens.Text = "0,00"
            pValor1 = 0
        End If
        'Atualiza a coluna de Valor Total
        dbgrdItens.Col = 4
        dbgrdItens.Text = FormatStringMask("@V ##.###.##0,00", StrVal(pValor1))
        
        pValor2 = ValStr(txtValor)
        pValor2 = pValor1 + pValor2
        pszValor = StrVal(pValor2)
        
'        lPula = True
'        txtValor = FormatStringMask("@V ##.###.##0,00", pszValor)
'        txtValorTotal = "R$" + String(6, " ") + FormatStringMask("@V ##.###.##0,00", pszValor)
'        lPula = False
    Else
        'Corrige o valor total
        pQuantidade = Val(dbgrdItens.Text)
        pOldValor = ValStr(dbgrdItens.Text)
        pValor1 = ValStr(txtValor) - ValStr(dbgrdItens.Text)
        pValor2 = SearchAdvancedProduto(plgC�digo, vbValValorUnit�rio) * pQuantidade
        pValor2 = pValor1 + pValor2
        pszValor = StrVal(pValor2)
        
'        lPula = True
'        txtValor = FormatStringMask("@V ##.###.##0,00", pszValor)
'        txtValorTotal = "R$" + String(6, " ") + FormatStringMask("@V ##.###.##0,00", pszValor)
'        lPula = False
        'dbgrdItens.Col = 4
        'dbgrdItens.Text = FormatStringMask("@V ##.###.##0,00", (SearchAdvancedProduto(plgC�digo, vbValValorUnit�rio) * pQuantidade))
        'pNewValor = ValStr(dbgrdItens.Text)
'
'        'Corrige o valor do desconto
        dbgrdItens.Col = 3
        mOldValue = dbgrdItens.Text
'        AcertaDesconto pOldValor, (ValStr(mOldValue) / 100), pNewValor
    End If
    
    mC�digoProduto = plgC�digo
    
    'Retorna a descri��o do produto na primeira coluna
    dbgrdItens.Col = ColIndex
    dbgrdItens.Text = SearchAdvancedProduto(plgC�digo, vbDescri��o)
    
    DoProduto = True
End Function
Private Sub ZeraDescontoPorItem()
    Dim Cont%
    
    For Cont = 0 To UBound(dbgrdItensArray, 2)
        dbgrdItensArray(3, Cont) = "0,00"
    Next
    
    dbgrdItens.Refresh
End Sub
Private Sub RefazDescontoPorItem()
    Dim Cont%
    Dim Quantidade%
    Dim Valor As Currency
    Dim Desconto As Currency
    Dim ValorDesconto As Currency
    Dim ValorDescontoTotal As Currency
    
    ValorDescontoTotal = 0
    
    For Cont = 0 To UBound(dbgrdItensArray, 2)
        Quantidade = StrVal(dbgrdItensArray(1, Cont))
        Valor = StrVal(dbgrdItensArray(2, Cont))
        Desconto = StrVal(dbgrdItensArray(3, Cont))
        
        Valor = Valor * Quantidade
        ValorDesconto = (Desconto / 100) * Valor
        
        ValorDescontoTotal = ValorDescontoTotal + ValorDesconto
    Next
    
    lPula = True
    'txtDesconto = FormatStringMask("@V ##.###.##0,00", ValStr(ValorDescontoTotal))
    lPula = False
End Sub
Private Sub txtValorBonus_Change()
    If Not lPula Then
        FormatMask "@K 99.999.999,99", txtValorBonus
    End If
End Sub
Private Sub txtValorBonus_KeyPress(KeyAscii As Integer)
    If Not lInserir Then
        lAlterar = True
        StatusBarAviso = "Altera��o do Or�amento"
        BarraDeStatus StatusBarAviso
    End If
End Sub
Private Sub txtValorBonus_LostFocus()
    Dim pValor1 As Currency
    Dim pValor2 As Currency
    Dim pValor3 As Currency
       
    If lPula Then
        Exit Sub
    End If
    
    lPula = True
    FormatMask "@V ##.###.##0,00", txtValorBonus
    lPula = False
    
    'Atualiza valor total
    pValor1 = ValStr(txtValor)
    pValor2 = ValStr(txtDesconto)
    pValor3 = ValStr(txtValorBonus)
    pValor3 = pValor1 - (pValor2 + pValor3)
    lPula = True
    txtValorTotal = "R$" + String(6, " ") + FormatStringMask("@V ##.###.##0,00", StrVal(pValor3))
    lPula = False
    
    'Atualiza campo de porcentagem
    pValor1 = ValStr(txtValor)
    pValor2 = ValStr(txtDesconto)
    pValor2 = pValor1 - pValor2
    pValor3 = ValStr(txtValorBonus)
    If pValor2 = 0 Then
        pValor3 = 0
    Else
        pValor3 = pValor3 * 100 / pValor2
    End If
    lPula = True
    txtPorcentagemBonus = FormatStringMask("@V ##0,00", StrVal(pValor3))
    lPula = False
End Sub
