VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form frmCaixaF�cil 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Caixa F�cil"
   ClientHeight    =   6525
   ClientLeft      =   1215
   ClientTop       =   1245
   ClientWidth     =   9540
   Icon            =   "CaixaFacil.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6525
   ScaleWidth      =   9540
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   345
      Left            =   8280
      TabIndex        =   23
      Top             =   6165
      Width           =   1245
   End
   Begin VB.Frame frDadosCadastrais 
      Height          =   1140
      Left            =   0
      TabIndex        =   15
      Top             =   0
      Width           =   9540
      Begin VB.TextBox txtDataVenda 
         Height          =   285
         Left            =   8250
         TabIndex        =   22
         Text            =   "  /  /"
         Top             =   690
         Width           =   990
      End
      Begin VB.TextBox txtCliente 
         Height          =   300
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   690
         Width           =   5235
      End
      Begin VB.TextBox txtDataOr�amento 
         Height          =   285
         Left            =   8250
         TabIndex        =   16
         Text            =   "  /  /"
         Top             =   240
         Width           =   990
      End
      Begin VB.TextBox txtOr�amento 
         Height          =   285
         Left            =   1200
         TabIndex        =   0
         Top             =   300
         Width           =   765
      End
      Begin VB.Label lblDataVenda 
         Caption         =   "Data da Venda"
         Height          =   405
         Left            =   7380
         TabIndex        =   21
         Top             =   630
         Width           =   765
      End
      Begin VB.Label lblCliente 
         Caption         =   "Cliente"
         Height          =   180
         Left            =   150
         TabIndex        =   20
         Top             =   720
         Width           =   645
      End
      Begin VB.Label lblDataOr�amento 
         Caption         =   "Data do Or�amento"
         Height          =   390
         Left            =   7350
         TabIndex        =   19
         Top             =   180
         Width           =   795
      End
      Begin VB.Label lblOr�amento 
         Caption         =   "Or�amento"
         Height          =   180
         Left            =   150
         TabIndex        =   18
         Top             =   330
         Width           =   825
      End
   End
   Begin VB.Frame frItens 
      Caption         =   " Itens "
      Height          =   3255
      Left            =   0
      TabIndex        =   13
      Top             =   1140
      Width           =   9540
      Begin MSDBGrid.DBGrid dbgrdItens 
         Height          =   2925
         Left            =   60
         OleObjectBlob   =   "CaixaFacil.frx":030A
         TabIndex        =   14
         Top             =   210
         Width           =   9405
      End
   End
   Begin VB.CommandButton cmdGravar 
      Caption         =   "&Localizar"
      Height          =   345
      Left            =   6990
      TabIndex        =   12
      Top             =   6165
      Width           =   1245
   End
   Begin VB.CommandButton cmdFormaDePagamento 
      Caption         =   "&Forma de Pagemanto"
      Height          =   345
      Left            =   45
      TabIndex        =   11
      Top             =   6165
      Width           =   1980
   End
   Begin VB.Frame frTotais 
      Caption         =   "Totais "
      Height          =   1695
      Left            =   0
      TabIndex        =   1
      Top             =   4410
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
         TabIndex        =   6
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
         Locked          =   -1  'True
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
         TabIndex        =   4
         Text            =   "R$"
         Top             =   1320
         Width           =   2655
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
         TabIndex        =   3
         Text            =   "         0,00"
         Top             =   930
         Width           =   1665
      End
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
         TabIndex        =   2
         Text            =   "  0,00"
         Top             =   930
         Width           =   855
      End
      Begin VB.Label lblBonus 
         Caption         =   "Bonus"
         Height          =   195
         Left            =   6180
         TabIndex        =   10
         Top             =   990
         Width           =   495
      End
      Begin VB.Label lblDesconto 
         Caption         =   "Desconto"
         Height          =   255
         Left            =   6930
         TabIndex        =   9
         Top             =   630
         Width           =   1065
      End
      Begin VB.Label lblTotalGeral 
         Caption         =   "Total do Or�amento"
         Height          =   225
         Left            =   5280
         TabIndex        =   8
         Top             =   1350
         Width           =   1425
      End
      Begin VB.Label lblSubTotal 
         Caption         =   "Sub Total"
         Height          =   255
         Left            =   6930
         TabIndex        =   7
         Top             =   240
         Width           =   705
      End
   End
End
Attribute VB_Name = "frmCaixaF�cil"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const MAXCOLS = 6
Const MAXCOLSPG = 3

Dim mRecno%
Dim mTotalRows%
Dim dbgrdItensArray() As String

Dim mNF As String
Dim mTotalDeNotas As Byte

Dim mRecnoPg%
Dim mTotalPagamentos As Integer
Dim mValorAVista As String
Dim mValorAPrazo As String
Dim mTipoDePagamento As Long
Dim mPrimeiroPagamentoBaixado As Boolean
Dim FormaDePagamentoArray() As String

Dim lPula As Boolean
Dim mDigitBonus As Boolean
Dim mDigitPorcent As Boolean
Dim mlRefazDesconto As Boolean
Dim lFechar As Boolean

Dim mC�digo As Integer
Dim mOldValue As String

Dim mC�digoDoCliente As String

Dim TBLVendas As Table
Dim VendasAberto As Boolean
Dim IndiceVendasAtivo$

Dim TBLVendasItens As Table
Dim VendasItensAberto As Boolean
Dim IndiceVendasItensAtivo$

Dim TBLVendasLote As Table
Dim VendasLoteAberto As Boolean
Dim IndiceVendasLotesAtivo$

Dim TBLPar�metros As Table
Dim Par�metrosAberto As Boolean

Dim TBLFormaDePagamento As Table
Dim FormaDePagamentoAberto As Boolean
Dim IndiceFormaDePagamentoAtivo$

Dim TBLPlanoDePagamento As Table
Dim PlanoDePagamentoAberto As Boolean
Dim IndicePlanoDePagamentoAtivo$

Dim ArrayFormaDePagamentoRecno() As Variant
Dim ArrayVendasItensRecno() As Variant

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
End Sub
Private Sub DesativaCampos()
    txtDataOr�amento.Enabled = False
    txtDataVenda.Enabled = False
    frItens.Enabled = False
    frTotais.Enabled = False
    cmdFormaDePagamento.Enabled = False
End Sub
Public Function ExcluirPagamentos() As Boolean
    On Error GoTo Erro
    
    Dim Recno As Byte, Cont As Byte
    
    WS.BeginTrans
    
    Recno = mRecnoPg - 1
    For Cont = 0 To Recno
        TBLFormaDePagamento.Bookmark = ArrayFormaDePagamentoRecno(Cont)
        TBLFormaDePagamento.Delete
    Next
    
    WS.CommitTrans
    
    ExcluirPagamentos = True
    
    FillGridPg TBLVendas("C�DIGO")
    
    Exit Function
    
Erro:
    GeraMensagemDeErro "CaixaF�cil - ExcluirPagamentos - Or�amento: " & txtOr�amento, True
    ExcluirPagamentos = False
End Function
Private Sub FillGrid(ByVal Chave As Long)
    dbgrdItens.ReBind
    
    ReDim dbgrdItensArray(MAXCOLS - 1, 0)
    ReDim ArrayVendasItensRecno(0)
    
    mTotalRows = 0
    mRecno = 0
    
    TBLVendasItens.Seek "=", Chave
    If Not TBLVendasItens.NoMatch Then
        Do While Not TBLVendasItens.EOF And TBLVendasItens("OR�AMENTO") = Chave
            mRecno = mRecno + 1
            mTotalRows = mTotalRows + 1
            ReDim Preserve dbgrdItensArray(MAXCOLS - 1, mTotalRows - 1)
            ReDim Preserve ArrayVendasItensRecno(mTotalRows - 1)
            
            ArrayVendasItensRecno(mTotalRows - 1) = TBLVendasItens.Bookmark
            dbgrdItensArray(0, mTotalRows - 1) = SearchProduto(TBLVendasItens("C�DIGO DO PRODUTO")) 'Nome do Produto
            dbgrdItensArray(1, mTotalRows - 1) = FormatStringMask("@V ######0", StrVal(TBLVendasItens("QUANTIDADE"))) 'Quantidade
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
    ReDim ArrayFormaDePagamentoRecno(0)
    
    mTotalPagamentos = 0
    mRecnoPg = 0
    
    TBLFormaDePagamento.Seek "=", Chave
    If Not TBLFormaDePagamento.NoMatch Then
        Do While Not TBLFormaDePagamento.EOF And TBLFormaDePagamento("OR�AMENTO") = Chave
            mRecnoPg = mRecnoPg + 1
            mTotalPagamentos = mTotalPagamentos + 1
            ReDim Preserve FormaDePagamentoArray(MAXCOLSPG - 1, mTotalPagamentos - 1)
            ReDim Preserve ArrayFormaDePagamentoRecno(mTotalPagamentos - 1)
            
            ArrayFormaDePagamentoRecno(mTotalPagamentos - 1) = TBLFormaDePagamento.Bookmark
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
    If SetRecords Then
        cmdGravar.Caption = "&Localizar"
        ZeraCampos
        DesativaCampos
        StatusBarAviso = "Venda aceita"
    Else
        StatusBarAviso = "Ocorreu uma falha"
    End If
    
    BarraDeStatus StatusBarAviso
    
    If txtOr�amento.Enabled Then
        txtOr�amento.SetFocus
    End If
End Sub
Public Function GetPagamentos(ByVal Coluna As Integer, ByVal Linha As Integer) As String
    GetPagamentos = FormaDePagamentoArray(Coluna, Linha)
End Function
Public Function GravaPagamento(ByRef Matriz() As String) As Boolean
    Dim Recno As Byte
    Dim Cont As Byte
    
    mTotalPagamentos = frmFormaDePagamento.mTotalPagamentos
    mTipoDePagamento = frmFormaDePagamento.mTipoDePagamento
    mValorAPrazo = frmFormaDePagamento.mValorAPrazo
    
    Recno = 0
    
    Recno = mRecnoPg - 1
    
    WS.BeginTrans
    
    On Error GoTo ErroVendas
    TBLVendas.Edit
    TBLVendas("TIPO DE PAGAMENTO") = mTipoDePagamento
    TBLVendas("QUANTIDADE DE VENCIMENTOS") = mTotalPagamentos
    TBLVendas("VALOR A PRAZO") = ValStr(mValorAPrazo)
    TBLVendas("USERNAME - ALTERA") = gUsu�rio
    TBLVendas("DATA - ALTERA") = Date
    TBLVendas("HORA - ALTERA") = Time
    TBLVendas.Update
    
    On Error GoTo ErroFormaDePagamento
    For Cont = 0 To mTotalPagamentos - 1
        If Cont + 1 <= mRecnoPg Then
            TBLFormaDePagamento.Bookmark = ArrayFormaDePagamentoRecno(Cont)
            TBLFormaDePagamento.Edit
        Else
            TBLFormaDePagamento.AddNew
        End If
        
        TBLFormaDePagamento("OR�AMENTO") = Val(txtOr�amento)
        TBLFormaDePagamento("DOCUMENTO") = Matriz(0, Cont)
        TBLFormaDePagamento("VENCIMENTO") = IIf(Trim(StrTran(Matriz(1, Cont), "/")) <> Empty, Matriz(1, Cont), vbNull)
        TBLFormaDePagamento("VALOR") = StrVal(Matriz(2, Cont))
        TBLFormaDePagamento.Update
    Next
    If Cont <= Recno Then
        mTotalPagamentos = Cont
        For Cont = mTotalPagamentos To Recno
            TBLFormaDePagamento.Bookmark = ArrayFormaDePagamentoRecno(Cont)
            TBLFormaDePagamento.Delete
        Next
    End If
    
    WS.CommitTrans
    
    GravaPagamento = True
    
    FillGridPg TBLVendas("C�DIGO")
    
    Exit Function
    
ErroVendas:
    TBLVendas.CancelUpdate
    GeraMensagemDeErro "CaixaF�cil - GravaPagamento - Or�amento: " & txtOr�amento, True
    GravaPagamento = False
    Exit Function
    
ErroFormaDePagamento:
    TBLFormaDePagamento.CancelUpdate
    GeraMensagemDeErro "CaixaF�cil - GravaPagamento - Or�amento: " & txtOr�amento, True
    GravaPagamento = False
End Function
Private Sub Localizar()
    If PosRecords Then
        GetRecords
        cmdGravar.Caption = "Fi&xar Venda"
        cmdFormaDePagamento.Enabled = True
        cmdCancelar.Enabled = True
    End If
End Sub
Private Function NotaFiscal() As Boolean
    On Error GoTo Erro
    
    Dim Cliente As String
    Dim Bookmark As Variant
        
    If MsgBox("Deseja emitir a nota fiscal agora?", vbInformation + vbYesNo, "Aviso") = vbNo Then
        NotaFiscal = True
    End If
    
    Cliente = SearchCliente(TBLVendas("C�DIGO DO CLIENTE"), byCodigo)
    
    If Cliente = "CONSUMIDOR" Then
        MsgBox "Cliente: " & Cliente & vbCr & "Deve-se identificar um cliente!", vbInformation, "Aviso"
    End If
    
    Set frmEncontrar.DBBancoDeDados = DBCadastro
    frmEncontrar.CampoPrincipal = Cliente
    frmEncontrar.NomeDaJanela = "Cliente"
    frmEncontrar.LabelDescription = "Nome/Raz�o Social"
    frmEncontrar.Mensagem = "Nenhum cliente foi selecionado!"
    frmEncontrar.BancoDeDados = "CADASTRO"
    frmEncontrar.Tabela = "CLIENTE"
    frmEncontrar.Indice = "2"
    frmEncontrar.CampoChave = "C�DIGO"
    frmEncontrar.CampoPreencheLista = "NOME - RAZ�O SOCIAL"
    frmEncontrar.Show vbModal
    Cliente = frmEncontrar.Chave
     
    Bookmark = TBLVendas.Bookmark
    TBLVendas.Edit
    TBLVendas("C�DIGO DO CLIENTE") = Cliente
    TBLVendas.Update
    TBLVendas.Bookmark = Bookmark
    
    NotaFiscal = False
    
    frmNotaFiscal.mOr�amento = TBLVendas("C�DIGO")
    frmNotaFiscal.Show vbModal
    mNF = frmNotaFiscal.mNF
    mTotalDeNotas = frmNotaFiscal.mTotalDeNotas
    
    Bookmark = TBLVendas.Bookmark
    TBLVendas.Edit
    TBLVendas("NOTA FISCAL") = mNF
    TBLVendas("TOTAL DE NOTAS FISCAIS") = mTotalDeNotas
    TBLVendas.Update
    TBLVendas.Bookmark = Bookmark
    
    NotaFiscal = True
    
    Exit Function
    
Erro:
    GeraMensagemDeErro "Nota Fiscal - Nota Fiscal"
    NotaFiscal = False
End Function
Private Function PosRecords() As Boolean
    TBLVendas.Seek "=", Val(txtOr�amento)
    If TBLVendas.NoMatch Then
        PosRecords = False
        MsgBox "N�o encontrei o or�amento " & txtOr�amento, vbInformation, "Aviso"
    Else
        If TBLVendas("TIPO") <> "O" Then
            MsgBox "Este or�amento n�o pode ser editado!", vbInformation, "Aviso"
            PosRecords = False
        Else
            PosRecords = True
        End If
    End If
End Function
Private Sub GetRecords()
    On Error Resume Next
    
    Dim pValor1 As Currency
    Dim pValor2 As Currency
    Dim pValor3 As Currency
    Dim pValor4 As Currency
    
    lPula = True
    txtOr�amento = TBLVendas("C�DIGO")
    mC�digoDoCliente = TBLVendas("C�DIGO DO CLIENTE")
    txtCliente = SearchCliente(TBLVendas("C�DIGO DO CLIENTE"), byCodigo)
    txtValor = FormatStringMask("@V ##.###.##0,00", StrVal(TBLVendas("VALOR TOTAL DA VENDA")))
    txtDesconto = FormatStringMask("@V ##.###.##0,00", StrVal(TBLVendas("DESCONTO TOTAL DA VENDA")))
    txtValorBonus = FormatStringMask("@V ##.###.##0,00", StrVal(TBLVendas("VALOR DO BONUS")))
    txtDataOr�amento = TBLVendas("DATA DO OR�AMENTO")
    CorrigeData DataMask, txtDataOr�amento, TBLVendas("DATA DO OR�AMENTO")
    
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
End Sub
Private Function SetRecords()
    On Error GoTo ErroVendas
    
    Dim Recno As Variant
    Dim Cont%
    Dim Msg$
    Dim Confirma��o As Integer, Msg1$, Msg2$
    Dim pAlterar As Boolean, pInserir As Boolean
    Dim C�digo$, Quantidade$, Pre�oUnit�rio$, Pre�oTotal$, Descri��o$, Tributa��o$, Total$
    Dim Status$, ValorTotal As Currency, DescontoTotal As Currency
    Dim AuxValor$, AuxTexto$
    Dim C�digoDoLote As String
    Dim D�gitoDoLote As String
    
    WS.BeginTrans 'Inicia uma Transa��o
        
    TBLVendas.Edit
    
    TBLVendas("TIPO") = "V"
    TBLVendas("DATA DA VENDA") = txtDataVenda
    TBLVendas("USERNAME - ALTERA") = gUsu�rio
    TBLVendas("DATA - ALTERA") = Date
    TBLVendas("HORA - ALTERA") = Time
    TBLVendas.Update
    
    On Error GoTo ErroEstoque
    TBLVendasItens.Seek "=", txtOr�amento
    Do While Not TBLVendasItens.EOF And TBLVendasItens("OR�AMENTO") = txtOr�amento
        If Not AtualizaProduto(TBLVendasItens("C�DIGO DO PRODUTO"), "-", TBLVendasItens("QUANTIDADE")) Then
            GoTo ErroVendas
        End If
        
        TBLVendasLote.Seek "=", txtOr�amento
        If Not TBLVendasLote.NoMatch Then
            Do While Not TBLVendasLote.EOF And TBLVendasLote("OR�AMENTO") = txtOr�amento
                C�digoDoLote = GetWordSeparatedBy(TBLVendasLote("C�DIGO DO LOTE"), 1, "-")
                D�gitoDoLote = GetWordSeparatedBy(TBLVendasLote("C�DIGO DO LOTE"), 1, "-")
                
                If Not AtualizaLote(TBLVendasLote("C�DIGO DO PRODUTO"), C�digoDoLote, D�gitoDoLote, TBLVendasLote("QUANTIDADE"), TBLVendasLote("M�LTIPLO")) Then
                    GoTo ErroVendas
                End If
                TBLVendasLote.MoveNext
                
                If TBLVendasLote.EOF Then
                    Exit Do
                End If
            Loop
        End If
                    
        TBLVendasItens.MoveNext
        If TBLVendasItens.EOF Then
            Exit Do
        End If
    Loop
    
    TBLFormaDePagamento.Index = "VENDAFORMADEPAGAMENTO1"
    TBLFormaDePagamento.Seek ">=", txtOr�amento
    If TBLFormaDePagamento.NoMatch Or TBLFormaDePagamento("OR�AMENTO") <> txtOr�amento Then
        MsgBox "Nenhuma Forma de Pagamento foi cadastrada para esta venda", vbCritical, "Aviso"
        GoTo Erro
    End If
    
    On Error GoTo ErroPlanoDePagamento
    TBLPlanoDePagamento.Seek "=", mTipoDePagamento
    If TBLPlanoDePagamento.NoMatch Then
        GoTo ErroPlanoDePagamento
    End If
    
    On Error GoTo ErroFormaDePagamento
    TBLFormaDePagamento.Edit
    TBLFormaDePagamento("BAIXADO") = TBLPlanoDePagamento("PRIMEIRO PAGAMENTO BAIXADO")
    TBLFormaDePagamento.Update
    
    TBLFormaDePagamento.Index = IndiceFormaDePagamentoAtivo
    
    On Error GoTo 0
    
'    Status = VerStatusECF
'
'    If Not AbrirCupomFiscal Then
'        GoTo ErroPDV
'    Else
'        If Mid(Status, 1, 2) = ".-" Then
'            AuxTexto = Mid(Status, 3, 4)
'            Status = Mid(Status, 7, Len(Status) - 7)
'            MsgBox Status, vbCritical, "Erro #" & AuxTexto
'            GoTo ErroPDV
'        End If
'    End If
    
    ValorTotal = 0
    
    TBLVendas.Seek "=", txtOr�amento
    TBLVendasItens.Seek "=", txtOr�amento
    
    DescontoTotal = TBLVendas("DESCONTO TOTAL DA VENDA") + TBLVendas("VALOR DO BONUS")
    
    Do While Not TBLVendasItens.EOF And TBLVendasItens("OR�AMENTO") = txtOr�amento
        C�digo = LeftBlankString(SearchAdvancedProduto(TBLVendasItens("C�DIGO DO PRODUTO"), vbC�digoDoFornecedor, vbIndice2), 13)
        Quantidade = LeftZeroString(Str(TBLVendasItens("QUANTIDADE")), 4) & "000"
        Pre�oUnit�rio = "0" & StrTran(FormatStringMask("@V 000000,00", StrVal(TBLVendasItens("VALOR UNIT�RIO"))), ",")
        Pre�oTotal = "0" & StrTran(FormatStringMask("@V 000000000,00", StrVal(TBLVendasItens("VALOR UNIT�RIO") * TBLVendasItens("QUANTIDADE"))), ",")
        Descri��o = RightBlankString(SearchAdvancedProduto(TBLVendasItens("C�DIGO DO PRODUTO"), vbDescri��o, vbIndice2), 24)
        Tributa��o = "I  "
        
'        RegistrarItemVendido C�digo, Quantidade, Pre�oUnit�rio, Pre�oTotal, Descri��o, Tributa��o
        
'        Status = VerStatusECF
'        If Mid(Status, 1, 2) = ".-" Then
'            AuxTexto = Mid(Status, 3, 4)
'            Status = Mid(Status, 7, Len(Status) - 7)
'            MsgBox Status, vbCritical, "Erro #" & AuxTexto
'            GoTo ErroPDV
'        End If
        
        ValorTotal = ValorTotal + TBLVendasItens("VALOR UNIT�RIO") * TBLVendasItens("QUANTIDADE")
        
        TBLVendasItens.MoveNext
        If TBLVendasItens.EOF Then
            Exit Do
        End If
    Loop
    
    If DescontoTotal > 0 Then
        AuxValor = StrTran(FormatStringMask("@V 0000000000,00", StrVal(DescontoTotal)), ",")
        AuxTexto = FormatStringMask("@V ##%", StrVal(DescontoTotal * 100 / ValorTotal))
        AuxTexto = RightBlankString(AuxTexto, 10)
        DescontoSobreCupomFiscal AuxTexto, AuxValor
    End If
    
'    frmTotal.ValorAPagar = ValorTotal - DescontoTotal
'    frmTotal.Show 1
'    Total = frmTotal.Total
'
'    Set frmTotal = Nothing
    
'    TotalizarCupomFiscal Total
'
'    Status = VerStatusECF
'
'    If Mid(Status, 1, 2) = ".-" Then
'        AuxTexto = Mid(Status, 3, 4)
'        Status = Mid(Status, 7, Len(Status) - 7)
'        MsgBox Status, vbCritical, "Erro #" & AuxTexto
'        GoTo ErroPDV
'    End If
'
'    FecharCupomFiscal
'    Status = VerStatusECF
'
'    If Mid(Status, 1, 2) = ".-" Then
'        AuxTexto = Mid(Status, 3, 4)
'        Status = Mid(Status, 7, Len(Status) - 7)
'        MsgBox Status, vbCritical, "Erro #" & AuxTexto
'        GoTo ErroPDV
'    End If
    
    WS.CommitTrans 'Grava as altera��es ou inclus�es se n�o houverem erros
    
    Log gUsu�rio, "Inclus�o - Caixa F�cil: " & txtOr�amento
        
    SetRecords = True
    
    Exit Function
    
ErroVendas:
    TBLVendas.CancelUpdate
    GeraMensagemDeErro "Sa�daVendas - SetRecords - ErroVendas - " & txtOr�amento, True
    TBLFormaDePagamento.Index = IndiceFormaDePagamentoAtivo
    SetRecords = False
    Exit Function
    
ErroEstoque:
    GeraMensagemDeErro "Sa�daVendas - SetRecords - ErroEstoque - " & txtOr�amento, True
    TBLFormaDePagamento.Index = IndiceFormaDePagamentoAtivo
    SetRecords = False
    Exit Function
    
ErroPlanoDePagamento:
    If Err <> 0 Then
        GeraMensagemDeErro "Sa�daVendas - SetRecords - ErroPlanoDePagamento - " & txtOr�amento, True
    Else
        MsgBox "N�o encontrei o Plano de Pagamento com c�digo: " & mTipoDePagamento
    End If
    SetRecords = False
    Exit Function
    
ErroFormaDePagamento:
    TBLFormaDePagamento.CancelUpdate
    GeraMensagemDeErro "Sa�daVendas - SetRecords - ErroFormaDePagamento - " & txtOr�amento, True
    TBLFormaDePagamento.Index = IndiceFormaDePagamentoAtivo
    SetRecords = False
    Exit Function
    
ErroPDV:
    WS.Rollback
    TBLFormaDePagamento.Index = IndiceFormaDePagamentoAtivo
    SetRecords = False
    Exit Function
    
Erro:
    TBLFormaDePagamento.Index = IndiceFormaDePagamentoAtivo
    SetRecords = False
End Function
Private Sub ZeraCampos()
    On Error Resume Next
    
    lPula = True
    txtOr�amento = Empty
    txtDataOr�amento = Empty
    txtDataVenda = Date
    CorrigeData DataMask, txtDataVenda, Date
    txtValor = FormatStringMask("@V ##.###.##0,00", "0,00")
    txtValorTotal = "R$" & String(6, " ") & FormatStringMask("@V ##.###.##0,00", "0,00")
    txtDesconto = FormatStringMask("@V ##.###.##0,00", "0,00")
    txtValorBonus = FormatStringMask("@V ##.###.##0,00", "0,00")
    txtPorcentagemBonus = FormatStringMask("@V ##0,00", "  0,00")
    txtCliente = Empty
    mC�digo = 0
    ReDim dbgrdItensArray(MAXCOLS - 1, 0)
    ReDim FormaDePagamentoArray(MAXCOLSPG - 1, 0)
    mTotalRows = 0
    mTotalPagamentos = 0
    mRecno = 0
    mRecnoPg = 0
    dbgrdItens.ReBind
    lPula = False
    mTotalPagamentos = Empty
    mValorAVista = Empty
    mValorAPrazo = Empty
    mTipoDePagamento = 0
End Sub
Private Sub cmdCancelar_Click()
    ZeraCampos
    DesativaCampos
    cmdGravar.Caption = "&Localizar"
    cmdGravar.Enabled = False
    cmdCancelar.Enabled = False
    Bot�oGravar False
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
    If cmdGravar.Caption = "&Localizar" Then
        Localizar
    Else
        If NotaFiscal Then
            Gravar
        End If
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
    If lFechar Then
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
    If Not VendasLoteAberto Then
        Unload Me
        Exit Sub
    End If
    If Not Par�metrosAberto Then
        Unload Me
        Exit Sub
    End If
    
    Bot�oGravar False
    Bot�oGravar True
    Navega��oInferior False
    Navega��oSuperior False
    Bot�oExcluir False
    Bot�oIncluir False
    Bot�oImprimir False
    
    BarraDeStatus StatusBarAviso
    dbgrdItens.Refresh

    If lAtualizar Then
        Bot�oAtualizar True
    Else
        Bot�oAtualizar False
    End If
End Sub
Private Sub Form_Deactivate()
    cmdGravar.Enabled = False
    Bot�oImprimir False
End Sub
Private Sub Form_Load()
    On Error GoTo Erro
    
    Dim Cont%
    
    ZeraCampos
    
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
    
    VendasLoteAberto = AbreTabela(Dicion�rio, "FINANCEIRO", "VENDA - LOTES", DBFinanceiro, TBLVendasLote, TBLTabela, dbOpenTable)
    
    If VendasLoteAberto Then
        IndiceVendasLotesAtivo = "VENDALOTES2"
        TBLVendasLote.Index = IndiceVendasLotesAtivo
    Else
        MsgBox "N�o consegui abrir a tabela 'Lote de Vendas' !", vbCritical, "Erro"
        GoTo Erro
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
    
    Navega��oInferior False
        
    StatusBarAviso = "Pronto"
    
    DesativaCampos
    
    Exit Sub
    
Erro:
    GeraMensagemDeErro "CaixaF�cil - Load"
    lFechar = True
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    mdiGeal.StatusBar.Panels("Posi��o").Visible = False
    ResizeStatusBar
End Sub
Private Sub Form_Unload(Cancel As Integer)
    If VendasAberto Then
        TBLVendas.Close
    End If
    If VendasItensAberto Then
        TBLVendasItens.Close
    End If
    If VendasLoteAberto Then
        TBLVendasLote.Close
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
Private Sub AcertaDesconto(ByVal oldvalue As Currency, ByVal Desconto As Currency, ByVal NewValor)
    Dim pValorDescontoTotal As Currency
    Dim pValorDesconto As Currency
    
    pValorDescontoTotal = StrVal(txtDesconto)
    pValorDesconto = oldvalue * Desconto
    pValorDesconto = StrVal(FormatStringMask("@V ##.###.##0,00", ValStr(pValorDesconto)))
    pValorDescontoTotal = pValorDescontoTotal - pValorDesconto
    pValorDesconto = NewValor * Desconto
    pValorDescontoTotal = pValorDescontoTotal + pValorDesconto
    
    lPula = True
    txtDesconto = FormatStringMask("@V ##.###.##0,00", ValStr(pValorDescontoTotal))
    lPula = False
End Sub
Private Sub txtOr�amento_Change()
    If txtOr�amento <> Empty Then
        cmdGravar.Enabled = True
    End If
End Sub
