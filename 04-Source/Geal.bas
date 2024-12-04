Attribute VB_Name = "GealFunc"
Option Explicit

Global PDV          As Object
Global Const lFalse As Byte = 1
Global Const lTrue  As Byte = 0

Global gEmpresa      As String
Global gCGC          As String

Global glPortaAberta As Boolean

Global Dicion�rio$
Global WS               As Workspace
Global DBDGeal          As Database
Global TBLArquivo       As Table, TBLTabela As Table, TBLCampo As Table, TBLIndice As Table
Global Dicion�rioAberto As Boolean

Global DBCadastro        As Database
Global DBDCadastroAberto As Boolean

Global DBUsu�rio        As Database
Global DBDUsu�rioAberto As Boolean

Global DBFinanceiro        As Database
Global DBDFinanceiroAberto As Boolean

Global DBUtilit�rio        As Database
Global DBDUtilit�rioAberto As Boolean

Global DBSistema        As Database
Global DBDSistemaAberto As Boolean

Global gUsu�rio$

Global gEstado As Byte

Global Aplica��oPath$

Global Const vbDescri��o          As Byte = 1
Global Const vbC�digo             As Byte = 2
Global Const vbValorUnit�rio      As Byte = 3
Global Const vbValValorUnit�rio   As Byte = 4
Global Const vbC�digoDoFornecedor As Byte = 5
Global Const vbLote               As Byte = 6
Global Const vbDescontoM�ximo     As Byte = 7
Global Const vbTributo            As Byte = 8

Global Const vbIndice1 As Byte = 1
Global Const vbIndice2 As Byte = 2
Global Const vbIndice3 As Byte = 3
Global Const vbIndice4 As Byte = 4

Global Const vbIncluir   As Byte = 1
Global Const vbAlterar   As Byte = 2
Global Const vbExcluir   As Byte = 3
Global Const vbConsultar As Byte = 4

Global Const DataMask       As String = "DD/MM/YYYY"
Global Const CheckDataMask  As String = "@D DD/MM/YYYY"
Global Const DataNula       As String = "  /  /    "

Global Const DataMaskMes      As String = "DD"
Global Const CheckDataMaskMes As String = "@D DD"
Global Const DataNulaMes      As String = "  "

Global Const byCodigo As Byte = 1
Global Const byCGCCPF As Byte = 2
Global Const byNome   As Byte = 3

'Constantes para o arquivo de registro do windows
Global Const APP_CATEGORY = "Microsoft Visual Basic AddIns"
Global gAPPNAME As String
Public Function AbreBaseDeDados(ByVal lIn�cio As Boolean, ByVal lExclusivo As Boolean) As Boolean
    Dim Cont As Long
    Dim NomeDaEmpresa As String
    Dim TBLPar�metros As Table
    Dim Par�metrosAberto As Boolean
    
    If lIn�cio Then
        'Path da aplica��o
        Aplica��oPath = App.Path
        
        Set WS = DBEngine.Workspaces(0)
    End If
    
    If lIn�cio Then
        'Abre o Dicion�rio de Dados
        
        NomeDaEmpresa = GetRegistryString("Geal", "Geral", "Empresa")
        frmSplash.lblLicenseTo = NomeDaEmpresa
        
        Dicion�rio = GetRegistryString("Geal", "Geral", "Dicion�rio")
          
        frmSplash.lblWarning = "Abrindo... dicion�rio de dados - " & Dicion�rio
        frmSplash.lblWarning.Refresh
        
        If Dicion�rio = Empty Then
            MsgBox "Erro na abertura do Dicion�rio de Dados!", vbExclamation, "Erro"
            AbreBaseDeDados = False
            Exit Function
        End If
        
        Dicion�rioAberto = AbreDicion�rio(WS, Dicion�rio, DBDGeal, TBLArquivo, TBLTabela, TBLCampo, TBLIndice)
        
        If Not Dicion�rioAberto Then
            MsgBox "Dicion�rio " + Dicion�rio + " n�o foi aberto !", vbExclamation, "Erro "
            AbreBaseDeDados = False
            Exit Function
        End If
    End If
    
    'Abre Cadastro
    If lIn�cio Then
        frmSplash.lblWarning = "Abrindo... base de dados - CADASTRO"
        frmSplash.lblWarning.Refresh
    End If
    
    DBDCadastroAberto = AbreArquivo(WS, Dicion�rio, "CADASTRO", DBCadastro, TBLArquivo, lExclusivo)
    
    If Not DBDCadastroAberto Then
        MsgBox "Erro na abertura do arquivo 'CADASTRO' ", vbExclamation, "Erro"
        AbreBaseDeDados = False
        Exit Function
    End If
    
    'Abre Usu�rio
    If lIn�cio Then
        frmSplash.lblWarning = "Abrindo... base de dados - USU�RIO"
        frmSplash.lblWarning.Refresh
    End If
    
    DBDUsu�rioAberto = AbreArquivo(WS, Dicion�rio, "USU�RIO", DBUsu�rio, TBLArquivo, lExclusivo)
    
    If Not DBDUsu�rioAberto Then
        MsgBox "Erro na abertura do arquivo 'USU�RIO' ", vbExclamation, "Erro"
        AbreBaseDeDados = False
        Exit Function
    End If
    
    'Abre Financeiro
    If lIn�cio Then
        frmSplash.lblWarning = "Abrindo... base de dados - FINANCEIRO"
        frmSplash.lblWarning.Refresh
    End If
    
    DBDFinanceiroAberto = AbreArquivo(WS, Dicion�rio, "FINANCEIRO", DBFinanceiro, TBLArquivo, lExclusivo)
    
    If Not DBDFinanceiroAberto Then
        MsgBox "Erro na abertura do arquivo 'FINANCEIRO' ", vbExclamation, "Erro"
        AbreBaseDeDados = False
        Exit Function
    End If
    
    'Abre Utilit�rio
    If lIn�cio Then
        frmSplash.lblWarning = "Abrindo... base de dados - UTILIT�RIO"
        frmSplash.lblWarning.Refresh
    End If
    
    DBDUtilit�rioAberto = AbreArquivo(WS, Dicion�rio, "UTILIT�RIO", DBUtilit�rio, TBLArquivo, lExclusivo)
    
    If Not DBDUtilit�rioAberto Then
        MsgBox "Erro na abertura do arquivo 'UTILIT�RIO' ", vbExclamation, "Erro"
        AbreBaseDeDados = False
        Exit Function
    End If
    
    'Abre Sistema
    If lIn�cio Then
        frmSplash.lblWarning = "Abrindo... base de dados - SISTEMA"
        frmSplash.lblWarning.Refresh
    End If
    
    DBDSistemaAberto = AbreArquivo(WS, Dicion�rio, "SISTEMA", DBSistema, TBLArquivo, lExclusivo)
    
    If Not DBDSistemaAberto Then
        MsgBox "Erro na abertura do arquivo 'SISTEMA' ", vbExclamation, "Erro"
        AbreBaseDeDados = False
        Exit Function
    End If
    
    'Pega o nome da Empresa e o CGC
    Par�metrosAberto = AbreTabela(Dicion�rio, "SISTEMA", "PAR�METROS", DBSistema, TBLPar�metros, TBLTabela, dbOpenTable)
    
    If Par�metrosAberto Then
    Else
        MsgBox "N�o consegui abrir a tabela 'Par�metros' !", vbCritical, "Erro"
        AbreBaseDeDados = False
        Exit Function
    End If
    
    gEmpresa = TBLPar�metros("EMPRESA")
    gCGC = TBLPar�metros("CGC")
    
    TBLPar�metros.Close
    
    AbreBaseDeDados = True
    
    If lIn�cio Then
        frmSplash.lblWarning = ""
        frmSplash.lblWarning.Refresh
    End If
End Function
Public Function AbrirCupomFiscal() As Boolean
    If PDV.AbrirCupomFiscal(Chr(27) & Chr(46) & "17}") Then
        AbrirCupomFiscal = True
    Else
        AbrirCupomFiscal = False
    End If
End Function
Public Function AbrirPorta(ByRef AbriPorta As Boolean)
    Dim Status As String
    
    If Not glPortaAberta Then
        If Not PDV.AbrirPorta(2, 5, 0, True) Then
            Status = VerStatusECF
            MsgBox Status, vbCritical, "Erro na abertura do ECF"
            AbrirPorta = False
            AbriPorta = False
            glPortaAberta = False
        Else
            AbrirPorta = True
            AbriPorta = True
            glPortaAberta = True
        End If
    Else
        AbrirPorta = True
        AbriPorta = False
    End If
End Function
Private Sub AcessoNegado(ByVal lAcesso As Boolean)
    lAcesso = Not lAcesso
    With mdiGeal
        .mnuArquivoAbrirAgencia.Enabled = lAcesso
        .mnuArquivoAbrirBanco.Enabled = lAcesso
        .mnuArquivoAbrirCliente.Enabled = lAcesso
        .mnuArquivoAbrirContaCorrente.Enabled = lAcesso
        .mnuArquivoAbrirFornecedor.Enabled = lAcesso
        .mnuArquivoAbrirFuncion�rio.Enabled = lAcesso
        .mnuArquivoAbrirProduto.Enabled = lAcesso
        .mnuArquivoAbrirDespesas.Enabled = lAcesso
        .mnuArquivoAbrir.Visible = lAcesso
        
        .mnuMovimentoEntradaCompra.Enabled = lAcesso
        .mnuMovimentoEntradaDevolu��oTroca.Enabled = lAcesso
        .mnuMovimentoEntrada.Visible = lAcesso
        
        .mnuMovimentoSa�daVenda.Enabled = lAcesso
        .mnuMovimentoSa�daDevolu��oTroca.Enabled = lAcesso
        .mnuMovimentoSa�da.Visible = lAcesso
        
        .mnuMovimentoMovimentoDi�rio.Enabled = lAcesso
        .mnuMovimentoContaCorrente.Enabled = lAcesso
        .mnuMovimentoCaixa.Enabled = lAcesso
        .mnuMovimentoCaixaF�cil.Enabled = lAcesso
        .mnuMovimentoDespesas.Enabled = lAcesso
        .mnuMovimento.Visible = lAcesso
        
        .mnuPar�metrosDepartamento.Enabled = lAcesso
        .mnuPar�metrosSe��o.Enabled = lAcesso
        .mnuPar�metrosDepartamentoSe��o.Enabled = lAcesso
        .mnuPar�metrosTipodeICM.Enabled = lAcesso
        .mnuPar�metrosTipodeEmbalagem.Enabled = lAcesso
        .mnuPar�metrosUnidades.Enabled = lAcesso
        .mnuPar�metrosLocalidadeDeEstoque.Enabled = lAcesso
        .mnuPar�metrosPlanoDePagamento.Enabled = lAcesso
        .mnuPar�metrosCaixa.Visible = lAcesso
        .mnuPar�metros.Visible = lAcesso
        
        .mnuPar�metrosUsu�rios.Visible = lAcesso
        .mnuPar�metrosGrupos.Visible = lAcesso
        .mnuPar�metrosSenhaDoSistema.Visible = lAcesso
        
        .mnuUtilit�riosConsultaSQL.Visible = lAcesso
        
        .mnuSep8.Visible = lAcesso
        .mnuSep9.Visible = lAcesso
        .mnuSep20.Visible = lAcesso
        .mnuSep11.Visible = lAcesso
        .mnuSep12.Visible = lAcesso
        .mnuSep16.Visible = lAcesso
        .mnuSep18.Visible = lAcesso
        .mnuPar�metroSep2.Visible = lAcesso
        .mnuPar�metroSep3.Visible = lAcesso
        .mnuPar�metroSep4.Visible = lAcesso
    End With
End Sub
Public Sub AllBot�es(ByVal Valor As Boolean)
    Bot�oAtualizar Valor
    Bot�oIncluir Valor
    Bot�oExcluir Valor
    Bot�oGravar Valor
    Bot�oImprimir Valor
    Navega��oInferior Valor
    Navega��oSuperior Valor
    BarraDeStatus "Pronto"
End Sub
Public Function Allow(ByVal Categoria As String, ByVal Direito As String, Optional ByVal Usu�rio As String) As Boolean
    On Error GoTo Erro
    
    Dim lRetorno As Boolean
    Dim TBLGrupos As Table
    Dim GruposAberto As Boolean
    Dim IndiceGruposAtivo$
    
    Dim TBLUsu�rioGrupo As Table
    Dim Usu�rioGrupoAberto As Boolean
    Dim IndiceUsu�rioGrupoAtivo$
    
    If IsMissing(Usu�rio) Or Usu�rio = Empty Then
        Usu�rio = gUsu�rio
    End If
    
    If Usu�rio = "ADMIN" Then
        Allow = True
    End If
    
    GruposAberto = AbreTabela(Dicion�rio, "USU�RIO", "GRUPO", DBUsu�rio, TBLGrupos, TBLTabela, dbOpenTable)
    
    If GruposAberto Then
        IndiceGruposAtivo = "GRUPO1"
        TBLGrupos.Index = IndiceGruposAtivo
    Else
        MsgBox "N�o consegui abrir a tabela 'GRUPO' !", vbCritical, "Erro"
        Exit Function
    End If
    
    Usu�rioGrupoAberto = AbreTabela(Dicion�rio, "USU�RIO", "USU�RIO - GRUPO", DBUsu�rio, TBLUsu�rioGrupo, TBLTabela, dbOpenTable)
    
    If Usu�rioGrupoAberto Then
        IndiceUsu�rioGrupoAtivo = "USU�RIOGRUPO1"
        TBLUsu�rioGrupo.Index = IndiceUsu�rioGrupoAtivo
    Else
        MsgBox "N�o consegui abrir a tabela 'GRUPO' !", vbCritical, "Erro"
        Exit Function
    End If
    
    TBLUsu�rioGrupo.Seek "=", Usu�rio
    
    If TBLUsu�rioGrupo.NoMatch Then
        Exit Function
    End If
    
    TBLUsu�rioGrupo.Seek "=", Usu�rio
    
    If TBLUsu�rioGrupo.NoMatch Then
        GoTo Fim
    End If
    
    lRetorno = False
    
    Do While Trim(TBLUsu�rioGrupo("USERNAME")) = Usu�rio
        TBLGrupos.Seek "=", TBLUsu�rioGrupo("C�DIGO DO GRUPO")
        If InStr(TBLGrupos(Categoria), Direito) Then
            lRetorno = True
        End If
        TBLUsu�rioGrupo.MoveNext
        If TBLUsu�rioGrupo.EOF Then
            Exit Do
        End If
    Loop
    
    Allow = lRetorno
    
Fim:
    If GruposAberto Then
        TBLGrupos.Close
    End If
    If Usu�rioGrupoAberto Then
        TBLUsu�rioGrupo.Close
    End If
    
    Exit Function
Erro:
    GeraMensagemDeErro "AllowInsert - Usu�rio:" & Usu�rio
    Allow = False
    If GruposAberto Then
        TBLGrupos.Close
    End If
    If Usu�rioGrupoAberto Then
        TBLUsu�rioGrupo.Close
    End If
End Function
Public Function AtualizaLote(ByVal C�digoDoProduto As Long, ByVal C�digoDoLote As String, ByVal D�gitoDoLote As String, ByVal Quantidade As Single, ByVal M�ltiplo As Single) As Boolean
    On Error GoTo Erro
    
    Dim TBLLote As Table
    Dim LoteAberto As Boolean
    
    'Abre tabela PRODUTO
    LoteAberto = AbreTabela(Dicion�rio, "CADASTRO", "LOTE DO PRODUTO", DBCadastro, TBLLote, TBLTabela, dbOpenTable)
    
    If LoteAberto Then
        TBLLote.Index = "LOTEDOPRODUTO1"
    Else
        MsgBox "N�o consegui abrir a tabela 'Lote do Produto' !", vbCritical, "Erro"
        Exit Function
    End If
    
    TBLLote.Seek "=", C�digoDoProduto, C�digoDoLote, D�gitoDoLote
    
    If TBLLote.NoMatch Then
        MsgBox "Produto: " & C�digoDoProduto & vbCr & "Lote: " & C�digoDoLote & "-" & D�gitoDoLote & vbCr & "N�o foi encontrado!", vbInformation, "Lote n�o existe!"
        AtualizaLote = False
    Else
        If TBLLote("QUANTIDADE") - (Quantidade * M�ltiplo) = 0 Then
            TBLLote.Delete
        Else
            TBLLote.Edit
            TBLLote("QUANTIDADE") = TBLLote("QUANTIDADE") - (Quantidade * M�ltiplo)
            TBLLote.Update
        End If
        AtualizaLote = True
    End If
    
    TBLLote.Close
    
    Exit Function
    
Erro:
    AtualizaLote = False
    TBLLote.CancelUpdate
    TBLLote.Close
End Function
Public Function AtualizaProduto(ByVal C�digo As Long, ByVal Opera��o, ByVal Valor As Long) As Boolean
    On Error GoTo Erro
    
    Dim TBLProduto As Table
    Dim ProdutoAberto As Boolean
    
    'Abre tabela PRODUTO
    ProdutoAberto = AbreTabela(Dicion�rio, "CADASTRO", "PRODUTO", DBCadastro, TBLProduto, TBLTabela, dbOpenTable)
    
    If ProdutoAberto Then
        TBLProduto.Index = "PRODUTO1"
    Else
        MsgBox "N�o consegui abrir a tabela 'Produto' !", vbCritical, "Erro"
        Exit Function
    End If
    
    TBLProduto.Seek "=", C�digo
    
    If TBLProduto.NoMatch Then
        AtualizaProduto = False
    End If
    
    TBLProduto.Edit
    If Opera��o = "+" Then
        TBLProduto("QUANTIDADE") = TBLProduto("QUANTIDADE") + Valor
    ElseIf Opera��o = "-" Then
        TBLProduto("QUANTIDADE") = TBLProduto("QUANTIDADE") - Valor
    End If
    
    TBLProduto.Update
    TBLProduto.Close
    
    AtualizaProduto = True
    
    Exit Function
Erro:
    AtualizaProduto = False
    TBLProduto.CancelUpdate
    TBLProduto.Close
End Function
Public Sub BarraDeStatus(Valor$)
    mdiGeal.StatusBar.Panels("Aviso").Text = Valor
End Sub
Public Sub Bot�oExcluir(ByVal Valor As Boolean)
    mdiGeal.mnuEditarExcluir.Enabled = Valor
    mdiGeal.Toolbar.Buttons("Excluir").Enabled = Valor
End Sub
Public Sub Bot�oGravar(ByVal Valor As Boolean)
    mdiGeal.mnuArquivoSalvar.Enabled = Valor
    mdiGeal.Toolbar.Buttons("Gravar").Enabled = Valor
End Sub
Public Sub Bot�oImprimir(ByVal Valor As Boolean)
    mdiGeal.mnuArquivoImprimir.Enabled = Valor
    mdiGeal.Toolbar.Buttons("Imprimir").Enabled = Valor
End Sub
Public Sub Bot�oAtualizar(ByVal Valor As Boolean)
    mdiGeal.mnuEditarAtualizar.Enabled = Valor
End Sub
Public Sub Bot�oIncluir(ByVal Valor As Boolean)
    mdiGeal.mnuEditarIncluir.Enabled = Valor
    mdiGeal.Toolbar.Buttons("Incluir").Enabled = Valor
End Sub
Public Function BuscaFuncion�rio(ByVal C�digo&) As String
    Dim Funcion�rioAberto As Boolean, TBLFuncion�rio As Table, IndiceFuncion�rioAtivo$
    
    Funcion�rioAberto = AbreTabela(Dicion�rio, "USU�RIO", "FUNCION�RIO", DBUsu�rio, TBLFuncion�rio, TBLTabela, dbOpenTable)
    
    If Funcion�rioAberto Then
        IndiceFuncion�rioAtivo = "FUNCION�RIO1"
        TBLFuncion�rio.Index = IndiceFuncion�rioAtivo
    Else
        MsgBox "N�o consegui abrir a tabela 'Funcion�rio' !", vbCritical, "Erro"
        BuscaFuncion�rio = ""
        Exit Function
    End If
    
    TBLFuncion�rio.Seek "=", C�digo
    
    If TBLFuncion�rio.NoMatch Then
        BuscaFuncion�rio = ""
    Else
        BuscaFuncion�rio = TBLFuncion�rio("NOME")
    End If
End Function
Public Function CancelarCupom() As Boolean
    If PDV.CancelarCupom(Chr(27) & Chr(46) & "05}") Then
        CancelarCupom = True
    Else
        CancelarCupom = False
    End If
End Function
Public Sub ChamaConfigura��es(ByVal Usu�rio$)
    On Error GoTo Erro
    
    Dim TBLGrupos As Table
    Dim GruposAberto As Boolean
    Dim IndiceGruposAtivo$
    
    Dim TBLUsu�rioGrupo As Table
    Dim Usu�rioGrupoAberto As Boolean
    Dim IndiceUsu�rioGrupoAtivo$
    
    Vis�oTotal
    AcessoNegado False
    
    If Usu�rio = "ADMIN" Then
        Exit Sub
    End If
    
    AcessoNegado True
    
    GruposAberto = AbreTabela(Dicion�rio, "USU�RIO", "GRUPO", DBUsu�rio, TBLGrupos, TBLTabela, dbOpenTable)
    
    If GruposAberto Then
        IndiceGruposAtivo = "GRUPO1"
        TBLGrupos.Index = IndiceGruposAtivo
    Else
        MsgBox "N�o consegui abrir a tabela 'GRUPO' !", vbCritical, "Erro"
        Exit Sub
    End If
    
    Usu�rioGrupoAberto = AbreTabela(Dicion�rio, "USU�RIO", "USU�RIO - GRUPO", DBUsu�rio, TBLUsu�rioGrupo, TBLTabela, dbOpenTable)
    
    If Usu�rioGrupoAberto Then
        IndiceUsu�rioGrupoAtivo = "USU�RIOGRUPO1"
        TBLUsu�rioGrupo.Index = IndiceUsu�rioGrupoAtivo
    Else
        MsgBox "N�o consegui abrir a tabela 'GRUPO' !", vbCritical, "Erro"
        Exit Sub
    End If
    
    TBLUsu�rioGrupo.Seek "=", Usu�rio
    
    If TBLUsu�rioGrupo.NoMatch Then
        Exit Sub
    End If
    
    TBLUsu�rioGrupo.Seek "=", Trim(Usu�rio)
    
    If TBLUsu�rioGrupo.NoMatch Then
        GoTo Fim
    End If
    
    Do While Trim(TBLUsu�rioGrupo("USERNAME")) = Trim(Usu�rio)
        TBLGrupos.Seek "=", TBLUsu�rioGrupo("C�DIGO DO GRUPO")
        
        'In�cio Arquivo
        'Ag�ncia
        If TBLGrupos("AG�NCIA") <> Empty Then
            mdiGeal.mnuArquivoAbrir.Visible = True
            mdiGeal.mnuArquivoAbrirAgencia.Visible = True
            mdiGeal.mnuArquivoAbrirAgencia.Enabled = True
            VisibleArquivo
        End If
        'Banco
        If TBLGrupos("BANCO") <> Empty Then
            mdiGeal.mnuArquivoAbrir.Visible = True
            mdiGeal.mnuArquivoAbrirBanco.Visible = True
            mdiGeal.mnuArquivoAbrirBanco.Enabled = True
            VisibleArquivo
        End If
        'Conta Corrente
        If TBLGrupos("CONTA CORRENTE") <> Empty Then
            mdiGeal.mnuArquivoAbrir.Visible = True
            mdiGeal.mnuArquivoAbrirContaCorrente.Visible = True
            mdiGeal.mnuArquivoAbrirContaCorrente.Enabled = True
            VisibleArquivo
        End If
        'Conta Corrente
        If TBLGrupos("CLIENTE") <> Empty Then
            mdiGeal.mnuArquivoAbrir.Visible = True
            mdiGeal.mnuArquivoAbrirCliente.Visible = True
            mdiGeal.mnuArquivoAbrirCliente.Enabled = True
            VisibleArquivo
        End If
        'Fornecedor
        If TBLGrupos("FORNECEDOR") <> Empty Then
            mdiGeal.mnuArquivoAbrir.Visible = True
            mdiGeal.mnuArquivoAbrirFornecedor.Visible = True
            mdiGeal.mnuArquivoAbrirFornecedor.Enabled = True
            VisibleArquivo
        End If
        'Funcion�rio
        If TBLGrupos("FUNCION�RIO") <> Empty Then
            mdiGeal.mnuArquivoAbrir.Visible = True
            mdiGeal.mnuArquivoAbrirFuncion�rio.Visible = True
            mdiGeal.mnuArquivoAbrirFuncion�rio.Enabled = True
            VisibleArquivo
        End If
        'Produto
        If TBLGrupos("PRODUTO") <> Empty Then
            mdiGeal.mnuArquivoAbrir.Visible = True
            mdiGeal.mnuArquivoAbrirProduto.Visible = True
            mdiGeal.mnuArquivoAbrirProduto.Enabled = True
            VisibleArquivo
        End If
        'Despesas
        If TBLGrupos("DESPESAS") <> Empty Then
            mdiGeal.mnuArquivoAbrir.Visible = True
            mdiGeal.mnuArquivoAbrirDespesas.Visible = True
            mdiGeal.mnuArquivoAbrirDespesas.Enabled = True
            VisibleArquivo
        End If
        'Fim Arquivo
        
        'In�cio Movimento
        'Compra
        If TBLGrupos("COMPRA") <> Empty Then
            mdiGeal.mnuMovimento.Visible = True
            mdiGeal.mnuMovimentoEntrada.Visible = True
            mdiGeal.mnuMovimentoEntradaCompra.Visible = True
            mdiGeal.mnuMovimentoEntradaCompra.Enabled = True
            VisibleMovimento
        End If
        'Devolu��o/Troca (Compra)
        If TBLGrupos("DEVOLU��O/TROCA (COMPRA)") <> Empty Then
            mdiGeal.mnuMovimento.Visible = True
            mdiGeal.mnuMovimentoEntrada.Visible = True
            mdiGeal.mnuMovimentoEntradaDevolu��oTroca.Visible = True
            mdiGeal.mnuMovimentoEntradaDevolu��oTroca.Enabled = True
            VisibleMovimento
        End If
        'Venda
        If TBLGrupos("VENDA") <> Empty Then
            mdiGeal.mnuMovimento.Visible = True
            mdiGeal.mnuMovimentoSa�da.Visible = True
            mdiGeal.mnuMovimentoSa�daVenda.Visible = True
            mdiGeal.mnuMovimentoSa�daVenda.Enabled = True
            VisibleMovimento
        End If
        'Devolu��o/Troca (Venda)
        If TBLGrupos("DEVOLU��O/TROCA (VENDA)") <> Empty Then
            mdiGeal.mnuMovimento.Visible = True
            mdiGeal.mnuMovimentoSa�da.Visible = True
            mdiGeal.mnuMovimentoSa�daDevolu��oTroca.Visible = True
            mdiGeal.mnuMovimentoSa�daDevolu��oTroca.Enabled = True
            VisibleMovimento
        End If
        'Movimento Di�rio
        If TBLGrupos("MOVIMENTO DI�RIO") <> Empty Then
            mdiGeal.mnuMovimento.Visible = True
            mdiGeal.mnuMovimentoMovimentoDi�rio.Visible = True
            mdiGeal.mnuMovimentoMovimentoDi�rio.Enabled = True
            mdiGeal.mnuSep8.Visible = True
            VisibleMovimento
        End If
        'Conta Corrente (Movimento)
        If TBLGrupos("CONTA CORRENTE (MOVIMENTO)") <> Empty Then
            mdiGeal.mnuMovimento.Visible = True
            mdiGeal.mnuMovimentoContaCorrente.Visible = True
            mdiGeal.mnuMovimentoContaCorrente.Enabled = True
            mdiGeal.mnuSep8.Visible = True
        End If
        'Caixa
        If TBLGrupos("CAIXA") <> Empty Then
            mdiGeal.mnuMovimento.Visible = True
            mdiGeal.mnuMovimentoCaixa.Visible = True
            mdiGeal.mnuMovimentoCaixa.Enabled = True
            mdiGeal.mnuSep9.Visible = True
            VisibleMovimento
        End If
        'Caixa F�cil
        If TBLGrupos("CAIXA F�CIL") <> Empty Then
            mdiGeal.mnuMovimento.Visible = True
            mdiGeal.mnuMovimentoCaixaF�cil.Visible = True
            mdiGeal.mnuMovimentoCaixaF�cil.Enabled = True
            mdiGeal.mnuSep9.Visible = True
            VisibleMovimento
        End If
        'Despesas
        If TBLGrupos("DESPESAS") <> Empty Then
            mdiGeal.mnuMovimento.Visible = True
            mdiGeal.mnuMovimentoDespesas.Visible = True
            mdiGeal.mnuMovimentoDespesas.Enabled = True
            mdiGeal.mnuSep20.Visible = True
            VisibleMovimento
        End If
        'Fim Movimento
        
        'In�cio Par�metros
        'Departamento
        If TBLGrupos("DEPARTAMENTO") <> Empty Then
            mdiGeal.mnuPar�metros.Visible = True
            mdiGeal.mnuPar�metrosDepartamento.Visible = True
            mdiGeal.mnuPar�metrosDepartamento.Enabled = True
            VisiblePar�metros
        End If
        'Se��o
        If TBLGrupos("SE��O") <> Empty Then
            mdiGeal.mnuPar�metros.Visible = True
            mdiGeal.mnuPar�metrosSe��o.Visible = True
            mdiGeal.mnuPar�metrosSe��o.Enabled = True
            VisiblePar�metros
        End If
        'Departamento - Se��o
        If TBLGrupos("DEPARTAMENTO - SE��O") <> Empty Then
            mdiGeal.mnuPar�metros.Visible = True
            mdiGeal.mnuPar�metrosDepartamentoSe��o.Visible = True
            mdiGeal.mnuPar�metrosDepartamentoSe��o.Enabled = True
            VisiblePar�metros
        End If
        'Tipo de ICM
        If TBLGrupos("TIPO DE ICM") <> Empty Then
            mdiGeal.mnuPar�metros.Visible = True
            mdiGeal.mnuPar�metrosTipodeICM.Visible = True
            mdiGeal.mnuPar�metrosTipodeICM.Enabled = True
            mdiGeal.mnuSep11.Visible = True
            VisiblePar�metros
        End If
        'Tipo de Embalagem
        If TBLGrupos("TIPO DE EMBALAGEM") <> Empty Then
            mdiGeal.mnuPar�metros.Visible = True
            mdiGeal.mnuPar�metrosTipodeEmbalagem.Visible = True
            mdiGeal.mnuPar�metrosTipodeEmbalagem.Enabled = True
            mdiGeal.mnuSep11.Visible = True
            VisiblePar�metros
        End If
        'Unidades
        If TBLGrupos("UNIDADES") <> Empty Then
            mdiGeal.mnuPar�metros.Visible = True
            mdiGeal.mnuPar�metrosUnidades.Visible = True
            mdiGeal.mnuPar�metrosUnidades.Enabled = True
            mdiGeal.mnuSep15.Visible = True
            VisiblePar�metros
        End If
        'Localidade de Estoque
        If TBLGrupos("LOCALIDADE DE ESTOQUE") <> Empty Then
            mdiGeal.mnuPar�metros.Visible = True
            mdiGeal.mnuPar�metrosLocalidadeDeEstoque.Visible = True
            mdiGeal.mnuPar�metrosLocalidadeDeEstoque.Enabled = True
            mdiGeal.mnuSep12.Visible = True
            VisiblePar�metros
        End If
        'Plano de Pagamento
        If TBLGrupos("PLANO DE PAGAMENTO") <> Empty Then
            mdiGeal.mnuPar�metros.Visible = True
            mdiGeal.mnuPar�metrosPlanoDePagamento.Visible = True
            mdiGeal.mnuPar�metrosPlanoDePagamento.Enabled = True
            mdiGeal.mnuSep16.Visible = True
            VisiblePar�metros
        End If
        'Fim Par�metros
        
        TBLUsu�rioGrupo.MoveNext
        If TBLUsu�rioGrupo.EOF Then
            Exit Do
        End If
    Loop
    
Fim:
    If GruposAberto Then
        TBLGrupos.Close
    End If
    If Usu�rioGrupoAberto Then
        TBLUsu�rioGrupo.Close
    End If
    
    Exit Sub
Erro:
    GeraMensagemDeErro "ChamaConfigura��es - Usu�rio:" & Usu�rio
    If GruposAberto Then
        TBLGrupos.Close
    End If
    If Usu�rioGrupoAberto Then
        TBLUsu�rioGrupo.Close
    End If
End Sub
Public Sub DeleteRegistryString(ByVal vsSection$, Optional ByVal vsSubSection, Optional ByVal vsKey)
    On Error Resume Next
    If Not IsMissing(vsSubSection) <> Empty Then
        vsSection = AddPath(vsSection, vsSubSection)
    End If
    DeleteSetting APP_CATEGORY & "\" & gAPPNAME, vsSection, vsKey
End Sub
Public Function DescontoSobreCupomFiscal(ByVal Texto$, ByVal Desconto As String) As Boolean
    If PDV.DescontoSobreCupomFiscal(Chr(27) & Chr(46) & "03" & Texto & Desconto & "N}") Then
        DescontoSobreCupomFiscal = True
    Else
        DescontoSobreCupomFiscal = False
    End If
End Function
Public Function FecharCupomFiscal() As Boolean
    If PDV.FecharCupomFiscal(Chr(27) & Chr(46) & "12}") Then
        FecharCupomFiscal = True
    Else
        FecharCupomFiscal = False
    End If
End Function
Public Sub FechaBaseDeDados()
    'Fecha todo os banco de dados abertos
    If DBDCadastroAberto Then
        DBCadastro.Close
    End If
    If DBDUsu�rioAberto Then
        DBUsu�rio.Close
    End If
    If DBDFinanceiroAberto Then
        DBFinanceiro.Close
    End If
    If DBDUtilit�rioAberto Then
        DBUtilit�rio.Close
    End If
    If DBDSistemaAberto Then
        DBSistema.Close
    End If
End Sub
Public Function FecharPorta() As Boolean
    PDV.FecharPorta
    glPortaAberta = False
End Function
Public Sub GeraMensagemDeErro(ByVal Opera��o, Optional ByVal Rollback As Boolean = False)
    Dim ErroAberto As Boolean, TBLErro As Table, IndiceErroAtivo$
    Dim ErroNumero&, ErroDescri��o$, Data, Hora
    
    Data = Date
    Hora = Time
    
    ErroNumero = Err.Number
    ErroDescri��o = Err.Description
    
    MensagemDeErro
    
    If Rollback Then
        WS.Rollback
    End If
    
    ErroAberto = AbreTabela(Dicion�rio, "SISTEMA", "ERRO", DBSistema, TBLErro, TBLTabela, dbOpenTable, True)
    
    If ErroAberto Then
        IndiceErroAtivo = "ERRO1"
        TBLErro.Index = IndiceErroAtivo
    Else
        MsgBox "N�o consegui abrir a tabela 'Erro' !", vbCritical, "Erro"
    End If
    
    If ErroAberto Then
        On Error Resume Next
        TBLErro.AddNew
        TBLErro("C�DIGO DO ERRO") = ErroNumero
        TBLErro("DESCRI��O") = ErroDescri��o
        TBLErro("USERNAME") = gUsu�rio
        TBLErro("OPERA��O") = Opera��o
        TBLErro("DATA") = Data
        TBLErro("HORA") = Hora
        TBLErro.Update
        TBLErro.Close
    End If
End Sub
Public Function GetRegistryString(ByVal vsSection$, ByVal vsSubSection$, ByVal vsItem$, Optional ByVal vsDefault) As String
    vsSection = AddPath(gAPPNAME, vsSection, vsSubSection)
    GetRegistryString = GetSetting(APP_CATEGORY, vsSection, vsItem, vsDefault)
End Function
Public Function IsCorrectFornecedor(txtObj As Object) As Boolean
    Dim FornecedorAberto As Boolean, TBLFornecedor As Table, IndiceFornecedorAtivo$
    
    FornecedorAberto = AbreTabela(Dicion�rio, "CADASTRO", "FORNECEDOR", DBCadastro, TBLFornecedor, TBLTabela, dbOpenTable)
    
    If FornecedorAberto Then
        IndiceFornecedorAtivo = "FORNECEDOR1"
        TBLFornecedor.Index = IndiceFornecedorAtivo
    Else
        MsgBox "N�o consegui abrir a tabela 'Fornecedor' !", vbCritical, "Erro"
        IsCorrectFornecedor = False
        Exit Function
    End If
    
    TBLFornecedor.Seek "=", txtObj
    
    If TBLFornecedor.NoMatch Then
        IsCorrectFornecedor = False
    Else
        IsCorrectFornecedor = True
    End If
End Function
Public Function IsCorrectFuncion�rio(txtObj As Object) As Boolean
    Dim Funcion�rioAberto As Boolean, TBLFuncion�rio As Table, IndiceFuncion�rioAtivo$
    
    Funcion�rioAberto = AbreTabela(Dicion�rio, "USU�RIO", "FUNCION�RIO", DBUsu�rio, TBLFuncion�rio, TBLTabela, dbOpenTable)
    
    If Funcion�rioAberto Then
        IndiceFuncion�rioAtivo = "FUNCION�RIO1"
        TBLFuncion�rio.Index = IndiceFuncion�rioAtivo
    Else
        MsgBox "N�o consegui abrir a tabela 'Funcion�rio' !", vbCritical, "Erro"
        IsCorrectFuncion�rio = False
        Exit Function
    End If
    
    If txtObj <> Empty Then
        TBLFuncion�rio.Seek "=", txtObj
        
        If TBLFuncion�rio.NoMatch Then
            IsCorrectFuncion�rio = False
        Else
            IsCorrectFuncion�rio = True
        End If
    Else
        IsCorrectFuncion�rio = False
    End If
End Function
Public Function ImpressaoDeCheque(ByVal Banco As String, ByVal Valor As String, ByVal Data As String, ByVal DadosAdicionais As String) As Boolean
    If PDV.ImpressaoDeCheque(Chr(27) & Chr(46) & "24" & Banco & Valor & "N" & DadosAdicionais & "4" & Data & "}") Then
        ImpressaoDeCheque = True
    Else
        ImpressaoDeCheque = False
    End If
End Function
Public Function LeituraX(ByVal Relat�rioGerencial As String) As Boolean
    If PDV.LeituraX(Chr(27) & Chr(46) & "13" & Relat�rioGerencial & "}") Then
        LeituraX = True
    Else
        LeituraX = False
    End If
End Function
Public Sub Log(ByVal Usu�rio As String, ByVal Opera��o As String)
    Dim LogAberto As Boolean, TBLLog As Table, IndiceLogAtivo$
    Dim Data, Hora
    
    LogAberto = AbreTabela(Dicion�rio, "SISTEMA", "LOG", DBSistema, TBLLog, TBLTabela, dbOpenTable, True)
    
    If Not LogAberto Then
        MsgBox "N�o consegui abrir a tabela 'Log' !", vbCritical, "Log"
        Exit Sub
    End If
    
    TBLLog.AddNew
    TBLLog("USERNAME") = Usu�rio
    TBLLog("OPERA��O") = Opera��o
    TBLLog("DATA") = Date
    TBLLog("HORA") = Time
    TBLLog.Update
End Sub
Public Sub Main()
    Dim Objeto As String
    
    frmSplash.Show
    
    Objeto = GetRegistryString("Geal", "Geral", "DllPDV")
    
    On Error Resume Next
    Set PDV = CreateObject(Objeto)
    If Err.Number <> 0 Then
        MensagemDeErro vbCr & "Main - Cria��o do Objeto " & Objeto & vbCr & "N�o ser� poss�vel utilizar o PDV"
    End If
    On Error GoTo 0
    
    If AbreBaseDeDados(True, False) Then
        If ValidaUsu�rio(frmSplash.Top, frmSplash.Height) Then
            SetRegistryString "Geal", "Geral", "Usu�rio", gUsu�rio
            ChamaConfigura��es gUsu�rio
            mdiGeal.Show
            Unload frmSplash
            Unload frmValidaUsu�rio
            Set frmSplash = Nothing
            Set frmValidaUsu�rio = Nothing
        Else
            Unload frmSplash
            Set frmSplash = Nothing
        End If
    Else
        Unload frmSplash
        Set frmSplash = Nothing
    End If
End Sub
Public Sub Navega��oInferior(ByVal Valor As Boolean)
    mdiGeal.mnuNavega��oPrimeiroRegistro.Enabled = Valor
    mdiGeal.mnuNavega��oRegistroAnterior.Enabled = Valor
    mdiGeal.Toolbar.Buttons("MoveFirst").Enabled = Valor
    mdiGeal.Toolbar.Buttons("MovePrevious").Enabled = Valor
End Sub
Public Sub Navega��oSuperior(ByVal Valor As Boolean)
    mdiGeal.mnuNavega��o�ltimoRegistro.Enabled = Valor
    mdiGeal.mnuNavega��oPr�ximoRegistro.Enabled = Valor
    mdiGeal.Toolbar.Buttons("MoveNext").Enabled = Valor
    mdiGeal.Toolbar.Buttons("MoveLast").Enabled = Valor
End Sub
Public Function Redu��oZ(ByVal Relat�rioGerencial As String, Optional ByVal Data As String) As Boolean
    If PDV.Redu��oZ(Chr(27) & Chr(46) & "14" & Relat�rioGerencial & "}") Then
        Redu��oZ = True
    Else
        Redu��oZ = False
    End If
End Function
Public Function RegistrarItemVendido(ByVal C�digo$, ByVal Quantidade$, ByVal Pre�oUnit�rio$, ByVal Pre�oTotal$, ByVal Descri��o$, ByVal Tributa��o$) As Boolean
    If PDV.RegistrarItemVendido(Chr(27) & Chr(46) & "01" & C�digo & Quantidade & Pre�oUnit�rio & Pre�oTotal & Descri��o & Tributa��o & "}") Then
        RegistrarItemVendido = True
    Else
        RegistrarItemVendido = False
    End If
End Function
Public Sub ResizeStatusBar()
    Dim Tamanho, Cont, Posi��o
    
    Tamanho = 0
    For Cont = 2 To mdiGeal.StatusBar.Panels.Count
        If mdiGeal.StatusBar.Panels(Cont).Visible = True Then
            Tamanho = Tamanho + mdiGeal.StatusBar.Panels(Cont).Width
        End If
    Next
    Posi��o = mdiGeal.Width - 500 - Tamanho
    If Posi��o >= 0 Then
        mdiGeal.StatusBar.Panels(1).Width = Posi��o
    End If
End Sub
Public Function SearchAdvancedProduto(ByVal C�digo As String, ByVal Tipo As Integer, Optional ByVal Indice As Integer) As Variant
    Dim TBLProduto As Table
    Dim ProdutoAberto As Boolean
    Dim IndiceProdutoAtivo$
    
    Dim TBLC�digoProduto As Table
    Dim C�digoProdutoAberto As Boolean
    Dim IndiceC�digoProdutoAtivo$
    
    Dim TBLPre�oProduto As Table
    Dim Pre�oProdutoAberto As Boolean
    Dim IndicePre�oProdutoAtivo$
    
    Dim TBLTipoDeICM As Table
    Dim TipoDeICMAberto As Boolean
    Dim IndiceTipoDeICMAtivo$
    
    If IsMissing(Indice) Or Indice = 0 Then
        Indice = 3 'C�digo do Fornecedor como padr�o a Tabela C�digo do Produto
    End If
    
    'Abre tabela PRODUTO
    ProdutoAberto = AbreTabela(Dicion�rio, "CADASTRO", "PRODUTO", DBCadastro, TBLProduto, TBLTabela, dbOpenTable)
    
    If ProdutoAberto Then
        IndiceProdutoAtivo = "PRODUTO1"
        TBLProduto.Index = IndiceProdutoAtivo
    Else
        MsgBox "N�o consegui abrir a tabela 'Produto' !", vbCritical, "Erro"
        Exit Function
    End If
    
    'Abre tabela C�DIGO DO PRODUTO
    C�digoProdutoAberto = AbreTabela(Dicion�rio, "CADASTRO", "C�DIGO DO PRODUTO", DBCadastro, TBLC�digoProduto, TBLTabela, dbOpenTable)
    'Se Indice 1 - +C�DIGO DO PRODUTO;+FORNECEDOR;+C�DIGO DO FORNECEDOR
    'Se Indice 2 - +C�DIGO DO PRODUTO
    'Se Indice 3 - +C�DIGO DO FORNECEDOR
    'Se Indice 4 - +C�DIGO DO PRODUTO;+FORNECEDOR
    'Se Indice 5 - +FORNECEDOR;+C�DIGO DO FORNECEDOR
    If C�digoProdutoAberto Then
        IndiceC�digoProdutoAtivo = "C�DIGODOPRODUTO" & Indice
        TBLC�digoProduto.Index = IndiceC�digoProdutoAtivo
    Else
        MsgBox "N�o consegui abrir a tabela 'C�digo do Produto' !", vbCritical, "Erro"
        Exit Function
    End If

    'Abre tabela PRE�O DO PRODUTO
    Pre�oProdutoAberto = AbreTabela(Dicion�rio, "CADASTRO", "PRE�O DO PRODUTO", DBCadastro, TBLPre�oProduto, TBLTabela, dbOpenTable)
    
    If Pre�oProdutoAberto Then
        IndicePre�oProdutoAtivo = "PRE�ODOPRODUTO1"
        TBLPre�oProduto.Index = IndicePre�oProdutoAtivo
    Else
        MsgBox "N�o consegui abrir a tabela 'Pre�oProduto' !", vbCritical, "Erro"
        Exit Function
    End If
     
    TipoDeICMAberto = AbreTabela(Dicion�rio, "CADASTRO", "TIPO DE ICM", DBCadastro, TBLTipoDeICM, TBLTabela, dbOpenTable)
    
    If TipoDeICMAberto Then
        IndiceTipoDeICMAtivo = "TIPODEICM1"
        TBLTipoDeICM.Index = IndiceTipoDeICMAtivo
    Else
        MsgBox "N�o consegui abrir a tabela 'Tipo De Embalagem' !", vbCritical, "Erro"
        Exit Function
    End If
    
    'Tipo do Retorno
    If Tipo = vbDescri��o Then
        TBLProduto.Seek "=", C�digo
        If TBLProduto.NoMatch Then
            SearchAdvancedProduto = ""
        Else
            SearchAdvancedProduto = TBLProduto("DESCRI��O")
        End If
    ElseIf Tipo = vbC�digo Then
        TBLC�digoProduto.Seek "=", C�digo
        If TBLC�digoProduto.NoMatch Then
            SearchAdvancedProduto = ""
        Else
            SearchAdvancedProduto = TBLC�digoProduto("C�DIGO DO PRODUTO")
        End If
    ElseIf Tipo = vbC�digoDoFornecedor Then
        TBLC�digoProduto.Seek "=", C�digo
        If TBLC�digoProduto.NoMatch Then
            SearchAdvancedProduto = ""
        Else
            SearchAdvancedProduto = TBLC�digoProduto("C�DIGO DO FORNECEDOR")
        End If
    ElseIf Tipo = vbValorUnit�rio Then
        TBLC�digoProduto.Seek "=", C�digo
        If TBLC�digoProduto.NoMatch Then
            SearchAdvancedProduto = Empty
            Exit Function
        End If
        TBLPre�oProduto.Seek "=", C�digo, TBLC�digoProduto("C�DIGO DO PRODUTO")
        If TBLPre�oProduto.NoMatch Then
            TBLPre�oProduto.Index = "PRE�ODOPRODUTO2"
            TBLPre�oProduto.Seek "=", C�digo
            If TBLPre�oProduto.NoMatch Then
                SearchAdvancedProduto = "0,00"
            Else
                SearchAdvancedProduto = FormatStringMask("@V ##.###.##0,00", ValStr(TBLPre�oProduto("PRE�O DE VENDA")))
            End If
        Else
            SearchAdvancedProduto = FormatStringMask("@V ##.###.##0,00", ValStr(TBLPre�oProduto("PRE�O DE VENDA")))
        End If
    ElseIf Tipo = vbValValorUnit�rio Then
        TBLC�digoProduto.Seek "=", C�digo
        If TBLC�digoProduto.NoMatch Then
            SearchAdvancedProduto = Empty
            Exit Function
        End If
        TBLPre�oProduto.Seek "=", C�digo, TBLC�digoProduto("C�DIGO DO PRODUTO")
        If TBLPre�oProduto.NoMatch Then
            TBLPre�oProduto.Index = "PRE�ODOPRODUTO2"
            TBLPre�oProduto.Seek "=", TBLC�digoProduto("C�DIGO DO PRODUTO")
            If TBLPre�oProduto.NoMatch Then
                SearchAdvancedProduto = 0
            Else
                SearchAdvancedProduto = TBLPre�oProduto("PRE�O DE VENDA")
            End If
        Else
            SearchAdvancedProduto = TBLPre�oProduto("PRE�O DE VENDA")
        End If
    ElseIf Tipo = vbLote Then
        TBLProduto.Seek "=", C�digo
        If TBLProduto.NoMatch Then
            SearchAdvancedProduto = False
        Else
            SearchAdvancedProduto = TBLProduto("LOTES")
        End If
    ElseIf Tipo = vbDescontoM�ximo Then
        TBLProduto.Seek "=", C�digo
        If TBLProduto.NoMatch Then
            SearchAdvancedProduto = False
        Else
            SearchAdvancedProduto = TBLProduto("DESCONTO M�XIMO")
        End If
    ElseIf Tipo = vbTributo Then
        TBLProduto.Seek "=", C�digo
        If TBLProduto.NoMatch Then
            SearchAdvancedProduto = Empty
        Else
            TBLTipoDeICM.Seek "=", TBLProduto("TIPO DE ICM")
            If TBLTipoDeICM.NoMatch Then
                SearchAdvancedProduto = Empty
            Else
                SearchAdvancedProduto = TBLTipoDeICM("C�DIGO DO PDV")
            End If
        End If
    End If
End Function
Public Function SearchCliente(ByVal Busca As String, ByVal CampoDeBusca As Byte) As String
    Dim TBLCliente As Table
    Dim ClienteAberto As Boolean

    ClienteAberto = AbreTabela(Dicion�rio, "CADASTRO", "CLIENTE", DBCadastro, TBLCliente, TBLTabela, dbOpenTable)
    
    If ClienteAberto Then
        If CampoDeBusca = byCodigo Then
            TBLCliente.Index = "CLIENTE1"
        ElseIf CampoDeBusca = byNome Then
            TBLCliente.Index = "CLIENTE2"
        ElseIf CampoDeBusca = byCGCCPF Then
            TBLCliente.Index = "CLIENTE3"
        End If
    Else
        MsgBox "N�o consegui abrir a tabela 'Cliente' !", vbCritical, "Erro"
        Exit Function
    End If
    
    TBLCliente.Seek "=", Busca
    
    If TBLCliente.NoMatch Then
        SearchCliente = Empty
    Else
        SearchCliente = TBLCliente("NOME - RAZ�O SOCIAL")
    End If
    
    TBLCliente.Close
End Function
Public Function SearchFornecedor(ByVal mCGCCPF As String) As String
    Dim TBLFornecedor As Table
    Dim FornecedorAberto As Boolean
    Dim IndiceFornecedorAtivo$
    
    FornecedorAberto = AbreTabela(Dicion�rio, "CADASTRO", "FORNECEDOR", DBCadastro, TBLFornecedor, TBLTabela, dbOpenTable)
    
    If FornecedorAberto Then
        IndiceFornecedorAtivo = "FORNECEDOR1"
        TBLFornecedor.Index = IndiceFornecedorAtivo
    Else
        MsgBox "N�o consegui abrir a tabela 'Fornecedor' !", vbCritical, "Erro"
        Exit Function
    End If
    
    TBLFornecedor.Seek "=", mCGCCPF
    
    If TBLFornecedor.NoMatch Then
        MsgBox "Fornecedor " & mCGCCPF & " n�o foi encontrado!", vbCritical, "Erro"
        Exit Function
    Else
        SearchFornecedor = TBLFornecedor("RAZ�O SOCIAL")
    End If
    TBLFornecedor.Close
End Function
Public Function SearchProduto(ByVal C�digo)
    Dim TBLProduto As Table
    Dim ProdutoAberto As Boolean
    Dim IndiceProdutoAtivo$
    
    ProdutoAberto = AbreTabela(Dicion�rio, "CADASTRO", "PRODUTO", DBCadastro, TBLProduto, TBLTabela, dbOpenTable)
    
    If ProdutoAberto Then
        IndiceProdutoAtivo = "PRODUTO1"
        TBLProduto.Index = IndiceProdutoAtivo
    Else
        MsgBox "N�o consegui abrir a tabela 'Produto' !", vbCritical, "Erro"
        SearchProduto = ""
        Exit Function
    End If
    
    TBLProduto.Seek "=", C�digo
    
    If TBLProduto.NoMatch Then
        MsgBox "N�o foi poss�vel encontrar o produto " & C�digo, vbCritical, "Erro"
        SearchProduto = ""
        TBLProduto.Close
        Exit Function
    Else
        SearchProduto = TBLProduto("DESCRI��O")
    End If
    
    TBLProduto.Close
End Function
Public Sub SetaFocus(ByRef Janela As Form)
    On Error Resume Next
    Janela.SetFocus
End Sub
Public Sub SetRegistryString(ByVal vsSection As String, ByVal vsSubSection, ByVal vsKey As String, ByVal vsSetting As String)
    vsSection = AddPath(gAPPNAME, vsSection)
    vsSection = AddPath(vsSection, vsSubSection)
    SaveSetting APP_CATEGORY, vsSection, vsKey, vsSetting
End Sub
Public Function StatusOkECF() As Boolean
    Dim Status As String
    
    Status = VerStatusECF

    If Mid(Status, 1, 2) = ".-" Then
        StatusOkECF = False
    Else
        StatusOkECF = True
    End If
End Function
Public Sub TestaInferior(ByRef TBLWork As Table, ByVal lAllowEdit As Boolean, ByVal lAllowDelete As Boolean, ByVal lAllowConsult As Boolean)
    On Error Resume Next
    
    mdiGeal.StatusBar.Panels("Posi��o").Visible = True
    mdiGeal.StatusBar.Panels("Posi��o").Text = "Registros: " & TBLWork.RecordCount
    ResizeStatusBar

    If TBLWork.RecordCount = 0 Then
        Navega��oInferior False
        Bot�oExcluir False
        Bot�oGravar False
        Exit Sub
    End If
    Bot�oExcluir lAllowDelete
    Bot�oGravar True
    TBLWork.MovePrevious
    If TBLWork.BOF Then
        Navega��oInferior False
    Else
        Navega��oInferior lAllowConsult
    End If
    TBLWork.MoveNext
End Sub
Public Sub TestaSuperior(ByRef TBLWork As Table, ByVal lAllowEdit As Boolean, ByVal lAllowDelete As Boolean, ByVal lAllowConsult As Boolean)
    On Error Resume Next
    
    mdiGeal.StatusBar.Panels("Posi��o").Visible = True
    mdiGeal.StatusBar.Panels("Posi��o").Text = "Registros: " & TBLWork.RecordCount
    ResizeStatusBar
    
    If TBLWork.RecordCount = 0 Then
        Navega��oSuperior False
        Bot�oExcluir False
        Bot�oGravar False
        Exit Sub
    End If
    Bot�oExcluir lAllowDelete
    Bot�oGravar True
    TBLWork.MoveNext
    If TBLWork.EOF Then
        Navega��oSuperior False
    Else
        Navega��oSuperior lAllowConsult
    End If
    TBLWork.MovePrevious
End Sub
Public Sub TestaInferiorArray(ByVal Elemento, ByRef Matriz(), ByVal lAllowEdit As Boolean, ByVal lAllowDelete As Boolean, ByVal lAllowConsult As Boolean, Optional ByVal Dimens As Byte = 1)
    mdiGeal.StatusBar.Panels("Posi��o").Text = "Registros: " & Elemento
    ResizeStatusBar
    
    If Elemento = 0 Then
        Navega��oInferior False
        Bot�oExcluir False
        Bot�oGravar False
        Exit Sub
    End If
    
    If UBound(Matriz, Dimens) = 0 Then
        Navega��oInferior False
        Bot�oExcluir False
        Bot�oGravar False
        Exit Sub
    End If
    Bot�oExcluir lAllowDelete
    Bot�oGravar True
    If Elemento = 1 Then
        Navega��oInferior False
    Else
        Navega��oInferior lAllowConsult
    End If
End Sub
Public Sub TestaSuperiorArray(ByVal Elemento, ByRef Matriz(), ByVal lAllowEdit As Boolean, ByVal lAllowDelete As Boolean, ByVal lAllowConsult As Boolean, Optional ByVal Dimens As Byte = 1)
    mdiGeal.StatusBar.Panels("Posi��o").Text = "Registros: " & Elemento
    ResizeStatusBar
    
    If Elemento = 0 Then
        Navega��oInferior False
        Bot�oExcluir False
        Bot�oGravar False
        Exit Sub
    End If
    If UBound(Matriz, Dimens) = 0 Then
        Navega��oSuperior False
        Bot�oExcluir False
        Bot�oGravar False
        Exit Sub
    End If
    Bot�oExcluir lAllowDelete
    Bot�oGravar True
    If Elemento = UBound(Matriz, Dimens) Then
        Navega��oSuperior False
    Else
        Navega��oSuperior lAllowConsult
    End If
End Sub
Public Function TotalizarCupomFiscal(ByVal Total$) As Boolean
    If PDV.TotalizarCupomFiscal(Chr(27) & Chr(46) & "1001" & Total & "}") Then
        TotalizarCupomFiscal = True
    Else
        TotalizarCupomFiscal = False
    End If
End Function
Private Function ValidaUsu�rio(ByVal WindowsTop As Long, WindowsHeight As Long) As Boolean
    'Valida Usu�rio
    frmValidaUsu�rio.GravaUsu�rio = True
    frmValidaUsu�rio.WindowTop = WindowsTop
    frmValidaUsu�rio.WindowHeight = WindowsHeight
    frmValidaUsu�rio.Show vbModeless
        
    Do While Not frmValidaUsu�rio.Fechado
        DoEvents
    Loop
    
    gUsu�rio = Trim(frmValidaUsu�rio.Usu�rio)
    
    Set frmValidaUsu�rio = Nothing
    
    If gUsu�rio = "" Then
        ValidaUsu�rio = False
    Else
        ValidaUsu�rio = True
    End If
End Function
Public Function VerStatusECF() As String
    Dim Status As String
    
    Status = Space(255)
    
    PDV.VerStatusECF Status, 255
    
    Status = StripTerminator(Status)
    
    VerStatusECF = Status
End Function
Private Sub Vis�oTotal()
    Dim lAcesso As Boolean
    lAcesso = True
    With mdiGeal
        .mnuArquivoAbrirAgencia.Visible = lAcesso
        .mnuArquivoAbrirBanco.Visible = lAcesso
        .mnuArquivoAbrirCliente.Visible = lAcesso
        .mnuArquivoAbrirContaCorrente.Visible = lAcesso
        .mnuArquivoAbrirFornecedor.Visible = lAcesso
        .mnuArquivoAbrirFuncion�rio.Visible = lAcesso
        .mnuArquivoAbrirProduto.Visible = lAcesso
        .mnuArquivoAbrirDespesas.Visible = lAcesso
        .mnuArquivoAbrir.Visible = lAcesso
        
        .mnuMovimentoEntradaCompra.Visible = lAcesso
        .mnuMovimentoEntradaDevolu��oTroca.Visible = lAcesso
        .mnuMovimentoEntrada.Visible = lAcesso
        
        .mnuMovimentoSa�daVenda.Visible = lAcesso
        .mnuMovimentoSa�daDevolu��oTroca.Visible = lAcesso
        .mnuMovimentoSa�da.Visible = lAcesso
        
        .mnuMovimentoMovimentoDi�rio.Visible = lAcesso
        .mnuMovimentoContaCorrente.Visible = lAcesso
        .mnuMovimentoCaixa.Visible = lAcesso
        .mnuMovimentoCaixaF�cil.Visible = lAcesso
        .mnuMovimentoDespesas.Visible = lAcesso
        .mnuMovimento.Visible = lAcesso
        
        .mnuPar�metrosDepartamento.Visible = lAcesso
        .mnuPar�metrosSe��o.Visible = lAcesso
        .mnuPar�metrosDepartamentoSe��o.Visible = lAcesso
        .mnuPar�metrosTipodeICM.Visible = lAcesso
        .mnuPar�metrosTipodeEmbalagem.Visible = lAcesso
        .mnuPar�metrosUnidades.Visible = lAcesso
        .mnuPar�metrosLocalidadeDeEstoque.Visible = lAcesso
        .mnuPar�metrosPlanoDePagamento.Visible = lAcesso
        .mnuPar�metros.Visible = lAcesso
        
        .mnuPar�metrosUsu�rios.Visible = lAcesso
        .mnuPar�metrosGrupos.Visible = lAcesso
        
        .mnuPar�metroSep2.Visible = lAcesso
        .mnuSep8.Visible = lAcesso
        .mnuSep9.Visible = lAcesso
        .mnuSep20.Visible = True
        .mnuSep11.Visible = lAcesso
        .mnuSep12.Visible = lAcesso
        .mnuSep16.Visible = lAcesso
    End With
End Sub
Private Sub VisibleArquivo()
    If Not mdiGeal.mnuArquivoAbrirAgencia.Enabled Then
        mdiGeal.mnuArquivoAbrirAgencia.Visible = False
    End If
    If Not mdiGeal.mnuArquivoAbrirBanco.Enabled Then
        mdiGeal.mnuArquivoAbrirBanco.Visible = False
    End If
    If Not mdiGeal.mnuArquivoAbrirCliente.Enabled Then
        mdiGeal.mnuArquivoAbrirCliente.Visible = False
    End If
    If Not mdiGeal.mnuArquivoAbrirContaCorrente.Enabled Then
        mdiGeal.mnuArquivoAbrirContaCorrente.Visible = False
    End If
    If Not mdiGeal.mnuArquivoAbrirFornecedor.Enabled Then
        mdiGeal.mnuArquivoAbrirFornecedor.Visible = False
    End If
    If Not mdiGeal.mnuArquivoAbrirFuncion�rio.Enabled Then
        mdiGeal.mnuArquivoAbrirFuncion�rio.Visible = False
    End If
    If Not mdiGeal.mnuArquivoAbrirProduto.Enabled Then
        mdiGeal.mnuArquivoAbrirProduto.Visible = False
    End If
    If Not mdiGeal.mnuArquivoAbrirDespesas.Enabled Then
        mdiGeal.mnuArquivoAbrirDespesas.Visible = False
    End If
End Sub
Private Sub VisibleMovimento()
    If mdiGeal.mnuMovimentoEntrada.Visible Then
        If Not mdiGeal.mnuMovimentoEntradaCompra.Enabled Then
            mdiGeal.mnuMovimentoEntradaCompra.Visible = False
        End If
        If Not mdiGeal.mnuMovimentoEntradaDevolu��oTroca.Enabled Then
            mdiGeal.mnuMovimentoEntradaDevolu��oTroca.Visible = False
        End If
    End If
    If mdiGeal.mnuMovimentoSa�da.Visible Then
        If Not mdiGeal.mnuMovimentoSa�daVenda.Enabled Then
            mdiGeal.mnuMovimentoSa�daVenda.Visible = False
        End If
        If Not mdiGeal.mnuMovimentoSa�daDevolu��oTroca.Enabled Then
            mdiGeal.mnuMovimentoSa�daDevolu��oTroca.Visible = False
        End If
    End If
    If Not mdiGeal.mnuMovimentoMovimentoDi�rio.Enabled Then
        mdiGeal.mnuMovimentoMovimentoDi�rio.Visible = False
    End If
    If Not mdiGeal.mnuMovimentoContaCorrente.Enabled Then
        mdiGeal.mnuMovimentoContaCorrente.Visible = False
    End If
    If Not mdiGeal.mnuMovimentoCaixa.Enabled Then
        mdiGeal.mnuMovimentoCaixa.Visible = False
    End If
    If Not mdiGeal.mnuMovimentoCaixaF�cil.Enabled Then
        mdiGeal.mnuMovimentoCaixaF�cil.Visible = False
    End If
    If Not mdiGeal.mnuMovimentoDespesas.Enabled Then
        mdiGeal.mnuMovimentoDespesas.Visible = False
    End If
End Sub
Private Sub VisiblePar�metros()
    If Not mdiGeal.mnuPar�metrosDepartamento.Enabled Then
        mdiGeal.mnuPar�metrosDepartamento.Visible = False
    End If
    If Not mdiGeal.mnuPar�metrosSe��o.Enabled Then
        mdiGeal.mnuPar�metrosSe��o.Visible = False
    End If
    If Not mdiGeal.mnuPar�metrosDepartamentoSe��o.Enabled Then
        mdiGeal.mnuPar�metrosDepartamentoSe��o.Visible = False
    End If
    If Not mdiGeal.mnuPar�metrosTipodeICM.Enabled Then
        mdiGeal.mnuPar�metrosTipodeICM.Visible = False
    End If
    If Not mdiGeal.mnuPar�metrosTipodeEmbalagem.Enabled Then
        mdiGeal.mnuPar�metrosTipodeEmbalagem.Visible = False
    End If
    If Not mdiGeal.mnuPar�metrosUnidades.Enabled Then
        mdiGeal.mnuPar�metrosUnidades.Visible = False
    End If
    If Not mdiGeal.mnuPar�metrosLocalidadeDeEstoque.Enabled Then
        mdiGeal.mnuPar�metrosLocalidadeDeEstoque.Visible = False
    End If
    If Not mdiGeal.mnuPar�metrosPlanoDePagamento.Enabled Then
        mdiGeal.mnuPar�metrosPlanoDePagamento.Visible = False
    End If
End Sub
