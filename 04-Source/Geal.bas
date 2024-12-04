Attribute VB_Name = "GealFunc"
Option Explicit

Global PDV          As Object
Global Const lFalse As Byte = 1
Global Const lTrue  As Byte = 0

Global gEmpresa      As String
Global gCGC          As String

Global glPortaAberta As Boolean

Global Dicionário$
Global WS               As Workspace
Global DBDGeal          As Database
Global TBLArquivo       As Table, TBLTabela As Table, TBLCampo As Table, TBLIndice As Table
Global DicionárioAberto As Boolean

Global DBCadastro        As Database
Global DBDCadastroAberto As Boolean

Global DBUsuário        As Database
Global DBDUsuárioAberto As Boolean

Global DBFinanceiro        As Database
Global DBDFinanceiroAberto As Boolean

Global DBUtilitário        As Database
Global DBDUtilitárioAberto As Boolean

Global DBSistema        As Database
Global DBDSistemaAberto As Boolean

Global gUsuário$

Global gEstado As Byte

Global AplicaçãoPath$

Global Const vbDescrição          As Byte = 1
Global Const vbCódigo             As Byte = 2
Global Const vbValorUnitário      As Byte = 3
Global Const vbValValorUnitário   As Byte = 4
Global Const vbCódigoDoFornecedor As Byte = 5
Global Const vbLote               As Byte = 6
Global Const vbDescontoMáximo     As Byte = 7
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
Public Function AbreBaseDeDados(ByVal lInício As Boolean, ByVal lExclusivo As Boolean) As Boolean
    Dim Cont As Long
    Dim NomeDaEmpresa As String
    Dim TBLParâmetros As Table
    Dim ParâmetrosAberto As Boolean
    
    If lInício Then
        'Path da aplicação
        AplicaçãoPath = App.Path
        
        Set WS = DBEngine.Workspaces(0)
    End If
    
    If lInício Then
        'Abre o Dicionário de Dados
        
        NomeDaEmpresa = GetRegistryString("Geal", "Geral", "Empresa")
        frmSplash.lblLicenseTo = NomeDaEmpresa
        
        Dicionário = GetRegistryString("Geal", "Geral", "Dicionário")
          
        frmSplash.lblWarning = "Abrindo... dicionário de dados - " & Dicionário
        frmSplash.lblWarning.Refresh
        
        If Dicionário = Empty Then
            MsgBox "Erro na abertura do Dicionário de Dados!", vbExclamation, "Erro"
            AbreBaseDeDados = False
            Exit Function
        End If
        
        DicionárioAberto = AbreDicionário(WS, Dicionário, DBDGeal, TBLArquivo, TBLTabela, TBLCampo, TBLIndice)
        
        If Not DicionárioAberto Then
            MsgBox "Dicionário " + Dicionário + " não foi aberto !", vbExclamation, "Erro "
            AbreBaseDeDados = False
            Exit Function
        End If
    End If
    
    'Abre Cadastro
    If lInício Then
        frmSplash.lblWarning = "Abrindo... base de dados - CADASTRO"
        frmSplash.lblWarning.Refresh
    End If
    
    DBDCadastroAberto = AbreArquivo(WS, Dicionário, "CADASTRO", DBCadastro, TBLArquivo, lExclusivo)
    
    If Not DBDCadastroAberto Then
        MsgBox "Erro na abertura do arquivo 'CADASTRO' ", vbExclamation, "Erro"
        AbreBaseDeDados = False
        Exit Function
    End If
    
    'Abre Usuário
    If lInício Then
        frmSplash.lblWarning = "Abrindo... base de dados - USUÁRIO"
        frmSplash.lblWarning.Refresh
    End If
    
    DBDUsuárioAberto = AbreArquivo(WS, Dicionário, "USUÁRIO", DBUsuário, TBLArquivo, lExclusivo)
    
    If Not DBDUsuárioAberto Then
        MsgBox "Erro na abertura do arquivo 'USUÁRIO' ", vbExclamation, "Erro"
        AbreBaseDeDados = False
        Exit Function
    End If
    
    'Abre Financeiro
    If lInício Then
        frmSplash.lblWarning = "Abrindo... base de dados - FINANCEIRO"
        frmSplash.lblWarning.Refresh
    End If
    
    DBDFinanceiroAberto = AbreArquivo(WS, Dicionário, "FINANCEIRO", DBFinanceiro, TBLArquivo, lExclusivo)
    
    If Not DBDFinanceiroAberto Then
        MsgBox "Erro na abertura do arquivo 'FINANCEIRO' ", vbExclamation, "Erro"
        AbreBaseDeDados = False
        Exit Function
    End If
    
    'Abre Utilitário
    If lInício Then
        frmSplash.lblWarning = "Abrindo... base de dados - UTILITÁRIO"
        frmSplash.lblWarning.Refresh
    End If
    
    DBDUtilitárioAberto = AbreArquivo(WS, Dicionário, "UTILITÁRIO", DBUtilitário, TBLArquivo, lExclusivo)
    
    If Not DBDUtilitárioAberto Then
        MsgBox "Erro na abertura do arquivo 'UTILITÁRIO' ", vbExclamation, "Erro"
        AbreBaseDeDados = False
        Exit Function
    End If
    
    'Abre Sistema
    If lInício Then
        frmSplash.lblWarning = "Abrindo... base de dados - SISTEMA"
        frmSplash.lblWarning.Refresh
    End If
    
    DBDSistemaAberto = AbreArquivo(WS, Dicionário, "SISTEMA", DBSistema, TBLArquivo, lExclusivo)
    
    If Not DBDSistemaAberto Then
        MsgBox "Erro na abertura do arquivo 'SISTEMA' ", vbExclamation, "Erro"
        AbreBaseDeDados = False
        Exit Function
    End If
    
    'Pega o nome da Empresa e o CGC
    ParâmetrosAberto = AbreTabela(Dicionário, "SISTEMA", "PARÂMETROS", DBSistema, TBLParâmetros, TBLTabela, dbOpenTable)
    
    If ParâmetrosAberto Then
    Else
        MsgBox "Não consegui abrir a tabela 'Parâmetros' !", vbCritical, "Erro"
        AbreBaseDeDados = False
        Exit Function
    End If
    
    gEmpresa = TBLParâmetros("EMPRESA")
    gCGC = TBLParâmetros("CGC")
    
    TBLParâmetros.Close
    
    AbreBaseDeDados = True
    
    If lInício Then
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
        .mnuArquivoAbrirFuncionário.Enabled = lAcesso
        .mnuArquivoAbrirProduto.Enabled = lAcesso
        .mnuArquivoAbrirDespesas.Enabled = lAcesso
        .mnuArquivoAbrir.Visible = lAcesso
        
        .mnuMovimentoEntradaCompra.Enabled = lAcesso
        .mnuMovimentoEntradaDevoluçãoTroca.Enabled = lAcesso
        .mnuMovimentoEntrada.Visible = lAcesso
        
        .mnuMovimentoSaídaVenda.Enabled = lAcesso
        .mnuMovimentoSaídaDevoluçãoTroca.Enabled = lAcesso
        .mnuMovimentoSaída.Visible = lAcesso
        
        .mnuMovimentoMovimentoDiário.Enabled = lAcesso
        .mnuMovimentoContaCorrente.Enabled = lAcesso
        .mnuMovimentoCaixa.Enabled = lAcesso
        .mnuMovimentoCaixaFácil.Enabled = lAcesso
        .mnuMovimentoDespesas.Enabled = lAcesso
        .mnuMovimento.Visible = lAcesso
        
        .mnuParâmetrosDepartamento.Enabled = lAcesso
        .mnuParâmetrosSeção.Enabled = lAcesso
        .mnuParâmetrosDepartamentoSeção.Enabled = lAcesso
        .mnuParâmetrosTipodeICM.Enabled = lAcesso
        .mnuParâmetrosTipodeEmbalagem.Enabled = lAcesso
        .mnuParâmetrosUnidades.Enabled = lAcesso
        .mnuParâmetrosLocalidadeDeEstoque.Enabled = lAcesso
        .mnuParâmetrosPlanoDePagamento.Enabled = lAcesso
        .mnuParâmetrosCaixa.Visible = lAcesso
        .mnuParâmetros.Visible = lAcesso
        
        .mnuParâmetrosUsuários.Visible = lAcesso
        .mnuParâmetrosGrupos.Visible = lAcesso
        .mnuParâmetrosSenhaDoSistema.Visible = lAcesso
        
        .mnuUtilitáriosConsultaSQL.Visible = lAcesso
        
        .mnuSep8.Visible = lAcesso
        .mnuSep9.Visible = lAcesso
        .mnuSep20.Visible = lAcesso
        .mnuSep11.Visible = lAcesso
        .mnuSep12.Visible = lAcesso
        .mnuSep16.Visible = lAcesso
        .mnuSep18.Visible = lAcesso
        .mnuParâmetroSep2.Visible = lAcesso
        .mnuParâmetroSep3.Visible = lAcesso
        .mnuParâmetroSep4.Visible = lAcesso
    End With
End Sub
Public Sub AllBotões(ByVal Valor As Boolean)
    BotãoAtualizar Valor
    BotãoIncluir Valor
    BotãoExcluir Valor
    BotãoGravar Valor
    BotãoImprimir Valor
    NavegaçãoInferior Valor
    NavegaçãoSuperior Valor
    BarraDeStatus "Pronto"
End Sub
Public Function Allow(ByVal Categoria As String, ByVal Direito As String, Optional ByVal Usuário As String) As Boolean
    On Error GoTo Erro
    
    Dim lRetorno As Boolean
    Dim TBLGrupos As Table
    Dim GruposAberto As Boolean
    Dim IndiceGruposAtivo$
    
    Dim TBLUsuárioGrupo As Table
    Dim UsuárioGrupoAberto As Boolean
    Dim IndiceUsuárioGrupoAtivo$
    
    If IsMissing(Usuário) Or Usuário = Empty Then
        Usuário = gUsuário
    End If
    
    If Usuário = "ADMIN" Then
        Allow = True
    End If
    
    GruposAberto = AbreTabela(Dicionário, "USUÁRIO", "GRUPO", DBUsuário, TBLGrupos, TBLTabela, dbOpenTable)
    
    If GruposAberto Then
        IndiceGruposAtivo = "GRUPO1"
        TBLGrupos.Index = IndiceGruposAtivo
    Else
        MsgBox "Não consegui abrir a tabela 'GRUPO' !", vbCritical, "Erro"
        Exit Function
    End If
    
    UsuárioGrupoAberto = AbreTabela(Dicionário, "USUÁRIO", "USUÁRIO - GRUPO", DBUsuário, TBLUsuárioGrupo, TBLTabela, dbOpenTable)
    
    If UsuárioGrupoAberto Then
        IndiceUsuárioGrupoAtivo = "USUÁRIOGRUPO1"
        TBLUsuárioGrupo.Index = IndiceUsuárioGrupoAtivo
    Else
        MsgBox "Não consegui abrir a tabela 'GRUPO' !", vbCritical, "Erro"
        Exit Function
    End If
    
    TBLUsuárioGrupo.Seek "=", Usuário
    
    If TBLUsuárioGrupo.NoMatch Then
        Exit Function
    End If
    
    TBLUsuárioGrupo.Seek "=", Usuário
    
    If TBLUsuárioGrupo.NoMatch Then
        GoTo Fim
    End If
    
    lRetorno = False
    
    Do While Trim(TBLUsuárioGrupo("USERNAME")) = Usuário
        TBLGrupos.Seek "=", TBLUsuárioGrupo("CÓDIGO DO GRUPO")
        If InStr(TBLGrupos(Categoria), Direito) Then
            lRetorno = True
        End If
        TBLUsuárioGrupo.MoveNext
        If TBLUsuárioGrupo.EOF Then
            Exit Do
        End If
    Loop
    
    Allow = lRetorno
    
Fim:
    If GruposAberto Then
        TBLGrupos.Close
    End If
    If UsuárioGrupoAberto Then
        TBLUsuárioGrupo.Close
    End If
    
    Exit Function
Erro:
    GeraMensagemDeErro "AllowInsert - Usuário:" & Usuário
    Allow = False
    If GruposAberto Then
        TBLGrupos.Close
    End If
    If UsuárioGrupoAberto Then
        TBLUsuárioGrupo.Close
    End If
End Function
Public Function AtualizaLote(ByVal CódigoDoProduto As Long, ByVal CódigoDoLote As String, ByVal DígitoDoLote As String, ByVal Quantidade As Single, ByVal Múltiplo As Single) As Boolean
    On Error GoTo Erro
    
    Dim TBLLote As Table
    Dim LoteAberto As Boolean
    
    'Abre tabela PRODUTO
    LoteAberto = AbreTabela(Dicionário, "CADASTRO", "LOTE DO PRODUTO", DBCadastro, TBLLote, TBLTabela, dbOpenTable)
    
    If LoteAberto Then
        TBLLote.Index = "LOTEDOPRODUTO1"
    Else
        MsgBox "Não consegui abrir a tabela 'Lote do Produto' !", vbCritical, "Erro"
        Exit Function
    End If
    
    TBLLote.Seek "=", CódigoDoProduto, CódigoDoLote, DígitoDoLote
    
    If TBLLote.NoMatch Then
        MsgBox "Produto: " & CódigoDoProduto & vbCr & "Lote: " & CódigoDoLote & "-" & DígitoDoLote & vbCr & "Não foi encontrado!", vbInformation, "Lote não existe!"
        AtualizaLote = False
    Else
        If TBLLote("QUANTIDADE") - (Quantidade * Múltiplo) = 0 Then
            TBLLote.Delete
        Else
            TBLLote.Edit
            TBLLote("QUANTIDADE") = TBLLote("QUANTIDADE") - (Quantidade * Múltiplo)
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
Public Function AtualizaProduto(ByVal Código As Long, ByVal Operação, ByVal Valor As Long) As Boolean
    On Error GoTo Erro
    
    Dim TBLProduto As Table
    Dim ProdutoAberto As Boolean
    
    'Abre tabela PRODUTO
    ProdutoAberto = AbreTabela(Dicionário, "CADASTRO", "PRODUTO", DBCadastro, TBLProduto, TBLTabela, dbOpenTable)
    
    If ProdutoAberto Then
        TBLProduto.Index = "PRODUTO1"
    Else
        MsgBox "Não consegui abrir a tabela 'Produto' !", vbCritical, "Erro"
        Exit Function
    End If
    
    TBLProduto.Seek "=", Código
    
    If TBLProduto.NoMatch Then
        AtualizaProduto = False
    End If
    
    TBLProduto.Edit
    If Operação = "+" Then
        TBLProduto("QUANTIDADE") = TBLProduto("QUANTIDADE") + Valor
    ElseIf Operação = "-" Then
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
Public Sub BotãoExcluir(ByVal Valor As Boolean)
    mdiGeal.mnuEditarExcluir.Enabled = Valor
    mdiGeal.Toolbar.Buttons("Excluir").Enabled = Valor
End Sub
Public Sub BotãoGravar(ByVal Valor As Boolean)
    mdiGeal.mnuArquivoSalvar.Enabled = Valor
    mdiGeal.Toolbar.Buttons("Gravar").Enabled = Valor
End Sub
Public Sub BotãoImprimir(ByVal Valor As Boolean)
    mdiGeal.mnuArquivoImprimir.Enabled = Valor
    mdiGeal.Toolbar.Buttons("Imprimir").Enabled = Valor
End Sub
Public Sub BotãoAtualizar(ByVal Valor As Boolean)
    mdiGeal.mnuEditarAtualizar.Enabled = Valor
End Sub
Public Sub BotãoIncluir(ByVal Valor As Boolean)
    mdiGeal.mnuEditarIncluir.Enabled = Valor
    mdiGeal.Toolbar.Buttons("Incluir").Enabled = Valor
End Sub
Public Function BuscaFuncionário(ByVal Código&) As String
    Dim FuncionárioAberto As Boolean, TBLFuncionário As Table, IndiceFuncionárioAtivo$
    
    FuncionárioAberto = AbreTabela(Dicionário, "USUÁRIO", "FUNCIONÁRIO", DBUsuário, TBLFuncionário, TBLTabela, dbOpenTable)
    
    If FuncionárioAberto Then
        IndiceFuncionárioAtivo = "FUNCIONÁRIO1"
        TBLFuncionário.Index = IndiceFuncionárioAtivo
    Else
        MsgBox "Não consegui abrir a tabela 'Funcionário' !", vbCritical, "Erro"
        BuscaFuncionário = ""
        Exit Function
    End If
    
    TBLFuncionário.Seek "=", Código
    
    If TBLFuncionário.NoMatch Then
        BuscaFuncionário = ""
    Else
        BuscaFuncionário = TBLFuncionário("NOME")
    End If
End Function
Public Function CancelarCupom() As Boolean
    If PDV.CancelarCupom(Chr(27) & Chr(46) & "05}") Then
        CancelarCupom = True
    Else
        CancelarCupom = False
    End If
End Function
Public Sub ChamaConfigurações(ByVal Usuário$)
    On Error GoTo Erro
    
    Dim TBLGrupos As Table
    Dim GruposAberto As Boolean
    Dim IndiceGruposAtivo$
    
    Dim TBLUsuárioGrupo As Table
    Dim UsuárioGrupoAberto As Boolean
    Dim IndiceUsuárioGrupoAtivo$
    
    VisãoTotal
    AcessoNegado False
    
    If Usuário = "ADMIN" Then
        Exit Sub
    End If
    
    AcessoNegado True
    
    GruposAberto = AbreTabela(Dicionário, "USUÁRIO", "GRUPO", DBUsuário, TBLGrupos, TBLTabela, dbOpenTable)
    
    If GruposAberto Then
        IndiceGruposAtivo = "GRUPO1"
        TBLGrupos.Index = IndiceGruposAtivo
    Else
        MsgBox "Não consegui abrir a tabela 'GRUPO' !", vbCritical, "Erro"
        Exit Sub
    End If
    
    UsuárioGrupoAberto = AbreTabela(Dicionário, "USUÁRIO", "USUÁRIO - GRUPO", DBUsuário, TBLUsuárioGrupo, TBLTabela, dbOpenTable)
    
    If UsuárioGrupoAberto Then
        IndiceUsuárioGrupoAtivo = "USUÁRIOGRUPO1"
        TBLUsuárioGrupo.Index = IndiceUsuárioGrupoAtivo
    Else
        MsgBox "Não consegui abrir a tabela 'GRUPO' !", vbCritical, "Erro"
        Exit Sub
    End If
    
    TBLUsuárioGrupo.Seek "=", Usuário
    
    If TBLUsuárioGrupo.NoMatch Then
        Exit Sub
    End If
    
    TBLUsuárioGrupo.Seek "=", Trim(Usuário)
    
    If TBLUsuárioGrupo.NoMatch Then
        GoTo Fim
    End If
    
    Do While Trim(TBLUsuárioGrupo("USERNAME")) = Trim(Usuário)
        TBLGrupos.Seek "=", TBLUsuárioGrupo("CÓDIGO DO GRUPO")
        
        'Início Arquivo
        'Agência
        If TBLGrupos("AGÊNCIA") <> Empty Then
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
        'Funcionário
        If TBLGrupos("FUNCIONÁRIO") <> Empty Then
            mdiGeal.mnuArquivoAbrir.Visible = True
            mdiGeal.mnuArquivoAbrirFuncionário.Visible = True
            mdiGeal.mnuArquivoAbrirFuncionário.Enabled = True
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
        
        'Início Movimento
        'Compra
        If TBLGrupos("COMPRA") <> Empty Then
            mdiGeal.mnuMovimento.Visible = True
            mdiGeal.mnuMovimentoEntrada.Visible = True
            mdiGeal.mnuMovimentoEntradaCompra.Visible = True
            mdiGeal.mnuMovimentoEntradaCompra.Enabled = True
            VisibleMovimento
        End If
        'Devolução/Troca (Compra)
        If TBLGrupos("DEVOLUÇÃO/TROCA (COMPRA)") <> Empty Then
            mdiGeal.mnuMovimento.Visible = True
            mdiGeal.mnuMovimentoEntrada.Visible = True
            mdiGeal.mnuMovimentoEntradaDevoluçãoTroca.Visible = True
            mdiGeal.mnuMovimentoEntradaDevoluçãoTroca.Enabled = True
            VisibleMovimento
        End If
        'Venda
        If TBLGrupos("VENDA") <> Empty Then
            mdiGeal.mnuMovimento.Visible = True
            mdiGeal.mnuMovimentoSaída.Visible = True
            mdiGeal.mnuMovimentoSaídaVenda.Visible = True
            mdiGeal.mnuMovimentoSaídaVenda.Enabled = True
            VisibleMovimento
        End If
        'Devolução/Troca (Venda)
        If TBLGrupos("DEVOLUÇÃO/TROCA (VENDA)") <> Empty Then
            mdiGeal.mnuMovimento.Visible = True
            mdiGeal.mnuMovimentoSaída.Visible = True
            mdiGeal.mnuMovimentoSaídaDevoluçãoTroca.Visible = True
            mdiGeal.mnuMovimentoSaídaDevoluçãoTroca.Enabled = True
            VisibleMovimento
        End If
        'Movimento Diário
        If TBLGrupos("MOVIMENTO DIÁRIO") <> Empty Then
            mdiGeal.mnuMovimento.Visible = True
            mdiGeal.mnuMovimentoMovimentoDiário.Visible = True
            mdiGeal.mnuMovimentoMovimentoDiário.Enabled = True
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
        'Caixa Fácil
        If TBLGrupos("CAIXA FÁCIL") <> Empty Then
            mdiGeal.mnuMovimento.Visible = True
            mdiGeal.mnuMovimentoCaixaFácil.Visible = True
            mdiGeal.mnuMovimentoCaixaFácil.Enabled = True
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
        
        'Início Parâmetros
        'Departamento
        If TBLGrupos("DEPARTAMENTO") <> Empty Then
            mdiGeal.mnuParâmetros.Visible = True
            mdiGeal.mnuParâmetrosDepartamento.Visible = True
            mdiGeal.mnuParâmetrosDepartamento.Enabled = True
            VisibleParâmetros
        End If
        'Seção
        If TBLGrupos("SEÇÃO") <> Empty Then
            mdiGeal.mnuParâmetros.Visible = True
            mdiGeal.mnuParâmetrosSeção.Visible = True
            mdiGeal.mnuParâmetrosSeção.Enabled = True
            VisibleParâmetros
        End If
        'Departamento - Seção
        If TBLGrupos("DEPARTAMENTO - SEÇÃO") <> Empty Then
            mdiGeal.mnuParâmetros.Visible = True
            mdiGeal.mnuParâmetrosDepartamentoSeção.Visible = True
            mdiGeal.mnuParâmetrosDepartamentoSeção.Enabled = True
            VisibleParâmetros
        End If
        'Tipo de ICM
        If TBLGrupos("TIPO DE ICM") <> Empty Then
            mdiGeal.mnuParâmetros.Visible = True
            mdiGeal.mnuParâmetrosTipodeICM.Visible = True
            mdiGeal.mnuParâmetrosTipodeICM.Enabled = True
            mdiGeal.mnuSep11.Visible = True
            VisibleParâmetros
        End If
        'Tipo de Embalagem
        If TBLGrupos("TIPO DE EMBALAGEM") <> Empty Then
            mdiGeal.mnuParâmetros.Visible = True
            mdiGeal.mnuParâmetrosTipodeEmbalagem.Visible = True
            mdiGeal.mnuParâmetrosTipodeEmbalagem.Enabled = True
            mdiGeal.mnuSep11.Visible = True
            VisibleParâmetros
        End If
        'Unidades
        If TBLGrupos("UNIDADES") <> Empty Then
            mdiGeal.mnuParâmetros.Visible = True
            mdiGeal.mnuParâmetrosUnidades.Visible = True
            mdiGeal.mnuParâmetrosUnidades.Enabled = True
            mdiGeal.mnuSep15.Visible = True
            VisibleParâmetros
        End If
        'Localidade de Estoque
        If TBLGrupos("LOCALIDADE DE ESTOQUE") <> Empty Then
            mdiGeal.mnuParâmetros.Visible = True
            mdiGeal.mnuParâmetrosLocalidadeDeEstoque.Visible = True
            mdiGeal.mnuParâmetrosLocalidadeDeEstoque.Enabled = True
            mdiGeal.mnuSep12.Visible = True
            VisibleParâmetros
        End If
        'Plano de Pagamento
        If TBLGrupos("PLANO DE PAGAMENTO") <> Empty Then
            mdiGeal.mnuParâmetros.Visible = True
            mdiGeal.mnuParâmetrosPlanoDePagamento.Visible = True
            mdiGeal.mnuParâmetrosPlanoDePagamento.Enabled = True
            mdiGeal.mnuSep16.Visible = True
            VisibleParâmetros
        End If
        'Fim Parâmetros
        
        TBLUsuárioGrupo.MoveNext
        If TBLUsuárioGrupo.EOF Then
            Exit Do
        End If
    Loop
    
Fim:
    If GruposAberto Then
        TBLGrupos.Close
    End If
    If UsuárioGrupoAberto Then
        TBLUsuárioGrupo.Close
    End If
    
    Exit Sub
Erro:
    GeraMensagemDeErro "ChamaConfigurações - Usuário:" & Usuário
    If GruposAberto Then
        TBLGrupos.Close
    End If
    If UsuárioGrupoAberto Then
        TBLUsuárioGrupo.Close
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
    If DBDUsuárioAberto Then
        DBUsuário.Close
    End If
    If DBDFinanceiroAberto Then
        DBFinanceiro.Close
    End If
    If DBDUtilitárioAberto Then
        DBUtilitário.Close
    End If
    If DBDSistemaAberto Then
        DBSistema.Close
    End If
End Sub
Public Function FecharPorta() As Boolean
    PDV.FecharPorta
    glPortaAberta = False
End Function
Public Sub GeraMensagemDeErro(ByVal Operação, Optional ByVal Rollback As Boolean = False)
    Dim ErroAberto As Boolean, TBLErro As Table, IndiceErroAtivo$
    Dim ErroNumero&, ErroDescrição$, Data, Hora
    
    Data = Date
    Hora = Time
    
    ErroNumero = Err.Number
    ErroDescrição = Err.Description
    
    MensagemDeErro
    
    If Rollback Then
        WS.Rollback
    End If
    
    ErroAberto = AbreTabela(Dicionário, "SISTEMA", "ERRO", DBSistema, TBLErro, TBLTabela, dbOpenTable, True)
    
    If ErroAberto Then
        IndiceErroAtivo = "ERRO1"
        TBLErro.Index = IndiceErroAtivo
    Else
        MsgBox "Não consegui abrir a tabela 'Erro' !", vbCritical, "Erro"
    End If
    
    If ErroAberto Then
        On Error Resume Next
        TBLErro.AddNew
        TBLErro("CÓDIGO DO ERRO") = ErroNumero
        TBLErro("DESCRIÇÃO") = ErroDescrição
        TBLErro("USERNAME") = gUsuário
        TBLErro("OPERAÇÃO") = Operação
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
    
    FornecedorAberto = AbreTabela(Dicionário, "CADASTRO", "FORNECEDOR", DBCadastro, TBLFornecedor, TBLTabela, dbOpenTable)
    
    If FornecedorAberto Then
        IndiceFornecedorAtivo = "FORNECEDOR1"
        TBLFornecedor.Index = IndiceFornecedorAtivo
    Else
        MsgBox "Não consegui abrir a tabela 'Fornecedor' !", vbCritical, "Erro"
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
Public Function IsCorrectFuncionário(txtObj As Object) As Boolean
    Dim FuncionárioAberto As Boolean, TBLFuncionário As Table, IndiceFuncionárioAtivo$
    
    FuncionárioAberto = AbreTabela(Dicionário, "USUÁRIO", "FUNCIONÁRIO", DBUsuário, TBLFuncionário, TBLTabela, dbOpenTable)
    
    If FuncionárioAberto Then
        IndiceFuncionárioAtivo = "FUNCIONÁRIO1"
        TBLFuncionário.Index = IndiceFuncionárioAtivo
    Else
        MsgBox "Não consegui abrir a tabela 'Funcionário' !", vbCritical, "Erro"
        IsCorrectFuncionário = False
        Exit Function
    End If
    
    If txtObj <> Empty Then
        TBLFuncionário.Seek "=", txtObj
        
        If TBLFuncionário.NoMatch Then
            IsCorrectFuncionário = False
        Else
            IsCorrectFuncionário = True
        End If
    Else
        IsCorrectFuncionário = False
    End If
End Function
Public Function ImpressaoDeCheque(ByVal Banco As String, ByVal Valor As String, ByVal Data As String, ByVal DadosAdicionais As String) As Boolean
    If PDV.ImpressaoDeCheque(Chr(27) & Chr(46) & "24" & Banco & Valor & "N" & DadosAdicionais & "4" & Data & "}") Then
        ImpressaoDeCheque = True
    Else
        ImpressaoDeCheque = False
    End If
End Function
Public Function LeituraX(ByVal RelatórioGerencial As String) As Boolean
    If PDV.LeituraX(Chr(27) & Chr(46) & "13" & RelatórioGerencial & "}") Then
        LeituraX = True
    Else
        LeituraX = False
    End If
End Function
Public Sub Log(ByVal Usuário As String, ByVal Operação As String)
    Dim LogAberto As Boolean, TBLLog As Table, IndiceLogAtivo$
    Dim Data, Hora
    
    LogAberto = AbreTabela(Dicionário, "SISTEMA", "LOG", DBSistema, TBLLog, TBLTabela, dbOpenTable, True)
    
    If Not LogAberto Then
        MsgBox "Não consegui abrir a tabela 'Log' !", vbCritical, "Log"
        Exit Sub
    End If
    
    TBLLog.AddNew
    TBLLog("USERNAME") = Usuário
    TBLLog("OPERAÇÃO") = Operação
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
        MensagemDeErro vbCr & "Main - Criação do Objeto " & Objeto & vbCr & "Não será possível utilizar o PDV"
    End If
    On Error GoTo 0
    
    If AbreBaseDeDados(True, False) Then
        If ValidaUsuário(frmSplash.Top, frmSplash.Height) Then
            SetRegistryString "Geal", "Geral", "Usuário", gUsuário
            ChamaConfigurações gUsuário
            mdiGeal.Show
            Unload frmSplash
            Unload frmValidaUsuário
            Set frmSplash = Nothing
            Set frmValidaUsuário = Nothing
        Else
            Unload frmSplash
            Set frmSplash = Nothing
        End If
    Else
        Unload frmSplash
        Set frmSplash = Nothing
    End If
End Sub
Public Sub NavegaçãoInferior(ByVal Valor As Boolean)
    mdiGeal.mnuNavegaçãoPrimeiroRegistro.Enabled = Valor
    mdiGeal.mnuNavegaçãoRegistroAnterior.Enabled = Valor
    mdiGeal.Toolbar.Buttons("MoveFirst").Enabled = Valor
    mdiGeal.Toolbar.Buttons("MovePrevious").Enabled = Valor
End Sub
Public Sub NavegaçãoSuperior(ByVal Valor As Boolean)
    mdiGeal.mnuNavegaçãoÚltimoRegistro.Enabled = Valor
    mdiGeal.mnuNavegaçãoPróximoRegistro.Enabled = Valor
    mdiGeal.Toolbar.Buttons("MoveNext").Enabled = Valor
    mdiGeal.Toolbar.Buttons("MoveLast").Enabled = Valor
End Sub
Public Function ReduçãoZ(ByVal RelatórioGerencial As String, Optional ByVal Data As String) As Boolean
    If PDV.ReduçãoZ(Chr(27) & Chr(46) & "14" & RelatórioGerencial & "}") Then
        ReduçãoZ = True
    Else
        ReduçãoZ = False
    End If
End Function
Public Function RegistrarItemVendido(ByVal Código$, ByVal Quantidade$, ByVal PreçoUnitário$, ByVal PreçoTotal$, ByVal Descrição$, ByVal Tributação$) As Boolean
    If PDV.RegistrarItemVendido(Chr(27) & Chr(46) & "01" & Código & Quantidade & PreçoUnitário & PreçoTotal & Descrição & Tributação & "}") Then
        RegistrarItemVendido = True
    Else
        RegistrarItemVendido = False
    End If
End Function
Public Sub ResizeStatusBar()
    Dim Tamanho, Cont, Posição
    
    Tamanho = 0
    For Cont = 2 To mdiGeal.StatusBar.Panels.Count
        If mdiGeal.StatusBar.Panels(Cont).Visible = True Then
            Tamanho = Tamanho + mdiGeal.StatusBar.Panels(Cont).Width
        End If
    Next
    Posição = mdiGeal.Width - 500 - Tamanho
    If Posição >= 0 Then
        mdiGeal.StatusBar.Panels(1).Width = Posição
    End If
End Sub
Public Function SearchAdvancedProduto(ByVal Código As String, ByVal Tipo As Integer, Optional ByVal Indice As Integer) As Variant
    Dim TBLProduto As Table
    Dim ProdutoAberto As Boolean
    Dim IndiceProdutoAtivo$
    
    Dim TBLCódigoProduto As Table
    Dim CódigoProdutoAberto As Boolean
    Dim IndiceCódigoProdutoAtivo$
    
    Dim TBLPreçoProduto As Table
    Dim PreçoProdutoAberto As Boolean
    Dim IndicePreçoProdutoAtivo$
    
    Dim TBLTipoDeICM As Table
    Dim TipoDeICMAberto As Boolean
    Dim IndiceTipoDeICMAtivo$
    
    If IsMissing(Indice) Or Indice = 0 Then
        Indice = 3 'Código do Fornecedor como padrão a Tabela Código do Produto
    End If
    
    'Abre tabela PRODUTO
    ProdutoAberto = AbreTabela(Dicionário, "CADASTRO", "PRODUTO", DBCadastro, TBLProduto, TBLTabela, dbOpenTable)
    
    If ProdutoAberto Then
        IndiceProdutoAtivo = "PRODUTO1"
        TBLProduto.Index = IndiceProdutoAtivo
    Else
        MsgBox "Não consegui abrir a tabela 'Produto' !", vbCritical, "Erro"
        Exit Function
    End If
    
    'Abre tabela CÓDIGO DO PRODUTO
    CódigoProdutoAberto = AbreTabela(Dicionário, "CADASTRO", "CÓDIGO DO PRODUTO", DBCadastro, TBLCódigoProduto, TBLTabela, dbOpenTable)
    'Se Indice 1 - +CÓDIGO DO PRODUTO;+FORNECEDOR;+CÓDIGO DO FORNECEDOR
    'Se Indice 2 - +CÓDIGO DO PRODUTO
    'Se Indice 3 - +CÓDIGO DO FORNECEDOR
    'Se Indice 4 - +CÓDIGO DO PRODUTO;+FORNECEDOR
    'Se Indice 5 - +FORNECEDOR;+CÓDIGO DO FORNECEDOR
    If CódigoProdutoAberto Then
        IndiceCódigoProdutoAtivo = "CÓDIGODOPRODUTO" & Indice
        TBLCódigoProduto.Index = IndiceCódigoProdutoAtivo
    Else
        MsgBox "Não consegui abrir a tabela 'Código do Produto' !", vbCritical, "Erro"
        Exit Function
    End If

    'Abre tabela PREÇO DO PRODUTO
    PreçoProdutoAberto = AbreTabela(Dicionário, "CADASTRO", "PREÇO DO PRODUTO", DBCadastro, TBLPreçoProduto, TBLTabela, dbOpenTable)
    
    If PreçoProdutoAberto Then
        IndicePreçoProdutoAtivo = "PREÇODOPRODUTO1"
        TBLPreçoProduto.Index = IndicePreçoProdutoAtivo
    Else
        MsgBox "Não consegui abrir a tabela 'PreçoProduto' !", vbCritical, "Erro"
        Exit Function
    End If
     
    TipoDeICMAberto = AbreTabela(Dicionário, "CADASTRO", "TIPO DE ICM", DBCadastro, TBLTipoDeICM, TBLTabela, dbOpenTable)
    
    If TipoDeICMAberto Then
        IndiceTipoDeICMAtivo = "TIPODEICM1"
        TBLTipoDeICM.Index = IndiceTipoDeICMAtivo
    Else
        MsgBox "Não consegui abrir a tabela 'Tipo De Embalagem' !", vbCritical, "Erro"
        Exit Function
    End If
    
    'Tipo do Retorno
    If Tipo = vbDescrição Then
        TBLProduto.Seek "=", Código
        If TBLProduto.NoMatch Then
            SearchAdvancedProduto = ""
        Else
            SearchAdvancedProduto = TBLProduto("DESCRIÇÃO")
        End If
    ElseIf Tipo = vbCódigo Then
        TBLCódigoProduto.Seek "=", Código
        If TBLCódigoProduto.NoMatch Then
            SearchAdvancedProduto = ""
        Else
            SearchAdvancedProduto = TBLCódigoProduto("CÓDIGO DO PRODUTO")
        End If
    ElseIf Tipo = vbCódigoDoFornecedor Then
        TBLCódigoProduto.Seek "=", Código
        If TBLCódigoProduto.NoMatch Then
            SearchAdvancedProduto = ""
        Else
            SearchAdvancedProduto = TBLCódigoProduto("CÓDIGO DO FORNECEDOR")
        End If
    ElseIf Tipo = vbValorUnitário Then
        TBLCódigoProduto.Seek "=", Código
        If TBLCódigoProduto.NoMatch Then
            SearchAdvancedProduto = Empty
            Exit Function
        End If
        TBLPreçoProduto.Seek "=", Código, TBLCódigoProduto("CÓDIGO DO PRODUTO")
        If TBLPreçoProduto.NoMatch Then
            TBLPreçoProduto.Index = "PREÇODOPRODUTO2"
            TBLPreçoProduto.Seek "=", Código
            If TBLPreçoProduto.NoMatch Then
                SearchAdvancedProduto = "0,00"
            Else
                SearchAdvancedProduto = FormatStringMask("@V ##.###.##0,00", ValStr(TBLPreçoProduto("PREÇO DE VENDA")))
            End If
        Else
            SearchAdvancedProduto = FormatStringMask("@V ##.###.##0,00", ValStr(TBLPreçoProduto("PREÇO DE VENDA")))
        End If
    ElseIf Tipo = vbValValorUnitário Then
        TBLCódigoProduto.Seek "=", Código
        If TBLCódigoProduto.NoMatch Then
            SearchAdvancedProduto = Empty
            Exit Function
        End If
        TBLPreçoProduto.Seek "=", Código, TBLCódigoProduto("CÓDIGO DO PRODUTO")
        If TBLPreçoProduto.NoMatch Then
            TBLPreçoProduto.Index = "PREÇODOPRODUTO2"
            TBLPreçoProduto.Seek "=", TBLCódigoProduto("CÓDIGO DO PRODUTO")
            If TBLPreçoProduto.NoMatch Then
                SearchAdvancedProduto = 0
            Else
                SearchAdvancedProduto = TBLPreçoProduto("PREÇO DE VENDA")
            End If
        Else
            SearchAdvancedProduto = TBLPreçoProduto("PREÇO DE VENDA")
        End If
    ElseIf Tipo = vbLote Then
        TBLProduto.Seek "=", Código
        If TBLProduto.NoMatch Then
            SearchAdvancedProduto = False
        Else
            SearchAdvancedProduto = TBLProduto("LOTES")
        End If
    ElseIf Tipo = vbDescontoMáximo Then
        TBLProduto.Seek "=", Código
        If TBLProduto.NoMatch Then
            SearchAdvancedProduto = False
        Else
            SearchAdvancedProduto = TBLProduto("DESCONTO MÁXIMO")
        End If
    ElseIf Tipo = vbTributo Then
        TBLProduto.Seek "=", Código
        If TBLProduto.NoMatch Then
            SearchAdvancedProduto = Empty
        Else
            TBLTipoDeICM.Seek "=", TBLProduto("TIPO DE ICM")
            If TBLTipoDeICM.NoMatch Then
                SearchAdvancedProduto = Empty
            Else
                SearchAdvancedProduto = TBLTipoDeICM("CÓDIGO DO PDV")
            End If
        End If
    End If
End Function
Public Function SearchCliente(ByVal Busca As String, ByVal CampoDeBusca As Byte) As String
    Dim TBLCliente As Table
    Dim ClienteAberto As Boolean

    ClienteAberto = AbreTabela(Dicionário, "CADASTRO", "CLIENTE", DBCadastro, TBLCliente, TBLTabela, dbOpenTable)
    
    If ClienteAberto Then
        If CampoDeBusca = byCodigo Then
            TBLCliente.Index = "CLIENTE1"
        ElseIf CampoDeBusca = byNome Then
            TBLCliente.Index = "CLIENTE2"
        ElseIf CampoDeBusca = byCGCCPF Then
            TBLCliente.Index = "CLIENTE3"
        End If
    Else
        MsgBox "Não consegui abrir a tabela 'Cliente' !", vbCritical, "Erro"
        Exit Function
    End If
    
    TBLCliente.Seek "=", Busca
    
    If TBLCliente.NoMatch Then
        SearchCliente = Empty
    Else
        SearchCliente = TBLCliente("NOME - RAZÃO SOCIAL")
    End If
    
    TBLCliente.Close
End Function
Public Function SearchFornecedor(ByVal mCGCCPF As String) As String
    Dim TBLFornecedor As Table
    Dim FornecedorAberto As Boolean
    Dim IndiceFornecedorAtivo$
    
    FornecedorAberto = AbreTabela(Dicionário, "CADASTRO", "FORNECEDOR", DBCadastro, TBLFornecedor, TBLTabela, dbOpenTable)
    
    If FornecedorAberto Then
        IndiceFornecedorAtivo = "FORNECEDOR1"
        TBLFornecedor.Index = IndiceFornecedorAtivo
    Else
        MsgBox "Não consegui abrir a tabela 'Fornecedor' !", vbCritical, "Erro"
        Exit Function
    End If
    
    TBLFornecedor.Seek "=", mCGCCPF
    
    If TBLFornecedor.NoMatch Then
        MsgBox "Fornecedor " & mCGCCPF & " não foi encontrado!", vbCritical, "Erro"
        Exit Function
    Else
        SearchFornecedor = TBLFornecedor("RAZÃO SOCIAL")
    End If
    TBLFornecedor.Close
End Function
Public Function SearchProduto(ByVal Código)
    Dim TBLProduto As Table
    Dim ProdutoAberto As Boolean
    Dim IndiceProdutoAtivo$
    
    ProdutoAberto = AbreTabela(Dicionário, "CADASTRO", "PRODUTO", DBCadastro, TBLProduto, TBLTabela, dbOpenTable)
    
    If ProdutoAberto Then
        IndiceProdutoAtivo = "PRODUTO1"
        TBLProduto.Index = IndiceProdutoAtivo
    Else
        MsgBox "Não consegui abrir a tabela 'Produto' !", vbCritical, "Erro"
        SearchProduto = ""
        Exit Function
    End If
    
    TBLProduto.Seek "=", Código
    
    If TBLProduto.NoMatch Then
        MsgBox "Não foi possível encontrar o produto " & Código, vbCritical, "Erro"
        SearchProduto = ""
        TBLProduto.Close
        Exit Function
    Else
        SearchProduto = TBLProduto("DESCRIÇÃO")
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
    
    mdiGeal.StatusBar.Panels("Posição").Visible = True
    mdiGeal.StatusBar.Panels("Posição").Text = "Registros: " & TBLWork.RecordCount
    ResizeStatusBar

    If TBLWork.RecordCount = 0 Then
        NavegaçãoInferior False
        BotãoExcluir False
        BotãoGravar False
        Exit Sub
    End If
    BotãoExcluir lAllowDelete
    BotãoGravar True
    TBLWork.MovePrevious
    If TBLWork.BOF Then
        NavegaçãoInferior False
    Else
        NavegaçãoInferior lAllowConsult
    End If
    TBLWork.MoveNext
End Sub
Public Sub TestaSuperior(ByRef TBLWork As Table, ByVal lAllowEdit As Boolean, ByVal lAllowDelete As Boolean, ByVal lAllowConsult As Boolean)
    On Error Resume Next
    
    mdiGeal.StatusBar.Panels("Posição").Visible = True
    mdiGeal.StatusBar.Panels("Posição").Text = "Registros: " & TBLWork.RecordCount
    ResizeStatusBar
    
    If TBLWork.RecordCount = 0 Then
        NavegaçãoSuperior False
        BotãoExcluir False
        BotãoGravar False
        Exit Sub
    End If
    BotãoExcluir lAllowDelete
    BotãoGravar True
    TBLWork.MoveNext
    If TBLWork.EOF Then
        NavegaçãoSuperior False
    Else
        NavegaçãoSuperior lAllowConsult
    End If
    TBLWork.MovePrevious
End Sub
Public Sub TestaInferiorArray(ByVal Elemento, ByRef Matriz(), ByVal lAllowEdit As Boolean, ByVal lAllowDelete As Boolean, ByVal lAllowConsult As Boolean, Optional ByVal Dimens As Byte = 1)
    mdiGeal.StatusBar.Panels("Posição").Text = "Registros: " & Elemento
    ResizeStatusBar
    
    If Elemento = 0 Then
        NavegaçãoInferior False
        BotãoExcluir False
        BotãoGravar False
        Exit Sub
    End If
    
    If UBound(Matriz, Dimens) = 0 Then
        NavegaçãoInferior False
        BotãoExcluir False
        BotãoGravar False
        Exit Sub
    End If
    BotãoExcluir lAllowDelete
    BotãoGravar True
    If Elemento = 1 Then
        NavegaçãoInferior False
    Else
        NavegaçãoInferior lAllowConsult
    End If
End Sub
Public Sub TestaSuperiorArray(ByVal Elemento, ByRef Matriz(), ByVal lAllowEdit As Boolean, ByVal lAllowDelete As Boolean, ByVal lAllowConsult As Boolean, Optional ByVal Dimens As Byte = 1)
    mdiGeal.StatusBar.Panels("Posição").Text = "Registros: " & Elemento
    ResizeStatusBar
    
    If Elemento = 0 Then
        NavegaçãoInferior False
        BotãoExcluir False
        BotãoGravar False
        Exit Sub
    End If
    If UBound(Matriz, Dimens) = 0 Then
        NavegaçãoSuperior False
        BotãoExcluir False
        BotãoGravar False
        Exit Sub
    End If
    BotãoExcluir lAllowDelete
    BotãoGravar True
    If Elemento = UBound(Matriz, Dimens) Then
        NavegaçãoSuperior False
    Else
        NavegaçãoSuperior lAllowConsult
    End If
End Sub
Public Function TotalizarCupomFiscal(ByVal Total$) As Boolean
    If PDV.TotalizarCupomFiscal(Chr(27) & Chr(46) & "1001" & Total & "}") Then
        TotalizarCupomFiscal = True
    Else
        TotalizarCupomFiscal = False
    End If
End Function
Private Function ValidaUsuário(ByVal WindowsTop As Long, WindowsHeight As Long) As Boolean
    'Valida Usuário
    frmValidaUsuário.GravaUsuário = True
    frmValidaUsuário.WindowTop = WindowsTop
    frmValidaUsuário.WindowHeight = WindowsHeight
    frmValidaUsuário.Show vbModeless
        
    Do While Not frmValidaUsuário.Fechado
        DoEvents
    Loop
    
    gUsuário = Trim(frmValidaUsuário.Usuário)
    
    Set frmValidaUsuário = Nothing
    
    If gUsuário = "" Then
        ValidaUsuário = False
    Else
        ValidaUsuário = True
    End If
End Function
Public Function VerStatusECF() As String
    Dim Status As String
    
    Status = Space(255)
    
    PDV.VerStatusECF Status, 255
    
    Status = StripTerminator(Status)
    
    VerStatusECF = Status
End Function
Private Sub VisãoTotal()
    Dim lAcesso As Boolean
    lAcesso = True
    With mdiGeal
        .mnuArquivoAbrirAgencia.Visible = lAcesso
        .mnuArquivoAbrirBanco.Visible = lAcesso
        .mnuArquivoAbrirCliente.Visible = lAcesso
        .mnuArquivoAbrirContaCorrente.Visible = lAcesso
        .mnuArquivoAbrirFornecedor.Visible = lAcesso
        .mnuArquivoAbrirFuncionário.Visible = lAcesso
        .mnuArquivoAbrirProduto.Visible = lAcesso
        .mnuArquivoAbrirDespesas.Visible = lAcesso
        .mnuArquivoAbrir.Visible = lAcesso
        
        .mnuMovimentoEntradaCompra.Visible = lAcesso
        .mnuMovimentoEntradaDevoluçãoTroca.Visible = lAcesso
        .mnuMovimentoEntrada.Visible = lAcesso
        
        .mnuMovimentoSaídaVenda.Visible = lAcesso
        .mnuMovimentoSaídaDevoluçãoTroca.Visible = lAcesso
        .mnuMovimentoSaída.Visible = lAcesso
        
        .mnuMovimentoMovimentoDiário.Visible = lAcesso
        .mnuMovimentoContaCorrente.Visible = lAcesso
        .mnuMovimentoCaixa.Visible = lAcesso
        .mnuMovimentoCaixaFácil.Visible = lAcesso
        .mnuMovimentoDespesas.Visible = lAcesso
        .mnuMovimento.Visible = lAcesso
        
        .mnuParâmetrosDepartamento.Visible = lAcesso
        .mnuParâmetrosSeção.Visible = lAcesso
        .mnuParâmetrosDepartamentoSeção.Visible = lAcesso
        .mnuParâmetrosTipodeICM.Visible = lAcesso
        .mnuParâmetrosTipodeEmbalagem.Visible = lAcesso
        .mnuParâmetrosUnidades.Visible = lAcesso
        .mnuParâmetrosLocalidadeDeEstoque.Visible = lAcesso
        .mnuParâmetrosPlanoDePagamento.Visible = lAcesso
        .mnuParâmetros.Visible = lAcesso
        
        .mnuParâmetrosUsuários.Visible = lAcesso
        .mnuParâmetrosGrupos.Visible = lAcesso
        
        .mnuParâmetroSep2.Visible = lAcesso
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
    If Not mdiGeal.mnuArquivoAbrirFuncionário.Enabled Then
        mdiGeal.mnuArquivoAbrirFuncionário.Visible = False
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
        If Not mdiGeal.mnuMovimentoEntradaDevoluçãoTroca.Enabled Then
            mdiGeal.mnuMovimentoEntradaDevoluçãoTroca.Visible = False
        End If
    End If
    If mdiGeal.mnuMovimentoSaída.Visible Then
        If Not mdiGeal.mnuMovimentoSaídaVenda.Enabled Then
            mdiGeal.mnuMovimentoSaídaVenda.Visible = False
        End If
        If Not mdiGeal.mnuMovimentoSaídaDevoluçãoTroca.Enabled Then
            mdiGeal.mnuMovimentoSaídaDevoluçãoTroca.Visible = False
        End If
    End If
    If Not mdiGeal.mnuMovimentoMovimentoDiário.Enabled Then
        mdiGeal.mnuMovimentoMovimentoDiário.Visible = False
    End If
    If Not mdiGeal.mnuMovimentoContaCorrente.Enabled Then
        mdiGeal.mnuMovimentoContaCorrente.Visible = False
    End If
    If Not mdiGeal.mnuMovimentoCaixa.Enabled Then
        mdiGeal.mnuMovimentoCaixa.Visible = False
    End If
    If Not mdiGeal.mnuMovimentoCaixaFácil.Enabled Then
        mdiGeal.mnuMovimentoCaixaFácil.Visible = False
    End If
    If Not mdiGeal.mnuMovimentoDespesas.Enabled Then
        mdiGeal.mnuMovimentoDespesas.Visible = False
    End If
End Sub
Private Sub VisibleParâmetros()
    If Not mdiGeal.mnuParâmetrosDepartamento.Enabled Then
        mdiGeal.mnuParâmetrosDepartamento.Visible = False
    End If
    If Not mdiGeal.mnuParâmetrosSeção.Enabled Then
        mdiGeal.mnuParâmetrosSeção.Visible = False
    End If
    If Not mdiGeal.mnuParâmetrosDepartamentoSeção.Enabled Then
        mdiGeal.mnuParâmetrosDepartamentoSeção.Visible = False
    End If
    If Not mdiGeal.mnuParâmetrosTipodeICM.Enabled Then
        mdiGeal.mnuParâmetrosTipodeICM.Visible = False
    End If
    If Not mdiGeal.mnuParâmetrosTipodeEmbalagem.Enabled Then
        mdiGeal.mnuParâmetrosTipodeEmbalagem.Visible = False
    End If
    If Not mdiGeal.mnuParâmetrosUnidades.Enabled Then
        mdiGeal.mnuParâmetrosUnidades.Visible = False
    End If
    If Not mdiGeal.mnuParâmetrosLocalidadeDeEstoque.Enabled Then
        mdiGeal.mnuParâmetrosLocalidadeDeEstoque.Visible = False
    End If
    If Not mdiGeal.mnuParâmetrosPlanoDePagamento.Enabled Then
        mdiGeal.mnuParâmetrosPlanoDePagamento.Visible = False
    End If
End Sub
