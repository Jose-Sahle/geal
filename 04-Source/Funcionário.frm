VERSION 5.00
Begin VB.Form frmFuncionário 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Funcioário"
   ClientHeight    =   5100
   ClientLeft      =   1575
   ClientTop       =   1515
   ClientWidth     =   6540
   Icon            =   "Funcionário.frx":0000
   LinkTopic       =   "frmFuncionário"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5100
   ScaleWidth      =   6540
   Begin VB.Frame frDadosContratuais 
      Caption         =   "Dados Contratuais "
      Height          =   1155
      Left            =   0
      TabIndex        =   20
      Top             =   3540
      Width           =   6525
      Begin VB.TextBox txtSalário 
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
         Left            =   4770
         TabIndex        =   9
         Text            =   " "
         Top             =   690
         Width           =   1665
      End
      Begin VB.TextBox txtDatadeSaída 
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
         Left            =   1410
         TabIndex        =   8
         Top             =   690
         Width           =   1305
      End
      Begin VB.TextBox txtDatadeEntrada 
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
         Left            =   1410
         TabIndex        =   7
         Top             =   270
         Width           =   1305
      End
      Begin VB.Label lblSalário 
         Caption         =   "Salário"
         Height          =   195
         Left            =   4140
         TabIndex        =   23
         Top             =   750
         Width           =   555
      End
      Begin VB.Label lblDatadeSaída 
         Caption         =   "Data de Saída"
         Height          =   225
         Left            =   150
         TabIndex        =   22
         Top             =   750
         Width           =   1185
      End
      Begin VB.Label lblDataDeEntrada 
         Caption         =   "Data de Entrada"
         Height          =   225
         Left            =   150
         TabIndex        =   21
         Top             =   300
         Width           =   1245
      End
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   345
      Left            =   5280
      TabIndex        =   11
      Top             =   4740
      Width           =   1245
   End
   Begin VB.CommandButton cmdGravar 
      Caption         =   "&Gravar"
      Height          =   345
      Left            =   3960
      TabIndex        =   10
      Top             =   4740
      Width           =   1245
   End
   Begin VB.Frame frDadosCadastrais 
      Caption         =   " Dados Cadastrais "
      Height          =   3525
      Left            =   0
      TabIndex        =   17
      Top             =   0
      Width           =   6525
      Begin VB.TextBox txtCpf 
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
         Left            =   1200
         TabIndex        =   6
         Text            =   "   .   .   -  "
         Top             =   3030
         Width           =   2310
      End
      Begin VB.TextBox txtCep 
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
         Left            =   5130
         TabIndex        =   5
         Text            =   "  .   -   "
         Top             =   2580
         Width           =   1305
      End
      Begin VB.TextBox txtUF 
         Height          =   285
         Left            =   1200
         TabIndex        =   4
         Top             =   2610
         Width           =   435
      End
      Begin VB.TextBox txtCidade 
         Height          =   285
         Left            =   1200
         TabIndex        =   3
         Top             =   2130
         Width           =   5235
      End
      Begin VB.TextBox txtBairro 
         Height          =   285
         Left            =   1200
         TabIndex        =   2
         Top             =   1680
         Width           =   5235
      End
      Begin VB.TextBox txtEndereço 
         Height          =   285
         Left            =   1200
         TabIndex        =   1
         Top             =   1230
         Width           =   5235
      End
      Begin VB.TextBox txtNome 
         Height          =   285
         Left            =   1200
         TabIndex        =   0
         Top             =   780
         Width           =   5235
      End
      Begin VB.TextBox txtCódigo 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1200
         TabIndex        =   24
         Top             =   330
         Width           =   525
      End
      Begin VB.Label lblCgcCpf 
         Caption         =   "C. P. F."
         Height          =   195
         Left            =   150
         TabIndex        =   12
         Top             =   3090
         Width           =   645
      End
      Begin VB.Label lblCep 
         Caption         =   "CEP"
         Height          =   195
         Left            =   4680
         TabIndex        =   13
         Top             =   2640
         Width           =   315
      End
      Begin VB.Label lblUF 
         Caption         =   "U. F."
         Height          =   225
         Left            =   150
         TabIndex        =   14
         Top             =   2640
         Width           =   405
      End
      Begin VB.Label lblCidade 
         Caption         =   "Cidade"
         Height          =   225
         Left            =   150
         TabIndex        =   15
         Top             =   2160
         Width           =   945
      End
      Begin VB.Label lblBairro 
         Caption         =   "Bairro"
         Height          =   225
         Left            =   150
         TabIndex        =   16
         Top             =   1710
         Width           =   945
      End
      Begin VB.Label lblEndereço 
         Caption         =   "Endereço"
         Height          =   195
         Left            =   150
         TabIndex        =   19
         Top             =   1260
         Width           =   975
      End
      Begin VB.Label lblNomeRazãoSocial 
         Caption         =   "Nome"
         Height          =   195
         Left            =   150
         TabIndex        =   18
         Top             =   810
         Width           =   1065
      End
      Begin VB.Label lblCódigo 
         Caption         =   "Código"
         Height          =   195
         Left            =   150
         TabIndex        =   25
         Top             =   360
         Width           =   1065
      End
   End
End
Attribute VB_Name = "frmFuncionário"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim TBLFuncionário As Table
Dim FuncionárioAberto As Boolean
Dim IndiceFuncionárioAtivo$

Dim TBLParâmetros As Table
Dim ParâmetrosAberto As Boolean

Dim lAllowInsert  As Boolean
Dim lAllowEdit    As Boolean
Dim lAllowDelete  As Boolean
Dim lAllowConsult As Boolean

Dim lInserir As Boolean
Dim lAlterar As Boolean

Dim mFechar As Boolean
Dim lPula As Boolean
Dim lPush As Boolean

Dim lInicio As Boolean
Dim StatusBar$

Public StatusBarAviso$

Dim DataBaseName(1 To 1) As String
Public Relatório$
Public TotalDatabaseName%

Public lAtualizar As Boolean
Private Sub AtivaCampos()
    BotãoImprimir True
    frDadosCadastrais.Enabled = True
    frDadosContratuais.Enabled = True
    BotãoGravar (lInserir Or lAllowEdit)
    cmdCancelar.Enabled = (lInserir Or lAllowEdit)
    cmdGravar.Enabled = (lInserir Or lAllowEdit)
End Sub
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
    
    lInserir = False
    lAlterar = False
    BotãoIncluir lAllowInsert
    
    If TBLFuncionário.RecordCount = 0 Then
        NavegaçãoInferior False
        NavegaçãoSuperior False
        BotãoGravar False
        cmdGravar.Enabled = False
        cmdCancelar.Enabled = False
        DesativaCampos
        lPush = True
        ZeraCampos
        lPush = False
        Cancelamento = True
        Exit Function
    End If
    
    Cancelamento = True
    
    TestaInferior TBLFuncionário, lAllowEdit, lAllowDelete, lAllowConsult
    TestaSuperior TBLFuncionário, lAllowEdit, lAllowDelete, lAllowConsult
    
    GetRecords
End Function
Public Function PushDataBaseName(ByVal Posição As Integer) As String
    PushDataBaseName = DataBaseName(Posição)
End Function
Private Sub DesativaCampos()
    BotãoImprimir False
    frDadosCadastrais.Enabled = False
    frDadosContratuais.Enabled = False
    cmdGravar.Enabled = False
    cmdCancelar.Enabled = False
    BotãoGravar False
End Sub
Public Sub Encontrar()
    If Not lAllowConsult Then
        Exit Sub
    End If
    Set frmEncontrar.DBBancoDeDados = DBUsuário
    frmEncontrar.NomeDaJanela = "Funcionário"
    frmEncontrar.LabelDescription = "Nome"
    frmEncontrar.Mensagem = "Nenhuma funcionário foi selecionado!"
    frmEncontrar.BancoDeDados = "USUÁRIO"
    frmEncontrar.Tabela = "FUNCIONÁRIO"
    frmEncontrar.Indice = "1"
    frmEncontrar.CampoChave = "CÓDIGO"
    frmEncontrar.CampoPreencheLista = "NOME"
    frmEncontrar.Show vbModal
    lPula = True
    txtCódigo = frmEncontrar.Chave
    lPula = False
    PosRecords
End Sub
Public Sub Excluir()
    Dim Confirmação As Integer, Msg1$, Msg2$
  
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
    
    TBLFuncionário.Delete
    
    If Err <> 0 Then
        GeraMensagemDeErro "Funcionário - Excluir - " & txtNome, True
        StatusBarAviso = "Falha na exclusão"
        BarraDeStatus StatusBarAviso
        Exit Sub
    End If
    
    WS.CommitTrans
    
    Log gUsuário, "Exclusão - Funcionário: " & txtCódigo & " - " & txtNome
    
    StatusBarAviso = "Exclusão bem sucedida"
    BarraDeStatus StatusBarAviso
    
    If TBLFuncionário.RecordCount = 0 Then
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
    
    If TBLFuncionário.BOF Then
        TBLFuncionário.MoveFirst
    ElseIf TBLFuncionário.EOF Then
        TBLFuncionário.MoveLast
    Else
        TBLFuncionário.MovePrevious
        If TBLFuncionário.BOF Then
            TBLFuncionário.MoveNext
        End If
    End If
    
    GetRecords
    
    TestaInferior TBLFuncionário, lAllowEdit, lAllowDelete, lAllowConsult
    TestaSuperior TBLFuncionário, lAllowEdit, lAllowDelete, lAllowConsult
End Sub
Public Sub Gravar()
    If lInserir Then
        'Pega o novo código interno de funcionário e atualiza na Tabela Parâmetros
        txtCódigo = TBLParâmetros("FUNCIONÁRIO") + 1
        TBLParâmetros.Edit
        TBLParâmetros("FUNCIONÁRIO") = txtCódigo
        TBLParâmetros.Update
        
        If SetRecords Then
            PosRecords
            lInserir = False
            StatusBarAviso = "Inclusão bem sucedida"
        Else
            StatusBarAviso = "Falha na inclusão"
            Exit Sub
        End If
    Else
        If TBLFuncionário.RecordCount > 0 And Not TBLFuncionário.BOF And Not TBLFuncionário.EOF Then
            If SetRecords Then
                PosRecords
                lAlterar = False
                StatusBarAviso = "Alteração bem sucedida"
            Else
                StatusBarAviso = "Falha na alteração"
            End If
        End If
    End If
    
    BarraDeStatus StatusBarAviso
    
    TestaInferior TBLFuncionário, lAllowEdit, lAllowDelete, lAllowConsult
    TestaSuperior TBLFuncionário, lAllowEdit, lAllowDelete, lAllowConsult
    
    If TBLFuncionário.RecordCount = 0 Then
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
    
    If txtNome.Enabled Then
        txtNome.SetFocus
    End If
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
    
    txtNome.SetFocus

End Sub
Public Sub MoveFirst()
    If lAlterar Then
       If Not Cancelamento Then
            Exit Sub
        End If
    End If
    
    StatusBarAviso = "Pronto"
    BarraDeStatus StatusBarAviso
    
    TBLFuncionário.MoveFirst
    
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
    
    TBLFuncionário.MoveLast
    
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
    
    TBLFuncionário.MoveNext
    If TBLFuncionário.EOF Then
        TBLFuncionário.MovePrevious
        Exit Sub
    End If
    
    StatusBarAviso = "Pronto"
    BarraDeStatus StatusBarAviso
    
    NavegaçãoInferior lAllowConsult
    TestaSuperior TBLFuncionário, lAllowEdit, lAllowDelete, lAllowConsult
    
    GetRecords
End Sub
Public Sub MovePrevious()
    If lAlterar Then
       If Not Cancelamento Then
            Exit Sub
        End If
    End If
    
    TBLFuncionário.MovePrevious
    If TBLFuncionário.BOF Then
        TBLFuncionário.MoveNext
        Exit Sub
    End If
    
    StatusBarAviso = "Pronto"
    BarraDeStatus StatusBarAviso
    
    NavegaçãoSuperior lAllowConsult
    TestaInferior TBLFuncionário, lAllowEdit, lAllowDelete, lAllowConsult
    
    GetRecords
End Sub
Public Sub PosRecords()
    If TBLFuncionário.RecordCount = 0 Then
        Exit Sub
    End If
    
    TBLFuncionário.Seek "=", txtCódigo
    If TBLFuncionário.NoMatch Then
        MsgBox "Não consegui encontrar o funcionário com o código  " + txtCódigo, vbExclamation, "Erro"
        TBLFuncionário.MoveFirst
        NavegaçãoInferior False
        NavegaçãoInferior lAllowConsult
    Else
        TestaInferior TBLFuncionário, lAllowEdit, lAllowDelete, lAllowConsult
        TestaSuperior TBLFuncionário, lAllowEdit, lAllowDelete, lAllowConsult
    End If
    GetRecords
End Sub
Private Sub GetRecords()
    On Error GoTo Erro
    
    lPush = True
    lPula = True
    If Not lAllowConsult Then
        ZeraCampos
        DesativaCampos
        lPush = False
        lPula = False
        Exit Sub
    End If
    txtCódigo = TBLFuncionário("CÓDIGO")
    txtNome = TBLFuncionário("NOME")
    txtEndereço = TBLFuncionário("ENDEREÇO")
    txtBairro = TBLFuncionário("BAIRRO")
    txtCidade = TBLFuncionário("CIDADE")
    txtUF = TBLFuncionário("UF")
    txtCep = TBLFuncionário("CEP")
    txtCpf = TBLFuncionário("CPF")
    
    If TBLFuncionário("DATA DE ENTRADA") <> vbNull Then
        txtDatadeEntrada = FormatStringMask(CheckDataMask, TBLFuncionário("DATA DE ENTRADA"))
        CorrigeData DataMask, txtDatadeEntrada, TBLFuncionário("DATA DE ENTRADA")
    Else
        txtDatadeEntrada = DataNula
    End If
    
    If TBLFuncionário("DATA DE SAÍDA") <> vbNull Then
        txtDatadeSaída = FormatStringMask(CheckDataMask, TBLFuncionário("DATA DE SAÍDA"))
        CorrigeData DataMask, txtDatadeSaída, TBLFuncionário("DATA DE SAÍDA")
    Else
        txtDatadeSaída = DataNula
    End If
    
    txtSalário = TBLFuncionário("SALÁRIO")
    txtSalário_LostFocus
    lPush = False
    lPula = False
    If Not lAllowEdit Then
        DesativaCampos
    End If
    If Not lAllowEdit Then
        DesativaCampos
    End If
    
    Exit Sub
    
Erro:
    GeraMensagemDeErro "Funcionário - GetRecords "
    Resume Next
End Sub
Private Function SetRecords()
    On Error GoTo Erro
    
    Dim Msg$
    Dim Confirmação As Integer, Msg1$, Msg2$
    
    WS.BeginTrans 'Inicia uma Transação
    
    If lInserir Then
        TBLFuncionário.AddNew
    Else
        TBLFuncionário.Edit
    End If
    
    If lInserir Then
        TBLFuncionário("CÓDIGO") = txtCódigo
    End If
    
    TBLFuncionário("NOME") = txtNome
    TBLFuncionário("ENDEREÇO") = txtEndereço
    TBLFuncionário("BAIRRO") = txtBairro
    TBLFuncionário("CIDADE") = txtCidade
    TBLFuncionário("UF") = txtUF
    TBLFuncionário("CEP") = txtCep
    TBLFuncionário("CPF") = txtCpf
    TBLFuncionário("DATA DE ENTRADA") = IIf(Trim(StrTran(txtDatadeEntrada, "/")) <> Empty, txtDatadeEntrada, vbNull)
    TBLFuncionário("DATA DE SAÍDA") = IIf(Trim(StrTran(txtDatadeSaída, "/")) <> Empty, txtDatadeSaída, vbNull)
    TBLFuncionário("SALÁRIO") = txtSalário
    
    If lInserir Then
        TBLFuncionário("USERNAME - CRIA") = gUsuário
        TBLFuncionário("DATA - CRIA") = Date
        TBLFuncionário("HORA - CRIA") = Time
        TBLFuncionário("USERNAME - ALTERA") = "VAZIO"
        TBLFuncionário("DATA - ALTERA") = vbNull
        TBLFuncionário("HORA - ALTERA") = vbNull
    End If
    If lAlterar Then
        TBLFuncionário("USERNAME - ALTERA") = gUsuário
        TBLFuncionário("DATA - ALTERA") = Date
        TBLFuncionário("HORA - ALTERA") = Time
    End If
    TBLFuncionário.Update
    
Erro:
    If Err <> 0 Then
        TBLFuncionário.CancelUpdate
        GeraMensagemDeErro "Funcionário - SetRecords - " & txtNome, True
        SetRecords = False
        Exit Function
    End If

    WS.CommitTrans 'Grava as alterações ou inclusões se não houverem erros
    
    If lInserir Then
        Log gUsuário, "Inclusão - Funcionário " & txtCódigo & " - " & txtNome
    Else
        Log gUsuário, "Alteração - Funcionário " & txtCódigo & " - " & txtNome
    End If
    
    SetRecords = True
End Function
Private Sub ZeraCampos()
    lPula = True
    txtCódigo = Empty
    txtNome = Empty
    txtEndereço = Empty
    txtBairro = Empty
    txtCidade = Empty
    txtUF = Empty
    txtCep = Empty
    txtCpf = Empty
    txtDatadeEntrada = DataNula
    txtDatadeSaída = DataNula
    txtSalário = FormatStringMask("@V ##.###.##0,00", "0,00")
    lPula = False
End Sub
Private Sub cmdCancelar_Click()
    Cancelamento
End Sub
Private Sub cmdGravar_Click()
    Gravar
End Sub
Private Sub Form_Activate()
    If mFechar Then
        Unload Me
        Exit Sub
    End If
    If Not FuncionárioAberto Then
        Unload Me
        Exit Sub
    End If
    
    If Not ParâmetrosAberto Then
        Unload Me
        Exit Sub
    End If
    
    TestaInferior TBLFuncionário, lAllowEdit, lAllowDelete, lAllowConsult
    TestaSuperior TBLFuncionário, lAllowEdit, lAllowDelete, lAllowConsult
    
    If TBLFuncionário.RecordCount = 0 Then
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
            txtNome.SetFocus
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
    
    If lAtualizar Then
        BotãoAtualizar True
    Else
        BotãoAtualizar False
    End If
    
    BarraDeStatus StatusBarAviso
    mdiGeal.StatusBar.Panels("Posição").Visible = True
    ResizeStatusBar
End Sub
Private Sub Form_Deactivate()
    cmdGravar.Enabled = False
    cmdCancelar.Enabled = False
    BotãoImprimir False
End Sub
Private Sub Form_Load()
    On Error GoTo Erro
    
    ZeraCampos
    
    lAllowInsert = Allow("FUNCIONÁRIO", "I")
    lAllowEdit = Allow("FUNCIONÁRIO", "A")
    lAllowDelete = Allow("FUNCIONÁRIO", "E")
    lAllowConsult = Allow("FUNCIONÁRIO", "C")
    
    lInserir = False
    lAlterar = False
    lPush = False
    lInicio = True
    
    FuncionárioAberto = AbreTabela(Dicionário, "USUÁRIO", "FUNCIONÁRIO", DBUsuário, TBLFuncionário, TBLTabela, dbOpenTable)
    
    If FuncionárioAberto Then
        IndiceFuncionárioAtivo = "FUNCIONÁRIO1"
        TBLFuncionário.Index = IndiceFuncionárioAtivo
    Else
        MsgBox "Não consegui abrir a tabela 'Funcionário' !", vbCritical, "Erro"
        Exit Sub
    End If
    
    ParâmetrosAberto = AbreTabela(Dicionário, "SISTEMA", "PARÂMETROS", DBSistema, TBLParâmetros, TBLTabela, dbOpenTable)
    
    If ParâmetrosAberto Then
    Else
        MsgBox "Não consegui abrir a tabela 'Parâmetros' !", vbCritical, "Erro"
        Exit Sub
    End If
    
    BotãoIncluir lAllowInsert
 
    If TBLFuncionário.RecordCount = 0 Then
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
        
    If TBLFuncionário.RecordCount = 0 Or TBLFuncionário.RecordCount = 1 Then
        NavegaçãoSuperior False
    Else
        NavegaçãoInferior lAllowConsult
    End If
   
    Relatório = AddPath(AplicaçãoPath, "REPORT\FUNCIONÁRIO.RPT")
    TotalDatabaseName = 1
    DataBaseName(1) = AddPath(AplicaçãoPath, "DATABASE\USUÁRIO.MDB")
    Exit Sub
    
Erro:
    GeraMensagemDeErro "Funcionário - Load"
    FuncionárioAberto = False
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
    
    Set frmFuncionário = Nothing
End Sub
Private Sub Form_Unload(Cancel As Integer)
    If FuncionárioAberto Then
        TBLFuncionário.Close
    End If
    If ParâmetrosAberto Then
        TBLParâmetros.Close
    End If
    If Forms.Count = 2 Then
        AllBotões False
    End If
End Sub
Private Sub txtBairro_Change()
    FormatMask "@!S30", txtBairro
End Sub
Private Sub txtBairro_KeyPress(KeyAscii As Integer)
    If Not lInserir Then
        If Not lPush Then
            lAlterar = True
            StatusBarAviso = "Alteração"
            BarraDeStatus StatusBarAviso
        End If
    End If
End Sub
Private Sub txtCep_Change()
    NumericOnly txtCep
    FormatMask "99.999-999", txtCep
End Sub
Private Sub txtCep_KeyPress(KeyAscii As Integer)
    If Not lInserir Then
        If Not lPush Then
            lAlterar = True
            StatusBarAviso = "Alteração"
            BarraDeStatus StatusBarAviso
        End If
    End If
End Sub
Private Sub txtCpf_Change()
    NumericOnly txtCpf
    FormatMask "999.999.999-99", txtCpf
End Sub
Private Sub txtCpf_KeyPress(KeyAscii As Integer)
    If Not lInserir Then
        If Not lPush Then
            lAlterar = True
            StatusBarAviso = "Alteração"
            BarraDeStatus StatusBarAviso
        End If
    End If
End Sub
Private Sub txtCpf_LostFocus()
    If Not IsCorrectCPF(txtCpf) Then
        MsgBox "C. P. F. incorreto !", vbCritical, "Erro"
        txtCpf.SetFocus
    End If
End Sub
Private Sub txtCidade_Change()
    FormatMask "@!S30", txtCidade
End Sub
Private Sub txtCidade_KeyPress(KeyAscii As Integer)
    If Not lInserir Then
        If Not lPush Then
            lAlterar = True
            StatusBarAviso = "Alteração"
            BarraDeStatus StatusBarAviso
        End If
    End If
End Sub
Private Sub txtDatadeEntrada_Change()
    If Not lPula Then
        lPula = True
        FormatMask DataMask, txtDatadeEntrada
        lPula = False
    End If
End Sub
Private Sub txtDatadeEntrada_KeyPress(KeyAscii As Integer)
    If Not lInserir Then
        lAlterar = True
        StatusBarAviso = "Alteração"
        BarraDeStatus StatusBarAviso
    End If
End Sub
Private Sub txtDatadeEntrada_LostFocus()
    If StrTran(txtDatadeEntrada.Text, "/") <> Space(8) Then
        lPula = True
        CorrigeData DataMask, txtDatadeEntrada, Date
        lPula = False
        If Not FormatMask(CheckDataMask, txtDatadeEntrada) Then
            Beep
            MsgBox "Data inválida !", vbCritical, "Erro"
            txtDatadeEntrada.SelStart = 0
            txtDatadeEntrada.SetFocus
        End If
    End If
End Sub
Private Sub txtDatadeSaída_Change()
    If Not lPula Then
        lPula = True
        FormatMask DataMask, txtDatadeSaída
        lPula = False
    End If
End Sub
Private Sub txtDatadeSaída_KeyPress(KeyAscii As Integer)
    If Not lInserir Then
        lAlterar = True
        StatusBarAviso = "Alteração"
        BarraDeStatus StatusBarAviso
    End If
End Sub
Private Sub txtDatadeSaída_LostFocus()
    If StrTran(txtDatadeSaída.Text, "/") <> Space(8) Then
        lPula = True
        CorrigeData DataMask, txtDatadeSaída, Date
        lPula = False
        If Not FormatMask(CheckDataMask, txtDatadeSaída) Then
            Beep
            MsgBox "Data inválida !", vbCritical, "Erro"
            txtDatadeSaída.SelStart = 0
            txtDatadeSaída.SetFocus
        End If
    End If
End Sub
Private Sub txtEndereço_Change()
    FormatMask "@S40", txtEndereço
End Sub
Private Sub txtEndereço_KeyPress(KeyAscii As Integer)
    If Not lInserir Then
        If Not lPush Then
            lAlterar = True
            StatusBarAviso = "Alteração"
            BarraDeStatus StatusBarAviso
        End If
    End If
End Sub
Private Sub txtNome_Change()
    FormatMask "@!S40", txtNome
End Sub
Private Sub txtNome_KeyPress(KeyAscii As Integer)
    If Not lInserir Then
        If Not lPush Then
            lAlterar = True
            StatusBarAviso = "Alteração"
            BarraDeStatus StatusBarAviso
        End If
    End If
End Sub
Private Sub txtSalário_Change()
    If Not lPula Then
        lPula = True
        FormatMask "@K 99.999.999,99", txtSalário
        lPula = False
    End If
End Sub
Private Sub txtSalário_KeyPress(KeyAscii As Integer)
    If Not lInserir Then
        lAlterar = True
        StatusBar = "Alteração"
        BarraDeStatus StatusBar
    End If
End Sub
Private Sub txtSalário_LostFocus()
    lPula = True
    FormatMask "@V ##.###.##0,00", txtSalário
    lPula = False
End Sub
Private Sub txtUF_Change()
    UpperOnly txtUF
    LetterOnly txtUF
End Sub
Private Sub txtUF_KeyPress(KeyAscii As Integer)
    If Not lInserir Then
        If Not lPush Then
            lAlterar = True
            StatusBarAviso = "Alteração"
            BarraDeStatus StatusBarAviso
        End If
    End If
End Sub
