VERSION 5.00
Begin VB.Form frmCliente 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cliente"
   ClientHeight    =   5130
   ClientLeft      =   1575
   ClientTop       =   1515
   ClientWidth     =   6540
   Icon            =   "Clientes.frx":0000
   LinkTopic       =   "frmCliente"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5130
   ScaleWidth      =   6540
   Begin VB.TextBox txtC�digo 
      Enabled         =   0   'False
      Height          =   285
      Left            =   780
      TabIndex        =   32
      Top             =   60
      Width           =   1065
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   345
      Left            =   5280
      TabIndex        =   18
      Top             =   4770
      Width           =   1245
   End
   Begin VB.CommandButton cmdGravar 
      Caption         =   "&Gravar"
      Height          =   345
      Left            =   3960
      TabIndex        =   17
      Top             =   4770
      Width           =   1245
   End
   Begin VB.Frame frNegativado 
      Caption         =   " Negativado"
      Height          =   525
      Left            =   0
      TabIndex        =   26
      Top             =   4170
      Width           =   6525
      Begin VB.OptionButton optNegativadoN�o 
         Caption         =   "N�o"
         Height          =   255
         Left            =   3480
         TabIndex        =   16
         Top             =   210
         Value           =   -1  'True
         Width           =   1395
      End
      Begin VB.OptionButton optNegativadoSim 
         Caption         =   "Sim"
         Height          =   255
         Left            =   600
         TabIndex        =   15
         Top             =   210
         Width           =   1395
      End
   End
   Begin VB.Frame frDadosCadastrais 
      Caption         =   " Dados Cadastrais "
      Height          =   3165
      Left            =   0
      TabIndex        =   19
      Top             =   990
      Width           =   6525
      Begin VB.TextBox txtInscrEstRG 
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
         Left            =   4350
         TabIndex        =   10
         Text            =   "*"
         Top             =   2010
         Width           =   2100
      End
      Begin VB.TextBox txtEmail 
         Height          =   285
         Left            =   1200
         TabIndex        =   14
         Text            =   "*"
         Top             =   2760
         Width           =   5235
      End
      Begin VB.TextBox txtFone1 
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
         Left            =   2880
         TabIndex        =   12
         Text            =   "*"
         Top             =   2370
         Width           =   1185
      End
      Begin VB.TextBox txtFone2 
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
         Left            =   5250
         TabIndex        =   13
         Text            =   "*"
         Top             =   2370
         Width           =   1185
      End
      Begin VB.TextBox txtDDD 
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
         TabIndex        =   11
         Text            =   "*"
         Top             =   2370
         Width           =   465
      End
      Begin VB.TextBox txtCgcCpf 
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
         TabIndex        =   9
         Text            =   "*"
         Top             =   1980
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
         TabIndex        =   8
         Text            =   "*"
         Top             =   1650
         Width           =   1305
      End
      Begin VB.TextBox txtUF 
         Height          =   285
         Left            =   1200
         TabIndex        =   7
         Text            =   "*"
         Top             =   1650
         Width           =   435
      End
      Begin VB.TextBox txtCidade 
         Height          =   285
         Left            =   1200
         TabIndex        =   6
         Text            =   "*"
         Top             =   1320
         Width           =   5235
      End
      Begin VB.TextBox txtBairro 
         Height          =   285
         Left            =   1200
         TabIndex        =   5
         Text            =   "*"
         Top             =   990
         Width           =   5235
      End
      Begin VB.TextBox txtEndere�o 
         Height          =   285
         Left            =   1200
         TabIndex        =   4
         Text            =   "*"
         Top             =   660
         Width           =   5235
      End
      Begin VB.TextBox txtNomeRaz�oSocial 
         Height          =   285
         Left            =   1200
         TabIndex        =   3
         Text            =   "*"
         Top             =   330
         Width           =   5235
      End
      Begin VB.Label lblInscrEstRg 
         Caption         =   "Inscr. Est.."
         Height          =   225
         Left            =   3570
         TabIndex        =   34
         Top             =   2040
         Width           =   795
      End
      Begin VB.Label lblEMail 
         Caption         =   "e-mail"
         Height          =   225
         Left            =   180
         TabIndex        =   31
         Top             =   2760
         Width           =   855
      End
      Begin VB.Label lblFone2 
         Caption         =   "Fone 2"
         Height          =   225
         Left            =   4620
         TabIndex        =   30
         Top             =   2430
         Width           =   585
      End
      Begin VB.Label lblFone1 
         Caption         =   "Fone (1)"
         Height          =   225
         Left            =   2190
         TabIndex        =   29
         Top             =   2430
         Width           =   765
      End
      Begin VB.Label lblDDD 
         Caption         =   "DDD"
         Height          =   225
         Left            =   150
         TabIndex        =   28
         Top             =   2430
         Width           =   765
      End
      Begin VB.Label lblCgcCpf 
         Caption         =   "C. P. F."
         Height          =   195
         Left            =   150
         TabIndex        =   27
         Top             =   2040
         Width           =   645
      End
      Begin VB.Label lblCep 
         Caption         =   "CEP"
         Height          =   195
         Left            =   4680
         TabIndex        =   25
         Top             =   1710
         Width           =   315
      End
      Begin VB.Label lblUF 
         Caption         =   "U. F."
         Height          =   225
         Left            =   150
         TabIndex        =   24
         Top             =   1680
         Width           =   405
      End
      Begin VB.Label lblCidade 
         Caption         =   "Cidade"
         Height          =   225
         Left            =   150
         TabIndex        =   23
         Top             =   1350
         Width           =   945
      End
      Begin VB.Label lblBairro 
         Caption         =   "Bairro"
         Height          =   225
         Left            =   150
         TabIndex        =   22
         Top             =   1020
         Width           =   945
      End
      Begin VB.Label lblEndere�o 
         Caption         =   "Endere�o"
         Height          =   195
         Left            =   150
         TabIndex        =   21
         Top             =   690
         Width           =   975
      End
      Begin VB.Label lblNomeRaz�oSocial 
         Caption         =   "Nome"
         Height          =   195
         Left            =   150
         TabIndex        =   20
         Top             =   360
         Width           =   1065
      End
   End
   Begin VB.Frame frTipo 
      Caption         =   " Tipo "
      Height          =   585
      Left            =   0
      TabIndex        =   0
      Top             =   390
      Width           =   6525
      Begin VB.OptionButton optPessoaJur�dica 
         Caption         =   "Pessoa Jur�dica"
         Height          =   195
         Left            =   2340
         TabIndex        =   2
         Top             =   270
         Width           =   2175
      End
      Begin VB.OptionButton optPessoaF�sica 
         Caption         =   "Pessoa F�sica"
         Height          =   255
         Left            =   450
         TabIndex        =   1
         Top             =   240
         Value           =   -1  'True
         Width           =   1755
      End
   End
   Begin VB.Label lblC�digo 
      Caption         =   "C�digo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   60
      TabIndex        =   33
      Top             =   90
      Width           =   675
   End
End
Attribute VB_Name = "frmCliente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim TBLCliente As Table
Dim ClienteAberto As Boolean
Dim IndiceClienteAtivo$

Dim TBLPar�metros As Table
Dim Par�metrosAberto As Boolean

Dim StatusBarAviso$

Dim lPula As Boolean

Public lInserir As Boolean
Public lAlterar As Boolean

Dim lAllowInsert  As Boolean
Dim lAllowEdit    As Boolean
Dim lAllowDelete  As Boolean
Dim lAllowConsult As Boolean

Dim mFechar As Boolean
Dim lPush As Boolean
Dim lPessoaF�sica As Boolean
Dim lPessoaJur�dica As Boolean
Dim lNegativadoSim As Boolean
Dim lNegativadoN�o As Boolean
Dim lInicio As Boolean

Dim DataBaseName(1 To 1) As String
Public Relat�rio$
Public TotalDatabaseName%

Public lAtualizar As Boolean
Private Sub AtivaCampos()
    Bot�oImprimir True
    frTipo.Enabled = True
    frDadosCadastrais.Enabled = True
    frNegativado.Enabled = True
    Bot�oGravar (lInserir Or lAlterar)
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
    
    lInserir = False
    lAlterar = False
    Bot�oIncluir lAllowInsert
    
    If TBLCliente.RecordCount = 0 Then
        Navega��oInferior False
        Navega��oSuperior False
        Bot�oGravar False
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
    
    TestaInferior TBLCliente, lAllowEdit, lAllowDelete, lAllowConsult
    TestaSuperior TBLCliente, lAllowEdit, lAllowDelete, lAllowConsult
    
    GetRecords
End Function
Public Function PushDataBaseName(ByVal Posi��o As Integer) As String
    PushDataBaseName = DataBaseName(Posi��o)
End Function
Private Sub DesativaCampos()
    Bot�oImprimir False
    frTipo.Enabled = False
    frDadosCadastrais.Enabled = False
    frNegativado.Enabled = False
    cmdGravar.Enabled = False
    cmdCancelar.Enabled = False
    Bot�oGravar False
End Sub
Public Sub Encontrar()
    If Not lAllowConsult Then
        Exit Sub
    End If
    Set frmEncontrar.DBBancoDeDados = DBCadastro
    frmEncontrar.NomeDaJanela = "Cliente"
    frmEncontrar.LabelDescription = "Nome/Raz�o Social"
    frmEncontrar.Mensagem = "Nenhum cliente foi selecionado!"
    frmEncontrar.BancoDeDados = "CADASTRO"
    frmEncontrar.Tabela = "CLIENTE"
    frmEncontrar.Indice = "2"
    frmEncontrar.CampoChave = "C�DIGO"
    frmEncontrar.CampoPreencheLista = "NOME - RAZ�O SOCIAL"
    frmEncontrar.Show vbModal
    lPula = True
    txtC�digo = frmEncontrar.Chave
    lPula = False
    PosRecords
End Sub
Public Sub Excluir()
    Dim Confirma��o As Integer, Msg1$, Msg2$
  
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
    
    TBLCliente.Delete
    
    If Err <> 0 Then
        GeraMensagemDeErro "Cliente - Excluir - " & TBLCliente("NOME - RAZ�O SOCIAL"), True
        StatusBarAviso = "Falha na exclus�o"
        BarraDeStatus StatusBarAviso
        Exit Sub
    End If
    
    WS.CommitTrans
    
    Log gUsu�rio, "Exclus�o - Cliente: " & txtNomeRaz�oSocial
    
    StatusBarAviso = "Exclus�o bem sucedida"
    BarraDeStatus StatusBarAviso
    
    If TBLCliente.RecordCount = 0 Then
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
    
    If TBLCliente.BOF Then
        TBLCliente.MoveFirst
    ElseIf TBLCliente.EOF Then
        TBLCliente.MoveLast
    Else
        TBLCliente.MovePrevious
        If TBLCliente.BOF Then
            TBLCliente.MoveNext
        End If
    End If
    
    GetRecords
    
    TestaInferior TBLCliente, lAllowEdit, lAllowDelete, lAllowConsult
    TestaSuperior TBLCliente, lAllowEdit, lAllowDelete, lAllowConsult
End Sub
Public Sub Gravar()
    If lInserir Then
        'Pega o novo c�digo interno de funcion�rio e atualiza na Tabela Par�metros
        On Error Resume Next
        txtC�digo = TBLPar�metros("CLIENTE") + 1
        If Err.Number <> 0 Then
            txtC�digo = "1"
        End If
        On Error GoTo 0
        TBLPar�metros.Edit
        TBLPar�metros("CLIENTE") = txtC�digo
        TBLPar�metros.Update
        If SetRecords Then
            PosRecords
            lInserir = False
            StatusBarAviso = "Inclus�o bem sucedida"
        Else
            StatusBarAviso = "Falha na inclus�o"
        End If
    Else
        If TBLCliente.RecordCount > 0 And Not TBLCliente.BOF And Not TBLCliente.EOF Then
            If SetRecords Then
                PosRecords
                lAlterar = False
                StatusBarAviso = "Altera��o bem sucedida"
            Else
                StatusBarAviso = "Falha na altera��o"
            End If
        End If
    End If
    
    BarraDeStatus StatusBarAviso
    
    TestaInferior TBLCliente, lAllowEdit, lAllowDelete, lAllowConsult
    TestaSuperior TBLCliente, lAllowEdit, lAllowDelete, lAllowConsult
    
    If TBLCliente.RecordCount = 0 Then
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
    
    If txtNomeRaz�oSocial.Enabled Then
        txtNomeRaz�oSocial.SetFocus
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
    
    Bot�oGravar (lInserir Or lAllowEdit)
    Bot�oIncluir False
    cmdGravar.Enabled = (lInserir Or lAllowEdit)
    cmdCancelar.Enabled = (lInserir Or lAllowEdit)
    
    Navega��oInferior False
    Navega��oSuperior False
    
    StatusBarAviso = "Inclus�o"
    BarraDeStatus StatusBarAviso
    
    optPessoaF�sica.SetFocus

End Sub
Public Sub MoveFirst()
    If lAlterar Then
       If Not Cancelamento Then
            Exit Sub
        End If
    End If
    
    StatusBarAviso = "Pronto"
    BarraDeStatus StatusBarAviso
    
    TBLCliente.MoveFirst
    
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
    
    TBLCliente.MoveLast
    
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
    
    TBLCliente.MoveNext
    If TBLCliente.EOF Then
        TBLCliente.MovePrevious
        Exit Sub
    End If
    
    StatusBarAviso = "Pronto"
    BarraDeStatus StatusBarAviso
    
    Navega��oInferior lAllowConsult
    TestaSuperior TBLCliente, lAllowEdit, lAllowDelete, lAllowConsult
    
    GetRecords
End Sub
Public Sub MovePrevious()
    If lAlterar Then
       If Not Cancelamento Then
            Exit Sub
        End If
    End If
    
    TBLCliente.MovePrevious
    If TBLCliente.BOF Then
        TBLCliente.MoveNext
        Exit Sub
    End If
    
    StatusBarAviso = "Pronto"
    BarraDeStatus StatusBarAviso
    
    Navega��oSuperior lAllowConsult
    TestaInferior TBLCliente, lAllowEdit, lAllowDelete, lAllowConsult
    
    GetRecords
End Sub
Public Sub PosRecords()
    If TBLCliente.RecordCount = 0 Then
        Exit Sub
    End If
    
    TBLCliente.Seek "=", txtC�digo
    If TBLCliente.NoMatch Then
        MsgBox "N�o consegui encontrar o cliente com CGC/CPF " + txtCgcCpf, vbExclamation, "Erro"
        TBLCliente.MoveFirst
        Navega��oInferior False
        Navega��oInferior lAllowConsult
    Else
        TestaInferior TBLCliente, lAllowEdit, lAllowDelete, lAllowConsult
        TestaSuperior TBLCliente, lAllowEdit, lAllowDelete, lAllowConsult
    End If
    GetRecords
End Sub
Private Sub GetRecords()
    On Error GoTo Erro
    
    lPush = True
    
    ZeraCampos
    
    If Not lAllowConsult Then
        ZeraCampos
        DesativaCampos
        lPula = False
        Exit Sub
    End If
    If TBLCliente("TIPO") = "F" Then
        optPessoaF�sica = True
        optPessoaJur�dica = False
        lPessoaF�sica = True
        lPessoaJur�dica = False
        
        lblNomeRaz�oSocial = "Nome"
        lblCgcCpf = "C. P. F."
        lblInscrEstRg = "R. G."
        FormatMask "###.###.###-##", txtCgcCpf
    Else
        optPessoaF�sica = False
        optPessoaJur�dica = True
        lPessoaF�sica = False
        lPessoaJur�dica = True
        
        lblNomeRaz�oSocial = "Raz�o Social"
        lblCgcCpf.Caption = "C. G. C."
        lblInscrEstRg = "Inscr. Est."
        FormatMask "##.###.###/####-##", txtCgcCpf
    End If
    
    txtC�digo = TBLCliente("C�DIGO")
    txtNomeRaz�oSocial = TBLCliente("NOME - RAZ�O SOCIAL")
    txtEndere�o = TBLCliente("ENDERE�O")
    txtBairro = TBLCliente("BAIRRO")
    txtCidade = TBLCliente("CIDADE")
    txtUF = TBLCliente("UF")
    txtCep = TBLCliente("CEP")
    txtCgcCpf = TBLCliente("CGC - CPF")
    txtInscrEstRG = TBLCliente("RG - INSCR ESTADUAL")
    txtDDD = TBLCliente("DDD")
    txtFone1 = TBLCliente("FONE (1)")
    txtFone2 = TBLCliente("FONE (2)")
    txtEMail = TBLCliente("E-MAIL")
    If TBLCliente("NEGATIVADO") Then
        optNegativadoSim = True
        optNegativadoN�o = False
        lNegativadoSim = True
        lNegativadoN�o = False
    Else
        optNegativadoSim = False
        optNegativadoN�o = True
        lNegativadoSim = False
        lNegativadoN�o = True
    End If
    lPush = False
    If Not lAllowEdit Then
        DesativaCampos
    End If
    
    Exit Sub
    
Erro:
    If Err.Number <> 94 Then
        GeraMensagemDeErro "Cliente - GetRecords"
    End If
    Resume Next
End Sub
Private Function SetRecords()
    On Error GoTo Erro
    
    Dim Msg$
    Dim Confirma��o As Integer, Msg1$, Msg2$
    
    WS.BeginTrans 'Inicia uma Transa��o
    
    If lInserir Then
        TBLCliente.AddNew
        TBLCliente("C�DIGO") = txtC�digo
    Else
        TBLCliente.Edit
    End If
    
    If optPessoaF�sica Then
        TBLCliente("TIPO") = "F"
    Else
        TBLCliente("TIPO") = "J"
    End If
    TBLCliente("NOME - RAZ�O SOCIAL") = txtNomeRaz�oSocial
    TBLCliente("ENDERE�O") = txtEndere�o
    TBLCliente("BAIRRO") = txtBairro
    TBLCliente("CIDADE") = txtCidade
    TBLCliente("UF") = txtUF
    TBLCliente("CEP") = txtCep
    TBLCliente("CGC - CPF") = txtCgcCpf
    TBLCliente("RG - INSCR ESTADUAL") = txtInscrEstRG
    TBLCliente("DDD") = txtDDD
    TBLCliente("FONE (1)") = txtFone1
    TBLCliente("FONE (2)") = txtFone2
    TBLCliente("E-MAIL") = txtEMail
    If optNegativadoSim Then
        TBLCliente("NEGATIVADO") = True
    Else
        TBLCliente("NEGATIVADO") = False
    End If
    If lInserir Then
        TBLCliente("USERNAME - CRIA") = gUsu�rio
        TBLCliente("DATA - CRIA") = Date
        TBLCliente("HORA - CRIA") = Time
        TBLCliente("USERNAME - ALTERA") = "VAZIO"
        TBLCliente("DATA - ALTERA") = vbNull
        TBLCliente("HORA - ALTERA") = vbNull
    End If
    If lAlterar Then
        TBLCliente("USERNAME - ALTERA") = gUsu�rio
        TBLCliente("DATA - ALTERA") = Date
        TBLCliente("HORA - ALTERA") = Time
    End If
    TBLCliente.Update
               
Erro:
    If Err <> 0 Then
        TBLCliente.CancelUpdate
        GeraMensagemDeErro "Cliente - SetRecords - " & TBLCliente("NOME - RAZ�O SOCIAL"), True
        SetRecords = False
        Exit Function
    End If

    WS.CommitTrans 'Grava as altera��es ou inclus�es se n�o houverem erros
    
    If lInserir Then
        Log gUsu�rio, "Inclus�o - Cliente: " & txtNomeRaz�oSocial
    Else
        Log gUsu�rio, "Altera��o - Cliente: " & txtNomeRaz�oSocial
    End If
    
    SetRecords = True
End Function
Private Sub ZeraCampos()
    optPessoaF�sica = True
    optPessoaJur�dica = False
    txtC�digo = Empty
    txtNomeRaz�oSocial = Empty
    txtEndere�o = Empty
    txtBairro = Empty
    txtCidade = Empty
    txtUF = Empty
    txtCep = Empty
    txtCgcCpf = Empty
    txtInscrEstRG = Empty
    txtDDD = Empty
    txtFone1 = Empty
    txtFone2 = Empty
    txtEMail = Empty
    optNegativadoSim = False
    optNegativadoN�o = True
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
    If Not ClienteAberto Then
        Unload Me
        Exit Sub
    End If
    If Not Par�metrosAberto Then
        Unload Me
        Exit Sub
    End If
    
    TestaInferior TBLCliente, lAllowEdit, lAllowDelete, lAllowConsult
    TestaSuperior TBLCliente, lAllowEdit, lAllowDelete, lAllowConsult
    If TBLCliente.RecordCount = 0 Then
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
            txtNomeRaz�oSocial.SetFocus
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
    
    If lAtualizar Then
        Bot�oAtualizar True
    Else
        Bot�oAtualizar False
    End If
    
    If lAtualizar Then
        Bot�oAtualizar True
    Else
        Bot�oAtualizar False
    End If
    
    BarraDeStatus StatusBarAviso
    mdiGeal.StatusBar.Panels("Posi��o").Visible = True
    ResizeStatusBar
End Sub
Private Sub Form_Deactivate()
    cmdGravar.Enabled = False
    cmdCancelar.Enabled = False
    Bot�oImprimir False
End Sub
Private Sub Form_Load()
    On Error GoTo Erro
    
    lAllowInsert = Allow("CLIENTE", "I")
    lAllowEdit = Allow("CLIENTE", "A")
    lAllowDelete = Allow("CLIENTE", "E")
    lAllowConsult = Allow("CLIENTE", "C")
    
    ZeraCampos
    
    lPula = False
    lInserir = False
    lAlterar = False
    lPush = False
    lInicio = True
    
    ClienteAberto = AbreTabela(Dicion�rio, "CADASTRO", "CLIENTE", DBCadastro, TBLCliente, TBLTabela, dbOpenTable)
    
    If ClienteAberto Then
        IndiceClienteAtivo = "CLIENTE1"
        TBLCliente.Index = IndiceClienteAtivo
    Else
        MsgBox "N�o consegui abrir a tabela 'Cliente' !", vbCritical, "Erro"
        Exit Sub
    End If
    
    Par�metrosAberto = AbreTabela(Dicion�rio, "SISTEMA", "PAR�METROS", DBSistema, TBLPar�metros, TBLTabela, dbOpenTable)
    
    If Par�metrosAberto Then
    Else
        MsgBox "N�o consegui abrir a tabela 'Par�metros' !", vbCritical, "Erro"
        Exit Sub
    End If
    
    Bot�oIncluir lAllowInsert
 
    If TBLCliente.RecordCount = 0 Then
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
        
    If TBLCliente.RecordCount = 0 Or TBLCliente.RecordCount = 1 Then
        Navega��oSuperior False
    Else
        Navega��oInferior lAllowConsult
    End If
   
    Relat�rio = AddPath(Aplica��oPath, "REPORT\CLIENTE.RPT")
    TotalDatabaseName = 1
    DataBaseName(1) = AddPath(Aplica��oPath, "DATABASE\CADASTRO.MDB")
    
    Exit Sub
    
Erro:
    GeraMensagemDeErro "Cliente - Load"
    ClienteAberto = False
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
    
    Set frmCliente = Nothing
End Sub
Private Sub Form_Unload(Cancel As Integer)
    If ClienteAberto Then
        TBLCliente.Close
    End If
    If Par�metrosAberto Then
        TBLPar�metros.Close
    End If
    If Forms.Count = 2 Then
        AllBot�es False
    End If
End Sub
Private Sub optNegativadoN�o_Click()
    If lPula Then
        Exit Sub
    End If
    If Not lInserir Then
        If Not lPush And optNegativadoN�o <> lNegativadoN�o Then
            lAlterar = True
            StatusBarAviso = "Altera��o"
            BarraDeStatus StatusBarAviso
        End If
    End If
End Sub
Private Sub optNegativadoN�o_KeyPress(KeyAscii As Integer)
    If Not lInserir Then
        If Not lPush And optNegativadoN�o <> lNegativadoN�o Then
            lAlterar = True
            StatusBarAviso = "Altera��o"
            BarraDeStatus StatusBarAviso
        End If
    End If
End Sub
Private Sub optNegativadoSim_Click()
    If lPula Then
        Exit Sub
    End If
    If Not lInserir Then
        If Not lPush And optNegativadoSim <> lNegativadoSim Then
            lAlterar = True
            StatusBarAviso = "Altera��o"
            BarraDeStatus StatusBarAviso
        End If
    End If
End Sub
Private Sub optNegativadoSim_KeyPress(KeyAscii As Integer)
    If Not lInserir Then
        If Not lPush And optNegativadoSim <> lNegativadoSim Then
            lAlterar = True
            StatusBarAviso = "Altera��o"
            BarraDeStatus StatusBarAviso
        End If
    End If
End Sub
Private Sub optPessoaF�sica_Click()
    If lPula Then
        Exit Sub
    End If
    lblNomeRaz�oSocial = "Nome"
    lblCgcCpf = "C. P. F."
    lblInscrEstRg = "R. G."
    FormatMask "###.###.###-##", txtCgcCpf
    If Not lInserir Then
        If Not lPush And optPessoaF�sica <> lPessoaF�sica Then
            lAlterar = True
            StatusBarAviso = "Altera��o"
            BarraDeStatus StatusBarAviso
        End If
    End If
End Sub
Private Sub optPessoaF�sica_KeyPress(KeyAscii As Integer)
    lblNomeRaz�oSocial = "Nome"
    lblCgcCpf = "C. P. F."
    FormatMask "###.###.###-##", txtCgcCpf
    If Not lInserir Then
        If Not lPush And optPessoaF�sica <> lPessoaF�sica Then
            lAlterar = True
            StatusBarAviso = "Altera��o"
            BarraDeStatus StatusBarAviso
        End If
    End If
End Sub
Private Sub optPessoaJur�dica_Click()
    If lPula Then
        Exit Sub
    End If
    lblNomeRaz�oSocial = "Raz�o Social"
    lblCgcCpf.Caption = "C. G. C."
    lblInscrEstRg = "Inscr. Est."
    FormatMask "##.###.###/####-##", txtCgcCpf
    If Not lInserir Then
        If Not lPush And optPessoaJur�dica <> lPessoaJur�dica Then
            lblCgcCpf.Caption = "C. G. C."
            FormatMask "##.###.###/####-##", txtCgcCpf
            lAlterar = True
            StatusBarAviso = "Altera��o"
            BarraDeStatus StatusBarAviso
        End If
    End If
End Sub
Private Sub optPessoaJur�dica_KeyPress(KeyAscii As Integer)
    lblNomeRaz�oSocial = "Raz�o Social"
    lblCgcCpf.Caption = "C. G. C."
    FormatMask "##.###.###/####-##", txtCgcCpf
    If Not lInserir Then
        If Not lPush And optPessoaJur�dica <> lPessoaJur�dica Then
            lAlterar = True
            StatusBarAviso = "Altera��o"
            BarraDeStatus StatusBarAviso
        End If
    End If
End Sub
Private Sub txtBairro_Change()
    If Not lPula Then
        FormatMask "@!S30", txtBairro
    End If
End Sub
Private Sub txtBairro_KeyPress(KeyAscii As Integer)
    If Not lInserir Then
        If Not lPush Then
            lAlterar = True
            StatusBarAviso = "Altera��o"
            BarraDeStatus StatusBarAviso
        End If
    End If
End Sub
Private Sub txtCep_Change()
    If Not lPula Then
        NumericOnly txtCep
        FormatMask "99.999-999", txtCep
    End If
End Sub
Private Sub txtCep_KeyPress(KeyAscii As Integer)
    If Not lInserir Then
        If Not lPush Then
            lAlterar = True
            StatusBarAviso = "Altera��o"
            BarraDeStatus StatusBarAviso
        End If
    End If
End Sub
Private Sub txtCgcCpf_Change()
    If lPula Then
        Exit Sub
    End If
    NumericOnly txtCgcCpf
    If optPessoaF�sica Then
        FormatMask "999.999.999-99", txtCgcCpf
    Else
        FormatMask "99.999.999/9999-99", txtCgcCpf
    End If
End Sub
Private Sub txtCgcCpf_KeyPress(KeyAscii As Integer)
    If Not lInserir Then
        If Not lPush Then
            lAlterar = True
            StatusBarAviso = "Altera��o"
            BarraDeStatus StatusBarAviso
        End If
    End If
End Sub
Private Sub txtCgcCpf_LostFocus()
    If optPessoaF�sica Then
        If Not IsCorrectCPF(txtCgcCpf) Then
            MsgBox "C. P. F. incorreto !", vbCritical, "Erro"
            txtCgcCpf.SetFocus
        End If
    Else
        If Not IsCorrectCGC(txtCgcCpf) Then
            MsgBox "C. G. C. incorreto !", vbCritical, "Erro"
            txtCgcCpf.SetFocus
        End If
    End If
End Sub
Private Sub txtCidade_Change()
    If Not lPula Then
        FormatMask "@!S30", txtCidade
    End If
End Sub
Private Sub txtCidade_KeyPress(KeyAscii As Integer)
    If Not lInserir Then
        If Not lPush Then
            lAlterar = True
            StatusBarAviso = "Altera��o"
            BarraDeStatus StatusBarAviso
        End If
    End If
End Sub
Private Sub txtDDD_Change()
    If Not lPula Then
        FormatMask "999", txtDDD
    End If
End Sub
Private Sub txtDDD_KeyPress(KeyAscii As Integer)
    If Not lInserir Then
        If Not lPush Then
            lAlterar = True
            StatusBarAviso = "Altera��o"
            BarraDeStatus StatusBarAviso
        End If
    End If
End Sub
Private Sub txtEMail_Change()
    If Not lPula Then
        FormatMask "@!S40", txtNomeRaz�oSocial
    End If
End Sub
Private Sub txtEMail_KeyPress(KeyAscii As Integer)
    If Not lInserir Then
        If Not lPush Then
            lAlterar = True
            StatusBarAviso = "Altera��o"
            BarraDeStatus StatusBarAviso
        End If
    End If
End Sub
Private Sub txtEndere�o_Change()
    If Not lPula Then
        FormatMask "@S40", txtEndere�o
    End If
End Sub
Private Sub txtEndere�o_KeyPress(KeyAscii As Integer)
    If Not lInserir Then
        If Not lPush Then
            lAlterar = True
            StatusBarAviso = "Altera��o"
            BarraDeStatus StatusBarAviso
        End If
    End If
End Sub
Private Sub txtFone1_Change()
    If Not lPula Then
        FormatMask "#999-9999", txtFone1
    End If
End Sub
Private Sub txtFone1_KeyPress(KeyAscii As Integer)
    If Not lInserir Then
        If Not lPush Then
            lAlterar = True
            StatusBarAviso = "Altera��o"
            BarraDeStatus StatusBarAviso
        End If
    End If
End Sub
Private Sub txtFone2_Change()
    If Not lPula Then
        FormatMask "#999-9999", txtFone2
    End If
End Sub
Private Sub txtFone2_KeyPress(KeyAscii As Integer)
    If Not lInserir Then
        If Not lPush Then
            lAlterar = True
            StatusBarAviso = "Altera��o"
            BarraDeStatus StatusBarAviso
        End If
    End If
End Sub
Private Sub txtNomeRaz�oSocial_Change()
    If Not lPula Then
        FormatMask "@!S40", txtNomeRaz�oSocial
    End If
End Sub
Private Sub txtNomeRaz�oSocial_KeyPress(KeyAscii As Integer)
    If Not lInserir Then
        If Not lPush Then
            lAlterar = True
            StatusBarAviso = "Altera��o"
            BarraDeStatus StatusBarAviso
        End If
    End If
End Sub
Private Sub txtUF_Change()
    If Not lPula Then
        UpperOnly txtUF
        LetterOnly txtUF
    End If
End Sub
Private Sub txtUF_KeyPress(KeyAscii As Integer)
    If Not lInserir Then
        If Not lPush Then
            lAlterar = True
            StatusBarAviso = "Altera��o"
            BarraDeStatus StatusBarAviso
        End If
    End If
End Sub
