VERSION 5.00
Begin VB.Form frmLocal 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Localidade do Produto"
   ClientHeight    =   3195
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6525
   Icon            =   "Local.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   3195
   ScaleWidth      =   6525
   Begin VB.Frame frLocal 
      Height          =   2760
      Left            =   0
      TabIndex        =   10
      Top             =   0
      Width           =   6525
      Begin VB.TextBox txtCódigo 
         Height          =   285
         Left            =   1200
         TabIndex        =   0
         Top             =   300
         Width           =   315
      End
      Begin VB.TextBox txtEndereço 
         Height          =   285
         Left            =   1200
         TabIndex        =   1
         Top             =   690
         Width           =   5235
      End
      Begin VB.TextBox txtBairro 
         Height          =   285
         Left            =   1200
         TabIndex        =   2
         Top             =   1080
         Width           =   5235
      End
      Begin VB.TextBox txtCidade 
         Height          =   285
         Left            =   1200
         TabIndex        =   3
         Top             =   1470
         Width           =   5235
      End
      Begin VB.TextBox txtUF 
         Height          =   285
         Left            =   1200
         TabIndex        =   4
         Top             =   1860
         Width           =   435
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
         Left            =   4380
         TabIndex        =   5
         Text            =   "  .   -   "
         Top             =   1860
         Width           =   1300
      End
      Begin VB.TextBox txtTelefone 
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
         Text            =   "(    )    -    "
         Top             =   2250
         Width           =   1900
      End
      Begin VB.TextBox txtFax 
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
         Left            =   4380
         TabIndex        =   7
         Text            =   "(    )    -    "
         Top             =   2250
         Width           =   1900
      End
      Begin VB.Label lblCódigo 
         Caption         =   "Código"
         Height          =   195
         Left            =   150
         TabIndex        =   18
         Top             =   330
         Width           =   1065
      End
      Begin VB.Label lblEndereço 
         Caption         =   "Endereço"
         Height          =   195
         Left            =   150
         TabIndex        =   17
         Top             =   720
         Width           =   720
      End
      Begin VB.Label lblBairro 
         Caption         =   "Bairro"
         Height          =   225
         Left            =   150
         TabIndex        =   16
         Top             =   1110
         Width           =   945
      End
      Begin VB.Label lblCidade 
         Caption         =   "Cidade"
         Height          =   225
         Left            =   150
         TabIndex        =   15
         Top             =   1500
         Width           =   945
      End
      Begin VB.Label lblUF 
         Caption         =   "U. F."
         Height          =   225
         Left            =   150
         TabIndex        =   14
         Top             =   1890
         Width           =   405
      End
      Begin VB.Label lblCep 
         Caption         =   "CEP"
         Height          =   195
         Left            =   3930
         TabIndex        =   13
         Top             =   1890
         Width           =   375
      End
      Begin VB.Label lblTelefone 
         Caption         =   "Telefone"
         Height          =   210
         Left            =   150
         TabIndex        =   12
         Top             =   2280
         Width           =   660
      End
      Begin VB.Label lblFax 
         Caption         =   "Fax"
         Height          =   195
         Left            =   3930
         TabIndex        =   11
         Top             =   2280
         Width           =   345
      End
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   345
      Left            =   5250
      TabIndex        =   9
      Top             =   2820
      Width           =   1245
   End
   Begin VB.CommandButton cmdGravar 
      Caption         =   "&Gravar"
      Default         =   -1  'True
      Height          =   345
      Left            =   3930
      TabIndex        =   8
      Top             =   2820
      Width           =   1245
   End
End
Attribute VB_Name = "frmLocal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim TBLLocal As Table
Dim LocalAberto As Boolean
Dim IndiceLocalAtivo$

Dim lAllowInsert  As Boolean
Dim lAllowEdit    As Boolean
Dim lAllowDelete  As Boolean
Dim lAllowConsult As Boolean

Dim lPula As Boolean
Dim lInserir As Boolean
Dim lAlterar As Boolean

Dim StatusBarAviso$

Public lAtualizar As Boolean
Private Sub AtivaCampos()
    frLocal.Enabled = True
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
    
    If TBLLocal.RecordCount = 0 Then
        NavegaçãoInferior False
        NavegaçãoSuperior False
        BotãoGravar False
        cmdGravar.Enabled = False
        cmdCancelar.Enabled = False
        DesativaCampos
        ZeraCampos
        Cancelamento = True
        Exit Function
    End If
    
    Cancelamento = True
    
    TestaInferior TBLLocal, lAllowEdit, lAllowDelete, lAllowConsult
    TestaSuperior TBLLocal, lAllowEdit, lAllowDelete, lAllowConsult
        
    GetRecords
End Function
Private Sub DesativaCampos()
    frLocal.Enabled = False
    cmdGravar.Enabled = False
    cmdCancelar.Enabled = False
    BotãoGravar False
End Sub
Public Sub Encontrar()
    If Not lAllowConsult Then
        Exit Sub
    End If
    Set frmEncontrar.DBBancoDeDados = DBCadastro
    frmEncontrar.NomeDaJanela = "Localidade"
    frmEncontrar.LabelDescription = "Endereço"
    frmEncontrar.Mensagem = "Nenhuma localidade foi selecionado!"
    frmEncontrar.BancoDeDados = "CADASTRO"
    frmEncontrar.Tabela = "LOCAL DO PRODUTO"
    frmEncontrar.Indice = "2"
    frmEncontrar.CampoChave = "CÓDIGO"
    frmEncontrar.CampoPreencheLista = "ENDEREÇO"
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
    
    TBLLocal.Delete
    
    If Err <> 0 Then
        GeraMensagemDeErro "Local - Excluir - " & txtEndereço, True
        StatusBarAviso = "Falha na exclusão"
        BarraDeStatus StatusBarAviso
        Exit Sub
    End If
    
    WS.CommitTrans
    
    Log gUsuário, "Exclusão - Produto: " & txtCódigo & " - " & txtEndereço
    
    StatusBarAviso = "Exclusão bem sucedida"
    BarraDeStatus StatusBarAviso
    
    If TBLLocal.RecordCount = 0 Then
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
    
    If TBLLocal.BOF Then
        TBLLocal.MoveFirst
    ElseIf TBLLocal.EOF Then
        TBLLocal.MoveLast
    Else
        TBLLocal.MovePrevious
        If TBLLocal.BOF Then
            TBLLocal.MoveNext
        End If
    End If
    
    GetRecords
    
    TestaInferior TBLLocal, lAllowEdit, lAllowDelete, lAllowConsult
    TestaSuperior TBLLocal, lAllowEdit, lAllowDelete, lAllowConsult
End Sub
Public Sub Gravar()
    If lInserir Then
        If SetRecords Then
            PosRecords
            lInserir = False
            StatusBarAviso = "Inclusão bem sucedida"
        Else
            StatusBarAviso = "Falha na inclusão"
        End If
    Else
        If TBLLocal.RecordCount > 0 And Not TBLLocal.BOF And Not TBLLocal.EOF Then
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
    
    TestaInferior TBLLocal, lAllowEdit, lAllowDelete, lAllowConsult
    TestaSuperior TBLLocal, lAllowEdit, lAllowDelete, lAllowConsult
    
    If TBLLocal.RecordCount = 0 Then
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
    
    If txtCódigo.Enabled Then
        txtCódigo.SetFocus
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
    
    txtCódigo.SetFocus
End Sub
Public Sub MoveFirst()
    If lAlterar Then
       If Not Cancelamento Then
            Exit Sub
        End If
    End If
    
    StatusBarAviso = "Pronto"
    BarraDeStatus StatusBarAviso
    
    TBLLocal.MoveFirst
    
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
    
    TBLLocal.MoveLast
    
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
    
    TBLLocal.MoveNext
    If TBLLocal.EOF Then
        TBLLocal.MovePrevious
        Exit Sub
    End If
    
    StatusBarAviso = "Pronto"
    BarraDeStatus StatusBarAviso
    
    NavegaçãoInferior lAllowConsult
    TestaSuperior TBLLocal, lAllowEdit, lAllowDelete, lAllowConsult
    
    GetRecords
End Sub
Public Sub MovePrevious()
    If lAlterar Then
       If Not Cancelamento Then
            Exit Sub
        End If
    End If
    
    TBLLocal.MovePrevious
    If TBLLocal.BOF Then
        TBLLocal.MoveNext
        Exit Sub
    End If
    
    StatusBarAviso = "Pronto"
    BarraDeStatus StatusBarAviso
    
    NavegaçãoSuperior lAllowConsult
    TestaInferior TBLLocal, lAllowEdit, lAllowDelete, lAllowConsult
    
    GetRecords
End Sub
Public Sub PosRecords()
    If TBLLocal.RecordCount = 0 Then
        Exit Sub
    End If
    
    TBLLocal.Seek "=", txtCódigo
    If TBLLocal.NoMatch Then
        MsgBox "Não consegui encontrar " + txtCódigo, vbExclamation, "Erro"
        TBLLocal.MoveFirst
        NavegaçãoInferior False
        NavegaçãoInferior lAllowConsult
    Else
        TestaInferior TBLLocal, lAllowEdit, lAllowDelete, lAllowConsult
        TestaSuperior TBLLocal, lAllowEdit, lAllowDelete, lAllowConsult
    End If
    GetRecords
End Sub
Private Sub GetRecords()
    On Error GoTo Erro
    
    If Not lAllowConsult Then
        ZeraCampos
        DesativaCampos
        Exit Sub
    End If
    txtCódigo = TBLLocal("CÓDIGO")
    txtEndereço = TBLLocal("ENDEREÇO")
    txtBairro = TBLLocal("BAIRRO")
    txtCidade = TBLLocal("CIDADE")
    txtUF = TBLLocal("UF")
    txtTelefone = TBLLocal("TELEFONE")
    txtFax = TBLLocal("FAX")
    txtCep = TBLLocal("CEP")
    If Not lAllowEdit Then
        DesativaCampos
    End If
    
    Exit Sub
    
Erro:
    GeraMensagemDeErro "Local - GetRecords "
    Resume Next
End Sub
Private Function SetRecords()
    On Error GoTo Erro
    
    Dim Msg$
    Dim Confirmação As Integer, Msg1$, Msg2$, AchouDepartamentoSeção As Boolean
    
    WS.BeginTrans 'Inicia uma Transação
    
    If lInserir Then
        TBLLocal.AddNew
    Else
        TBLLocal.Edit
    End If
    
    TBLLocal("CÓDIGO") = txtCódigo
    TBLLocal("ENDEREÇO") = txtEndereço
    TBLLocal("BAIRRO") = txtBairro
    TBLLocal("CIDADE") = txtCidade
    TBLLocal("UF") = txtUF
    TBLLocal("TELEFONE") = txtTelefone
    TBLLocal("FAX") = txtFax
    TBLLocal("CEP") = txtCep
    If lInserir Then
        TBLLocal("USERNAME - CRIA") = gUsuário
        TBLLocal("DATA - CRIA") = Date
        TBLLocal("HORA - CRIA") = Time
        TBLLocal("USERNAME - ALTERA") = "VAZIO"
        TBLLocal("DATA - ALTERA") = vbNull
        TBLLocal("HORA - ALTERA") = vbNull
    End If
    If lAlterar Then
        TBLLocal("USERNAME - ALTERA") = gUsuário
        TBLLocal("DATA - ALTERA") = Date
        TBLLocal("HORA - ALTERA") = Time
    End If
    TBLLocal.Update
        
Erro:
    If Err <> 0 Then
        TBLLocal.CancelUpdate
        GeraMensagemDeErro "Local - SetRecords - " & txtEndereço, True
        SetRecords = False
        Exit Function
    End If

    WS.CommitTrans 'Grava as alterações ou inclusões se não houverem erros
        
    If lInserir Then
        Log gUsuário, "Inclusão - Local " & txtCódigo & " - " & txtEndereço
    Else
        Log gUsuário, "Aleração - Local " & txtCódigo & " - " & txtEndereço
    End If
    
    SetRecords = True
End Function
Private Sub ZeraCampos()
    txtCódigo = Empty
    txtEndereço = Empty
    txtBairro = Empty
    txtCidade = Empty
    txtUF = Empty
    txtTelefone = Empty
    txtFax = Empty
    txtCep = Empty
End Sub
Private Sub cmdCancelar_Click()
    Cancelamento
End Sub
Private Sub cmdGravar_Click()
    Gravar
End Sub
Private Sub Form_Activate()
    If Not LocalAberto Then
        Unload frmLocal
        Exit Sub
    End If
    TestaInferior TBLLocal, lAllowEdit, lAllowDelete, lAllowConsult
    TestaSuperior TBLLocal, lAllowEdit, lAllowDelete, lAllowConsult
    If TBLLocal.RecordCount = 0 Then
        BotãoGravar False
        cmdGravar.Enabled = False
        cmdCancelar.Enabled = False
    Else
        BotãoGravar (lInserir Or lAllowEdit)
        cmdGravar.Enabled = (lInserir Or lAllowEdit)
        cmdCancelar.Enabled = (lInserir Or lAllowEdit)
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
End Sub
Private Sub Form_Deactivate()
    cmdGravar.Enabled = False
    cmdCancelar.Enabled = False
End Sub
Private Sub Form_Load()
    ZeraCampos
    
    lPula = False
    lInserir = False
    lAlterar = False
    
    lAllowInsert = Allow("LOCALIDADE DE ESTOQUE", "I")
    lAllowEdit = Allow("LOCALIDADE DE ESTOQUE", "A")
    lAllowDelete = Allow("LOCALIDADE DE ESTOQUE", "E")
    lAllowConsult = Allow("LOCALIDADE DE ESTOQUE", "C")
    
    LocalAberto = AbreTabela(Dicionário, "CADASTRO", "LOCAL DO PRODUTO", DBCadastro, TBLLocal, TBLTabela, dbOpenTable)
    
    If LocalAberto Then
        IndiceLocalAtivo = "LOCALDOPRODUTO1"
        TBLLocal.Index = IndiceLocalAtivo
    Else
        MsgBox "Não consegui abrir a tabela 'LOCAL' !", vbCritical, "Erro"
        Exit Sub
    End If
    
    BotãoIncluir lAllowInsert
 
    If TBLLocal.RecordCount = 0 Then
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
        
    If TBLLocal.RecordCount = 0 Or TBLLocal.RecordCount = 1 Then
        NavegaçãoSuperior False
    Else
        NavegaçãoInferior lAllowConsult
    End If
    
    StatusBarAviso = "Pronto"
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
    
    Set frmLocal = Nothing
End Sub
Private Sub Form_Unload(Cancel As Integer)
    If LocalAberto Then
        TBLLocal.Close
    End If
    If Forms.Count = 2 Then
        AllBotões False
    End If
End Sub
Private Sub txtBairro_Change()
    If Not lPula Then
        FormatMask "@!S30", txtBairro
    End If
End Sub
Private Sub txtCep_Change()
    If Not lPula Then
        FormatMask "99.999-999", txtCep
    End If
End Sub
Private Sub txtCidade_Change()
    If Not lPula Then
        FormatMask "@!S30", txtCidade
    End If
End Sub
Private Sub txtCódigo_Change()
    If Not lPula Then
        FormatMask "99", txtCódigo
    End If
End Sub
Private Sub txtCódigo_LostFocus()
    FormatMask "@N 00", txtCódigo
End Sub
Private Sub txtEndereço_Change()
    If Not lPula Then
        FormatMask "@S40", txtEndereço
    End If
End Sub
Private Sub txtFax_Change()
    If Not lPula Then
        FormatMask "(####)####-####", txtFax
    End If
End Sub
Private Sub txtTelefone_Change()
    If Not lPula Then
        FormatMask "(####)####-####", txtTelefone
    End If
End Sub
Private Sub txtUF_Change()
    If Not lPula Then
        FormatMask "@! AA", txtUF
    End If
End Sub
