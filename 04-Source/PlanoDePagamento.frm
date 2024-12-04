VERSION 5.00
Begin VB.Form frmPlanoDePagamento 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Plano de Pagamento"
   ClientHeight    =   3780
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6120
   Icon            =   "PlanoDePagamento.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   3780
   ScaleWidth      =   6120
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   345
      Left            =   4860
      TabIndex        =   14
      Top             =   3420
      Width           =   1245
   End
   Begin VB.CommandButton cmdGravar 
      Caption         =   "&Gravar"
      Height          =   345
      Left            =   3540
      TabIndex        =   13
      Top             =   3420
      Width           =   1245
   End
   Begin VB.Frame frPlanoDePagamento 
      Height          =   3375
      Left            =   0
      TabIndex        =   15
      Top             =   0
      Width           =   6105
      Begin VB.TextBox txtIntervaloDePagamentos 
         Height          =   285
         Left            =   5160
         TabIndex        =   8
         Top             =   1350
         Width           =   465
      End
      Begin VB.CheckBox chkVenda 
         Alignment       =   1  'Right Justify
         Caption         =   "Venda"
         Height          =   225
         Left            =   1470
         TabIndex        =   7
         Top             =   2970
         Width           =   945
      End
      Begin VB.CheckBox chkCompra 
         Alignment       =   1  'Right Justify
         Caption         =   "Compra"
         Height          =   225
         Left            =   210
         TabIndex        =   6
         Top             =   2970
         Width           =   945
      End
      Begin VB.CheckBox chkCaixa 
         Alignment       =   1  'Right Justify
         Caption         =   "Opção no caixa"
         Height          =   255
         Left            =   3270
         TabIndex        =   12
         Top             =   2970
         Width           =   2445
      End
      Begin VB.CheckBox chkPrimeiroPagamentoBaixado 
         Alignment       =   1  'Right Justify
         Caption         =   "Primeiro pagamento baixado"
         Height          =   405
         Left            =   3270
         TabIndex        =   11
         Top             =   2520
         Width           =   2415
      End
      Begin VB.CheckBox chkPermiteAlterarAutoInclusão 
         Alignment       =   1  'Right Justify
         Caption         =   "Permite Alterar Auto-Inclusão"
         Height          =   375
         Left            =   210
         TabIndex        =   5
         Top             =   2310
         Width           =   2205
      End
      Begin VB.CheckBox chkAutoInclusão 
         Alignment       =   1  'Right Justify
         Caption         =   "Auto-Inclusão"
         Height          =   195
         Left            =   210
         TabIndex        =   4
         Top             =   2010
         Width           =   2205
      End
      Begin VB.TextBox txtRepasse 
         Height          =   285
         Left            =   4980
         TabIndex        =   10
         Top             =   2190
         Width           =   675
      End
      Begin VB.TextBox txtCustoFinanceiro 
         Height          =   285
         Left            =   4980
         TabIndex        =   9
         Top             =   1770
         Width           =   675
      End
      Begin VB.TextBox txtQTVencimentos 
         Height          =   285
         Left            =   1950
         TabIndex        =   3
         Top             =   1590
         Width           =   465
      End
      Begin VB.TextBox txtCarênciaPosVenda 
         Height          =   285
         Left            =   1950
         TabIndex        =   2
         Top             =   1200
         Width           =   465
      End
      Begin VB.TextBox txtDescrição 
         Height          =   285
         Left            =   1320
         TabIndex        =   1
         Top             =   660
         Width           =   4605
      End
      Begin VB.TextBox txtCódigo 
         Height          =   285
         Left            =   1320
         TabIndex        =   0
         Top             =   270
         Width           =   465
      End
      Begin VB.Label lblIntervalosDePagamentos 
         Caption         =   "Intervalos de Pagamento"
         Height          =   435
         Left            =   3300
         TabIndex        =   23
         Top             =   1230
         Width           =   1335
      End
      Begin VB.Label lblRepasse 
         Caption         =   "Repasse"
         Height          =   285
         Left            =   3300
         TabIndex        =   22
         Top             =   2220
         Width           =   705
      End
      Begin VB.Label lblCustoFinaceiro 
         Caption         =   "Custo Financeiro"
         Height          =   225
         Left            =   3300
         TabIndex        =   21
         Top             =   1830
         Width           =   1335
      End
      Begin VB.Label lblDias 
         Caption         =   "dia(s)"
         Height          =   225
         Left            =   2490
         TabIndex        =   20
         Top             =   1290
         Width           =   405
      End
      Begin VB.Label lblQTVencimento 
         Caption         =   "Quantidade de Vencimentos"
         Height          =   405
         Left            =   210
         TabIndex        =   19
         Top             =   1500
         Width           =   1065
      End
      Begin VB.Label lblCarênciaPosVenda 
         Caption         =   "Carência"
         Height          =   225
         Left            =   210
         TabIndex        =   18
         Top             =   1230
         Width           =   945
      End
      Begin VB.Line Line4 
         BorderColor     =   &H00808080&
         X1              =   3000
         X2              =   3000
         Y1              =   1110
         Y2              =   3360
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00FFFFFF&
         X1              =   3030
         X2              =   3030
         Y1              =   1140
         Y2              =   3360
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00808080&
         X1              =   30
         X2              =   6090
         Y1              =   1110
         Y2              =   1110
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         X1              =   30
         X2              =   6090
         Y1              =   1130
         Y2              =   1130
      End
      Begin VB.Label lblDescrição 
         Caption         =   "Descrição"
         Height          =   195
         Left            =   270
         TabIndex        =   17
         Top             =   690
         Width           =   765
      End
      Begin VB.Label lblCódigo 
         Caption         =   "Código"
         Height          =   225
         Left            =   270
         TabIndex        =   16
         Top             =   300
         Width           =   885
      End
   End
End
Attribute VB_Name = "frmPlanoDePagamento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim TBLPlanoDePagamento As Table
Dim PlanoDePagamentoAberto As Boolean
Dim IndicePlanoDePagamentoAtivo$

Dim lAllowInsert  As Boolean
Dim lAllowEdit    As Boolean
Dim lAllowDelete  As Boolean
Dim lAllowConsult As Boolean

Dim lInserir As Boolean
Dim lAlterar As Boolean

Dim lFechar As Boolean
Dim lPula As Boolean

Dim StatusBarAviso$

Dim DataBaseName(1 To 1) As String
Public Relatório$
Public TotalDatabaseName%

Public lAtualizar As Boolean
Private Sub AtivaCampos()
    frPlanoDePagamento.Enabled = True
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
    
    If TBLPlanoDePagamento.RecordCount = 0 Then
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
    
    TestaInferior TBLPlanoDePagamento, lAllowEdit, lAllowDelete, lAllowConsult
    TestaSuperior TBLPlanoDePagamento, lAllowEdit, lAllowDelete, lAllowConsult
    
    GetRecords
End Function
Private Sub DesativaCampos()
    frPlanoDePagamento.Enabled = False
    cmdGravar.Enabled = False
    cmdCancelar.Enabled = False
    BotãoGravar False
End Sub
Public Sub Encontrar()
    If Not lAllowConsult Then
        Exit Sub
    End If
    Set frmEncontrar.DBBancoDeDados = DBFinanceiro
    frmEncontrar.NomeDaJanela = "Plano de Pagamento"
    frmEncontrar.LabelDescription = "Descrição"
    frmEncontrar.Mensagem = "Nenhum Plano de Pagamento foi selecionado!"
    frmEncontrar.BancoDeDados = "FINANCEIRO"
    frmEncontrar.Tabela = "PLANO DE PAGAMENTO"
    frmEncontrar.Indice = "2"
    frmEncontrar.CampoChave = "CÓDIGO"
    frmEncontrar.CampoPreencheLista = "DESCRIÇÃO"
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
    
    TBLPlanoDePagamento.Delete
    
    If Err <> 0 Then
        GeraMensagemDeErro "Plano de Pagamento - Excluir - " & txtDescrição
        WS.Rollback 'Caso haja erro volta os valores normais
        StatusBarAviso = "Falha na exclusão"
        BarraDeStatus StatusBarAviso
        Exit Sub
    End If
    
    WS.CommitTrans
    
    Log gUsuário, "Exclusão - Plano de Pagamento: " & txtCódigo & " - " & txtDescrição
    
    StatusBarAviso = "Exclusão bem sucedida"
    BarraDeStatus StatusBarAviso
    
    If TBLPlanoDePagamento.RecordCount = 0 Then
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
    
    If TBLPlanoDePagamento.BOF Then
        TBLPlanoDePagamento.MoveFirst
    ElseIf TBLPlanoDePagamento.EOF Then
        TBLPlanoDePagamento.MoveLast
    Else
        TBLPlanoDePagamento.MovePrevious
        If TBLPlanoDePagamento.BOF Then
            TBLPlanoDePagamento.MoveNext
        End If
    End If
    
    GetRecords
    
    TestaInferior TBLPlanoDePagamento, lAllowEdit, lAllowDelete, lAllowConsult
    TestaSuperior TBLPlanoDePagamento, lAllowEdit, lAllowDelete, lAllowConsult
End Sub
Private Sub GetRecords()
    lPula = True
    If Not lAllowConsult Then
        ZeraCampos
        DesativaCampos
        lPula = False
        Exit Sub
    End If
    txtCódigo = TBLPlanoDePagamento("CÓDIGO")
    txtDescrição = TBLPlanoDePagamento("DESCRIÇÃO")
    txtCarênciaPosVenda = TBLPlanoDePagamento("CARÊNCIA")
    txtIntervaloDePagamentos = TBLPlanoDePagamento("INTERVALO DE PAGAMENTOS")
    txtQTVencimentos = TBLPlanoDePagamento("QUANTIDADE DE VENCIMENTOS")
    txtCustoFinanceiro = TBLPlanoDePagamento("CUSTO FINANCEIRO")
    chkAutoInclusão.Value = IIf(TBLPlanoDePagamento("AUTO-INCLUSÃO"), 1, 0)
    chkPermiteAlterarAutoInclusão.Value = IIf(TBLPlanoDePagamento("PERMITE ALTERAR AUTO-INCLUSÃO"), 1, 0)
    chkPrimeiroPagamentoBaixado.Value = IIf(TBLPlanoDePagamento("PRIMEIRO PAGAMENTO BAIXADO"), 1, 0)
    chkCaixa.Value = IIf(TBLPlanoDePagamento("CAIXA"), 1, 0)
    chkCompra.Value = IIf(TBLPlanoDePagamento("COMPRA"), 1, 0)
    chkVenda.Value = IIf(TBLPlanoDePagamento("VENDA"), 1, 0)
    lPula = False
    txtCustoFinanceiro_LostFocus
    lPula = True
    txtRepasse = TBLPlanoDePagamento("REPASSE")
    lPula = False
    txtRepasse_LostFocus
    If Not lAllowEdit Then
        DesativaCampos
    End If
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
        If TBLPlanoDePagamento.RecordCount > 0 And Not TBLPlanoDePagamento.BOF And Not TBLPlanoDePagamento.EOF Then
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
    
    TestaInferior TBLPlanoDePagamento, lAllowEdit, lAllowDelete, lAllowConsult
    TestaSuperior TBLPlanoDePagamento, lAllowEdit, lAllowDelete, lAllowConsult
    
    If TBLPlanoDePagamento.RecordCount = 0 Then
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
    
    TBLPlanoDePagamento.MoveFirst
    
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
    
    TBLPlanoDePagamento.MoveLast
    
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
    
    TBLPlanoDePagamento.MoveNext
    If TBLPlanoDePagamento.EOF Then
        TBLPlanoDePagamento.MovePrevious
        Exit Sub
    End If
    
    StatusBarAviso = "Pronto"
    BarraDeStatus StatusBarAviso
    
    NavegaçãoInferior lAllowConsult
    TestaSuperior TBLPlanoDePagamento, lAllowEdit, lAllowDelete, lAllowConsult
    
    GetRecords
End Sub
Public Sub MovePrevious()
    If lAlterar Then
       If Not Cancelamento Then
            Exit Sub
        End If
    End If
    
    TBLPlanoDePagamento.MovePrevious
    If TBLPlanoDePagamento.BOF Then
        TBLPlanoDePagamento.MoveNext
        Exit Sub
    End If
    
    StatusBarAviso = "Pronto"
    BarraDeStatus StatusBarAviso
    
    NavegaçãoSuperior lAllowConsult
    TestaInferior TBLPlanoDePagamento, lAllowEdit, lAllowDelete, lAllowConsult
    
    GetRecords
End Sub
Public Sub PosRecords()
    If TBLPlanoDePagamento.RecordCount = 0 Then
        Exit Sub
    End If
    
    TBLPlanoDePagamento.Seek "=", txtCódigo
    If TBLPlanoDePagamento.NoMatch Then
        MsgBox "Não consegui encontrar " + txtCódigo, vbExclamation, "Erro"
        TBLPlanoDePagamento.MoveFirst
        NavegaçãoInferior False
        NavegaçãoInferior lAllowConsult
    Else
        TestaInferior TBLPlanoDePagamento, lAllowEdit, lAllowDelete, lAllowConsult
        TestaSuperior TBLPlanoDePagamento, lAllowEdit, lAllowDelete, lAllowConsult
    End If
    GetRecords
End Sub
Public Function PushDataBaseName(ByVal Posição As Integer) As String
    PushDataBaseName = DataBaseName(Posição)
End Function
Private Function SetRecords()
    On Error GoTo Erro
    
    Dim Msg$
    Dim Confirmação As Integer, Msg1$, Msg2$
    
    WS.BeginTrans 'Inicia uma Transação
    
    If lInserir Then
        TBLPlanoDePagamento.AddNew
    Else
        TBLPlanoDePagamento.Edit
    End If
    
    TBLPlanoDePagamento("CÓDIGO") = txtCódigo
    TBLPlanoDePagamento("DESCRIÇÃO") = txtDescrição
    TBLPlanoDePagamento("CARÊNCIA") = txtCarênciaPosVenda
    TBLPlanoDePagamento("INTERVALO DE PAGAMENTOS") = txtIntervaloDePagamentos
    TBLPlanoDePagamento("QUANTIDADE DE VENCIMENTOS") = txtQTVencimentos
    TBLPlanoDePagamento("CUSTO FINANCEIRO") = txtCustoFinanceiro
    TBLPlanoDePagamento("REPASSE") = txtRepasse
    TBLPlanoDePagamento("AUTO-INCLUSÃO") = chkAutoInclusão.Value
    TBLPlanoDePagamento("PERMITE ALTERAR AUTO-INCLUSÃO") = chkPermiteAlterarAutoInclusão.Value
    TBLPlanoDePagamento("PRIMEIRO PAGAMENTO BAIXADO") = chkPrimeiroPagamentoBaixado.Value
    TBLPlanoDePagamento("CAIXA") = chkCaixa.Value
    TBLPlanoDePagamento("COMPRA") = chkCompra.Value
    TBLPlanoDePagamento("VENDA") = chkVenda.Value
    If lInserir Then
        TBLPlanoDePagamento("USERNAME - CRIA") = gUsuário
        TBLPlanoDePagamento("DATA - CRIA") = Date
        TBLPlanoDePagamento("HORA - CRIA") = Time
        TBLPlanoDePagamento("USERNAME - ALTERA") = "VAZIO"
        TBLPlanoDePagamento("DATA - ALTERA") = vbNull
        TBLPlanoDePagamento("HORA - ALTERA") = vbNull
    End If
    If lAlterar Then
        TBLPlanoDePagamento("USERNAME - ALTERA") = gUsuário
        TBLPlanoDePagamento("DATA - ALTERA") = Date
        TBLPlanoDePagamento("HORA - ALTERA") = Time
    End If
    TBLPlanoDePagamento.Update
            
    WS.CommitTrans 'Grava as alterações ou inclusões se não houver erros
    
    If lInserir Then
        Log gUsuário, "Inclusão - Plano de Pagamento: " & txtCódigo & " - " & txtDescrição
    Else
        Log gUsuário, "Alteração - Plano de Pagamento: " & txtCódigo & " - " & txtDescrição
    End If
    
    SetRecords = True
    
    Exit Function
    
Erro:
    GeraMensagemDeErro "Plano de Pagamento - SetRecords - " & txtDescrição
    On Error Resume Next
    SetRecords = False
    TBLPlanoDePagamento.CancelUpdate
    WS.Rollback  'Caso haja erro volta os valores normais
    On Error GoTo 0
End Function
Private Sub ZeraCampos()
    lPula = True
    txtCódigo = Empty
    txtDescrição = Empty
    txtCarênciaPosVenda = "0"
    txtIntervaloDePagamentos = "0"
    txtQTVencimentos = "0"
    txtCustoFinanceiro = "0,00"
    lPula = False
    txtCustoFinanceiro_LostFocus
    lPula = True
    txtRepasse = "0,00"
    lPula = False
    txtRepasse_LostFocus
    lPula = False
End Sub
Private Sub chkAutoInclusão_Click()
    If Not lInserir And Not lPula Then
        lAlterar = True
        StatusBarAviso = "Alteração"
        BarraDeStatus StatusBarAviso
    End If
End Sub
Private Sub chkAutoInclusão_KeyPress(KeyAscii As Integer)
    If Not lInserir Then
        lAlterar = True
        StatusBarAviso = "Alteração"
        BarraDeStatus StatusBarAviso
    End If
End Sub
Private Sub chkCaixa_Click()
    If Not lInserir And Not lPula Then
        lAlterar = True
        StatusBarAviso = "Alteração"
        BarraDeStatus StatusBarAviso
    End If
End Sub
Private Sub chkCaixa_KeyPress(KeyAscii As Integer)
    If Not lInserir Then
        lAlterar = True
        StatusBarAviso = "Alteração"
        BarraDeStatus StatusBarAviso
    End If
End Sub
Private Sub chkCompra_Click()
    If Not lInserir And Not lPula Then
        lAlterar = True
        StatusBarAviso = "Alteração"
        BarraDeStatus StatusBarAviso
    End If
End Sub
Private Sub chkCompra_KeyPress(KeyAscii As Integer)
    If Not lInserir Then
        lAlterar = True
        StatusBarAviso = "Alteração"
        BarraDeStatus StatusBarAviso
    End If
End Sub
Private Sub chkPermiteAlterarAutoInclusão_Click()
    If Not lInserir And Not lPula Then
        lAlterar = True
        StatusBarAviso = "Alteração"
        BarraDeStatus StatusBarAviso
    End If
End Sub
Private Sub chkPermiteAlterarAutoInclusão_KeyPress(KeyAscii As Integer)
    If Not lInserir Then
        lAlterar = True
        StatusBarAviso = "Alteração"
        BarraDeStatus StatusBarAviso
    End If
End Sub
Private Sub chkPrimeiroPagamentoBaixado_Click()
    If Not lInserir And Not lPula Then
        lAlterar = True
        StatusBarAviso = "Alteração"
        BarraDeStatus StatusBarAviso
    End If
End Sub
Private Sub chkPrimeiroPagamentoBaixado_KeyPress(KeyAscii As Integer)
    If Not lInserir Then
        lAlterar = True
        StatusBarAviso = "Alteração"
        BarraDeStatus StatusBarAviso
    End If
End Sub
Private Sub chkVenda_Click()
    If Not lInserir And Not lPula Then
        lAlterar = True
        StatusBarAviso = "Alteração"
        BarraDeStatus StatusBarAviso
    End If
End Sub
Private Sub chkVenda_KeyPress(KeyAscii As Integer)
    If Not lInserir Then
        lAlterar = True
        StatusBarAviso = "Alteração"
        BarraDeStatus StatusBarAviso
    End If
End Sub
Private Sub cmdCancelar_Click()
    Cancelamento
End Sub
Private Sub cmdGravar_Click()
    Gravar
End Sub
Private Sub Form_Activate()
    If lFechar Then
        Unload Me
        Exit Sub
    End If
    
    If Not PlanoDePagamentoAberto Then
        Unload Me
        Exit Sub
    End If
    
    TestaInferior TBLPlanoDePagamento, lAllowEdit, lAllowDelete, lAllowConsult
    TestaSuperior TBLPlanoDePagamento, lAllowEdit, lAllowDelete, lAllowConsult
    
    If TBLPlanoDePagamento.RecordCount = 0 Then
        BotãoGravar False
        cmdGravar.Enabled = False
        cmdCancelar.Enabled = False
        BotãoImprimir False
    Else
        BotãoGravar (lInserir Or lAllowEdit)
        cmdGravar.Enabled = (lInserir Or lAllowEdit)
        cmdCancelar.Enabled = (lInserir Or lAllowEdit)
        BotãoImprimir True
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
    BotãoImprimir False
End Sub
Private Sub Form_Load()
    On Error GoTo Erro
    
    ZeraCampos
    
    lAllowInsert = Allow("PLANO DE PAGAMENTO", "I")
    lAllowEdit = Allow("PLANO DE PAGAMENTO", "A")
    lAllowDelete = Allow("PLANO DE PAGAMENTO", "E")
    lAllowConsult = Allow("PLANO DE PAGAMENTO", "C")
    
    lInserir = False
    lAlterar = False
    lPula = False
    
    PlanoDePagamentoAberto = AbreTabela(Dicionário, "FINANCEIRO", "PLANO DE PAGAMENTO", DBFinanceiro, TBLPlanoDePagamento, TBLTabela, dbOpenTable)
    
    If PlanoDePagamentoAberto Then
        IndicePlanoDePagamentoAtivo = "PLANODEPAGAMENTO1"
        TBLPlanoDePagamento.Index = IndicePlanoDePagamentoAtivo
    Else
        MsgBox "Não consegui abrir a tabela 'Plano de Pagamento' !", vbCritical, "Erro"
        Exit Sub
    End If
    
    BotãoIncluir lAllowInsert
 
    If TBLPlanoDePagamento.RecordCount = 0 Then
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
        
    If TBLPlanoDePagamento.RecordCount = 0 Or TBLPlanoDePagamento.RecordCount = 1 Then
        NavegaçãoSuperior False
    Else
        NavegaçãoInferior lAllowConsult
    End If
    
    StatusBarAviso = "Pronto"
    Relatório = AddPath(AplicaçãoPath, "REPORT\PlanoDePagamento.RPT")
    TotalDatabaseName = 1
    DataBaseName(1) = AddPath(AplicaçãoPath, "DATABASE\CADASTRO.MDB")
    lFechar = False
    Exit Sub
    
Erro:
    GeraMensagemDeErro "Plano de Pagamento - Load"
    lFechar = True
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
    
    Set frmPlanoDePagamento = Nothing
End Sub
Private Sub Form_Unload(Cancel As Integer)
    If PlanoDePagamentoAberto Then
        TBLPlanoDePagamento.Close
    End If
    If Forms.Count = 2 Then
        AllBotões False
    End If
End Sub
Private Sub txtCarênciaPosVenda_Change()
    If lPula Then
        Exit Sub
    End If
    FormatMask "99", txtCarênciaPosVenda
End Sub
Private Sub txtCarênciaPosVenda_KeyPress(KeyAscii As Integer)
    If Not lInserir Then
        lAlterar = True
        StatusBarAviso = "Alteração"
        BarraDeStatus StatusBarAviso
    End If
End Sub
Private Sub txtCódigo_Change()
    If lPula Then
        Exit Sub
    End If
    FormatMask "9999", txtCódigo
End Sub
Private Sub txtCódigo_KeyPress(KeyAscii As Integer)
    If Not lInserir Then
        lAlterar = True
        StatusBarAviso = "Alteração"
        BarraDeStatus StatusBarAviso
    End If
End Sub
Private Sub txtCódigo_LostFocus()
    lPula = True
    LeftBlank txtCódigo
    lPula = False
End Sub
Private Sub txtCustoFinanceiro_Change()
    If lPula Then
        Exit Sub
    End If
    FormatMask "@K 99,99", txtCustoFinanceiro
End Sub
Private Sub txtCustoFinanceiro_KeyPress(KeyAscii As Integer)
    If Not lInserir Then
        lAlterar = True
        StatusBarAviso = "Alteração do Orçamento"
        BarraDeStatus StatusBarAviso
    End If
End Sub
Private Sub txtCustoFinanceiro_LostFocus()
    If lPula Then
        Exit Sub
    End If
        
    lPula = True
    FormatMask "@V #0,00", txtCustoFinanceiro
    lPula = False
End Sub
Private Sub txtDescrição_Change()
    If lPula Then
        Exit Sub
    End If
    FormatMask "@!S40", txtDescrição
End Sub
Private Sub txtDescrição_KeyPress(KeyAscii As Integer)
    If Not lInserir Then
        lAlterar = True
        StatusBarAviso = "Alteração"
        BarraDeStatus StatusBarAviso
    End If
End Sub
Private Sub txtIntervaloDePagamentos_Change()
    If lPula Then
        Exit Sub
    End If
    FormatMask "99", txtCarênciaPosVenda
End Sub
Private Sub txtIntervaloDePagamentos_KeyPress(KeyAscii As Integer)
    If Not lInserir Then
        lAlterar = True
        StatusBarAviso = "Alteração"
        BarraDeStatus StatusBarAviso
    End If
End Sub
Private Sub txtQTVencimentos_Change()
    If lPula Then
        Exit Sub
    End If
    FormatMask "99", txtQTVencimentos
End Sub
Private Sub txtQTVencimentos_KeyPress(KeyAscii As Integer)
    If Not lInserir Then
        lAlterar = True
        StatusBarAviso = "Alteração"
        BarraDeStatus StatusBarAviso
    End If
End Sub
Private Sub txtRepasse_Change()
    If lPula Then
        Exit Sub
    End If
    FormatMask "@K 99,99", txtRepasse
End Sub
Private Sub txtRepasse_KeyPress(KeyAscii As Integer)
    If Not lInserir Then
        lAlterar = True
        StatusBarAviso = "Alteração do Orçamento"
        BarraDeStatus StatusBarAviso
    End If
End Sub
Private Sub txtRepasse_LostFocus()
    If lPula Then
        Exit Sub
    End If
        
    lPula = True
    FormatMask "@V #0,00", txtRepasse
    lPula = False
End Sub
