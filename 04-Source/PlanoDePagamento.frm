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
         Caption         =   "Op��o no caixa"
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
      Begin VB.CheckBox chkPermiteAlterarAutoInclus�o 
         Alignment       =   1  'Right Justify
         Caption         =   "Permite Alterar Auto-Inclus�o"
         Height          =   375
         Left            =   210
         TabIndex        =   5
         Top             =   2310
         Width           =   2205
      End
      Begin VB.CheckBox chkAutoInclus�o 
         Alignment       =   1  'Right Justify
         Caption         =   "Auto-Inclus�o"
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
      Begin VB.TextBox txtCar�nciaPosVenda 
         Height          =   285
         Left            =   1950
         TabIndex        =   2
         Top             =   1200
         Width           =   465
      End
      Begin VB.TextBox txtDescri��o 
         Height          =   285
         Left            =   1320
         TabIndex        =   1
         Top             =   660
         Width           =   4605
      End
      Begin VB.TextBox txtC�digo 
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
      Begin VB.Label lblCar�nciaPosVenda 
         Caption         =   "Car�ncia"
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
      Begin VB.Label lblDescri��o 
         Caption         =   "Descri��o"
         Height          =   195
         Left            =   270
         TabIndex        =   17
         Top             =   690
         Width           =   765
      End
      Begin VB.Label lblC�digo 
         Caption         =   "C�digo"
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
Public Relat�rio$
Public TotalDatabaseName%

Public lAtualizar As Boolean
Private Sub AtivaCampos()
    frPlanoDePagamento.Enabled = True
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
    
    lInserir = False
    lAlterar = False
    Bot�oIncluir lAllowInsert
    
    If TBLPlanoDePagamento.RecordCount = 0 Then
        Navega��oInferior False
        Navega��oSuperior False
        Bot�oGravar False
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
    Bot�oGravar False
End Sub
Public Sub Encontrar()
    If Not lAllowConsult Then
        Exit Sub
    End If
    Set frmEncontrar.DBBancoDeDados = DBFinanceiro
    frmEncontrar.NomeDaJanela = "Plano de Pagamento"
    frmEncontrar.LabelDescription = "Descri��o"
    frmEncontrar.Mensagem = "Nenhum Plano de Pagamento foi selecionado!"
    frmEncontrar.BancoDeDados = "FINANCEIRO"
    frmEncontrar.Tabela = "PLANO DE PAGAMENTO"
    frmEncontrar.Indice = "2"
    frmEncontrar.CampoChave = "C�DIGO"
    frmEncontrar.CampoPreencheLista = "DESCRI��O"
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
    
    TBLPlanoDePagamento.Delete
    
    If Err <> 0 Then
        GeraMensagemDeErro "Plano de Pagamento - Excluir - " & txtDescri��o
        WS.Rollback 'Caso haja erro volta os valores normais
        StatusBarAviso = "Falha na exclus�o"
        BarraDeStatus StatusBarAviso
        Exit Sub
    End If
    
    WS.CommitTrans
    
    Log gUsu�rio, "Exclus�o - Plano de Pagamento: " & txtC�digo & " - " & txtDescri��o
    
    StatusBarAviso = "Exclus�o bem sucedida"
    BarraDeStatus StatusBarAviso
    
    If TBLPlanoDePagamento.RecordCount = 0 Then
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
    txtC�digo = TBLPlanoDePagamento("C�DIGO")
    txtDescri��o = TBLPlanoDePagamento("DESCRI��O")
    txtCar�nciaPosVenda = TBLPlanoDePagamento("CAR�NCIA")
    txtIntervaloDePagamentos = TBLPlanoDePagamento("INTERVALO DE PAGAMENTOS")
    txtQTVencimentos = TBLPlanoDePagamento("QUANTIDADE DE VENCIMENTOS")
    txtCustoFinanceiro = TBLPlanoDePagamento("CUSTO FINANCEIRO")
    chkAutoInclus�o.Value = IIf(TBLPlanoDePagamento("AUTO-INCLUS�O"), 1, 0)
    chkPermiteAlterarAutoInclus�o.Value = IIf(TBLPlanoDePagamento("PERMITE ALTERAR AUTO-INCLUS�O"), 1, 0)
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
            StatusBarAviso = "Inclus�o bem sucedida"
        Else
            StatusBarAviso = "Falha na inclus�o"
        End If
    Else
        If TBLPlanoDePagamento.RecordCount > 0 And Not TBLPlanoDePagamento.BOF And Not TBLPlanoDePagamento.EOF Then
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
    
    TestaInferior TBLPlanoDePagamento, lAllowEdit, lAllowDelete, lAllowConsult
    TestaSuperior TBLPlanoDePagamento, lAllowEdit, lAllowDelete, lAllowConsult
    
    If TBLPlanoDePagamento.RecordCount = 0 Then
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
    
    If txtC�digo.Enabled Then
        txtC�digo.SetFocus
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
    
    txtC�digo.SetFocus
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
    
    TBLPlanoDePagamento.MoveLast
    
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
    
    TBLPlanoDePagamento.MoveNext
    If TBLPlanoDePagamento.EOF Then
        TBLPlanoDePagamento.MovePrevious
        Exit Sub
    End If
    
    StatusBarAviso = "Pronto"
    BarraDeStatus StatusBarAviso
    
    Navega��oInferior lAllowConsult
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
    
    Navega��oSuperior lAllowConsult
    TestaInferior TBLPlanoDePagamento, lAllowEdit, lAllowDelete, lAllowConsult
    
    GetRecords
End Sub
Public Sub PosRecords()
    If TBLPlanoDePagamento.RecordCount = 0 Then
        Exit Sub
    End If
    
    TBLPlanoDePagamento.Seek "=", txtC�digo
    If TBLPlanoDePagamento.NoMatch Then
        MsgBox "N�o consegui encontrar " + txtC�digo, vbExclamation, "Erro"
        TBLPlanoDePagamento.MoveFirst
        Navega��oInferior False
        Navega��oInferior lAllowConsult
    Else
        TestaInferior TBLPlanoDePagamento, lAllowEdit, lAllowDelete, lAllowConsult
        TestaSuperior TBLPlanoDePagamento, lAllowEdit, lAllowDelete, lAllowConsult
    End If
    GetRecords
End Sub
Public Function PushDataBaseName(ByVal Posi��o As Integer) As String
    PushDataBaseName = DataBaseName(Posi��o)
End Function
Private Function SetRecords()
    On Error GoTo Erro
    
    Dim Msg$
    Dim Confirma��o As Integer, Msg1$, Msg2$
    
    WS.BeginTrans 'Inicia uma Transa��o
    
    If lInserir Then
        TBLPlanoDePagamento.AddNew
    Else
        TBLPlanoDePagamento.Edit
    End If
    
    TBLPlanoDePagamento("C�DIGO") = txtC�digo
    TBLPlanoDePagamento("DESCRI��O") = txtDescri��o
    TBLPlanoDePagamento("CAR�NCIA") = txtCar�nciaPosVenda
    TBLPlanoDePagamento("INTERVALO DE PAGAMENTOS") = txtIntervaloDePagamentos
    TBLPlanoDePagamento("QUANTIDADE DE VENCIMENTOS") = txtQTVencimentos
    TBLPlanoDePagamento("CUSTO FINANCEIRO") = txtCustoFinanceiro
    TBLPlanoDePagamento("REPASSE") = txtRepasse
    TBLPlanoDePagamento("AUTO-INCLUS�O") = chkAutoInclus�o.Value
    TBLPlanoDePagamento("PERMITE ALTERAR AUTO-INCLUS�O") = chkPermiteAlterarAutoInclus�o.Value
    TBLPlanoDePagamento("PRIMEIRO PAGAMENTO BAIXADO") = chkPrimeiroPagamentoBaixado.Value
    TBLPlanoDePagamento("CAIXA") = chkCaixa.Value
    TBLPlanoDePagamento("COMPRA") = chkCompra.Value
    TBLPlanoDePagamento("VENDA") = chkVenda.Value
    If lInserir Then
        TBLPlanoDePagamento("USERNAME - CRIA") = gUsu�rio
        TBLPlanoDePagamento("DATA - CRIA") = Date
        TBLPlanoDePagamento("HORA - CRIA") = Time
        TBLPlanoDePagamento("USERNAME - ALTERA") = "VAZIO"
        TBLPlanoDePagamento("DATA - ALTERA") = vbNull
        TBLPlanoDePagamento("HORA - ALTERA") = vbNull
    End If
    If lAlterar Then
        TBLPlanoDePagamento("USERNAME - ALTERA") = gUsu�rio
        TBLPlanoDePagamento("DATA - ALTERA") = Date
        TBLPlanoDePagamento("HORA - ALTERA") = Time
    End If
    TBLPlanoDePagamento.Update
            
    WS.CommitTrans 'Grava as altera��es ou inclus�es se n�o houver erros
    
    If lInserir Then
        Log gUsu�rio, "Inclus�o - Plano de Pagamento: " & txtC�digo & " - " & txtDescri��o
    Else
        Log gUsu�rio, "Altera��o - Plano de Pagamento: " & txtC�digo & " - " & txtDescri��o
    End If
    
    SetRecords = True
    
    Exit Function
    
Erro:
    GeraMensagemDeErro "Plano de Pagamento - SetRecords - " & txtDescri��o
    On Error Resume Next
    SetRecords = False
    TBLPlanoDePagamento.CancelUpdate
    WS.Rollback  'Caso haja erro volta os valores normais
    On Error GoTo 0
End Function
Private Sub ZeraCampos()
    lPula = True
    txtC�digo = Empty
    txtDescri��o = Empty
    txtCar�nciaPosVenda = "0"
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
Private Sub chkAutoInclus�o_Click()
    If Not lInserir And Not lPula Then
        lAlterar = True
        StatusBarAviso = "Altera��o"
        BarraDeStatus StatusBarAviso
    End If
End Sub
Private Sub chkAutoInclus�o_KeyPress(KeyAscii As Integer)
    If Not lInserir Then
        lAlterar = True
        StatusBarAviso = "Altera��o"
        BarraDeStatus StatusBarAviso
    End If
End Sub
Private Sub chkCaixa_Click()
    If Not lInserir And Not lPula Then
        lAlterar = True
        StatusBarAviso = "Altera��o"
        BarraDeStatus StatusBarAviso
    End If
End Sub
Private Sub chkCaixa_KeyPress(KeyAscii As Integer)
    If Not lInserir Then
        lAlterar = True
        StatusBarAviso = "Altera��o"
        BarraDeStatus StatusBarAviso
    End If
End Sub
Private Sub chkCompra_Click()
    If Not lInserir And Not lPula Then
        lAlterar = True
        StatusBarAviso = "Altera��o"
        BarraDeStatus StatusBarAviso
    End If
End Sub
Private Sub chkCompra_KeyPress(KeyAscii As Integer)
    If Not lInserir Then
        lAlterar = True
        StatusBarAviso = "Altera��o"
        BarraDeStatus StatusBarAviso
    End If
End Sub
Private Sub chkPermiteAlterarAutoInclus�o_Click()
    If Not lInserir And Not lPula Then
        lAlterar = True
        StatusBarAviso = "Altera��o"
        BarraDeStatus StatusBarAviso
    End If
End Sub
Private Sub chkPermiteAlterarAutoInclus�o_KeyPress(KeyAscii As Integer)
    If Not lInserir Then
        lAlterar = True
        StatusBarAviso = "Altera��o"
        BarraDeStatus StatusBarAviso
    End If
End Sub
Private Sub chkPrimeiroPagamentoBaixado_Click()
    If Not lInserir And Not lPula Then
        lAlterar = True
        StatusBarAviso = "Altera��o"
        BarraDeStatus StatusBarAviso
    End If
End Sub
Private Sub chkPrimeiroPagamentoBaixado_KeyPress(KeyAscii As Integer)
    If Not lInserir Then
        lAlterar = True
        StatusBarAviso = "Altera��o"
        BarraDeStatus StatusBarAviso
    End If
End Sub
Private Sub chkVenda_Click()
    If Not lInserir And Not lPula Then
        lAlterar = True
        StatusBarAviso = "Altera��o"
        BarraDeStatus StatusBarAviso
    End If
End Sub
Private Sub chkVenda_KeyPress(KeyAscii As Integer)
    If Not lInserir Then
        lAlterar = True
        StatusBarAviso = "Altera��o"
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
        Bot�oGravar False
        cmdGravar.Enabled = False
        cmdCancelar.Enabled = False
        Bot�oImprimir False
    Else
        Bot�oGravar (lInserir Or lAllowEdit)
        cmdGravar.Enabled = (lInserir Or lAllowEdit)
        cmdCancelar.Enabled = (lInserir Or lAllowEdit)
        Bot�oImprimir True
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
    
    BarraDeStatus StatusBarAviso
End Sub
Private Sub Form_Deactivate()
    cmdGravar.Enabled = False
    cmdCancelar.Enabled = False
    Bot�oImprimir False
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
    
    PlanoDePagamentoAberto = AbreTabela(Dicion�rio, "FINANCEIRO", "PLANO DE PAGAMENTO", DBFinanceiro, TBLPlanoDePagamento, TBLTabela, dbOpenTable)
    
    If PlanoDePagamentoAberto Then
        IndicePlanoDePagamentoAtivo = "PLANODEPAGAMENTO1"
        TBLPlanoDePagamento.Index = IndicePlanoDePagamentoAtivo
    Else
        MsgBox "N�o consegui abrir a tabela 'Plano de Pagamento' !", vbCritical, "Erro"
        Exit Sub
    End If
    
    Bot�oIncluir lAllowInsert
 
    If TBLPlanoDePagamento.RecordCount = 0 Then
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
        
    If TBLPlanoDePagamento.RecordCount = 0 Or TBLPlanoDePagamento.RecordCount = 1 Then
        Navega��oSuperior False
    Else
        Navega��oInferior lAllowConsult
    End If
    
    StatusBarAviso = "Pronto"
    Relat�rio = AddPath(Aplica��oPath, "REPORT\PlanoDePagamento.RPT")
    TotalDatabaseName = 1
    DataBaseName(1) = AddPath(Aplica��oPath, "DATABASE\CADASTRO.MDB")
    lFechar = False
    Exit Sub
    
Erro:
    GeraMensagemDeErro "Plano de Pagamento - Load"
    lFechar = True
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
    
    Set frmPlanoDePagamento = Nothing
End Sub
Private Sub Form_Unload(Cancel As Integer)
    If PlanoDePagamentoAberto Then
        TBLPlanoDePagamento.Close
    End If
    If Forms.Count = 2 Then
        AllBot�es False
    End If
End Sub
Private Sub txtCar�nciaPosVenda_Change()
    If lPula Then
        Exit Sub
    End If
    FormatMask "99", txtCar�nciaPosVenda
End Sub
Private Sub txtCar�nciaPosVenda_KeyPress(KeyAscii As Integer)
    If Not lInserir Then
        lAlterar = True
        StatusBarAviso = "Altera��o"
        BarraDeStatus StatusBarAviso
    End If
End Sub
Private Sub txtC�digo_Change()
    If lPula Then
        Exit Sub
    End If
    FormatMask "9999", txtC�digo
End Sub
Private Sub txtC�digo_KeyPress(KeyAscii As Integer)
    If Not lInserir Then
        lAlterar = True
        StatusBarAviso = "Altera��o"
        BarraDeStatus StatusBarAviso
    End If
End Sub
Private Sub txtC�digo_LostFocus()
    lPula = True
    LeftBlank txtC�digo
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
        StatusBarAviso = "Altera��o do Or�amento"
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
Private Sub txtDescri��o_Change()
    If lPula Then
        Exit Sub
    End If
    FormatMask "@!S40", txtDescri��o
End Sub
Private Sub txtDescri��o_KeyPress(KeyAscii As Integer)
    If Not lInserir Then
        lAlterar = True
        StatusBarAviso = "Altera��o"
        BarraDeStatus StatusBarAviso
    End If
End Sub
Private Sub txtIntervaloDePagamentos_Change()
    If lPula Then
        Exit Sub
    End If
    FormatMask "99", txtCar�nciaPosVenda
End Sub
Private Sub txtIntervaloDePagamentos_KeyPress(KeyAscii As Integer)
    If Not lInserir Then
        lAlterar = True
        StatusBarAviso = "Altera��o"
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
        StatusBarAviso = "Altera��o"
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
        StatusBarAviso = "Altera��o do Or�amento"
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
