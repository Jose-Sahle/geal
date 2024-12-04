VERSION 5.00
Begin VB.Form frmApontamentoDeDespesas 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Apontamentos de Serviços"
   ClientHeight    =   3495
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5475
   Icon            =   "ApontamentoDeServiços.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   3495
   ScaleWidth      =   5475
   Begin VB.CommandButton cmdGravar 
      Caption         =   "&Gravar"
      Height          =   345
      Left            =   2880
      TabIndex        =   4
      Top             =   3120
      Width           =   1245
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   345
      Left            =   4200
      TabIndex        =   5
      Top             =   3120
      Width           =   1245
   End
   Begin VB.Frame frApontamentosDeDespesas 
      Height          =   3045
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   5475
      Begin VB.TextBox txtObservação 
         Height          =   1035
         Left            =   150
         MultiLine       =   -1  'True
         TabIndex        =   2
         Top             =   1320
         Width           =   5145
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
         Height          =   315
         Left            =   3630
         TabIndex        =   3
         Text            =   "99.999.999,99"
         Top             =   2580
         Width           =   1665
      End
      Begin VB.TextBox txtVencimento 
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
         Left            =   1050
         TabIndex        =   0
         Text            =   "99/99/9999"
         Top             =   240
         Width           =   1335
      End
      Begin VB.ComboBox cmbDespesa 
         Height          =   315
         Left            =   1050
         TabIndex        =   1
         Top             =   660
         Width           =   4245
      End
      Begin VB.Label lblObservação 
         Caption         =   "Observação"
         Height          =   225
         Left            =   150
         TabIndex        =   10
         Top             =   1080
         Width           =   1305
      End
      Begin VB.Label lblValor 
         Caption         =   "Valor"
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
         Left            =   3000
         TabIndex        =   9
         Top             =   2640
         Width           =   525
      End
      Begin VB.Label lblVencimento 
         Caption         =   "Vencimento"
         Height          =   255
         Left            =   150
         TabIndex        =   8
         Top             =   300
         Width           =   945
      End
      Begin VB.Label lblServiços 
         Caption         =   "Serviço"
         Height          =   225
         Left            =   150
         TabIndex        =   7
         Top             =   690
         Width           =   645
      End
   End
End
Attribute VB_Name = "frmApontamentoDeDespesas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim TBLDespesas          As Table
Dim DespesasAberto       As Boolean
Dim IndiceDespesasAtivo  As String

Dim TBLTipoDeDespesas          As Table
Dim TipoDeDespesasAberto       As Boolean
Dim IndiceTipoDeDespesasAtivo  As String

Dim lPula As Boolean

Dim lInserir As Boolean
Dim lAlterar As Boolean

Dim lFechar As Boolean

Dim lAllowInsert  As Boolean
Dim lAllowEdit As Boolean
Dim lAllowDelete As Boolean
Dim lAllowConsult As Boolean

Dim CódigoDoDespesa() As Integer

Dim mUsuário As String
Dim mHora As Date

Dim StatusBarAviso$

Public lAtualizar As Boolean
Private Function Ascan(ByRef Matriz() As Integer, ByVal Expressão As String) As Integer
    Dim Cont As Integer
    Dim Retorno As Integer
    
    Retorno = -1
    
    For Cont = LBound(Matriz) To UBound(Matriz)
        If Matriz(Cont) = Expressão Then
            Retorno = Cont
            Exit For
        End If
    Next
    
    Ascan = Retorno
End Function
Private Sub AtivaCampos()
    frApontamentosDeDespesas.Enabled = True
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
    
    If TBLDespesas.RecordCount = 0 Then
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
    
    TestaInferior TBLDespesas, lAllowEdit, lAllowDelete, lAllowConsult
    TestaSuperior TBLDespesas, lAllowEdit, lAllowDelete, lAllowConsult
    
    GetRecords
End Function
Private Sub DesativaCampos()
    frApontamentosDeDespesas.Enabled = False
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
    
    TBLDespesas.Delete
    
    If Err <> 0 Then
        GeraMensagemDeErro "Despesas - Excluir - " & txtVencimento & " - " & txtObservação, True
        StatusBarAviso = "Falha na exclusão"
        BarraDeStatus StatusBarAviso
        Exit Sub
    End If
    
    WS.CommitTrans
    
    Log gUsuário, "Exclusão - Apontamento de Despesas: " & txtVencimento & " - " & txtObservação
    
    StatusBarAviso = "Exclusão bem sucedida"
    BarraDeStatus StatusBarAviso
    
    If TBLDespesas.RecordCount = 0 Then
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
    
    If TBLDespesas.BOF Then
        TBLDespesas.MoveFirst
    ElseIf TBLDespesas.EOF Then
        TBLDespesas.MoveLast
    Else
        TBLDespesas.MovePrevious
        If TBLDespesas.BOF Then
            TBLDespesas.MoveNext
        End If
    End If
    
    GetRecords
    
    TestaInferior TBLDespesas, lAllowEdit, lAllowDelete, lAllowConsult
    TestaSuperior TBLDespesas, lAllowEdit, lAllowDelete, lAllowConsult
End Sub
Public Sub FillDespesas()
    Dim Cont As Integer
    
    TBLTipoDeDespesas.MoveFirst
    
    cmbDespesa.Clear
    
    ReDim CódigoDoDespesa(0 To TBLTipoDeDespesas.RecordCount - 1)
    
    For Cont = 0 To TBLTipoDeDespesas.RecordCount - 1
        cmbDespesa.AddItem TBLTipoDeDespesas("DESCRIÇÃO")
        CódigoDoDespesa(Cont) = TBLTipoDeDespesas("CÓDIGO")
        TBLTipoDeDespesas.MoveNext
    Next
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
        If TBLDespesas.RecordCount > 0 And Not TBLDespesas.BOF And Not TBLDespesas.EOF Then
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
    
    TestaInferior TBLDespesas, lAllowEdit, lAllowDelete, lAllowConsult
    TestaSuperior TBLDespesas, lAllowEdit, lAllowDelete, lAllowConsult
    
    If TBLDespesas.RecordCount = 0 Then
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
    
    If txtVencimento.Enabled Then
        txtVencimento.SetFocus
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
    
    txtVencimento.SetFocus
End Sub
Public Sub MoveFirst()
    If lAlterar Then
       If Not Cancelamento Then
            Exit Sub
        End If
    End If
    
    StatusBarAviso = "Pronto"
    BarraDeStatus StatusBarAviso
    
    TBLDespesas.MoveFirst
    
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
    
    TBLDespesas.MoveLast
    
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
    
    TBLDespesas.MoveNext
    If TBLDespesas.EOF Then
        TBLDespesas.MovePrevious
        Exit Sub
    End If
    
    StatusBarAviso = "Pronto"
    BarraDeStatus StatusBarAviso
    
    NavegaçãoInferior lAllowConsult
    TestaSuperior TBLDespesas, lAllowEdit, lAllowDelete, lAllowConsult
    
    GetRecords
End Sub
Public Sub MovePrevious()
    If lAlterar Then
       If Not Cancelamento Then
            Exit Sub
        End If
    End If
    
    TBLDespesas.MovePrevious
    If TBLDespesas.BOF Then
        TBLDespesas.MoveNext
        Exit Sub
    End If
    
    StatusBarAviso = "Pronto"
    BarraDeStatus StatusBarAviso
    
    NavegaçãoSuperior lAllowConsult
    TestaInferior TBLDespesas, lAllowEdit, lAllowDelete, lAllowConsult
    
    GetRecords
End Sub
Public Sub PosRecords()
    Dim Código As Integer
    
    If TBLDespesas.RecordCount = 0 Then
        Exit Sub
    End If
    
    Código = CódigoDoDespesa(cmbDespesa.ListIndex)
    
    TBLDespesas.Seek "=", Código, txtVencimento, mHora, mUsuário
    
    If TBLDespesas.NoMatch Then
        MsgBox "Não consegui encontrar o Despesa", vbExclamation, "Erro"
        TBLDespesas.MoveFirst
        NavegaçãoInferior False
        NavegaçãoInferior lAllowConsult
    Else
        TestaInferior TBLDespesas, lAllowEdit, lAllowDelete, lAllowConsult
        TestaSuperior TBLDespesas, lAllowEdit, lAllowDelete, lAllowConsult
    End If
    
    GetRecords
End Sub
Private Sub GetRecords()
    Dim Pos As Integer
    
    lPula = True
    
    If Not lAllowConsult Then
        ZeraCampos
        DesativaCampos
        lPula = False
        Exit Sub
    End If
    
    If TBLDespesas("DATA DO VENCIMENTO") <> vbNull Then
        txtVencimento = FormatStringMask(CheckDataMask, TBLDespesas("DATA DO VENCIMENTO"))
        CorrigeData DataMask, txtVencimento, TBLDespesas("DATA DO VENCIMENTO")
    Else
        txtVencimento = DataNula
    End If
    
    Pos = Ascan(CódigoDoDespesa, TBLDespesas("CÓDIGO DA DESPESA"))
    cmbDespesa.ListIndex = Pos
    
    txtObservação = TBLDespesas("OBSERVAÇÃO")
    
    txtValor = TBLDespesas("VALOR DA DESPESA")
    lPula = True
    txtValor_LostFocus
    lPula = False
    
    mHora = TBLDespesas("HORA")
    mUsuário = TBLDespesas("USERNAME - CRIA")
    
    lPula = False
    
    If Not lAllowEdit Then
        DesativaCampos
    End If
End Sub
Private Function SetRecords()
    On Error GoTo Erro
    
    Dim Msg$
    Dim Confirmação As Integer, Msg1$, Msg2$
    
    WS.BeginTrans 'Inicia uma Transação
    
    If lInserir Then
        TBLDespesas.AddNew
    Else
        TBLDespesas.Edit
    End If
    
    TBLDespesas("DATA DO VENCIMENTO") = IIf(Trim(StrTran(txtVencimento, "/")) <> Empty, txtVencimento, vbNull)
    TBLDespesas("CÓDIGO DA DESPESA") = CódigoDoDespesa(cmbDespesa.ListIndex)
    TBLDespesas("OBSERVAÇÃO") = txtObservação
    TBLDespesas("VALOR DA DESPESA") = ValStr(txtValor)
    
    If lInserir Then
        mHora = Time
        mUsuário = gUsuário
        TBLDespesas("HORA") = mHora
        TBLDespesas("USERNAME - CRIA") = gUsuário
        TBLDespesas("DATA - CRIA") = Date
        TBLDespesas("HORA - CRIA") = Time
        TBLDespesas("USERNAME - ALTERA") = vbNull
        TBLDespesas("DATA - ALTERA") = vbNull
        TBLDespesas("HORA - ALTERA") = vbNull
    End If
    If lAlterar Then
        TBLDespesas("USERNAME - ALTERA") = gUsuário
        TBLDespesas("DATA - ALTERA") = Date
        TBLDespesas("HORA - ALTERA") = Time
    End If
    TBLDespesas.Update
            
    WS.CommitTrans 'Grava as alterações ou inclusões se não houverem erros
    
    SetRecords = True
    
    If lInserir Then
        Log gUsuário, "Inclusão - Apontamento de Despesas: " & cmbDespesa.Text & " - " & txtVencimento
    Else
        Log gUsuário, "Alteração - Apontamento de Despesas: " & cmbDespesa.Text & " - " & txtVencimento
    End If
    
    Exit Function
    
Erro:
    TBLDespesas.CancelUpdate
    GeraMensagemDeErro "Apontamento de Despesas - SetRecords - " & cmbDespesa.Text & " - " & txtVencimento, True
    SetRecords = False
    On Error GoTo 0
End Function
Private Sub ZeraCampos()
    lPula = True
    txtVencimento = "  /  /  "
    txtVencimento_LostFocus
    lPula = True
    cmbDespesa.ListIndex = 0
    txtObservação = Empty
    txtValor = Empty
    txtValor_LostFocus
    lPula = False
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
    
    If Not DespesasAberto Then
        Unload Me
        Exit Sub
    End If
    
    If Not TipoDeDespesasAberto Then
        Unload Me
        Exit Sub
    End If
    
    TestaInferior TBLDespesas, lAllowEdit, lAllowDelete, lAllowConsult
    TestaSuperior TBLDespesas, lAllowEdit, lAllowDelete, lAllowConsult
    
    If TBLDespesas.RecordCount = 0 Then
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
    
    lAllowInsert = Allow("DESPESAS", "I")
    lAllowEdit = Allow("DESPESAS", "A")
    lAllowDelete = Allow("DESPESAS", "E")
    lAllowConsult = Allow("DESPESAS", "C")
        
    lInserir = False
    lAlterar = False
    lPula = False
    
    TipoDeDespesasAberto = AbreTabela(Dicionário, "CADASTRO", "DESPESAS", DBCadastro, TBLTipoDeDespesas, TBLTabela, dbOpenTable)
    
    If TipoDeDespesasAberto Then
        IndiceTipoDeDespesasAtivo = "DESPESAS1"
        TBLTipoDeDespesas.Index = IndiceTipoDeDespesasAtivo
    Else
        MsgBox "Não consegui abrir a tabela 'Cadastro - Despesas' !", vbCritical, "Erro"
        Exit Sub
    End If
    
    DespesasAberto = AbreTabela(Dicionário, "FINANCEIRO", "DESPESA", DBFinanceiro, TBLDespesas, TBLTabela, dbOpenTable)
    
    If DespesasAberto Then
        IndiceDespesasAtivo = "DESPESA1"
        TBLDespesas.Index = IndiceDespesasAtivo
    Else
        MsgBox "Não consegui abrir a tabela 'Financeiro - Despesas' !", vbCritical, "Erro"
        Exit Sub
    End If
    
    FillDespesas
    
    ZeraCampos
    
    BotãoIncluir lAllowInsert
 
    If TBLDespesas.RecordCount = 0 Then
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
        
    If TBLDespesas.RecordCount = 0 Or TBLDespesas.RecordCount = 1 Then
        NavegaçãoSuperior False
    Else
        NavegaçãoInferior lAllowConsult
    End If
    
    StatusBarAviso = "Pronto"
    
    lFechar = False
    
    Exit Sub
    
Erro:
    GeraMensagemDeErro "Apontamento de Despesas - Load"
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
    
    Set frmApontamentoDeDespesas = Nothing
End Sub
Private Sub Form_Unload(Cancel As Integer)
    If DespesasAberto Then
        TBLDespesas.Close
    End If
    
    If TipoDeDespesasAberto Then
        TBLTipoDeDespesas.Close
    End If
    
    If Forms.Count = 2 Then
        AllBotões False
    End If
End Sub
Private Sub txtObservação_KeyPress(KeyAscii As Integer)
    If Not lInserir Then
        lAlterar = True
        StatusBarAviso = "Alteração"
        BarraDeStatus StatusBarAviso
    End If
End Sub
Private Sub txtValor_Change()
    If Not lPula Then
        FormatMask "@K 99.999.999,99", txtValor
    End If
End Sub
Private Sub txtValor_LostFocus()
    lPula = True
    FormatMask "@V ##.###.##0,00", txtValor
    lPula = False
End Sub
Private Sub txtVencimento_Change()
    If Not lPula Then
        lPula = True
        FormatMask DataMask, txtVencimento
        lPula = False
    End If
End Sub
Private Sub txtVencimento_KeyPress(KeyAscii As Integer)
    If Not lInserir Then
        lAlterar = True
        StatusBarAviso = "Alteração"
        BarraDeStatus StatusBarAviso
    End If
End Sub
Private Sub txtVencimento_LostFocus()
    If StrTran(txtVencimento.Text, "/") <> Space(6) Then
        lPula = True
        CorrigeData DataMask, txtVencimento, Date
        lPula = False
        If Not FormatMask(CheckDataMask, txtVencimento) Then
            Beep
            MsgBox "Data inválida !", vbCritical, "Erro"
            txtVencimento.SelStart = 0
            txtVencimento.SetFocus
        End If
    End If
End Sub

