VERSION 5.00
Begin VB.Form frmDespesas 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Despesas"
   ClientHeight    =   2115
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6150
   Icon            =   "Serviços.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   2115
   ScaleWidth      =   6150
   Begin VB.CommandButton cmdGravar 
      Caption         =   "&Gravar"
      Height          =   345
      Left            =   3510
      TabIndex        =   4
      Top             =   1740
      Width           =   1245
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   345
      Left            =   4830
      TabIndex        =   5
      Top             =   1740
      Width           =   1245
   End
   Begin VB.Frame frDespesas 
      Height          =   1695
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   6135
      Begin VB.TextBox txtData 
         Height          =   315
         Left            =   5490
         TabIndex        =   2
         Top             =   750
         Width           =   495
      End
      Begin VB.TextBox txtDescrição 
         Height          =   315
         Left            =   960
         TabIndex        =   3
         Top             =   1170
         Width           =   5055
      End
      Begin VB.ComboBox cmbTipo 
         Height          =   315
         Left            =   960
         TabIndex        =   1
         Top             =   750
         Width           =   3375
      End
      Begin VB.TextBox txtCódigo 
         Height          =   315
         Left            =   960
         TabIndex        =   0
         Top             =   330
         Width           =   465
      End
      Begin VB.Label lblData 
         Caption         =   "Dia"
         Height          =   195
         Left            =   5010
         TabIndex        =   10
         Top             =   780
         Width           =   465
      End
      Begin VB.Label lblDescrição 
         Caption         =   "Descrição"
         Height          =   225
         Left            =   180
         TabIndex        =   9
         Top             =   1170
         Width           =   765
      End
      Begin VB.Label lblTipo 
         Caption         =   "Tipo"
         Height          =   225
         Left            =   180
         TabIndex        =   8
         Top             =   780
         Width           =   495
      End
      Begin VB.Label lblCódigo 
         Caption         =   "Código"
         Height          =   195
         Left            =   180
         TabIndex        =   7
         Top             =   360
         Width           =   615
      End
   End
End
Attribute VB_Name = "frmDespesas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim TBLDespesas As Table
Dim DespesasAberto As Boolean
Dim IndiceDespesasAtivo$

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
    frDespesas.Enabled = True
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
    frDespesas.Enabled = False
    cmdGravar.Enabled = False
    cmdCancelar.Enabled = False
    BotãoGravar False
End Sub
Public Sub Encontrar()
    If Not lAllowConsult Then
        Exit Sub
    End If
    Set frmEncontrar.DBBancoDeDados = DBCadastro
    frmEncontrar.NomeDaJanela = "Despesas"
    frmEncontrar.LabelDescription = "Descrição"
    frmEncontrar.Mensagem = "Nenhuma Despesa foi selecionado!"
    frmEncontrar.BancoDeDados = "CADASTRO"
    frmEncontrar.Tabela = "DESPESAS"
    frmEncontrar.Indice = "1"
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
    
    TBLDespesas.Delete
    
    If Err <> 0 Then
        GeraMensagemDeErro "Despesas - Excluir - " & txtDescrição, True
        StatusBarAviso = "Falha na exclusão"
        BarraDeStatus StatusBarAviso
        Exit Sub
    End If
    
    WS.CommitTrans
    
    Log gUsuário, "Exclusão - Despesas: " & txtCódigo & " - " & txtDescrição
    
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
    If TBLDespesas.RecordCount = 0 Then
        Exit Sub
    End If
    
    TBLDespesas.Seek "=", txtCódigo
    If TBLDespesas.NoMatch Then
        MsgBox "Não consegui encontrar " + txtCódigo, vbExclamation, "Erro"
        TBLDespesas.MoveFirst
        NavegaçãoInferior False
        NavegaçãoInferior lAllowConsult
    Else
        TestaInferior TBLDespesas, lAllowEdit, lAllowDelete, lAllowConsult
        TestaSuperior TBLDespesas, lAllowEdit, lAllowDelete, lAllowConsult
    End If
    GetRecords
End Sub
Public Function PushDataBaseName(ByVal Posição As Integer) As String
    PushDataBaseName = DataBaseName(Posição)
End Function
Private Sub GetRecords()
    lPula = True
    If Not lAllowConsult Then
        ZeraCampos
        DesativaCampos
        lPula = False
        Exit Sub
    End If
    
    txtCódigo = TBLDespesas("CÓDIGO")
    txtDescrição = TBLDespesas("DESCRIÇÃO")
    txtData = TBLDespesas("DATA")
    
    cmbTipo.ListIndex = Val(TBLDespesas("TIPO")) - 1
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
    
    TBLDespesas("CÓDIGO") = txtCódigo
    TBLDespesas("DESCRIÇÃO") = txtDescrição
    TBLDespesas("DATA") = txtData
    TBLDespesas("TIPO") = Trim(Str(cmbTipo.ListIndex + 1))
    If lInserir Then
        TBLDespesas("USERNAME - CRIA") = gUsuário
        TBLDespesas("DATA - CRIA") = Date
        TBLDespesas("HORA - CRIA") = Time
        TBLDespesas("USERNAME - ALTERA") = "VAZIO"
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
        Log gUsuário, "Inclusão - Despesas: " & txtCódigo & " - " & txtDescrição
    Else
        Log gUsuário, "Alteração - Despesas: " & txtCódigo & " - " & txtDescrição
    End If
    
    Exit Function
    
Erro:
    TBLDespesas.CancelUpdate
    GeraMensagemDeErro "Despesas - SetRecords - " & txtDescrição, True
    SetRecords = False
    On Error GoTo 0
End Function
Private Sub ZeraCampos()
    lPula = True
    txtCódigo = Empty
    txtDescrição = Empty
    cmbTipo.ListIndex = 0
    txtData = DataNulaMes
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
    
    cmbTipo.Clear
    
    cmbTipo.AddItem "1-Despesa mensal obrigatória com data fixa"
    cmbTipo.AddItem "2-Despesa mensal obrigatória sem data fixa"
    cmbTipo.AddItem "3-Despesa mensal não obrigatória"
        
    cmbTipo.ListIndex = 0
    
    ZeraCampos
    
    lInserir = False
    lAlterar = False
    lPula = False
    
    DespesasAberto = AbreTabela(Dicionário, "CADASTRO", "DESPESAS", DBCadastro, TBLDespesas, TBLTabela, dbOpenTable)
    
    If DespesasAberto Then
        IndiceDespesasAtivo = "DESPESAS1"
        TBLDespesas.Index = IndiceDespesasAtivo
    Else
        MsgBox "Não consegui abrir a tabela 'Despesas' !", vbCritical, "Erro"
        Exit Sub
    End If
    
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
    Relatório = AddPath(AplicaçãoPath, "REPORT\Despesas.RPT")
    TotalDatabaseName = 1
    DataBaseName(1) = AddPath(AplicaçãoPath, "DATABASE\CADASTRO.MDB")
    
    lFechar = False
    
    Exit Sub
    
Erro:
    GeraMensagemDeErro "Despesas - Load"
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
    
    Set frmDespesas = Nothing
End Sub
Private Sub Form_Unload(Cancel As Integer)
    If DespesasAberto Then
        TBLDespesas.Close
    End If
    If Forms.Count = 2 Then
        AllBotões False
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
Private Sub txtData_Change()
    If Not lPula Then
        lPula = True
        FormatMask DataMaskMes, txtData
        lPula = False
    End If
End Sub
Private Sub txtData_KeyPress(KeyAscii As Integer)
    If Not lInserir Then
        lAlterar = True
        StatusBarAviso = "Alteração"
        BarraDeStatus StatusBarAviso
    End If
End Sub
Private Sub txtData_LostFocus()
    If txtData <> Space(2) Then
        lPula = True
        'CorrigeData DataMaskMes, txtData, Date
        lPula = False
        If Not FormatMask(CheckDataMaskMes, txtData) Then
            Beep
            MsgBox "Data inválida !", vbCritical, "Erro"
            txtData.SelStart = 0
            txtData.SetFocus
        End If
    End If
End Sub
Private Sub txtDescrição_Change()
    If lPula Then
        Exit Sub
    End If
    FormatMask "@!S30", txtDescrição
End Sub
Private Sub txtDescrição_KeyPress(KeyAscii As Integer)
    If Not lInserir Then
        lAlterar = True
        StatusBarAviso = "Alteração"
        BarraDeStatus StatusBarAviso
    End If
End Sub

