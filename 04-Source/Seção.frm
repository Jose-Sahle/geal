VERSION 5.00
Begin VB.Form frmSeção 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Seção"
   ClientHeight    =   1695
   ClientLeft      =   1140
   ClientTop       =   1515
   ClientWidth     =   6480
   Icon            =   "Seção.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   1695
   ScaleWidth      =   6480
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   345
      Left            =   5205
      TabIndex        =   3
      Top             =   1320
      Width           =   1245
   End
   Begin VB.CommandButton cmdGravar 
      Caption         =   "&Gravar"
      Height          =   345
      Left            =   3885
      TabIndex        =   2
      Top             =   1320
      Width           =   1245
   End
   Begin VB.Frame frSeção 
      Height          =   1275
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   6465
      Begin VB.TextBox txtDescrição 
         Height          =   285
         Left            =   1200
         TabIndex        =   1
         Top             =   750
         Width           =   5000
      End
      Begin VB.TextBox txtCódigo 
         Height          =   285
         Left            =   1200
         TabIndex        =   0
         Top             =   300
         Width           =   750
      End
      Begin VB.Label lblDescrição 
         Caption         =   "Descrição"
         Height          =   195
         Left            =   150
         TabIndex        =   6
         Top             =   780
         Width           =   885
      End
      Begin VB.Label lblCódigo 
         Caption         =   "Código"
         Height          =   200
         Left            =   150
         TabIndex        =   5
         Top             =   330
         Width           =   855
      End
   End
End
Attribute VB_Name = "frmSeção"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim TBLSeção As Table
Dim SeçãoAberto As Boolean
Dim IndiceSeçãoAtivo$
Dim txtCódigoAnterior As String

Dim lAllowInsert  As Boolean
Dim lAllowEdit    As Boolean
Dim lAllowDelete  As Boolean
Dim lAllowConsult As Boolean

Dim lPula As Boolean
Dim lInserir As Boolean
Dim lAlterar As Boolean
Dim mFechar As Boolean

Dim StatusBarAviso$

Dim DataBaseName(1 To 1) As String
Public Relatório$
Public TotalDatabaseName%

Public lAtualizar As Boolean
Private Sub AtivaCampos()
    BotãoImprimir True
    frSeção.Enabled = True
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
    
    If TBLSeção.RecordCount = 0 Then
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
    
    TestaInferior TBLSeção, lAllowEdit, lAllowDelete, lAllowConsult
    TestaSuperior TBLSeção, lAllowEdit, lAllowDelete, lAllowConsult
    
    GetRecords
End Function
Private Sub DesativaCampos()
    BotãoImprimir False
    frSeção.Enabled = False
    cmdGravar.Enabled = False
    cmdCancelar.Enabled = False
    BotãoGravar False
End Sub
Public Sub Seção()
    Set frmEncontrar.DBBancoDeDados = DBCadastro
    frmEncontrar.NomeDaJanela = "Seção"
    frmEncontrar.LabelDescription = "Descrição"
    frmEncontrar.Mensagem = "Nenhuma seção foi selecionada!"
    frmEncontrar.BancoDeDados = "CADASTRO"
    frmEncontrar.Tabela = "SEÇÃO"
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
    Dim TBLDepartamentoSeção As Table

    If lAlterar Then
       If Not Cancelamento Then
           Exit Sub
       End If
    End If

    If AbreTabela(Dicionário, "CADASTRO", "DEPARTAMENTO - SEÇÃO", DBCadastro, TBLDepartamentoSeção, TBLTabela, dbOpenTable) Then
        TBLDepartamentoSeção.Index = "DEPARTAMENTOSEÇÃO2"
        TBLDepartamentoSeção.Seek ">=", txtCódigo
        If Not TBLDepartamentoSeção.NoMatch Then
            If TBLDepartamentoSeção("CÓDIGO Da SEÇÃO") = txtCódigo Then
                MsgBox "Relação violada!" + vbCr + "Para apagar esta seção, antes é necessário apagar" + vbCr + "todos os 'departamentos-seção' dela dependente.", vbExclamation, "Aviso"
                TBLDepartamentoSeção.Close
                Exit Sub
            End If
        End If
    Else
        Exit Sub
    End If
    TBLDepartamentoSeção.Close

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
    
    TBLSeção.Delete
    
    If Err <> 0 Then
        GeraMensagemDeErro "Seção - Excluir - " & txtDescrição, True
        StatusBarAviso = "Falha na exclusão"
        BarraDeStatus StatusBarAviso
        Exit Sub
    End If
    
    WS.CommitTrans
    
    Log gUsuário, "Exclusão - Seção: " & txtCódigo & " - " & txtDescrição
    
    StatusBarAviso = "Exclusão bem sucedida"
    BarraDeStatus StatusBarAviso
    
    If TBLSeção.RecordCount = 0 Then
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
    
    If TBLSeção.BOF Then
        TBLSeção.MoveFirst
    ElseIf TBLSeção.EOF Then
        TBLSeção.MoveLast
    Else
        TBLSeção.MovePrevious
        If TBLSeção.BOF Then
            TBLSeção.MoveNext
        End If
    End If
    
    GetRecords
    
    TestaInferior TBLSeção, lAllowEdit, lAllowDelete, lAllowConsult
    TestaSuperior TBLSeção, lAllowEdit, lAllowDelete, lAllowConsult
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
        If TBLSeção.RecordCount > 0 And Not TBLSeção.BOF And Not TBLSeção.EOF Then
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
    
    TestaInferior TBLSeção, lAllowEdit, lAllowDelete, lAllowConsult
    TestaSuperior TBLSeção, lAllowEdit, lAllowDelete, lAllowConsult
    
    If TBLSeção.RecordCount = 0 Then
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
    
    TBLSeção.MoveFirst
    
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
    
    TBLSeção.MoveLast
    
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
    
    TBLSeção.MoveNext
    If TBLSeção.EOF Then
        TBLSeção.MovePrevious
        Exit Sub
    End If
    
    StatusBarAviso = "Pronto"
    BarraDeStatus StatusBarAviso
    
    NavegaçãoInferior lAllowConsult
    TestaSuperior TBLSeção, lAllowEdit, lAllowDelete, lAllowConsult
    
    GetRecords
End Sub
Public Sub MovePrevious()
    If lAlterar Then
       If Not Cancelamento Then
            Exit Sub
        End If
    End If
    
    TBLSeção.MovePrevious
    If TBLSeção.BOF Then
        TBLSeção.MoveNext
        Exit Sub
    End If
    
    StatusBarAviso = "Pronto"
    BarraDeStatus StatusBarAviso
    
    NavegaçãoSuperior lAllowConsult
    TestaInferior TBLSeção, lAllowEdit, lAllowDelete, lAllowConsult
    
    GetRecords
End Sub
Public Sub PosRecords()
    If TBLSeção.RecordCount = 0 Then
        Exit Sub
    End If
    
    TBLSeção.Seek "=", txtCódigo
    If TBLSeção.NoMatch Then
        MsgBox "Não consegui encontrar " + txtCódigo, vbExclamation, "Erro"
        TBLSeção.MoveFirst
        NavegaçãoInferior False
        NavegaçãoInferior lAllowConsult
    Else
        TestaInferior TBLSeção, lAllowEdit, lAllowDelete, lAllowConsult
        TestaSuperior TBLSeção, lAllowEdit, lAllowDelete, lAllowConsult
    End If
    GetRecords
End Sub
Public Function PushDataBaseName(ByVal Posição As Integer) As String
    PushDataBaseName = DataBaseName(Posição)
End Function
Private Sub GetRecords()
    If Not lAllowConsult Then
        ZeraCampos
        DesativaCampos
        Exit Sub
    End If
    txtCódigo = TBLSeção("CÓDIGO")
    txtCódigoAnterior = txtCódigo
    txtDescrição = TBLSeção("DESCRIÇÃO")
    If Not lAllowEdit Then
        DesativaCampos
    End If
End Sub
Private Function SetRecords()
    On Error Resume Next
    
    Dim Msg$
    Dim Confirmação As Integer, Msg1$, Msg2$, AchouDepartamentoSeção As Boolean
    Dim TBLDepartamentoSeção As Table
    Dim SQL As String
    Dim Cont%

    If (txtCódigo <> txtCódigoAnterior) And Not lInserir Then
        If AbreTabela(Dicionário, "CADASTRO", "DEPARTAMENTO - SEÇÃO", DBCadastro, TBLDepartamentoSeção, TBLTabela, dbOpenTable) Then
            TBLDepartamentoSeção.Index = "DEPARTAMENTOSEÇÃO2"
            TBLDepartamentoSeção.Seek ">=", txtCódigoAnterior
            If Not TBLDepartamentoSeção.NoMatch Then
                If TBLDepartamentoSeção("CÓDIGO DA SEÇÃO") = txtCódigoAnterior Then
                    AchouDepartamentoSeção = True
                    Confirmação = MsgBox("Você necessita alterar os 'Departamentos-Seção' relacionados com esta seção !" + vbCr + "Deseja realizar agora as alterações de" + vbCr + "todas os 'departamentos-seção' dela dependente?", vbYesNo + vbQuestion + vbDefaultButton2, "Aviso")
                End If
            Else
                AchouDepartamentoSeção = False
            End If
        Else
            Exit Function
        End If
        TBLDepartamentoSeção.Close
        
        If AchouDepartamentoSeção Then
            If Confirmação = vbNo Then
                SetRecords = False
                Exit Function
            End If
        End If
    Else
        AchouDepartamentoSeção = False
    End If
    
    On Error GoTo Erro
    
    WS.BeginTrans 'Inicia uma Transação
        
    If lInserir Then
        TBLSeção.AddNew
    Else
        TBLSeção.Edit
    End If
    
    TBLSeção("CÓDIGO") = txtCódigo
    TBLSeção("DESCRIÇÃO") = txtDescrição
    If lInserir Then
        TBLSeção("USERNAME - CRIA") = gUsuário
        TBLSeção("DATA - CRIA") = Date
        TBLSeção("HORA - CRIA") = Time
        TBLSeção("USERNAME - ALTERA") = "VAZIO"
        TBLSeção("DATA - ALTERA") = vbNull
        TBLSeção("HORA - ALTERA") = vbNull
    End If
    If lAlterar Then
        TBLSeção("USERNAME - ALTERA") = gUsuário
        TBLSeção("DATA - ALTERA") = Date
        TBLSeção("HORA - ALTERA") = Time
    End If
    TBLSeção.Update
        
    If AchouDepartamentoSeção Then
        SQL = "Update [DEPARTAMENTO - SEÇÃO] Set [CÓDIGO DA SEÇÃO]= '" + txtCódigo + "' Where [CÓDIGO DA SEÇÃO]= '" + txtCódigoAnterior + "'"
        DBCadastro.Execute SQL
    End If
        
Erro:
    If Err <> 0 Then
        TBLSeção.CancelUpdate
        GeraMensagemDeErro "Seção - SetRecords - " & txtDescrição, True
        SetRecords = False
        Exit Function
    End If

    WS.CommitTrans 'Grava as alterações ou inclusões se não houverem erros
    
    'Se a janela Departamento-Seção estiver aberta atualiza seus valores se necessário.
    If Not lInserir Then
        For Cont = 1 To Forms.Count - 1
            If Forms(Cont).Name = "frmDepartamentoSeção" Then
                If Forms(Cont).txtCódigoSeção = txtCódigoAnterior Then
                    Forms(Cont).txtCódigoSeção = txtCódigo
                    Forms(Cont).txtDescriçãoSeção = txtDescrição
                    Forms(Cont).PosRecords
                End If
            End If
        Next
    End If
    
    If lInserir Then
        Log gUsuário, "Inclusão - Seção: " & txtCódigo & " - " & txtDescrição
    Else
        Log gUsuário, "Alteração - Seção: " & txtCódigo & " - " & txtDescrição
    End If
    
    SetRecords = True
End Function
Private Sub ZeraCampos()
    txtCódigo = Empty
    txtDescrição = Empty
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
    If Not SeçãoAberto Then
        Unload Me
        Exit Sub
    End If
    TestaInferior TBLSeção, lAllowEdit, lAllowDelete, lAllowConsult
    TestaSuperior TBLSeção, lAllowEdit, lAllowDelete, lAllowConsult
    If TBLSeção.RecordCount = 0 Then
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
    
    lAllowInsert = Allow("SEÇÃO", "I")
    lAllowEdit = Allow("SEÇÃO", "A")
    lAllowDelete = Allow("SEÇÃO", "E")
    lAllowConsult = Allow("SEÇÃO", "C")
    
    ZeraCampos
    
    lPula = False
    lInserir = False
    lAlterar = False
    
    SeçãoAberto = AbreTabela(Dicionário, "CADASTRO", "SEÇÃO", DBCadastro, TBLSeção, TBLTabela, dbOpenTable)
    
    If SeçãoAberto Then
        IndiceSeçãoAtivo = "SEÇÃO1"
        TBLSeção.Index = IndiceSeçãoAtivo
    Else
        MsgBox "Não consegui abrir a tabela 'Seção' !", vbCritical, "Erro"
        Exit Sub
    End If
    
    BotãoIncluir lAllowInsert
 
    If TBLSeção.RecordCount = 0 Then
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
        
    If TBLSeção.RecordCount = 0 Or TBLSeção.RecordCount = 1 Then
        NavegaçãoSuperior False
    Else
        NavegaçãoInferior lAllowConsult
    End If
    
    StatusBarAviso = "Pronto"
    Relatório = AddPath(AplicaçãoPath, "REPORT\SEÇÃO.RPT")
    TotalDatabaseName = 1
    DataBaseName(1) = AddPath(AplicaçãoPath, "DATABASE\CADASTRO.MDB")
    mFechar = False
    Exit Sub
    
Erro:
    GeraMensagemDeErro "Seção - Load"
    mFechar = True
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
    
    Set frmSeção = Nothing
End Sub
Private Sub Form_Unload(Cancel As Integer)
    If SeçãoAberto Then
        TBLSeção.Close
    End If
    If Forms.Count = 2 Then
        AllBotões False
    End If
End Sub
Private Sub txtCódigo_Change()
    If Not lPula Then
        FormatMask "9999", txtCódigo
    End If
End Sub
Private Sub txtCódigo_KeyPress(KeyAscii As Integer)
    If Not lInserir Then
        lAlterar = True
        StatusBarAviso = "Alteração"
        BarraDeStatus StatusBarAviso
    End If
End Sub
Private Sub txtCódigo_LostFocus()
    LeftBlank txtCódigo
End Sub
Private Sub txtDescrição_Change()
    If Not lPula Then
        FormatMask "@!S30", txtDescrição
    End If
End Sub
Private Sub txtDescrição_KeyPress(KeyAscii As Integer)
    If Not lInserir Then
        lAlterar = True
        StatusBarAviso = "Alteração"
        BarraDeStatus StatusBarAviso
    End If
End Sub
