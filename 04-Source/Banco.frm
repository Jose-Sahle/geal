VERSION 5.00
Begin VB.Form frmBanco 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Banco"
   ClientHeight    =   1770
   ClientLeft      =   1140
   ClientTop       =   1515
   ClientWidth     =   6330
   Icon            =   "Banco.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   1770
   ScaleWidth      =   6330
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Enabled         =   0   'False
      Height          =   345
      Left            =   5070
      TabIndex        =   4
      Top             =   1395
      Width           =   1245
   End
   Begin VB.CommandButton cmdGravar 
      Caption         =   "&Gravar"
      Height          =   345
      Left            =   3750
      TabIndex        =   2
      Top             =   1395
      Width           =   1245
   End
   Begin VB.Frame frBanco 
      Height          =   1350
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   6330
      Begin VB.TextBox txtCódigoBanco 
         Height          =   285
         Left            =   1200
         TabIndex        =   0
         Top             =   300
         Width           =   700
      End
      Begin VB.TextBox txtDescriçãoBanco 
         Height          =   285
         Left            =   1200
         TabIndex        =   1
         Top             =   750
         Width           =   5000
      End
      Begin VB.Label lblCódigoBanco 
         Caption         =   "Código"
         Height          =   210
         Left            =   150
         TabIndex        =   6
         Top             =   330
         Width           =   660
      End
      Begin VB.Label lblDescriçãoBanco 
         Caption         =   "Descrição"
         Height          =   180
         Left            =   150
         TabIndex        =   5
         Top             =   780
         Width           =   960
      End
   End
End
Attribute VB_Name = "frmBanco"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim TBLBanco As Table
Dim BancoAberto As Boolean
Dim IndiceBancoAtivo$
Dim txtCódigoBancoAnterior As String

Dim lPula As Boolean

Dim lInserir As Boolean
Dim lAlterar As Boolean

Dim lFechar As Boolean

Dim lAllowInsert  As Boolean
Dim lAllowEdit As Boolean
Dim lAllowDelete As Boolean
Dim lAllowConsult As Boolean

Dim StatusBarAviso$

Dim DataBaseName(1 To 1) As String
Public Relatório$
Public TotalDatabaseName%

Public lAtualizar As Boolean
Private Sub AtivaCampos()
    BotãoImprimir True
    frBanco.Enabled = True
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
    
    If TBLBanco.RecordCount = 0 Then
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
    
    TestaInferior TBLBanco, lAllowEdit, lAllowDelete, lAllowConsult
    TestaSuperior TBLBanco, lAllowEdit, lAllowDelete, lAllowConsult
    
    GetRecords
End Function
Private Sub DesativaCampos()
    BotãoImprimir False
    frBanco.Enabled = False
    cmdGravar.Enabled = False
    cmdCancelar.Enabled = False
    BotãoGravar False
End Sub
Public Sub Encontrar()
    If Not lAllowConsult Then
        Exit Sub
    End If
    Set frmEncontrar.DBBancoDeDados = DBFinanceiro
    frmEncontrar.NomeDaJanela = "Banco"
    frmEncontrar.LabelDescription = "Descrição"
    frmEncontrar.Mensagem = "Nenhum banco foi selecionado!"
    frmEncontrar.BancoDeDados = "FINANCEIRO"
    frmEncontrar.Tabela = "BANCO"
    frmEncontrar.Indice = "1"
    frmEncontrar.CampoChave = "CÓDIGO"
    frmEncontrar.CampoPreencheLista = "DESCRIÇÃO"
    frmEncontrar.Show vbModal
    lPula = True
    txtCódigoBanco = frmEncontrar.Chave
    lPula = False
    PosRecords
End Sub
Public Sub Excluir()
    Dim Confirmação As Integer, Msg1$, Msg2$
    Dim TBLAgência As Table

    If lAlterar Then
       If Not Cancelamento Then
            Exit Sub
        End If
    End If

    If AbreTabela(Dicionário, "FINANCEIRO", "AGÊNCIA", DBFinanceiro, TBLAgência, TBLTabela, dbOpenTable) Then
        TBLAgência.Index = "AGÊNCIA1"
        TBLAgência.Seek ">=", txtCódigoBanco
        If Not TBLAgência.NoMatch Then
            If TBLAgência("CÓDIGO DO BANCO") = txtCódigoBanco Then
                MsgBox "Relação violada!" + vbCr + "Para apagar este banco, antes é necessário apagar" + vbCr + "todas as agências dele dependente.", vbExclamation, "Aviso"
                TBLAgência.Close
                Exit Sub
            End If
        End If
    Else
        Exit Sub
    End If
    TBLAgência.Close
    
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
    
    TBLBanco.Delete
    
    If Err <> 0 Then
        GeraMensagemDeErro "Banco - Excluir - " & txtDescriçãoBanco, True
        StatusBarAviso = "Falha na exclusão"
        BarraDeStatus StatusBarAviso
        Exit Sub
    End If
    
    WS.CommitTrans
    
    Log gUsuário, "Exclusão - Banco: " & txtCódigoBanco & " - " & txtDescriçãoBanco
    
    StatusBarAviso = "Exclusão bem sucedida"
    BarraDeStatus StatusBarAviso
    
    If TBLBanco.RecordCount = 0 Then
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
    
    If TBLBanco.BOF Then
        TBLBanco.MoveFirst
    ElseIf TBLBanco.EOF Then
        TBLBanco.MoveLast
    Else
        TBLBanco.MovePrevious
        If TBLBanco.BOF Then
            TBLBanco.MoveNext
        End If
    End If
    
    GetRecords
    
    TestaInferior TBLBanco, lAllowEdit, lAllowDelete, lAllowConsult
    TestaSuperior TBLBanco, lAllowEdit, lAllowDelete, lAllowConsult
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
        If TBLBanco.RecordCount > 0 And Not TBLBanco.BOF And Not TBLBanco.EOF Then
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
    
    TestaInferior TBLBanco, lAllowEdit, lAllowDelete, lAllowConsult
    TestaSuperior TBLBanco, lAllowEdit, lAllowDelete, lAllowConsult
    
    If TBLBanco.RecordCount = 0 Then
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
    
    If txtCódigoBanco.Enabled Then
        txtCódigoBanco.SetFocus
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
    
    txtCódigoBanco.SetFocus
End Sub
Public Sub MoveFirst()
    If lAlterar Then
       If Not Cancelamento Then
            Exit Sub
        End If
    End If
    
    StatusBarAviso = "Pronto"
    BarraDeStatus StatusBarAviso
    
    TBLBanco.MoveFirst
    
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
    
    TBLBanco.MoveLast
    
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
    
    TBLBanco.MoveNext
    If TBLBanco.EOF Then
        TBLBanco.MovePrevious
        Exit Sub
    End If
    
    StatusBarAviso = "Pronto"
    BarraDeStatus StatusBarAviso
    
    NavegaçãoInferior lAllowConsult
    TestaSuperior TBLBanco, lAllowEdit, lAllowDelete, lAllowConsult
    
    GetRecords
End Sub
Public Sub MovePrevious()
    If lAlterar Then
       If Not Cancelamento Then
            Exit Sub
        End If
    End If
    
    TBLBanco.MovePrevious
    If TBLBanco.BOF Then
        TBLBanco.MoveNext
        Exit Sub
    End If
    
    StatusBarAviso = "Pronto"
    BarraDeStatus StatusBarAviso
    
    NavegaçãoSuperior lAllowConsult
    TestaInferior TBLBanco, lAllowEdit, lAllowDelete, lAllowConsult
    
    GetRecords
End Sub
Public Sub PosRecords()
    If TBLBanco.RecordCount = 0 Then
        Exit Sub
    End If
    
    TBLBanco.Seek "=", txtCódigoBanco
    If TBLBanco.NoMatch Then
        MsgBox "Não consegui encontrar " + txtCódigoBanco, vbExclamation, "Erro"
        TBLBanco.MoveFirst
        NavegaçãoInferior False
        NavegaçãoInferior lAllowConsult
    Else
        TestaInferior TBLBanco, lAllowEdit, lAllowDelete, lAllowConsult
        TestaSuperior TBLBanco, lAllowEdit, lAllowDelete, lAllowConsult
    End If
    GetRecords
End Sub
Public Function PushDataBaseName(ByVal Posição As Integer) As String
    PushDataBaseName = DataBaseName(Posição)
End Function
Private Sub GetRecords()
    On Error GoTo Erro
    
    If Not lAllowConsult Then
        ZeraCampos
        DesativaCampos
        Exit Sub
    End If
    txtCódigoBanco = TBLBanco("CÓDIGO")
    txtCódigoBancoAnterior = txtCódigoBanco
    txtDescriçãoBanco = TBLBanco("DESCRIÇÃO")
    If Not lAllowEdit Then
        DesativaCampos
    End If
    
    Exit Sub
    
Erro:
    GeraMensagemDeErro "Banco - GetRecords"
    Resume Next
End Sub
Private Function SetRecords()
    On Error Resume Next
    
    Dim Msg$
    Dim Confirmação As Integer, Msg1$, Msg2$, AchouAgência As Boolean, AchouContaCorrente As Boolean
    Dim TBLAgência As Table
    Dim TBLContaCorrente As Table
    Dim SQL As String
    Dim Cont%

    If (txtCódigoBanco <> txtCódigoBancoAnterior) And Not lInserir Then
        If AbreTabela(Dicionário, "FINANCEIRO", "AGÊNCIA", DBFinanceiro, TBLAgência, TBLTabela, dbOpenTable) Then
            TBLAgência.Index = "AGÊNCIA1"
            TBLAgência.Seek ">=", txtCódigoBancoAnterior
            If Not TBLAgência.NoMatch Then
                If TBLAgência("CÓDIGO DO BANCO") = txtCódigoBancoAnterior Then
                    AchouAgência = True
                    Confirmação = MsgBox("Você necessita alterar as agências relacionadas com este banco !" + vbCr + "Deseja realizar agora as alterações de" + vbCr + "todas as agências dele dependente?", vbYesNo + vbQuestion + vbDefaultButton2, "Aviso")
                End If
            Else
                AchouAgência = False
            End If
        Else
            Exit Function
        End If
        TBLAgência.Close
        
        If AchouAgência Then
            If Confirmação = vbNo Then
                SetRecords = False
                Exit Function
            End If
        End If
        
        If AbreTabela(Dicionário, "FINANCEIRO", "CONTA CORRENTE", DBFinanceiro, TBLContaCorrente, TBLTabela, dbOpenTable) Then
            TBLContaCorrente.Index = "CONTACORRENTE1"
            TBLContaCorrente.Seek ">=", txtCódigoBancoAnterior
            If Not TBLContaCorrente.NoMatch Then
                If TBLContaCorrente("CÓDIGO DO BANCO") = txtCódigoBancoAnterior Then
                    AchouContaCorrente = True
                End If
            Else
                AchouContaCorrente = False
            End If
        Else
            Exit Function
        End If
        TBLContaCorrente.Close
    Else
        AchouAgência = False
        AchouContaCorrente = False
    End If
    
    On Error GoTo Erro
    
    WS.BeginTrans 'Inicia uma Transação
    
    If lInserir Then
        TBLBanco.AddNew
    Else
        TBLBanco.Edit
    End If
    
    TBLBanco("CÓDIGO") = Trim(txtCódigoBanco)
    TBLBanco("DESCRIÇÃO") = Trim(txtDescriçãoBanco)
    If lInserir Then
        TBLBanco("USERNAME - CRIA") = gUsuário
        TBLBanco("DATA - CRIA") = Date
        TBLBanco("HORA - CRIA") = Time
        TBLBanco("USERNAME - ALTERA") = "VAZIO"
        TBLBanco("DATA - ALTERA") = vbNull
        TBLBanco("HORA - ALTERA") = vbNull
    End If
    If lAlterar Then
        TBLBanco("USERNAME - ALTERA") = gUsuário
        TBLBanco("DATA - ALTERA") = Date
        TBLBanco("HORA - ALTERA") = Time
    End If
    TBLBanco.Update
        
    If AchouAgência Then
        SQL = "Update AGÊNCIA Set [CÓDIGO DO BANCO]= '" + txtCódigoBanco + "' Where [CÓDIGO DO BANCO]= '" + txtCódigoBancoAnterior + "'"
        DBFinanceiro.Execute SQL
    End If
    If AchouContaCorrente Then
        SQL = "Update [CONTA CORRENTE] Set [CÓDIGO DO BANCO]= '" + txtCódigoBanco + "' Where [CÓDIGO DO BANCO]= '" + txtCódigoBancoAnterior + "'"
        DBFinanceiro.Execute SQL
    End If
    
Erro:
    If Err <> 0 Then
        TBLBanco.CancelUpdate
        GeraMensagemDeErro "Banco - SetRecords - " & txtDescriçãoBanco, True
        SetRecords = False
        Exit Function
    End If

    WS.CommitTrans 'Grava as alterações ou inclusões se não houverem erros
    
    'Se a janela Agência estiver aberta atualiza seus valores se necessário.
    If Not lInserir Then
        For Cont = 1 To Forms.Count - 1
            If Forms(Cont).Name = "frmAgência" Or Forms(Cont).Name = "frmContaCorrente" Then
                If Forms(Cont).txtCódigoBanco = txtCódigoBancoAnterior Then
                    Forms(Cont).txtCódigoBanco = txtCódigoBanco
                    Forms(Cont).txtDescriçãoBanco = txtDescriçãoBanco
                    Forms(Cont).PosRecords
                End If
            End If
        Next
    End If
    
    If lInserir Then
        Log gUsuário, "Inclusão - Banco: " & txtCódigoBanco & " - " & txtDescriçãoBanco
    Else
        Log gUsuário, "Alteração - Banco: " & txtCódigoBanco & " - " & txtDescriçãoBanco
    End If
    
    SetRecords = True
End Function
Private Sub ZeraCampos()
    txtCódigoBanco = Empty
    txtCódigoBancoAnterior = Empty
    txtDescriçãoBanco = Empty
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
    If Not BancoAberto Then
        Unload Me
        Exit Sub
    End If
    TestaInferior TBLBanco, lAllowEdit, lAllowDelete, lAllowConsult
    TestaSuperior TBLBanco, lAllowEdit, lAllowDelete, lAllowConsult
    If TBLBanco.RecordCount = 0 Then
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
    
    lAllowInsert = Allow("BANCO", "I")
    lAllowEdit = Allow("BANCO", "A")
    lAllowDelete = Allow("BANCO", "E")
    lAllowConsult = Allow("BANCO", "C")
    
    ZeraCampos
    
    lPula = False
    lInserir = False
    lAlterar = False
    
    BancoAberto = AbreTabela(Dicionário, "FINANCEIRO", "BANCO", DBFinanceiro, TBLBanco, TBLTabela, dbOpenTable)
    
    If BancoAberto Then
        IndiceBancoAtivo = "BANCO1"
        TBLBanco.Index = IndiceBancoAtivo
    Else
        MsgBox "Não consegui abrir a tabela 'Banco' !", vbCritical, "Erro"
        Exit Sub
    End If
    
    BotãoIncluir lAllowInsert
 
    If TBLBanco.RecordCount = 0 Then
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
        
    If TBLBanco.RecordCount = 0 Or TBLBanco.RecordCount = 1 Then
        NavegaçãoSuperior False
    Else
        NavegaçãoInferior lAllowConsult
    End If
    
    StatusBarAviso = "Pronto"
    Relatório = AddPath(AplicaçãoPath, "REPORT\BANCO.RPT")
    TotalDatabaseName = 1
    DataBaseName(1) = AddPath(AplicaçãoPath, "DATABASE\FINANCEIRO.MDB")
    lFechar = False
    Exit Sub
    
Erro:
    GeraMensagemDeErro "Banco - Load"
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
    
    Set frmBanco = Nothing
End Sub
Private Sub Form_Unload(Cancel As Integer)
    If BancoAberto Then
        TBLBanco.Close
    End If
    If Forms.Count = 2 Then
        AllBotões False
    End If
End Sub
Private Sub txtCódigoBanco_Change()
    If Not lPula Then
        FormatMask "9999", txtCódigoBanco
    End If
End Sub
Private Sub txtCódigoBanco_KeyPress(KeyAscii As Integer)
    If Not lInserir Then
        lAlterar = True
        StatusBarAviso = "Alteração"
        BarraDeStatus StatusBarAviso
    End If
End Sub
Private Sub txtCódigoBanco_LostFocus()
    If txtCódigoBanco.Enabled Then
        LeftBlank txtCódigoBanco
    End If
End Sub
Private Sub txtDescriçãoBanco_Change()
    If Not lPula Then
        FormatMask "@!S30", txtDescriçãoBanco
    End If
End Sub
Private Sub txtDescriçãoBanco_KeyPress(KeyAscii As Integer)
    If Not lInserir Then
        lAlterar = True
        StatusBarAviso = "Alteração"
        BarraDeStatus StatusBarAviso
    End If
End Sub

