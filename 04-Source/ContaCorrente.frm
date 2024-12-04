VERSION 5.00
Begin VB.Form frmContaCorrente 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Conta Corrente"
   ClientHeight    =   3705
   ClientLeft      =   1770
   ClientTop       =   1515
   ClientWidth     =   6330
   Icon            =   "ContaCorrente.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3705
   ScaleWidth      =   6330
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   345
      Left            =   5040
      TabIndex        =   4
      Top             =   3330
      Width           =   1245
   End
   Begin VB.CommandButton cmdGravar 
      Caption         =   "&Gravar"
      Height          =   345
      Left            =   3720
      TabIndex        =   3
      Top             =   3330
      Width           =   1245
   End
   Begin VB.Frame frContaCorrente 
      Height          =   570
      Left            =   0
      TabIndex        =   13
      Top             =   2700
      Width           =   6330
      Begin VB.TextBox txtContaCorrente 
         Height          =   285
         Left            =   1530
         TabIndex        =   2
         Top             =   195
         Width           =   1035
      End
      Begin VB.Label lblContaCorrente 
         Caption         =   "Conta Corrente"
         Height          =   180
         Left            =   165
         TabIndex        =   14
         Top             =   225
         Width           =   1215
      End
   End
   Begin VB.Frame frAgência 
      Caption         =   "Agência"
      Height          =   1350
      Left            =   0
      TabIndex        =   9
      Top             =   1350
      Width           =   6330
      Begin VB.TextBox txtCódigoAgência 
         Height          =   285
         Left            =   1200
         TabIndex        =   1
         Top             =   300
         Width           =   1035
      End
      Begin VB.TextBox txtDescriçãoAgência 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1200
         TabIndex        =   10
         Top             =   750
         Width           =   5000
      End
      Begin VB.Label lblCódigoAgência 
         Caption         =   "Código"
         Height          =   210
         Left            =   150
         TabIndex        =   12
         Top             =   330
         Width           =   660
      End
      Begin VB.Label lblDescriçãoAgência 
         Caption         =   "Descrição"
         Height          =   180
         Left            =   150
         TabIndex        =   11
         Top             =   780
         Width           =   960
      End
   End
   Begin VB.Frame frBanco 
      Caption         =   " Banco "
      Height          =   1350
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   6330
      Begin VB.TextBox txtDescriçãoBanco 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1200
         TabIndex        =   8
         Top             =   750
         Width           =   5000
      End
      Begin VB.TextBox txtCódigoBanco 
         Height          =   285
         Left            =   1200
         TabIndex        =   0
         Top             =   300
         Width           =   700
      End
      Begin VB.Label lblDescriçãoBanco 
         Caption         =   "Descrição"
         Height          =   180
         Left            =   150
         TabIndex        =   7
         Top             =   780
         Width           =   960
      End
      Begin VB.Label lblCódigoBanco 
         Caption         =   "Código"
         Height          =   210
         Left            =   150
         TabIndex        =   6
         Top             =   330
         Width           =   660
      End
   End
End
Attribute VB_Name = "frmContaCorrente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim TBLBanco As Table
Dim BancoAberto As Boolean

Dim TBLAgência As Table
Dim AgênciaAberto As Boolean

Dim TBLContaCorrente As Table
Dim ContaCorrenteAberto As Boolean

Dim IndiceAtivoBanco$, IndiceAtivoAgência$, IndiceAtivoContaCorrente$

Dim lAllowInsert  As Boolean
Dim lAllowEdit    As Boolean
Dim lAllowDelete  As Boolean
Dim lAllowConsult As Boolean

Dim lPula As Boolean
Dim lInserir As Boolean
Dim lAlterar As Boolean
Dim mFechar As Boolean

Dim StatusBarAviso$

Public lAtualizar As Boolean
Private Sub AtivaCampos()
    frBanco.Enabled = True
    frAgência.Enabled = True
    frContaCorrente.Enabled = True
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
    
    If TBLContaCorrente.RecordCount = 0 Then
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
    
    TestaInferior TBLContaCorrente, lAllowEdit, lAllowDelete, lAllowConsult
    TestaSuperior TBLContaCorrente, lAllowEdit, lAllowDelete, lAllowConsult
    
    GetRecords
End Function
Private Sub DesativaCampos()
    frBanco.Enabled = False
    frAgência.Enabled = False
    frContaCorrente.Enabled = False
    cmdGravar.Enabled = False
    cmdCancelar.Enabled = False
    BotãoGravar False
End Sub
Public Sub Encontrar()
    If Not lAllowConsult Then
        Exit Sub
    End If
    Set frmEncontrar.DBBancoDeDados = DBFinanceiro
    frmEncontrar.NomeDaJanela = "Conta Corrente"
    frmEncontrar.LabelDescription = "Código"
    frmEncontrar.Mensagem = "Nenhuma conta corrente foi selecionada!"
    frmEncontrar.BancoDeDados = "FINANCEIRO"
    frmEncontrar.Tabela = "CONTA CORRENTE"
    frmEncontrar.Indice = "1"
    frmEncontrar.CampoChave = "CÓDIGO DO BANCO,CÓDIGO DA AGÊNCIA,CÓDIGO"
    frmEncontrar.CampoPreencheLista = "CÓDIGO DO BANCO,CÓDIGO DA AGÊNCIA,CÓDIGO"
    frmEncontrar.Show vbModal
    lPula = True
    txtCódigoBanco = GetWordSeparatedBy(frmEncontrar.Chave, 1)
    txtCódigoAgência = GetWordSeparatedBy(frmEncontrar.Chave, 2)
    txtContaCorrente = GetWordSeparatedBy(frmEncontrar.Chave, 3)
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
    
    TBLContaCorrente.Delete
    
    If Err <> 0 Then
        GeraMensagemDeErro "Conta Corrente - Excluir - " & txtDescriçãoBanco & " - " & txtDescriçãoAgência & " - " & txtContaCorrente, True
        StatusBarAviso = "Falha na exclusão"
        BarraDeStatus StatusBarAviso
        Exit Sub
    End If
    
    WS.CommitTrans
    
    Log gUsuário, "Exclusão - Conta Corrente: " & txtContaCorrente & vbCr & "Banco: " & txtDescriçãoBanco & vbCr & "Agência:" & txtDescriçãoAgência
    
    StatusBarAviso = "Exclusão bem sucedida"
    BarraDeStatus StatusBarAviso
    
    If TBLContaCorrente.RecordCount = 0 Then
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
    
    If TBLContaCorrente.BOF Then
        TBLContaCorrente.MoveFirst
    ElseIf TBLContaCorrente.EOF Then
        TBLContaCorrente.MoveLast
    Else
        TBLContaCorrente.MovePrevious
        If TBLContaCorrente.BOF Then
            TBLContaCorrente.MoveNext
        End If
    End If
    
    GetRecords
    
    TestaInferior TBLContaCorrente, lAllowEdit, lAllowDelete, lAllowConsult
    TestaSuperior TBLContaCorrente, lAllowEdit, lAllowDelete, lAllowConsult
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
        If TBLContaCorrente.RecordCount > 0 And Not TBLContaCorrente.BOF And Not TBLContaCorrente.EOF Then
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
    
    TestaInferior TBLContaCorrente, lAllowEdit, lAllowDelete, lAllowConsult
    TestaSuperior TBLContaCorrente, lAllowEdit, lAllowDelete, lAllowConsult
    
    If TBLContaCorrente.RecordCount = 0 Then
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
    
    TBLContaCorrente.MoveFirst
    
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
    
    TBLContaCorrente.MoveLast
    
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
    
    TBLContaCorrente.MoveNext
    If TBLContaCorrente.EOF Then
        TBLContaCorrente.MovePrevious
        Exit Sub
    End If
    
    StatusBarAviso = "Pronto"
    BarraDeStatus StatusBarAviso
    
    NavegaçãoInferior lAllowConsult
    TestaSuperior TBLContaCorrente, lAllowEdit, lAllowDelete, lAllowConsult
    
    GetRecords
End Sub
Public Sub MovePrevious()
    If lAlterar Then
       If Not Cancelamento Then
            Exit Sub
        End If
    End If
    
    TBLContaCorrente.MovePrevious
    If TBLContaCorrente.BOF Then
        TBLContaCorrente.MoveNext
        Exit Sub
    End If
    
    StatusBarAviso = "Pronto"
    BarraDeStatus StatusBarAviso
    
    NavegaçãoSuperior lAllowConsult
    TestaInferior TBLContaCorrente, lAllowEdit, lAllowDelete, lAllowConsult
    
    GetRecords
End Sub
Public Sub PosRecords()
    If TBLContaCorrente.RecordCount = 0 Then
        Exit Sub
    End If
    TBLContaCorrente.Seek "=", txtCódigoBanco, txtCódigoAgência, txtContaCorrente
    If TBLContaCorrente.NoMatch Then
        MsgBox "Não consegui encontrar a Conta Corrente" + txtContaCorrente, vbExclamation, "Erro"
        TBLContaCorrente.MoveFirst
        NavegaçãoInferior False
        NavegaçãoInferior lAllowConsult
    Else
        TestaInferior TBLContaCorrente, lAllowEdit, lAllowDelete, lAllowConsult
        TestaSuperior TBLContaCorrente, lAllowEdit, lAllowDelete, lAllowConsult
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
    txtCódigoBanco = TBLContaCorrente("CÓDIGO DO BANCO")
    TBLBanco.Seek "=", txtCódigoBanco
    txtDescriçãoBanco = TBLBanco("DESCRIÇÃO")
    txtCódigoAgência = TBLContaCorrente("CÓDIGO DA AGÊNCIA")
    TBLAgência.Seek "=", txtCódigoBanco, txtCódigoAgência
    txtDescriçãoAgência = TBLAgência("DESCRIÇÃO")
    txtContaCorrente = TBLContaCorrente("CÓDIGO")
    If Not lAllowEdit Then
        DesativaCampos
    End If
    
    Exit Sub
    
Erro:
    GeraMensagemDeErro "Conta Corrente - GetRecords "
    Resume Next
End Sub
Private Function SetRecords()
    On Error GoTo Erro
    
    Dim Msg$
    
    WS.BeginTrans 'Inicia transações
        
    If lInserir Then
        TBLContaCorrente.AddNew
    Else
        TBLContaCorrente.Edit
    End If
    
    TBLContaCorrente("CÓDIGO DO BANCO") = txtCódigoBanco
    TBLContaCorrente("CÓDIGO DA AGÊNCIA") = txtCódigoAgência
    TBLContaCorrente("CÓDIGO") = txtContaCorrente
    If lInserir Then
        TBLContaCorrente("USERNAME - CRIA") = gUsuário
        TBLContaCorrente("DATA - CRIA") = Date
        TBLContaCorrente("HORA - CRIA") = Time
        TBLContaCorrente("USERNAME - ALTERA") = "VAZIO"
        TBLContaCorrente("DATA - ALTERA") = vbNull
        TBLContaCorrente("HORA - ALTERA") = vbNull
    End If
    If lAlterar Then
        TBLContaCorrente("USERNAME - ALTERA") = gUsuário
        TBLContaCorrente("DATA - ALTERA") = Date
        TBLContaCorrente("HORA - ALTERA") = Time
    End If
    TBLContaCorrente.Update
    
Erro:
    If Err <> 0 Then
        GeraMensagemDeErro "Conta Corrente - SetRecords - " & txtDescriçãoBanco & " - " & txtDescriçãoAgência & " - " & txtContaCorrente, True
        On Error Resume Next
        SetRecords = False
        Exit Function
    End If
    
    WS.CommitTrans 'Grava as alterações ou inclusões se não houverem erros
    
    If lInserir Then
        Log gUsuário, "Inclusão - Conta Corrente: " & txtContaCorrente & vbCr & "Banco" & txtDescriçãoBanco & vbCr & "Agência: " & txtDescriçãoAgência
    Else
        Log gUsuário, "Alteração - Conta Corrente: " & txtContaCorrente & vbCr & "Banco" & txtDescriçãoBanco & vbCr & "Agência: " & txtDescriçãoAgência
    End If
    
    SetRecords = True
End Function
Private Sub ZeraCampos()
    txtCódigoBanco = Empty
    txtDescriçãoBanco = Empty
    txtCódigoAgência = Empty
    txtDescriçãoAgência = Empty
    txtContaCorrente = Empty
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
    If Not BancoAberto Then
        Unload Me
        Exit Sub
    End If
    If Not AgênciaAberto Then
        Unload Me
        Exit Sub
    End If
    If Not ContaCorrenteAberto Then
        Unload Me
        Exit Sub
    End If
    
    TestaInferior TBLContaCorrente, lAllowEdit, lAllowDelete, lAllowConsult
    TestaSuperior TBLContaCorrente, lAllowEdit, lAllowDelete, lAllowConsult
    
    If TBLContaCorrente.RecordCount = 0 Then
        cmdGravar.Enabled = False
        cmdCancelar.Enabled = False
    Else
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
    On Error GoTo Erro
    
    lAllowInsert = Allow("CONTA CORRENTE", "I")
    lAllowEdit = Allow("CONTA CORRENTE", "A")
    lAllowDelete = Allow("CONTA CORRENTE", "E")
    lAllowConsult = Allow("CONTA CORRENTE", "C")
    
    ZeraCampos
    
    lPula = False
    lInserir = False
    lAlterar = False
    
    BancoAberto = AbreTabela(Dicionário, "FINANCEIRO", "BANCO", DBFinanceiro, TBLBanco, TBLTabela, dbOpenTable)
    
    If BancoAberto Then
        IndiceAtivoBanco = "BANCO1"
        TBLBanco.Index = IndiceAtivoBanco
    Else
        MsgBox "Não consegui abrir a tabela 'Banco' !", vbCritical, "Erro"
        Exit Sub
    End If
    
    AgênciaAberto = AbreTabela(Dicionário, "FINANCEIRO", "AGÊNCIA", DBFinanceiro, TBLAgência, TBLTabela, dbOpenTable)
    
    If AgênciaAberto Then
        IndiceAtivoAgência = "AGÊNCIA1"
        TBLAgência.Index = IndiceAtivoAgência
    Else
        MsgBox "Não consegui abrir a tabela 'Agência' !", vbCritical, "Erro"
        Exit Sub
    End If
    
    ContaCorrenteAberto = AbreTabela(Dicionário, "FINANCEIRO", "CONTA CORRENTE", DBFinanceiro, TBLContaCorrente, TBLTabela, dbOpenTable)
    
    If ContaCorrenteAberto Then
        IndiceAtivoContaCorrente = "CONTACORRENTE1"
        TBLContaCorrente.Index = IndiceAtivoContaCorrente
    Else
        MsgBox "Não consegui abrir a tabela 'Conta Corrente' !", vbCritical, "Erro"
        Exit Sub
    End If
    
    BotãoIncluir lAllowInsert
 
    If TBLContaCorrente.RecordCount = 0 Then
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
        
    If TBLContaCorrente.RecordCount = 0 Or TBLContaCorrente.RecordCount = 1 Then
        NavegaçãoSuperior False
    Else
        NavegaçãoInferior lAllowConsult
    End If

    StatusBarAviso = "Pronto"
    mFechar = False
    Exit Sub
    
Erro:
    GeraMensagemDeErro "Conta Corrente - Load"
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
    
    Set frmContaCorrente = Nothing
End Sub
Private Sub Form_Unload(Cancel As Integer)
    If BancoAberto Then
        TBLBanco.Close
    End If
    If AgênciaAberto Then
        TBLAgência.Close
    End If
    If ContaCorrenteAberto Then
        TBLContaCorrente.Close
    End If
    If Forms.Count = 2 Then
        AllBotões False
    End If
End Sub
Private Sub txtCódigoAgência_Change()
    If Not lPula Then
        FormatMask "@S10", txtCódigoAgência
    End If
End Sub
Private Sub txtCódigoAgência_KeyPress(KeyAscii As Integer)
    If Not lInserir Then
        lAlterar = True
        StatusBarAviso = "Alteração"
        BarraDeStatus StatusBarAviso
    End If
End Sub
Private Sub txtCódigoAgência_LostFocus()
    If mdiGeal.ActiveForm.Name = "frmContaCorrente" Then
        If txtCódigoAgência.Enabled Then
            LeftBlank txtCódigoAgência
            TBLAgência.Seek "=", txtCódigoBanco, txtCódigoAgência
            If TBLAgência.NoMatch Then
                MsgBox "Não encontrei a agência !" + txtCódigoAgência, vbExclamation, "Aviso"
                txtCódigoAgência = Empty
                txtCódigoAgência.SetFocus
                Exit Sub
            End If
            txtDescriçãoAgência = TBLAgência("DESCRIÇÃO")
        Else
            If txtCódigoBanco.Enabled Then
                txtCódigoAgência.Enabled = True
            End If
        End If
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
    If mdiGeal.ActiveForm.Name = "frmContaCorrente" Then
        If txtCódigoBanco.Enabled Then
            LeftBlank txtCódigoBanco
            TBLBanco.Seek "=", txtCódigoBanco
            If TBLBanco.NoMatch Then
                MsgBox "Não encontrei o banco " + txtCódigoBanco, vbExclamation, "Aviso"
                txtCódigoBanco = Empty
                txtCódigoAgência.Enabled = False
                txtCódigoBanco.SetFocus
                Exit Sub
            End If
            txtDescriçãoBanco = TBLBanco("DESCRIÇÃO")
        End If
    End If
End Sub
Private Sub txtContaCorrente_Change()
    If Not lPula Then
        FormatMask "@S10", txtContaCorrente
    End If
End Sub
Private Sub txtContaCorrente_KeyPress(KeyAscii As Integer)
    If Not lInserir Then
        lAlterar = True
        StatusBarAviso = "Alteração"
        BarraDeStatus StatusBarAviso
    End If
End Sub


