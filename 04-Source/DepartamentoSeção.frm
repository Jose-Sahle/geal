VERSION 5.00
Begin VB.Form frmDepartamentoSeção 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Departamento - Seção"
   ClientHeight    =   1695
   ClientLeft      =   870
   ClientTop       =   1515
   ClientWidth     =   8250
   Icon            =   "DepartamentoSeção.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   1695
   ScaleWidth      =   8250
   Begin VB.Frame frDepartamentoSeção 
      Height          =   1275
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   8235
      Begin VB.TextBox txtDescriçãoSeção 
         Enabled         =   0   'False
         Height          =   285
         Left            =   2100
         MaxLength       =   30
         TabIndex        =   8
         Top             =   750
         Width           =   5000
      End
      Begin VB.TextBox txtCódigoSeção 
         Height          =   285
         Left            =   1230
         TabIndex        =   1
         Top             =   750
         Width           =   750
      End
      Begin VB.TextBox txtCódigoDepartamento 
         Height          =   285
         Left            =   1230
         TabIndex        =   0
         Top             =   300
         Width           =   750
      End
      Begin VB.TextBox txtDescriçãoDepartamento 
         Enabled         =   0   'False
         Height          =   285
         Left            =   2100
         MaxLength       =   30
         TabIndex        =   5
         Top             =   300
         Width           =   5000
      End
      Begin VB.Label lblDepartamento 
         Caption         =   "Departamento"
         Height          =   195
         Left            =   150
         TabIndex        =   7
         Top             =   330
         Width           =   1035
      End
      Begin VB.Label lblSeção 
         Caption         =   "Seção"
         Height          =   195
         Left            =   150
         TabIndex        =   6
         Top             =   780
         Width           =   885
      End
   End
   Begin VB.CommandButton cmdGravar 
      Caption         =   "&Gravar"
      Height          =   345
      Left            =   5685
      TabIndex        =   2
      Top             =   1320
      Width           =   1245
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   345
      Left            =   7005
      TabIndex        =   3
      Top             =   1320
      Width           =   1245
   End
End
Attribute VB_Name = "frmDepartamentoSeção"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim TBLDepartamentoSeção As Table
Dim DepartamentoSeçãoAberto As Boolean
Dim IndiceDepartamentoSeçãoAtivo$

Dim TBLDepartamento As Table
Dim DepartamentoAberto As Boolean
Dim IndiceDepartamentoAtivo$

Dim TBLSeção As Table
Dim SeçãoAberto As Boolean
Dim IndiceSeçãoAtivo$

Dim lAllowInsert  As Boolean
Dim lAllowEdit    As Boolean
Dim lAllowDelete  As Boolean
Dim lAllowConsult As Boolean

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
    frDepartamentoSeção.Enabled = True
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
    
    If TBLDepartamentoSeção.RecordCount = 0 Then
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
    
    TestaInferior TBLDepartamentoSeção, lAllowEdit, lAllowDelete, lAllowConsult
    TestaSuperior TBLDepartamentoSeção, lAllowEdit, lAllowDelete, lAllowConsult
    
    GetRecords
End Function
Private Sub DesativaCampos()
    BotãoImprimir False
    frDepartamentoSeção.Enabled = False
    cmdGravar.Enabled = False
    cmdCancelar.Enabled = False
    BotãoGravar False
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
    
    TBLDepartamentoSeção.Delete
    
    If Err <> 0 Then
        GeraMensagemDeErro "Departamento - Seção - Excluir - " & txtCódigoDepartamento & txtCódigoSeção, True
        StatusBarAviso = "Falha na exclusão"
        BarraDeStatus StatusBarAviso
        Exit Sub
    End If
    
    WS.CommitTrans
    
    Log gUsuário, "Exclusão - Departamento - Seção: " & txtDescriçãoDepartamento & " - " & txtDescriçãoSeção
    
    StatusBarAviso = "Exclusão bem sucedida"
    BarraDeStatus StatusBarAviso
    
    If TBLDepartamentoSeção.RecordCount = 0 Then
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
    
    If TBLDepartamentoSeção.BOF Then
        TBLDepartamentoSeção.MoveFirst
    ElseIf TBLDepartamentoSeção.EOF Then
        TBLDepartamentoSeção.MoveLast
    Else
        TBLDepartamentoSeção.MovePrevious
        If TBLDepartamentoSeção.BOF Then
            TBLDepartamentoSeção.MoveNext
        End If
    End If
    
    GetRecords
    
    TestaInferior TBLDepartamentoSeção, lAllowEdit, lAllowDelete, lAllowConsult
    TestaSuperior TBLDepartamentoSeção, lAllowEdit, lAllowDelete, lAllowConsult
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
        If TBLDepartamentoSeção.RecordCount > 0 And Not TBLDepartamentoSeção.BOF And Not TBLDepartamentoSeção.EOF Then
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
    
    TestaInferior TBLDepartamentoSeção, lAllowEdit, lAllowDelete, lAllowConsult
    TestaSuperior TBLDepartamentoSeção, lAllowEdit, lAllowDelete, lAllowConsult
    
    If TBLDepartamentoSeção.RecordCount = 0 Then
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
    
    If txtCódigoDepartamento.Enabled Then
        txtCódigoDepartamento.SetFocus
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
    
    txtCódigoDepartamento.SetFocus
End Sub
Public Sub MoveFirst()
    If lAlterar Then
       If Not Cancelamento Then
            Exit Sub
        End If
    End If
    
    StatusBarAviso = "Pronto"
    BarraDeStatus StatusBarAviso
    
    TBLDepartamentoSeção.MoveFirst
    
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
    
    TBLDepartamentoSeção.MoveLast
    
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
    
    TBLDepartamentoSeção.MoveNext
    If TBLDepartamentoSeção.EOF Then
        TBLDepartamentoSeção.MovePrevious
        Exit Sub
    End If
    
    StatusBarAviso = "Pronto"
    BarraDeStatus StatusBarAviso
    
    NavegaçãoInferior lAllowConsult
    TestaSuperior TBLDepartamentoSeção, lAllowEdit, lAllowDelete, lAllowConsult
    
    GetRecords
End Sub
Public Sub MovePrevious()
    If lAlterar Then
       If Not Cancelamento Then
            Exit Sub
        End If
    End If
    
    TBLDepartamentoSeção.MovePrevious
    
    If TBLDepartamentoSeção.BOF Then
        TBLDepartamentoSeção.MoveNext
        Exit Sub
    End If
    
    StatusBarAviso = "Pronto"
    BarraDeStatus StatusBarAviso
    
    NavegaçãoSuperior lAllowConsult
    TestaInferior TBLDepartamentoSeção, lAllowEdit, lAllowDelete, lAllowConsult
    
    GetRecords
End Sub
Public Sub PosRecords()
    If TBLDepartamentoSeção.RecordCount = 0 Then
        Exit Sub
    End If
    
    TBLDepartamentoSeção.Seek "=", txtCódigoDepartamento, txtCódigoSeção
    If TBLDepartamentoSeção.NoMatch Then
        MsgBox "Não consegui encontrar " + txtCódigoDepartamento + " - " + txtCódigoSeção, vbExclamation, "Erro"
        TBLDepartamentoSeção.MoveFirst
        NavegaçãoInferior False
        NavegaçãoInferior lAllowConsult
    Else
        TestaInferior TBLDepartamentoSeção, lAllowEdit, lAllowDelete, lAllowConsult
        TestaSuperior TBLDepartamentoSeção, lAllowEdit, lAllowDelete, lAllowConsult
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
    txtCódigoDepartamento = TBLDepartamentoSeção("CÓDIGO DO DEPTO")
    TBLDepartamento.Seek "=", txtCódigoDepartamento
    txtDescriçãoDepartamento = TBLDepartamento("DESCRIÇÃO")
    txtCódigoSeção = TBLDepartamentoSeção("CÓDIGO DA SEÇÃO")
    TBLSeção.Seek "=", txtCódigoSeção
    txtDescriçãoSeção = TBLSeção("DESCRIÇÃO")
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
        TBLDepartamentoSeção.AddNew
    Else
        TBLDepartamentoSeção.Edit
    End If
    
    TBLDepartamentoSeção("CÓDIGO DO DEPTO") = txtCódigoDepartamento
    TBLDepartamentoSeção("CÓDIGO DA SEÇÃO") = txtCódigoSeção
    If lInserir Then
        TBLDepartamentoSeção("USERNAME - CRIA") = gUsuário
        TBLDepartamentoSeção("DATA - CRIA") = Date
        TBLDepartamentoSeção("HORA - CRIA") = Time
        TBLDepartamentoSeção("USERNAME - ALTERA") = "VAZIO"
        TBLDepartamentoSeção("DATA - ALTERA") = vbNull
        TBLDepartamentoSeção("HORA - ALTERA") = vbNull
    End If
    If lAlterar Then
        TBLDepartamentoSeção("USERNAME - ALTERA") = gUsuário
        TBLDepartamentoSeção("DATA - ALTERA") = Date
        TBLDepartamentoSeção("HORA - ALTERA") = Time
    End If
    TBLDepartamentoSeção.Update
        
Erro:
    If Err <> 0 Then
        TBLDepartamentoSeção.CancelUpdate
        GeraMensagemDeErro "Departamento - Seção - SetRecords - " & txtCódigoDepartamento & txtCódigoSeção, True
        SetRecords = False
        Exit Function
    End If

    WS.CommitTrans 'Grava as alterações ou inclusões se não houverem erros
    
    If lInserir Then
        Log gUsuário, "Inclusão - Departamento - Seção: " & txtDescriçãoDepartamento & " - " & txtDescriçãoSeção
    Else
        Log gUsuário, "Alteração - Departamento - Seção: " & txtDescriçãoDepartamento & " - " & txtDescriçãoSeção
    End If
    
    SetRecords = True
End Function
Private Sub ZeraCampos()
    txtCódigoDepartamento = Empty
    txtDescriçãoDepartamento = Empty
    txtCódigoSeção = Empty
    txtDescriçãoSeção = Empty
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
    If Not DepartamentoSeçãoAberto Then
        Unload Me
        Exit Sub
    End If
    If Not DepartamentoAberto Then
        Unload Me
        Exit Sub
    End If
    If Not SeçãoAberto Then
        Unload Me
        Exit Sub
    End If
    
    TestaInferior TBLDepartamentoSeção, lAllowEdit, lAllowDelete, lAllowConsult
    TestaSuperior TBLDepartamentoSeção, lAllowEdit, lAllowDelete, lAllowConsult
    
    If TBLDepartamentoSeção.RecordCount = 0 Then
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
    
    lAllowInsert = Allow("DEPARTAMENTO - SEÇÃO", "I")
    lAllowEdit = Allow("DEPARTAMENTO - SEÇÃO", "A")
    lAllowDelete = Allow("DEPARTAMENTO - SEÇÃO", "E")
    lAllowConsult = Allow("DEPARTAMENTO - SEÇÃO", "C")
    
    ZeraCampos
    
    lInserir = False
    lAlterar = False
    
    DepartamentoSeçãoAberto = AbreTabela(Dicionário, "CADASTRO", "DEPARTAMENTO - SEÇÃO", DBCadastro, TBLDepartamentoSeção, TBLTabela, dbOpenTable)
    
    If DepartamentoSeçãoAberto Then
        IndiceDepartamentoSeçãoAtivo = "DEPARTAMENTOSEÇÃO1"
        TBLDepartamentoSeção.Index = IndiceDepartamentoSeçãoAtivo
    Else
        MsgBox "Não consegui abrir a tabela 'Departamento - Seção' !", vbCritical, "Erro"
        Exit Sub
    End If
    
    DepartamentoAberto = AbreTabela(Dicionário, "CADASTRO", "DEPARTAMENTO", DBCadastro, TBLDepartamento, TBLTabela, dbOpenTable)
    
    If DepartamentoAberto Then
        IndiceDepartamentoAtivo = "DEPARTAMENTO1"
        TBLDepartamento.Index = IndiceDepartamentoAtivo
    Else
        MsgBox "Não consegui abrir a tabela 'Departamento' !", vbCritical, "Erro"
        Exit Sub
    End If
    
    SeçãoAberto = AbreTabela(Dicionário, "CADASTRO", "SEÇÃO", DBCadastro, TBLSeção, TBLTabela, dbOpenTable)
    
    If SeçãoAberto Then
        IndiceSeçãoAtivo = "SEÇÃO1"
        TBLSeção.Index = IndiceSeçãoAtivo
    Else
        MsgBox "Não consegui abrir a tabela 'Seção' !", vbCritical, "Erro"
        Exit Sub
    End If

    BotãoIncluir lAllowInsert
 
    If TBLDepartamentoSeção.RecordCount = 0 Then
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
        
    If TBLDepartamentoSeção.RecordCount = 0 Or TBLDepartamentoSeção.RecordCount = 1 Then
        NavegaçãoSuperior False
    Else
        NavegaçãoInferior lAllowConsult
    End If
    
    StatusBarAviso = "Pronto"
    Relatório = AddPath(AplicaçãoPath, "REPORT\DEPTOSEÇÃO.RPT")
    TotalDatabaseName = 1
    DataBaseName(1) = AddPath(AplicaçãoPath, "DATABASE\CADASTRO.MDB")
    mFechar = False
    Exit Sub
    
Erro:
    GeraMensagemDeErro -"Departamento - Seção - Load"
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
    
    Set frmDepartamentoSeção = Nothing
End Sub
Private Sub Form_Unload(Cancel As Integer)
    If DepartamentoSeçãoAberto Then
        TBLDepartamentoSeção.Close
    End If
    If DepartamentoAberto Then
        TBLDepartamento.Close
    End If
    If SeçãoAberto Then
        TBLSeção.Close
    End If
    If Forms.Count = 2 Then
        AllBotões False
    End If
End Sub
Private Sub txtCódigoDepartamento_Change()
    FormatMask "9999", txtCódigoDepartamento
End Sub
Private Sub txtCódigoDepartamento_KeyPress(KeyAscii As Integer)
    If Not lInserir Then
        lAlterar = True
        StatusBarAviso = "Alteração"
        BarraDeStatus StatusBarAviso
    End If
End Sub
Private Sub txtCódigoDepartamento_LostFocus()
    If mdiGeal.ActiveForm.Name = "frmDepartamentoSeção" Then
        If txtCódigoDepartamento.Enabled Then
            LeftBlank txtCódigoDepartamento
            TBLDepartamento.Seek "=", txtCódigoDepartamento
            If TBLDepartamento.NoMatch Then
                MsgBox "Não encontrei o departamento " + txtCódigoDepartamento, vbExclamation, "Aviso"
                frmEncontra.BancoDeDados = "CADASTRO"
                frmEncontra.Tabela = "DEPARTAMENTO"
                frmEncontra.Inicio = 1
                frmEncontra.Fim = 4
                frmEncontra.Caption = "Departamnento"
                frmEncontra.Show vbModal
                txtCódigoDepartamento = frmEncontra.Código
                TBLDepartamento.Seek "=", txtCódigoDepartamento
            End If
            txtDescriçãoDepartamento = TBLDepartamento("DESCRIÇÃO")
        End If
    End If
End Sub
Private Sub txtCódigoSeção_Change()
    FormatMask "9999", txtCódigoSeção
End Sub
Private Sub txtCódigoSeção_KeyPress(KeyAscii As Integer)
    If Not lInserir Then
        lAlterar = True
        StatusBarAviso = "Alteração"
        BarraDeStatus StatusBarAviso
    End If
End Sub
Private Sub txtCódigoSeção_LostFocus()
    If mdiGeal.ActiveForm.Name = "frmDepartamentoSeção" Then
        If txtCódigoSeção.Enabled Then
            LeftBlank txtCódigoSeção
            TBLSeção.Seek "=", txtCódigoSeção
            If TBLSeção.NoMatch Then
                MsgBox "Não encontrei a seção " + txtCódigoSeção, vbExclamation, "Aviso"
                frmEncontra.BancoDeDados = "CADASTRO"
                frmEncontra.Tabela = "SEÇÃO"
                frmEncontra.Inicio = 1
                frmEncontra.Fim = 4
                frmEncontra.Show vbModal
                frmEncontra.Caption = "Seção"
                txtCódigoSeção = frmEncontra.Código
                TBLSeção.Seek "=", txtCódigoSeção
            End If
            txtDescriçãoSeção = TBLSeção("DESCRIÇÃO")
        End If
    End If
End Sub
