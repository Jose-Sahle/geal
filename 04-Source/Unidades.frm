VERSION 5.00
Begin VB.Form frmUnidades 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Unidades"
   ClientHeight    =   1770
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3240
   Icon            =   "Unidades.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   1770
   ScaleWidth      =   3240
   Begin VB.CommandButton cmdGravar 
      Caption         =   "&Gravar"
      Height          =   345
      Left            =   660
      TabIndex        =   2
      Top             =   1320
      Width           =   1245
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Enabled         =   0   'False
      Height          =   345
      Left            =   1980
      TabIndex        =   3
      Top             =   1320
      Width           =   1245
   End
   Begin VB.Frame frUnidades 
      Height          =   1215
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   3225
      Begin VB.TextBox txtDescrição 
         Height          =   285
         Left            =   1200
         TabIndex        =   1
         Top             =   720
         Width           =   1635
      End
      Begin VB.TextBox txtCódigo 
         Height          =   285
         Left            =   1200
         TabIndex        =   0
         Top             =   300
         Width           =   315
      End
      Begin VB.Label lblDescrição 
         Caption         =   "Descrição"
         Height          =   195
         Left            =   150
         TabIndex        =   6
         Top             =   750
         Width           =   765
      End
      Begin VB.Label lblCódigo 
         Caption         =   "Código"
         Height          =   195
         Left            =   150
         TabIndex        =   5
         Top             =   330
         Width           =   615
      End
   End
End
Attribute VB_Name = "frmUnidades"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim TBLUnidades As Table
Dim UnidadesAberto As Boolean
Dim IndiceUnidadesAtivo$

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
    frUnidades.Enabled = True
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
    
    If TBLUnidades.RecordCount = 0 Then
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
    
    TestaInferior TBLUnidades, lAllowEdit, lAllowDelete, lAllowConsult
    TestaSuperior TBLUnidades, lAllowEdit, lAllowDelete, lAllowConsult
    
    GetRecords
End Function
Private Sub DesativaCampos()
    frUnidades.Enabled = False
    cmdGravar.Enabled = False
    cmdCancelar.Enabled = False
    BotãoGravar False
End Sub
Public Sub Encontrar()
    If Not lAllowConsult Then
        Exit Sub
    End If
    Set frmEncontrar.DBBancoDeDados = DBCadastro
    frmEncontrar.NomeDaJanela = "Unidades"
    frmEncontrar.LabelDescription = "Descrição"
    frmEncontrar.Mensagem = "Nenhuma unidade foi selecionado!"
    frmEncontrar.BancoDeDados = "CADASTRO"
    frmEncontrar.Tabela = "UNIDADES"
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
    
    TBLUnidades.Delete
    
    If Err <> 0 Then
        GeraMensagemDeErro "Unidades - Excluir - " & txtDescrição, True
        StatusBarAviso = "Falha na exclusão"
        BarraDeStatus StatusBarAviso
        Exit Sub
    End If
    
    WS.CommitTrans
    
    Log gUsuário, "Exclusão - Unidades: " & txtCódigo & " - " & txtDescrição
    
    StatusBarAviso = "Exclusão bem sucedida"
    BarraDeStatus StatusBarAviso
    
    If TBLUnidades.RecordCount = 0 Then
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
    
    If TBLUnidades.BOF Then
        TBLUnidades.MoveFirst
    ElseIf TBLUnidades.EOF Then
        TBLUnidades.MoveLast
    Else
        TBLUnidades.MovePrevious
        If TBLUnidades.BOF Then
            TBLUnidades.MoveNext
        End If
    End If
    
    GetRecords
    
    TestaInferior TBLUnidades, lAllowEdit, lAllowDelete, lAllowConsult
    TestaSuperior TBLUnidades, lAllowEdit, lAllowDelete, lAllowConsult
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
        If TBLUnidades.RecordCount > 0 And Not TBLUnidades.BOF And Not TBLUnidades.EOF Then
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
    
    TestaInferior TBLUnidades, lAllowEdit, lAllowDelete, lAllowConsult
    TestaSuperior TBLUnidades, lAllowEdit, lAllowDelete, lAllowConsult
    
    If TBLUnidades.RecordCount = 0 Then
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
    
    TBLUnidades.MoveFirst
    
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
    
    TBLUnidades.MoveLast
    
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
    
    TBLUnidades.MoveNext
    If TBLUnidades.EOF Then
        TBLUnidades.MovePrevious
        Exit Sub
    End If
    
    StatusBarAviso = "Pronto"
    BarraDeStatus StatusBarAviso
    
    NavegaçãoInferior lAllowConsult
    TestaSuperior TBLUnidades, lAllowEdit, lAllowDelete, lAllowConsult
    
    GetRecords
End Sub
Public Sub MovePrevious()
    If lAlterar Then
       If Not Cancelamento Then
            Exit Sub
        End If
    End If
    
    TBLUnidades.MovePrevious
    If TBLUnidades.BOF Then
        TBLUnidades.MoveNext
        Exit Sub
    End If
    
    StatusBarAviso = "Pronto"
    BarraDeStatus StatusBarAviso
    
    NavegaçãoSuperior lAllowConsult
    TestaInferior TBLUnidades, lAllowEdit, lAllowDelete, lAllowConsult
    
    GetRecords
End Sub
Public Sub PosRecords()
    If TBLUnidades.RecordCount = 0 Then
        Exit Sub
    End If
    
    TBLUnidades.Seek "=", txtCódigo
    If TBLUnidades.NoMatch Then
        MsgBox "Não consegui encontrar " + txtCódigo, vbExclamation, "Erro"
        TBLUnidades.MoveFirst
        NavegaçãoInferior False
        NavegaçãoInferior lAllowConsult
    Else
        TestaInferior TBLUnidades, lAllowEdit, lAllowDelete, lAllowConsult
        TestaSuperior TBLUnidades, lAllowEdit, lAllowDelete, lAllowConsult
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
    txtCódigo = TBLUnidades("CÓDIGO")
    txtDescrição = TBLUnidades("DESCRIÇÃO")
    If Not lAllowEdit Then
        DesativaCampos
    End If
    
    Exit Sub
    
Erro:
    GeraMensagemDeErro "Unidades - GetRecords "
    Resume Next
End Sub
Private Function SetRecords()
    On Error GoTo Erro
    
    Dim Msg$
    Dim Confirmação As Integer, Msg1$, Msg2$, AchouDepartamentoSeção As Boolean
    
    WS.BeginTrans 'Inicia uma Transação
        
    If lInserir Then
        TBLUnidades.AddNew
    Else
        TBLUnidades.Edit
    End If
    
    TBLUnidades("CÓDIGO") = txtCódigo
    TBLUnidades("DESCRIÇÃO") = txtDescrição
    If lInserir Then
        TBLUnidades("USERNAME - CRIA") = gUsuário
        TBLUnidades("DATA - CRIA") = Date
        TBLUnidades("HORA - CRIA") = Time
        TBLUnidades("USERNAME - ALTERA") = "VAZIO"
        TBLUnidades("DATA - ALTERA") = vbNull
        TBLUnidades("HORA - ALTERA") = vbNull
    End If
    If lAlterar Then
        TBLUnidades("USERNAME - ALTERA") = gUsuário
        TBLUnidades("DATA - ALTERA") = Date
        TBLUnidades("HORA - ALTERA") = Time
    End If
    TBLUnidades.Update
        
Erro:
    If Err <> 0 Then
        TBLUnidades.CancelUpdate
        GeraMensagemDeErro "Unidades - SetRecords - " & txtDescrição, True
        SetRecords = False
        Exit Function
    End If

    WS.CommitTrans 'Grava as alterações ou inclusões se não houverem erros
        
    If lInserir Then
        Log gUsuário, "Inclusão - Unidades " & txtCódigo & " - " & txtDescrição
    Else
        Log gUsuário, "Alteração - Unidades " & txtCódigo & " - " & txtDescrição
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
    If Not UnidadesAberto Then
        Unload Me
        Exit Sub
    End If
    TestaInferior TBLUnidades, lAllowEdit, lAllowDelete, lAllowConsult
    TestaSuperior TBLUnidades, lAllowEdit, lAllowDelete, lAllowConsult
    If TBLUnidades.RecordCount = 0 Then
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
    On Error GoTo Erro
    
    ZeraCampos
    
    lAllowInsert = Allow("UNIDADES", "I")
    lAllowEdit = Allow("UNIDADES", "A")
    lAllowDelete = Allow("UNIDADES", "E")
    lAllowConsult = Allow("UNIDADES", "C")
    
    lPula = False
    lInserir = False
    lAlterar = False
    
    UnidadesAberto = AbreTabela(Dicionário, "CADASTRO", "UNIDADES", DBCadastro, TBLUnidades, TBLTabela, dbOpenTable)
    
    If UnidadesAberto Then
        IndiceUnidadesAtivo = "UNIDADES1"
        TBLUnidades.Index = IndiceUnidadesAtivo
    Else
        MsgBox "Não consegui abrir a tabela 'Unidades' !", vbCritical, "Erro"
        Exit Sub
    End If
    
    BotãoIncluir lAllowInsert
 
    If TBLUnidades.RecordCount = 0 Then
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
        
    If TBLUnidades.RecordCount = 0 Or TBLUnidades.RecordCount = 1 Then
        NavegaçãoSuperior False
    Else
        NavegaçãoInferior lAllowConsult
    End If
    
    StatusBarAviso = "Pronto"
    mFechar = False
    Exit Sub
    
Erro:
    GeraMensagemDeErro "Unidades - Load"
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
    
    Set frmUnidades = Nothing
End Sub
Private Sub Form_Unload(Cancel As Integer)
    If UnidadesAberto Then
        TBLUnidades.Close
    End If
    If Forms.Count = 2 Then
        AllBotões False
    End If
End Sub
Private Sub txtCódigo_Change()
    If Not lPula Then
        FormatMask "99", txtCódigo
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
    LeftZero txtCódigo
End Sub
Private Sub txtDescrição_Change()
    If Not lPula Then
        FormatMask "@!S10", txtDescrição
    End If
End Sub
Private Sub txtDescrição_KeyPress(KeyAscii As Integer)
    If Not lInserir Then
        lAlterar = True
        StatusBarAviso = "Alteração"
        BarraDeStatus StatusBarAviso
    End If
End Sub

