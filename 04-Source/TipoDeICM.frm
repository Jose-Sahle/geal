VERSION 5.00
Begin VB.Form frmTipoDeICM 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tipo de ICM"
   ClientHeight    =   2145
   ClientLeft      =   1605
   ClientTop       =   2790
   ClientWidth     =   6480
   Icon            =   "TipoDeICM.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2145
   ScaleWidth      =   6480
   Begin VB.Frame frICM 
      Height          =   1635
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   6465
      Begin VB.TextBox txtCódigoDoPDV 
         Height          =   285
         Left            =   5820
         TabIndex        =   3
         Top             =   1200
         Width           =   375
      End
      Begin VB.TextBox txtICM 
         Height          =   285
         Left            =   1200
         TabIndex        =   2
         Text            =   " 0,00"
         Top             =   1200
         Width           =   555
      End
      Begin VB.TextBox txtCódigo 
         Height          =   285
         Left            =   1200
         TabIndex        =   0
         Top             =   300
         Width           =   750
      End
      Begin VB.TextBox txtDescrição 
         Height          =   285
         Left            =   1200
         TabIndex        =   1
         Top             =   750
         Width           =   5000
      End
      Begin VB.Label CódigoDoPDV 
         Caption         =   "Código do PDV"
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
         Left            =   4350
         TabIndex        =   10
         Top             =   1260
         Width           =   1605
      End
      Begin VB.Label lblICM 
         Caption         =   "ICM"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Left            =   180
         TabIndex        =   9
         Top             =   1260
         Width           =   420
      End
      Begin VB.Label lblCódigo 
         Caption         =   "Código"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   200
         Left            =   150
         TabIndex        =   8
         Top             =   330
         Width           =   855
      End
      Begin VB.Label lblDescrição 
         Caption         =   "Descrição"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   150
         TabIndex        =   7
         Top             =   780
         Width           =   885
      End
   End
   Begin VB.CommandButton cmdGravar 
      Caption         =   "&Gravar"
      Height          =   345
      Left            =   3885
      TabIndex        =   4
      Top             =   1740
      Width           =   1245
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   345
      Left            =   5205
      TabIndex        =   5
      Top             =   1740
      Width           =   1245
   End
End
Attribute VB_Name = "frmTipoDeICM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim TBLTipoDeICM As Table
Dim TipoDeICMAberto As Boolean
Dim IndiceTipoDeICMAtivo$

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
    frICM.Enabled = True
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
    
    If TBLTipoDeICM.RecordCount = 0 Then
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
    
    TestaInferior TBLTipoDeICM, lAllowEdit, lAllowDelete, lAllowConsult
    TestaSuperior TBLTipoDeICM, lAllowEdit, lAllowDelete, lAllowConsult
    
    GetRecords
End Function
Private Sub DesativaCampos()
    BotãoImprimir False
    frICM.Enabled = False
    cmdGravar.Enabled = False
    cmdCancelar.Enabled = False
    BotãoGravar False
End Sub
Public Sub Encontrar()
    If Not lAllowConsult Then
        Exit Sub
    End If
    Set frmEncontrar.DBBancoDeDados = DBCadastro
    frmEncontrar.NomeDaJanela = "Tipo de ICM"
    frmEncontrar.LabelDescription = "Descrição"
    frmEncontrar.Mensagem = "Nenhum Tipo de ICM foi selecionado!"
    frmEncontrar.BancoDeDados = "CADASTRO"
    frmEncontrar.Tabela = "TIPO DE ICM"
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
    
    TBLTipoDeICM.Delete
    
    If Err <> 0 Then
        GeraMensagemDeErro "TipoDeICM - Excluir", True
        StatusBarAviso = "Falha na exclusão"
        BarraDeStatus StatusBarAviso
        Exit Sub
    End If
    
    WS.CommitTrans
    
    Log gUsuário, "Exclusão - Tipo de ICM: " & txtCódigo & " - " & txtDescrição
    
    StatusBarAviso = "Exclusão bem sucedida"
    BarraDeStatus StatusBarAviso
    
    If TBLTipoDeICM.RecordCount = 0 Then
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
    
    If TBLTipoDeICM.BOF Then
        TBLTipoDeICM.MoveFirst
    ElseIf TBLTipoDeICM.EOF Then
        TBLTipoDeICM.MoveLast
    Else
        TBLTipoDeICM.MovePrevious
        If TBLTipoDeICM.BOF Then
            TBLTipoDeICM.MoveNext
        End If
    End If
    
    GetRecords
    
    TestaInferior TBLTipoDeICM, lAllowEdit, lAllowDelete, lAllowConsult
    TestaSuperior TBLTipoDeICM, lAllowEdit, lAllowDelete, lAllowConsult
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
        If TBLTipoDeICM.RecordCount > 0 And Not TBLTipoDeICM.BOF And Not TBLTipoDeICM.EOF Then
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
    
    TestaInferior TBLTipoDeICM, lAllowEdit, lAllowDelete, lAllowConsult
    TestaSuperior TBLTipoDeICM, lAllowEdit, lAllowDelete, lAllowConsult
    
    If TBLTipoDeICM.RecordCount = 0 Then
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
    
    TBLTipoDeICM.MoveFirst
    
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
    
    TBLTipoDeICM.MoveLast
    
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
    
    TBLTipoDeICM.MoveNext
    If TBLTipoDeICM.EOF Then
        TBLTipoDeICM.MovePrevious
        Exit Sub
    End If
    
    StatusBarAviso = "Pronto"
    BarraDeStatus StatusBarAviso
    
    NavegaçãoInferior lAllowConsult
    TestaSuperior TBLTipoDeICM, lAllowEdit, lAllowDelete, lAllowConsult
    
    GetRecords
End Sub
Public Sub MovePrevious()
    If lAlterar Then
       If Not Cancelamento Then
            Exit Sub
        End If
    End If
    
    TBLTipoDeICM.MovePrevious
    If TBLTipoDeICM.BOF Then
        TBLTipoDeICM.MoveNext
        Exit Sub
    End If
    
    StatusBarAviso = "Pronto"
    BarraDeStatus StatusBarAviso
    
    NavegaçãoSuperior lAllowConsult
    TestaInferior TBLTipoDeICM, lAllowEdit, lAllowDelete, lAllowConsult
    
    GetRecords
End Sub
Public Sub PosRecords()
    If TBLTipoDeICM.RecordCount = 0 Then
        Exit Sub
    End If
    
    TBLTipoDeICM.Seek "=", txtCódigo
    If TBLTipoDeICM.NoMatch Then
        MsgBox "Não consegui encontrar " + txtCódigo, vbExclamation, "Erro"
        TBLTipoDeICM.MoveFirst
        NavegaçãoInferior False
        NavegaçãoInferior lAllowConsult
    Else
        TestaInferior TBLTipoDeICM, lAllowEdit, lAllowDelete, lAllowConsult
        TestaSuperior TBLTipoDeICM, lAllowEdit, lAllowDelete, lAllowConsult
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
    
    txtCódigo = TBLTipoDeICM("CÓDIGO")
    txtDescrição = TBLTipoDeICM("DESCRIÇÃO")
    txtICM = StrVal(TBLTipoDeICM("ICM"))
    txtICM_LostFocus
    txtCódigoDoPDV = TBLTipoDeICM("CÓDIGO DO PDV")
    
    If Not lAllowEdit Then
        DesativaCampos
    End If
    
    Exit Sub
    
Erro:
    GeraMensagemDeErro "Tipo de ICM - GetRecords"
    Resume Next
End Sub
Private Function SetRecords()
    On Error GoTo Erro
    
    Dim Msg$
    Dim Confirmação As Integer, Msg1$, Msg2$
    
    WS.BeginTrans 'Inicia uma Transação
    
    If lInserir Then
        TBLTipoDeICM.AddNew
    Else
        TBLTipoDeICM.Edit
    End If
    
    TBLTipoDeICM("CÓDIGO") = txtCódigo
    TBLTipoDeICM("DESCRIÇÃO") = txtDescrição
    TBLTipoDeICM("ICM") = ValStr(txtICM)
    TBLTipoDeICM("CÓDIGO DO PDV") = txtCódigoDoPDV
    If lInserir Then
        TBLTipoDeICM("USERNAME - CRIA") = gUsuário
        TBLTipoDeICM("DATA - CRIA") = Date
        TBLTipoDeICM("HORA - CRIA") = Time
        TBLTipoDeICM("USERNAME - ALTERA") = "VAZIO"
        TBLTipoDeICM("DATA - ALTERA") = vbNull
        TBLTipoDeICM("HORA - ALTERA") = vbNull
    End If
    If lAlterar Then
        TBLTipoDeICM("USERNAME - ALTERA") = gUsuário
        TBLTipoDeICM("DATA - ALTERA") = Date
        TBLTipoDeICM("HORA - ALTERA") = Time
    End If
    TBLTipoDeICM.Update
        
Erro:
    If Err <> 0 Then
        TBLTipoDeICM.CancelUpdate
        GeraMensagemDeErro "TipoDeICM - SetRecords", True
        SetRecords = False
        Exit Function
    End If

    WS.CommitTrans 'Grava as alterações ou inclusões se não houverem erros
    
    If lInserir Then
        Log gUsuário, "Inclusão - Tipo de ICM: " & txtCódigo & " - " & txtDescrição
    Else
        Log gUsuário, "Alteração - Tipo de ICM: " & txtCódigo & " - " & txtDescrição
    End If
    
    SetRecords = True
End Function
Private Sub ZeraCampos()
    txtCódigo = Empty
    txtDescrição = Empty
    txtICM = "0,00"
    txtICM_LostFocus
    txtCódigoDoPDV = Empty
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
    If Not TipoDeICMAberto Then
        Unload Me
        Exit Sub
    End If
    
    TestaInferior TBLTipoDeICM, lAllowEdit, lAllowDelete, lAllowConsult
    TestaSuperior TBLTipoDeICM, lAllowEdit, lAllowDelete, lAllowConsult
    
    If TBLTipoDeICM.RecordCount = 0 Then
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
    
    lAllowInsert = Allow("TIPO DE ICM", "I")
    lAllowEdit = Allow("TIPO DE ICM", "A")
    lAllowDelete = Allow("TIPO DE ICM", "E")
    lAllowConsult = Allow("TIPO DE ICM", "C")
    
    ZeraCampos
    
    lPula = False
    lInserir = False
    lAlterar = False
    
    TipoDeICMAberto = AbreTabela(Dicionário, "CADASTRO", "TIPO DE ICM", DBCadastro, TBLTipoDeICM, TBLTabela, dbOpenTable)
    
    If TipoDeICMAberto Then
        IndiceTipoDeICMAtivo = "TIPODEICM1"
        TBLTipoDeICM.Index = IndiceTipoDeICMAtivo
    Else
        MsgBox "Não consegui abrir a tabela 'Tipo De Embalagem' !", vbCritical, "Erro"
        Exit Sub
    End If
    
    BotãoIncluir lAllowInsert
 
    If TBLTipoDeICM.RecordCount = 0 Then
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
        
    If TBLTipoDeICM.RecordCount = 0 Or TBLTipoDeICM.RecordCount = 1 Then
        NavegaçãoSuperior False
    Else
        NavegaçãoInferior lAllowConsult
    End If
    
    StatusBarAviso = "Pronto"
    Relatório = AddPath(AplicaçãoPath, "REPORT\TIPO DE ICM.RPT")
    TotalDatabaseName = 1
    DataBaseName(1) = AddPath(AplicaçãoPath, "DATABASE\CADASTRO.MDB")
    mFechar = False
    Exit Sub
    
Erro:
    GeraMensagemDeErro "TipoDeICM - Load"
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
    
    Set frmTipoDeICM = Nothing
End Sub
Private Sub Form_Unload(Cancel As Integer)
    If TipoDeICMAberto Then
        TBLTipoDeICM.Close
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
Private Sub txtICM_Change()
    FormatMask "@K 99,99", txtICM
End Sub
Private Sub txtICM_KeyPress(KeyAscii As Integer)
    If Not lInserir Then
        lAlterar = True
        StatusBarAviso = "Alteração"
        BarraDeStatus StatusBarAviso
    End If
End Sub
Private Sub txtICM_LostFocus()
    FormatMask "@V #0,00", txtICM
End Sub

