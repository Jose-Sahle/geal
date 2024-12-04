VERSION 5.00
Begin VB.Form frmTipoDeEmbalagem 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tipo de Embalagem"
   ClientHeight    =   1695
   ClientLeft      =   1650
   ClientTop       =   2145
   ClientWidth     =   6480
   Icon            =   "TipoDeEmbalagem.frx":0000
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
      TabIndex        =   7
      Top             =   1320
      Width           =   1245
   End
   Begin VB.CommandButton cmdGravar 
      Caption         =   "&Gravar"
      Height          =   345
      Left            =   3885
      TabIndex        =   6
      Top             =   1320
      Width           =   1245
   End
   Begin VB.Frame frEmbalagem 
      Height          =   1275
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   6465
      Begin VB.TextBox txtAbreviado 
         Height          =   285
         Left            =   5220
         TabIndex        =   1
         Top             =   300
         Width           =   975
      End
      Begin VB.TextBox txtDescri��o 
         Height          =   285
         Left            =   1200
         MaxLength       =   30
         TabIndex        =   2
         Top             =   750
         Width           =   5000
      End
      Begin VB.TextBox txtC�digo 
         Height          =   285
         Left            =   1200
         TabIndex        =   0
         Top             =   300
         Width           =   750
      End
      Begin VB.Label lblDescri��o 
         Caption         =   "Descri��o"
         Height          =   195
         Left            =   150
         TabIndex        =   5
         Top             =   780
         Width           =   885
      End
      Begin VB.Label lblC�digo 
         Caption         =   "C�digo"
         Height          =   200
         Left            =   150
         TabIndex        =   4
         Top             =   330
         Width           =   855
      End
   End
End
Attribute VB_Name = "frmTipoDeEmbalagem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim TBLTipoDeEmbalagem As Table
Dim TipoDeEmbalagemAberto As Boolean
Dim IndiceTipoDeEmbalagemAtivo$

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
Public Relat�rio$
Public TotalDatabaseName%

Public lAtualizar As Boolean
Private Sub AtivaCampos()
    Bot�oImprimir True
    frEmbalagem.Enabled = True
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
    
    If TBLTipoDeEmbalagem.RecordCount = 0 Then
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
    
    TestaInferior TBLTipoDeEmbalagem, lAllowEdit, lAllowDelete, lAllowConsult
    TestaSuperior TBLTipoDeEmbalagem, lAllowEdit, lAllowDelete, lAllowConsult
    
    GetRecords
End Function
Private Sub DesativaCampos()
    Bot�oImprimir False
    frEmbalagem.Enabled = False
    cmdGravar.Enabled = False
    cmdCancelar.Enabled = False
    Bot�oGravar False
End Sub
Public Sub Encontrar()
    If Not lAllowConsult Then
        Exit Sub
    End If
    Set frmEncontrar.DBBancoDeDados = DBCadastro
    frmEncontrar.NomeDaJanela = "Tipo de Embalagem"
    frmEncontrar.LabelDescription = "Descri��o"
    frmEncontrar.Mensagem = "Nenhum Tipo de Embalagem foi selecionado!"
    frmEncontrar.BancoDeDados = "CADASTRO"
    frmEncontrar.Tabela = "TIPO DE EMBALAGEM"
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
    
    TBLTipoDeEmbalagem.Delete
    
    If Err <> 0 Then
        GeraMensagemDeErro "TipoDeEmbalagem - Excluir", True
        StatusBarAviso = "Falha na exclus�o"
        BarraDeStatus StatusBarAviso
        Exit Sub
    End If
    
    WS.CommitTrans
    
    Log gUsu�rio, "Exclus�o - Tipo de Embalagem: " & txtC�digo & " - " & txtDescri��o
    
    StatusBarAviso = "Exclus�o bem sucedida"
    BarraDeStatus StatusBarAviso
    
    If TBLTipoDeEmbalagem.RecordCount = 0 Then
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
    
    If TBLTipoDeEmbalagem.BOF Then
        TBLTipoDeEmbalagem.MoveFirst
    ElseIf TBLTipoDeEmbalagem.EOF Then
        TBLTipoDeEmbalagem.MoveLast
    Else
        TBLTipoDeEmbalagem.MovePrevious
        If TBLTipoDeEmbalagem.BOF Then
            TBLTipoDeEmbalagem.MoveNext
        End If
    End If
    
    GetRecords
    
    TestaInferior TBLTipoDeEmbalagem, lAllowEdit, lAllowDelete, lAllowConsult
    TestaSuperior TBLTipoDeEmbalagem, lAllowEdit, lAllowDelete, lAllowConsult
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
        If TBLTipoDeEmbalagem.RecordCount > 0 And Not TBLTipoDeEmbalagem.BOF And Not TBLTipoDeEmbalagem.EOF Then
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
    
    TestaInferior TBLTipoDeEmbalagem, lAllowEdit, lAllowDelete, lAllowConsult
    TestaSuperior TBLTipoDeEmbalagem, lAllowEdit, lAllowDelete, lAllowConsult
    
    If TBLTipoDeEmbalagem.RecordCount = 0 Then
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
    
    TBLTipoDeEmbalagem.MoveFirst
    
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
    
    TBLTipoDeEmbalagem.MoveLast
    
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
    
    TBLTipoDeEmbalagem.MoveNext
    If TBLTipoDeEmbalagem.EOF Then
        TBLTipoDeEmbalagem.MovePrevious
        Exit Sub
    End If
    
    StatusBarAviso = "Pronto"
    BarraDeStatus StatusBarAviso
    
    Navega��oInferior lAllowConsult
    TestaSuperior TBLTipoDeEmbalagem, lAllowEdit, lAllowDelete, lAllowConsult
    
    GetRecords
End Sub
Public Sub MovePrevious()
    If lAlterar Then
       If Not Cancelamento Then
            Exit Sub
        End If
    End If
    
    TBLTipoDeEmbalagem.MovePrevious
    If TBLTipoDeEmbalagem.BOF Then
        TBLTipoDeEmbalagem.MoveNext
        Exit Sub
    End If
    
    StatusBarAviso = "Pronto"
    BarraDeStatus StatusBarAviso
    
    Navega��oSuperior lAllowConsult
    TestaInferior TBLTipoDeEmbalagem, lAllowEdit, lAllowDelete, lAllowConsult
    
    GetRecords
End Sub
Public Sub PosRecords()
    If TBLTipoDeEmbalagem.RecordCount = 0 Then
        Exit Sub
    End If
    
    TBLTipoDeEmbalagem.Seek "=", txtC�digo
    If TBLTipoDeEmbalagem.NoMatch Then
        MsgBox "N�o consegui encontrar " + txtC�digo, vbExclamation, "Erro"
        TBLTipoDeEmbalagem.MoveFirst
        Navega��oInferior False
        Navega��oInferior lAllowConsult
    Else
        TestaInferior TBLTipoDeEmbalagem, lAllowEdit, lAllowDelete, lAllowConsult
        TestaSuperior TBLTipoDeEmbalagem, lAllowEdit, lAllowDelete, lAllowConsult
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
    txtC�digo = TBLTipoDeEmbalagem("C�DIGO")
    txtDescri��o = TBLTipoDeEmbalagem("DESCRI��O")
    txtAbreviado = TBLTipoDeEmbalagem("ABREVIADO")
    If Not lAllowEdit Then
        DesativaCampos
    End If
    
    Exit Sub
    
Erro:
    GeraMensagemDeErro "Tipo de Embalagem - GetRecords "
    Resume Next
End Sub
Private Function SetRecords()
    On Error GoTo Erro
    
    Dim Msg$
    Dim Confirma��o As Integer, Msg1$, Msg2$
    
    WS.BeginTrans 'Inicia uma Transa��o
    
    If lInserir Then
        TBLTipoDeEmbalagem.AddNew
    Else
        TBLTipoDeEmbalagem.Edit
    End If
    
    TBLTipoDeEmbalagem("C�DIGO") = txtC�digo
    TBLTipoDeEmbalagem("DESCRI��O") = txtDescri��o
    TBLTipoDeEmbalagem("ABREVIADO") = txtAbreviado
    If lInserir Then
        TBLTipoDeEmbalagem("USERNAME - CRIA") = gUsu�rio
        TBLTipoDeEmbalagem("DATA - CRIA") = Date
        TBLTipoDeEmbalagem("HORA - CRIA") = Time
        TBLTipoDeEmbalagem("USERNAME - ALTERA") = "VAZIO"
        TBLTipoDeEmbalagem("DATA - ALTERA") = vbNull
        TBLTipoDeEmbalagem("HORA - ALTERA") = vbNull
    End If
    If lAlterar Then
        TBLTipoDeEmbalagem("USERNAME - ALTERA") = gUsu�rio
        TBLTipoDeEmbalagem("DATA - ALTERA") = Date
        TBLTipoDeEmbalagem("HORA - ALTERA") = Time
    End If
    TBLTipoDeEmbalagem.Update
        
Erro:
    If Err <> 0 Then
        TBLTipoDeEmbalagem.CancelUpdate
        GeraMensagemDeErro "TipoDeEmbalagem - SetRecords", True
        SetRecords = False
        Exit Function
    End If

    WS.CommitTrans 'Grava as altera��es ou inclus�es se n�o houverem erros
    
    If lInserir Then
        Log gUsu�rio, "Inclus�o - Tipo de Embalagem: " & txtC�digo & " - " & txtDescri��o
    Else
        Log gUsu�rio, "Altera��o - Tipo de Embalagem: " & txtC�digo & " - " & txtDescri��o
    End If
    
    SetRecords = True
End Function
Private Sub ZeraCampos()
    txtC�digo = Empty
    txtDescri��o = Empty
    txtAbreviado = Empty
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
    If Not TipoDeEmbalagemAberto Then
        Unload Me
        Exit Sub
    End If
    
    TestaInferior TBLTipoDeEmbalagem, lAllowEdit, lAllowDelete, lAllowConsult
    TestaSuperior TBLTipoDeEmbalagem, lAllowEdit, lAllowDelete, lAllowConsult
    
    If TBLTipoDeEmbalagem.RecordCount = 0 Then
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
    
    lAllowInsert = Allow("TIPO DE EMBALAGEM", "I")
    lAllowEdit = Allow("TIPO DE EMBALAGEM", "A")
    lAllowDelete = Allow("TIPO DE EMBALAGEM", "E")
    lAllowConsult = Allow("TIPO DE EMBALAGEM", "C")
    
    
    ZeraCampos
    
    lPula = False
    lInserir = False
    lAlterar = False
    
    TipoDeEmbalagemAberto = AbreTabela(Dicion�rio, "CADASTRO", "TIPO DE EMBALAGEM", DBCadastro, TBLTipoDeEmbalagem, TBLTabela, dbOpenTable)
    
    If TipoDeEmbalagemAberto Then
        IndiceTipoDeEmbalagemAtivo = "TIPODEEMBALAGEM1"
        TBLTipoDeEmbalagem.Index = IndiceTipoDeEmbalagemAtivo
    Else
        MsgBox "N�o consegui abrir a tabela 'Tipo De Embalagem' !", vbCritical, "Erro"
        Exit Sub
    End If
    
    Bot�oIncluir lAllowInsert
 
    If TBLTipoDeEmbalagem.RecordCount = 0 Then
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
        
    If TBLTipoDeEmbalagem.RecordCount = 0 Or TBLTipoDeEmbalagem.RecordCount = 1 Then
        Navega��oSuperior False
    Else
        Navega��oInferior lAllowConsult
    End If
    
    StatusBarAviso = "Pronto"
    Relat�rio = AddPath(Aplica��oPath, "REPORT\TIPO DE EMBALAGEM.RPT")
    TotalDatabaseName = 1
    DataBaseName(1) = AddPath(Aplica��oPath, "DATABASE\CADASTRO.MDB")
    mFechar = False
    Exit Sub
    
Erro:
    GeraMensagemDeErro "TipoDeEmbalagem - Load"
    mFechar = True
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
    
    Set frmTipoDeEmbalagem = Nothing
End Sub
Private Sub Form_Unload(Cancel As Integer)
    If TipoDeEmbalagemAberto Then
        TBLTipoDeEmbalagem.Close
    End If
    If Forms.Count = 2 Then
        AllBot�es False
    End If
End Sub
Private Sub txtC�digo_Change()
    If Not lPula Then
        FormatMask "9999", txtC�digo
    End If
End Sub
Private Sub txtC�digo_KeyPress(KeyAscii As Integer)
    If Not lInserir Then
        lAlterar = True
        StatusBarAviso = "Altera��o"
        BarraDeStatus StatusBarAviso
    End If
End Sub
Private Sub txtC�digo_LostFocus()
    LeftBlank txtC�digo
End Sub
Private Sub txtDescri��o_Change()
    If Not lPula Then
        FormatMask "@!S30", txtDescri��o
    End If
End Sub
Private Sub txtDescri��o_KeyPress(KeyAscii As Integer)
    If Not lInserir Then
        lAlterar = True
        StatusBarAviso = "Altera��o"
        BarraDeStatus StatusBarAviso
    End If
End Sub
