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
      Begin VB.TextBox txtDescri��o 
         Height          =   285
         Left            =   1200
         TabIndex        =   1
         Top             =   720
         Width           =   1635
      End
      Begin VB.TextBox txtC�digo 
         Height          =   285
         Left            =   1200
         TabIndex        =   0
         Top             =   300
         Width           =   315
      End
      Begin VB.Label lblDescri��o 
         Caption         =   "Descri��o"
         Height          =   195
         Left            =   150
         TabIndex        =   6
         Top             =   750
         Width           =   765
      End
      Begin VB.Label lblC�digo 
         Caption         =   "C�digo"
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
    
    If TBLUnidades.RecordCount = 0 Then
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
    
    TestaInferior TBLUnidades, lAllowEdit, lAllowDelete, lAllowConsult
    TestaSuperior TBLUnidades, lAllowEdit, lAllowDelete, lAllowConsult
    
    GetRecords
End Function
Private Sub DesativaCampos()
    frUnidades.Enabled = False
    cmdGravar.Enabled = False
    cmdCancelar.Enabled = False
    Bot�oGravar False
End Sub
Public Sub Encontrar()
    If Not lAllowConsult Then
        Exit Sub
    End If
    Set frmEncontrar.DBBancoDeDados = DBCadastro
    frmEncontrar.NomeDaJanela = "Unidades"
    frmEncontrar.LabelDescription = "Descri��o"
    frmEncontrar.Mensagem = "Nenhuma unidade foi selecionado!"
    frmEncontrar.BancoDeDados = "CADASTRO"
    frmEncontrar.Tabela = "UNIDADES"
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
    
    TBLUnidades.Delete
    
    If Err <> 0 Then
        GeraMensagemDeErro "Unidades - Excluir - " & txtDescri��o, True
        StatusBarAviso = "Falha na exclus�o"
        BarraDeStatus StatusBarAviso
        Exit Sub
    End If
    
    WS.CommitTrans
    
    Log gUsu�rio, "Exclus�o - Unidades: " & txtC�digo & " - " & txtDescri��o
    
    StatusBarAviso = "Exclus�o bem sucedida"
    BarraDeStatus StatusBarAviso
    
    If TBLUnidades.RecordCount = 0 Then
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
            StatusBarAviso = "Inclus�o bem sucedida"
        Else
            StatusBarAviso = "Falha na inclus�o"
        End If
    Else
        If TBLUnidades.RecordCount > 0 And Not TBLUnidades.BOF And Not TBLUnidades.EOF Then
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
    
    TestaInferior TBLUnidades, lAllowEdit, lAllowDelete, lAllowConsult
    TestaSuperior TBLUnidades, lAllowEdit, lAllowDelete, lAllowConsult
    
    If TBLUnidades.RecordCount = 0 Then
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
    
    TBLUnidades.MoveFirst
    
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
    
    TBLUnidades.MoveLast
    
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
    
    TBLUnidades.MoveNext
    If TBLUnidades.EOF Then
        TBLUnidades.MovePrevious
        Exit Sub
    End If
    
    StatusBarAviso = "Pronto"
    BarraDeStatus StatusBarAviso
    
    Navega��oInferior lAllowConsult
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
    
    Navega��oSuperior lAllowConsult
    TestaInferior TBLUnidades, lAllowEdit, lAllowDelete, lAllowConsult
    
    GetRecords
End Sub
Public Sub PosRecords()
    If TBLUnidades.RecordCount = 0 Then
        Exit Sub
    End If
    
    TBLUnidades.Seek "=", txtC�digo
    If TBLUnidades.NoMatch Then
        MsgBox "N�o consegui encontrar " + txtC�digo, vbExclamation, "Erro"
        TBLUnidades.MoveFirst
        Navega��oInferior False
        Navega��oInferior lAllowConsult
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
    txtC�digo = TBLUnidades("C�DIGO")
    txtDescri��o = TBLUnidades("DESCRI��O")
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
    Dim Confirma��o As Integer, Msg1$, Msg2$, AchouDepartamentoSe��o As Boolean
    
    WS.BeginTrans 'Inicia uma Transa��o
        
    If lInserir Then
        TBLUnidades.AddNew
    Else
        TBLUnidades.Edit
    End If
    
    TBLUnidades("C�DIGO") = txtC�digo
    TBLUnidades("DESCRI��O") = txtDescri��o
    If lInserir Then
        TBLUnidades("USERNAME - CRIA") = gUsu�rio
        TBLUnidades("DATA - CRIA") = Date
        TBLUnidades("HORA - CRIA") = Time
        TBLUnidades("USERNAME - ALTERA") = "VAZIO"
        TBLUnidades("DATA - ALTERA") = vbNull
        TBLUnidades("HORA - ALTERA") = vbNull
    End If
    If lAlterar Then
        TBLUnidades("USERNAME - ALTERA") = gUsu�rio
        TBLUnidades("DATA - ALTERA") = Date
        TBLUnidades("HORA - ALTERA") = Time
    End If
    TBLUnidades.Update
        
Erro:
    If Err <> 0 Then
        TBLUnidades.CancelUpdate
        GeraMensagemDeErro "Unidades - SetRecords - " & txtDescri��o, True
        SetRecords = False
        Exit Function
    End If

    WS.CommitTrans 'Grava as altera��es ou inclus�es se n�o houverem erros
        
    If lInserir Then
        Log gUsu�rio, "Inclus�o - Unidades " & txtC�digo & " - " & txtDescri��o
    Else
        Log gUsu�rio, "Altera��o - Unidades " & txtC�digo & " - " & txtDescri��o
    End If
    
    SetRecords = True
End Function
Private Sub ZeraCampos()
    txtC�digo = Empty
    txtDescri��o = Empty
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
        Bot�oGravar False
        cmdGravar.Enabled = False
        cmdCancelar.Enabled = False
    Else
        Bot�oGravar (lInserir Or lAllowEdit)
        cmdGravar.Enabled = (lInserir Or lAllowEdit)
        cmdCancelar.Enabled = (lInserir Or lAllowEdit)
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
    
    UnidadesAberto = AbreTabela(Dicion�rio, "CADASTRO", "UNIDADES", DBCadastro, TBLUnidades, TBLTabela, dbOpenTable)
    
    If UnidadesAberto Then
        IndiceUnidadesAtivo = "UNIDADES1"
        TBLUnidades.Index = IndiceUnidadesAtivo
    Else
        MsgBox "N�o consegui abrir a tabela 'Unidades' !", vbCritical, "Erro"
        Exit Sub
    End If
    
    Bot�oIncluir lAllowInsert
 
    If TBLUnidades.RecordCount = 0 Then
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
        
    If TBLUnidades.RecordCount = 0 Or TBLUnidades.RecordCount = 1 Then
        Navega��oSuperior False
    Else
        Navega��oInferior lAllowConsult
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
    
    Set frmUnidades = Nothing
End Sub
Private Sub Form_Unload(Cancel As Integer)
    If UnidadesAberto Then
        TBLUnidades.Close
    End If
    If Forms.Count = 2 Then
        AllBot�es False
    End If
End Sub
Private Sub txtC�digo_Change()
    If Not lPula Then
        FormatMask "99", txtC�digo
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
    LeftZero txtC�digo
End Sub
Private Sub txtDescri��o_Change()
    If Not lPula Then
        FormatMask "@!S10", txtDescri��o
    End If
End Sub
Private Sub txtDescri��o_KeyPress(KeyAscii As Integer)
    If Not lInserir Then
        lAlterar = True
        StatusBarAviso = "Altera��o"
        BarraDeStatus StatusBarAviso
    End If
End Sub

