VERSION 5.00
Begin VB.Form frmDespesas 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Despesas"
   ClientHeight    =   2115
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6150
   Icon            =   "Servi�os.frx":0000
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
      Begin VB.TextBox txtDescri��o 
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
      Begin VB.TextBox txtC�digo 
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
      Begin VB.Label lblDescri��o 
         Caption         =   "Descri��o"
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
      Begin VB.Label lblC�digo 
         Caption         =   "C�digo"
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
Public Relat�rio$
Public TotalDatabaseName%

Public lAtualizar As Boolean
Private Sub AtivaCampos()
    frDespesas.Enabled = True
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
    
    If TBLDespesas.RecordCount = 0 Then
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
    
    TestaInferior TBLDespesas, lAllowEdit, lAllowDelete, lAllowConsult
    TestaSuperior TBLDespesas, lAllowEdit, lAllowDelete, lAllowConsult
    
    GetRecords
End Function
Private Sub DesativaCampos()
    frDespesas.Enabled = False
    cmdGravar.Enabled = False
    cmdCancelar.Enabled = False
    Bot�oGravar False
End Sub
Public Sub Encontrar()
    If Not lAllowConsult Then
        Exit Sub
    End If
    Set frmEncontrar.DBBancoDeDados = DBCadastro
    frmEncontrar.NomeDaJanela = "Despesas"
    frmEncontrar.LabelDescription = "Descri��o"
    frmEncontrar.Mensagem = "Nenhuma Despesa foi selecionado!"
    frmEncontrar.BancoDeDados = "CADASTRO"
    frmEncontrar.Tabela = "DESPESAS"
    frmEncontrar.Indice = "1"
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
    
    TBLDespesas.Delete
    
    If Err <> 0 Then
        GeraMensagemDeErro "Despesas - Excluir - " & txtDescri��o, True
        StatusBarAviso = "Falha na exclus�o"
        BarraDeStatus StatusBarAviso
        Exit Sub
    End If
    
    WS.CommitTrans
    
    Log gUsu�rio, "Exclus�o - Despesas: " & txtC�digo & " - " & txtDescri��o
    
    StatusBarAviso = "Exclus�o bem sucedida"
    BarraDeStatus StatusBarAviso
    
    If TBLDespesas.RecordCount = 0 Then
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
            StatusBarAviso = "Inclus�o bem sucedida"
        Else
            StatusBarAviso = "Falha na inclus�o"
        End If
    Else
        If TBLDespesas.RecordCount > 0 And Not TBLDespesas.BOF And Not TBLDespesas.EOF Then
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
    
    TestaInferior TBLDespesas, lAllowEdit, lAllowDelete, lAllowConsult
    TestaSuperior TBLDespesas, lAllowEdit, lAllowDelete, lAllowConsult
    
    If TBLDespesas.RecordCount = 0 Then
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
    
    TBLDespesas.MoveFirst
    
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
    
    TBLDespesas.MoveLast
    
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
    
    TBLDespesas.MoveNext
    If TBLDespesas.EOF Then
        TBLDespesas.MovePrevious
        Exit Sub
    End If
    
    StatusBarAviso = "Pronto"
    BarraDeStatus StatusBarAviso
    
    Navega��oInferior lAllowConsult
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
    
    Navega��oSuperior lAllowConsult
    TestaInferior TBLDespesas, lAllowEdit, lAllowDelete, lAllowConsult
    
    GetRecords
End Sub
Public Sub PosRecords()
    If TBLDespesas.RecordCount = 0 Then
        Exit Sub
    End If
    
    TBLDespesas.Seek "=", txtC�digo
    If TBLDespesas.NoMatch Then
        MsgBox "N�o consegui encontrar " + txtC�digo, vbExclamation, "Erro"
        TBLDespesas.MoveFirst
        Navega��oInferior False
        Navega��oInferior lAllowConsult
    Else
        TestaInferior TBLDespesas, lAllowEdit, lAllowDelete, lAllowConsult
        TestaSuperior TBLDespesas, lAllowEdit, lAllowDelete, lAllowConsult
    End If
    GetRecords
End Sub
Public Function PushDataBaseName(ByVal Posi��o As Integer) As String
    PushDataBaseName = DataBaseName(Posi��o)
End Function
Private Sub GetRecords()
    lPula = True
    If Not lAllowConsult Then
        ZeraCampos
        DesativaCampos
        lPula = False
        Exit Sub
    End If
    
    txtC�digo = TBLDespesas("C�DIGO")
    txtDescri��o = TBLDespesas("DESCRI��O")
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
    Dim Confirma��o As Integer, Msg1$, Msg2$
    
    WS.BeginTrans 'Inicia uma Transa��o
    
    If lInserir Then
        TBLDespesas.AddNew
    Else
        TBLDespesas.Edit
    End If
    
    TBLDespesas("C�DIGO") = txtC�digo
    TBLDespesas("DESCRI��O") = txtDescri��o
    TBLDespesas("DATA") = txtData
    TBLDespesas("TIPO") = Trim(Str(cmbTipo.ListIndex + 1))
    If lInserir Then
        TBLDespesas("USERNAME - CRIA") = gUsu�rio
        TBLDespesas("DATA - CRIA") = Date
        TBLDespesas("HORA - CRIA") = Time
        TBLDespesas("USERNAME - ALTERA") = "VAZIO"
        TBLDespesas("DATA - ALTERA") = vbNull
        TBLDespesas("HORA - ALTERA") = vbNull
    End If
    If lAlterar Then
        TBLDespesas("USERNAME - ALTERA") = gUsu�rio
        TBLDespesas("DATA - ALTERA") = Date
        TBLDespesas("HORA - ALTERA") = Time
    End If
    TBLDespesas.Update
            
    WS.CommitTrans 'Grava as altera��es ou inclus�es se n�o houverem erros
    
    SetRecords = True
    
    If lInserir Then
        Log gUsu�rio, "Inclus�o - Despesas: " & txtC�digo & " - " & txtDescri��o
    Else
        Log gUsu�rio, "Altera��o - Despesas: " & txtC�digo & " - " & txtDescri��o
    End If
    
    Exit Function
    
Erro:
    TBLDespesas.CancelUpdate
    GeraMensagemDeErro "Despesas - SetRecords - " & txtDescri��o, True
    SetRecords = False
    On Error GoTo 0
End Function
Private Sub ZeraCampos()
    lPula = True
    txtC�digo = Empty
    txtDescri��o = Empty
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
    
    lAllowInsert = Allow("DESPESAS", "I")
    lAllowEdit = Allow("DESPESAS", "A")
    lAllowDelete = Allow("DESPESAS", "E")
    lAllowConsult = Allow("DESPESAS", "C")
    
    cmbTipo.Clear
    
    cmbTipo.AddItem "1-Despesa mensal obrigat�ria com data fixa"
    cmbTipo.AddItem "2-Despesa mensal obrigat�ria sem data fixa"
    cmbTipo.AddItem "3-Despesa mensal n�o obrigat�ria"
        
    cmbTipo.ListIndex = 0
    
    ZeraCampos
    
    lInserir = False
    lAlterar = False
    lPula = False
    
    DespesasAberto = AbreTabela(Dicion�rio, "CADASTRO", "DESPESAS", DBCadastro, TBLDespesas, TBLTabela, dbOpenTable)
    
    If DespesasAberto Then
        IndiceDespesasAtivo = "DESPESAS1"
        TBLDespesas.Index = IndiceDespesasAtivo
    Else
        MsgBox "N�o consegui abrir a tabela 'Despesas' !", vbCritical, "Erro"
        Exit Sub
    End If
    
    Bot�oIncluir lAllowInsert
 
    If TBLDespesas.RecordCount = 0 Then
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
        
    If TBLDespesas.RecordCount = 0 Or TBLDespesas.RecordCount = 1 Then
        Navega��oSuperior False
    Else
        Navega��oInferior lAllowConsult
    End If
    
    StatusBarAviso = "Pronto"
    Relat�rio = AddPath(Aplica��oPath, "REPORT\Despesas.RPT")
    TotalDatabaseName = 1
    DataBaseName(1) = AddPath(Aplica��oPath, "DATABASE\CADASTRO.MDB")
    
    lFechar = False
    
    Exit Sub
    
Erro:
    GeraMensagemDeErro "Despesas - Load"
    lFechar = True
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
    
    Set frmDespesas = Nothing
End Sub
Private Sub Form_Unload(Cancel As Integer)
    If DespesasAberto Then
        TBLDespesas.Close
    End If
    If Forms.Count = 2 Then
        AllBot�es False
    End If
End Sub
Private Sub txtC�digo_Change()
    If lPula Then
        Exit Sub
    End If
    FormatMask "9999", txtC�digo
End Sub
Private Sub txtC�digo_KeyPress(KeyAscii As Integer)
    If Not lInserir Then
        lAlterar = True
        StatusBarAviso = "Altera��o"
        BarraDeStatus StatusBarAviso
    End If
End Sub
Private Sub txtC�digo_LostFocus()
    lPula = True
    LeftBlank txtC�digo
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
        StatusBarAviso = "Altera��o"
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
            MsgBox "Data inv�lida !", vbCritical, "Erro"
            txtData.SelStart = 0
            txtData.SetFocus
        End If
    End If
End Sub
Private Sub txtDescri��o_Change()
    If lPula Then
        Exit Sub
    End If
    FormatMask "@!S30", txtDescri��o
End Sub
Private Sub txtDescri��o_KeyPress(KeyAscii As Integer)
    If Not lInserir Then
        lAlterar = True
        StatusBarAviso = "Altera��o"
        BarraDeStatus StatusBarAviso
    End If
End Sub

