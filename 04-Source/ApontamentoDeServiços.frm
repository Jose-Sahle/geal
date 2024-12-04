VERSION 5.00
Begin VB.Form frmApontamentoDeDespesas 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Apontamentos de Servi�os"
   ClientHeight    =   3495
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5475
   Icon            =   "ApontamentoDeServi�os.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   3495
   ScaleWidth      =   5475
   Begin VB.CommandButton cmdGravar 
      Caption         =   "&Gravar"
      Height          =   345
      Left            =   2880
      TabIndex        =   4
      Top             =   3120
      Width           =   1245
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   345
      Left            =   4200
      TabIndex        =   5
      Top             =   3120
      Width           =   1245
   End
   Begin VB.Frame frApontamentosDeDespesas 
      Height          =   3045
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   5475
      Begin VB.TextBox txtObserva��o 
         Height          =   1035
         Left            =   150
         MultiLine       =   -1  'True
         TabIndex        =   2
         Top             =   1320
         Width           =   5145
      End
      Begin VB.TextBox txtValor 
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   3630
         TabIndex        =   3
         Text            =   "99.999.999,99"
         Top             =   2580
         Width           =   1665
      End
      Begin VB.TextBox txtVencimento 
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1050
         TabIndex        =   0
         Text            =   "99/99/9999"
         Top             =   240
         Width           =   1335
      End
      Begin VB.ComboBox cmbDespesa 
         Height          =   315
         Left            =   1050
         TabIndex        =   1
         Top             =   660
         Width           =   4245
      End
      Begin VB.Label lblObserva��o 
         Caption         =   "Observa��o"
         Height          =   225
         Left            =   150
         TabIndex        =   10
         Top             =   1080
         Width           =   1305
      End
      Begin VB.Label lblValor 
         Caption         =   "Valor"
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
         Left            =   3000
         TabIndex        =   9
         Top             =   2640
         Width           =   525
      End
      Begin VB.Label lblVencimento 
         Caption         =   "Vencimento"
         Height          =   255
         Left            =   150
         TabIndex        =   8
         Top             =   300
         Width           =   945
      End
      Begin VB.Label lblServi�os 
         Caption         =   "Servi�o"
         Height          =   225
         Left            =   150
         TabIndex        =   7
         Top             =   690
         Width           =   645
      End
   End
End
Attribute VB_Name = "frmApontamentoDeDespesas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim TBLDespesas          As Table
Dim DespesasAberto       As Boolean
Dim IndiceDespesasAtivo  As String

Dim TBLTipoDeDespesas          As Table
Dim TipoDeDespesasAberto       As Boolean
Dim IndiceTipoDeDespesasAtivo  As String

Dim lPula As Boolean

Dim lInserir As Boolean
Dim lAlterar As Boolean

Dim lFechar As Boolean

Dim lAllowInsert  As Boolean
Dim lAllowEdit As Boolean
Dim lAllowDelete As Boolean
Dim lAllowConsult As Boolean

Dim C�digoDoDespesa() As Integer

Dim mUsu�rio As String
Dim mHora As Date

Dim StatusBarAviso$

Public lAtualizar As Boolean
Private Function Ascan(ByRef Matriz() As Integer, ByVal Express�o As String) As Integer
    Dim Cont As Integer
    Dim Retorno As Integer
    
    Retorno = -1
    
    For Cont = LBound(Matriz) To UBound(Matriz)
        If Matriz(Cont) = Express�o Then
            Retorno = Cont
            Exit For
        End If
    Next
    
    Ascan = Retorno
End Function
Private Sub AtivaCampos()
    frApontamentosDeDespesas.Enabled = True
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
    frApontamentosDeDespesas.Enabled = False
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
        GeraMensagemDeErro "Despesas - Excluir - " & txtVencimento & " - " & txtObserva��o, True
        StatusBarAviso = "Falha na exclus�o"
        BarraDeStatus StatusBarAviso
        Exit Sub
    End If
    
    WS.CommitTrans
    
    Log gUsu�rio, "Exclus�o - Apontamento de Despesas: " & txtVencimento & " - " & txtObserva��o
    
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
Public Sub FillDespesas()
    Dim Cont As Integer
    
    TBLTipoDeDespesas.MoveFirst
    
    cmbDespesa.Clear
    
    ReDim C�digoDoDespesa(0 To TBLTipoDeDespesas.RecordCount - 1)
    
    For Cont = 0 To TBLTipoDeDespesas.RecordCount - 1
        cmbDespesa.AddItem TBLTipoDeDespesas("DESCRI��O")
        C�digoDoDespesa(Cont) = TBLTipoDeDespesas("C�DIGO")
        TBLTipoDeDespesas.MoveNext
    Next
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
    
    If txtVencimento.Enabled Then
        txtVencimento.SetFocus
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
    
    txtVencimento.SetFocus
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
    Dim C�digo As Integer
    
    If TBLDespesas.RecordCount = 0 Then
        Exit Sub
    End If
    
    C�digo = C�digoDoDespesa(cmbDespesa.ListIndex)
    
    TBLDespesas.Seek "=", C�digo, txtVencimento, mHora, mUsu�rio
    
    If TBLDespesas.NoMatch Then
        MsgBox "N�o consegui encontrar o Despesa", vbExclamation, "Erro"
        TBLDespesas.MoveFirst
        Navega��oInferior False
        Navega��oInferior lAllowConsult
    Else
        TestaInferior TBLDespesas, lAllowEdit, lAllowDelete, lAllowConsult
        TestaSuperior TBLDespesas, lAllowEdit, lAllowDelete, lAllowConsult
    End If
    
    GetRecords
End Sub
Private Sub GetRecords()
    Dim Pos As Integer
    
    lPula = True
    
    If Not lAllowConsult Then
        ZeraCampos
        DesativaCampos
        lPula = False
        Exit Sub
    End If
    
    If TBLDespesas("DATA DO VENCIMENTO") <> vbNull Then
        txtVencimento = FormatStringMask(CheckDataMask, TBLDespesas("DATA DO VENCIMENTO"))
        CorrigeData DataMask, txtVencimento, TBLDespesas("DATA DO VENCIMENTO")
    Else
        txtVencimento = DataNula
    End If
    
    Pos = Ascan(C�digoDoDespesa, TBLDespesas("C�DIGO DA DESPESA"))
    cmbDespesa.ListIndex = Pos
    
    txtObserva��o = TBLDespesas("OBSERVA��O")
    
    txtValor = TBLDespesas("VALOR DA DESPESA")
    lPula = True
    txtValor_LostFocus
    lPula = False
    
    mHora = TBLDespesas("HORA")
    mUsu�rio = TBLDespesas("USERNAME - CRIA")
    
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
    
    TBLDespesas("DATA DO VENCIMENTO") = IIf(Trim(StrTran(txtVencimento, "/")) <> Empty, txtVencimento, vbNull)
    TBLDespesas("C�DIGO DA DESPESA") = C�digoDoDespesa(cmbDespesa.ListIndex)
    TBLDespesas("OBSERVA��O") = txtObserva��o
    TBLDespesas("VALOR DA DESPESA") = ValStr(txtValor)
    
    If lInserir Then
        mHora = Time
        mUsu�rio = gUsu�rio
        TBLDespesas("HORA") = mHora
        TBLDespesas("USERNAME - CRIA") = gUsu�rio
        TBLDespesas("DATA - CRIA") = Date
        TBLDespesas("HORA - CRIA") = Time
        TBLDespesas("USERNAME - ALTERA") = vbNull
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
        Log gUsu�rio, "Inclus�o - Apontamento de Despesas: " & cmbDespesa.Text & " - " & txtVencimento
    Else
        Log gUsu�rio, "Altera��o - Apontamento de Despesas: " & cmbDespesa.Text & " - " & txtVencimento
    End If
    
    Exit Function
    
Erro:
    TBLDespesas.CancelUpdate
    GeraMensagemDeErro "Apontamento de Despesas - SetRecords - " & cmbDespesa.Text & " - " & txtVencimento, True
    SetRecords = False
    On Error GoTo 0
End Function
Private Sub ZeraCampos()
    lPula = True
    txtVencimento = "  /  /  "
    txtVencimento_LostFocus
    lPula = True
    cmbDespesa.ListIndex = 0
    txtObserva��o = Empty
    txtValor = Empty
    txtValor_LostFocus
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
    
    If Not TipoDeDespesasAberto Then
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
        
    lInserir = False
    lAlterar = False
    lPula = False
    
    TipoDeDespesasAberto = AbreTabela(Dicion�rio, "CADASTRO", "DESPESAS", DBCadastro, TBLTipoDeDespesas, TBLTabela, dbOpenTable)
    
    If TipoDeDespesasAberto Then
        IndiceTipoDeDespesasAtivo = "DESPESAS1"
        TBLTipoDeDespesas.Index = IndiceTipoDeDespesasAtivo
    Else
        MsgBox "N�o consegui abrir a tabela 'Cadastro - Despesas' !", vbCritical, "Erro"
        Exit Sub
    End If
    
    DespesasAberto = AbreTabela(Dicion�rio, "FINANCEIRO", "DESPESA", DBFinanceiro, TBLDespesas, TBLTabela, dbOpenTable)
    
    If DespesasAberto Then
        IndiceDespesasAtivo = "DESPESA1"
        TBLDespesas.Index = IndiceDespesasAtivo
    Else
        MsgBox "N�o consegui abrir a tabela 'Financeiro - Despesas' !", vbCritical, "Erro"
        Exit Sub
    End If
    
    FillDespesas
    
    ZeraCampos
    
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
    
    lFechar = False
    
    Exit Sub
    
Erro:
    GeraMensagemDeErro "Apontamento de Despesas - Load"
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
    
    Set frmApontamentoDeDespesas = Nothing
End Sub
Private Sub Form_Unload(Cancel As Integer)
    If DespesasAberto Then
        TBLDespesas.Close
    End If
    
    If TipoDeDespesasAberto Then
        TBLTipoDeDespesas.Close
    End If
    
    If Forms.Count = 2 Then
        AllBot�es False
    End If
End Sub
Private Sub txtObserva��o_KeyPress(KeyAscii As Integer)
    If Not lInserir Then
        lAlterar = True
        StatusBarAviso = "Altera��o"
        BarraDeStatus StatusBarAviso
    End If
End Sub
Private Sub txtValor_Change()
    If Not lPula Then
        FormatMask "@K 99.999.999,99", txtValor
    End If
End Sub
Private Sub txtValor_LostFocus()
    lPula = True
    FormatMask "@V ##.###.##0,00", txtValor
    lPula = False
End Sub
Private Sub txtVencimento_Change()
    If Not lPula Then
        lPula = True
        FormatMask DataMask, txtVencimento
        lPula = False
    End If
End Sub
Private Sub txtVencimento_KeyPress(KeyAscii As Integer)
    If Not lInserir Then
        lAlterar = True
        StatusBarAviso = "Altera��o"
        BarraDeStatus StatusBarAviso
    End If
End Sub
Private Sub txtVencimento_LostFocus()
    If StrTran(txtVencimento.Text, "/") <> Space(6) Then
        lPula = True
        CorrigeData DataMask, txtVencimento, Date
        lPula = False
        If Not FormatMask(CheckDataMask, txtVencimento) Then
            Beep
            MsgBox "Data inv�lida !", vbCritical, "Erro"
            txtVencimento.SelStart = 0
            txtVencimento.SetFocus
        End If
    End If
End Sub

