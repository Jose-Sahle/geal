VERSION 5.00
Begin VB.Form frmLocal 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Localidade do Produto"
   ClientHeight    =   3195
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6525
   Icon            =   "Local.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   3195
   ScaleWidth      =   6525
   Begin VB.Frame frLocal 
      Height          =   2760
      Left            =   0
      TabIndex        =   10
      Top             =   0
      Width           =   6525
      Begin VB.TextBox txtC�digo 
         Height          =   285
         Left            =   1200
         TabIndex        =   0
         Top             =   300
         Width           =   315
      End
      Begin VB.TextBox txtEndere�o 
         Height          =   285
         Left            =   1200
         TabIndex        =   1
         Top             =   690
         Width           =   5235
      End
      Begin VB.TextBox txtBairro 
         Height          =   285
         Left            =   1200
         TabIndex        =   2
         Top             =   1080
         Width           =   5235
      End
      Begin VB.TextBox txtCidade 
         Height          =   285
         Left            =   1200
         TabIndex        =   3
         Top             =   1470
         Width           =   5235
      End
      Begin VB.TextBox txtUF 
         Height          =   285
         Left            =   1200
         TabIndex        =   4
         Top             =   1860
         Width           =   435
      End
      Begin VB.TextBox txtCep 
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
         Left            =   4380
         TabIndex        =   5
         Text            =   "  .   -   "
         Top             =   1860
         Width           =   1300
      End
      Begin VB.TextBox txtTelefone 
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
         Left            =   1200
         TabIndex        =   6
         Text            =   "(    )    -    "
         Top             =   2250
         Width           =   1900
      End
      Begin VB.TextBox txtFax 
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
         Left            =   4380
         TabIndex        =   7
         Text            =   "(    )    -    "
         Top             =   2250
         Width           =   1900
      End
      Begin VB.Label lblC�digo 
         Caption         =   "C�digo"
         Height          =   195
         Left            =   150
         TabIndex        =   18
         Top             =   330
         Width           =   1065
      End
      Begin VB.Label lblEndere�o 
         Caption         =   "Endere�o"
         Height          =   195
         Left            =   150
         TabIndex        =   17
         Top             =   720
         Width           =   720
      End
      Begin VB.Label lblBairro 
         Caption         =   "Bairro"
         Height          =   225
         Left            =   150
         TabIndex        =   16
         Top             =   1110
         Width           =   945
      End
      Begin VB.Label lblCidade 
         Caption         =   "Cidade"
         Height          =   225
         Left            =   150
         TabIndex        =   15
         Top             =   1500
         Width           =   945
      End
      Begin VB.Label lblUF 
         Caption         =   "U. F."
         Height          =   225
         Left            =   150
         TabIndex        =   14
         Top             =   1890
         Width           =   405
      End
      Begin VB.Label lblCep 
         Caption         =   "CEP"
         Height          =   195
         Left            =   3930
         TabIndex        =   13
         Top             =   1890
         Width           =   375
      End
      Begin VB.Label lblTelefone 
         Caption         =   "Telefone"
         Height          =   210
         Left            =   150
         TabIndex        =   12
         Top             =   2280
         Width           =   660
      End
      Begin VB.Label lblFax 
         Caption         =   "Fax"
         Height          =   195
         Left            =   3930
         TabIndex        =   11
         Top             =   2280
         Width           =   345
      End
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   345
      Left            =   5250
      TabIndex        =   9
      Top             =   2820
      Width           =   1245
   End
   Begin VB.CommandButton cmdGravar 
      Caption         =   "&Gravar"
      Default         =   -1  'True
      Height          =   345
      Left            =   3930
      TabIndex        =   8
      Top             =   2820
      Width           =   1245
   End
End
Attribute VB_Name = "frmLocal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim TBLLocal As Table
Dim LocalAberto As Boolean
Dim IndiceLocalAtivo$

Dim lAllowInsert  As Boolean
Dim lAllowEdit    As Boolean
Dim lAllowDelete  As Boolean
Dim lAllowConsult As Boolean

Dim lPula As Boolean
Dim lInserir As Boolean
Dim lAlterar As Boolean

Dim StatusBarAviso$

Public lAtualizar As Boolean
Private Sub AtivaCampos()
    frLocal.Enabled = True
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
    
    If TBLLocal.RecordCount = 0 Then
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
    
    TestaInferior TBLLocal, lAllowEdit, lAllowDelete, lAllowConsult
    TestaSuperior TBLLocal, lAllowEdit, lAllowDelete, lAllowConsult
        
    GetRecords
End Function
Private Sub DesativaCampos()
    frLocal.Enabled = False
    cmdGravar.Enabled = False
    cmdCancelar.Enabled = False
    Bot�oGravar False
End Sub
Public Sub Encontrar()
    If Not lAllowConsult Then
        Exit Sub
    End If
    Set frmEncontrar.DBBancoDeDados = DBCadastro
    frmEncontrar.NomeDaJanela = "Localidade"
    frmEncontrar.LabelDescription = "Endere�o"
    frmEncontrar.Mensagem = "Nenhuma localidade foi selecionado!"
    frmEncontrar.BancoDeDados = "CADASTRO"
    frmEncontrar.Tabela = "LOCAL DO PRODUTO"
    frmEncontrar.Indice = "2"
    frmEncontrar.CampoChave = "C�DIGO"
    frmEncontrar.CampoPreencheLista = "ENDERE�O"
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
    
    TBLLocal.Delete
    
    If Err <> 0 Then
        GeraMensagemDeErro "Local - Excluir - " & txtEndere�o, True
        StatusBarAviso = "Falha na exclus�o"
        BarraDeStatus StatusBarAviso
        Exit Sub
    End If
    
    WS.CommitTrans
    
    Log gUsu�rio, "Exclus�o - Produto: " & txtC�digo & " - " & txtEndere�o
    
    StatusBarAviso = "Exclus�o bem sucedida"
    BarraDeStatus StatusBarAviso
    
    If TBLLocal.RecordCount = 0 Then
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
    
    If TBLLocal.BOF Then
        TBLLocal.MoveFirst
    ElseIf TBLLocal.EOF Then
        TBLLocal.MoveLast
    Else
        TBLLocal.MovePrevious
        If TBLLocal.BOF Then
            TBLLocal.MoveNext
        End If
    End If
    
    GetRecords
    
    TestaInferior TBLLocal, lAllowEdit, lAllowDelete, lAllowConsult
    TestaSuperior TBLLocal, lAllowEdit, lAllowDelete, lAllowConsult
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
        If TBLLocal.RecordCount > 0 And Not TBLLocal.BOF And Not TBLLocal.EOF Then
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
    
    TestaInferior TBLLocal, lAllowEdit, lAllowDelete, lAllowConsult
    TestaSuperior TBLLocal, lAllowEdit, lAllowDelete, lAllowConsult
    
    If TBLLocal.RecordCount = 0 Then
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
    
    TBLLocal.MoveFirst
    
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
    
    TBLLocal.MoveLast
    
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
    
    TBLLocal.MoveNext
    If TBLLocal.EOF Then
        TBLLocal.MovePrevious
        Exit Sub
    End If
    
    StatusBarAviso = "Pronto"
    BarraDeStatus StatusBarAviso
    
    Navega��oInferior lAllowConsult
    TestaSuperior TBLLocal, lAllowEdit, lAllowDelete, lAllowConsult
    
    GetRecords
End Sub
Public Sub MovePrevious()
    If lAlterar Then
       If Not Cancelamento Then
            Exit Sub
        End If
    End If
    
    TBLLocal.MovePrevious
    If TBLLocal.BOF Then
        TBLLocal.MoveNext
        Exit Sub
    End If
    
    StatusBarAviso = "Pronto"
    BarraDeStatus StatusBarAviso
    
    Navega��oSuperior lAllowConsult
    TestaInferior TBLLocal, lAllowEdit, lAllowDelete, lAllowConsult
    
    GetRecords
End Sub
Public Sub PosRecords()
    If TBLLocal.RecordCount = 0 Then
        Exit Sub
    End If
    
    TBLLocal.Seek "=", txtC�digo
    If TBLLocal.NoMatch Then
        MsgBox "N�o consegui encontrar " + txtC�digo, vbExclamation, "Erro"
        TBLLocal.MoveFirst
        Navega��oInferior False
        Navega��oInferior lAllowConsult
    Else
        TestaInferior TBLLocal, lAllowEdit, lAllowDelete, lAllowConsult
        TestaSuperior TBLLocal, lAllowEdit, lAllowDelete, lAllowConsult
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
    txtC�digo = TBLLocal("C�DIGO")
    txtEndere�o = TBLLocal("ENDERE�O")
    txtBairro = TBLLocal("BAIRRO")
    txtCidade = TBLLocal("CIDADE")
    txtUF = TBLLocal("UF")
    txtTelefone = TBLLocal("TELEFONE")
    txtFax = TBLLocal("FAX")
    txtCep = TBLLocal("CEP")
    If Not lAllowEdit Then
        DesativaCampos
    End If
    
    Exit Sub
    
Erro:
    GeraMensagemDeErro "Local - GetRecords "
    Resume Next
End Sub
Private Function SetRecords()
    On Error GoTo Erro
    
    Dim Msg$
    Dim Confirma��o As Integer, Msg1$, Msg2$, AchouDepartamentoSe��o As Boolean
    
    WS.BeginTrans 'Inicia uma Transa��o
    
    If lInserir Then
        TBLLocal.AddNew
    Else
        TBLLocal.Edit
    End If
    
    TBLLocal("C�DIGO") = txtC�digo
    TBLLocal("ENDERE�O") = txtEndere�o
    TBLLocal("BAIRRO") = txtBairro
    TBLLocal("CIDADE") = txtCidade
    TBLLocal("UF") = txtUF
    TBLLocal("TELEFONE") = txtTelefone
    TBLLocal("FAX") = txtFax
    TBLLocal("CEP") = txtCep
    If lInserir Then
        TBLLocal("USERNAME - CRIA") = gUsu�rio
        TBLLocal("DATA - CRIA") = Date
        TBLLocal("HORA - CRIA") = Time
        TBLLocal("USERNAME - ALTERA") = "VAZIO"
        TBLLocal("DATA - ALTERA") = vbNull
        TBLLocal("HORA - ALTERA") = vbNull
    End If
    If lAlterar Then
        TBLLocal("USERNAME - ALTERA") = gUsu�rio
        TBLLocal("DATA - ALTERA") = Date
        TBLLocal("HORA - ALTERA") = Time
    End If
    TBLLocal.Update
        
Erro:
    If Err <> 0 Then
        TBLLocal.CancelUpdate
        GeraMensagemDeErro "Local - SetRecords - " & txtEndere�o, True
        SetRecords = False
        Exit Function
    End If

    WS.CommitTrans 'Grava as altera��es ou inclus�es se n�o houverem erros
        
    If lInserir Then
        Log gUsu�rio, "Inclus�o - Local " & txtC�digo & " - " & txtEndere�o
    Else
        Log gUsu�rio, "Alera��o - Local " & txtC�digo & " - " & txtEndere�o
    End If
    
    SetRecords = True
End Function
Private Sub ZeraCampos()
    txtC�digo = Empty
    txtEndere�o = Empty
    txtBairro = Empty
    txtCidade = Empty
    txtUF = Empty
    txtTelefone = Empty
    txtFax = Empty
    txtCep = Empty
End Sub
Private Sub cmdCancelar_Click()
    Cancelamento
End Sub
Private Sub cmdGravar_Click()
    Gravar
End Sub
Private Sub Form_Activate()
    If Not LocalAberto Then
        Unload frmLocal
        Exit Sub
    End If
    TestaInferior TBLLocal, lAllowEdit, lAllowDelete, lAllowConsult
    TestaSuperior TBLLocal, lAllowEdit, lAllowDelete, lAllowConsult
    If TBLLocal.RecordCount = 0 Then
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
    ZeraCampos
    
    lPula = False
    lInserir = False
    lAlterar = False
    
    lAllowInsert = Allow("LOCALIDADE DE ESTOQUE", "I")
    lAllowEdit = Allow("LOCALIDADE DE ESTOQUE", "A")
    lAllowDelete = Allow("LOCALIDADE DE ESTOQUE", "E")
    lAllowConsult = Allow("LOCALIDADE DE ESTOQUE", "C")
    
    LocalAberto = AbreTabela(Dicion�rio, "CADASTRO", "LOCAL DO PRODUTO", DBCadastro, TBLLocal, TBLTabela, dbOpenTable)
    
    If LocalAberto Then
        IndiceLocalAtivo = "LOCALDOPRODUTO1"
        TBLLocal.Index = IndiceLocalAtivo
    Else
        MsgBox "N�o consegui abrir a tabela 'LOCAL' !", vbCritical, "Erro"
        Exit Sub
    End If
    
    Bot�oIncluir lAllowInsert
 
    If TBLLocal.RecordCount = 0 Then
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
        
    If TBLLocal.RecordCount = 0 Or TBLLocal.RecordCount = 1 Then
        Navega��oSuperior False
    Else
        Navega��oInferior lAllowConsult
    End If
    
    StatusBarAviso = "Pronto"
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
    
    Set frmLocal = Nothing
End Sub
Private Sub Form_Unload(Cancel As Integer)
    If LocalAberto Then
        TBLLocal.Close
    End If
    If Forms.Count = 2 Then
        AllBot�es False
    End If
End Sub
Private Sub txtBairro_Change()
    If Not lPula Then
        FormatMask "@!S30", txtBairro
    End If
End Sub
Private Sub txtCep_Change()
    If Not lPula Then
        FormatMask "99.999-999", txtCep
    End If
End Sub
Private Sub txtCidade_Change()
    If Not lPula Then
        FormatMask "@!S30", txtCidade
    End If
End Sub
Private Sub txtC�digo_Change()
    If Not lPula Then
        FormatMask "99", txtC�digo
    End If
End Sub
Private Sub txtC�digo_LostFocus()
    FormatMask "@N 00", txtC�digo
End Sub
Private Sub txtEndere�o_Change()
    If Not lPula Then
        FormatMask "@S40", txtEndere�o
    End If
End Sub
Private Sub txtFax_Change()
    If Not lPula Then
        FormatMask "(####)####-####", txtFax
    End If
End Sub
Private Sub txtTelefone_Change()
    If Not lPula Then
        FormatMask "(####)####-####", txtTelefone
    End If
End Sub
Private Sub txtUF_Change()
    If Not lPula Then
        FormatMask "@! AA", txtUF
    End If
End Sub
