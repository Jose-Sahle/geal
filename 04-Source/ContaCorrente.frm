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
   Begin VB.Frame frAg�ncia 
      Caption         =   "Ag�ncia"
      Height          =   1350
      Left            =   0
      TabIndex        =   9
      Top             =   1350
      Width           =   6330
      Begin VB.TextBox txtC�digoAg�ncia 
         Height          =   285
         Left            =   1200
         TabIndex        =   1
         Top             =   300
         Width           =   1035
      End
      Begin VB.TextBox txtDescri��oAg�ncia 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1200
         TabIndex        =   10
         Top             =   750
         Width           =   5000
      End
      Begin VB.Label lblC�digoAg�ncia 
         Caption         =   "C�digo"
         Height          =   210
         Left            =   150
         TabIndex        =   12
         Top             =   330
         Width           =   660
      End
      Begin VB.Label lblDescri��oAg�ncia 
         Caption         =   "Descri��o"
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
      Begin VB.TextBox txtDescri��oBanco 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1200
         TabIndex        =   8
         Top             =   750
         Width           =   5000
      End
      Begin VB.TextBox txtC�digoBanco 
         Height          =   285
         Left            =   1200
         TabIndex        =   0
         Top             =   300
         Width           =   700
      End
      Begin VB.Label lblDescri��oBanco 
         Caption         =   "Descri��o"
         Height          =   180
         Left            =   150
         TabIndex        =   7
         Top             =   780
         Width           =   960
      End
      Begin VB.Label lblC�digoBanco 
         Caption         =   "C�digo"
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

Dim TBLAg�ncia As Table
Dim Ag�nciaAberto As Boolean

Dim TBLContaCorrente As Table
Dim ContaCorrenteAberto As Boolean

Dim IndiceAtivoBanco$, IndiceAtivoAg�ncia$, IndiceAtivoContaCorrente$

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
    frAg�ncia.Enabled = True
    frContaCorrente.Enabled = True
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
    
    If TBLContaCorrente.RecordCount = 0 Then
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
    
    TestaInferior TBLContaCorrente, lAllowEdit, lAllowDelete, lAllowConsult
    TestaSuperior TBLContaCorrente, lAllowEdit, lAllowDelete, lAllowConsult
    
    GetRecords
End Function
Private Sub DesativaCampos()
    frBanco.Enabled = False
    frAg�ncia.Enabled = False
    frContaCorrente.Enabled = False
    cmdGravar.Enabled = False
    cmdCancelar.Enabled = False
    Bot�oGravar False
End Sub
Public Sub Encontrar()
    If Not lAllowConsult Then
        Exit Sub
    End If
    Set frmEncontrar.DBBancoDeDados = DBFinanceiro
    frmEncontrar.NomeDaJanela = "Conta Corrente"
    frmEncontrar.LabelDescription = "C�digo"
    frmEncontrar.Mensagem = "Nenhuma conta corrente foi selecionada!"
    frmEncontrar.BancoDeDados = "FINANCEIRO"
    frmEncontrar.Tabela = "CONTA CORRENTE"
    frmEncontrar.Indice = "1"
    frmEncontrar.CampoChave = "C�DIGO DO BANCO,C�DIGO DA AG�NCIA,C�DIGO"
    frmEncontrar.CampoPreencheLista = "C�DIGO DO BANCO,C�DIGO DA AG�NCIA,C�DIGO"
    frmEncontrar.Show vbModal
    lPula = True
    txtC�digoBanco = GetWordSeparatedBy(frmEncontrar.Chave, 1)
    txtC�digoAg�ncia = GetWordSeparatedBy(frmEncontrar.Chave, 2)
    txtContaCorrente = GetWordSeparatedBy(frmEncontrar.Chave, 3)
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
    
    TBLContaCorrente.Delete
    
    If Err <> 0 Then
        GeraMensagemDeErro "Conta Corrente - Excluir - " & txtDescri��oBanco & " - " & txtDescri��oAg�ncia & " - " & txtContaCorrente, True
        StatusBarAviso = "Falha na exclus�o"
        BarraDeStatus StatusBarAviso
        Exit Sub
    End If
    
    WS.CommitTrans
    
    Log gUsu�rio, "Exclus�o - Conta Corrente: " & txtContaCorrente & vbCr & "Banco: " & txtDescri��oBanco & vbCr & "Ag�ncia:" & txtDescri��oAg�ncia
    
    StatusBarAviso = "Exclus�o bem sucedida"
    BarraDeStatus StatusBarAviso
    
    If TBLContaCorrente.RecordCount = 0 Then
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
            StatusBarAviso = "Inclus�o bem sucedida"
        Else
            StatusBarAviso = "Falha na inclus�o"
        End If
    Else
        If TBLContaCorrente.RecordCount > 0 And Not TBLContaCorrente.BOF And Not TBLContaCorrente.EOF Then
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
    
    TestaInferior TBLContaCorrente, lAllowEdit, lAllowDelete, lAllowConsult
    TestaSuperior TBLContaCorrente, lAllowEdit, lAllowDelete, lAllowConsult
    
    If TBLContaCorrente.RecordCount = 0 Then
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
    If txtC�digoBanco.Enabled Then
        txtC�digoBanco.SetFocus
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
    
    txtC�digoBanco.SetFocus
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
    
    TBLContaCorrente.MoveLast
    
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
    
    TBLContaCorrente.MoveNext
    If TBLContaCorrente.EOF Then
        TBLContaCorrente.MovePrevious
        Exit Sub
    End If
    
    StatusBarAviso = "Pronto"
    BarraDeStatus StatusBarAviso
    
    Navega��oInferior lAllowConsult
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
    
    Navega��oSuperior lAllowConsult
    TestaInferior TBLContaCorrente, lAllowEdit, lAllowDelete, lAllowConsult
    
    GetRecords
End Sub
Public Sub PosRecords()
    If TBLContaCorrente.RecordCount = 0 Then
        Exit Sub
    End If
    TBLContaCorrente.Seek "=", txtC�digoBanco, txtC�digoAg�ncia, txtContaCorrente
    If TBLContaCorrente.NoMatch Then
        MsgBox "N�o consegui encontrar a Conta Corrente" + txtContaCorrente, vbExclamation, "Erro"
        TBLContaCorrente.MoveFirst
        Navega��oInferior False
        Navega��oInferior lAllowConsult
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
    txtC�digoBanco = TBLContaCorrente("C�DIGO DO BANCO")
    TBLBanco.Seek "=", txtC�digoBanco
    txtDescri��oBanco = TBLBanco("DESCRI��O")
    txtC�digoAg�ncia = TBLContaCorrente("C�DIGO DA AG�NCIA")
    TBLAg�ncia.Seek "=", txtC�digoBanco, txtC�digoAg�ncia
    txtDescri��oAg�ncia = TBLAg�ncia("DESCRI��O")
    txtContaCorrente = TBLContaCorrente("C�DIGO")
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
    
    WS.BeginTrans 'Inicia transa��es
        
    If lInserir Then
        TBLContaCorrente.AddNew
    Else
        TBLContaCorrente.Edit
    End If
    
    TBLContaCorrente("C�DIGO DO BANCO") = txtC�digoBanco
    TBLContaCorrente("C�DIGO DA AG�NCIA") = txtC�digoAg�ncia
    TBLContaCorrente("C�DIGO") = txtContaCorrente
    If lInserir Then
        TBLContaCorrente("USERNAME - CRIA") = gUsu�rio
        TBLContaCorrente("DATA - CRIA") = Date
        TBLContaCorrente("HORA - CRIA") = Time
        TBLContaCorrente("USERNAME - ALTERA") = "VAZIO"
        TBLContaCorrente("DATA - ALTERA") = vbNull
        TBLContaCorrente("HORA - ALTERA") = vbNull
    End If
    If lAlterar Then
        TBLContaCorrente("USERNAME - ALTERA") = gUsu�rio
        TBLContaCorrente("DATA - ALTERA") = Date
        TBLContaCorrente("HORA - ALTERA") = Time
    End If
    TBLContaCorrente.Update
    
Erro:
    If Err <> 0 Then
        GeraMensagemDeErro "Conta Corrente - SetRecords - " & txtDescri��oBanco & " - " & txtDescri��oAg�ncia & " - " & txtContaCorrente, True
        On Error Resume Next
        SetRecords = False
        Exit Function
    End If
    
    WS.CommitTrans 'Grava as altera��es ou inclus�es se n�o houverem erros
    
    If lInserir Then
        Log gUsu�rio, "Inclus�o - Conta Corrente: " & txtContaCorrente & vbCr & "Banco" & txtDescri��oBanco & vbCr & "Ag�ncia: " & txtDescri��oAg�ncia
    Else
        Log gUsu�rio, "Altera��o - Conta Corrente: " & txtContaCorrente & vbCr & "Banco" & txtDescri��oBanco & vbCr & "Ag�ncia: " & txtDescri��oAg�ncia
    End If
    
    SetRecords = True
End Function
Private Sub ZeraCampos()
    txtC�digoBanco = Empty
    txtDescri��oBanco = Empty
    txtC�digoAg�ncia = Empty
    txtDescri��oAg�ncia = Empty
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
    If Not Ag�nciaAberto Then
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
    
    lAllowInsert = Allow("CONTA CORRENTE", "I")
    lAllowEdit = Allow("CONTA CORRENTE", "A")
    lAllowDelete = Allow("CONTA CORRENTE", "E")
    lAllowConsult = Allow("CONTA CORRENTE", "C")
    
    ZeraCampos
    
    lPula = False
    lInserir = False
    lAlterar = False
    
    BancoAberto = AbreTabela(Dicion�rio, "FINANCEIRO", "BANCO", DBFinanceiro, TBLBanco, TBLTabela, dbOpenTable)
    
    If BancoAberto Then
        IndiceAtivoBanco = "BANCO1"
        TBLBanco.Index = IndiceAtivoBanco
    Else
        MsgBox "N�o consegui abrir a tabela 'Banco' !", vbCritical, "Erro"
        Exit Sub
    End If
    
    Ag�nciaAberto = AbreTabela(Dicion�rio, "FINANCEIRO", "AG�NCIA", DBFinanceiro, TBLAg�ncia, TBLTabela, dbOpenTable)
    
    If Ag�nciaAberto Then
        IndiceAtivoAg�ncia = "AG�NCIA1"
        TBLAg�ncia.Index = IndiceAtivoAg�ncia
    Else
        MsgBox "N�o consegui abrir a tabela 'Ag�ncia' !", vbCritical, "Erro"
        Exit Sub
    End If
    
    ContaCorrenteAberto = AbreTabela(Dicion�rio, "FINANCEIRO", "CONTA CORRENTE", DBFinanceiro, TBLContaCorrente, TBLTabela, dbOpenTable)
    
    If ContaCorrenteAberto Then
        IndiceAtivoContaCorrente = "CONTACORRENTE1"
        TBLContaCorrente.Index = IndiceAtivoContaCorrente
    Else
        MsgBox "N�o consegui abrir a tabela 'Conta Corrente' !", vbCritical, "Erro"
        Exit Sub
    End If
    
    Bot�oIncluir lAllowInsert
 
    If TBLContaCorrente.RecordCount = 0 Then
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
        
    If TBLContaCorrente.RecordCount = 0 Or TBLContaCorrente.RecordCount = 1 Then
        Navega��oSuperior False
    Else
        Navega��oInferior lAllowConsult
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
    
    Set frmContaCorrente = Nothing
End Sub
Private Sub Form_Unload(Cancel As Integer)
    If BancoAberto Then
        TBLBanco.Close
    End If
    If Ag�nciaAberto Then
        TBLAg�ncia.Close
    End If
    If ContaCorrenteAberto Then
        TBLContaCorrente.Close
    End If
    If Forms.Count = 2 Then
        AllBot�es False
    End If
End Sub
Private Sub txtC�digoAg�ncia_Change()
    If Not lPula Then
        FormatMask "@S10", txtC�digoAg�ncia
    End If
End Sub
Private Sub txtC�digoAg�ncia_KeyPress(KeyAscii As Integer)
    If Not lInserir Then
        lAlterar = True
        StatusBarAviso = "Altera��o"
        BarraDeStatus StatusBarAviso
    End If
End Sub
Private Sub txtC�digoAg�ncia_LostFocus()
    If mdiGeal.ActiveForm.Name = "frmContaCorrente" Then
        If txtC�digoAg�ncia.Enabled Then
            LeftBlank txtC�digoAg�ncia
            TBLAg�ncia.Seek "=", txtC�digoBanco, txtC�digoAg�ncia
            If TBLAg�ncia.NoMatch Then
                MsgBox "N�o encontrei a ag�ncia !" + txtC�digoAg�ncia, vbExclamation, "Aviso"
                txtC�digoAg�ncia = Empty
                txtC�digoAg�ncia.SetFocus
                Exit Sub
            End If
            txtDescri��oAg�ncia = TBLAg�ncia("DESCRI��O")
        Else
            If txtC�digoBanco.Enabled Then
                txtC�digoAg�ncia.Enabled = True
            End If
        End If
    End If
End Sub
Private Sub txtC�digoBanco_Change()
    If Not lPula Then
        FormatMask "9999", txtC�digoBanco
    End If
End Sub
Private Sub txtC�digoBanco_KeyPress(KeyAscii As Integer)
    If Not lInserir Then
        lAlterar = True
        StatusBarAviso = "Altera��o"
        BarraDeStatus StatusBarAviso
    End If
End Sub
Private Sub txtC�digoBanco_LostFocus()
    If mdiGeal.ActiveForm.Name = "frmContaCorrente" Then
        If txtC�digoBanco.Enabled Then
            LeftBlank txtC�digoBanco
            TBLBanco.Seek "=", txtC�digoBanco
            If TBLBanco.NoMatch Then
                MsgBox "N�o encontrei o banco " + txtC�digoBanco, vbExclamation, "Aviso"
                txtC�digoBanco = Empty
                txtC�digoAg�ncia.Enabled = False
                txtC�digoBanco.SetFocus
                Exit Sub
            End If
            txtDescri��oBanco = TBLBanco("DESCRI��O")
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
        StatusBarAviso = "Altera��o"
        BarraDeStatus StatusBarAviso
    End If
End Sub


