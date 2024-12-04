VERSION 5.00
Begin VB.Form frmBanco 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Banco"
   ClientHeight    =   1770
   ClientLeft      =   1140
   ClientTop       =   1515
   ClientWidth     =   6330
   Icon            =   "Banco.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   1770
   ScaleWidth      =   6330
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Enabled         =   0   'False
      Height          =   345
      Left            =   5070
      TabIndex        =   4
      Top             =   1395
      Width           =   1245
   End
   Begin VB.CommandButton cmdGravar 
      Caption         =   "&Gravar"
      Height          =   345
      Left            =   3750
      TabIndex        =   2
      Top             =   1395
      Width           =   1245
   End
   Begin VB.Frame frBanco 
      Height          =   1350
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   6330
      Begin VB.TextBox txtC�digoBanco 
         Height          =   285
         Left            =   1200
         TabIndex        =   0
         Top             =   300
         Width           =   700
      End
      Begin VB.TextBox txtDescri��oBanco 
         Height          =   285
         Left            =   1200
         TabIndex        =   1
         Top             =   750
         Width           =   5000
      End
      Begin VB.Label lblC�digoBanco 
         Caption         =   "C�digo"
         Height          =   210
         Left            =   150
         TabIndex        =   6
         Top             =   330
         Width           =   660
      End
      Begin VB.Label lblDescri��oBanco 
         Caption         =   "Descri��o"
         Height          =   180
         Left            =   150
         TabIndex        =   5
         Top             =   780
         Width           =   960
      End
   End
End
Attribute VB_Name = "frmBanco"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim TBLBanco As Table
Dim BancoAberto As Boolean
Dim IndiceBancoAtivo$
Dim txtC�digoBancoAnterior As String

Dim lPula As Boolean

Dim lInserir As Boolean
Dim lAlterar As Boolean

Dim lFechar As Boolean

Dim lAllowInsert  As Boolean
Dim lAllowEdit As Boolean
Dim lAllowDelete As Boolean
Dim lAllowConsult As Boolean

Dim StatusBarAviso$

Dim DataBaseName(1 To 1) As String
Public Relat�rio$
Public TotalDatabaseName%

Public lAtualizar As Boolean
Private Sub AtivaCampos()
    Bot�oImprimir True
    frBanco.Enabled = True
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
    
    If TBLBanco.RecordCount = 0 Then
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
    
    TestaInferior TBLBanco, lAllowEdit, lAllowDelete, lAllowConsult
    TestaSuperior TBLBanco, lAllowEdit, lAllowDelete, lAllowConsult
    
    GetRecords
End Function
Private Sub DesativaCampos()
    Bot�oImprimir False
    frBanco.Enabled = False
    cmdGravar.Enabled = False
    cmdCancelar.Enabled = False
    Bot�oGravar False
End Sub
Public Sub Encontrar()
    If Not lAllowConsult Then
        Exit Sub
    End If
    Set frmEncontrar.DBBancoDeDados = DBFinanceiro
    frmEncontrar.NomeDaJanela = "Banco"
    frmEncontrar.LabelDescription = "Descri��o"
    frmEncontrar.Mensagem = "Nenhum banco foi selecionado!"
    frmEncontrar.BancoDeDados = "FINANCEIRO"
    frmEncontrar.Tabela = "BANCO"
    frmEncontrar.Indice = "1"
    frmEncontrar.CampoChave = "C�DIGO"
    frmEncontrar.CampoPreencheLista = "DESCRI��O"
    frmEncontrar.Show vbModal
    lPula = True
    txtC�digoBanco = frmEncontrar.Chave
    lPula = False
    PosRecords
End Sub
Public Sub Excluir()
    Dim Confirma��o As Integer, Msg1$, Msg2$
    Dim TBLAg�ncia As Table

    If lAlterar Then
       If Not Cancelamento Then
            Exit Sub
        End If
    End If

    If AbreTabela(Dicion�rio, "FINANCEIRO", "AG�NCIA", DBFinanceiro, TBLAg�ncia, TBLTabela, dbOpenTable) Then
        TBLAg�ncia.Index = "AG�NCIA1"
        TBLAg�ncia.Seek ">=", txtC�digoBanco
        If Not TBLAg�ncia.NoMatch Then
            If TBLAg�ncia("C�DIGO DO BANCO") = txtC�digoBanco Then
                MsgBox "Rela��o violada!" + vbCr + "Para apagar este banco, antes � necess�rio apagar" + vbCr + "todas as ag�ncias dele dependente.", vbExclamation, "Aviso"
                TBLAg�ncia.Close
                Exit Sub
            End If
        End If
    Else
        Exit Sub
    End If
    TBLAg�ncia.Close
    
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
    
    TBLBanco.Delete
    
    If Err <> 0 Then
        GeraMensagemDeErro "Banco - Excluir - " & txtDescri��oBanco, True
        StatusBarAviso = "Falha na exclus�o"
        BarraDeStatus StatusBarAviso
        Exit Sub
    End If
    
    WS.CommitTrans
    
    Log gUsu�rio, "Exclus�o - Banco: " & txtC�digoBanco & " - " & txtDescri��oBanco
    
    StatusBarAviso = "Exclus�o bem sucedida"
    BarraDeStatus StatusBarAviso
    
    If TBLBanco.RecordCount = 0 Then
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
    
    If TBLBanco.BOF Then
        TBLBanco.MoveFirst
    ElseIf TBLBanco.EOF Then
        TBLBanco.MoveLast
    Else
        TBLBanco.MovePrevious
        If TBLBanco.BOF Then
            TBLBanco.MoveNext
        End If
    End If
    
    GetRecords
    
    TestaInferior TBLBanco, lAllowEdit, lAllowDelete, lAllowConsult
    TestaSuperior TBLBanco, lAllowEdit, lAllowDelete, lAllowConsult
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
        If TBLBanco.RecordCount > 0 And Not TBLBanco.BOF And Not TBLBanco.EOF Then
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
    
    TestaInferior TBLBanco, lAllowEdit, lAllowDelete, lAllowConsult
    TestaSuperior TBLBanco, lAllowEdit, lAllowDelete, lAllowConsult
    
    If TBLBanco.RecordCount = 0 Then
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
    
    TBLBanco.MoveFirst
    
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
    
    TBLBanco.MoveLast
    
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
    
    TBLBanco.MoveNext
    If TBLBanco.EOF Then
        TBLBanco.MovePrevious
        Exit Sub
    End If
    
    StatusBarAviso = "Pronto"
    BarraDeStatus StatusBarAviso
    
    Navega��oInferior lAllowConsult
    TestaSuperior TBLBanco, lAllowEdit, lAllowDelete, lAllowConsult
    
    GetRecords
End Sub
Public Sub MovePrevious()
    If lAlterar Then
       If Not Cancelamento Then
            Exit Sub
        End If
    End If
    
    TBLBanco.MovePrevious
    If TBLBanco.BOF Then
        TBLBanco.MoveNext
        Exit Sub
    End If
    
    StatusBarAviso = "Pronto"
    BarraDeStatus StatusBarAviso
    
    Navega��oSuperior lAllowConsult
    TestaInferior TBLBanco, lAllowEdit, lAllowDelete, lAllowConsult
    
    GetRecords
End Sub
Public Sub PosRecords()
    If TBLBanco.RecordCount = 0 Then
        Exit Sub
    End If
    
    TBLBanco.Seek "=", txtC�digoBanco
    If TBLBanco.NoMatch Then
        MsgBox "N�o consegui encontrar " + txtC�digoBanco, vbExclamation, "Erro"
        TBLBanco.MoveFirst
        Navega��oInferior False
        Navega��oInferior lAllowConsult
    Else
        TestaInferior TBLBanco, lAllowEdit, lAllowDelete, lAllowConsult
        TestaSuperior TBLBanco, lAllowEdit, lAllowDelete, lAllowConsult
    End If
    GetRecords
End Sub
Public Function PushDataBaseName(ByVal Posi��o As Integer) As String
    PushDataBaseName = DataBaseName(Posi��o)
End Function
Private Sub GetRecords()
    On Error GoTo Erro
    
    If Not lAllowConsult Then
        ZeraCampos
        DesativaCampos
        Exit Sub
    End If
    txtC�digoBanco = TBLBanco("C�DIGO")
    txtC�digoBancoAnterior = txtC�digoBanco
    txtDescri��oBanco = TBLBanco("DESCRI��O")
    If Not lAllowEdit Then
        DesativaCampos
    End If
    
    Exit Sub
    
Erro:
    GeraMensagemDeErro "Banco - GetRecords"
    Resume Next
End Sub
Private Function SetRecords()
    On Error Resume Next
    
    Dim Msg$
    Dim Confirma��o As Integer, Msg1$, Msg2$, AchouAg�ncia As Boolean, AchouContaCorrente As Boolean
    Dim TBLAg�ncia As Table
    Dim TBLContaCorrente As Table
    Dim SQL As String
    Dim Cont%

    If (txtC�digoBanco <> txtC�digoBancoAnterior) And Not lInserir Then
        If AbreTabela(Dicion�rio, "FINANCEIRO", "AG�NCIA", DBFinanceiro, TBLAg�ncia, TBLTabela, dbOpenTable) Then
            TBLAg�ncia.Index = "AG�NCIA1"
            TBLAg�ncia.Seek ">=", txtC�digoBancoAnterior
            If Not TBLAg�ncia.NoMatch Then
                If TBLAg�ncia("C�DIGO DO BANCO") = txtC�digoBancoAnterior Then
                    AchouAg�ncia = True
                    Confirma��o = MsgBox("Voc� necessita alterar as ag�ncias relacionadas com este banco !" + vbCr + "Deseja realizar agora as altera��es de" + vbCr + "todas as ag�ncias dele dependente?", vbYesNo + vbQuestion + vbDefaultButton2, "Aviso")
                End If
            Else
                AchouAg�ncia = False
            End If
        Else
            Exit Function
        End If
        TBLAg�ncia.Close
        
        If AchouAg�ncia Then
            If Confirma��o = vbNo Then
                SetRecords = False
                Exit Function
            End If
        End If
        
        If AbreTabela(Dicion�rio, "FINANCEIRO", "CONTA CORRENTE", DBFinanceiro, TBLContaCorrente, TBLTabela, dbOpenTable) Then
            TBLContaCorrente.Index = "CONTACORRENTE1"
            TBLContaCorrente.Seek ">=", txtC�digoBancoAnterior
            If Not TBLContaCorrente.NoMatch Then
                If TBLContaCorrente("C�DIGO DO BANCO") = txtC�digoBancoAnterior Then
                    AchouContaCorrente = True
                End If
            Else
                AchouContaCorrente = False
            End If
        Else
            Exit Function
        End If
        TBLContaCorrente.Close
    Else
        AchouAg�ncia = False
        AchouContaCorrente = False
    End If
    
    On Error GoTo Erro
    
    WS.BeginTrans 'Inicia uma Transa��o
    
    If lInserir Then
        TBLBanco.AddNew
    Else
        TBLBanco.Edit
    End If
    
    TBLBanco("C�DIGO") = Trim(txtC�digoBanco)
    TBLBanco("DESCRI��O") = Trim(txtDescri��oBanco)
    If lInserir Then
        TBLBanco("USERNAME - CRIA") = gUsu�rio
        TBLBanco("DATA - CRIA") = Date
        TBLBanco("HORA - CRIA") = Time
        TBLBanco("USERNAME - ALTERA") = "VAZIO"
        TBLBanco("DATA - ALTERA") = vbNull
        TBLBanco("HORA - ALTERA") = vbNull
    End If
    If lAlterar Then
        TBLBanco("USERNAME - ALTERA") = gUsu�rio
        TBLBanco("DATA - ALTERA") = Date
        TBLBanco("HORA - ALTERA") = Time
    End If
    TBLBanco.Update
        
    If AchouAg�ncia Then
        SQL = "Update AG�NCIA Set [C�DIGO DO BANCO]= '" + txtC�digoBanco + "' Where [C�DIGO DO BANCO]= '" + txtC�digoBancoAnterior + "'"
        DBFinanceiro.Execute SQL
    End If
    If AchouContaCorrente Then
        SQL = "Update [CONTA CORRENTE] Set [C�DIGO DO BANCO]= '" + txtC�digoBanco + "' Where [C�DIGO DO BANCO]= '" + txtC�digoBancoAnterior + "'"
        DBFinanceiro.Execute SQL
    End If
    
Erro:
    If Err <> 0 Then
        TBLBanco.CancelUpdate
        GeraMensagemDeErro "Banco - SetRecords - " & txtDescri��oBanco, True
        SetRecords = False
        Exit Function
    End If

    WS.CommitTrans 'Grava as altera��es ou inclus�es se n�o houverem erros
    
    'Se a janela Ag�ncia estiver aberta atualiza seus valores se necess�rio.
    If Not lInserir Then
        For Cont = 1 To Forms.Count - 1
            If Forms(Cont).Name = "frmAg�ncia" Or Forms(Cont).Name = "frmContaCorrente" Then
                If Forms(Cont).txtC�digoBanco = txtC�digoBancoAnterior Then
                    Forms(Cont).txtC�digoBanco = txtC�digoBanco
                    Forms(Cont).txtDescri��oBanco = txtDescri��oBanco
                    Forms(Cont).PosRecords
                End If
            End If
        Next
    End If
    
    If lInserir Then
        Log gUsu�rio, "Inclus�o - Banco: " & txtC�digoBanco & " - " & txtDescri��oBanco
    Else
        Log gUsu�rio, "Altera��o - Banco: " & txtC�digoBanco & " - " & txtDescri��oBanco
    End If
    
    SetRecords = True
End Function
Private Sub ZeraCampos()
    txtC�digoBanco = Empty
    txtC�digoBancoAnterior = Empty
    txtDescri��oBanco = Empty
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
    If Not BancoAberto Then
        Unload Me
        Exit Sub
    End If
    TestaInferior TBLBanco, lAllowEdit, lAllowDelete, lAllowConsult
    TestaSuperior TBLBanco, lAllowEdit, lAllowDelete, lAllowConsult
    If TBLBanco.RecordCount = 0 Then
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
    
    lAllowInsert = Allow("BANCO", "I")
    lAllowEdit = Allow("BANCO", "A")
    lAllowDelete = Allow("BANCO", "E")
    lAllowConsult = Allow("BANCO", "C")
    
    ZeraCampos
    
    lPula = False
    lInserir = False
    lAlterar = False
    
    BancoAberto = AbreTabela(Dicion�rio, "FINANCEIRO", "BANCO", DBFinanceiro, TBLBanco, TBLTabela, dbOpenTable)
    
    If BancoAberto Then
        IndiceBancoAtivo = "BANCO1"
        TBLBanco.Index = IndiceBancoAtivo
    Else
        MsgBox "N�o consegui abrir a tabela 'Banco' !", vbCritical, "Erro"
        Exit Sub
    End If
    
    Bot�oIncluir lAllowInsert
 
    If TBLBanco.RecordCount = 0 Then
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
        
    If TBLBanco.RecordCount = 0 Or TBLBanco.RecordCount = 1 Then
        Navega��oSuperior False
    Else
        Navega��oInferior lAllowConsult
    End If
    
    StatusBarAviso = "Pronto"
    Relat�rio = AddPath(Aplica��oPath, "REPORT\BANCO.RPT")
    TotalDatabaseName = 1
    DataBaseName(1) = AddPath(Aplica��oPath, "DATABASE\FINANCEIRO.MDB")
    lFechar = False
    Exit Sub
    
Erro:
    GeraMensagemDeErro "Banco - Load"
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
    
    Set frmBanco = Nothing
End Sub
Private Sub Form_Unload(Cancel As Integer)
    If BancoAberto Then
        TBLBanco.Close
    End If
    If Forms.Count = 2 Then
        AllBot�es False
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
    If txtC�digoBanco.Enabled Then
        LeftBlank txtC�digoBanco
    End If
End Sub
Private Sub txtDescri��oBanco_Change()
    If Not lPula Then
        FormatMask "@!S30", txtDescri��oBanco
    End If
End Sub
Private Sub txtDescri��oBanco_KeyPress(KeyAscii As Integer)
    If Not lInserir Then
        lAlterar = True
        StatusBarAviso = "Altera��o"
        BarraDeStatus StatusBarAviso
    End If
End Sub

