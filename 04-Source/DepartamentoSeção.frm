VERSION 5.00
Begin VB.Form frmDepartamentoSe��o 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Departamento - Se��o"
   ClientHeight    =   1695
   ClientLeft      =   870
   ClientTop       =   1515
   ClientWidth     =   8250
   Icon            =   "DepartamentoSe��o.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   1695
   ScaleWidth      =   8250
   Begin VB.Frame frDepartamentoSe��o 
      Height          =   1275
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   8235
      Begin VB.TextBox txtDescri��oSe��o 
         Enabled         =   0   'False
         Height          =   285
         Left            =   2100
         MaxLength       =   30
         TabIndex        =   8
         Top             =   750
         Width           =   5000
      End
      Begin VB.TextBox txtC�digoSe��o 
         Height          =   285
         Left            =   1230
         TabIndex        =   1
         Top             =   750
         Width           =   750
      End
      Begin VB.TextBox txtC�digoDepartamento 
         Height          =   285
         Left            =   1230
         TabIndex        =   0
         Top             =   300
         Width           =   750
      End
      Begin VB.TextBox txtDescri��oDepartamento 
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
      Begin VB.Label lblSe��o 
         Caption         =   "Se��o"
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
Attribute VB_Name = "frmDepartamentoSe��o"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim TBLDepartamentoSe��o As Table
Dim DepartamentoSe��oAberto As Boolean
Dim IndiceDepartamentoSe��oAtivo$

Dim TBLDepartamento As Table
Dim DepartamentoAberto As Boolean
Dim IndiceDepartamentoAtivo$

Dim TBLSe��o As Table
Dim Se��oAberto As Boolean
Dim IndiceSe��oAtivo$

Dim lAllowInsert  As Boolean
Dim lAllowEdit    As Boolean
Dim lAllowDelete  As Boolean
Dim lAllowConsult As Boolean

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
    frDepartamentoSe��o.Enabled = True
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
    
    If TBLDepartamentoSe��o.RecordCount = 0 Then
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
    
    TestaInferior TBLDepartamentoSe��o, lAllowEdit, lAllowDelete, lAllowConsult
    TestaSuperior TBLDepartamentoSe��o, lAllowEdit, lAllowDelete, lAllowConsult
    
    GetRecords
End Function
Private Sub DesativaCampos()
    Bot�oImprimir False
    frDepartamentoSe��o.Enabled = False
    cmdGravar.Enabled = False
    cmdCancelar.Enabled = False
    Bot�oGravar False
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
    
    TBLDepartamentoSe��o.Delete
    
    If Err <> 0 Then
        GeraMensagemDeErro "Departamento - Se��o - Excluir - " & txtC�digoDepartamento & txtC�digoSe��o, True
        StatusBarAviso = "Falha na exclus�o"
        BarraDeStatus StatusBarAviso
        Exit Sub
    End If
    
    WS.CommitTrans
    
    Log gUsu�rio, "Exclus�o - Departamento - Se��o: " & txtDescri��oDepartamento & " - " & txtDescri��oSe��o
    
    StatusBarAviso = "Exclus�o bem sucedida"
    BarraDeStatus StatusBarAviso
    
    If TBLDepartamentoSe��o.RecordCount = 0 Then
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
    
    If TBLDepartamentoSe��o.BOF Then
        TBLDepartamentoSe��o.MoveFirst
    ElseIf TBLDepartamentoSe��o.EOF Then
        TBLDepartamentoSe��o.MoveLast
    Else
        TBLDepartamentoSe��o.MovePrevious
        If TBLDepartamentoSe��o.BOF Then
            TBLDepartamentoSe��o.MoveNext
        End If
    End If
    
    GetRecords
    
    TestaInferior TBLDepartamentoSe��o, lAllowEdit, lAllowDelete, lAllowConsult
    TestaSuperior TBLDepartamentoSe��o, lAllowEdit, lAllowDelete, lAllowConsult
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
        If TBLDepartamentoSe��o.RecordCount > 0 And Not TBLDepartamentoSe��o.BOF And Not TBLDepartamentoSe��o.EOF Then
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
    
    TestaInferior TBLDepartamentoSe��o, lAllowEdit, lAllowDelete, lAllowConsult
    TestaSuperior TBLDepartamentoSe��o, lAllowEdit, lAllowDelete, lAllowConsult
    
    If TBLDepartamentoSe��o.RecordCount = 0 Then
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
    
    If txtC�digoDepartamento.Enabled Then
        txtC�digoDepartamento.SetFocus
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
    
    txtC�digoDepartamento.SetFocus
End Sub
Public Sub MoveFirst()
    If lAlterar Then
       If Not Cancelamento Then
            Exit Sub
        End If
    End If
    
    StatusBarAviso = "Pronto"
    BarraDeStatus StatusBarAviso
    
    TBLDepartamentoSe��o.MoveFirst
    
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
    
    TBLDepartamentoSe��o.MoveLast
    
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
    
    TBLDepartamentoSe��o.MoveNext
    If TBLDepartamentoSe��o.EOF Then
        TBLDepartamentoSe��o.MovePrevious
        Exit Sub
    End If
    
    StatusBarAviso = "Pronto"
    BarraDeStatus StatusBarAviso
    
    Navega��oInferior lAllowConsult
    TestaSuperior TBLDepartamentoSe��o, lAllowEdit, lAllowDelete, lAllowConsult
    
    GetRecords
End Sub
Public Sub MovePrevious()
    If lAlterar Then
       If Not Cancelamento Then
            Exit Sub
        End If
    End If
    
    TBLDepartamentoSe��o.MovePrevious
    
    If TBLDepartamentoSe��o.BOF Then
        TBLDepartamentoSe��o.MoveNext
        Exit Sub
    End If
    
    StatusBarAviso = "Pronto"
    BarraDeStatus StatusBarAviso
    
    Navega��oSuperior lAllowConsult
    TestaInferior TBLDepartamentoSe��o, lAllowEdit, lAllowDelete, lAllowConsult
    
    GetRecords
End Sub
Public Sub PosRecords()
    If TBLDepartamentoSe��o.RecordCount = 0 Then
        Exit Sub
    End If
    
    TBLDepartamentoSe��o.Seek "=", txtC�digoDepartamento, txtC�digoSe��o
    If TBLDepartamentoSe��o.NoMatch Then
        MsgBox "N�o consegui encontrar " + txtC�digoDepartamento + " - " + txtC�digoSe��o, vbExclamation, "Erro"
        TBLDepartamentoSe��o.MoveFirst
        Navega��oInferior False
        Navega��oInferior lAllowConsult
    Else
        TestaInferior TBLDepartamentoSe��o, lAllowEdit, lAllowDelete, lAllowConsult
        TestaSuperior TBLDepartamentoSe��o, lAllowEdit, lAllowDelete, lAllowConsult
    End If
    GetRecords
End Sub
Public Function PushDataBaseName(ByVal Posi��o As Integer) As String
    PushDataBaseName = DataBaseName(Posi��o)
End Function
Private Sub GetRecords()
    If Not lAllowConsult Then
        ZeraCampos
        DesativaCampos
        Exit Sub
    End If
    txtC�digoDepartamento = TBLDepartamentoSe��o("C�DIGO DO DEPTO")
    TBLDepartamento.Seek "=", txtC�digoDepartamento
    txtDescri��oDepartamento = TBLDepartamento("DESCRI��O")
    txtC�digoSe��o = TBLDepartamentoSe��o("C�DIGO DA SE��O")
    TBLSe��o.Seek "=", txtC�digoSe��o
    txtDescri��oSe��o = TBLSe��o("DESCRI��O")
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
        TBLDepartamentoSe��o.AddNew
    Else
        TBLDepartamentoSe��o.Edit
    End If
    
    TBLDepartamentoSe��o("C�DIGO DO DEPTO") = txtC�digoDepartamento
    TBLDepartamentoSe��o("C�DIGO DA SE��O") = txtC�digoSe��o
    If lInserir Then
        TBLDepartamentoSe��o("USERNAME - CRIA") = gUsu�rio
        TBLDepartamentoSe��o("DATA - CRIA") = Date
        TBLDepartamentoSe��o("HORA - CRIA") = Time
        TBLDepartamentoSe��o("USERNAME - ALTERA") = "VAZIO"
        TBLDepartamentoSe��o("DATA - ALTERA") = vbNull
        TBLDepartamentoSe��o("HORA - ALTERA") = vbNull
    End If
    If lAlterar Then
        TBLDepartamentoSe��o("USERNAME - ALTERA") = gUsu�rio
        TBLDepartamentoSe��o("DATA - ALTERA") = Date
        TBLDepartamentoSe��o("HORA - ALTERA") = Time
    End If
    TBLDepartamentoSe��o.Update
        
Erro:
    If Err <> 0 Then
        TBLDepartamentoSe��o.CancelUpdate
        GeraMensagemDeErro "Departamento - Se��o - SetRecords - " & txtC�digoDepartamento & txtC�digoSe��o, True
        SetRecords = False
        Exit Function
    End If

    WS.CommitTrans 'Grava as altera��es ou inclus�es se n�o houverem erros
    
    If lInserir Then
        Log gUsu�rio, "Inclus�o - Departamento - Se��o: " & txtDescri��oDepartamento & " - " & txtDescri��oSe��o
    Else
        Log gUsu�rio, "Altera��o - Departamento - Se��o: " & txtDescri��oDepartamento & " - " & txtDescri��oSe��o
    End If
    
    SetRecords = True
End Function
Private Sub ZeraCampos()
    txtC�digoDepartamento = Empty
    txtDescri��oDepartamento = Empty
    txtC�digoSe��o = Empty
    txtDescri��oSe��o = Empty
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
    If Not DepartamentoSe��oAberto Then
        Unload Me
        Exit Sub
    End If
    If Not DepartamentoAberto Then
        Unload Me
        Exit Sub
    End If
    If Not Se��oAberto Then
        Unload Me
        Exit Sub
    End If
    
    TestaInferior TBLDepartamentoSe��o, lAllowEdit, lAllowDelete, lAllowConsult
    TestaSuperior TBLDepartamentoSe��o, lAllowEdit, lAllowDelete, lAllowConsult
    
    If TBLDepartamentoSe��o.RecordCount = 0 Then
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
    
    lAllowInsert = Allow("DEPARTAMENTO - SE��O", "I")
    lAllowEdit = Allow("DEPARTAMENTO - SE��O", "A")
    lAllowDelete = Allow("DEPARTAMENTO - SE��O", "E")
    lAllowConsult = Allow("DEPARTAMENTO - SE��O", "C")
    
    ZeraCampos
    
    lInserir = False
    lAlterar = False
    
    DepartamentoSe��oAberto = AbreTabela(Dicion�rio, "CADASTRO", "DEPARTAMENTO - SE��O", DBCadastro, TBLDepartamentoSe��o, TBLTabela, dbOpenTable)
    
    If DepartamentoSe��oAberto Then
        IndiceDepartamentoSe��oAtivo = "DEPARTAMENTOSE��O1"
        TBLDepartamentoSe��o.Index = IndiceDepartamentoSe��oAtivo
    Else
        MsgBox "N�o consegui abrir a tabela 'Departamento - Se��o' !", vbCritical, "Erro"
        Exit Sub
    End If
    
    DepartamentoAberto = AbreTabela(Dicion�rio, "CADASTRO", "DEPARTAMENTO", DBCadastro, TBLDepartamento, TBLTabela, dbOpenTable)
    
    If DepartamentoAberto Then
        IndiceDepartamentoAtivo = "DEPARTAMENTO1"
        TBLDepartamento.Index = IndiceDepartamentoAtivo
    Else
        MsgBox "N�o consegui abrir a tabela 'Departamento' !", vbCritical, "Erro"
        Exit Sub
    End If
    
    Se��oAberto = AbreTabela(Dicion�rio, "CADASTRO", "SE��O", DBCadastro, TBLSe��o, TBLTabela, dbOpenTable)
    
    If Se��oAberto Then
        IndiceSe��oAtivo = "SE��O1"
        TBLSe��o.Index = IndiceSe��oAtivo
    Else
        MsgBox "N�o consegui abrir a tabela 'Se��o' !", vbCritical, "Erro"
        Exit Sub
    End If

    Bot�oIncluir lAllowInsert
 
    If TBLDepartamentoSe��o.RecordCount = 0 Then
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
        
    If TBLDepartamentoSe��o.RecordCount = 0 Or TBLDepartamentoSe��o.RecordCount = 1 Then
        Navega��oSuperior False
    Else
        Navega��oInferior lAllowConsult
    End If
    
    StatusBarAviso = "Pronto"
    Relat�rio = AddPath(Aplica��oPath, "REPORT\DEPTOSE��O.RPT")
    TotalDatabaseName = 1
    DataBaseName(1) = AddPath(Aplica��oPath, "DATABASE\CADASTRO.MDB")
    mFechar = False
    Exit Sub
    
Erro:
    GeraMensagemDeErro -"Departamento - Se��o - Load"
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
    
    Set frmDepartamentoSe��o = Nothing
End Sub
Private Sub Form_Unload(Cancel As Integer)
    If DepartamentoSe��oAberto Then
        TBLDepartamentoSe��o.Close
    End If
    If DepartamentoAberto Then
        TBLDepartamento.Close
    End If
    If Se��oAberto Then
        TBLSe��o.Close
    End If
    If Forms.Count = 2 Then
        AllBot�es False
    End If
End Sub
Private Sub txtC�digoDepartamento_Change()
    FormatMask "9999", txtC�digoDepartamento
End Sub
Private Sub txtC�digoDepartamento_KeyPress(KeyAscii As Integer)
    If Not lInserir Then
        lAlterar = True
        StatusBarAviso = "Altera��o"
        BarraDeStatus StatusBarAviso
    End If
End Sub
Private Sub txtC�digoDepartamento_LostFocus()
    If mdiGeal.ActiveForm.Name = "frmDepartamentoSe��o" Then
        If txtC�digoDepartamento.Enabled Then
            LeftBlank txtC�digoDepartamento
            TBLDepartamento.Seek "=", txtC�digoDepartamento
            If TBLDepartamento.NoMatch Then
                MsgBox "N�o encontrei o departamento " + txtC�digoDepartamento, vbExclamation, "Aviso"
                frmEncontra.BancoDeDados = "CADASTRO"
                frmEncontra.Tabela = "DEPARTAMENTO"
                frmEncontra.Inicio = 1
                frmEncontra.Fim = 4
                frmEncontra.Caption = "Departamnento"
                frmEncontra.Show vbModal
                txtC�digoDepartamento = frmEncontra.C�digo
                TBLDepartamento.Seek "=", txtC�digoDepartamento
            End If
            txtDescri��oDepartamento = TBLDepartamento("DESCRI��O")
        End If
    End If
End Sub
Private Sub txtC�digoSe��o_Change()
    FormatMask "9999", txtC�digoSe��o
End Sub
Private Sub txtC�digoSe��o_KeyPress(KeyAscii As Integer)
    If Not lInserir Then
        lAlterar = True
        StatusBarAviso = "Altera��o"
        BarraDeStatus StatusBarAviso
    End If
End Sub
Private Sub txtC�digoSe��o_LostFocus()
    If mdiGeal.ActiveForm.Name = "frmDepartamentoSe��o" Then
        If txtC�digoSe��o.Enabled Then
            LeftBlank txtC�digoSe��o
            TBLSe��o.Seek "=", txtC�digoSe��o
            If TBLSe��o.NoMatch Then
                MsgBox "N�o encontrei a se��o " + txtC�digoSe��o, vbExclamation, "Aviso"
                frmEncontra.BancoDeDados = "CADASTRO"
                frmEncontra.Tabela = "SE��O"
                frmEncontra.Inicio = 1
                frmEncontra.Fim = 4
                frmEncontra.Show vbModal
                frmEncontra.Caption = "Se��o"
                txtC�digoSe��o = frmEncontra.C�digo
                TBLSe��o.Seek "=", txtC�digoSe��o
            End If
            txtDescri��oSe��o = TBLSe��o("DESCRI��O")
        End If
    End If
End Sub
