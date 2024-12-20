VERSION 5.00
Begin VB.Form frmDepartamento 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Departamento"
   ClientHeight    =   1695
   ClientLeft      =   2085
   ClientTop       =   2145
   ClientWidth     =   6480
   Icon            =   "Departamento.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   1695
   ScaleWidth      =   6480
   Begin VB.Frame frDepartamento 
      Height          =   1275
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   6465
      Begin VB.TextBox txtC�digo 
         Height          =   285
         Left            =   1200
         TabIndex        =   0
         Top             =   300
         Width           =   750
      End
      Begin VB.TextBox txtDescri��o 
         Height          =   285
         Left            =   1200
         TabIndex        =   1
         Top             =   750
         Width           =   5000
      End
      Begin VB.Label lblC�digo 
         Caption         =   "C�digo"
         Height          =   200
         Left            =   150
         TabIndex        =   6
         Top             =   330
         Width           =   855
      End
      Begin VB.Label lblDescri��o 
         Caption         =   "Descri��o"
         Height          =   195
         Left            =   150
         TabIndex        =   5
         Top             =   780
         Width           =   885
      End
   End
   Begin VB.CommandButton cmdGravar 
      Caption         =   "&Gravar"
      Height          =   345
      Left            =   3885
      TabIndex        =   2
      Top             =   1320
      Width           =   1245
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   345
      Left            =   5205
      TabIndex        =   3
      Top             =   1320
      Width           =   1245
   End
End
Attribute VB_Name = "frmDepartamento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim TBLDepartamento As Table
Dim DepartamentoAberto As Boolean
Dim IndiceDepartamentoAtivo$
Dim txtC�digoAnterior As String

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
    frDepartamento.Enabled = True
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
    
    If TBLDepartamento.RecordCount = 0 Then
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
    
    TestaInferior TBLDepartamento, lAllowEdit, lAllowDelete, lAllowConsult
    TestaSuperior TBLDepartamento, lAllowEdit, lAllowDelete, lAllowConsult
    
    GetRecords
End Function
Private Sub DesativaCampos()
    Bot�oImprimir False
    frDepartamento.Enabled = False
    cmdGravar.Enabled = False
    cmdCancelar.Enabled = False
    Bot�oGravar False
End Sub
Public Sub Encontrar()
    If Not lAllowConsult Then
        Exit Sub
    End If
    Set frmEncontrar.DBBancoDeDados = DBCadastro
    frmEncontrar.NomeDaJanela = "Departamento"
    frmEncontrar.LabelDescription = "Descri��o"
    frmEncontrar.Mensagem = "Nenhum departamento foi selecionado!"
    frmEncontrar.BancoDeDados = "CADASTRO"
    frmEncontrar.Tabela = "DEPARTAMENTO"
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
    Dim TBLDepartamentoSe��o As Table

    If lAlterar Then
       If Not Cancelamento Then
            Exit Sub
        End If
    End If

    If AbreTabela(Dicion�rio, "CADASTRO", "DEPARTAMENTO - SE��O", DBCadastro, TBLDepartamentoSe��o, TBLTabela, dbOpenTable) Then
        TBLDepartamentoSe��o.Index = "DEPARTAMENTOSE��O1"
        TBLDepartamentoSe��o.Seek ">=", txtC�digo
        If Not TBLDepartamentoSe��o.NoMatch Then
            If TBLDepartamentoSe��o("C�DIGO DO DEPTO") = txtC�digo Then
                MsgBox "Rela��o violada!" + vbCr + "Para apagar este departamento, antes � necess�rio apagar" + vbCr + "todos os 'departamentos-se��o' dele dependente.", vbExclamation, "Aviso"
                TBLDepartamentoSe��o.Close
                Exit Sub
            End If
        End If
    End If
    TBLDepartamentoSe��o.Close

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
    
    TBLDepartamento.Delete
    
    If Err <> 0 Then
        GeraMensagemDeErro "Departamento - Excluir - " & txtDescri��o, True
        StatusBarAviso = "Falha na exclus�o"
        BarraDeStatus StatusBarAviso
        Exit Sub
    End If
    
    WS.CommitTrans
    
    Log gUsu�rio, "Exclus�o - Departamento: " & txtC�digo & " - " & txtDescri��o
    
    StatusBarAviso = "Exclus�o bem sucedida"
    BarraDeStatus StatusBarAviso
    
    If TBLDepartamento.RecordCount = 0 Then
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
    
    If TBLDepartamento.BOF Then
        TBLDepartamento.MoveFirst
    ElseIf TBLDepartamento.EOF Then
        TBLDepartamento.MoveLast
    Else
        TBLDepartamento.MovePrevious
        If TBLDepartamento.BOF Then
            TBLDepartamento.MoveNext
        End If
    End If
    
    GetRecords
    
    TestaInferior TBLDepartamento, lAllowEdit, lAllowDelete, lAllowConsult
    TestaSuperior TBLDepartamento, lAllowEdit, lAllowDelete, lAllowConsult
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
        If TBLDepartamento.RecordCount > 0 And Not TBLDepartamento.BOF And Not TBLDepartamento.EOF Then
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
    
    TestaInferior TBLDepartamento, lAllowEdit, lAllowDelete, lAllowConsult
    TestaSuperior TBLDepartamento, lAllowEdit, lAllowDelete, lAllowConsult
    
    If TBLDepartamento.RecordCount = 0 Then
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
    
    TBLDepartamento.MoveFirst
    
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
    
    TBLDepartamento.MoveLast
    
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
    
    TBLDepartamento.MoveNext
    If TBLDepartamento.EOF Then
        TBLDepartamento.MovePrevious
        Exit Sub
    End If
    
    StatusBarAviso = "Pronto"
    BarraDeStatus StatusBarAviso
    
    Navega��oInferior lAllowConsult
    TestaSuperior TBLDepartamento, lAllowEdit, lAllowDelete, lAllowConsult
    
    GetRecords
End Sub
Public Sub MovePrevious()
    If lAlterar Then
       If Not Cancelamento Then
            Exit Sub
        End If
    End If
    
    TBLDepartamento.MovePrevious
    If TBLDepartamento.BOF Then
        TBLDepartamento.MoveNext
        Exit Sub
    End If
    
    StatusBarAviso = "Pronto"
    BarraDeStatus StatusBarAviso
    
    Navega��oSuperior lAllowConsult
    TestaInferior TBLDepartamento, lAllowEdit, lAllowDelete, lAllowConsult
    
    GetRecords
End Sub
Public Sub PosRecords()
    If TBLDepartamento.RecordCount = 0 Then
        Exit Sub
    End If
    
    TBLDepartamento.Seek "=", txtC�digo
    If TBLDepartamento.NoMatch Then
        MsgBox "N�o consegui encontrar " + txtC�digo, vbExclamation, "Erro"
        TBLDepartamento.MoveFirst
        Navega��oInferior False
        Navega��oInferior lAllowConsult
    Else
        TestaInferior TBLDepartamento, lAllowEdit, lAllowDelete, lAllowConsult
        TestaSuperior TBLDepartamento, lAllowEdit, lAllowDelete, lAllowConsult
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
    txtC�digo = TBLDepartamento("C�DIGO")
    txtC�digoAnterior = txtC�digo
    txtDescri��o = TBLDepartamento("DESCRI��O")
    If Not lAllowEdit Then
        DesativaCampos
    End If
    
    Exit Sub
    
Erro:
    GeraMensagemDeErro "Departamento - GetRecords "
    Resume Next
End Sub
Private Function SetRecords()
    On Error Resume Next
    
    Dim Msg$
    Dim Confirma��o As Integer, Msg1$, Msg2$, AchouDepartamentoSe��o As Boolean
    Dim TBLDepartamentoSe��o As Table
    Dim SQL As String
    Dim Cont%

    If (txtC�digo <> txtC�digoAnterior) And Not lInserir Then
        If AbreTabela(Dicion�rio, "CADASTRO", "DEPARTAMENTO - SE��O", DBCadastro, TBLDepartamentoSe��o, TBLTabela, dbOpenTable) Then
            TBLDepartamentoSe��o.Index = "DEPARTAMENTOSE��O1"
            TBLDepartamentoSe��o.Seek ">=", txtC�digoAnterior
            If Not TBLDepartamentoSe��o.NoMatch Then
                If TBLDepartamentoSe��o("C�DIGO DO DEPTO") = txtC�digoAnterior Then
                    AchouDepartamentoSe��o = True
                    Confirma��o = MsgBox("Voc� necessita alterar os 'Departamentos-Se��o' relacionados com este departamento !" + vbCr + "Deseja realizar agora as altera��es de" + vbCr + "todas os 'departamentos-se��o' dele dependente?", vbYesNo + vbQuestion + vbDefaultButton2, "Aviso")
                End If
            Else
                AchouDepartamentoSe��o = False
            End If
        Else
            Exit Function
        End If
        TBLDepartamentoSe��o.Close
        
        If AchouDepartamentoSe��o Then
            If Confirma��o = vbNo Then
                SetRecords = False
                Exit Function
            End If
        End If
    Else
        AchouDepartamentoSe��o = False
    End If
    
    On Error GoTo Erro
    
    WS.BeginTrans 'Inicia uma Transa��o
    
    If lInserir Then
        TBLDepartamento.AddNew
    Else
        TBLDepartamento.Edit
    End If
    
    TBLDepartamento("C�DIGO") = txtC�digo
    TBLDepartamento("DESCRI��O") = txtDescri��o
    If lInserir Then
        TBLDepartamento("USERNAME - CRIA") = gUsu�rio
        TBLDepartamento("DATA - CRIA") = Date
        TBLDepartamento("HORA - CRIA") = Time
        TBLDepartamento("USERNAME - ALTERA") = "VAZIO"
        TBLDepartamento("DATA - ALTERA") = vbNull
        TBLDepartamento("HORA - ALTERA") = vbNull
    End If
    If lAlterar Then
        TBLDepartamento("USERNAME - ALTERA") = gUsu�rio
        TBLDepartamento("DATA - ALTERA") = Date
        TBLDepartamento("HORA - ALTERA") = Time
    End If
    TBLDepartamento.Update
        
    If AchouDepartamentoSe��o Then
        SQL = "Update [DEPARTAMENTO - SE��O] Set [C�DIGO DO DEPTO]= '" + txtC�digo + "' Where [C�DIGO DO DEPTO]= '" + txtC�digoAnterior + "'"
        DBCadastro.Execute SQL
    End If
        
Erro:
    If Err <> 0 Then
        TBLDepartamento.CancelUpdate
        GeraMensagemDeErro "Departamento - SetRecords" - txtDescri��o, True
        SetRecords = False
        Exit Function
    End If

    WS.CommitTrans 'Grava as altera��es ou inclus�es se n�o houverem erros
    
    'Se a janela Departamento-Se��o estiver aberta atualiza seus valores se necess�rio.
    If Not lInserir Then
        For Cont = 1 To Forms.Count - 1
            If Forms(Cont).Name = "frmDepartamentoSe��o" Then
                If Forms(Cont).txtC�digoDepartamento = txtC�digoAnterior Then
                    Forms(Cont).txtC�digoDepartamento = txtC�digo
                    Forms(Cont).txtDescri��oDepartamento = txtDescri��o
                    Forms(Cont).PosRecords
                End If
            End If
        Next
    End If
    
    If lInserir Then
        Log gUsu�rio, "Inclus�o - Departamento: " & txtC�digo & " - " & txtDescri��o
    Else
        Log gUsu�rio, "Altera��o - Departamento: " & txtC�digo & " - " & txtDescri��o
    End If
    
    SetRecords = True
End Function
Private Sub ZeraCampos()
    txtC�digo = Empty
    txtC�digoAnterior = Empty
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
    If Not DepartamentoAberto Then
        Unload Me
        Exit Sub
    End If
    TestaInferior TBLDepartamento, lAllowEdit, lAllowDelete, lAllowConsult
    TestaSuperior TBLDepartamento, lAllowEdit, lAllowDelete, lAllowConsult
    If TBLDepartamento.RecordCount = 0 Then
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
    
    lAllowInsert = Allow("DEPARTAMENTO", "I")
    lAllowEdit = Allow("DEPARTAMENTO", "A")
    lAllowDelete = Allow("DEPARTAMENTO", "E")
    lAllowConsult = Allow("DEPARTAMENTO", "C")
    
    ZeraCampos
    
    lPula = False
    lInserir = False
    lAlterar = False
    
    DepartamentoAberto = AbreTabela(Dicion�rio, "CADASTRO", "DEPARTAMENTO", DBCadastro, TBLDepartamento, TBLTabela, dbOpenTable)
    
    If DepartamentoAberto Then
        IndiceDepartamentoAtivo = "DEPARTAMENTO1"
        TBLDepartamento.Index = IndiceDepartamentoAtivo
    Else
        MsgBox "N�o consegui abrir a tabela 'Departamento' !", vbCritical, "Erro"
        Exit Sub
    End If
    
    Bot�oIncluir lAllowInsert
 
    If TBLDepartamento.RecordCount = 0 Then
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
        
    If TBLDepartamento.RecordCount = 0 Or TBLDepartamento.RecordCount = 1 Then
        Navega��oSuperior False
    Else
        Navega��oInferior lAllowConsult
    End If
    
    StatusBarAviso = "Pronto"
    Relat�rio = AddPath(Aplica��oPath, "REPORT\DEPARTAMENTO.RPT")
    TotalDatabaseName = 1
    DataBaseName(1) = AddPath(Aplica��oPath, "DATABASE\CADASTRO.MDB")
    mFechar = False
    Exit Sub
    
Erro:
    GeraMensagemDeErro "Departamento - Load"
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
    
    Set frmDepartamento = Nothing
End Sub
Private Sub Form_Unload(Cancel As Integer)
    If DepartamentoAberto Then
        TBLDepartamento.Close
    End If
    If Forms.Count = 2 Then
        AllBot�es False
    End If
End Sub
Private Sub txtC�digo_Change()
    If lPula Then
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
    If lPula Then
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
