VERSION 5.00
Begin VB.Form frmAgência 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Agência"
   ClientHeight    =   5670
   ClientLeft      =   2730
   ClientTop       =   1365
   ClientWidth     =   6330
   ClipControls    =   0   'False
   Icon            =   "Agência.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5670
   ScaleWidth      =   6330
   Begin VB.Frame frObservações 
      Caption         =   "Observações"
      Height          =   1185
      Left            =   0
      TabIndex        =   24
      Top             =   4050
      Width           =   6315
      Begin VB.TextBox txtObservações 
         Height          =   825
         Left            =   90
         MultiLine       =   -1  'True
         TabIndex        =   9
         Top             =   240
         Width           =   6105
      End
   End
   Begin VB.Frame frDadosCadastrais 
      Height          =   1335
      Left            =   0
      TabIndex        =   18
      Top             =   2700
      Width           =   6315
      Begin VB.TextBox txtFax 
         Height          =   285
         Left            =   5160
         TabIndex        =   8
         Top             =   930
         Width           =   1035
      End
      Begin VB.TextBox txtFone 
         Height          =   285
         Left            =   3060
         TabIndex        =   7
         Top             =   930
         Width           =   1035
      End
      Begin VB.TextBox txtDDD 
         Height          =   285
         Left            =   1200
         TabIndex        =   6
         Top             =   930
         Width           =   555
      End
      Begin VB.TextBox txtEndereço 
         Height          =   285
         Left            =   1200
         TabIndex        =   5
         Top             =   570
         Width           =   5000
      End
      Begin VB.TextBox txtContato 
         Height          =   285
         Left            =   1200
         TabIndex        =   4
         Top             =   210
         Width           =   5000
      End
      Begin VB.Label lblFax 
         Caption         =   "Fax"
         Height          =   255
         Left            =   4800
         TabIndex        =   23
         Top             =   960
         Width           =   315
      End
      Begin VB.Label lblFone 
         Caption         =   "Fone"
         Height          =   255
         Left            =   2640
         TabIndex        =   22
         Top             =   960
         Width           =   555
      End
      Begin VB.Label lblDDD 
         Caption         =   "DDD"
         Height          =   165
         Left            =   180
         TabIndex        =   21
         Top             =   960
         Width           =   405
      End
      Begin VB.Label lblEndereço 
         Caption         =   "Endereço"
         Height          =   195
         Left            =   150
         TabIndex        =   20
         Top             =   600
         Width           =   885
      End
      Begin VB.Label lblContato 
         Caption         =   "Contato"
         Height          =   195
         Left            =   150
         TabIndex        =   19
         Top             =   270
         Width           =   765
      End
   End
   Begin VB.Frame frAgência 
      Caption         =   "Agência"
      Height          =   1350
      Left            =   0
      TabIndex        =   15
      Top             =   1350
      Width           =   6330
      Begin VB.TextBox txtCódigoAgência 
         Height          =   285
         Left            =   1200
         TabIndex        =   2
         Top             =   300
         Width           =   1035
      End
      Begin VB.TextBox txtDescriçãoAgência 
         Height          =   285
         Left            =   1200
         TabIndex        =   3
         Top             =   750
         Width           =   5000
      End
      Begin VB.Label lblCódigoAgência 
         Caption         =   "Código"
         Height          =   210
         Left            =   150
         TabIndex        =   17
         Top             =   330
         Width           =   660
      End
      Begin VB.Label lblDescriçãoAgência 
         Caption         =   "Nome"
         Height          =   180
         Left            =   150
         TabIndex        =   16
         Top             =   780
         Width           =   960
      End
   End
   Begin VB.Frame frBanco 
      Caption         =   " Banco "
      Height          =   1350
      Left            =   0
      TabIndex        =   12
      Top             =   0
      Width           =   6330
      Begin VB.TextBox txtDescriçãoBanco 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1200
         TabIndex        =   1
         Top             =   750
         Width           =   5000
      End
      Begin VB.TextBox txtCódigoBanco 
         Height          =   285
         Left            =   1200
         TabIndex        =   0
         Top             =   300
         Width           =   700
      End
      Begin VB.Label lblDescriçãoBanco 
         Caption         =   "Descrição"
         Height          =   180
         Left            =   150
         TabIndex        =   14
         Top             =   780
         Width           =   960
      End
      Begin VB.Label lblCódigoBanco 
         Caption         =   "Código"
         Height          =   210
         Left            =   150
         TabIndex        =   13
         Top             =   330
         Width           =   660
      End
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Enabled         =   0   'False
      Height          =   345
      Left            =   5070
      TabIndex        =   11
      Top             =   5280
      Width           =   1245
   End
   Begin VB.CommandButton cmdGravar 
      Caption         =   "&Gravar"
      Enabled         =   0   'False
      Height          =   345
      Left            =   3780
      TabIndex        =   10
      Top             =   5280
      Width           =   1245
   End
End
Attribute VB_Name = "frmAgência"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim TBLBanco      As Table
Dim BancoAberto   As Boolean
Dim TBLAgência    As Table
Dim AgênciaAberto As Boolean
Dim IndiceAtivoBanco$, IndiceAtivoAgência$

Dim txtCódigoBancoAnterior$, txtCódigoAgênciaAnterior$

Dim lPula    As Boolean
Dim lInserir As Boolean
Dim lAlterar As Boolean
Dim mFechar  As Boolean

Dim lAllowInsert  As Boolean
Dim lAllowEdit    As Boolean
Dim lAllowDelete  As Boolean
Dim lAllowConsult As Boolean
Dim lAllowCagar   As Boolean

Dim StatusBarAviso$

Dim DataBaseName(1 To 1) As String
Public Relatório$

Public TotalDatabaseName%

Public lAtualizar As Boolean
Private Sub AtivaCampos()
    BotãoImprimir True
    frBanco.Enabled = True
    frAgência.Enabled = True
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
    
    If TBLAgência.RecordCount = 0 Then
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
    
    TestaInferior TBLAgência, lAllowEdit, lAllowDelete, lAllowConsult
    TestaSuperior TBLAgência, lAllowEdit, lAllowDelete, lAllowConsult
    
    GetRecords
End Function
Private Sub DesativaCampos()
    BotãoImprimir False
    frBanco.Enabled = False
    frAgência.Enabled = False
    cmdGravar.Enabled = False
    cmdCancelar.Enabled = False
    BotãoGravar False
End Sub
Public Sub Encontrar()
    If Not lAllowConsult Then
        Exit Sub
    End If
    Set frmEncontrar.DBBancoDeDados = DBFinanceiro
    frmEncontrar.NomeDaJanela = "Agência"
    frmEncontrar.LabelDescription = "Descrição"
    frmEncontrar.Mensagem = "Nenhuma agência foi selecionada!"
    frmEncontrar.BancoDeDados = "FINANCEIRO"
    frmEncontrar.Tabela = "AGÊNCIA"
    frmEncontrar.Indice = "1"
    frmEncontrar.CampoChave = "CÓDIGO DO BANCO,CÓDIGO"
    frmEncontrar.CampoPreencheLista = "DESCRIÇÃO"
    frmEncontrar.Show vbModal
    lPula = True
    txtCódigoBanco = GetWordSeparatedBy(frmEncontrar.Chave, 1)
    txtCódigoAgência = GetWordSeparatedBy(frmEncontrar.Chave, 2)
    lPula = False
    PosRecords
End Sub
Public Sub Excluir()
    Dim Confirmação As Integer, Msg1$, Msg2$
    Dim TBLContaCorrente As Table
    
    If lAlterar Then
       If Not Cancelamento Then
            Exit Sub
        End If
    End If
    
    If AbreTabela(Dicionário, "FINANCEIRO", "CONTA CORRENTE", DBFinanceiro, TBLContaCorrente, TBLTabela, dbOpenTable) Then
        TBLContaCorrente.Index = "CONTACORRENTE1"
        TBLContaCorrente.Seek ">=", txtCódigoBanco, txtCódigoAgência
        If Not TBLContaCorrente.NoMatch Then
            If TBLContaCorrente("CÓDIGO DO BANCO") = txtCódigoBanco And TBLContaCorrente("CÓDIGO DA AGÊNCIA") = txtCódigoAgência Then
                MsgBox "Relação violada!" + vbCr + "Para apagar esta agência, antes é necessário apagar" + vbCr + "todas as conta correntes dela dependente.", vbExclamation, "Aviso"
                TBLContaCorrente.Close
                Exit Sub
            End If
        End If
    Else
        Exit Sub
    End If
    TBLContaCorrente.Close
    
    StatusBarAviso = "Exclusão"
    BarraDeStatus StatusBarAviso
    
    Msg1 = "Você está preste a apagar um registro !"
    Msg2 = "Tem certeza?"
    Msg2 = String(((Len(Msg1) - Len(Msg2)) / 2), " ") + Msg2
    Confirmação = MsgBox(Msg1 + vbCr + Msg2, vbYesNo + vbQuestion + vbDefaultButton2, "Confirmação")
    
    If Confirmação = vbNo Then
        StatusBarAviso = "Exclusão cancelada"
        BarraDeStatus StatusBarAviso
        Exit Sub
    End If
    
    WS.BeginTrans
    
    TBLAgência.Delete
    
    If Err <> 0 Then
        GeraMensagemDeErro "Agência - Excluir - " & txtDescriçãoAgência, True
        StatusBarAviso = "Falha na exclusão"
        BarraDeStatus StatusBarAviso
        Exit Sub
    End If
    
    WS.CommitTrans
    
    Log gUsuário, "Exclusão - Agência" & vbCr & "Banco:" & txtCódigoBanco & " - " & vbCr & "Agência: " & txtCódigoAgência & " - " & txtDescriçãoAgência
    
    StatusBarAviso = "Exclusão bem sucedida"
    BarraDeStatus StatusBarAviso
    
    If TBLAgência.RecordCount = 0 Then
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
    
    If TBLAgência.BOF Then
        TBLAgência.MoveFirst
    ElseIf TBLAgência.EOF Then
        TBLAgência.MoveLast
    Else
        TBLAgência.MovePrevious
        If TBLAgência.BOF Then
            TBLAgência.MoveNext
        End If
    End If
    
    GetRecords
    
    TestaInferior TBLAgência, lAllowEdit, lAllowDelete, lAllowConsult
    TestaSuperior TBLAgência, lAllowEdit, lAllowDelete, lAllowConsult
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
        If TBLAgência.RecordCount > 0 And Not TBLAgência.BOF And Not TBLAgência.EOF Then
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
    
    TestaInferior TBLAgência, lAllowEdit, lAllowDelete, lAllowConsult
    TestaSuperior TBLAgência, lAllowEdit, lAllowDelete, lAllowConsult
    
    If TBLAgência.RecordCount = 0 Then
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
    
    If txtCódigoBanco.Enabled Then
        txtCódigoBanco.SetFocus
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
    
    txtCódigoBanco.SetFocus
End Sub
Public Sub MoveFirst()
    If lAlterar Then
       If Not Cancelamento Then
            Exit Sub
        End If
    End If
    
    StatusBarAviso = "Pronto"
    BarraDeStatus StatusBarAviso
    
    TBLAgência.MoveFirst
    
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
    
    TBLAgência.MoveLast
    
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
    
    TBLAgência.MoveNext
    If TBLAgência.EOF Then
        TBLAgência.MovePrevious
        Exit Sub
    End If
    
    StatusBarAviso = "Pronto"
    BarraDeStatus StatusBarAviso
    
    NavegaçãoInferior lAllowConsult
    TestaSuperior TBLAgência, lAllowEdit, lAllowDelete, lAllowConsult
    
    GetRecords
End Sub
Public Sub MovePrevious()
    If lAlterar Then
       If Not Cancelamento Then
            Exit Sub
        End If
    End If
    
    TBLAgência.MovePrevious
    If TBLAgência.BOF Then
        TBLAgência.MoveNext
        Exit Sub
    End If
    
    StatusBarAviso = "Pronto"
    BarraDeStatus StatusBarAviso
    
    NavegaçãoSuperior lAllowConsult
    TestaInferior TBLAgência, lAllowEdit, lAllowDelete, lAllowConsult
    
    GetRecords
End Sub
Public Sub PosRecords()
    If TBLAgência.RecordCount = 0 Then
        Exit Sub
    End If
    
    TBLAgência.Seek "=", txtCódigoBanco, txtCódigoAgência
    If TBLAgência.NoMatch Then
        MsgBox "Não consegui encontrar a agência" + txtCódigoAgência, vbExclamation, "Erro"
        TBLAgência.MoveFirst
        NavegaçãoInferior False
        NavegaçãoInferior lAllowConsult
    Else
        TestaInferior TBLAgência, lAllowEdit, lAllowDelete, lAllowConsult
        TestaSuperior TBLAgência, lAllowEdit, lAllowDelete, lAllowConsult
    End If
    GetRecords
End Sub
Public Function PushDataBaseName(ByVal Posição As Integer) As String
    PushDataBaseName = DataBaseName(Posição)
End Function
Private Sub GetRecords()
    On Error GoTo Erro
    
    If Not lAllowConsult Then
        ZeraCampos
        DesativaCampos
        Exit Sub
    End If
    txtCódigoBanco = TBLAgência("CÓDIGO DO BANCO")
    txtCódigoBancoAnterior = txtCódigoBanco
    TBLBanco.Seek "=", txtCódigoBanco
    txtDescriçãoBanco = TBLBanco("DESCRIÇÃO")
    txtCódigoAgência = TBLAgência("CÓDIGO")
    txtCódigoAgênciaAnterior = txtCódigoAgência
    txtDescriçãoAgência = TBLAgência("DESCRIÇÃO")
    txtContato = TBLAgência("CONTATO")
    txtEndereço = TBLAgência("ENDEREÇO")
    txtDDD = TBLAgência("DDD")
    txtFone = TBLAgência("TELEFONE")
    txtFax = TBLAgência("FAX")
    txtObservações = TBLAgência("OBSERVAÇÕES")
    If Not lAllowEdit Then
        DesativaCampos
    End If
    
    Exit Sub
    
Erro:
    GeraMensagemDeErro "Agência - GetRecords "
    Resume Next
End Sub
Private Function SetRecords()
    On Error GoTo Erro
    
    Dim Msg$
    Dim Confirmação As Integer, Msg1$, Msg2$, AchouContaCorrente As Boolean
    Dim TBLContaCorrente As Table
    Dim SQL As String
    Dim Cont%
    
    If ((txtCódigoBanco <> txtCódigoBancoAnterior) Or (txtCódigoAgência <> txtCódigoAgênciaAnterior)) And Not lInserir Then
        If AbreTabela(Dicionário, "FINANCEIRO", "CONTA CORRENTE", DBFinanceiro, TBLContaCorrente, TBLTabela, dbOpenTable) Then
            TBLContaCorrente.Index = "CONTACORRENTE1"
            TBLContaCorrente.Seek ">=", txtCódigoBancoAnterior, txtCódigoAgênciaAnterior
            If Not TBLContaCorrente.NoMatch Then
                If TBLContaCorrente("CÓDIGO DO BANCO") = txtCódigoBancoAnterior And TBLContaCorrente("CÓDIGO DA AGÊNCIA") = txtCódigoAgênciaAnterior Then
                    AchouContaCorrente = True
                    Confirmação = MsgBox("Você necessita alterar as contas correntes relacionadas com esta agência !" + vbCr + "Deseja realizar agora as alterações de" + vbCr + "todas as contas dela dependente?", vbYesNo + vbQuestion + vbDefaultButton2, "Aviso")
                End If
            Else
                AchouContaCorrente = False
            End If
        Else
            Exit Function
        End If
        TBLContaCorrente.Close
        
        If AchouContaCorrente Then
            If Confirmação = vbNo Then
                SetRecords = False
                Exit Function
            End If
        End If
    Else
        AchouContaCorrente = False
    End If
    
    On Error GoTo ErroInclusão
    
    WS.BeginTrans 'Inicia transações
    
    If lInserir Then
        TBLAgência.AddNew
    Else
        TBLAgência.Edit
    End If
    
    TBLAgência("CÓDIGO DO BANCO") = txtCódigoBanco
    TBLAgência("CÓDIGO") = txtCódigoAgência
    TBLAgência("DESCRIÇÃO") = txtDescriçãoAgência
    TBLAgência("CONTATO") = txtContato
    TBLAgência("ENDEREÇO") = txtEndereço
    TBLAgência("DDD") = txtDDD
    TBLAgência("TELEFONE") = txtFone
    TBLAgência("FAX") = txtFax
    TBLAgência("OBSERVAÇÕES") = txtObservações
    If lInserir Then
        TBLAgência("USERNAME - CRIA") = gUsuário
        TBLAgência("DATA - CRIA") = Date
        TBLAgência("HORA - CRIA") = Time
        TBLAgência("USERNAME - ALTERA") = "VAZIO"
        TBLAgência("DATA - ALTERA") = vbNull
        TBLAgência("HORA - ALTERA") = vbNull
    End If
    If lAlterar Then
        TBLAgência("USERNAME - ALTERA") = gUsuário
        TBLAgência("DATA - ALTERA") = Date
        TBLAgência("HORA - ALTERA") = Time
    End If
    TBLAgência.Update
    
    If AchouContaCorrente Then
        SQL = "Update [CONTA CORRENTE] Set [CÓDIGO DO BANCO]= '" + txtCódigoBanco + "',[CÓDIGO DA AGÊNCIA]= '" + txtCódigoAgência + "' Where [CÓDIGO DA AGÊNCIA]= '" + txtCódigoAgênciaAnterior + "' AND [CÓDIGO DO BANCO] = '" + txtCódigoBancoAnterior + "'"
        DBFinanceiro.Execute SQL
    End If
    
    WS.CommitTrans 'Grava as alterações ou inclusões se não houverem erros
    
    'Se a janela Agência estiver aberta atualiza seus valores se necessário.
    If Not lInserir Then
        For Cont = 1 To Forms.Count - 1
            If Forms(Cont).Name = "frmContaCorrente" Then
                If Forms(Cont).txtCódigoBanco = txtCódigoBancoAnterior Then
                    Forms(Cont).txtCódigoBanco = txtCódigoBanco
                    Forms(Cont).txtDescriçãoBanco = txtDescriçãoBanco
                    If Forms(Cont).txtCódigoAgência = txtCódigoAgênciaAnterior Then
                        Forms(Cont).txtCódigoAgência = txtCódigoAgência
                        Forms(Cont).txtDescriçãoAgência = txtDescriçãoAgência
                    End If
                    Forms(Cont).PosRecords
                End If
            End If
        Next
    End If
    
    SetRecords = True
    
    If lInserir Then
        Log gUsuário, "Inclusão - Agência" & vbCr & "Banco:" & txtCódigoBanco & " - " & vbCr & "Agência: " & txtCódigoAgência & " - " & txtDescriçãoAgência
    Else
        Log gUsuário, "Alteração - Agência" & vbCr & "Banco:" & txtCódigoBanco & " - " & vbCr & "Agência: " & txtCódigoAgência & " - " & txtDescriçãoAgência
    End If
    
    Exit Function
    
Erro:
    GeraMensagemDeErro "Agência - SetRecords - " & txtDescriçãoAgência
    SetRecords = False
    Exit Function
    
ErroInclusão:
    TBLAgência.CancelUpdate
    GeraMensagemDeErro "Agência - SetRecords - " & txtDescriçãoAgência, True
    SetRecords = False
    Exit Function
End Function
Private Sub ZeraCampos()
    txtCódigoBanco = Empty
    txtCódigoBancoAnterior = Empty
    txtCódigoAgência = Empty
    txtCódigoAgênciaAnterior = Empty
    txtDescriçãoBanco = Empty
    txtDescriçãoAgência = Empty
    txtContato = Empty
    txtEndereço = Empty
    txtDDD = Empty
    txtFone = Empty
    txtFax = Empty
    txtObservações = Empty
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
    If Not AgênciaAberto Then
        Unload Me
        Exit Sub
    End If
    
    TestaInferior TBLAgência, lAllowEdit, lAllowDelete, lAllowConsult
    TestaSuperior TBLAgência, lAllowEdit, lAllowDelete, lAllowConsult
    
    If TBLAgência.RecordCount = 0 Then
        cmdGravar.Enabled = False
        cmdCancelar.Enabled = False
        BotãoImprimir False
    Else
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
    
    lAllowInsert = Allow("AGÊNCIA", "I")
    lAllowEdit = Allow("AGÊNCIA", "A")
    lAllowDelete = Allow("AGÊNCIA", "E")
    lAllowConsult = Allow("AGÊNCIA", "C")
    
    ZeraCampos
    
    lInserir = False
    lAlterar = False
    lPula = False
    
    BancoAberto = AbreTabela(Dicionário, "FINANCEIRO", "BANCO", DBFinanceiro, TBLBanco, TBLTabela, dbOpenTable)
    
    If BancoAberto Then
        IndiceAtivoBanco = "BANCO1"
        TBLBanco.Index = IndiceAtivoBanco
    Else
        MsgBox "Não consegui abrir a tabela 'Banco' !", vbCritical, "Erro"
        Exit Sub
    End If
    
    AgênciaAberto = AbreTabela(Dicionário, "FINANCEIRO", "AGÊNCIA", DBFinanceiro, TBLAgência, TBLTabela, dbOpenTable)
    
    If AgênciaAberto Then
        IndiceAtivoAgência = "AGÊNCIA1"
        TBLAgência.Index = IndiceAtivoAgência
    Else
        MsgBox "Não consegui abrir a tabela 'Agência' !", vbCritical, "Erro"
        Exit Sub
    End If
    
    BotãoIncluir lAllowInsert
 
    If TBLAgência.RecordCount = 0 Then
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
        
    If TBLAgência.RecordCount = 0 Or TBLAgência.RecordCount = 1 Then
        NavegaçãoSuperior False
    Else
        NavegaçãoInferior lAllowConsult
    End If
        
    StatusBarAviso = "Pronto"
    Relatório = AddPath(AplicaçãoPath, "REPORT\AGÊNCIA.RPT")
    TotalDatabaseName = 1
    DataBaseName(1) = AddPath(AplicaçãoPath, "DATABASE\FINANCEIRO.MDB")
    mFechar = False
    
    Exit Sub
    
Erro:
    GeraMensagemDeErro "Agência - Load"
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
    
    Set frmAgência = Nothing
End Sub
Private Sub Form_Unload(Cancel As Integer)
    If BancoAberto Then
        TBLBanco.Close
    End If
    If AgênciaAberto Then
        TBLAgência.Close
    End If
    If Forms.Count = 2 Then
        AllBotões False
    End If
End Sub
Private Sub txtCódigoAgência_Change()
    If Not lPula Then
        FormatMask "@S10", txtCódigoAgência
    End If
End Sub
Private Sub txtCódigoAgência_KeyPress(KeyAscii As Integer)
    If Not lInserir Then
        lAlterar = True
        StatusBarAviso = "Alteração"
        BarraDeStatus StatusBarAviso
    End If
End Sub
Private Sub txtCódigoAgência_LostFocus()
    If txtCódigoAgência.Enabled Then
        LeftBlank txtCódigoAgência
    End If
End Sub
Private Sub txtCódigoBanco_Change()
    FormatMask "9999", txtCódigoBanco
End Sub
Private Sub txtCódigoBanco_KeyPress(KeyAscii As Integer)
    If Not lInserir Then
        lAlterar = True
        StatusBarAviso = "Alteração"
        BarraDeStatus StatusBarAviso
    End If
End Sub
Private Sub txtCódigoBanco_LostFocus()
    If mdiGeal.ActiveForm.Name = "frmAgência" Then
        If txtCódigoBanco.Enabled Then
            LeftBlank txtCódigoBanco
            TBLBanco.Seek "=", txtCódigoBanco
            If TBLBanco.NoMatch Then
                MsgBox "Não encontrei o banco !" + txtCódigoBanco, vbExclamation, "Aviso"
                txtCódigoBanco = Empty
                txtCódigoBanco.SetFocus
                Exit Sub
            End If
            txtDescriçãoBanco = TBLBanco("DESCRIÇÃO")
        End If
    End If
End Sub
Private Sub txtContato_KeyPress(KeyAscii As Integer)
    If Not lInserir Then
        lAlterar = True
        StatusBarAviso = "Alteração"
        BarraDeStatus StatusBarAviso
    End If
End Sub
Private Sub txtDDD_Change()
    If Not lPula Then
        FormatMask "9999", txtDDD
    End If
End Sub
Private Sub txtDDD_KeyPress(KeyAscii As Integer)
    If Not lInserir Then
        lAlterar = True
        StatusBarAviso = "Alteração"
        BarraDeStatus StatusBarAviso
    End If
End Sub
Private Sub txtDescriçãoAgência_Change()
    If Not lPula Then
        FormatMask "@!S30", txtDescriçãoAgência
    End If
End Sub
Private Sub txtDescriçãoAgência_KeyPress(KeyAscii As Integer)
    If Not lInserir Then
        lAlterar = True
        StatusBarAviso = "Alteração"
        BarraDeStatus StatusBarAviso
    End If
End Sub
Private Sub txtEndereço_KeyPress(KeyAscii As Integer)
    If Not lInserir Then
        lAlterar = True
        StatusBarAviso = "Alteração"
        BarraDeStatus StatusBarAviso
    End If
End Sub
Private Sub txtFax_Change()
    If Not lPula Then
        FormatMask "####-####", txtFax
    End If
End Sub
Private Sub txtFax_KeyPress(KeyAscii As Integer)
    If Not lInserir Then
        lAlterar = True
        StatusBarAviso = "Alteração"
        BarraDeStatus StatusBarAviso
    End If
End Sub
Private Sub txtFone_Change()
    If Not lPula Then
        FormatMask "####-####", txtFone
    End If
End Sub
Private Sub txtFone_KeyPress(KeyAscii As Integer)
    If Not lInserir Then
        lAlterar = True
        StatusBarAviso = "Alteração"
        BarraDeStatus StatusBarAviso
    End If
End Sub
Private Sub txtObservações_KeyPress(KeyAscii As Integer)
    If Not lInserir Then
        lAlterar = True
        StatusBarAviso = "Alteração"
        BarraDeStatus StatusBarAviso
    End If
End Sub
