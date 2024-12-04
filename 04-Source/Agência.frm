VERSION 5.00
Begin VB.Form frmAg�ncia 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ag�ncia"
   ClientHeight    =   5670
   ClientLeft      =   2730
   ClientTop       =   1365
   ClientWidth     =   6330
   ClipControls    =   0   'False
   Icon            =   "Ag�ncia.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5670
   ScaleWidth      =   6330
   Begin VB.Frame frObserva��es 
      Caption         =   "Observa��es"
      Height          =   1185
      Left            =   0
      TabIndex        =   24
      Top             =   4050
      Width           =   6315
      Begin VB.TextBox txtObserva��es 
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
      Begin VB.TextBox txtEndere�o 
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
      Begin VB.Label lblEndere�o 
         Caption         =   "Endere�o"
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
   Begin VB.Frame frAg�ncia 
      Caption         =   "Ag�ncia"
      Height          =   1350
      Left            =   0
      TabIndex        =   15
      Top             =   1350
      Width           =   6330
      Begin VB.TextBox txtC�digoAg�ncia 
         Height          =   285
         Left            =   1200
         TabIndex        =   2
         Top             =   300
         Width           =   1035
      End
      Begin VB.TextBox txtDescri��oAg�ncia 
         Height          =   285
         Left            =   1200
         TabIndex        =   3
         Top             =   750
         Width           =   5000
      End
      Begin VB.Label lblC�digoAg�ncia 
         Caption         =   "C�digo"
         Height          =   210
         Left            =   150
         TabIndex        =   17
         Top             =   330
         Width           =   660
      End
      Begin VB.Label lblDescri��oAg�ncia 
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
      Begin VB.TextBox txtDescri��oBanco 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1200
         TabIndex        =   1
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
         TabIndex        =   14
         Top             =   780
         Width           =   960
      End
      Begin VB.Label lblC�digoBanco 
         Caption         =   "C�digo"
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
Attribute VB_Name = "frmAg�ncia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim TBLBanco      As Table
Dim BancoAberto   As Boolean
Dim TBLAg�ncia    As Table
Dim Ag�nciaAberto As Boolean
Dim IndiceAtivoBanco$, IndiceAtivoAg�ncia$

Dim txtC�digoBancoAnterior$, txtC�digoAg�nciaAnterior$

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
Public Relat�rio$

Public TotalDatabaseName%

Public lAtualizar As Boolean
Private Sub AtivaCampos()
    Bot�oImprimir True
    frBanco.Enabled = True
    frAg�ncia.Enabled = True
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
    
    If TBLAg�ncia.RecordCount = 0 Then
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
    
    TestaInferior TBLAg�ncia, lAllowEdit, lAllowDelete, lAllowConsult
    TestaSuperior TBLAg�ncia, lAllowEdit, lAllowDelete, lAllowConsult
    
    GetRecords
End Function
Private Sub DesativaCampos()
    Bot�oImprimir False
    frBanco.Enabled = False
    frAg�ncia.Enabled = False
    cmdGravar.Enabled = False
    cmdCancelar.Enabled = False
    Bot�oGravar False
End Sub
Public Sub Encontrar()
    If Not lAllowConsult Then
        Exit Sub
    End If
    Set frmEncontrar.DBBancoDeDados = DBFinanceiro
    frmEncontrar.NomeDaJanela = "Ag�ncia"
    frmEncontrar.LabelDescription = "Descri��o"
    frmEncontrar.Mensagem = "Nenhuma ag�ncia foi selecionada!"
    frmEncontrar.BancoDeDados = "FINANCEIRO"
    frmEncontrar.Tabela = "AG�NCIA"
    frmEncontrar.Indice = "1"
    frmEncontrar.CampoChave = "C�DIGO DO BANCO,C�DIGO"
    frmEncontrar.CampoPreencheLista = "DESCRI��O"
    frmEncontrar.Show vbModal
    lPula = True
    txtC�digoBanco = GetWordSeparatedBy(frmEncontrar.Chave, 1)
    txtC�digoAg�ncia = GetWordSeparatedBy(frmEncontrar.Chave, 2)
    lPula = False
    PosRecords
End Sub
Public Sub Excluir()
    Dim Confirma��o As Integer, Msg1$, Msg2$
    Dim TBLContaCorrente As Table
    
    If lAlterar Then
       If Not Cancelamento Then
            Exit Sub
        End If
    End If
    
    If AbreTabela(Dicion�rio, "FINANCEIRO", "CONTA CORRENTE", DBFinanceiro, TBLContaCorrente, TBLTabela, dbOpenTable) Then
        TBLContaCorrente.Index = "CONTACORRENTE1"
        TBLContaCorrente.Seek ">=", txtC�digoBanco, txtC�digoAg�ncia
        If Not TBLContaCorrente.NoMatch Then
            If TBLContaCorrente("C�DIGO DO BANCO") = txtC�digoBanco And TBLContaCorrente("C�DIGO DA AG�NCIA") = txtC�digoAg�ncia Then
                MsgBox "Rela��o violada!" + vbCr + "Para apagar esta ag�ncia, antes � necess�rio apagar" + vbCr + "todas as conta correntes dela dependente.", vbExclamation, "Aviso"
                TBLContaCorrente.Close
                Exit Sub
            End If
        End If
    Else
        Exit Sub
    End If
    TBLContaCorrente.Close
    
    StatusBarAviso = "Exclus�o"
    BarraDeStatus StatusBarAviso
    
    Msg1 = "Voc� est� preste a apagar um registro !"
    Msg2 = "Tem certeza?"
    Msg2 = String(((Len(Msg1) - Len(Msg2)) / 2), " ") + Msg2
    Confirma��o = MsgBox(Msg1 + vbCr + Msg2, vbYesNo + vbQuestion + vbDefaultButton2, "Confirma��o")
    
    If Confirma��o = vbNo Then
        StatusBarAviso = "Exclus�o cancelada"
        BarraDeStatus StatusBarAviso
        Exit Sub
    End If
    
    WS.BeginTrans
    
    TBLAg�ncia.Delete
    
    If Err <> 0 Then
        GeraMensagemDeErro "Ag�ncia - Excluir - " & txtDescri��oAg�ncia, True
        StatusBarAviso = "Falha na exclus�o"
        BarraDeStatus StatusBarAviso
        Exit Sub
    End If
    
    WS.CommitTrans
    
    Log gUsu�rio, "Exclus�o - Ag�ncia" & vbCr & "Banco:" & txtC�digoBanco & " - " & vbCr & "Ag�ncia: " & txtC�digoAg�ncia & " - " & txtDescri��oAg�ncia
    
    StatusBarAviso = "Exclus�o bem sucedida"
    BarraDeStatus StatusBarAviso
    
    If TBLAg�ncia.RecordCount = 0 Then
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
    
    If TBLAg�ncia.BOF Then
        TBLAg�ncia.MoveFirst
    ElseIf TBLAg�ncia.EOF Then
        TBLAg�ncia.MoveLast
    Else
        TBLAg�ncia.MovePrevious
        If TBLAg�ncia.BOF Then
            TBLAg�ncia.MoveNext
        End If
    End If
    
    GetRecords
    
    TestaInferior TBLAg�ncia, lAllowEdit, lAllowDelete, lAllowConsult
    TestaSuperior TBLAg�ncia, lAllowEdit, lAllowDelete, lAllowConsult
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
        If TBLAg�ncia.RecordCount > 0 And Not TBLAg�ncia.BOF And Not TBLAg�ncia.EOF Then
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
    
    TestaInferior TBLAg�ncia, lAllowEdit, lAllowDelete, lAllowConsult
    TestaSuperior TBLAg�ncia, lAllowEdit, lAllowDelete, lAllowConsult
    
    If TBLAg�ncia.RecordCount = 0 Then
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
    
    TBLAg�ncia.MoveFirst
    
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
    
    TBLAg�ncia.MoveLast
    
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
    
    TBLAg�ncia.MoveNext
    If TBLAg�ncia.EOF Then
        TBLAg�ncia.MovePrevious
        Exit Sub
    End If
    
    StatusBarAviso = "Pronto"
    BarraDeStatus StatusBarAviso
    
    Navega��oInferior lAllowConsult
    TestaSuperior TBLAg�ncia, lAllowEdit, lAllowDelete, lAllowConsult
    
    GetRecords
End Sub
Public Sub MovePrevious()
    If lAlterar Then
       If Not Cancelamento Then
            Exit Sub
        End If
    End If
    
    TBLAg�ncia.MovePrevious
    If TBLAg�ncia.BOF Then
        TBLAg�ncia.MoveNext
        Exit Sub
    End If
    
    StatusBarAviso = "Pronto"
    BarraDeStatus StatusBarAviso
    
    Navega��oSuperior lAllowConsult
    TestaInferior TBLAg�ncia, lAllowEdit, lAllowDelete, lAllowConsult
    
    GetRecords
End Sub
Public Sub PosRecords()
    If TBLAg�ncia.RecordCount = 0 Then
        Exit Sub
    End If
    
    TBLAg�ncia.Seek "=", txtC�digoBanco, txtC�digoAg�ncia
    If TBLAg�ncia.NoMatch Then
        MsgBox "N�o consegui encontrar a ag�ncia" + txtC�digoAg�ncia, vbExclamation, "Erro"
        TBLAg�ncia.MoveFirst
        Navega��oInferior False
        Navega��oInferior lAllowConsult
    Else
        TestaInferior TBLAg�ncia, lAllowEdit, lAllowDelete, lAllowConsult
        TestaSuperior TBLAg�ncia, lAllowEdit, lAllowDelete, lAllowConsult
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
    txtC�digoBanco = TBLAg�ncia("C�DIGO DO BANCO")
    txtC�digoBancoAnterior = txtC�digoBanco
    TBLBanco.Seek "=", txtC�digoBanco
    txtDescri��oBanco = TBLBanco("DESCRI��O")
    txtC�digoAg�ncia = TBLAg�ncia("C�DIGO")
    txtC�digoAg�nciaAnterior = txtC�digoAg�ncia
    txtDescri��oAg�ncia = TBLAg�ncia("DESCRI��O")
    txtContato = TBLAg�ncia("CONTATO")
    txtEndere�o = TBLAg�ncia("ENDERE�O")
    txtDDD = TBLAg�ncia("DDD")
    txtFone = TBLAg�ncia("TELEFONE")
    txtFax = TBLAg�ncia("FAX")
    txtObserva��es = TBLAg�ncia("OBSERVA��ES")
    If Not lAllowEdit Then
        DesativaCampos
    End If
    
    Exit Sub
    
Erro:
    GeraMensagemDeErro "Ag�ncia - GetRecords "
    Resume Next
End Sub
Private Function SetRecords()
    On Error GoTo Erro
    
    Dim Msg$
    Dim Confirma��o As Integer, Msg1$, Msg2$, AchouContaCorrente As Boolean
    Dim TBLContaCorrente As Table
    Dim SQL As String
    Dim Cont%
    
    If ((txtC�digoBanco <> txtC�digoBancoAnterior) Or (txtC�digoAg�ncia <> txtC�digoAg�nciaAnterior)) And Not lInserir Then
        If AbreTabela(Dicion�rio, "FINANCEIRO", "CONTA CORRENTE", DBFinanceiro, TBLContaCorrente, TBLTabela, dbOpenTable) Then
            TBLContaCorrente.Index = "CONTACORRENTE1"
            TBLContaCorrente.Seek ">=", txtC�digoBancoAnterior, txtC�digoAg�nciaAnterior
            If Not TBLContaCorrente.NoMatch Then
                If TBLContaCorrente("C�DIGO DO BANCO") = txtC�digoBancoAnterior And TBLContaCorrente("C�DIGO DA AG�NCIA") = txtC�digoAg�nciaAnterior Then
                    AchouContaCorrente = True
                    Confirma��o = MsgBox("Voc� necessita alterar as contas correntes relacionadas com esta ag�ncia !" + vbCr + "Deseja realizar agora as altera��es de" + vbCr + "todas as contas dela dependente?", vbYesNo + vbQuestion + vbDefaultButton2, "Aviso")
                End If
            Else
                AchouContaCorrente = False
            End If
        Else
            Exit Function
        End If
        TBLContaCorrente.Close
        
        If AchouContaCorrente Then
            If Confirma��o = vbNo Then
                SetRecords = False
                Exit Function
            End If
        End If
    Else
        AchouContaCorrente = False
    End If
    
    On Error GoTo ErroInclus�o
    
    WS.BeginTrans 'Inicia transa��es
    
    If lInserir Then
        TBLAg�ncia.AddNew
    Else
        TBLAg�ncia.Edit
    End If
    
    TBLAg�ncia("C�DIGO DO BANCO") = txtC�digoBanco
    TBLAg�ncia("C�DIGO") = txtC�digoAg�ncia
    TBLAg�ncia("DESCRI��O") = txtDescri��oAg�ncia
    TBLAg�ncia("CONTATO") = txtContato
    TBLAg�ncia("ENDERE�O") = txtEndere�o
    TBLAg�ncia("DDD") = txtDDD
    TBLAg�ncia("TELEFONE") = txtFone
    TBLAg�ncia("FAX") = txtFax
    TBLAg�ncia("OBSERVA��ES") = txtObserva��es
    If lInserir Then
        TBLAg�ncia("USERNAME - CRIA") = gUsu�rio
        TBLAg�ncia("DATA - CRIA") = Date
        TBLAg�ncia("HORA - CRIA") = Time
        TBLAg�ncia("USERNAME - ALTERA") = "VAZIO"
        TBLAg�ncia("DATA - ALTERA") = vbNull
        TBLAg�ncia("HORA - ALTERA") = vbNull
    End If
    If lAlterar Then
        TBLAg�ncia("USERNAME - ALTERA") = gUsu�rio
        TBLAg�ncia("DATA - ALTERA") = Date
        TBLAg�ncia("HORA - ALTERA") = Time
    End If
    TBLAg�ncia.Update
    
    If AchouContaCorrente Then
        SQL = "Update [CONTA CORRENTE] Set [C�DIGO DO BANCO]= '" + txtC�digoBanco + "',[C�DIGO DA AG�NCIA]= '" + txtC�digoAg�ncia + "' Where [C�DIGO DA AG�NCIA]= '" + txtC�digoAg�nciaAnterior + "' AND [C�DIGO DO BANCO] = '" + txtC�digoBancoAnterior + "'"
        DBFinanceiro.Execute SQL
    End If
    
    WS.CommitTrans 'Grava as altera��es ou inclus�es se n�o houverem erros
    
    'Se a janela Ag�ncia estiver aberta atualiza seus valores se necess�rio.
    If Not lInserir Then
        For Cont = 1 To Forms.Count - 1
            If Forms(Cont).Name = "frmContaCorrente" Then
                If Forms(Cont).txtC�digoBanco = txtC�digoBancoAnterior Then
                    Forms(Cont).txtC�digoBanco = txtC�digoBanco
                    Forms(Cont).txtDescri��oBanco = txtDescri��oBanco
                    If Forms(Cont).txtC�digoAg�ncia = txtC�digoAg�nciaAnterior Then
                        Forms(Cont).txtC�digoAg�ncia = txtC�digoAg�ncia
                        Forms(Cont).txtDescri��oAg�ncia = txtDescri��oAg�ncia
                    End If
                    Forms(Cont).PosRecords
                End If
            End If
        Next
    End If
    
    SetRecords = True
    
    If lInserir Then
        Log gUsu�rio, "Inclus�o - Ag�ncia" & vbCr & "Banco:" & txtC�digoBanco & " - " & vbCr & "Ag�ncia: " & txtC�digoAg�ncia & " - " & txtDescri��oAg�ncia
    Else
        Log gUsu�rio, "Altera��o - Ag�ncia" & vbCr & "Banco:" & txtC�digoBanco & " - " & vbCr & "Ag�ncia: " & txtC�digoAg�ncia & " - " & txtDescri��oAg�ncia
    End If
    
    Exit Function
    
Erro:
    GeraMensagemDeErro "Ag�ncia - SetRecords - " & txtDescri��oAg�ncia
    SetRecords = False
    Exit Function
    
ErroInclus�o:
    TBLAg�ncia.CancelUpdate
    GeraMensagemDeErro "Ag�ncia - SetRecords - " & txtDescri��oAg�ncia, True
    SetRecords = False
    Exit Function
End Function
Private Sub ZeraCampos()
    txtC�digoBanco = Empty
    txtC�digoBancoAnterior = Empty
    txtC�digoAg�ncia = Empty
    txtC�digoAg�nciaAnterior = Empty
    txtDescri��oBanco = Empty
    txtDescri��oAg�ncia = Empty
    txtContato = Empty
    txtEndere�o = Empty
    txtDDD = Empty
    txtFone = Empty
    txtFax = Empty
    txtObserva��es = Empty
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
    
    TestaInferior TBLAg�ncia, lAllowEdit, lAllowDelete, lAllowConsult
    TestaSuperior TBLAg�ncia, lAllowEdit, lAllowDelete, lAllowConsult
    
    If TBLAg�ncia.RecordCount = 0 Then
        cmdGravar.Enabled = False
        cmdCancelar.Enabled = False
        Bot�oImprimir False
    Else
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
    
    lAllowInsert = Allow("AG�NCIA", "I")
    lAllowEdit = Allow("AG�NCIA", "A")
    lAllowDelete = Allow("AG�NCIA", "E")
    lAllowConsult = Allow("AG�NCIA", "C")
    
    ZeraCampos
    
    lInserir = False
    lAlterar = False
    lPula = False
    
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
    
    Bot�oIncluir lAllowInsert
 
    If TBLAg�ncia.RecordCount = 0 Then
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
        
    If TBLAg�ncia.RecordCount = 0 Or TBLAg�ncia.RecordCount = 1 Then
        Navega��oSuperior False
    Else
        Navega��oInferior lAllowConsult
    End If
        
    StatusBarAviso = "Pronto"
    Relat�rio = AddPath(Aplica��oPath, "REPORT\AG�NCIA.RPT")
    TotalDatabaseName = 1
    DataBaseName(1) = AddPath(Aplica��oPath, "DATABASE\FINANCEIRO.MDB")
    mFechar = False
    
    Exit Sub
    
Erro:
    GeraMensagemDeErro "Ag�ncia - Load"
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
    
    Set frmAg�ncia = Nothing
End Sub
Private Sub Form_Unload(Cancel As Integer)
    If BancoAberto Then
        TBLBanco.Close
    End If
    If Ag�nciaAberto Then
        TBLAg�ncia.Close
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
    If txtC�digoAg�ncia.Enabled Then
        LeftBlank txtC�digoAg�ncia
    End If
End Sub
Private Sub txtC�digoBanco_Change()
    FormatMask "9999", txtC�digoBanco
End Sub
Private Sub txtC�digoBanco_KeyPress(KeyAscii As Integer)
    If Not lInserir Then
        lAlterar = True
        StatusBarAviso = "Altera��o"
        BarraDeStatus StatusBarAviso
    End If
End Sub
Private Sub txtC�digoBanco_LostFocus()
    If mdiGeal.ActiveForm.Name = "frmAg�ncia" Then
        If txtC�digoBanco.Enabled Then
            LeftBlank txtC�digoBanco
            TBLBanco.Seek "=", txtC�digoBanco
            If TBLBanco.NoMatch Then
                MsgBox "N�o encontrei o banco !" + txtC�digoBanco, vbExclamation, "Aviso"
                txtC�digoBanco = Empty
                txtC�digoBanco.SetFocus
                Exit Sub
            End If
            txtDescri��oBanco = TBLBanco("DESCRI��O")
        End If
    End If
End Sub
Private Sub txtContato_KeyPress(KeyAscii As Integer)
    If Not lInserir Then
        lAlterar = True
        StatusBarAviso = "Altera��o"
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
        StatusBarAviso = "Altera��o"
        BarraDeStatus StatusBarAviso
    End If
End Sub
Private Sub txtDescri��oAg�ncia_Change()
    If Not lPula Then
        FormatMask "@!S30", txtDescri��oAg�ncia
    End If
End Sub
Private Sub txtDescri��oAg�ncia_KeyPress(KeyAscii As Integer)
    If Not lInserir Then
        lAlterar = True
        StatusBarAviso = "Altera��o"
        BarraDeStatus StatusBarAviso
    End If
End Sub
Private Sub txtEndere�o_KeyPress(KeyAscii As Integer)
    If Not lInserir Then
        lAlterar = True
        StatusBarAviso = "Altera��o"
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
        StatusBarAviso = "Altera��o"
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
        StatusBarAviso = "Altera��o"
        BarraDeStatus StatusBarAviso
    End If
End Sub
Private Sub txtObserva��es_KeyPress(KeyAscii As Integer)
    If Not lInserir Then
        lAlterar = True
        StatusBarAviso = "Altera��o"
        BarraDeStatus StatusBarAviso
    End If
End Sub
