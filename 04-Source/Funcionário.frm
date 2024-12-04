VERSION 5.00
Begin VB.Form frmFuncion�rio 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Funcio�rio"
   ClientHeight    =   5100
   ClientLeft      =   1575
   ClientTop       =   1515
   ClientWidth     =   6540
   Icon            =   "Funcion�rio.frx":0000
   LinkTopic       =   "frmFuncion�rio"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5100
   ScaleWidth      =   6540
   Begin VB.Frame frDadosContratuais 
      Caption         =   "Dados Contratuais "
      Height          =   1155
      Left            =   0
      TabIndex        =   20
      Top             =   3540
      Width           =   6525
      Begin VB.TextBox txtSal�rio 
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
         Left            =   4770
         TabIndex        =   9
         Text            =   " "
         Top             =   690
         Width           =   1665
      End
      Begin VB.TextBox txtDatadeSa�da 
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
         Left            =   1410
         TabIndex        =   8
         Top             =   690
         Width           =   1305
      End
      Begin VB.TextBox txtDatadeEntrada 
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
         Left            =   1410
         TabIndex        =   7
         Top             =   270
         Width           =   1305
      End
      Begin VB.Label lblSal�rio 
         Caption         =   "Sal�rio"
         Height          =   195
         Left            =   4140
         TabIndex        =   23
         Top             =   750
         Width           =   555
      End
      Begin VB.Label lblDatadeSa�da 
         Caption         =   "Data de Sa�da"
         Height          =   225
         Left            =   150
         TabIndex        =   22
         Top             =   750
         Width           =   1185
      End
      Begin VB.Label lblDataDeEntrada 
         Caption         =   "Data de Entrada"
         Height          =   225
         Left            =   150
         TabIndex        =   21
         Top             =   300
         Width           =   1245
      End
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   345
      Left            =   5280
      TabIndex        =   11
      Top             =   4740
      Width           =   1245
   End
   Begin VB.CommandButton cmdGravar 
      Caption         =   "&Gravar"
      Height          =   345
      Left            =   3960
      TabIndex        =   10
      Top             =   4740
      Width           =   1245
   End
   Begin VB.Frame frDadosCadastrais 
      Caption         =   " Dados Cadastrais "
      Height          =   3525
      Left            =   0
      TabIndex        =   17
      Top             =   0
      Width           =   6525
      Begin VB.TextBox txtCpf 
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
         Text            =   "   .   .   -  "
         Top             =   3030
         Width           =   2310
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
         Left            =   5130
         TabIndex        =   5
         Text            =   "  .   -   "
         Top             =   2580
         Width           =   1305
      End
      Begin VB.TextBox txtUF 
         Height          =   285
         Left            =   1200
         TabIndex        =   4
         Top             =   2610
         Width           =   435
      End
      Begin VB.TextBox txtCidade 
         Height          =   285
         Left            =   1200
         TabIndex        =   3
         Top             =   2130
         Width           =   5235
      End
      Begin VB.TextBox txtBairro 
         Height          =   285
         Left            =   1200
         TabIndex        =   2
         Top             =   1680
         Width           =   5235
      End
      Begin VB.TextBox txtEndere�o 
         Height          =   285
         Left            =   1200
         TabIndex        =   1
         Top             =   1230
         Width           =   5235
      End
      Begin VB.TextBox txtNome 
         Height          =   285
         Left            =   1200
         TabIndex        =   0
         Top             =   780
         Width           =   5235
      End
      Begin VB.TextBox txtC�digo 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1200
         TabIndex        =   24
         Top             =   330
         Width           =   525
      End
      Begin VB.Label lblCgcCpf 
         Caption         =   "C. P. F."
         Height          =   195
         Left            =   150
         TabIndex        =   12
         Top             =   3090
         Width           =   645
      End
      Begin VB.Label lblCep 
         Caption         =   "CEP"
         Height          =   195
         Left            =   4680
         TabIndex        =   13
         Top             =   2640
         Width           =   315
      End
      Begin VB.Label lblUF 
         Caption         =   "U. F."
         Height          =   225
         Left            =   150
         TabIndex        =   14
         Top             =   2640
         Width           =   405
      End
      Begin VB.Label lblCidade 
         Caption         =   "Cidade"
         Height          =   225
         Left            =   150
         TabIndex        =   15
         Top             =   2160
         Width           =   945
      End
      Begin VB.Label lblBairro 
         Caption         =   "Bairro"
         Height          =   225
         Left            =   150
         TabIndex        =   16
         Top             =   1710
         Width           =   945
      End
      Begin VB.Label lblEndere�o 
         Caption         =   "Endere�o"
         Height          =   195
         Left            =   150
         TabIndex        =   19
         Top             =   1260
         Width           =   975
      End
      Begin VB.Label lblNomeRaz�oSocial 
         Caption         =   "Nome"
         Height          =   195
         Left            =   150
         TabIndex        =   18
         Top             =   810
         Width           =   1065
      End
      Begin VB.Label lblC�digo 
         Caption         =   "C�digo"
         Height          =   195
         Left            =   150
         TabIndex        =   25
         Top             =   360
         Width           =   1065
      End
   End
End
Attribute VB_Name = "frmFuncion�rio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim TBLFuncion�rio As Table
Dim Funcion�rioAberto As Boolean
Dim IndiceFuncion�rioAtivo$

Dim TBLPar�metros As Table
Dim Par�metrosAberto As Boolean

Dim lAllowInsert  As Boolean
Dim lAllowEdit    As Boolean
Dim lAllowDelete  As Boolean
Dim lAllowConsult As Boolean

Dim lInserir As Boolean
Dim lAlterar As Boolean

Dim mFechar As Boolean
Dim lPula As Boolean
Dim lPush As Boolean

Dim lInicio As Boolean
Dim StatusBar$

Public StatusBarAviso$

Dim DataBaseName(1 To 1) As String
Public Relat�rio$
Public TotalDatabaseName%

Public lAtualizar As Boolean
Private Sub AtivaCampos()
    Bot�oImprimir True
    frDadosCadastrais.Enabled = True
    frDadosContratuais.Enabled = True
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
    
    If TBLFuncion�rio.RecordCount = 0 Then
        Navega��oInferior False
        Navega��oSuperior False
        Bot�oGravar False
        cmdGravar.Enabled = False
        cmdCancelar.Enabled = False
        DesativaCampos
        lPush = True
        ZeraCampos
        lPush = False
        Cancelamento = True
        Exit Function
    End If
    
    Cancelamento = True
    
    TestaInferior TBLFuncion�rio, lAllowEdit, lAllowDelete, lAllowConsult
    TestaSuperior TBLFuncion�rio, lAllowEdit, lAllowDelete, lAllowConsult
    
    GetRecords
End Function
Public Function PushDataBaseName(ByVal Posi��o As Integer) As String
    PushDataBaseName = DataBaseName(Posi��o)
End Function
Private Sub DesativaCampos()
    Bot�oImprimir False
    frDadosCadastrais.Enabled = False
    frDadosContratuais.Enabled = False
    cmdGravar.Enabled = False
    cmdCancelar.Enabled = False
    Bot�oGravar False
End Sub
Public Sub Encontrar()
    If Not lAllowConsult Then
        Exit Sub
    End If
    Set frmEncontrar.DBBancoDeDados = DBUsu�rio
    frmEncontrar.NomeDaJanela = "Funcion�rio"
    frmEncontrar.LabelDescription = "Nome"
    frmEncontrar.Mensagem = "Nenhuma funcion�rio foi selecionado!"
    frmEncontrar.BancoDeDados = "USU�RIO"
    frmEncontrar.Tabela = "FUNCION�RIO"
    frmEncontrar.Indice = "1"
    frmEncontrar.CampoChave = "C�DIGO"
    frmEncontrar.CampoPreencheLista = "NOME"
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
    
    TBLFuncion�rio.Delete
    
    If Err <> 0 Then
        GeraMensagemDeErro "Funcion�rio - Excluir - " & txtNome, True
        StatusBarAviso = "Falha na exclus�o"
        BarraDeStatus StatusBarAviso
        Exit Sub
    End If
    
    WS.CommitTrans
    
    Log gUsu�rio, "Exclus�o - Funcion�rio: " & txtC�digo & " - " & txtNome
    
    StatusBarAviso = "Exclus�o bem sucedida"
    BarraDeStatus StatusBarAviso
    
    If TBLFuncion�rio.RecordCount = 0 Then
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
    
    If TBLFuncion�rio.BOF Then
        TBLFuncion�rio.MoveFirst
    ElseIf TBLFuncion�rio.EOF Then
        TBLFuncion�rio.MoveLast
    Else
        TBLFuncion�rio.MovePrevious
        If TBLFuncion�rio.BOF Then
            TBLFuncion�rio.MoveNext
        End If
    End If
    
    GetRecords
    
    TestaInferior TBLFuncion�rio, lAllowEdit, lAllowDelete, lAllowConsult
    TestaSuperior TBLFuncion�rio, lAllowEdit, lAllowDelete, lAllowConsult
End Sub
Public Sub Gravar()
    If lInserir Then
        'Pega o novo c�digo interno de funcion�rio e atualiza na Tabela Par�metros
        txtC�digo = TBLPar�metros("FUNCION�RIO") + 1
        TBLPar�metros.Edit
        TBLPar�metros("FUNCION�RIO") = txtC�digo
        TBLPar�metros.Update
        
        If SetRecords Then
            PosRecords
            lInserir = False
            StatusBarAviso = "Inclus�o bem sucedida"
        Else
            StatusBarAviso = "Falha na inclus�o"
            Exit Sub
        End If
    Else
        If TBLFuncion�rio.RecordCount > 0 And Not TBLFuncion�rio.BOF And Not TBLFuncion�rio.EOF Then
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
    
    TestaInferior TBLFuncion�rio, lAllowEdit, lAllowDelete, lAllowConsult
    TestaSuperior TBLFuncion�rio, lAllowEdit, lAllowDelete, lAllowConsult
    
    If TBLFuncion�rio.RecordCount = 0 Then
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
    
    If txtNome.Enabled Then
        txtNome.SetFocus
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
    
    txtNome.SetFocus

End Sub
Public Sub MoveFirst()
    If lAlterar Then
       If Not Cancelamento Then
            Exit Sub
        End If
    End If
    
    StatusBarAviso = "Pronto"
    BarraDeStatus StatusBarAviso
    
    TBLFuncion�rio.MoveFirst
    
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
    
    TBLFuncion�rio.MoveLast
    
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
    
    TBLFuncion�rio.MoveNext
    If TBLFuncion�rio.EOF Then
        TBLFuncion�rio.MovePrevious
        Exit Sub
    End If
    
    StatusBarAviso = "Pronto"
    BarraDeStatus StatusBarAviso
    
    Navega��oInferior lAllowConsult
    TestaSuperior TBLFuncion�rio, lAllowEdit, lAllowDelete, lAllowConsult
    
    GetRecords
End Sub
Public Sub MovePrevious()
    If lAlterar Then
       If Not Cancelamento Then
            Exit Sub
        End If
    End If
    
    TBLFuncion�rio.MovePrevious
    If TBLFuncion�rio.BOF Then
        TBLFuncion�rio.MoveNext
        Exit Sub
    End If
    
    StatusBarAviso = "Pronto"
    BarraDeStatus StatusBarAviso
    
    Navega��oSuperior lAllowConsult
    TestaInferior TBLFuncion�rio, lAllowEdit, lAllowDelete, lAllowConsult
    
    GetRecords
End Sub
Public Sub PosRecords()
    If TBLFuncion�rio.RecordCount = 0 Then
        Exit Sub
    End If
    
    TBLFuncion�rio.Seek "=", txtC�digo
    If TBLFuncion�rio.NoMatch Then
        MsgBox "N�o consegui encontrar o funcion�rio com o c�digo  " + txtC�digo, vbExclamation, "Erro"
        TBLFuncion�rio.MoveFirst
        Navega��oInferior False
        Navega��oInferior lAllowConsult
    Else
        TestaInferior TBLFuncion�rio, lAllowEdit, lAllowDelete, lAllowConsult
        TestaSuperior TBLFuncion�rio, lAllowEdit, lAllowDelete, lAllowConsult
    End If
    GetRecords
End Sub
Private Sub GetRecords()
    On Error GoTo Erro
    
    lPush = True
    lPula = True
    If Not lAllowConsult Then
        ZeraCampos
        DesativaCampos
        lPush = False
        lPula = False
        Exit Sub
    End If
    txtC�digo = TBLFuncion�rio("C�DIGO")
    txtNome = TBLFuncion�rio("NOME")
    txtEndere�o = TBLFuncion�rio("ENDERE�O")
    txtBairro = TBLFuncion�rio("BAIRRO")
    txtCidade = TBLFuncion�rio("CIDADE")
    txtUF = TBLFuncion�rio("UF")
    txtCep = TBLFuncion�rio("CEP")
    txtCpf = TBLFuncion�rio("CPF")
    
    If TBLFuncion�rio("DATA DE ENTRADA") <> vbNull Then
        txtDatadeEntrada = FormatStringMask(CheckDataMask, TBLFuncion�rio("DATA DE ENTRADA"))
        CorrigeData DataMask, txtDatadeEntrada, TBLFuncion�rio("DATA DE ENTRADA")
    Else
        txtDatadeEntrada = DataNula
    End If
    
    If TBLFuncion�rio("DATA DE SA�DA") <> vbNull Then
        txtDatadeSa�da = FormatStringMask(CheckDataMask, TBLFuncion�rio("DATA DE SA�DA"))
        CorrigeData DataMask, txtDatadeSa�da, TBLFuncion�rio("DATA DE SA�DA")
    Else
        txtDatadeSa�da = DataNula
    End If
    
    txtSal�rio = TBLFuncion�rio("SAL�RIO")
    txtSal�rio_LostFocus
    lPush = False
    lPula = False
    If Not lAllowEdit Then
        DesativaCampos
    End If
    If Not lAllowEdit Then
        DesativaCampos
    End If
    
    Exit Sub
    
Erro:
    GeraMensagemDeErro "Funcion�rio - GetRecords "
    Resume Next
End Sub
Private Function SetRecords()
    On Error GoTo Erro
    
    Dim Msg$
    Dim Confirma��o As Integer, Msg1$, Msg2$
    
    WS.BeginTrans 'Inicia uma Transa��o
    
    If lInserir Then
        TBLFuncion�rio.AddNew
    Else
        TBLFuncion�rio.Edit
    End If
    
    If lInserir Then
        TBLFuncion�rio("C�DIGO") = txtC�digo
    End If
    
    TBLFuncion�rio("NOME") = txtNome
    TBLFuncion�rio("ENDERE�O") = txtEndere�o
    TBLFuncion�rio("BAIRRO") = txtBairro
    TBLFuncion�rio("CIDADE") = txtCidade
    TBLFuncion�rio("UF") = txtUF
    TBLFuncion�rio("CEP") = txtCep
    TBLFuncion�rio("CPF") = txtCpf
    TBLFuncion�rio("DATA DE ENTRADA") = IIf(Trim(StrTran(txtDatadeEntrada, "/")) <> Empty, txtDatadeEntrada, vbNull)
    TBLFuncion�rio("DATA DE SA�DA") = IIf(Trim(StrTran(txtDatadeSa�da, "/")) <> Empty, txtDatadeSa�da, vbNull)
    TBLFuncion�rio("SAL�RIO") = txtSal�rio
    
    If lInserir Then
        TBLFuncion�rio("USERNAME - CRIA") = gUsu�rio
        TBLFuncion�rio("DATA - CRIA") = Date
        TBLFuncion�rio("HORA - CRIA") = Time
        TBLFuncion�rio("USERNAME - ALTERA") = "VAZIO"
        TBLFuncion�rio("DATA - ALTERA") = vbNull
        TBLFuncion�rio("HORA - ALTERA") = vbNull
    End If
    If lAlterar Then
        TBLFuncion�rio("USERNAME - ALTERA") = gUsu�rio
        TBLFuncion�rio("DATA - ALTERA") = Date
        TBLFuncion�rio("HORA - ALTERA") = Time
    End If
    TBLFuncion�rio.Update
    
Erro:
    If Err <> 0 Then
        TBLFuncion�rio.CancelUpdate
        GeraMensagemDeErro "Funcion�rio - SetRecords - " & txtNome, True
        SetRecords = False
        Exit Function
    End If

    WS.CommitTrans 'Grava as altera��es ou inclus�es se n�o houverem erros
    
    If lInserir Then
        Log gUsu�rio, "Inclus�o - Funcion�rio " & txtC�digo & " - " & txtNome
    Else
        Log gUsu�rio, "Altera��o - Funcion�rio " & txtC�digo & " - " & txtNome
    End If
    
    SetRecords = True
End Function
Private Sub ZeraCampos()
    lPula = True
    txtC�digo = Empty
    txtNome = Empty
    txtEndere�o = Empty
    txtBairro = Empty
    txtCidade = Empty
    txtUF = Empty
    txtCep = Empty
    txtCpf = Empty
    txtDatadeEntrada = DataNula
    txtDatadeSa�da = DataNula
    txtSal�rio = FormatStringMask("@V ##.###.##0,00", "0,00")
    lPula = False
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
    If Not Funcion�rioAberto Then
        Unload Me
        Exit Sub
    End If
    
    If Not Par�metrosAberto Then
        Unload Me
        Exit Sub
    End If
    
    TestaInferior TBLFuncion�rio, lAllowEdit, lAllowDelete, lAllowConsult
    TestaSuperior TBLFuncion�rio, lAllowEdit, lAllowDelete, lAllowConsult
    
    If TBLFuncion�rio.RecordCount = 0 Then
        Bot�oGravar False
        cmdGravar.Enabled = False
        cmdCancelar.Enabled = False
        Bot�oImprimir False
    Else
        Bot�oGravar (lInserir Or lAllowEdit)
        cmdGravar.Enabled = (lInserir Or lAllowEdit)
        cmdCancelar.Enabled = (lInserir Or lAllowEdit)
        Bot�oImprimir True
        If lInicio Then
            txtNome.SetFocus
            lInicio = False
        End If
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
    mdiGeal.StatusBar.Panels("Posi��o").Visible = True
    ResizeStatusBar
End Sub
Private Sub Form_Deactivate()
    cmdGravar.Enabled = False
    cmdCancelar.Enabled = False
    Bot�oImprimir False
End Sub
Private Sub Form_Load()
    On Error GoTo Erro
    
    ZeraCampos
    
    lAllowInsert = Allow("FUNCION�RIO", "I")
    lAllowEdit = Allow("FUNCION�RIO", "A")
    lAllowDelete = Allow("FUNCION�RIO", "E")
    lAllowConsult = Allow("FUNCION�RIO", "C")
    
    lInserir = False
    lAlterar = False
    lPush = False
    lInicio = True
    
    Funcion�rioAberto = AbreTabela(Dicion�rio, "USU�RIO", "FUNCION�RIO", DBUsu�rio, TBLFuncion�rio, TBLTabela, dbOpenTable)
    
    If Funcion�rioAberto Then
        IndiceFuncion�rioAtivo = "FUNCION�RIO1"
        TBLFuncion�rio.Index = IndiceFuncion�rioAtivo
    Else
        MsgBox "N�o consegui abrir a tabela 'Funcion�rio' !", vbCritical, "Erro"
        Exit Sub
    End If
    
    Par�metrosAberto = AbreTabela(Dicion�rio, "SISTEMA", "PAR�METROS", DBSistema, TBLPar�metros, TBLTabela, dbOpenTable)
    
    If Par�metrosAberto Then
    Else
        MsgBox "N�o consegui abrir a tabela 'Par�metros' !", vbCritical, "Erro"
        Exit Sub
    End If
    
    Bot�oIncluir lAllowInsert
 
    If TBLFuncion�rio.RecordCount = 0 Then
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
        
    If TBLFuncion�rio.RecordCount = 0 Or TBLFuncion�rio.RecordCount = 1 Then
        Navega��oSuperior False
    Else
        Navega��oInferior lAllowConsult
    End If
   
    Relat�rio = AddPath(Aplica��oPath, "REPORT\FUNCION�RIO.RPT")
    TotalDatabaseName = 1
    DataBaseName(1) = AddPath(Aplica��oPath, "DATABASE\USU�RIO.MDB")
    Exit Sub
    
Erro:
    GeraMensagemDeErro "Funcion�rio - Load"
    Funcion�rioAberto = False
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
    
    Set frmFuncion�rio = Nothing
End Sub
Private Sub Form_Unload(Cancel As Integer)
    If Funcion�rioAberto Then
        TBLFuncion�rio.Close
    End If
    If Par�metrosAberto Then
        TBLPar�metros.Close
    End If
    If Forms.Count = 2 Then
        AllBot�es False
    End If
End Sub
Private Sub txtBairro_Change()
    FormatMask "@!S30", txtBairro
End Sub
Private Sub txtBairro_KeyPress(KeyAscii As Integer)
    If Not lInserir Then
        If Not lPush Then
            lAlterar = True
            StatusBarAviso = "Altera��o"
            BarraDeStatus StatusBarAviso
        End If
    End If
End Sub
Private Sub txtCep_Change()
    NumericOnly txtCep
    FormatMask "99.999-999", txtCep
End Sub
Private Sub txtCep_KeyPress(KeyAscii As Integer)
    If Not lInserir Then
        If Not lPush Then
            lAlterar = True
            StatusBarAviso = "Altera��o"
            BarraDeStatus StatusBarAviso
        End If
    End If
End Sub
Private Sub txtCpf_Change()
    NumericOnly txtCpf
    FormatMask "999.999.999-99", txtCpf
End Sub
Private Sub txtCpf_KeyPress(KeyAscii As Integer)
    If Not lInserir Then
        If Not lPush Then
            lAlterar = True
            StatusBarAviso = "Altera��o"
            BarraDeStatus StatusBarAviso
        End If
    End If
End Sub
Private Sub txtCpf_LostFocus()
    If Not IsCorrectCPF(txtCpf) Then
        MsgBox "C. P. F. incorreto !", vbCritical, "Erro"
        txtCpf.SetFocus
    End If
End Sub
Private Sub txtCidade_Change()
    FormatMask "@!S30", txtCidade
End Sub
Private Sub txtCidade_KeyPress(KeyAscii As Integer)
    If Not lInserir Then
        If Not lPush Then
            lAlterar = True
            StatusBarAviso = "Altera��o"
            BarraDeStatus StatusBarAviso
        End If
    End If
End Sub
Private Sub txtDatadeEntrada_Change()
    If Not lPula Then
        lPula = True
        FormatMask DataMask, txtDatadeEntrada
        lPula = False
    End If
End Sub
Private Sub txtDatadeEntrada_KeyPress(KeyAscii As Integer)
    If Not lInserir Then
        lAlterar = True
        StatusBarAviso = "Altera��o"
        BarraDeStatus StatusBarAviso
    End If
End Sub
Private Sub txtDatadeEntrada_LostFocus()
    If StrTran(txtDatadeEntrada.Text, "/") <> Space(8) Then
        lPula = True
        CorrigeData DataMask, txtDatadeEntrada, Date
        lPula = False
        If Not FormatMask(CheckDataMask, txtDatadeEntrada) Then
            Beep
            MsgBox "Data inv�lida !", vbCritical, "Erro"
            txtDatadeEntrada.SelStart = 0
            txtDatadeEntrada.SetFocus
        End If
    End If
End Sub
Private Sub txtDatadeSa�da_Change()
    If Not lPula Then
        lPula = True
        FormatMask DataMask, txtDatadeSa�da
        lPula = False
    End If
End Sub
Private Sub txtDatadeSa�da_KeyPress(KeyAscii As Integer)
    If Not lInserir Then
        lAlterar = True
        StatusBarAviso = "Altera��o"
        BarraDeStatus StatusBarAviso
    End If
End Sub
Private Sub txtDatadeSa�da_LostFocus()
    If StrTran(txtDatadeSa�da.Text, "/") <> Space(8) Then
        lPula = True
        CorrigeData DataMask, txtDatadeSa�da, Date
        lPula = False
        If Not FormatMask(CheckDataMask, txtDatadeSa�da) Then
            Beep
            MsgBox "Data inv�lida !", vbCritical, "Erro"
            txtDatadeSa�da.SelStart = 0
            txtDatadeSa�da.SetFocus
        End If
    End If
End Sub
Private Sub txtEndere�o_Change()
    FormatMask "@S40", txtEndere�o
End Sub
Private Sub txtEndere�o_KeyPress(KeyAscii As Integer)
    If Not lInserir Then
        If Not lPush Then
            lAlterar = True
            StatusBarAviso = "Altera��o"
            BarraDeStatus StatusBarAviso
        End If
    End If
End Sub
Private Sub txtNome_Change()
    FormatMask "@!S40", txtNome
End Sub
Private Sub txtNome_KeyPress(KeyAscii As Integer)
    If Not lInserir Then
        If Not lPush Then
            lAlterar = True
            StatusBarAviso = "Altera��o"
            BarraDeStatus StatusBarAviso
        End If
    End If
End Sub
Private Sub txtSal�rio_Change()
    If Not lPula Then
        lPula = True
        FormatMask "@K 99.999.999,99", txtSal�rio
        lPula = False
    End If
End Sub
Private Sub txtSal�rio_KeyPress(KeyAscii As Integer)
    If Not lInserir Then
        lAlterar = True
        StatusBar = "Altera��o"
        BarraDeStatus StatusBar
    End If
End Sub
Private Sub txtSal�rio_LostFocus()
    lPula = True
    FormatMask "@V ##.###.##0,00", txtSal�rio
    lPula = False
End Sub
Private Sub txtUF_Change()
    UpperOnly txtUF
    LetterOnly txtUF
End Sub
Private Sub txtUF_KeyPress(KeyAscii As Integer)
    If Not lInserir Then
        If Not lPush Then
            lAlterar = True
            StatusBarAviso = "Altera��o"
            BarraDeStatus StatusBarAviso
        End If
    End If
End Sub
