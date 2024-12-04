VERSION 5.00
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#2.5#0"; "SS32X25.OCX"
Begin VB.Form frmMovimentoDeContaCorrente 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Movimento de Conta Corrente"
   ClientHeight    =   5895
   ClientLeft      =   1335
   ClientTop       =   2040
   ClientWidth     =   10680
   Icon            =   "MovimentoDeContaCorrente.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5895
   ScaleWidth      =   10680
   Begin VB.CommandButton cmdGravar 
      Caption         =   "&Gravar"
      Height          =   345
      Left            =   8070
      TabIndex        =   3
      Top             =   5550
      Width           =   1245
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   345
      Left            =   9390
      TabIndex        =   4
      Top             =   5550
      Width           =   1245
   End
   Begin VB.Frame frMovimento 
      Caption         =   " Movimento "
      Height          =   4155
      Left            =   0
      TabIndex        =   10
      Top             =   1350
      Width           =   8055
      Begin FPSpread.vaSpread dbgrdMovimento 
         Height          =   3855
         Left            =   60
         TabIndex        =   12
         Top             =   210
         Width           =   7935
         _Version        =   131077
         _ExtentX        =   13996
         _ExtentY        =   6800
         _StockProps     =   64
         AllowUserFormulas=   -1  'True
         DisplayRowHeaders=   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxCols         =   5
         MaxRows         =   15
         OperationMode   =   1
         ScrollBars      =   2
         SpreadDesigner  =   "MovimentoDeContaCorrente.frx":030A
         UserResize      =   1
      End
   End
   Begin VB.Frame frSaldo 
      Caption         =   " Saldo "
      Height          =   4155
      Left            =   8070
      TabIndex        =   9
      Top             =   1350
      Width           =   2595
      Begin FPSpread.vaSpread dbgrdSaldo 
         Height          =   3855
         Left            =   60
         TabIndex        =   11
         Top             =   210
         Width           =   2445
         _Version        =   131077
         _ExtentX        =   4313
         _ExtentY        =   6800
         _StockProps     =   64
         AllowUserFormulas=   -1  'True
         DisplayRowHeaders=   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxCols         =   2
         MaxRows         =   15
         OperationMode   =   1
         ScrollBars      =   2
         SpreadDesigner  =   "MovimentoDeContaCorrente.frx":0453
         UserResize      =   1
      End
   End
   Begin VB.Frame frContaCorrente 
      Height          =   1350
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   10665
      Begin VB.ComboBox cmbContaCorrente 
         Height          =   315
         Left            =   990
         Locked          =   -1  'True
         TabIndex        =   2
         Top             =   900
         Width           =   2085
      End
      Begin VB.ComboBox cmbAgência 
         Height          =   315
         Left            =   990
         Locked          =   -1  'True
         TabIndex        =   1
         Top             =   570
         Width           =   5625
      End
      Begin VB.ComboBox cmbBanco 
         Height          =   315
         Left            =   990
         Locked          =   -1  'True
         TabIndex        =   0
         Top             =   240
         Width           =   5625
      End
      Begin VB.Label lblConta 
         Caption         =   "Conta"
         Height          =   195
         Left            =   270
         TabIndex        =   8
         Top             =   945
         Width           =   420
      End
      Begin VB.Label lblAgência 
         Caption         =   "Agência"
         Height          =   210
         Left            =   270
         TabIndex        =   7
         Top             =   615
         Width           =   615
      End
      Begin VB.Label lblBanco 
         Caption         =   "Banco"
         Height          =   225
         Left            =   270
         TabIndex        =   6
         Top             =   255
         Width           =   660
      End
   End
End
Attribute VB_Name = "frmMovimentoDeContaCorrente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim mFechar As Boolean

Dim lInserir As Boolean
Dim lAlterar As Boolean

Dim lAllowInsert  As Boolean
Dim lAllowEdit    As Boolean
Dim lAllowDelete  As Boolean
Dim lAllowConsult As Boolean

Dim TBLMovimentoContaCorrente As Table
Dim MovimentoContaCorrenteAberto As Boolean
Dim IndiceMovimentoContaCorrenteAtivo As String

Dim TBLBanco As Table
Dim BancoAberto As Boolean
Dim IndiceBancoAtivo As String

Dim TBLAgência    As Table
Dim AgênciaAberto As Boolean
Dim IndiceAgênciaAtivo As String

Dim TBLContaCorrente As Table
Dim ContaCorrenteAberto As Boolean
Dim IndiceContaCorrenteAtivo As String

Dim TBLFormaDePagamentoVenda As Table
Dim FormaDePagamentoVendaAberto As Boolean
Dim IndiceFormaDePagamentoVendaAtivo As String

Dim TBLFormaDePagamentoCompra As Table
Dim FormaDePagamentoCompraAberto As Boolean
Dim IndiceFormaDePagamentoCompraAtivo As String

Dim TBLParâmetros As Table
Dim ParâmetrosAberto As Boolean

Dim mData As String
Dim mHistórico As String
Dim mDébitoCrédito As String
Dim mValor As Currency
Dim mSaldo As Currency
Dim mTotalDeRegistros As Long

Dim StatusBarAviso As String

Public lAtualizar As Boolean
Public Sub Excluir()

End Sub
Private Sub FillAgência(ByVal CódigoBanco As Long)
    TBLAgência.Index = "AGÊNCIA2"
    
    TBLAgência.Seek "=", CódigoBanco
    
    If TBLAgência.NoMatch Then
        TBLAgência.Index = IndiceAgênciaAtivo
        Exit Sub
    End If
    
    Do While TBLAgência("CÓDIGO DO BANCO") = CódigoBanco
        cmbAgência.AddItem TBLAgência("CÓDIGO") & " - " & TBLAgência("DESCRIÇÃO")
        TBLAgência.MoveNext
        If TBLAgência.EOF Then
            Exit Do
        End If
    Loop
    
    TBLAgência.Index = IndiceAgênciaAtivo
End Sub
Private Sub FillBanco()
    TBLBanco.MoveFirst
    
    Do While Not TBLBanco.EOF
        cmbBanco.AddItem TBLBanco("CÓDIGO") & " - " & TBLBanco("DESCRIÇÃO")
        TBLBanco.MoveNext
    Loop
End Sub
Private Sub FillContaCorrente(ByVal CódigoBanco As Long, ByVal CódigoAgência As String)
    TBLContaCorrente.Index = "CONTACORRENTE2"
    
    TBLContaCorrente.Seek "=", CódigoBanco, CódigoAgência
    
    If TBLContaCorrente.NoMatch Then
        TBLContaCorrente.Index = IndiceContaCorrenteAtivo
        Exit Sub
    End If
    
    Do While TBLContaCorrente("CÓDIGO DO BANCO") = CódigoBanco And TBLContaCorrente("CÓDIGO DA AGÊNCIA") = CódigoAgência
        cmbContaCorrente.AddItem TBLContaCorrente("CÓDIGO")
        TBLContaCorrente.MoveNext
        
        If TBLContaCorrente.EOF Then
            Exit Do
        End If
    Loop
    
    TBLContaCorrente.Index = IndiceContaCorrenteAtivo
End Sub
Public Sub Gravar()

End Sub
Public Sub Incluir()
    lInserir = True
    
    frmEditaContaCorrente.Show
    
    If frmEditaContaCorrente.lCancelado Then
        Exit Sub
    End If
    
    mData = frmEditaContaCorrente.Data
    mHistórico = frmEditaContaCorrente.Histórico
    mDébitoCrédito = frmEditaContaCorrente.DébitoCrédito
    mValor = frmEditaContaCorrente.Valor
    
    dbgrdMovimento.Col = 5
    dbgrdMovimento.Row = mTotalDeRegistros
    
    mSaldo = IIf(mDébitoCrédito = "C", mSaldo + ValStr(dbgrdMovimento.Text), mSaldo - ValStr(dbgrdMovimento.Text))
    
    If SetRecords() Then
        mTotalDeRegistros = mTotalDeRegistros + 1
        If dbgrdMovimento.MaxRows < mTotalDeRegistros Then
            dbgrdMovimento.MaxRows = mTotalDeRegistros
        End If
        
        dbgrdMovimento.Row = mTotalDeRegistros
        
        dbgrdMovimento.Col = 1
        dbgrdMovimento.Text = mData
        
        dbgrdMovimento.Col = 2
        dbgrdMovimento.Text = mHistórico
        
        dbgrdMovimento.Col = 3
        dbgrdMovimento.Text = mDébitoCrédito
        
        dbgrdMovimento.Col = 4
        dbgrdMovimento.Text = mValor
        
        dbgrdMovimento.Col = 5
        dbgrdMovimento.Text = mSaldo
    End If
End Sub
Public Sub MoveFirst()

End Sub
Public Sub MoveLast()

End Sub
Public Sub MoveNext()

End Sub
Public Sub MovePrevious()

End Sub
Sub PosRecords()

End Sub
Private Sub GetRecords()
    On Error GoTo Erro
    
    Exit Sub
    
Erro:
    
End Sub
Private Function SetRecords() As Boolean
    On Error GoTo Erro
    
    Dim CódigoDoBanco As Long
    Dim CódigoDaAgência As String
    Dim CódigoDaContaCorrente As String
    Dim ID As Long
    
    CódigoDoBanco = Trim(GetWordSeparatedBy(cmbBanco.Text, 1, "-"))
    CódigoDaAgência = Trim(GetWordSeparatedBy(cmbAgência.Text, 1, "-"))
    CódigoDaContaCorrente = Trim(GetWordSeparatedBy(cmbContaCorrente.Text, 1, "-"))
    
    'Pega o novo código interno do produto e atualiza na Tabela Parâmetros
    TBLParâmetros.Edit
    ID = TBLParâmetros("MOVIMENTO DE CONTA CORRENTE") + 1
    TBLParâmetros("MOVIMENTO DE CONTA CORRENTE") = ID
    TBLParâmetros.Update
    
    WS.BeginTrans
    
    TBLMovimentoContaCorrente.AddNew
    TBLMovimentoContaCorrente("ID") = ID
    TBLMovimentoContaCorrente("CÓDIGO DO BANCO") = CódigoDoBanco
    TBLMovimentoContaCorrente("CÓDIGO DA AGÊNCIA") = CódigoDaAgência
    TBLMovimentoContaCorrente("CÓDIGO DA CONTA CORRENTE") = CódigoDaContaCorrente
    TBLMovimentoContaCorrente("HISTÓRICO") = mHistórico
    TBLMovimentoContaCorrente("VALOR") = mValor
    TBLMovimentoContaCorrente("DÉBITO/CRÉDITO") = mDébitoCrédito
    TBLMovimentoContaCorrente("DATA DA OPERAÇÃO") = mData
    TBLMovimentoContaCorrente("SALDO") = mSaldo
    
    TBLMovimentoContaCorrente("USERNAME - CRIA") = gUsuário
    TBLMovimentoContaCorrente("HORA - CRIA") = Time
    TBLMovimentoContaCorrente("DATA - CRIA") = Date
    
    WS.CommitTrans
    
    SetRecords = True
    
    Exit Function
    
Erro:
    GeraMensagemDeErro "Movimento de Conta Corrente", True
    SetRecords = False
End Function
Private Sub ZeraCampos()
    Dim Cont As Byte
    Dim Cont1 As Byte
    
    cmbBanco.Clear
    cmbBanco.Text = Empty
    
    cmbAgência.Clear
    cmbAgência.Text = Empty
    
    cmbContaCorrente.Clear
    cmbContaCorrente.Text = Empty
    
    dbgrdMovimento.MaxRows = 15
    dbgrdSaldo.MaxRows = 15
    
    For Cont = 1 To 15
        dbgrdMovimento.Row = Cont
        dbgrdSaldo.Row = Cont
        For Cont1 = 1 To 5
            dbgrdMovimento.Col = Cont1
            dbgrdMovimento.Text = Empty
        Next
        For Cont1 = 1 To 2
            dbgrdSaldo.Col = Cont1
            dbgrdSaldo.Text = Empty
        Next
    Next
End Sub
Private Sub cmbAgência_Click()
    Dim CódigoBanco As Long
    Dim CódigoAgência As String
    
    cmbAgência.Locked = True
    
    cmbContaCorrente.Clear
    cmbContaCorrente.Text = Empty
    
    CódigoBanco = Trim(GetWordSeparatedBy(cmbBanco.Text, 1, "-"))
    CódigoAgência = Trim(GetWordSeparatedBy(cmbAgência.Text, 1, "-"))
    
    FillContaCorrente CódigoBanco, CódigoAgência
End Sub
Private Sub cmbAgência_DropDown()
    cmbAgência.Locked = False
End Sub
Private Sub cmbBanco_Click()
    Dim CódigoBanco As Long
    cmbBanco.Locked = True
    cmbAgência.Clear
    cmbAgência.Text = Empty
    cmbContaCorrente.Clear
    cmbContaCorrente.Text = Empty
    
    CódigoBanco = Trim(GetWordSeparatedBy(cmbBanco.Text, 1, "-"))
    
    FillAgência CódigoBanco
End Sub
Private Sub cmbBanco_DropDown()
    cmbBanco.Locked = False
End Sub
Private Sub cmbContaCorrente_Click()
    cmbContaCorrente.Locked = True
End Sub
Private Sub cmbContaCorrente_DropDown()
    cmbContaCorrente.Locked = False
End Sub
Private Sub Form_Activate()
    If mFechar Then
        Unload Me
        Exit Sub
    End If
    
    If Not MovimentoContaCorrenteAberto Then
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
    
    If Not ContaCorrenteAberto Then
        Unload Me
        Exit Sub
    End If
    
    If TBLBanco.RecordCount = 0 Then
        MsgBox "Nenhum BANCO foi cadastrado!", vbInformation, "Aviso"
        Unload Me
        Exit Sub
    End If
    
    If TBLAgência.RecordCount = 0 Then
        MsgBox "Nenhuma AGÊNCIA foi cadastrada!", vbInformation, "Aviso"
        Unload Me
        Exit Sub
    End If
    
    If TBLContaCorrente.RecordCount = 0 Then
        MsgBox "Nenhuma CONTA CORRENTE foi cadastrada!", vbInformation, "Aviso"
        Unload Me
        Exit Sub
    End If
    
    BotãoIncluir lAllowInsert
    BotãoExcluir False
    BotãoGravar False
    BotãoImprimir False
    
    NavegaçãoInferior False
    NavegaçãoSuperior False
    
    StatusBarAviso = "Pronto"

    If lAtualizar Then
        BotãoAtualizar True
    Else
        BotãoAtualizar False
    End If
    
    BarraDeStatus StatusBarAviso
End Sub
Private Sub Form_Load()
    On Error GoTo Erro
    
    lAllowInsert = Allow("CONTA CORRENTE (MOVIMENTO)", "I")
    'lAllowEdit = Allow("CONTA CORRENTE (MOVIMENTO)", "A")
    'lAllowDelete = Allow("CONTA CORRENTE (MOVIMENTO)", "E")
    lAllowConsult = Allow("CONTA CORRENTE (MOVIMENTO)", "C")
    
    ZeraCampos
    
    'Movimento de Conta Corrente
    MovimentoContaCorrenteAberto = AbreTabela(Dicionário, "FINANCEIRO", "MOVIMENTO - CONTA CORRENTE", DBFinanceiro, TBLMovimentoContaCorrente, TBLTabela, dbOpenTable)
    
    If MovimentoContaCorrenteAberto Then
        IndiceMovimentoContaCorrenteAtivo = "MOVIMENTOCONTACORRENTE1"
        TBLMovimentoContaCorrente.Index = IndiceMovimentoContaCorrenteAtivo
    Else
        MsgBox "Não consegui abrir a tabela 'Movimento - Conta Corrente' !", vbCritical, "Erro"
        Exit Sub
    End If
    
    'Banco
    BancoAberto = AbreTabela(Dicionário, "FINANCEIRO", "BANCO", DBFinanceiro, TBLBanco, TBLTabela, dbOpenTable)
    
    If BancoAberto Then
        IndiceBancoAtivo = "BANCO1"
        TBLBanco.Index = IndiceBancoAtivo
    Else
        MsgBox "Não consegui abrir a tabela 'Banco' !", vbCritical, "Erro"
        Exit Sub
    End If
    
    'Agência
    AgênciaAberto = AbreTabela(Dicionário, "FINANCEIRO", "AGÊNCIA", DBFinanceiro, TBLAgência, TBLTabela, dbOpenTable)
    
    If AgênciaAberto Then
        IndiceAgênciaAtivo = "AGÊNCIA1"
        TBLAgência.Index = IndiceAgênciaAtivo
    Else
        MsgBox "Não consegui abrir a tabela 'Agência' !", vbCritical, "Erro"
        Exit Sub
    End If
    
    'Conta Corrente
    ContaCorrenteAberto = AbreTabela(Dicionário, "FINANCEIRO", "CONTA CORRENTE", DBFinanceiro, TBLContaCorrente, TBLTabela, dbOpenTable)
    
    If ContaCorrenteAberto Then
        IndiceContaCorrenteAtivo = "CONTACORRENTE1"
        TBLContaCorrente.Index = IndiceContaCorrenteAtivo
    Else
        MsgBox "Não consegui abrir a tabela 'Conta Corrente' !", vbCritical, "Erro"
        Exit Sub
    End If
    
    ''Movimento
    dbgrdMovimento.MAXCOLS = 5
    dbgrdMovimento.MaxRows = 15
    dbgrdMovimento.Row = 0
    
    'Data
    dbgrdMovimento.Col = 1
    dbgrdMovimento.Text = "Data"
    dbgrdMovimento.FontBold = True
    dbgrdMovimento.ColWidth(dbgrdMovimento.Col) = 7.25
    
    'Movimento
    dbgrdMovimento.Col = 2
    dbgrdMovimento.Text = "Histórico"
    dbgrdMovimento.FontBold = True
    dbgrdMovimento.ColWidth(dbgrdMovimento.Col) = 30.75
    
    'Débito/Crédito
    dbgrdMovimento.Col = 3
    dbgrdMovimento.Text = "D/C"
    dbgrdMovimento.FontBold = True
    dbgrdMovimento.ColWidth(dbgrdMovimento.Col) = 3.875
    
    'Valor
    dbgrdMovimento.Col = 4
    dbgrdMovimento.Text = "Valor"
    dbgrdMovimento.FontBold = True
    dbgrdMovimento.ColWidth(dbgrdMovimento.Col) = 10.75
    
    'Saldo
    dbgrdMovimento.Col = 5
    dbgrdMovimento.Text = "Saldo"
    dbgrdMovimento.FontBold = True
    dbgrdMovimento.ColWidth(dbgrdMovimento.Col) = 10.75
    
    ''Saldo
    dbgrdSaldo.MAXCOLS = 2
    dbgrdSaldo.MaxRows = 15
    dbgrdSaldo.Row = 0
    
    'Data
    dbgrdSaldo.Col = 1
    dbgrdSaldo.Text = "Data"
    dbgrdMovimento.FontBold = True
    dbgrdSaldo.ColWidth(dbgrdSaldo.Col) = 7.25
    
    'Valor
    dbgrdSaldo.Col = 2
    dbgrdSaldo.Text = "Valor"
    dbgrdMovimento.FontBold = True
    dbgrdSaldo.ColWidth(dbgrdSaldo.Col) = 10.75
        
    FillBanco
    
    Exit Sub
    
Erro:
    GeraMensagemDeErro "Moviento de Conta Corrente - Load"
    mFechar = True
End Sub
Private Sub Form_Unload(Cancel As Integer)
    If MovimentoContaCorrenteAberto Then
        TBLMovimentoContaCorrente.Close
    End If
    
    If BancoAberto Then
        TBLBanco.Close
    End If
    
    If AgênciaAberto Then
        TBLAgência.Close
    End If
    
    If ContaCorrenteAberto Then
        TBLContaCorrente.Close
    End If
    
    Set frmEditaContaCorrente = Nothing
End Sub
