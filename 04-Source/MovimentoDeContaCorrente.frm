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
      Begin VB.ComboBox cmbAg�ncia 
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
      Begin VB.Label lblAg�ncia 
         Caption         =   "Ag�ncia"
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

Dim TBLAg�ncia    As Table
Dim Ag�nciaAberto As Boolean
Dim IndiceAg�nciaAtivo As String

Dim TBLContaCorrente As Table
Dim ContaCorrenteAberto As Boolean
Dim IndiceContaCorrenteAtivo As String

Dim TBLFormaDePagamentoVenda As Table
Dim FormaDePagamentoVendaAberto As Boolean
Dim IndiceFormaDePagamentoVendaAtivo As String

Dim TBLFormaDePagamentoCompra As Table
Dim FormaDePagamentoCompraAberto As Boolean
Dim IndiceFormaDePagamentoCompraAtivo As String

Dim TBLPar�metros As Table
Dim Par�metrosAberto As Boolean

Dim mData As String
Dim mHist�rico As String
Dim mD�bitoCr�dito As String
Dim mValor As Currency
Dim mSaldo As Currency
Dim mTotalDeRegistros As Long

Dim StatusBarAviso As String

Public lAtualizar As Boolean
Public Sub Excluir()

End Sub
Private Sub FillAg�ncia(ByVal C�digoBanco As Long)
    TBLAg�ncia.Index = "AG�NCIA2"
    
    TBLAg�ncia.Seek "=", C�digoBanco
    
    If TBLAg�ncia.NoMatch Then
        TBLAg�ncia.Index = IndiceAg�nciaAtivo
        Exit Sub
    End If
    
    Do While TBLAg�ncia("C�DIGO DO BANCO") = C�digoBanco
        cmbAg�ncia.AddItem TBLAg�ncia("C�DIGO") & " - " & TBLAg�ncia("DESCRI��O")
        TBLAg�ncia.MoveNext
        If TBLAg�ncia.EOF Then
            Exit Do
        End If
    Loop
    
    TBLAg�ncia.Index = IndiceAg�nciaAtivo
End Sub
Private Sub FillBanco()
    TBLBanco.MoveFirst
    
    Do While Not TBLBanco.EOF
        cmbBanco.AddItem TBLBanco("C�DIGO") & " - " & TBLBanco("DESCRI��O")
        TBLBanco.MoveNext
    Loop
End Sub
Private Sub FillContaCorrente(ByVal C�digoBanco As Long, ByVal C�digoAg�ncia As String)
    TBLContaCorrente.Index = "CONTACORRENTE2"
    
    TBLContaCorrente.Seek "=", C�digoBanco, C�digoAg�ncia
    
    If TBLContaCorrente.NoMatch Then
        TBLContaCorrente.Index = IndiceContaCorrenteAtivo
        Exit Sub
    End If
    
    Do While TBLContaCorrente("C�DIGO DO BANCO") = C�digoBanco And TBLContaCorrente("C�DIGO DA AG�NCIA") = C�digoAg�ncia
        cmbContaCorrente.AddItem TBLContaCorrente("C�DIGO")
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
    mHist�rico = frmEditaContaCorrente.Hist�rico
    mD�bitoCr�dito = frmEditaContaCorrente.D�bitoCr�dito
    mValor = frmEditaContaCorrente.Valor
    
    dbgrdMovimento.Col = 5
    dbgrdMovimento.Row = mTotalDeRegistros
    
    mSaldo = IIf(mD�bitoCr�dito = "C", mSaldo + ValStr(dbgrdMovimento.Text), mSaldo - ValStr(dbgrdMovimento.Text))
    
    If SetRecords() Then
        mTotalDeRegistros = mTotalDeRegistros + 1
        If dbgrdMovimento.MaxRows < mTotalDeRegistros Then
            dbgrdMovimento.MaxRows = mTotalDeRegistros
        End If
        
        dbgrdMovimento.Row = mTotalDeRegistros
        
        dbgrdMovimento.Col = 1
        dbgrdMovimento.Text = mData
        
        dbgrdMovimento.Col = 2
        dbgrdMovimento.Text = mHist�rico
        
        dbgrdMovimento.Col = 3
        dbgrdMovimento.Text = mD�bitoCr�dito
        
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
    
    Dim C�digoDoBanco As Long
    Dim C�digoDaAg�ncia As String
    Dim C�digoDaContaCorrente As String
    Dim ID As Long
    
    C�digoDoBanco = Trim(GetWordSeparatedBy(cmbBanco.Text, 1, "-"))
    C�digoDaAg�ncia = Trim(GetWordSeparatedBy(cmbAg�ncia.Text, 1, "-"))
    C�digoDaContaCorrente = Trim(GetWordSeparatedBy(cmbContaCorrente.Text, 1, "-"))
    
    'Pega o novo c�digo interno do produto e atualiza na Tabela Par�metros
    TBLPar�metros.Edit
    ID = TBLPar�metros("MOVIMENTO DE CONTA CORRENTE") + 1
    TBLPar�metros("MOVIMENTO DE CONTA CORRENTE") = ID
    TBLPar�metros.Update
    
    WS.BeginTrans
    
    TBLMovimentoContaCorrente.AddNew
    TBLMovimentoContaCorrente("ID") = ID
    TBLMovimentoContaCorrente("C�DIGO DO BANCO") = C�digoDoBanco
    TBLMovimentoContaCorrente("C�DIGO DA AG�NCIA") = C�digoDaAg�ncia
    TBLMovimentoContaCorrente("C�DIGO DA CONTA CORRENTE") = C�digoDaContaCorrente
    TBLMovimentoContaCorrente("HIST�RICO") = mHist�rico
    TBLMovimentoContaCorrente("VALOR") = mValor
    TBLMovimentoContaCorrente("D�BITO/CR�DITO") = mD�bitoCr�dito
    TBLMovimentoContaCorrente("DATA DA OPERA��O") = mData
    TBLMovimentoContaCorrente("SALDO") = mSaldo
    
    TBLMovimentoContaCorrente("USERNAME - CRIA") = gUsu�rio
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
    
    cmbAg�ncia.Clear
    cmbAg�ncia.Text = Empty
    
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
Private Sub cmbAg�ncia_Click()
    Dim C�digoBanco As Long
    Dim C�digoAg�ncia As String
    
    cmbAg�ncia.Locked = True
    
    cmbContaCorrente.Clear
    cmbContaCorrente.Text = Empty
    
    C�digoBanco = Trim(GetWordSeparatedBy(cmbBanco.Text, 1, "-"))
    C�digoAg�ncia = Trim(GetWordSeparatedBy(cmbAg�ncia.Text, 1, "-"))
    
    FillContaCorrente C�digoBanco, C�digoAg�ncia
End Sub
Private Sub cmbAg�ncia_DropDown()
    cmbAg�ncia.Locked = False
End Sub
Private Sub cmbBanco_Click()
    Dim C�digoBanco As Long
    cmbBanco.Locked = True
    cmbAg�ncia.Clear
    cmbAg�ncia.Text = Empty
    cmbContaCorrente.Clear
    cmbContaCorrente.Text = Empty
    
    C�digoBanco = Trim(GetWordSeparatedBy(cmbBanco.Text, 1, "-"))
    
    FillAg�ncia C�digoBanco
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
    
    If Not Ag�nciaAberto Then
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
    
    If TBLAg�ncia.RecordCount = 0 Then
        MsgBox "Nenhuma AG�NCIA foi cadastrada!", vbInformation, "Aviso"
        Unload Me
        Exit Sub
    End If
    
    If TBLContaCorrente.RecordCount = 0 Then
        MsgBox "Nenhuma CONTA CORRENTE foi cadastrada!", vbInformation, "Aviso"
        Unload Me
        Exit Sub
    End If
    
    Bot�oIncluir lAllowInsert
    Bot�oExcluir False
    Bot�oGravar False
    Bot�oImprimir False
    
    Navega��oInferior False
    Navega��oSuperior False
    
    StatusBarAviso = "Pronto"

    If lAtualizar Then
        Bot�oAtualizar True
    Else
        Bot�oAtualizar False
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
    MovimentoContaCorrenteAberto = AbreTabela(Dicion�rio, "FINANCEIRO", "MOVIMENTO - CONTA CORRENTE", DBFinanceiro, TBLMovimentoContaCorrente, TBLTabela, dbOpenTable)
    
    If MovimentoContaCorrenteAberto Then
        IndiceMovimentoContaCorrenteAtivo = "MOVIMENTOCONTACORRENTE1"
        TBLMovimentoContaCorrente.Index = IndiceMovimentoContaCorrenteAtivo
    Else
        MsgBox "N�o consegui abrir a tabela 'Movimento - Conta Corrente' !", vbCritical, "Erro"
        Exit Sub
    End If
    
    'Banco
    BancoAberto = AbreTabela(Dicion�rio, "FINANCEIRO", "BANCO", DBFinanceiro, TBLBanco, TBLTabela, dbOpenTable)
    
    If BancoAberto Then
        IndiceBancoAtivo = "BANCO1"
        TBLBanco.Index = IndiceBancoAtivo
    Else
        MsgBox "N�o consegui abrir a tabela 'Banco' !", vbCritical, "Erro"
        Exit Sub
    End If
    
    'Ag�ncia
    Ag�nciaAberto = AbreTabela(Dicion�rio, "FINANCEIRO", "AG�NCIA", DBFinanceiro, TBLAg�ncia, TBLTabela, dbOpenTable)
    
    If Ag�nciaAberto Then
        IndiceAg�nciaAtivo = "AG�NCIA1"
        TBLAg�ncia.Index = IndiceAg�nciaAtivo
    Else
        MsgBox "N�o consegui abrir a tabela 'Ag�ncia' !", vbCritical, "Erro"
        Exit Sub
    End If
    
    'Conta Corrente
    ContaCorrenteAberto = AbreTabela(Dicion�rio, "FINANCEIRO", "CONTA CORRENTE", DBFinanceiro, TBLContaCorrente, TBLTabela, dbOpenTable)
    
    If ContaCorrenteAberto Then
        IndiceContaCorrenteAtivo = "CONTACORRENTE1"
        TBLContaCorrente.Index = IndiceContaCorrenteAtivo
    Else
        MsgBox "N�o consegui abrir a tabela 'Conta Corrente' !", vbCritical, "Erro"
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
    dbgrdMovimento.Text = "Hist�rico"
    dbgrdMovimento.FontBold = True
    dbgrdMovimento.ColWidth(dbgrdMovimento.Col) = 30.75
    
    'D�bito/Cr�dito
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
    
    If Ag�nciaAberto Then
        TBLAg�ncia.Close
    End If
    
    If ContaCorrenteAberto Then
        TBLContaCorrente.Close
    End If
    
    Set frmEditaContaCorrente = Nothing
End Sub
