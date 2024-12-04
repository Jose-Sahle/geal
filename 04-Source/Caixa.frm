VERSION 5.00
Begin VB.Form frmCaixa 
   BackColor       =   &H00C0C000&
   BorderStyle     =   0  'None
   ClientHeight    =   9000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12000
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   24
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Caixa.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9000
   ScaleWidth      =   12000
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdCheque 
      Caption         =   "Impress�o de Cheques (F6)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Left            =   10500
      TabIndex        =   25
      TabStop         =   0   'False
      Top             =   8310
      Width           =   1245
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "Cancelar  (F4)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Left            =   9240
      TabIndex        =   24
      TabStop         =   0   'False
      Top             =   8310
      Width           =   1245
   End
   Begin VB.CommandButton cmdAjuda 
      Caption         =   "Ajuda (F1)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Left            =   8250
      TabIndex        =   23
      TabStop         =   0   'False
      Top             =   8310
      Width           =   975
   End
   Begin VB.CommandButton cmdFormaDePagamento 
      Caption         =   "Pagamento (F12)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Left            =   10500
      TabIndex        =   22
      TabStop         =   0   'False
      Top             =   7620
      Width           =   1245
   End
   Begin VB.Frame frTotalAPagar 
      BackColor       =   &H00C0C000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4545
      Left            =   7140
      TabIndex        =   18
      Top             =   10000
      Width           =   4635
      Begin VB.TextBox txtTotalAPagar 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         ForeColor       =   &H00FFFFFF&
         Height          =   585
         Left            =   210
         Locked          =   -1  'True
         TabIndex        =   19
         TabStop         =   0   'False
         Text            =   "Text1"
         Top             =   1710
         Width           =   4245
      End
      Begin VB.Label lblTotalAPagar 
         BackStyle       =   0  'Transparent
         Caption         =   "Total a Pagar"
         Height          =   555
         Left            =   120
         TabIndex        =   20
         Top             =   1050
         Width           =   3465
      End
      Begin VB.Line Line24 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   5
         X1              =   4470
         X2              =   4470
         Y1              =   1680
         Y2              =   2295
      End
      Begin VB.Line Line23 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   5
         X1              =   150
         X2              =   4470
         Y1              =   2310
         Y2              =   2310
      End
      Begin VB.Line Line22 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   5
         X1              =   150
         X2              =   4470
         Y1              =   1650
         Y2              =   1650
      End
      Begin VB.Line Line21 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   5
         X1              =   150
         X2              =   150
         Y1              =   1680
         Y2              =   2295
      End
   End
   Begin VB.Frame frVenda 
      BackColor       =   &H00C0C000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4635
      Left            =   7080
      TabIndex        =   9
      Top             =   240
      Width           =   4785
      Begin VB.TextBox txtPre�oUnit�rio 
         BackColor       =   &H00C00000&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   585
         Left            =   90
         Locked          =   -1  'True
         TabIndex        =   12
         TabStop         =   0   'False
         Text            =   "Text1"
         Top             =   540
         Width           =   4245
      End
      Begin VB.TextBox txtPre�oTotal 
         BackColor       =   &H00C00000&
         BorderStyle     =   0  'None
         ForeColor       =   &H00FFFFFF&
         Height          =   585
         Left            =   60
         Locked          =   -1  'True
         TabIndex        =   11
         TabStop         =   0   'False
         Text            =   "Text1"
         Top             =   3720
         Width           =   4245
      End
      Begin VB.TextBox txtQuantidade 
         BackColor       =   &H00C00000&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   585
         Left            =   90
         Locked          =   -1  'True
         TabIndex        =   10
         TabStop         =   0   'False
         Text            =   "Text1"
         Top             =   2100
         Width           =   4245
      End
      Begin VB.Line Line4 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   5
         X1              =   30
         X2              =   30
         Y1              =   510
         Y2              =   1125
      End
      Begin VB.Line Line6 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   5
         X1              =   30
         X2              =   4350
         Y1              =   480
         Y2              =   480
      End
      Begin VB.Line Line7 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   5
         X1              =   30
         X2              =   4350
         Y1              =   1140
         Y2              =   1140
      End
      Begin VB.Line Line8 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   5
         X1              =   4350
         X2              =   4350
         Y1              =   480
         Y2              =   1095
      End
      Begin VB.Label lblPre�oUnit�rio 
         BackStyle       =   0  'Transparent
         Caption         =   "Pre�o Unit�rio"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   30
         TabIndex        =   17
         Top             =   0
         Width           =   3465
      End
      Begin VB.Line Line9 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   5
         X1              =   30
         X2              =   30
         Y1              =   3690
         Y2              =   4305
      End
      Begin VB.Line Line10 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   5
         X1              =   30
         X2              =   4350
         Y1              =   3660
         Y2              =   3660
      End
      Begin VB.Line Line11 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   5
         X1              =   30
         X2              =   4350
         Y1              =   4320
         Y2              =   4320
      End
      Begin VB.Line Line12 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   5
         X1              =   4350
         X2              =   4350
         Y1              =   3690
         Y2              =   4305
      End
      Begin VB.Label lblPre�oTotal 
         BackStyle       =   0  'Transparent
         Caption         =   "Pre�o Total"
         Height          =   555
         Left            =   0
         TabIndex        =   16
         Top             =   3060
         Width           =   3465
      End
      Begin VB.Line Line13 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   5
         X1              =   30
         X2              =   30
         Y1              =   2070
         Y2              =   2685
      End
      Begin VB.Line Line14 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   5
         X1              =   30
         X2              =   4350
         Y1              =   2040
         Y2              =   2040
      End
      Begin VB.Line Line15 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   5
         X1              =   60
         X2              =   4380
         Y1              =   2700
         Y2              =   2700
      End
      Begin VB.Line Line16 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   5
         X1              =   4380
         X2              =   4380
         Y1              =   2040
         Y2              =   2655
      End
      Begin VB.Label lblQuantidade 
         BackStyle       =   0  'Transparent
         Caption         =   "Quantidade"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   30
         TabIndex        =   15
         Top             =   1560
         Width           =   3465
      End
      Begin VB.Label lblVezes 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "X"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   480
         TabIndex        =   14
         Top             =   1230
         Width           =   3465
      End
      Begin VB.Label lblIgual 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "="
         Height          =   405
         Left            =   540
         TabIndex        =   13
         Top             =   2670
         Width           =   3465
      End
   End
   Begin VB.ListBox lstCupom 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2400
      Left            =   7350
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   4950
      Width           =   4395
   End
   Begin VB.CommandButton cmdFecharCupom 
      Caption         =   "Fechar Cupom (F11)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Left            =   9720
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   7620
      Width           =   765
   End
   Begin VB.CommandButton cmdTotalizarCupom 
      Caption         =   "Totalizar Cupom (F10)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Left            =   8820
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   7620
      Width           =   885
   End
   Begin VB.CommandButton cmdAbrirCupom 
      Caption         =   "Abrir Cupom (F9)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Left            =   8040
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   7620
      Width           =   765
   End
   Begin VB.TextBox txtComando 
      BackColor       =   &H00C00000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   585
      Left            =   240
      TabIndex        =   3
      Top             =   8220
      Width           =   6885
   End
   Begin VB.TextBox txtDescri��o 
      BackColor       =   &H00C00000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   585
      Left            =   240
      Locked          =   -1  'True
      TabIndex        =   2
      TabStop         =   0   'False
      Text            =   "Text1"
      Top             =   5910
      Width           =   6885
   End
   Begin VB.PictureBox pctProdutoEmpresa 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4965
      Left            =   210
      Picture         =   "Caixa.frx":000C
      ScaleHeight     =   4905
      ScaleWidth      =   6555
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   240
      Width           =   6615
   End
   Begin VB.Label lblAviso 
      BackStyle       =   0  'Transparent
      Caption         =   "* - Itens cancelados"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   7410
      TabIndex        =   21
      Top             =   7320
      Width           =   3075
   End
   Begin VB.Label lblMensagem 
      BackStyle       =   0  'Transparent
      Caption         =   "Mensagem:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   585
      Left            =   150
      TabIndex        =   7
      Top             =   7680
      Width           =   6945
   End
   Begin VB.Line Line20 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   5
      X1              =   210
      X2              =   210
      Y1              =   8190
      Y2              =   8805
   End
   Begin VB.Line Line19 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   5
      X1              =   210
      X2              =   7160
      Y1              =   8160
      Y2              =   8160
   End
   Begin VB.Line Line18 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   5
      X1              =   210
      X2              =   7160
      Y1              =   8820
      Y2              =   8820
   End
   Begin VB.Line Line17 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   5
      X1              =   7155
      X2              =   7155
      Y1              =   8190
      Y2              =   8805
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   5
      X1              =   7155
      X2              =   7155
      Y1              =   5880
      Y2              =   6495
   End
   Begin VB.Line Line5 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   5
      X1              =   210
      X2              =   7160
      Y1              =   6510
      Y2              =   6510
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   5
      X1              =   210
      X2              =   7160
      Y1              =   5850
      Y2              =   5850
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   5
      X1              =   210
      X2              =   210
      Y1              =   5880
      Y2              =   6495
   End
   Begin VB.Label lblDescri��o 
      BackStyle       =   0  'Transparent
      Caption         =   "Descri��o:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   585
      Left            =   180
      TabIndex        =   1
      Top             =   5340
      Width           =   3465
   End
End
Attribute VB_Name = "frmCaixa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const MAXCOLS = 5
Const MAXCOLSPG = 3

Dim lAllowCancelLast As Boolean

Dim mScala As Single

Dim mTotalRows%
Dim ItensArray() As String
Dim mOr�amento As Long

Dim lOr�amento As Boolean
Dim lOr�amentoNovo As Boolean
Dim lPula As Boolean
Dim lAbriPorta As Boolean
Dim lFechar As Boolean
Dim lFechaNoFim As Boolean
Dim lPermiteCancelar As Boolean

Dim TBLVendas As Table
Dim VendasAberto As Boolean
Dim IndiceVendasAtivo$

Dim TBLVendasItens As Table
Dim VendasItensAberto As Boolean
Dim IndiceVendasItensAtivo$

Dim TBLVendasLote As Table
Dim VendasLoteAberto As Boolean
Dim IndiceVendasLotesAtivo$

Dim TBLPar�metros As Table
Dim Par�metrosAberto As Boolean

Dim TBLFormaDePagamento As Table
Dim FormaDePagamentoAberto As Boolean
Dim IndiceFormaDePagamentoAtivo$

Dim TBLPlanoDePagamento As Table
Dim PlanoDePagamentoAberto As Boolean
Dim IndicePlanoDePagamentoAtivo$

Dim TBLConfigura��oCaixa As Table
Dim Configura��oCaixaAberto As Boolean

Dim TBLCaixa As Table
Dim CaixaAberto As Boolean
Dim IndiceCaixaAtivo$

Dim TBLCaixaAbertura As Table
Dim CaixaAberturaAberto As Boolean
Dim IndiceCaixaAberturaAtivo$

Dim TBLCaixaMovimento As Table
Dim CaixaMovimentoAberto As Boolean
Dim IndiceCaixaMovimentoAtivo$

Dim TBLCaixaSangriaEntrada As Table
Dim CaixaSangriaEntradaAberto As Boolean
Dim IndiceCaixaSangriaEntradaAtivo$

Dim mRecnoPg%
Dim mTotalPagamentos As Integer
Dim mValorAVista As String
Dim mValorAPrazo As String
Dim mTipoDePagamento As Long
Dim FormaDePagamentoArray() As String

Dim CupomAberto As Boolean
Dim mTotal As Currency
Dim N�meroCaixa As String
Dim C�digoDoCaixa As Long
Dim C�digoDaAbertura As Long
Private Function AbrirCupom() As Boolean
    On Error GoTo ErroPDV
    
    Dim Status As String, AuxTexto As String
    
    Status = VerStatusECF

    If Not AbrirCupomFiscal Then
        GoTo ErroPDV
    Else
        If Mid(Status, 1, 2) = ".-" Then
            AuxTexto = Mid(Status, 3, 4)
            Status = Mid(Status, 7, Len(Status) - 7)
            MsgBox Status, vbCritical, "Erro #" & AuxTexto
            GoTo ErroPDV
        End If
    End If

    AbrirCupom = True
    
    Exit Function
    
ErroPDV:
    AbrirCupom = False
End Function
Private Function DetalheCupom() As Currency
    Dim Cont As Integer
    Dim ValoTotal As Currency
    Dim C�digo$, Quantidade$, Pre�oUnit�rio$, Pre�oTotal$, Descri��o$, Tributa��o$, Total$
    Dim Status$, ValorTotal As Currency, DescontoTotal As Currency
    Dim AuxValor$, AuxTexto$
    
    ValorTotal = 0
    For Cont = 0 To mTotalRows - 1
        If ItensArray(4, Cont) <> "C" Then
            C�digo = LeftBlankString(ItensArray(3, Cont), 13)
            Quantidade = LeftZeroString(ItensArray(2, Cont), 4) & "000"
            Pre�oUnit�rio = "0" & StrTran(FormatStringMask("@V 000000,00", ItensArray(1, Cont)), ",")
            Pre�oTotal = "0" & StrTran(FormatStringMask("@V 000000000,00", StrVal(ValStr(ItensArray(1, Cont)) * ValStr(ItensArray(2, Cont)))), ",")
            AuxTexto = Mid(SearchAdvancedProduto(ItensArray(0, Cont), vbDescri��o, vbIndice2), 1, 24)
            Descri��o = RightBlankString(AuxTexto, 24)
            
            Tributa��o = RightBlankString(SearchAdvancedProduto(ItensArray(0, Cont), vbTributo), 3)
            
            RegistrarItemVendido C�digo, Quantidade, Pre�oUnit�rio, Pre�oTotal, Descri��o, Tributa��o
            Status = VerStatusECF
            
            If Mid(Status, 1, 2) = ".-" Then
                AuxTexto = Mid(Status, 3, 4)
                Status = Mid(Status, 7, Len(Status) - 7)
                MsgBox Status, vbCritical, "Erro #" & AuxTexto
                DetalheCupom = 0
                Exit Function
            End If
            ValorTotal = ValorTotal + ValStr(ItensArray(1, Cont)) * ValStr(ItensArray(2, Cont))
        End If
    Next
        
    DetalheCupom = ValorTotal
End Function
Private Sub DoProduto(ByVal vC�digo As String)
    Dim pC�digo As String
    Dim plgC�digo As Long
    Dim pValor As Currency
    Dim pNumero As Long
    Dim pszValor As String
    
    pC�digo = UCase(vC�digo)
    plgC�digo = Val(SearchAdvancedProduto(pC�digo, vbC�digo))
               
    If plgC�digo = 0 Then
        frmEncontraProduto.Show 1
        plgC�digo = Val(frmEncontraProduto.C�digo)
    End If

    txtDescri��o = SearchAdvancedProduto(plgC�digo, vbDescri��o)
    
    vC�digo = SearchAdvancedProduto(plgC�digo, vbC�digoDoFornecedor, vbIndice2)
    pValor = SearchAdvancedProduto(plgC�digo, vbValValorUnit�rio, vbIndice2)
    
    Me.Font = txtPre�oUnit�rio.Font
    Me.FontBold = txtPre�oUnit�rio.FontBold
    Me.FontSize = txtPre�oUnit�rio.FontSize
    Me.ScaleWidth = mdiGeal.ScaleWidth
    
    pszValor = Trim(FormatStringMask("@V ##.###.##0,00", pValor))
    
    pNumero = Me.TextWidth("R$") + Me.TextWidth(pszValor) + 1
    pNumero = txtPre�oUnit�rio.Width - pNumero
    pNumero = Int(pNumero / Me.TextWidth(" "))
    txtPre�oUnit�rio = "R$" + String(pNumero, " ") + pszValor
    
    txtQuantidade = LeftBlankString("1", 24)
    
    Me.Font = txtPre�oTotal.Font
    Me.FontBold = txtPre�oTotal.FontBold
    Me.FontSize = txtPre�oTotal.FontSize
    Me.ScaleWidth = mdiGeal.ScaleWidth
    
    pNumero = Me.TextWidth("R$") + Me.TextWidth(pszValor) + 1
    pNumero = txtPre�oTotal.Width - pNumero
    pNumero = Int(pNumero / Me.TextWidth(" "))
    txtPre�oTotal = "R$" + String(pNumero, " ") + pszValor
    
    mTotalRows = mTotalRows + 1
    ReDim Preserve ItensArray(MAXCOLS - 1, mTotalRows - 1)
    ItensArray(0, mTotalRows - 1) = plgC�digo
    ItensArray(1, mTotalRows - 1) = StrVal(pValor)
    ItensArray(2, mTotalRows - 1) = "1"
    ItensArray(3, mTotalRows - 1) = vC�digo
    ItensArray(4, mTotalRows - 1) = "A"
End Sub
Private Sub DoQuantidade(ByVal vQuantidade As Long)
    Dim pValor As Currency
    Dim pNumero As Long
    Dim pszValor
        
    If vQuantidade > 1 Then
        pValor = ValStr(ItensArray(1, mTotalRows - 1))
        pValor = pValor * vQuantidade
        
        txtQuantidade = LeftBlankString(vQuantidade, 24)
        
        Me.Font = txtPre�oTotal.Font
        Me.FontBold = txtPre�oTotal.FontBold
        Me.FontSize = txtPre�oTotal.FontSize
        Me.ScaleWidth = mdiGeal.ScaleWidth
        pszValor = FormatStringMask("@V ##.###.##0,00", pValor)
        pNumero = Me.TextWidth("R$") + Me.TextWidth(pszValor) + 1
        pNumero = txtPre�oTotal.Width - pNumero
        pNumero = Int(pNumero / Me.TextWidth(" "))
        txtPre�oTotal = "R$" + String(pNumero, " ") + pszValor
        
        ItensArray(2, mTotalRows - 1) = Str(vQuantidade)
    End If
    
    If lstCupom.ListCount > 0 Then
        lstCupom.RemoveItem lstCupom.ListCount - 1
        lstCupom.RemoveItem lstCupom.ListCount - 1
        lstCupom.RemoveItem lstCupom.ListCount - 1
    End If
    
    Me.Font = lstCupom.Font
    Me.FontBold = lstCupom.FontBold
    Me.FontSize = lstCupom.FontSize
    Me.ScaleWidth = mdiGeal.ScaleWidth
    
    lstCupom.AddItem RightBlankString(ItensArray(3, mTotalRows - 1), 13) & "[" & Mid(txtDescri��o, 1, 24) & "]"
    lstCupom.AddItem LeftBlankString(ItensArray(2, mTotalRows - 1), 7) & Space(2) & LeftBlankString(StrTran(txtPre�oUnit�rio, "R$"), 9) & " = " & StrTran(txtPre�oTotal, "R$")
    lstCupom.AddItem " "
    
    FillSubTotal
End Sub
Private Function FechaCupom() As Boolean
    Dim Status As String, AuxTexto As String
    
    Status = String(255, " ")
    
    FechaCupom = FecharCupomFiscal
    Status = VerStatusECF
    
    If Mid(Status, 1, 2) = ".-" Then
        AuxTexto = Mid(Status, 3, 4)
        Status = Mid(Status, 7, Len(Status) - 7)
        MsgBox Status, vbCritical, "Erro #" & AuxTexto
        FechaCupom = False
        Exit Function
    End If
End Function
Private Sub FillSubTotal()
    Dim pValor As Currency
    Dim pNumero As Long
    Dim pszValor As String
    
    lstCupom.AddItem " "
    lstCupom.AddItem String(42, "_")
    
    pValor = Subtotal
    
    pszValor = Trim(FormatStringMask("@V ##.###.##0,00", pValor))
    pNumero = Me.TextWidth("Sub Total.......................R$") + Me.TextWidth(pszValor) + 1
    pNumero = lstCupom.Width - pNumero - 200
    pNumero = Int(pNumero / Me.TextWidth(" "))
    lstCupom.AddItem "Sub Total.......................R$" + String(pNumero, " ") + pszValor
    
    lstCupom.ListIndex = lstCupom.ListCount - 1
    lstCupom.ListIndex = -1
End Sub
Private Function GravaBaseDeDados() As Boolean
    Dim Cont As Byte
    Dim C�digoDoLote As String
    Dim D�gitoDoLote As String
    
    If mOr�amento <> 0 Then
        If PosRecordsOr�amento(mOr�amento) Then
            TBLVendas.Edit
            TBLVendas("TIPO") = "V"
            TBLVendas("DATA DA VENDA") = Date
            TBLVendas("USERNAME - ALTERA") = gUsu�rio
            TBLVendas("DATA - ALTERA") = Date
            TBLVendas("HORA - ALTERA") = Time
            TBLVendas.Update
            
            If PosRecordsItens(mOr�amento) Then
                Do While Not TBLVendasItens.EOF And TBLVendasItens("OR�AMENTO") = mOr�amento
                    If Not AtualizaProduto(TBLVendasItens("C�DIGO DO PRODUTO"), "-", TBLVendasItens("QUANTIDADE")) Then
                        GoTo ErroVendas
                    End If
                    
                    TBLVendasLote.Seek "=", mOr�amento
                    If Not TBLVendasLote.NoMatch Then
                        Do While Not TBLVendasLote.EOF And TBLVendasLote("OR�AMENTO") = mOr�amento
                            C�digoDoLote = GetWordSeparatedBy(TBLVendasLote("C�DIGO DO LOTE"), 1, "-")
                            D�gitoDoLote = GetWordSeparatedBy(TBLVendasLote("C�DIGO DO LOTE"), 1, "-")
                            
                            If Not AtualizaLote(TBLVendasLote("C�DIGO DO PRODUTO"), C�digoDoLote, D�gitoDoLote, TBLVendasLote("QUANTIDADE"), TBLVendasLote("M�LTIPLO")) Then
                                GoTo ErroVendas
                            End If
                            TBLVendasLote.MoveNext
                            
                            If TBLVendasLote.EOF Then
                                Exit Do
                            End If
                        Loop
                    End If
                    
                    TBLVendasItens.MoveNext
                    If TBLVendasItens.EOF Then
                        Exit Do
                    End If
                Loop
            End If
            
            If Not PosRecordsFormaDePagamento(mOr�amento) Then
                PosRecordsOr�amento mOr�amento
                TBLVendas.Edit
                TBLVendas("TIPO DE PAGAMENTO") = mTipoDePagamento
                TBLVendas("QUANTIDADE DE VENCIMENTOS") = mTotalPagamentos
                TBLVendas("VALOR A PRAZO") = ValStr(mValorAPrazo)
                TBLVendas("USERNAME - ALTERA") = gUsu�rio
                TBLVendas("DATA - ALTERA") = Date
                TBLVendas("HORA - ALTERA") = Time
                TBLVendas.Update
                
                PosRecordsOr�amento mOr�amento
                
                For Cont = 0 To mTotalPagamentos - 1
                    TBLFormaDePagamento.AddNew
                    TBLFormaDePagamento("OR�AMENTO") = mOr�amento
                    TBLFormaDePagamento("DOCUMENTO") = FormaDePagamentoArray(0, Cont)
                    TBLFormaDePagamento("VENCIMENTO") = IIf(Trim(StrTran(FormaDePagamentoArray(1, Cont), "/")) <> Empty, FormaDePagamentoArray(1, Cont), vbNull)
                    TBLFormaDePagamento("VALOR") = StrVal(FormaDePagamentoArray(2, Cont))
                    TBLFormaDePagamento.Update
                Next
            End If
        End If
    Else
        'Pega o novo c�digo interno do produto e atualiza na Tabela Par�metros
        TBLPar�metros.Edit
        mOr�amento = TBLPar�metros("OR�AMENTO") + 1
        TBLPar�metros("OR�AMENTO") = mOr�amento
        TBLPar�metros.Update
        
        'Inclus�o da venda (cabe�alho)
        TBLVendas.AddNew
        TBLVendas("C�DIGO") = Val(mOr�amento)
        'TBLVendas("CGC - CPF") = "000.000.000-00"
        TBLVendas("C�DIGO DO CLIENTE") = 0
        TBLVendas("TIPO") = "V"
        TBLVendas("VALOR TOTAL DA VENDA") = Subtotal
        TBLVendas("DESCONTO TOTAL DA VENDA") = 0
        TBLVendas("VALOR DO BONUS") = 0
        TBLVendas("DATA DA VENDA") = Date
        TBLVendas("DATA DO OR�AMENTO") = Date
        TBLVendas("BAIXADO") = False
        TBLVendas("TIPO DE PAGAMENTO") = mTipoDePagamento
        TBLVendas("QUANTIDADE DE VENCIMENTOS") = mTotalPagamentos
        TBLVendas("VALOR A PRAZO") = ValStr(mValorAPrazo)
        TBLVendas("OBSERVA��O") = vbNull
        TBLVendas("USERNAME - CRIA") = gUsu�rio
        TBLVendas("DATA - CRIA") = Date
        TBLVendas("HORA - CRIA") = Time
        TBLVendas("USERNAME - ALTERA") = gUsu�rio
        TBLVendas("DATA - ALTERA") = Date
        TBLVendas("HORA - ALTERA") = Time
        TBLVendas.Update
        
        'Itens da venda
        For Cont = 0 To mTotalRows - 1
            If ItensArray(4, Cont) <> "C" Then
                TBLVendasItens.AddNew
                TBLVendasItens("OR�AMENTO") = Val(mOr�amento)
                TBLVendasItens("C�DIGO DO PRODUTO") = SearchAdvancedProduto(ItensArray(3, Cont), vbC�digo, vbIndice3)
                TBLVendasItens("QUANTIDADE") = ValStr(ItensArray(2, Cont))
                TBLVendasItens("VALOR UNIT�RIO") = ValStr(ItensArray(1, Cont))
                TBLVendasItens("DESCONTO") = 0
                TBLVendasItens.Update
            End If
        Next
        
        'Forma de Pagamento
        For Cont = 0 To mTotalPagamentos - 1
            TBLFormaDePagamento.AddNew
            TBLFormaDePagamento("OR�AMENTO") = mOr�amento
            TBLFormaDePagamento("DOCUMENTO") = FormaDePagamentoArray(0, Cont)
            TBLFormaDePagamento("VENCIMENTO") = IIf(Trim(StrTran(FormaDePagamentoArray(1, Cont), "/")) <> Empty, FormaDePagamentoArray(1, Cont), vbNull)
            TBLFormaDePagamento("VALOR") = StrVal(FormaDePagamentoArray(2, Cont))
            TBLFormaDePagamento.Update
        Next
    End If
    
    'Inclui na tabela de Movimento de Caixa
    TBLCaixaMovimento.AddNew
    TBLCaixaMovimento("C�DIGO DA ABERTURA") = C�digoDaAbertura
    TBLCaixaMovimento("OR�AMENTO") = mOr�amento
    TBLCaixaMovimento("HORA") = Time
    TBLCaixaMovimento.Update
    
    GravaBaseDeDados = True
    
    Exit Function
ErroVendas:
    GravaBaseDeDados = False
End Function
Public Function GravaPagamento(ByRef Matriz() As String) As Boolean
    Dim Cont As Integer, Cont1 As Integer
    
    ReDim FormaDePagamentoArray(MAXCOLSPG, UBound(Matriz, 2) + 1)
    
    mTotalPagamentos = frmFormaDePagamento.mTotalPagamentos
    mTipoDePagamento = frmFormaDePagamento.mTipoDePagamento
    
    For Cont = 0 To UBound(Matriz, 2)
        For Cont1 = 0 To MAXCOLSPG - 1
            FormaDePagamentoArray(Cont1, Cont) = Matriz(Cont1, Cont)
        Next
    Next
    
    GravaPagamento = True
    
End Function
Private Sub IniciaCaixa()
    ZeraCampos
End Sub
Private Function Inicializa��o() As Boolean
    TBLCaixa.Seek "=", N�meroCaixa, True, False
    
    If TBLCaixa.NoMatch Then
        MsgBox "O Caixa dever ser aberto antes de se escolher esta op��o", vbInformation, "Aviso"
        Inicializa��o = False
        Exit Function
    End If
    
    frmAberturaOpera��oDoCaixa.Show 1
    
    C�digoDoCaixa = frmAberturaOpera��oDoCaixa.C�digoDoCaixa
    C�digoDaAbertura = frmAberturaOpera��oDoCaixa.C�digoDaAbertura
    Inicializa��o = frmAberturaOpera��oDoCaixa.lSuccessfull
    
    Set frmAberturaOpera��oDoCaixa = Nothing
End Function
Private Sub NovaVenda()
    mTotalRows = 0
    lOr�amentoNovo = True
    ReDim ItensArray(MAXCOLS - 1, mTotalRows) As String
End Sub
Private Function PosRecordsFormaDePagamento(ByVal C�digo As Long) As Boolean
    TBLFormaDePagamento.Index = "VENDAFORMADEPAGAMENTO1"
    TBLFormaDePagamento.Seek ">=", C�digo
    If TBLFormaDePagamento.NoMatch Then
        PosRecordsFormaDePagamento = False
    Else
        If TBLFormaDePagamento("OR�AMENTO") <> C�digo Then
            PosRecordsFormaDePagamento = False
        Else
            PosRecordsFormaDePagamento = True
        End If
    End If
End Function
Private Function PosRecordsItens(ByVal C�digo As Long) As Boolean
    TBLVendasItens.Seek "=", C�digo
    
    If TBLVendasItens.NoMatch Then
        MsgBox "Or�amento n�o cadastrado!", vbInformation, "Aviso"
        PosRecordsItens = False
    Else
        PosRecordsItens = True
    End If
End Function
Private Function PosRecordsOr�amento(ByVal C�digo As Long) As Boolean
    TBLVendas.Seek "=", C�digo
    
    If TBLVendas.NoMatch Then
        PosRecordsOr�amento = False
        MsgBox "Or�amento n�o foi encontrado", vbInformation, "Aviso"
    Else
        PosRecordsOr�amento = True
    End If
End Function
Private Function PreencheOr�amento() As Boolean
    Dim pDescri��o As String
    
    mOr�amento = txtComando
    If Not PosRecordsOr�amento(mOr�amento) Then
        lOr�amentoNovo = False
        PreencheOr�amento = False
        Exit Function
    Else
        If TBLVendas("TIPO") <> "O" Then
            MsgBox "Or�amento j� esta fechado!", vbInformation, "Aviso"
            lOr�amentoNovo = False
            PreencheOr�amento = False
            Exit Function
        End If
    End If
    
    If Not PosRecordsItens(mOr�amento) Then
        lOr�amentoNovo = False
        PreencheOr�amento = False
        Exit Function
    End If
    
    Do While TBLVendasItens("OR�AMENTO") = Val(txtComando)
        'Preenche a matriz
        mTotalRows = mTotalRows + 1
        ReDim Preserve ItensArray(MAXCOLS - 1, mTotalRows - 1)
        ItensArray(0, mTotalRows - 1) = TBLVendasItens("C�DIGO DO PRODUTO")
        ItensArray(1, mTotalRows - 1) = StrVal(TBLVendasItens("VALOR UNIT�RIO"))
        ItensArray(2, mTotalRows - 1) = Str(TBLVendasItens("QUANTIDADE"))
        ItensArray(3, mTotalRows - 1) = SearchAdvancedProduto(TBLVendasItens("C�DIGO DO PRODUTO"), vbC�digoDoFornecedor, vbIndice2)
        ItensArray(4, mTotalRows - 1) = "O"
        
        'Imprime no ListBox
        Me.Font = lstCupom.Font
        Me.FontBold = lstCupom.FontBold
        Me.FontSize = lstCupom.FontSize
        Me.ScaleWidth = mdiGeal.ScaleWidth
        
        pDescri��o = SearchAdvancedProduto(TBLVendasItens("C�DIGO DO PRODUTO"), vbDescri��o)
        
        lstCupom.AddItem RightBlankString(ItensArray(3, mTotalRows - 1), 13) & "[" & Mid(pDescri��o, 1, 24) & "]"
        lstCupom.AddItem LeftBlankString(ItensArray(2, mTotalRows - 1), 7) & Space(2) & LeftBlankString(FormatStringMask("@V ##.###.##0,00", TBLVendasItens("VALOR UNIT�RIO")), 9) & " = " & FormatStringMask("@V ##.###.##0,00", TBLVendasItens("VALOR UNIT�RIO") * TBLVendasItens("QUANTIDADE"))
        lstCupom.AddItem " "
                
        TBLVendasItens.MoveNext
        
        If TBLVendasItens.EOF Then
            Exit Do
        End If
    Loop
    
    FillSubTotal
    
    PreencheOr�amento = True
    
    lOr�amentoNovo = True
End Function
Private Sub Receber()
    Dim pValor1 As Currency
    Dim pValor2 As Currency
    
    lPula = True
    FormatMask "@V ##.###.##0,00", txtComando
    lPula = False
    
    pValor1 = ValStr(StrTran(txtTotalAPagar, "R$"))
    pValor2 = ValStr(txtComando)
    
    If pValor1 > pValor2 Then
        MsgBox "Valor inv�lido !", vbCritical, "Aviso"
        Exit Sub
    End If
    
    mTotal = pValor2
    
    pValor1 = pValor2 - pValor1
    
    lblMensagem = "Troco"
    txtComando = "R$" + String(6, " ") + FormatStringMask("@V ##.###.##0,00", pValor1)
    
    If Not PosRecordsFormaDePagamento(mOr�amento) Then
        cmdFormaDePagamento.Enabled = True
    Else
        cmdFecharCupom.Enabled = True
    End If
End Sub
Private Sub SetKey(ByVal KeyCode As Integer, Shift As Integer)
    On Error GoTo Erro
    
    Dim Cont As Integer, Inicio As Integer, Fim As Integer
    
    If KeyCode = 27 Then 'Tecla ESC - Fechar opera��o
        If Not CupomAberto Then
            frmFechamentoOpera��oDoCaixa.C�digoDaAbertura = C�digoDaAbertura
            frmFechamentoOpera��oDoCaixa.C�digoDoCaixa = C�digoDoCaixa
            frmFechamentoOpera��oDoCaixa.Show 1
            If frmFechamentoOpera��oDoCaixa.lSuccessfull Then
                Unload Me
            End If
        End If
    ElseIf KeyCode = 13 Then 'Tecla ENTER
        Select Case lblMensagem
            Case "Or�amento"
                If txtComando = Empty Then
                    NovaVenda
                Else
                    If Not PreencheOr�amento Then
                        Exit Sub
                    End If
                End If
                lblMensagem = "C�digo do Produto"
                txtComando = Empty
            Case "C�digo do Produto"
                lOr�amento = False
                DoProduto txtComando
                lblMensagem = "Quantidade"
                lPermiteCancelar = True
                txtComando = Empty
            Case "Quantidade"
                If txtComando = Empty Then
                    txtComando = "1"
                End If
                DoQuantidade txtComando
                lblMensagem = "C�digo do Produto"
                txtComando = Empty
            Case "Valor"
                Receber
                txtComando.Locked = False
        End Select
        txtComando.SetFocus
    ElseIf KeyCode = 8 And ((txtComando = Empty And lblMensagem <> "Troco") Or (txtComando <> Empty And lblMensagem = "Troco")) And mTotalRows > 0 And lAllowCancelLast Then 'Tecla BACKSPACE - Retornar item anterior
        Select Case lblMensagem
            Case "Or�amento"
            Case "C�digo do Produto"
                If Not lPermiteCancelar Then
                    MsgBox "Item n�o pode ser cancelado ou alterado !", vbInformation, "Aviso"
                    Exit Sub
                End If
                If lOr�amento Then
                    IniciaCaixa
                    cmdAbrirCupom_Click
                    lblMensagem = "Or�amento"
                Else
                    For Cont = 1 To 3
                        lstCupom.RemoveItem lstCupom.ListCount - 1
                    Next
                    lblMensagem = "Quantidade"
                End If
                txtComando = Empty
            Case "Quantidade"
                If mTotalRows = 0 Then
                    ReDim ItensArray(MAXCOLS - 1, mTotalRows)
                Else
                    mTotalRows = mTotalRows - 1
                    ReDim Preserve ItensArray(MAXCOLS - 1, mTotalRows - 1)
                End If
                For Cont = 1 To 3
                    lstCupom.RemoveItem lstCupom.ListCount - 1
                Next
                FillSubTotal
                lPermiteCancelar = False
                lblMensagem = "C�digo do Produto"
                txtComando = Empty
            Case "Valor"
            Case "Troco"
                lblMensagem = "Valor"
                txtComando.Locked = False
                txtComando = Empty
        End Select
    ElseIf KeyCode = 112 Then 'Tecla F1 - Ajuda
        cmdAjuda_Click
    ElseIf KeyCode = 115 Then 'Tecla F4 - Cancelar
        cmdCancelar_Click
        txtComando.SetFocus
    ElseIf KeyCode = 120 Then 'Tecla F9 - Abrir cupom
        If cmdAbrirCupom.Enabled Then
            cmdAbrirCupom_Click
            txtComando.SetFocus
        End If
    ElseIf KeyCode = 121 Then 'Tecla F10 - Totalizar cupom
        If cmdTotalizarCupom.Enabled Then
            cmdTotalizarCupom_Click
            txtComando.SetFocus
        End If
    ElseIf KeyCode = 122 Then 'Tecla F11 - Fechar cupom
        If cmdFecharCupom.Enabled Then
            cmdFecharCupom_Click
            txtComando.SetFocus
        End If
    ElseIf KeyCode = 123 Then 'Tecla F12 - Forma de pagamento
        If cmdFormaDePagamento.Enabled Then
            cmdFormaDePagamento_Click
            txtComando.SetFocus
        End If
    ElseIf KeyCode = 83 And Shift = 2 Then 'CTRL-S - Sangria
        frmSangriaEntrada.Tipo = "S"
        frmSangriaEntrada.C�digoDaAbertura = C�digoDaAbertura
        frmSangriaEntrada.N�meroCaixa = N�meroCaixa
        frmSangriaEntrada.T�tulo = "Sangrida do Caixa"
        frmSangriaEntrada.Show 1
    ElseIf KeyCode = 69 And Shift = 2 Then 'CTRL-E - Entrada
        frmSangriaEntrada.Tipo = "E"
        frmSangriaEntrada.C�digoDaAbertura = C�digoDaAbertura
        frmSangriaEntrada.N�meroCaixa = N�meroCaixa
        frmSangriaEntrada.T�tulo = "Entrada no Caixa"
        frmSangriaEntrada.Show 1
    End If
    
    Exit Sub
    
Erro:
    GeraMensagemDeErro " "
End Sub
Private Function Subtotal() As Currency
    Dim Cont As Integer
    Dim pValor As Currency
    
    pValor = 0
    For Cont = 0 To mTotalRows - 1
        If ItensArray(4, Cont) <> "C" Then
            pValor = pValor + StrVal(ItensArray(1, Cont)) * Val(ItensArray(2, Cont))
        End If
    Next
    
    Subtotal = pValor
End Function
Private Function TotalizaCupom() As Boolean
    Dim Total As String
    Dim Status As String, AuxTexto As String
    
    Status = String(255, " ")
    
    Total = StrTran(FormatStringMask("@V 0000000000,00", StrVal(mTotal)), ",")
    
    TotalizaCupom = TotalizarCupomFiscal(Total)
    
    Status = VerStatusECF
    
    If Mid(Status, 1, 2) = ".-" Then
        AuxTexto = Mid(Status, 3, 4)
        Status = Mid(Status, 7, Len(Status) - 7)
        MsgBox Status, vbCritical, "Erro #" & AuxTexto
        TotalizaCupom = False
    End If
End Function
Private Sub ZeraCampos()
    mTotalRows = 0
    ReDim ItensArray(MAXCOLS - 1, mTotalRows) As String
    cmdAbrirCupom.Enabled = True
    cmdTotalizarCupom.Enabled = False
    cmdFecharCupom.Enabled = False
    cmdFormaDePagamento.Enabled = False
    CupomAberto = False
    txtComando.Locked = True
    frTotalAPagar.Visible = False
    frVenda.Visible = True
    frTotalAPagar.Top = 240
    frVenda.Top = 240
    lblMensagem = Empty
    txtDescri��o = Empty
    txtPre�oUnit�rio = Empty
    txtPre�oTotal = Empty
    txtQuantidade = Empty
    txtTotalAPagar = Empty
    txtComando = Empty
    lstCupom.Clear
    mOr�amento = 0
    lOr�amentoNovo = False
End Sub
Private Sub cmdAbrirCupom_Click()
    On Error GoTo ErroPDV
    
    Dim Status$
        
    CupomAberto = True
    txtComando.Locked = False
    txtComando.SetFocus
    
    lblMensagem = "Or�amento"
    
    cmdAbrirCupom.Enabled = False
    cmdTotalizarCupom.Enabled = True
    
    lOr�amento = True
    
    Exit Sub
    
ErroPDV:
    CupomAberto = False
    txtComando = Status
End Sub
Private Sub cmdAjuda_Click()
    MsgBox "F1" & vbTab & " - Ajuda" & vbCrLf & _
           "F4" & vbTab & " - Cancelar or�amento atual" & vbCrLf & _
           "F9" & vbTab & " - Abrir cupom fiscal" & vbCrLf & _
           "F10" & vbTab & " - Totalizar cupom fiscal" & vbCrLf & _
           "F11" & vbTab & " - Fechar cupom fiscal" & vbCrLf & _
           "F12" & vbTab & " - Forma de Pagamento" & vbCrLf & _
           "Ctrl-E" & vbTab & " - Entrada no caixa" & vbCrLf & _
           "Ctrl-S" & vbTab & " - Sangria" & vbCrLf & _
           "Esc" & vbTab & " - Fechar caixa", vbInformation, "Aviso"
End Sub
Private Sub cmdCancelar_Click()
    If MsgBox("Tem certeza que quer cancelar este or�amento?", vbQuestion + vbYesNo + vbDefaultButton2, "Confirma��o") = vbYes Then
        IniciaCaixa
    End If
End Sub
Private Sub cmdCheque_Click()
    frmImpressaoDeCheque.Show vbModal
End Sub
Private Sub cmdFecharCupom_Click()
    On Error GoTo Erro
    
    Dim ValorTotal As Currency, DescontoTotal As Currency
    Dim AuxValor As String, AuxTexto As String
    
    If lFechaNoFim Then
        
        WS.BeginTrans
        
        If Not GravaBaseDeDados Then
            GoTo Erro
        End If
        
        If Not AbrirCupom Then
            Exit Sub
        End If
        
        ValorTotal = DetalheCupom
        
        If DescontoTotal > 0 Then
            AuxValor = StrTran(FormatStringMask("@V 0000000000,00", StrVal(DescontoTotal)), ",")
            AuxTexto = FormatStringMask("@V ##%", StrVal(DescontoTotal * 100 / ValorTotal))
            AuxTexto = RightBlankString(AuxTexto, 10)
            DescontoSobreCupomFiscal AuxTexto, AuxValor
        End If
        
        If Not TotalizaCupom Then
            GoTo Cancela
        End If
        
        If Not FechaCupom Then
            WS.Rollback
            Exit Sub
        End If
        
        WS.CommitTrans
    End If
    
    CupomAberto = False
    IniciaCaixa
    
    Exit Sub
    
Cancela:
    FechaCupom
    WS.Rollback
    CupomAberto = False
    IniciaCaixa
    Exit Sub
    
Erro:
    GeraMensagemDeErro "Caixa - Fechar Cupom", True
End Sub
Private Sub cmdFormaDePagamento_Click()
    frmFormaDePagamento.mValorAVista = Trim(StrTran(Subtotal, "R$"))
    frmFormaDePagamento.mTotalPagamentos = mTotalPagamentos
    frmFormaDePagamento.mTipoDePagamento = mTipoDePagamento
    frmFormaDePagamento.lEdit = True
    frmFormaDePagamento.lCaixa = True
    frmFormaDePagamento.lNotCancel = True
    frmFormaDePagamento.lCompra = False
    Set frmFormaDePagamento.ptrForm = Me
    Set frmFormaDePagamento.TBLPlanoDePagamento = TBLPlanoDePagamento
    frmFormaDePagamento.Show 1
    cmdFecharCupom.Enabled = True
    cmdFormaDePagamento.Enabled = False
End Sub
Private Sub cmdTotalizarCupom_Click()
    Dim Cont As Integer
    Dim pValor As Currency
    Dim pNumero As Long
    Dim pszValor As String
    
    pValor = Subtotal
    
    Me.Font = txtTotalAPagar.Font
    Me.FontBold = txtTotalAPagar.FontBold
    Me.FontSize = txtTotalAPagar.FontSize
    Me.ScaleWidth = mdiGeal.ScaleWidth
    pszValor = FormatStringMask("@V ##.###.##0,00", pValor)
    pNumero = Me.TextWidth("R$") + Me.TextWidth(pszValor) + 1
    pNumero = txtTotalAPagar.Width - pNumero
    pNumero = Int(pNumero / Me.TextWidth(" "))
    txtTotalAPagar = "R$" + String(pNumero, " ") + FormatStringMask("@V ##.###.##0,00", pValor)
    
    cmdTotalizarCupom.Enabled = False
    frVenda.Visible = False
    frTotalAPagar.Visible = True
    lblMensagem = "Valor"
End Sub
Private Sub Form_Activate()
    If lFechar Then
        Unload Me
        Exit Sub
    End If
    If Not VendasAberto Then
        Unload Me
        Exit Sub
    End If
    If Not VendasItensAberto Then
        Unload Me
        Exit Sub
    End If
    If Not VendasLoteAberto Then
        Unload Me
        Exit Sub
    End If
    If Not Par�metrosAberto Then
        Unload Me
        Exit Sub
    End If
End Sub
Private Sub Form_DblClick()
    Unload Me
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    SetKey KeyCode, Shift
End Sub
Private Sub Form_Load()
    On Error GoTo Erro
    
    lPula = False
    lFechar = False
    lFechaNoFim = True
    
    If Not Allow("CAIXA", "O", gUsu�rio) Then
        MsgBox "Usu�rio n�o tem direitos para operar caixa", vbCritical, "Aviso"
        lFechar = True
        Exit Sub
    End If
    
    lAllowCancelLast = Allow("CAIXA", "U", gUsu�rio)
    
    IniciaCaixa
            
    Configura��oCaixaAberto = AbreTabela(Dicion�rio, "SISTEMA", "CAIXA", DBSistema, TBLConfigura��oCaixa, TBLTabela, dbOpenTable)
    
    If Configura��oCaixaAberto Then
        TBLConfigura��oCaixa.Index = "CAIXA1"
    Else
        MsgBox "N�o consegui abrir a tabela 'Configura��o de Caixa' !", vbCritical, "Erro"
        Exit Sub
    End If
    
    N�meroCaixa = GetRegistryString("Caixa", "Posto", "N�mero", "")
    
    If N�meroCaixa <> Empty Then
        TBLConfigura��oCaixa.Seek "=", N�meroCaixa
        If TBLConfigura��oCaixa.NoMatch Then
            MsgBox "Existe uma inconsist�ncia no Posto de Caixa " & N�meroCaixa, vbCritical, "Inconsist�ncia"
            lFechar = True
            Exit Sub
        End If
    Else
        MsgBox "Existe uma inconsist�ncia no Posto de Caixa " & N�meroCaixa, vbCritical, "Inconsist�ncia"
        lFechar = True
        Exit Sub
    End If
    
    VendasAberto = AbreTabela(Dicion�rio, "FINANCEIRO", "VENDA", DBFinanceiro, TBLVendas, TBLTabela, dbOpenTable)
    
    If VendasAberto Then
        IndiceVendasAtivo = "VENDA1"
        TBLVendas.Index = IndiceVendasAtivo
    Else
        MsgBox "N�o consegui abrir a tabela 'Vendas' !", vbCritical, "Erro"
        GoTo Erro
    End If
    
    VendasItensAberto = AbreTabela(Dicion�rio, "FINANCEIRO", "VENDA - ITENS", DBFinanceiro, TBLVendasItens, TBLTabela, dbOpenTable)
    
    If VendasItensAberto Then
        IndiceVendasItensAtivo = "VENDAITENS1"
        TBLVendasItens.Index = IndiceVendasItensAtivo
    Else
        MsgBox "N�o consegui abrir a tabela 'Itens de Venda' !", vbCritical, "Erro"
        GoTo Erro
    End If
    
    VendasLoteAberto = AbreTabela(Dicion�rio, "FINANCEIRO", "VENDA - LOTES", DBFinanceiro, TBLVendasLote, TBLTabela, dbOpenTable)
    
    If VendasLoteAberto Then
        IndiceVendasLotesAtivo = "VENDALOTES2"
        TBLVendasLote.Index = IndiceVendasLotesAtivo
    Else
        MsgBox "N�o consegui abrir a tabela 'Lote de Vendas' !", vbCritical, "Erro"
        GoTo Erro
    End If
    
    FormaDePagamentoAberto = AbreTabela(Dicion�rio, "FINANCEIRO", "VENDA - FORMA DE PAGAMENTO", DBFinanceiro, TBLFormaDePagamento, TBLTabela, dbOpenTable)
    
    If FormaDePagamentoAberto Then
        IndiceFormaDePagamentoAtivo = "VENDAFORMADEPAGAMENTO2"
        TBLFormaDePagamento.Index = IndiceFormaDePagamentoAtivo
    Else
        MsgBox "N�o consegui abrir a tabela 'Forma de Pagamento - Venda' !", vbCritical, "Erro"
        GoTo Erro
    End If
    
    PlanoDePagamentoAberto = AbreTabela(Dicion�rio, "FINANCEIRO", "PLANO DE PAGAMENTO", DBFinanceiro, TBLPlanoDePagamento, TBLTabela, dbOpenTable)
    
    If PlanoDePagamentoAberto Then
        IndicePlanoDePagamentoAtivo = "PLANODEPAGAMENTO1"
        TBLPlanoDePagamento.Index = IndicePlanoDePagamentoAtivo
    Else
        MsgBox "N�o consegui abrir a tabela 'Forma de Pagamento' !", vbCritical, "Erro"
        GoTo Erro
    End If
    
    CaixaAberto = AbreTabela(Dicion�rio, "FINANCEIRO", "CAIXA", DBFinanceiro, TBLCaixa, TBLTabela, dbOpenTable)
    
    If CaixaAberto Then
        IndiceCaixaAtivo = "CAIXA3"
        TBLCaixa.Index = IndiceCaixaAtivo
    Else
        MsgBox "N�o consegui abrir a tabela 'CAIXA' !", vbCritical, "Erro"
        GoTo Erro
    End If
    
    CaixaAberturaAberto = AbreTabela(Dicion�rio, "FINANCEIRO", "CAIXA - ABERTURA", DBFinanceiro, TBLCaixaAbertura, TBLTabela, dbOpenTable)
    
    If CaixaAberturaAberto Then
        IndiceCaixaAberturaAtivo = "CAIXAABERTURA1"
        TBLCaixaAbertura.Index = IndiceCaixaAberturaAtivo
    Else
        MsgBox "N�o consegui abrir a tabela 'Caixa - Abertura' !", vbCritical, "Erro"
        GoTo Erro
    End If
    
    CaixaMovimentoAberto = AbreTabela(Dicion�rio, "FINANCEIRO", "CAIXA - MOVIMENTO", DBFinanceiro, TBLCaixaMovimento, TBLTabela, dbOpenTable)
    
    If CaixaMovimentoAberto Then
        IndiceCaixaMovimentoAtivo = "CAIXAMOVIMENTO1"
        TBLCaixaMovimento.Index = IndiceCaixaMovimentoAtivo
    Else
        MsgBox "N�o consegui abrir a tabela 'CAIXA - MOVIMENTO' !", vbCritical, "Erro"
        GoTo Erro
    End If
    
    CaixaSangriaEntradaAberto = AbreTabela(Dicion�rio, "FINANCEIRO", "CAIXA - SANGRIA - ENTRADA", DBFinanceiro, TBLCaixaSangriaEntrada, TBLTabela, dbOpenTable)
    
    If CaixaSangriaEntradaAberto Then
        IndiceCaixaSangriaEntradaAtivo = "CAIXASANGRIAENTRADA1"
        TBLCaixaSangriaEntrada.Index = IndiceCaixaSangriaEntradaAtivo
    Else
        MsgBox "N�o consegui abrir a tabela 'Caixa - Sangria - Entrada' !", vbCritical, "Erro"
        GoTo Erro
    End If
    
    Par�metrosAberto = AbreTabela(Dicion�rio, "SISTEMA", "PAR�METROS", DBSistema, TBLPar�metros, TBLTabela, dbOpenTable)
    
    If Par�metrosAberto Then
    Else
        MsgBox "N�o consegui abrir a tabela 'Par�metros' !", vbCritical, "Erro"
        GoTo Erro
    End If
    
    If Not Inicializa��o Then
        lFechar = True
        Exit Sub
    End If
    
    If Not AbrirPorta(lAbriPorta) Then
        lFechar = True
        GoTo Erro
    End If
        
    Exit Sub
    
Erro:
    GeraMensagemDeErro "CaixaF�cil - Load"
    lFechar = True
End Sub
Private Sub Form_Unload(Cancel As Integer)
    If Configura��oCaixaAberto Then
        TBLConfigura��oCaixa.Close
    End If
    If VendasAberto Then
        TBLVendas.Close
    End If
    If VendasItensAberto Then
        TBLVendasItens.Close
    End If
    If VendasLoteAberto Then
        TBLVendasLote.Close
    End If
    If PlanoDePagamentoAberto Then
        TBLPlanoDePagamento.Close
    End If
    If FormaDePagamentoAberto Then
        TBLFormaDePagamento.Close
    End If
    If CaixaAberto Then
        TBLCaixa.Close
    End If
    If CaixaAberturaAberto Then
        TBLCaixaAbertura.Close
    End If
    If CaixaMovimentoAberto Then
        TBLCaixaMovimento.Close
    End If
    If CaixaSangriaEntradaAberto Then
        TBLCaixaSangriaEntrada.Close
    End If
    If Par�metrosAberto Then
        TBLPar�metros.Close
    End If
    If Forms.Count = 2 Then
        AllBot�es False
    End If
    
    If lAbriPorta Then
        FecharPorta
    End If
    
    Set frmCaixa = Nothing
End Sub
Private Sub lstCupom_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim xPosi��o As Integer
    Dim Usu�rio As String
    Dim lAllowCancel As Boolean
            
    
    If KeyCode = 46 Then
        If Right(lstCupom.List(lstCupom.ListIndex), 1) <> "]" Then
            MsgBox "Nenhum item selecionado!", vbInformation, "Aviso"
        Else
            frmValidaUsu�rio.Show 1
            
            Usu�rio = frmValidaUsu�rio.Usu�rio
            
            Set frmValidaUsu�rio = Nothing
            
            If Usu�rio = Empty Then
                MsgBox "Nenhum usu�rio foi selecionado!", vbInformation, "Aviso"
                Exit Sub
            End If
    
            lAllowCancel = Allow("CAIXA", "C", Usu�rio)
            
            If Not lAllowCancel Then
                MsgBox "Acesso negado!", vbCritical, "Aviso"
                Exit Sub
            End If
            
            xPosi��o = Int(lstCupom.ListIndex / 3)
            ItensArray(4, xPosi��o) = "C"
            lstCupom.List(lstCupom.ListIndex) = lstCupom.List(lstCupom.ListIndex) & " *"
            lstCupom.RemoveItem lstCupom.ListCount - 1
            lstCupom.RemoveItem lstCupom.ListCount - 1
            lstCupom.RemoveItem lstCupom.ListCount - 1
            FillSubTotal
        End If
    Else
        SetKey KeyCode, Shift
    End If
End Sub
Private Sub pctProdutoEmpresa_KeyDown(KeyCode As Integer, Shift As Integer)
    SetKey KeyCode, Shift
End Sub
Private Sub txtComando_Change()
    If lblMensagem = "Valor" And Not lPula Then
        FormatMask "@K 99.999.999,99", txtComando
    End If
End Sub
Private Sub txtComando_KeyDown(KeyCode As Integer, Shift As Integer)
    SetKey KeyCode, Shift
End Sub
Private Sub txtPre�oTotal_KeyDown(KeyCode As Integer, Shift As Integer)
    SetKey KeyCode, Shift
End Sub
Private Sub txtPre�oUnit�rio_KeyDown(KeyCode As Integer, Shift As Integer)
    SetKey KeyCode, Shift
End Sub
Private Sub txtQuantidade_KeyDown(KeyCode As Integer, Shift As Integer)
    SetKey KeyCode, Shift
End Sub
Private Sub txtTotalAPagar_KeyDown(KeyCode As Integer, Shift As Integer)
    SetKey KeyCode, Shift
End Sub
