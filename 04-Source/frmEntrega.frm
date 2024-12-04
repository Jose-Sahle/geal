VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form frmEntrega 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Controle de Entregas"
   ClientHeight    =   6570
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9540
   Icon            =   "frmEntrega.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6570
   ScaleWidth      =   9540
   Begin VB.CommandButton cmdGravar 
      Caption         =   "&Localizar"
      Height          =   345
      Left            =   6960
      TabIndex        =   23
      Top             =   6180
      Width           =   1245
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   345
      Left            =   8280
      TabIndex        =   22
      Top             =   6180
      Width           =   1245
   End
   Begin VB.CommandButton cmdImprimir 
      Caption         =   "&Imprimir"
      Height          =   345
      Left            =   30
      TabIndex        =   21
      Top             =   6195
      Width           =   1980
   End
   Begin VB.Frame frTotais 
      Caption         =   "Totais "
      Height          =   1695
      Left            =   0
      TabIndex        =   11
      Top             =   4410
      Width           =   9525
      Begin VB.TextBox txtPorcentagemBonus 
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
         Left            =   6750
         TabIndex        =   16
         Text            =   "  0,00"
         Top             =   930
         Width           =   855
      End
      Begin VB.TextBox txtValorBonus 
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
         Left            =   7740
         TabIndex        =   15
         Text            =   "         0,00"
         Top             =   930
         Width           =   1665
      End
      Begin VB.TextBox txtValorTotal 
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
         Left            =   6750
         TabIndex        =   14
         Text            =   "R$"
         Top             =   1320
         Width           =   2655
      End
      Begin VB.TextBox txtDesconto 
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
         Left            =   7740
         Locked          =   -1  'True
         TabIndex        =   13
         Top             =   540
         Width           =   1665
      End
      Begin VB.TextBox txtValor 
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
         Left            =   7740
         Locked          =   -1  'True
         TabIndex        =   12
         Top             =   150
         Width           =   1665
      End
      Begin VB.Label lblSubTotal 
         Caption         =   "Sub Total"
         Height          =   255
         Left            =   6930
         TabIndex        =   20
         Top             =   240
         Width           =   705
      End
      Begin VB.Label lblTotalGeral 
         Caption         =   "Total do Or�amento"
         Height          =   225
         Left            =   5280
         TabIndex        =   19
         Top             =   1350
         Width           =   1425
      End
      Begin VB.Label lblDesconto 
         Caption         =   "Desconto"
         Height          =   255
         Left            =   6930
         TabIndex        =   18
         Top             =   630
         Width           =   1065
      End
      Begin VB.Label lblBonus 
         Caption         =   "Bonus"
         Height          =   195
         Left            =   6180
         TabIndex        =   17
         Top             =   990
         Width           =   495
      End
   End
   Begin VB.Frame frItens 
      Caption         =   " Itens "
      Height          =   3255
      Left            =   0
      TabIndex        =   9
      Top             =   1140
      Width           =   9540
      Begin MSDBGrid.DBGrid dbgrdItens 
         Height          =   2925
         Left            =   60
         OleObjectBlob   =   "frmEntrega.frx":030A
         TabIndex        =   10
         Top             =   210
         Width           =   9405
      End
   End
   Begin VB.Frame frDadosCadastrais 
      Height          =   1140
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9540
      Begin VB.TextBox txtDataDaEntrega 
         Height          =   285
         Left            =   8430
         TabIndex        =   24
         Text            =   "  /  /"
         Top             =   570
         Width           =   990
      End
      Begin VB.TextBox txtOr�amento 
         Height          =   285
         Left            =   1200
         TabIndex        =   4
         Top             =   300
         Width           =   765
      End
      Begin VB.TextBox txtDataOr�amento 
         Height          =   285
         Left            =   7290
         TabIndex        =   3
         Text            =   "  /  /"
         Top             =   240
         Width           =   990
      End
      Begin VB.TextBox txtCliente 
         Height          =   300
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   690
         Width           =   5085
      End
      Begin VB.TextBox txtDataVenda 
         Height          =   285
         Left            =   7290
         TabIndex        =   1
         Text            =   "  /  /"
         Top             =   690
         Width           =   990
      End
      Begin VB.Label lblDataDaEntrega 
         Caption         =   "Data da Entrega"
         Height          =   390
         Left            =   8610
         TabIndex        =   25
         Top             =   150
         Width           =   675
      End
      Begin VB.Label lblOr�amento 
         Caption         =   "Or�amento"
         Height          =   180
         Left            =   150
         TabIndex        =   8
         Top             =   330
         Width           =   825
      End
      Begin VB.Label lblDataOr�amento 
         Caption         =   "Data do Or�amento"
         Height          =   390
         Left            =   6390
         TabIndex        =   7
         Top             =   180
         Width           =   795
      End
      Begin VB.Label lblCliente 
         Caption         =   "Cliente"
         Height          =   180
         Left            =   150
         TabIndex        =   6
         Top             =   720
         Width           =   645
      End
      Begin VB.Label lblDataVenda 
         Caption         =   "Data da Venda"
         Height          =   405
         Left            =   6420
         TabIndex        =   5
         Top             =   630
         Width           =   765
      End
   End
End
Attribute VB_Name = "frmEntrega"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const MAXCOLS = 6

Dim mRecno%
Dim mTotalRows%
Dim dbgrdItensArray() As String

Dim lPula As Boolean
Dim mDigitBonus As Boolean
Dim mDigitPorcent As Boolean
Dim mlRefazDesconto As Boolean

Dim lFechar As Boolean

Dim mC�digo As Integer
Dim mOldValue As String


Dim mC�digoDoCliente As String

Dim TBLVendas As Table
Dim VendasAberto As Boolean
Dim IndiceVendasAtivo$

Dim TBLVendasItens As Table
Dim VendasItensAberto As Boolean
Dim IndiceVendasItensAtivo$

Dim TBLVendasLote As Table
Dim VendasLoteAberto As Boolean
Dim IndiceVendasLotesAtivo As String

Dim TBLEntregas As Table
Dim EntregasAberto As Boolean
Dim IndiceEntregasAtivo As String
Dim ArrayVendasItensRecno() As Variant

Dim TBLCliente        As Table
Dim ClienteAberto     As Boolean

Dim TBLPar�metros As Table
Dim Par�metrosAberto As Boolean

Dim TBLProduto    As Table
Dim ProdutoAberto As Boolean

Dim TBLC�digoDoProduto As Table
Dim C�digoDoProdutoAberto As Boolean

Dim TBLTipoDeEmbalagem As Table
Dim TipoDeEmbalagemAberto As Boolean

Dim StatusBarAviso$

Public lAtualizar As Boolean
Private Sub AtivaCampos()
    frDadosCadastrais.Enabled = True
    frItens.Enabled = True
    frTotais.Enabled = True
    cmdImprimir.Enabled = True
End Sub
Private Sub Cabe�alhoDaNotaFiscal()
    Printer.FontBold = True
    
    Printer.CurrentY = 2
    Printer.CurrentX = 13.2
    Printer.Print "X"
    
    Printer.CurrentY = 3.2
    Printer.CurrentX = 0.05
    Printer.Print TBLCliente("NOME - RAZ�O SOCIAL")

    Printer.CurrentY = 3.2
    Printer.CurrentX = 13
    Printer.Print TBLCliente("CGC - CPF")
    
    Printer.CurrentY = 3.2
    Printer.CurrentX = 18
    Printer.Print Date
    
    Printer.CurrentY = 4
    Printer.CurrentX = 0.05
    Printer.Print TBLCliente("ENDERE�O")
    
    Printer.CurrentY = 4
    Printer.CurrentX = 11.1
    Printer.Print TBLCliente("BAIRRO")

    Printer.CurrentY = 4
    Printer.CurrentX = 15.3
    Printer.Print TBLCliente("CEP")

    Printer.CurrentY = 4
    Printer.CurrentX = 18
    Printer.Print txtDataDaEntrega
    
    Printer.CurrentY = 4.5
    Printer.CurrentX = 0.05
    Printer.Print TBLCliente("CIDADE")
    
    Printer.CurrentY = 4.5
    Printer.CurrentX = 8.3
    If "(" & TBLCliente("DDD") & ")" & TBLCliente("FONE (1)") <> "()" Then
        Printer.Print "(" & TBLCliente("DDD") & ")" & TBLCliente("FONE (1)")
    End If

    Printer.CurrentY = 4.5
    Printer.CurrentX = 12
    Printer.Print TBLCliente("UF")

    Printer.CurrentY = 4.5
    Printer.CurrentX = 18
    Printer.Print Time
    
    Printer.FontBold = False
End Sub
Private Sub ConfigurarImpressora()
    On Error Resume Next
    
    Printer.TrackDefault = False
    Printer.ScaleMode = vbCentimeters
    Printer.FontBold = False
    Printer.FontItalic = False
    Printer.FontSize = 8
End Sub
Private Sub DesativaCampos()
    txtDataOr�amento.Enabled = False
    txtDataVenda.Enabled = False
    frItens.Enabled = False
    frTotais.Enabled = False
    cmdImprimir.Enabled = False
End Sub
Private Sub DetalheDaNotaFiscal()
    Dim LinhaCorrente As Single
    Dim TotalDeLinhas As Byte
    
    LinhaCorrente = 6
    TBLVendasItens.Seek "=", txtOr�amento
    
    Do While TBLVendasItens("OR�AMENTO") = txtOr�amento
        TBLC�digoDoProduto.Seek "=", TBLVendasItens("C�DIGO DO PRODUTO"), TBLPar�metros("CGC")
        Printer.CurrentY = LinhaCorrente
        Printer.CurrentX = 0.05
        Printer.Print TBLC�digoDoProduto("C�DIGO DO FORNECEDOR")
                
        TBLProduto.Seek "=", TBLVendasItens("C�DIGO DO PRODUTO")
        Printer.CurrentY = LinhaCorrente
        Printer.CurrentX = 2
        Printer.Print TBLProduto("DESCRI��O")
            
        TBLTipoDeEmbalagem.Seek "=", TBLProduto("TIPO DE EMBALAGEM")
        Printer.CurrentY = LinhaCorrente
        Printer.CurrentX = 11.7
        Printer.Print TBLTipoDeEmbalagem("ABREVIADO")
        
        Printer.CurrentY = LinhaCorrente
        Printer.CurrentX = 12.7
        Printer.Print FormatStringMask("@V ######0,00", StrVal(TBLVendasItens("QUANTIDADE")))
        
        Printer.CurrentY = LinhaCorrente
        Printer.CurrentX = 14.5
        Printer.Print FormatStringMask("@V ##.###.##0,00", StrVal(TBLVendasItens("VALOR UNIT�RIO")))
        
        Printer.CurrentY = LinhaCorrente
        Printer.CurrentX = 17.8
        Printer.Print FormatStringMask("@V ##.###.##0,00", StrVal(TBLVendasItens("QUANTIDADE") * TBLVendasItens("VALOR UNIT�RIO")))
        
        LinhaCorrente = LinhaCorrente + 0.35
        TotalDeLinhas = TotalDeLinhas + 1
        
        TBLVendasItens.MoveNext
        
        If TBLVendasItens.EOF Then
            Exit Do
        End If
        
        If TotalDeLinhas = 37 Then
            Printer.CurrentY = 22
            Printer.CurrentX = 0.35
            Printer.Print "Or�amento: " & txtOr�amento
            
            Printer.CurrentY = 23
            Printer.CurrentX = 0.35
            Printer.Print "Continua na pr�xima folha..."
            
            ConfigurarImpressora
            Cabe�alhoDaNotaFiscal
            TotalDeLinhas = 1
            LinhaCorrente = 6
        End If
    Loop
End Sub
Private Function Imprimir() As Boolean
    Dim lImpressao As Boolean
    
    frmImprimirEntrega.Show vbModal
    
    On Error Resume Next
    
    Set Printer = frmImprimirEntrega.Impressora
    
    If Err.Number <> 0 Then
        MsgBox "Nenhuma impressora foi selecionada!", vbInformation, "Aviso"
        Exit Function
    End If
    
    On Error GoTo Erro
    
    ConfigurarImpressora
    
    lImpressao = True
    
    Cabe�alhoDaNotaFiscal
    
    DetalheDaNotaFiscal
    
    RodaPeDaNotaFiscal
        
        
    Imprimir = True
    
    Printer.TrackDefault = True
    
    Exit Function
    
Erro:
    GeraMensagemDeErro "Entrega - Imprimir"
    If lImpressao Then
        Printer.KillDoc
    End If
    Printer.TrackDefault = True
End Function
Private Sub FillGrid(ByVal Chave As Long)
    dbgrdItens.ReBind
    
    ReDim dbgrdItensArray(MAXCOLS - 1, 0)
    ReDim ArrayVendasItensRecno(0)
    
    mTotalRows = 0
    mRecno = 0
    
    TBLVendasItens.Seek "=", Chave
    If Not TBLVendasItens.NoMatch Then
        Do While Not TBLVendasItens.EOF And TBLVendasItens("OR�AMENTO") = Chave
            mRecno = mRecno + 1
            mTotalRows = mTotalRows + 1
            ReDim Preserve dbgrdItensArray(MAXCOLS - 1, mTotalRows - 1)
            ReDim Preserve ArrayVendasItensRecno(mTotalRows - 1)
            
            ArrayVendasItensRecno(mTotalRows - 1) = TBLVendasItens.Bookmark
            dbgrdItensArray(0, mTotalRows - 1) = SearchProduto(TBLVendasItens("C�DIGO DO PRODUTO")) 'Nome do Produto
            dbgrdItensArray(1, mTotalRows - 1) = FormatStringMask("@V ######0", StrVal(TBLVendasItens("QUANTIDADE"))) 'Quantidade
            dbgrdItensArray(2, mTotalRows - 1) = FormatStringMask("@V ##.###.##0,00", StrVal(TBLVendasItens("VALOR UNIT�RIO"))) 'Pre�o Unit�rio
            dbgrdItensArray(3, mTotalRows - 1) = FormatStringMask("@V ##.###.##0,00", StrVal(TBLVendasItens("DESCONTO"))) 'Desconto no valor do produto
            dbgrdItensArray(4, mTotalRows - 1) = FormatStringMask("@V ##.###.##0,00", StrVal((TBLVendasItens("VALOR UNIT�RIO") * TBLVendasItens("QUANTIDADE")))) 'Pre�o de Venda
            dbgrdItensArray(5, mTotalRows - 1) = TBLVendasItens("C�DIGO DO PRODUTO") 'C�digo do Produto

            TBLVendasItens.MoveNext

            If TBLVendasItens.EOF Then
                Exit Do
            End If
        Loop
    End If
    
    dbgrdItens.Refresh
End Sub
Public Sub Gravar()
    If SetRecords Then
        cmdGravar.Caption = "&Localizar"
        ZeraCampos
        DesativaCampos
        StatusBarAviso = "Venda aceita"
    Else
        StatusBarAviso = "Ocorreu uma falha"
    End If
    
    BarraDeStatus StatusBarAviso
    
    If txtOr�amento.Enabled Then
        txtOr�amento.SetFocus
    End If
End Sub
Private Function Localizar() As Boolean
    If PosRecords Then
        GetRecords
        cmdImprimir.Enabled = True
        cmdCancelar.Enabled = True
        cmdGravar.Enabled = False
        Localizar = True
    Else
        Localizar = False
    End If
End Function
Private Function PosRecords() As Boolean
    TBLVendas.Seek "=", Val(txtOr�amento)
    If TBLVendas.NoMatch Then
        PosRecords = False
        MsgBox "N�o encontrei o or�amento " & txtOr�amento, vbInformation, "Aviso"
    Else
        If TBLVendas("ENTREGAS") Then
            MsgBox "O formul�rio de Entrega j� foi emitido!", vbInformation, "Aviso"
            PosRecords = False
        Else
            PosRecords = True
        End If
    End If
End Function
Private Sub GetRecords()
    On Error Resume Next
    
    Dim pValor1 As Currency
    Dim pValor2 As Currency
    Dim pValor3 As Currency
    Dim pValor4 As Currency
    
    lPula = True
    txtOr�amento = TBLVendas("C�DIGO")
    mC�digoDoCliente = TBLVendas("C�DIGO DO CLIENTE")
    txtCliente = SearchCliente(mC�digoDoCliente, byCodigo)
    txtValor = FormatStringMask("@V ##.###.##0,00", StrVal(TBLVendas("VALOR TOTAL DA VENDA")))
    txtDesconto = FormatStringMask("@V ##.###.##0,00", StrVal(TBLVendas("DESCONTO TOTAL DA VENDA")))
    txtValorBonus = FormatStringMask("@V ##.###.##0,00", StrVal(TBLVendas("VALOR DO BONUS")))
    txtDataVenda = TBLVendas("DATA DA VENDA")
    CorrigeData DataMask, txtDataVenda, TBLVendas("DATA DA VENDA")
    txtDataOr�amento = TBLVendas("DATA DO OR�AMENTO")
    CorrigeData DataMask, txtDataOr�amento, TBLVendas("DATA DO OR�AMENTO")
    
    'Calcula valor total
    pValor1 = ValStr(txtValor)
    pValor2 = ValStr(txtDesconto)
    pValor3 = ValStr(txtValorBonus)
    pValor4 = pValor1 - pValor2 - pValor3
    
    txtValorTotal = "R$" + String(6, " ") + FormatStringMask("@V ##.###.##0,00", StrVal(pValor4))
    
    'Calcula porcentagem de bonus
    If (pValor1 - pValor2) = 0 Then
        pValor4 = 0
    Else
        pValor4 = pValor3 * 100 / (pValor1 - pValor2)
    End If
    txtPorcentagemBonus = FormatStringMask("@V ##0,00", StrVal(pValor4))
    
    FillGrid TBLVendas("C�DIGO")
    
    lPula = False
End Sub
Private Sub RodaPeDaNotaFiscal()
    Dim ValorDoDesconto As Currency
    Dim BaseDeC�lculo   As Currency
        
    Printer.FontSize = 8
    Printer.FontBold = False
    
    ValorDoDesconto = TBLVendas("DESCONTO TOTAL DA VENDA") + TBLVendas("VALOR DO BONUS")
    
    If ValorDoDesconto > 0 Then
        Printer.CurrentY = 18.5
        Printer.CurrentX = 2
        Printer.Print "DESCONTO"
        
        Printer.CurrentY = 19.5
        Printer.CurrentX = 17.8
        Printer.Print FormatStringMask("@V ##.###.##0,00", StrVal(ValorDoDesconto))
    End If
        
    Printer.CurrentY = 20.5
    Printer.CurrentX = 17.8
    Printer.Print FormatStringMask("@V ##.###.##0,00", StrVal(TBLVendas("VALOR TOTAL DA VENDA") - ValorDoDesconto))
    
    Printer.CurrentY = 22
    Printer.CurrentX = 0.35
    Printer.Print "Or�amento: " & txtOr�amento
    
    Printer.EndDoc
End Sub
Private Function SetRecords() As Boolean
    On Error GoTo Erro
    
    WS.BeginTrans
    
    TBLVendas.Seek "=", txtOr�amento
    
    TBLVendas.Edit
    TBLVendas("ENTREGAS") = True
    TBLVendas.Update
    
    TBLVendasItens.Seek "=", txtOr�amento
    Do While TBLVendasItens("OR�AMENTO") = txtOr�amento
        TBLEntregas.AddNew
        TBLEntregas("OR�AMENTO") = txtOr�amento
        TBLEntregas("C�DIGO DO PRODUTO") = TBLVendasItens("C�DIGO DO PRODUTO")
        TBLEntregas("DATA DE ENTREGA") = IIf(Trim(StrTran(txtDataDaEntrega, "/")) <> Empty, txtDataDaEntrega, vbNull)
    
        TBLEntregas("USERNAME - CRIA") = gUsu�rio
        TBLEntregas("DATA - CRIA") = Date
        TBLEntregas("HORA - CRIA") = Time
        TBLEntregas("USERNAME - ALTERA") = "VAZIO"
        TBLEntregas("DATA - ALTERA") = vbNull
        TBLEntregas("HORA - ALTERA") = vbNull
        
        TBLVendasItens.MoveNext
        
        If TBLVendasItens.EOF Then
            Exit Do
        End If
    Loop
    
    WS.CommitTrans
    
    SetRecords = True
    
    Exit Function
    
Erro:
    GeraMensagemDeErro "Entrega - SetRecords - Or�amento " & txtOr�amento, True
    SetRecords = False
End Function
Private Sub ZeraCampos()
    On Error Resume Next
    
    lPula = True
    txtOr�amento = Empty
    txtDataOr�amento = Empty
    txtDataVenda = Empty
    txtValor = FormatStringMask("@V ##.###.##0,00", "0,00")
    txtValorTotal = "R$" & String(6, " ") & FormatStringMask("@V ##.###.##0,00", "0,00")
    txtDesconto = FormatStringMask("@V ##.###.##0,00", "0,00")
    txtValorBonus = FormatStringMask("@V ##.###.##0,00", "0,00")
    txtPorcentagemBonus = FormatStringMask("@V ##0,00", "  0,00")
    txtCliente = Empty
    mC�digo = 0
    ReDim dbgrdItensArray(MAXCOLS - 1, 0)
    mTotalRows = 0
    mRecno = 0
    dbgrdItens.ReBind
    lPula = False
End Sub
Private Sub cmdCancelar_Click()
    ZeraCampos
    DesativaCampos
    cmdGravar.Enabled = True
    cmdImprimir.Enabled = False
    cmdCancelar.Enabled = False
    Bot�oGravar False
End Sub
Private Sub cmdGravar_Click()
    If Localizar Then
        cmdGravar.Enabled = False
    End If
End Sub
Private Sub cmdImprimir_Click()
    If Trim(StrTran(txtDataDaEntrega, "/")) = Empty Then
        MsgBox "A data de entrega n�o foi preenchida!", vbExclamation, "Aviso"
        Exit Sub
    End If
    
    TBLVendas.Seek "=", txtOr�amento
    If TBLVendas.NoMatch Then
        MsgBox "Or�amento " & txtOr�amento & " n�o foi encontrado!", vbCritical, "Aviso"
        Exit Sub
    End If
    
    TBLVendasItens.Seek "=", txtOr�amento
    If TBLVendasItens.NoMatch Then
        MsgBox "Nenhum item foi encontrado para o or�amento " & txtOr�amento & "!", vbCritical, "Aviso"
        Exit Sub
    End If
    
    TBLCliente.Seek "=", TBLVendas("C�DIGO DO CLIENTE")
    If TBLCliente.NoMatch Then
        MsgBox "Dados do cliente com 'C�DIGO' " & TBLVendas("C�DIGO DO CLIENTE") & " n�o foi encontrado!", vbCritical, "Aviso"
        Exit Sub
    End If
    
    If MsgBox("Ajuste o formul�rio, e clique Ok para iniciar a impress�o!" & vbCr & "Ou tecle 'Cancelar/Cancel' para desistir", vbInformation + vbOKCancel, "Confirma��o") = vbOK Then
        If Imprimir Then
            cmdImprimir.Enabled = False
            cmdCancelar.Enabled = False
            cmdGravar.Enabled = True
            If Not SetRecords Then
                MsgBox "Repita a opera��o!", vbCritical, "Erro"
            End If
        Else
            MsgBox "Falha na impress�o!", vbCritical, "Erro"
        End If
    End If
End Sub
Private Sub dbgrdItens_ColResize(ByVal ColIndex As Integer, Cancel As Integer)
    Cancel = 1
End Sub
Private Sub dbgrdItens_RowResize(Cancel As Integer)
    Cancel = 1
End Sub
Private Sub dbgrdItens_UnboundAddData(ByVal RowBuf As MSDBGrid.RowBuffer, NewRowBookmark As Variant)
    Dim Col%
        
    mTotalRows = mTotalRows + 1
    ReDim Preserve dbgrdItensArray(MAXCOLS - 1, mTotalRows - 1)
    
    'Sets the bookmark to the last row.
    NewRowBookmark = mTotalRows - 1
    
    ' The following loop adds a new record to the database.
    For Col = 0 To UBound(dbgrdItensArray, 1)
        If Not IsNull(RowBuf.Value(0, Col)) Then
            dbgrdItensArray(Col, mTotalRows - 1) = RowBuf.Value(0, Col)
        Else
            ' If no value set for column, then use the
            ' DefaultValue
            dbgrdItensArray(Col, mTotalRows - 1) = dbgrdItens.Columns(Col).DefaultValue
        End If
    Next
End Sub
Private Sub dbgrdItens_UnboundDeleteRow(Bookmark As Variant)
    Dim iCol As Integer, iRow As Integer
    
    ' Move all rows above the deleted row down in the
    ' array.
    
    For iRow = Bookmark + 1 To mTotalRows - 1
        For iCol = 0 To MAXCOLS - 1
            dbgrdItensArray(iCol, iRow - 1) = dbgrdItensArray(iCol, iRow)
        Next iCol
    Next iRow
    
    mTotalRows = mTotalRows - 1
End Sub
Private Sub dbgrdItens_UnboundReadData(ByVal RowBuf As MSDBGrid.RowBuffer, StartLocation As Variant, ByVal ReadPriorRows As Boolean)
    Dim CurRow&, iRow As Integer, iCol As Integer, iRowsFetched As Integer, iIncr As Integer
    ' DBGrid is requesting rows so give them to it
    
    If mTotalRows = 0 Then Exit Sub
    
    If ReadPriorRows Then
        iIncr = -1
    Else
        iIncr = 1
    End If
    
    ' If StartLocation is Null then start reading at the end
    ' or beginning of the data set.
    If IsNull(StartLocation) Then
        If ReadPriorRows Then
            CurRow = RowBuf.RowCount - 1
        Else
            CurRow = 0
        End If
    Else
        ' Find the position to start reading based on the
        ' StartLocation bookmark and the iIncr variable
        CurRow = CLng(StartLocation) + iIncr
    End If
    
    ' Transfer data from our data set array to the RowBuf object
    ' which DBGrid uses to display the data
    For iRow = 0 To RowBuf.RowCount - 1
        If CurRow < 0 Or CurRow >= mTotalRows Then Exit For
        For iCol = 0 To UBound(dbgrdItensArray, 1)
            RowBuf.Value(iRow, iCol) = dbgrdItensArray(iCol, CurRow&)
        Next iCol
        ' Set bookmark using CurRow& which is also our
        ' array index
        RowBuf.Bookmark(iRow) = CStr(CurRow)
        CurRow = CurRow + iIncr
        iRowsFetched = iRowsFetched + 1
    Next iRow
    RowBuf.RowCount = iRowsFetched
End Sub
Private Sub dbgrdItens_UnboundWriteData(ByVal RowBuf As MSDBGrid.RowBuffer, WriteLocation As Variant)
    Dim iCol As Integer
    ' Data is being updated
    'MsgBox WriteLocation
    ' Update each column in the data set array
    For iCol = 0 To MAXCOLS - 1
        If Not IsNull(RowBuf.Value(0, iCol)) Then
            dbgrdItensArray(iCol, WriteLocation) = RowBuf.Value(0, iCol)
        End If
    Next iCol
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
    If Not EntregasAberto Then
        Unload Me
        Exit Sub
    End If
    If Not ClienteAberto Then
        Unload Me
        Exit Sub
    End If
    
    If Not Par�metrosAberto Then
        Unload Me
        Exit Sub
    End If
    If Not ProdutoAberto Then
        Unload Me
        Exit Sub
    End If
    If Not C�digoDoProdutoAberto Then
        Unload Me
        Exit Sub
    End If
    If Not TipoDeEmbalagemAberto Then
        Unload Me
        Exit Sub
    End If
    
    Bot�oGravar False
    Bot�oGravar True
    Navega��oInferior False
    Navega��oSuperior False
    Bot�oExcluir False
    Bot�oIncluir False
    Bot�oImprimir False
    
    BarraDeStatus StatusBarAviso
    dbgrdItens.Refresh

    If lAtualizar Then
        Bot�oAtualizar True
    Else
        Bot�oAtualizar False
    End If
End Sub
Private Sub Form_Deactivate()
    cmdGravar.Enabled = False
    Bot�oImprimir False
End Sub
Private Sub Form_Load()
    On Error GoTo Erro
    
    Dim Cont%
    
    ZeraCampos
    
    VendasAberto = AbreTabela(Dicion�rio, "FINANCEIRO", "VENDA", DBFinanceiro, TBLVendas, TBLTabela, dbOpenTable)
    
    If VendasAberto Then
        IndiceVendasAtivo = "VENDA1"
        TBLVendas.Index = IndiceVendasAtivo
    Else
        MsgBox "N�o consegui abrir a tabela 'Vendas' !", vbCritical, "Erro"
        Exit Sub
    End If
    
    VendasItensAberto = AbreTabela(Dicion�rio, "FINANCEIRO", "VENDA - ITENS", DBFinanceiro, TBLVendasItens, TBLTabela, dbOpenTable)
    
    If VendasItensAberto Then
        IndiceVendasItensAtivo = "VENDAITENS1"
        TBLVendasItens.Index = IndiceVendasItensAtivo
    Else
        MsgBox "N�o consegui abrir a tabela 'Itens de Venda' !", vbCritical, "Erro"
        Exit Sub
    End If
    
    VendasLoteAberto = AbreTabela(Dicion�rio, "FINANCEIRO", "VENDA - LOTES", DBFinanceiro, TBLVendasLote, TBLTabela, dbOpenTable)
    
    If VendasLoteAberto Then
        IndiceVendasLotesAtivo = "VENDALOTES2"
        TBLVendasLote.Index = IndiceVendasLotesAtivo
    Else
        MsgBox "N�o consegui abrir a tabela 'Lote de Vendas' !", vbCritical, "Erro"
        GoTo Erro
    End If
    
    EntregasAberto = AbreTabela(Dicion�rio, "FINANCEIRO", "ENTREGAS", DBFinanceiro, TBLEntregas, TBLTabela, dbOpenTable)
    
    If EntregasAberto Then
        IndiceEntregasAtivo = "ENTREGAS1"
        TBLEntregas.Index = IndiceEntregasAtivo
    Else
        MsgBox "N�o consegui abrir a tabela 'Entregas' !", vbCritical, "Erro"
        GoTo Erro
    End If
    
    ClienteAberto = AbreTabela(Dicion�rio, "CADASTRO", "CLIENTE", DBCadastro, TBLCliente, TBLTabela, dbOpenTable)
    
    If ClienteAberto Then
        TBLCliente.Index = "CLIENTE1"
    Else
        MsgBox "N�o consegui abrir a tabela 'Cliente' !", vbCritical, "Erro"
        Exit Sub
    End If
    
    Par�metrosAberto = AbreTabela(Dicion�rio, "SISTEMA", "PAR�METROS", DBSistema, TBLPar�metros, TBLTabela, dbOpenTable)
    
    If Par�metrosAberto Then
    Else
        MsgBox "N�o consegui abrir a tabela 'Par�metros' !", vbCritical, "Erro"
        Exit Sub
    End If
    
    ProdutoAberto = AbreTabela(Dicion�rio, "CADASTRO", "PRODUTO", DBCadastro, TBLProduto, TBLTabela, dbOpenTable)
    
    If ProdutoAberto Then
        TBLProduto.Index = "PRODUTO1"
    Else
        MsgBox "N�o consegui abrir a tabela 'Produto' !", vbCritical, "Erro"
        Exit Sub
    End If
    
    C�digoDoProdutoAberto = AbreTabela(Dicion�rio, "CADASTRO", "C�DIGO DO PRODUTO", DBCadastro, TBLC�digoDoProduto, TBLTabela, dbOpenTable)
    
    If C�digoDoProdutoAberto Then
        TBLC�digoDoProduto.Index = "C�DIGODOPRODUTO4"
    Else
        MsgBox "N�o consegui abrir a tabela 'C�digo do Produto' !", vbCritical, "Erro"
        Exit Sub
    End If
    
    TipoDeEmbalagemAberto = AbreTabela(Dicion�rio, "CADASTRO", "TIPO DE EMBALAGEM", DBCadastro, TBLTipoDeEmbalagem, TBLTabela, dbOpenTable)
    
    If TipoDeEmbalagemAberto Then
        TBLTipoDeEmbalagem.Index = "TIPODEEMBALAGEM1"
    Else
        MsgBox "N�o consegui abrir a tabela 'Tipo De Embalagem' !", vbCritical, "Erro"
        Exit Sub
    End If
    
    dbgrdItens.Columns.Add 1
    dbgrdItens.Columns.Add 1
    dbgrdItens.Columns.Add 1
    dbgrdItens.Columns.Add 1
    
    For Cont = 0 To dbgrdItens.Columns.Count - 1
        dbgrdItens.Columns(Cont).Visible = True
    Next
       
    dbgrdItens.Columns(0).Caption = "Produto"
    dbgrdItens.Columns(0).Width = 3045
    dbgrdItens.Columns(0).DefaultValue = " "
    dbgrdItens.Columns(0).Alignment = dbgLeft
    
    dbgrdItens.Columns(1).Caption = "Quantidade"
    dbgrdItens.Columns(1).Width = 1000
    dbgrdItens.Columns(1).DefaultValue = "0"
    dbgrdItens.Columns(1).Alignment = dbgRight
    
    dbgrdItens.Columns(2).Caption = "Valor Unit�rio"
    dbgrdItens.Columns(2).Width = 1910
    dbgrdItens.Columns(2).DefaultValue = "0,00"
    dbgrdItens.Columns(2).Alignment = dbgRight
    
    dbgrdItens.Columns(3).Caption = "Desconto"
    dbgrdItens.Columns(3).Width = 1000
    dbgrdItens.Columns(3).DefaultValue = "0,00"
    dbgrdItens.Columns(3).Alignment = dbgRight
    
    dbgrdItens.Columns(4).Caption = "Valor Total"
    dbgrdItens.Columns(4).Width = 1910
    dbgrdItens.Columns(4).DefaultValue = "0,00"
    dbgrdItens.Columns(4).Alignment = dbgRight
    
    dbgrdItens.Columns(5).Caption = "" 'C�digo do Produto
    dbgrdItens.Columns(5).Width = 1
    dbgrdItens.Columns(5).DefaultValue = "0"
    
    dbgrdItens.ReBind
    
    Navega��oInferior False
        
    StatusBarAviso = "Pronto"
    
    DesativaCampos
    
    Exit Sub
    
Erro:
    GeraMensagemDeErro "Entrega - Load"
    lFechar = True
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    mdiGeal.StatusBar.Panels("Posi��o").Visible = False
    ResizeStatusBar
End Sub
Private Sub Form_Unload(Cancel As Integer)
    If VendasAberto Then
        TBLVendas.Close
    End If
    If VendasItensAberto Then
        TBLVendasItens.Close
    End If
    If VendasLoteAberto Then
        TBLVendasLote.Close
    End If
    If EntregasAberto Then
        TBLEntregas.Close
    End If
    If ClienteAberto Then
        TBLCliente.Close
    End If
    If Par�metrosAberto Then
        TBLPar�metros.Close
    End If
    If ProdutoAberto Then
        TBLProduto.Close
    End If
    If C�digoDoProdutoAberto Then
        TBLC�digoDoProduto.Close
    End If
    If TipoDeEmbalagemAberto Then
        TBLTipoDeEmbalagem.Close
    End If
    
    If Forms.Count = 2 Then
        AllBot�es False
    End If
End Sub
Private Sub txtDataDaEntrega_Change()
    If Not lPula Then
        lPula = True
        FormatMask DataMask, txtDataDaEntrega
        lPula = False
    End If
End Sub
Private Sub txtDataDaEntrega_LostFocus()
    If Trim(StrTran(txtDataDaEntrega.Text, "/")) <> Empty Then
        lPula = True
        CorrigeData DataMask, txtDataDaEntrega, Date
        lPula = False
        If Not FormatMask(CheckDataMask, txtDataDaEntrega) Then
            Beep
            MsgBox "Data inv�lida !", vbCritical, "Erro"
            txtDataDaEntrega.SelStart = 0
            txtDataDaEntrega.SetFocus
        End If
    End If
End Sub
Private Sub txtOr�amento_Change()
    If txtOr�amento <> Empty Then
        cmdGravar.Enabled = True
    End If
End Sub

