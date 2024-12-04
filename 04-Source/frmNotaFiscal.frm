VERSION 5.00
Begin VB.Form frmNotaFiscal 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Nota Fiscal"
   ClientHeight    =   2205
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6150
   Icon            =   "frmNotaFiscal.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2205
   ScaleWidth      =   6150
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   345
      Left            =   4890
      TabIndex        =   9
      Top             =   1860
      Width           =   1245
   End
   Begin VB.CommandButton cmdEmitir 
      Caption         =   "&Emitir"
      Height          =   345
      Left            =   3600
      TabIndex        =   8
      Top             =   1860
      Width           =   1245
   End
   Begin VB.Frame frDadosDaNotaFiscal 
      Caption         =   "Dados da Nota Fiscal"
      Height          =   1185
      Left            =   0
      TabIndex        =   2
      Top             =   630
      Width           =   6135
      Begin VB.TextBox txtConfirmaNota 
         Height          =   315
         Left            =   2610
         TabIndex        =   5
         Top             =   240
         Width           =   1005
      End
      Begin VB.ComboBox cmbCFOP 
         Height          =   315
         ItemData        =   "frmNotaFiscal.frx":030A
         Left            =   690
         List            =   "frmNotaFiscal.frx":0327
         TabIndex        =   7
         Top             =   660
         Width           =   5325
      End
      Begin VB.TextBox txtNotaFiscal 
         Height          =   315
         Left            =   1500
         TabIndex        =   4
         Top             =   240
         Width           =   1005
      End
      Begin VB.Label lblCFOP 
         Caption         =   "CFOP"
         Height          =   225
         Left            =   180
         TabIndex        =   6
         Top             =   720
         Width           =   555
      End
      Begin VB.Label lblN�meroDaNotaFiscal 
         Caption         =   "N�mero da Nota Fiscal"
         Height          =   225
         Left            =   210
         TabIndex        =   3
         Top             =   270
         Width           =   1215
      End
   End
   Begin VB.Frame frImpressora 
      Caption         =   "Impressora"
      Height          =   615
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6135
      Begin VB.ComboBox cmbImpressora 
         Height          =   315
         Left            =   150
         TabIndex        =   1
         Top             =   210
         Width           =   5865
      End
   End
End
Attribute VB_Name = "frmNotaFiscal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim AllImpressoras() As Printer

Dim mlFechar          As Boolean

Dim mBaseDeCalculoICM As Currency
Dim mValorDoICM       As Currency
Dim mValorDaNota      As Currency

Dim TBLCliente        As Table
Dim ClienteAberto     As Boolean

Dim TBLVenda          As Table
Dim VendasAberto      As Boolean

Dim TBLVendaItens     As Table
Dim VendasItensAberto As Boolean

Dim TBLPar�metros As Table
Dim Par�metrosAberto As Boolean

Dim TBLProduto    As Table
Dim ProdutoAberto As Boolean

Dim TBLC�digoDoProduto As Table
Dim C�digoDoProdutoAberto As Boolean

Dim TBLTipoDeEmbalagem As Table
Dim TipoDeEmbalagemAberto As Boolean

Public mOr�amento     As Long
Public mNF            As String
Public mTotalDeNotas  As Byte
Private Sub Cabe�alhoDaNotaFiscal()
    Printer.FontBold = True
    
    Printer.CurrentY = 0.75
    Printer.CurrentX = 13.36
    Printer.Print "X"
    
    Printer.CurrentY = 3.1
    Printer.CurrentX = 0.05
    Printer.Print Trim(GetWordSeparatedBy(cmbCFOP.Text, 2, "-"))
    
    Printer.CurrentY = 3.1
    Printer.CurrentX = 6.5
    Printer.Print Trim(GetWordSeparatedBy(cmbCFOP.Text, 1, "-"))
    
    Printer.CurrentY = 4
    Printer.CurrentX = 0.05
    Printer.Print TBLCliente("NOME - RAZ�O SOCIAL")

    Printer.CurrentY = 4
    Printer.CurrentX = 13
    Printer.Print TBLCliente("CGC - CPF")
    
    Printer.CurrentY = 4
    Printer.CurrentX = 18
    Printer.Print Date
    
    Printer.CurrentY = 4.7
    Printer.CurrentX = 0.05
    Printer.Print TBLCliente("ENDERE�O")
    
    Printer.CurrentY = 4.7
    Printer.CurrentX = 11.1
    Printer.Print TBLCliente("BAIRRO")

    Printer.CurrentY = 4.7
    Printer.CurrentX = 15.3
    Printer.Print TBLCliente("CEP")

    Printer.CurrentY = 4.7
    Printer.CurrentX = 18
    Printer.Print Date
    
    Printer.CurrentY = 5.3
    Printer.CurrentX = 0.05
    Printer.Print TBLCliente("CIDADE")
    
    Printer.CurrentY = 5.3
    Printer.CurrentX = 8.3
    If "(" & TBLCliente("DDD") & ")" & TBLCliente("FONE (1)") <> "()" Then
        Printer.Print "(" & TBLCliente("DDD") & ")" & TBLCliente("FONE (1)")
    End If

    Printer.CurrentY = 5.3
    Printer.CurrentX = 12
    Printer.Print TBLCliente("UF")

    Printer.CurrentY = 5.3
    Printer.CurrentX = 18
    Printer.Print Time
    
    Printer.FontBold = False
End Sub
Private Sub ConfigurarImpressora()
    On Error Resume Next
    
    Printer.TrackDefault = False
    Set Printer = AllImpressoras((cmbImpressora.ListIndex + 1))
    Printer.ScaleMode = vbCentimeters
    Printer.FontBold = False
    Printer.FontItalic = False
    Printer.FontSize = 8
End Sub
Private Sub DetalheDaNotaFiscal()
    Dim LinhaCorrente As Single
    Dim TotalDeLinhas As Byte
    
    mBaseDeCalculoICM = 0
    mValorDoICM = 0
    mValorDaNota = 0
    TotalDeLinhas = 1
    LinhaCorrente = 8.2
    mTotalDeNotas = 1
    TBLVendaItens.Seek "=", mOr�amento
    Do While TBLVendaItens("OR�AMENTO") = mOr�amento
        TBLC�digoDoProduto.Seek "=", TBLVendaItens("C�DIGO DO PRODUTO"), TBLPar�metros("CGC")
        Printer.CurrentY = LinhaCorrente
        Printer.CurrentX = 0.05
        Printer.Print TBLC�digoDoProduto("C�DIGO DO FORNECEDOR")
                
        TBLProduto.Seek "=", TBLVendaItens("C�DIGO DO PRODUTO")
        Printer.CurrentY = LinhaCorrente
        Printer.CurrentX = 2
        Printer.Print TBLProduto("DESCRI��O")
            
        TBLTipoDeEmbalagem.Seek "=", TBLProduto("TIPO DE EMBALAGEM")
        Printer.CurrentY = LinhaCorrente
        Printer.CurrentX = 11.7
        Printer.Print TBLTipoDeEmbalagem("ABREVIADO")
        
        Printer.CurrentY = LinhaCorrente
        Printer.CurrentX = 12.7
        Printer.Print FormatStringMask("@V ######0,00", StrVal(TBLVendaItens("QUANTIDADE")))
        
        Printer.CurrentY = LinhaCorrente
        Printer.CurrentX = 14.1
        Printer.Print FormatStringMask("@V ##.###.##0,00", StrVal(TBLVendaItens("VALOR UNIT�RIO")))
        
        Printer.CurrentY = LinhaCorrente
        Printer.CurrentX = 16.8
        Printer.Print FormatStringMask("@V ##.###.##0,00", StrVal(TBLVendaItens("QUANTIDADE") * TBLVendaItens("VALOR UNIT�RIO")))
        
'        Printer.CurrentY = LinhaCorrente
'        Printer.CurrentX = 19.5
'        Printer.Print TBLProduto("ICM")
        
        LinhaCorrente = LinhaCorrente + 0.35
        TotalDeLinhas = TotalDeLinhas + 1
        
        mValorDaNota = mValorDaNota + TBLVendaItens("QUANTIDADE") * TBLVendaItens("VALOR UNIT�RIO")
        
        If TBLProduto("ICM") > 0 Then
            mBaseDeCalculoICM = mBaseDeCalculoICM + TBLVendaItens("QUANTIDADE") * TBLVendaItens("VALOR UNIT�RIO")
            mValorDoICM = mValorDoICM + ((TBLVendaItens("QUANTIDADE") * TBLVendaItens("VALOR UNIT�RIO")) * (TBLProduto("ICM") / 100))
        End If
        
        If TotalDeLinhas = 37 Then
            mTotalDeNotas = mTotalDeNotas + 1
            RodaPeDaNotaFiscal
            MsgBox "Aguarde o fim de impress�o desta Nota Fiscal, ent�o, ajuste a folha da pr�xima!" & vbCr & vbCr & vbTab & vbTab & "Clique OK quando pronto", vbInformation, "Aviso"
            ConfigurarImpressora
            Cabe�alhoDaNotaFiscal
            mBaseDeCalculoICM = 0
            mValorDaNota = 0
            mValorDoICM = 0
            TotalDeLinhas = 1
            LinhaCorrente = 8.1
        End If
        
        TBLVendaItens.MoveNext
        
        If TBLVendaItens.EOF Then
            Exit Do
        End If
    Loop
End Sub
Private Sub RodaPeDaNotaFiscal()
    Dim ValorDoDesconto As Currency
    Dim BaseDeC�lculo   As Currency
        
    Printer.FontSize = 8
    Printer.FontBold = False
    
    If mTotalDeNotas = 1 Then
        ValorDoDesconto = TBLVenda("DESCONTO TOTAL DA VENDA") + TBLVenda("VALOR DO BONUS")
    Else
        ValorDoDesconto = TBLVenda("DESCONTO TOTAL DA VENDA") + TBLVenda("VALOR DO BONUS")
        ValorDoDesconto = ValorDoDesconto * mValorDaNota / TBLVenda("VALOR TOTAL DA VENDA")
    End If
    
'    Printer.CurrentY = 23
'    Printer.CurrentX = 0.05
'    Printer.Print FormatStringMask("@V ##.###.##0,00", StrVal(mBaseDeCalculoICM))
'
'    Printer.CurrentY = 23
'    Printer.CurrentX = 4.2
'    Printer.Print FormatStringMask("@V ##.###.##0,00", StrVal(mValorDoICM))
    
    If ValorDoDesconto > 0 Then
        Printer.CurrentY = 21.5
        Printer.CurrentX = 2
        Printer.Print "DESCONTO"
        
        Printer.CurrentY = 21.5
        Printer.CurrentX = 16.8
        Printer.Print FormatStringMask("@V ##.###.##0,00", StrVal(ValorDoDesconto))
    End If
        
    Printer.CurrentY = 23.7
    Printer.CurrentX = 16.5
    Printer.Print FormatStringMask("@V ##.###.##0,00", StrVal(mValorDaNota - ValorDoDesconto))
    
    Printer.CurrentY = 26.5
    Printer.CurrentX = 0.65
    Printer.Print "OR�AMENTO: " & mOr�amento
    
    Printer.CurrentY = 26.85
    Printer.CurrentX = 0.65
    Printer.Print "Este documento n�o transfere cr�dito de ICMS"
    
    Printer.CurrentY = 27.2
    Printer.CurrentX = 0.65
    Printer.Print "Base de C�lculo: " & FormatStringMask("@V ##.###.##0,00", StrVal(mValorDaNota - ValorDoDesconto))
    
    Printer.CurrentY = 27.55
    Printer.CurrentX = 0.65
    Printer.Print "Al�quota: 2,4375%"
    
    Printer.CurrentY = 27.9
    Printer.CurrentX = 0.65
    Printer.Print "Valor do ICMS: " & FormatStringMask("@V ##.###.##0,00", StrVal((mValorDaNota - ValorDoDesconto) * 0.024375))
    
    Printer.EndDoc
End Sub
Private Sub cmdCancelar_Click()
    mNF = Empty
    Unload Me
End Sub
Private Sub cmdEmitir_Click()
    On Error GoTo Erro
    
    Dim lImpress�o As Boolean
    
    lImpress�o = False
    
    mTotalDeNotas = 0
    
    If Trim(txtNotaFiscal) = Empty Then
        MsgBox "Preencha o campo de Nota Fiscal", vbCritical, "Aviso"
        Exit Sub
    End If
    
    If UCase(Trim(txtNotaFiscal)) <> UCase(Trim(txtConfirmaNota)) Then
        MsgBox "A confirma��o da Nota Fiscal n�o confere!", vbInformation, "Aviso"
        Exit Sub
    End If
    
    TBLVenda.Seek "=", mOr�amento
    If TBLVenda.NoMatch Then
        MsgBox "Or�amento " & mOr�amento & " n�o foi encontrado!", vbCritical, "Aviso"
        Exit Sub
    End If
    
    TBLVendaItens.Seek "=", mOr�amento
    If TBLVendaItens.NoMatch Then
        MsgBox "Nenhum item foi encontrado para o or�amento " & mOr�amento & "!", vbCritical, "Aviso"
        Exit Sub
    End If
    
    TBLCliente.Seek "=", TBLVenda("C�DIGO DO CLIENTE")
    If TBLCliente.NoMatch Then
        MsgBox "Dados do cliente com 'C�DIGO' " & TBLVenda("C�DIGO DO CLIENTE") & " n�o foi encontrado!", vbCritical, "Aviso"
        Exit Sub
    End If
    
    'Configura��o da impressora
    ConfigurarImpressora
    
    lImpress�o = True
    
    Cabe�alhoDaNotaFiscal
    
    DetalheDaNotaFiscal
    
    RodaPeDaNotaFiscal
    
    mNF = txtNotaFiscal
    
    Printer.TrackDefault = True

    Unload Me
    
    Exit Sub
    
Erro:
    GeraMensagemDeErro "Emiss�o de Nota Fiscal - Emitir"
    If lImpress�o Then
        Printer.KillDoc
    End If
    Printer.TrackDefault = True
End Sub
Private Sub Form_Activate()
    If mlFechar Then
        Unload Me
        Exit Sub
    End If
    
    If Not ClienteAberto Then
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
    
    If Not Par�metrosAberto Then
        Unload Me
        Exit Sub
    End If
End Sub
Private Sub Form_Load()
    On Error GoTo Erro
    
    Dim Cont       As Byte
    Dim Impressora As Printer
    
    mlFechar = False
    
    VendasAberto = AbreTabela(Dicion�rio, "FINANCEIRO", "VENDA", DBFinanceiro, TBLVenda, TBLTabela, dbOpenTable)
    
    If VendasAberto Then
        TBLVenda.Index = "VENDA1"
    Else
        MsgBox "N�o consegui abrir a tabela 'Vendas' !", vbCritical, "Erro"
        Exit Sub
    End If
    
    VendasItensAberto = AbreTabela(Dicion�rio, "FINANCEIRO", "VENDA - ITENS", DBFinanceiro, TBLVendaItens, TBLTabela, dbOpenTable)
    
    If VendasItensAberto Then
        TBLVendaItens.Index = "VENDAITENS1"
    Else
        MsgBox "N�o consegui abrir a tabela 'Itens de Venda' !", vbCritical, "Erro"
        Exit Sub
    End If
    
    ClienteAberto = AbreTabela(Dicion�rio, "CADASTRO", "CLIENTE", DBCadastro, TBLCliente, TBLTabela, dbOpenTable)
    
    If ClienteAberto Then
        TBLCliente.Index = "CLIENTE1"
    Else
        MsgBox "N�o consegui abrir a tabela 'Cliente' !", vbCritical, "Erro"
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
    
    Par�metrosAberto = AbreTabela(Dicion�rio, "SISTEMA", "PAR�METROS", DBSistema, TBLPar�metros, TBLTabela, dbOpenTable)
    
    If Par�metrosAberto Then
    Else
        MsgBox "N�o consegui abrir a tabela 'Par�metros' !", vbCritical, "Erro"
        Exit Sub
    End If
    
    ReDim Preserve AllImpressoras(1 To Printers.Count)
    
    Cont = 0
    For Each Impressora In Printers
        Cont = Cont + 1
        cmbImpressora.AddItem Impressora.DeviceName
        Set AllImpressoras(Cont) = Impressora
    Next
    
    Exit Sub
    
Erro:
    GeraMensagemDeErro "Nota Fiscal - Load"
    mlFechar = True
End Sub
Private Sub Form_Unload(Cancel As Integer)
    If ClienteAberto Then
        TBLCliente.Close
    End If
    
    If VendasAberto Then
        TBLVenda.Close
    End If
    
    If VendasItensAberto Then
        TBLVendaItens.Close
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
    
    If Par�metrosAberto Then
        TBLPar�metros.Close
    End If
    
End Sub
