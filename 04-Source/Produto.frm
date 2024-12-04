VERSION 5.00
Begin VB.Form frmProduto 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Produto"
   ClientHeight    =   6240
   ClientLeft      =   1215
   ClientTop       =   810
   ClientWidth     =   8040
   Icon            =   "Produto.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6240
   ScaleWidth      =   8040
   Begin VB.CommandButton cmdGravar 
      Caption         =   "&Gravar"
      Default         =   -1  'True
      Height          =   345
      Left            =   5445
      TabIndex        =   20
      Top             =   5820
      Width           =   1245
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   345
      Left            =   6765
      TabIndex        =   21
      Top             =   5820
      Width           =   1245
   End
   Begin VB.Frame frVariados 
      Height          =   720
      Left            =   0
      TabIndex        =   45
      Top             =   5040
      Width           =   8025
      Begin VB.CommandButton cmdLotes 
         Caption         =   "&Lotes"
         Height          =   345
         Left            =   6615
         TabIndex        =   19
         Top             =   240
         Width           =   1245
      End
      Begin VB.CommandButton cmdPreços 
         Caption         =   "&Preços"
         Height          =   345
         Left            =   3420
         TabIndex        =   18
         Top             =   240
         Width           =   1245
      End
      Begin VB.CommandButton cmdCódigo 
         Caption         =   "Có&digos"
         Height          =   345
         Left            =   150
         TabIndex        =   17
         Top             =   240
         Width           =   1245
      End
   End
   Begin VB.Frame frDescontoPromoção 
      Caption         =   " Promoção/Desconto"
      Height          =   1230
      Left            =   0
      TabIndex        =   33
      Top             =   3810
      Width           =   8025
      Begin VB.VScrollBar vscrDescontoMáximo 
         Height          =   315
         Left            =   2715
         TabIndex        =   43
         TabStop         =   0   'False
         Top             =   675
         Width           =   210
      End
      Begin VB.TextBox txtDescontoMáximo 
         Height          =   285
         Left            =   1950
         TabIndex        =   16
         Text            =   " 0,00"
         Top             =   690
         Width           =   765
      End
      Begin VB.TextBox txtTérmino 
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
         Left            =   6540
         TabIndex        =   15
         Text            =   "  /  /"
         Top             =   270
         Width           =   1290
      End
      Begin VB.TextBox txtInício 
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
         Left            =   4080
         TabIndex        =   14
         Text            =   "  /  /"
         Top             =   270
         Width           =   1290
      End
      Begin VB.VScrollBar vscrDescontoDePromoção 
         Height          =   315
         Left            =   2715
         TabIndex        =   38
         TabStop         =   0   'False
         Top             =   285
         Width           =   210
      End
      Begin VB.TextBox txtDescontoDePromoção 
         Height          =   285
         Left            =   1950
         TabIndex        =   13
         Text            =   " 0,00"
         Top             =   300
         Width           =   765
      End
      Begin VB.Label lblDescontoMáximoPorcentagem 
         Caption         =   "%"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   2985
         TabIndex        =   44
         Top             =   675
         Width           =   225
      End
      Begin VB.Label lblDescontoMáximo 
         Caption         =   "Desconto Máximo"
         Height          =   195
         Left            =   150
         TabIndex        =   42
         Top             =   720
         Width           =   1380
      End
      Begin VB.Label lblTérmino 
         Caption         =   "Término"
         Height          =   195
         Left            =   5850
         TabIndex        =   41
         Top             =   330
         Width           =   570
      End
      Begin VB.Label lblInício 
         Caption         =   "Início"
         Height          =   225
         Left            =   3600
         TabIndex        =   40
         Top             =   330
         Width           =   495
      End
      Begin VB.Label lblDescontoDePromoçãoPorcentagem 
         Caption         =   "%"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   2970
         TabIndex        =   39
         Top             =   285
         Width           =   225
      End
      Begin VB.Label lblDescontoDePromoção 
         Caption         =   "Desconto de Promoção"
         Height          =   180
         Left            =   150
         TabIndex        =   37
         Top             =   330
         Width           =   1740
      End
   End
   Begin VB.Frame frImpostos 
      Caption         =   " Impostos "
      Height          =   1125
      Left            =   0
      TabIndex        =   28
      Top             =   2670
      Width           =   8025
      Begin VB.VScrollBar vscrIPI 
         Height          =   315
         Left            =   1965
         TabIndex        =   32
         TabStop         =   0   'False
         Top             =   680
         Width           =   210
      End
      Begin VB.TextBox txtIPI 
         Height          =   285
         Left            =   1200
         TabIndex        =   12
         Text            =   " 0,00"
         Top             =   690
         Width           =   765
      End
      Begin VB.ComboBox cmbTipoDeICM 
         Height          =   315
         Left            =   1200
         TabIndex        =   10
         Top             =   300
         Width           =   5205
      End
      Begin VB.TextBox txtICM 
         Height          =   285
         Left            =   6900
         Locked          =   -1  'True
         TabIndex        =   11
         Text            =   " 0,00"
         Top             =   300
         Width           =   765
      End
      Begin VB.Label lblIPI 
         Caption         =   "IPI"
         Height          =   225
         Left            =   150
         TabIndex        =   31
         Top             =   720
         Width           =   285
      End
      Begin VB.Label lblTipoDeICM 
         Caption         =   "Tipo de ICM"
         Height          =   240
         Left            =   150
         TabIndex        =   30
         Top             =   330
         Width           =   990
      End
      Begin VB.Label lblICM 
         Caption         =   "ICM"
         Height          =   165
         Left            =   6525
         TabIndex        =   29
         Top             =   360
         Width           =   270
      End
   End
   Begin VB.Frame frEspecificações 
      Caption         =   " Especificações "
      Height          =   2655
      Left            =   0
      TabIndex        =   22
      Top             =   0
      Width           =   8025
      Begin VB.CheckBox chkLotes 
         Caption         =   "Divisão por Lotes"
         Height          =   225
         Left            =   4290
         TabIndex        =   8
         Top             =   2280
         Width           =   1545
      End
      Begin VB.ComboBox cmbLocal 
         Height          =   315
         Left            =   1650
         TabIndex        =   5
         Top             =   1830
         Width           =   6225
      End
      Begin VB.ComboBox cmbTipoDeEmbalagem 
         Height          =   315
         Left            =   1650
         TabIndex        =   4
         Top             =   1470
         Width           =   6225
      End
      Begin VB.ComboBox cmbSeção 
         Height          =   315
         Left            =   1200
         TabIndex        =   3
         Top             =   1080
         Width           =   6675
      End
      Begin VB.ComboBox cmbDepartamento 
         Height          =   315
         Left            =   1200
         TabIndex        =   2
         Top             =   690
         Width           =   6675
      End
      Begin VB.TextBox txtQuantidadeDeLotes 
         Enabled         =   0   'False
         Height          =   285
         Left            =   7110
         TabIndex        =   9
         Top             =   2250
         Width           =   765
      End
      Begin VB.ComboBox cmbUnidades 
         Height          =   315
         Left            =   2190
         TabIndex        =   7
         Text            =   "Unid."
         Top             =   2250
         Width           =   1635
      End
      Begin VB.VScrollBar vscrQuantidade 
         Height          =   315
         Left            =   1950
         TabIndex        =   35
         TabStop         =   0   'False
         Top             =   2230
         Width           =   210
      End
      Begin VB.TextBox txtQuantidade 
         Height          =   285
         Left            =   1140
         TabIndex        =   6
         Top             =   2250
         Width           =   825
      End
      Begin VB.VScrollBar vscrPeso 
         Height          =   315
         Left            =   7635
         TabIndex        =   25
         TabStop         =   0   'False
         Top             =   285
         Width           =   210
      End
      Begin VB.TextBox txtPeso 
         Height          =   285
         Left            =   6870
         TabIndex        =   1
         Text            =   "0,000"
         Top             =   300
         Width           =   765
      End
      Begin VB.TextBox txtDescrição 
         Height          =   285
         Left            =   1200
         TabIndex        =   0
         Top             =   300
         Width           =   5085
      End
      Begin VB.Label lblLocal 
         Caption         =   "Local do Estoque"
         Height          =   285
         Left            =   150
         TabIndex        =   47
         Top             =   1890
         Width           =   1365
      End
      Begin VB.Label lblSeção 
         Caption         =   "Seção"
         Height          =   195
         Left            =   150
         TabIndex        =   46
         Top             =   1110
         Width           =   765
      End
      Begin VB.Label lblQuantidadeDeLotes 
         Caption         =   "Qtd. de Lotes"
         Height          =   285
         Left            =   6000
         TabIndex        =   36
         Top             =   2280
         Width           =   1005
      End
      Begin VB.Label lblQuantidade 
         Caption         =   "Quantidade"
         Height          =   225
         Left            =   150
         TabIndex        =   34
         Top             =   2280
         Width           =   975
      End
      Begin VB.Label lblTipoDeEmbalagem 
         Caption         =   "Tipo de Embalagem"
         Height          =   270
         Left            =   150
         TabIndex        =   27
         Top             =   1500
         Width           =   1470
      End
      Begin VB.Label lblDepartamentoSeção 
         Caption         =   "Departamento"
         Height          =   255
         Left            =   150
         TabIndex        =   26
         Top             =   720
         Width           =   1050
      End
      Begin VB.Label lblPeso 
         Caption         =   "Peso"
         Height          =   240
         Left            =   6450
         TabIndex        =   24
         Top             =   330
         Width           =   420
      End
      Begin VB.Label lblDescrição 
         Caption         =   "Descrição"
         Height          =   210
         Left            =   150
         TabIndex        =   23
         Top             =   330
         Width           =   990
      End
   End
End
Attribute VB_Name = "frmProduto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const MAXCOL = 4

Dim Código  As Long
Dim mCódigo As Long

Dim lAllowInsert     As Boolean
Dim lAllowEdit       As Boolean
Dim lAllowDelete     As Boolean
Dim lAllowConsult    As Boolean
Dim lAllowEditAmount As Boolean

Dim TBLProduto    As Table
Dim ProdutoAberto As Boolean
Dim IndiceProdutoAtivo$

Dim TBLCódigoDoProduto As Table
Dim CódigoDoProdutoAberto As Boolean
Dim IndiceCódigoDoProdutoAtivo$

Dim TBLPreço As Table
Dim PreçoAberto As Boolean
Dim IndicePreçoAtivo$

Dim TBLLote As Table
Dim LoteAberto As Boolean
Dim IndiceLoteAtivo$

Dim TBLDepartamento As Table
Dim DepartamentoAberto As Boolean
Dim IndiceDepartamentoAtivo$

Dim TBLSeção As Table
Dim SeçãoAberto As Boolean
Dim IndiceSeçãoAtivo$

Dim TBLDepartamentoSeção As Table
Dim DepartamentoSeçãoAberto As Boolean
Dim IndiceDepartamentoSeçãoAtivo$

Dim TBLTipoDeICM As Table
Dim TipoDeICMAberto As Boolean
Dim IndiceTipoDeICMAtivo$

Dim TBLTipoDeEmbalagem As Table
Dim TipoDeEmbalagemAberto As Boolean
Dim IndiceTipoDeEmbalagemAtivo$

Dim TBLUnidades As Table
Dim UnidadesAberto As Boolean
Dim IndiceUnidadesAtivo$

Dim TBLLocal As Table
Dim LocalAberto As Boolean
Dim IndiceLocalAtivo$

Dim TBLFornecedor As Table
Dim FornecedorAberto As Boolean
Dim IndiceFornecedorAtivo$

Dim TBLParâmetros As Table
Dim ParâmetrosAberto As Boolean

Dim ArrayDepartamento() As String
Dim ArraySeção() As String
Dim ArrayTipoDeICM() As String
Dim ArrayTipoDeEmbalagem() As String
Dim ArrayUnidades() As String
Dim ArrayLocal() As String

'Matriz para a chamada da janela de inclusão do código do produto
Public ArrayProdutoTotal%
Dim ArrayProdutoFornecedor() As Variant
Dim ArrayProdutoCódigo() As Variant

'Matriz para a chamada da janela de inclusão de lotes
Public ArrayLoteTotal%
Dim ArrayLote() As Variant

'Matriz para a chamada da janela de inclusão de preço
Public ArrayPreçoTotal%
Dim ArrayPreçoFornecedor() As Variant
Dim ArrayPreçoCusto() As Variant
Dim ArrayPreçoVenda() As Variant
Dim ArrayPreçoLucro() As Variant

Public lInserir As Boolean
Public lAlterar As Boolean

Dim mFechar As Boolean

Public lAlterarArrayProduto As Boolean
Public lAlterarArrayLote As Boolean
Public lAlterarArrayPreço As Boolean

Public StatusBarAviso$

Dim lPula As Boolean

Dim DataBaseName(1 To 1) As String
Public Relatório$
Public TotalDatabaseName%

Public mÚltimoDígito As Byte

Public lAtualizar As Boolean
Private Sub AtivaCampos()
    BotãoImprimir True
    frEspecificações.Enabled = True
    frImpostos.Enabled = True
    frDescontoPromoção.Enabled = True
    frVariados.Enabled = True
    BotãoGravar (lInserir Or lAllowEdit)
    cmdCancelar.Enabled = (lInserir Or lAllowEdit)
    cmdGravar.Enabled = (lInserir Or lAllowEdit)
End Sub
Public Sub AtualizaQuantidade()
    Dim Cont As Integer, AuxTotal As Single
    
    txtQuantidadeDeLotes = ArrayLoteTotal
    
    If ArrayLoteTotal > 0 Then
        AuxTotal = 0
        For Cont = 1 To ArrayLoteTotal
            AuxTotal = AuxTotal + ValStr(ArrayLote(3, Cont)) * ValStr(ArrayLote(4, Cont))
        Next
        lPula = True
        txtQuantidade = FormatStringMask("@V #####0,00", StrVal(AuxTotal))
        lPula = False
    Else
        lPula = True
        txtQuantidade = FormatStringMask("@V #####0,00", StrVal(0))
        txtQuantidadeDeLotes = Empty
        lPula = False
    End If
End Sub
Public Sub Atualizar()
    FillDepartamento
    TBLDepartamento.MoveFirst
    TBLDepartamento.Move cmbDepartamento.ListIndex
    FillSeção TBLDepartamento("CÓDIGO")
    FillTipoDeEmbalagem
    FillLocal
    FillUnidades
    FillTipoDeICM
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
    
    BotãoIncluir lAllowInsert
    ClearArrayProduto
    ClearArrayPreço
    ClearArrayLote
    
    If TBLProduto.RecordCount = 0 Then
        NavegaçãoInferior False
        NavegaçãoSuperior False
        BotãoGravar False
        cmdGravar.Enabled = False
        cmdCancelar.Enabled = False
        DesativaCampos
        ZeraCampos
        Cancelamento = True
        lInserir = False
        lAlterar = False
        Exit Function
    End If
    
    lInserir = False
    lAlterar = False
    
    Cancelamento = True
    
    TestaInferior TBLProduto, lAllowEdit, lAllowDelete, lAllowConsult
    TestaSuperior TBLProduto, lAllowEdit, lAllowDelete, lAllowConsult
    
    GetRecords
End Function
Public Sub ClearArrayLote()
    lAlterarArrayLote = False
    frmLotes.lAlteração = False
    
    ArrayLoteTotal = 0
    ReDim ArrayLote(MAXCOL, 1)
End Sub
Public Sub ClearArrayPreço()
    lAlterarArrayPreço = False
    frmPreços.lAlteração = False
    
    ArrayPreçoTotal = 0
    ReDim ArrayPreçoFornecedor(1 To 1)
    ReDim ArrayPreçoCusto(1 To 1)
    ReDim ArrayPreçoVenda(1 To 1)
    ReDim ArrayPreçoLucro(1 To 1)
End Sub
Public Sub ClearArrayProduto()
    lAlterarArrayProduto = False
    frmCódigoDoProduto.lAlteração = False
    
    ArrayProdutoTotal = 0
    ReDim ArrayProdutoFornecedor(1 To 1)
    ReDim ArrayProdutoCódigo(1 To 1)
End Sub
Private Sub DesativaCampos()
    BotãoImprimir False
    frEspecificações.Enabled = False
    frImpostos.Enabled = False
    frDescontoPromoção.Enabled = False
    frVariados.Enabled = False
    cmdGravar.Enabled = False
    cmdCancelar.Enabled = False
    BotãoGravar False
End Sub
Public Sub Encontrar()
    If Not lAllowConsult Then
        Exit Sub
    End If
    If lInserir Then
        MsgBox "Você está em uma inclusão!", vbExclamation, Caption
        StatusBarAviso = "Finalize a inclusão"
        Exit Sub
    End If
    If lAlterar Then
        MsgBox "Você está em uma alteração!", vbExclamation, Caption
        StatusBarAviso = "Finalize a alteração"
        Exit Sub
    End If
    
    Set frmEncontraProduto.Janela = Me
    frmEncontraProduto.NoModal = True
    frmEncontraProduto.Show 0
End Sub
Public Sub Excluir()
    Dim Confirmação As Integer, Msg1$, Msg2$, CódigoDoProduto As Variant
    Dim SQL As String

    If lAlterar Then
       If Not Cancelamento Then
            Exit Sub
        End If
    End If

    StatusBarAviso = "Exclusão"
    BarraDeStatus StatusBarAviso
    
    Msg1 = "Você está preste a apagar um registro !"
    Msg2 = "Tem certeza?"
    Msg2 = String(((Len(Msg1) - Len(Msg2)) / 2), " ") + Msg2
    Confirmação = MsgBox(Msg1 + vbCr + Msg2, vbYesNo + vbQuestion + vbDefaultButton2, "Confirmação")
    
    If Confirmação = vbNo Then
        Exit Sub
    End If
    
    WS.BeginTrans
    
    CódigoDoProduto = TBLProduto("CÓDIGO")
    TBLProduto.Delete
    
    SQL = "Delete * From [CÓDIGO DO PRODUTO] Where [CÓDIGO DO PRODUTO]= " + Str(CódigoDoProduto)
    DBCadastro.Execute SQL
    
    SQL = "Delete * From [PREÇO DO PRODUTO] Where [CÓDIGO DO PRODUTO]= " + Str(CódigoDoProduto)
    DBCadastro.Execute SQL
    
    SQL = "Delete * From [LOTE DO PRODUTO] Where [CÓDIGO DO PRODUTO]= " + Str(CódigoDoProduto)
    DBCadastro.Execute SQL
    
    If Err <> 0 Then
        GeraMensagemDeErro "Produto - Excluir - " & txtDescrição, True
        StatusBarAviso = "Falha na exclusão"
        BarraDeStatus StatusBarAviso
        Exit Sub
    End If
    
    Log gUsuário, "Exclusão - Produto: " & txtDescrição
    
    WS.CommitTrans
        
    ClearArrayProduto
    ClearArrayPreço
    ClearArrayLote
    
    StatusBarAviso = "Exclusão bem sucedida"
    BarraDeStatus StatusBarAviso
    
    If TBLProduto.RecordCount = 0 Then
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
    
    If TBLProduto.BOF Then
        TBLProduto.MoveFirst
    ElseIf TBLProduto.EOF Then
        TBLProduto.MoveLast
    Else
        TBLProduto.MovePrevious
        If TBLProduto.BOF Then
            TBLProduto.MoveNext
        End If
    End If
    
    GetRecords
    
    TestaInferior TBLProduto, lAllowEdit, lAllowDelete, lAllowConsult
    TestaSuperior TBLProduto, lAllowEdit, lAllowDelete, lAllowConsult
End Sub
Private Sub FillArrayLote()
    TBLLote.Seek "=", Código
    
    If TBLLote.NoMatch Then
        Exit Sub
    End If
    
    ArrayLoteTotal = 0
    
    Do While Not TBLLote.EOF
        If TBLLote("CÓDIGO DO PRODUTO") <> TBLProduto("CÓDIGO") Then
            Exit Do
        End If
        
        ArrayLoteTotal = ArrayLoteTotal + 1
        
        SizeArrayLote ArrayLoteTotal
        
        ArrayLote(1, ArrayLoteTotal) = TBLLote("CÓDIGO DO LOTE")
        ArrayLote(2, ArrayLoteTotal) = TBLLote("DÍGITO DO LOTE")
        ArrayLote(3, ArrayLoteTotal) = TBLLote("MÚLTIPLO")
        ArrayLote(4, ArrayLoteTotal) = TBLLote("QUANTIDADE")
        
        TBLLote.MoveNext
    Loop
End Sub
Private Sub FillArrayPreço()
    Dim Aux$, Posição%
    
    TBLPreço.Seek "=", Código
    
    If TBLPreço.NoMatch Then
        Exit Sub
    End If
    
    ArrayPreçoTotal = 0
    
    Do While Not TBLPreço.EOF
        If TBLPreço("CÓDIGO DO PRODUTO") <> TBLProduto("CÓDIGO") Then
            Exit Do
        End If
    
        ArrayPreçoTotal = ArrayPreçoTotal + 1
        
        SizeArrayPreço ArrayPreçoTotal
        
        ArrayPreçoFornecedor(ArrayPreçoTotal) = TBLPreço("CÓDIGO DO FORNECEDOR")
        ArrayPreçoCusto(ArrayPreçoTotal) = FormatStringMask("@V ##.###.##0,00", StrVal(TBLPreço("PREÇO DE CUSTO")))
        ArrayPreçoVenda(ArrayPreçoTotal) = FormatStringMask("@V ##.###.##0,00", StrVal(TBLPreço("PREÇO DE VENDA")))
        ArrayPreçoLucro(ArrayPreçoTotal) = FormatStringMask("@V ##0,00", StrVal(TBLPreço("MARGEM DE LUCRO")))
        
        TBLPreço.MoveNext
    Loop
End Sub
Private Sub FillArrayProduto()
    TBLCódigoDoProduto.Seek "=", Código
    
    If TBLCódigoDoProduto.NoMatch Then
        Exit Sub
    End If
    
    ArrayProdutoTotal = 0
    
    Do While Not TBLCódigoDoProduto.EOF
        If TBLCódigoDoProduto("CÓDIGO DO PRODUTO") <> TBLProduto("CÓDIGO") Then
            Exit Do
        End If
        
        ArrayProdutoTotal = ArrayProdutoTotal + 1
        
        SizeArrayProduto ArrayProdutoTotal
        
            
        ArrayProdutoFornecedor(ArrayProdutoTotal) = TBLCódigoDoProduto("FORNECEDOR")
        ArrayProdutoCódigo(ArrayProdutoTotal) = TBLCódigoDoProduto("CÓDIGO DO FORNECEDOR")
        
        TBLCódigoDoProduto.MoveNext
    Loop
End Sub
Private Sub FillDepartamento()
    Dim Cont%
    
    ReDim ArrayDepartamento(0 To TBLDepartamento.RecordCount - 1)
    
    Cont = 0
    
    cmbDepartamento.Clear
    
    TBLDepartamento.MoveFirst
    
    Do While Not TBLDepartamento.EOF
        cmbDepartamento.AddItem TBLDepartamento("DESCRIÇÃO")
        ArrayDepartamento(Cont) = TBLDepartamento("CÓDIGO")
        Cont = Cont + 1
        TBLDepartamento.MoveNext
    Loop
    cmbDepartamento.ListIndex = 0
End Sub
Private Sub FillLocal()
    Dim Cont%
    
    ReDim ArrayLocal(0 To TBLLocal.RecordCount - 1)
    
    Cont = 0
    
    cmbLocal.Clear
    
    TBLLocal.MoveFirst
    
    Do While Not TBLLocal.EOF
        cmbLocal.AddItem TBLLocal("ENDEREÇO")
        ArrayLocal(Cont) = TBLLocal("CÓDIGO")
        Cont = Cont + 1
        TBLLocal.MoveNext
    Loop
    cmbLocal.ListIndex = 0
End Sub
Private Sub FillSeção(ByVal Código)
    Dim SQL$, TBLAux As Recordset, Cont%
    
    SQL = "SELECT SEÇÃO.CÓDIGO,SEÇÃO.DESCRIÇÃO FROM (DEPARTAMENTO INNER JOIN [DEPARTAMENTO - SEÇÃO] ON DEPARTAMENTO.CÓDIGO = [DEPARTAMENTO - SEÇÃO].[CÓDIGO DO DEPTO]) INNER JOIN SEÇÃO ON [DEPARTAMENTO - SEÇÃO].[CÓDIGO DA SEÇÃO] = SEÇÃO.CÓDIGO Where (([DEPARTAMENTO - SEÇÃO].[CÓDIGO DO DEPTO] = '" + Código + "') And ([Seção].[Código] = [DEPARTAMENTO - SEÇÃO].[CÓDIGO DA SEÇÃO])) ORDER BY SEÇÃO.DESCRIÇÃO"
    
    Set TBLAux = DBCadastro.OpenRecordset(SQL)
    
    cmbSeção.Clear
    
    If TBLAux.RecordCount = 0 Then
        Exit Sub
    End If
    
    ReDim ArraySeção(0 To TBLAux.RecordCount - 1)
    
    Cont = 0
    
    TBLAux.MoveFirst
    
    Do While Not TBLAux.EOF
        cmbSeção.AddItem TBLAux("DESCRIÇÃO")
        ArraySeção(Cont) = TBLAux("CÓDIGO")
        Cont = Cont + 1
        TBLAux.MoveNext
    Loop
    
    cmbSeção.ListIndex = 0
End Sub
Private Sub FillTipoDeEmbalagem()
    Dim Cont%
    
    ReDim ArrayTipoDeEmbalagem(0 To TBLTipoDeEmbalagem.RecordCount - 1)
    
    Cont = 0
    
    cmbTipoDeEmbalagem.Clear
    
    TBLTipoDeEmbalagem.MoveFirst
    
    Do While Not TBLTipoDeEmbalagem.EOF
        cmbTipoDeEmbalagem.AddItem TBLTipoDeEmbalagem("DESCRIÇÃO")
        ArrayTipoDeEmbalagem(Cont) = TBLTipoDeEmbalagem("CÓDIGO")
        Cont = Cont + 1
        TBLTipoDeEmbalagem.MoveNext
    Loop
    
    cmbTipoDeEmbalagem.ListIndex = 0
End Sub
Private Sub FillTipoDeICM()
    Dim Cont%
    
    ReDim ArrayTipoDeICM(1 To 2, 0 To TBLTipoDeICM.RecordCount - 1)
    
    Cont = 0
    
    cmbTipoDeICM.Clear
    
    TBLTipoDeICM.MoveFirst
    
    Do While Not TBLTipoDeICM.EOF
        cmbTipoDeICM.AddItem TBLTipoDeICM("DESCRIÇÃO")
        ArrayTipoDeICM(1, Cont) = TBLTipoDeICM("CÓDIGO")
        ArrayTipoDeICM(2, Cont) = TBLTipoDeICM("ICM")
        Cont = Cont + 1
        TBLTipoDeICM.MoveNext
    Loop
    cmbTipoDeICM.ListIndex = 0
End Sub
Private Sub FillUnidades()
    Dim Cont%
    
    ReDim ArrayUnidades(0 To TBLUnidades.RecordCount - 1)
    
    Cont = 0
    
    cmbUnidades.Clear
    
    TBLUnidades.MoveFirst
    
    Do While Not TBLUnidades.EOF
        cmbUnidades.AddItem TBLUnidades("DESCRIÇÃO")
        ArrayUnidades(Cont) = TBLUnidades("CÓDIGO")
        Cont = Cont + 1
        TBLUnidades.MoveNext
    Loop
    cmbUnidades.ListIndex = 0
End Sub
Public Function GetArrayLote(ByVal Item As Byte, ByVal Elemento As Integer) As String
    GetArrayLote = ArrayLote(Item, Elemento)
End Function
Public Function GetArrayPreço(ByVal Nome As String, ByVal Elemento As Integer) As String
    If Nome = "Fornecedor" Then
        GetArrayPreço = ArrayPreçoFornecedor(Elemento)
    ElseIf Nome = "Custo" Then
        GetArrayPreço = ArrayPreçoCusto(Elemento)
    ElseIf Nome = "Venda" Then
        GetArrayPreço = ArrayPreçoVenda(Elemento)
    ElseIf Nome = "Lucro" Then
        GetArrayPreço = ArrayPreçoLucro(Elemento)
    End If
End Function
Public Function GetArrayProduto(ByVal Nome As String, ByVal Elemento As Integer) As String
    If Nome = "Fornecedor" Then
        GetArrayProduto = ArrayProdutoFornecedor(Elemento)
    ElseIf Nome = "Código" Then
        GetArrayProduto = ArrayProdutoCódigo(Elemento)
    End If
End Function
Public Sub Gravar()
    If lInserir Then
        
        'Pega o novo código interno do produto e atualiza na Tabela Parâmetros
        TBLParâmetros.Edit
        mCódigo = TBLParâmetros("PRODUTO") + 1
        TBLParâmetros("PRODUTO") = mCódigo
        TBLParâmetros.Update
        
        If SetRecords Then
            PosRecords
            lInserir = False
            ClearArrayProduto
            ClearArrayPreço
            ClearArrayLote
            StatusBarAviso = "Inclusão bem sucedida"
        Else
            StatusBarAviso = "Falha na inclusão"
        End If
    ElseIf lAlterar Then
        If TBLProduto.RecordCount > 0 And Not TBLProduto.BOF And Not TBLProduto.EOF Then
            mCódigo = TBLProduto("CÓDIGO")
            If SetRecords Then
                PosRecords
                lAlterar = False
                ClearArrayProduto
                ClearArrayPreço
                ClearArrayLote
                StatusBarAviso = "Alteração bem sucedida"
            Else
                StatusBarAviso = "Falha na alteração"
            End If
        End If
    End If
    
    BarraDeStatus StatusBarAviso
    
    TestaInferior TBLProduto, lAllowEdit, lAllowDelete, lAllowConsult
    TestaSuperior TBLProduto, lAllowEdit, lAllowDelete, lAllowConsult
    
    If TBLProduto.RecordCount = 0 Then
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
    
    If txtDescrição.Enabled Then
        txtDescrição.SetFocus
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
    
    txtDescrição.SetFocus
End Sub
Public Sub MoveFirst()
    If lAlterar Then
       If Not Cancelamento Then
            Exit Sub
        End If
    End If
    
    StatusBarAviso = "Pronto"
    BarraDeStatus StatusBarAviso
    
    TBLProduto.MoveFirst
    
    ClearArrayProduto
    ClearArrayPreço
    ClearArrayLote
    
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
    
    TBLProduto.MoveLast
    
    ClearArrayProduto
    ClearArrayPreço
    ClearArrayLote
    
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
    
    TBLProduto.MoveNext
    
    If TBLProduto.EOF Then
        TBLProduto.MovePrevious
        Exit Sub
    End If
    
    ClearArrayProduto
    ClearArrayPreço
    ClearArrayLote
    
    StatusBarAviso = "Pronto"
    BarraDeStatus StatusBarAviso
    
    NavegaçãoInferior lAllowConsult
    TestaSuperior TBLProduto, lAllowEdit, lAllowDelete, lAllowConsult
    
    GetRecords
End Sub
Public Sub MovePrevious()
    If lAlterar Then
       If Not Cancelamento Then
            Exit Sub
        End If
    End If
    
    TBLProduto.MovePrevious
    
    If TBLProduto.BOF Then
        TBLProduto.MoveNext
        Exit Sub
    End If
    
    ClearArrayProduto
    ClearArrayPreço
    ClearArrayLote
    
    StatusBarAviso = "Pronto"
    BarraDeStatus StatusBarAviso
    
    NavegaçãoSuperior lAllowConsult
    TestaInferior TBLProduto, lAllowEdit, lAllowDelete, lAllowConsult
    
    GetRecords
End Sub
Public Sub Posicionar(ByVal Código)
    mCódigo = Val(Código)
    PosRecords
    SetFocus
End Sub
Public Sub PosRecords()
    If TBLProduto.RecordCount = 0 Then
            Exit Sub
    End If

    TBLProduto.Seek "=", mCódigo
    If TBLProduto.NoMatch Then
        'MsgBox "Não consegui encontrar o cliente com CGC/CPF " + txtCgcCpf, vbExclamation, "Erro"
        TBLProduto.MoveFirst
        NavegaçãoInferior False
        NavegaçãoInferior lAllowConsult
    Else
        TestaInferior TBLProduto, lAllowEdit, lAllowDelete, lAllowConsult
        TestaSuperior TBLProduto, lAllowEdit, lAllowDelete, lAllowConsult
    End If
    GetRecords
End Sub
Public Function PushDataBaseName(ByVal Posição As Integer) As String
    PushDataBaseName = DataBaseName(Posição)
End Function
Private Sub GetRecords()
    On Error GoTo Erro
    
    Dim Aux$, Posição%, AuxBookMark, Cont As Integer, AuxTotal As Single
    
    lPula = True
    
    If Not lAllowConsult Then
        ZeraCampos
        DesativaCampos
        lPula = False
        Exit Sub
    End If
    
    Código = TBLProduto("CÓDIGO")
    txtDescrição = TBLProduto("DESCRIÇÃO")
    
    Aux = Trim(StrVal(TBLProduto("PESO")))
    txtPeso = FormatStringMask("@V ##0,000", Aux)
    
    TBLDepartamento.Index = "DEPARTAMENTO1"
    TBLDepartamento.Seek "=", Mid(TBLProduto("DEPTO - SEÇÃO"), 1, 4)
    cmbDepartamento.Text = TBLDepartamento("DESCRIÇÃO")
    cmbDepartamento_LostFocus
    AuxBookMark = TBLDepartamento.Bookmark
    TBLDepartamento.Index = IndiceDepartamentoAtivo
    TBLDepartamento.Bookmark = AuxBookMark
    FillSeção TBLDepartamento("CÓDIGO")
        
    TBLSeção.Index = "SEÇÃO1"
    TBLSeção.Seek "=", Mid(TBLProduto("DEPTO - SEÇÃO"), 5, 4)
    cmbSeção.Text = TBLSeção("DESCRIÇÃO")
    cmbSeção_LostFocus
    AuxBookMark = TBLSeção.Bookmark
    TBLSeção.Index = IndiceSeçãoAtivo
    TBLSeção.Bookmark = AuxBookMark
    
    TBLUnidades.Index = "UNIDADES1"
    TBLUnidades.Seek "=", TBLProduto("UNIDADES")
    cmbUnidades.Text = TBLUnidades("DESCRIÇÃO")
    cmbUnidades_LostFocus
    AuxBookMark = TBLUnidades.Bookmark
    TBLUnidades.Index = IndiceUnidadesAtivo
    TBLUnidades.Bookmark = AuxBookMark
    
    TBLTipoDeEmbalagem.Index = "TIPODEEMBALAGEM1"
    TBLTipoDeEmbalagem.Seek "=", TBLProduto("TIPO DE EMBALAGEM")
    cmbTipoDeEmbalagem.Text = TBLTipoDeEmbalagem("DESCRIÇÃO")
    cmbTipoDeEmbalagem_LostFocus
    AuxBookMark = TBLTipoDeEmbalagem.Bookmark
    TBLTipoDeEmbalagem.Index = IndiceTipoDeEmbalagemAtivo
    TBLTipoDeEmbalagem.Bookmark = AuxBookMark
    
    TBLTipoDeICM.Index = "TIPODEICM1"
    TBLTipoDeICM.Seek "=", TBLProduto("TIPO DE ICM")
    cmbTipoDeICM.Text = TBLTipoDeICM("DESCRIÇÃO")
    cmbTipoDeICM_LostFocus
    AuxBookMark = TBLTipoDeICM.Bookmark
    TBLTipoDeICM.Index = IndiceTipoDeICMAtivo
    TBLTipoDeICM.Bookmark = AuxBookMark
    
    TBLLocal.Index = "LOCALDOPRODUTO1"
    TBLLocal.Seek "=", TBLProduto("LOCAL")
    If TBLLocal.NoMatch Then
        TBLLocal.MoveFirst
    End If
    cmbLocal.Text = TBLLocal("ENDEREÇO")
    cmbLocal_LostFocus
    AuxBookMark = TBLLocal.Bookmark
    TBLLocal.Index = IndiceLocalAtivo
    TBLLocal.Bookmark = AuxBookMark
    
    txtQuantidade = FormatStringMask("@V ######0,00", Trim(StrVal(TBLProduto("QUANTIDADE"))))
    txtQuantidade.Locked = Not lAllowEditAmount
    
    txtQuantidadeDeLotes = FormatStringMask("@N #######", Trim(Str(TBLProduto("QUANTIDADE DE LOTES"))))
    chkLotes.Value = IIf(TBLProduto("LOTES"), 1, 0)
    
    If TBLProduto("LOTES") Then
        FillArrayLote
        txtQuantidade.Enabled = False
        vscrQuantidade.Enabled = False
        cmdLotes.Enabled = True
        AuxTotal = 0
        mÚltimoDígito = 0
        For Cont = 1 To ArrayLoteTotal
            AuxTotal = AuxTotal + ArrayLote(3, Cont) * ArrayLote(4, Cont)
            If mÚltimoDígito < ArrayLote(2, Cont) Then
                mÚltimoDígito = ArrayLote(2, Cont)
            End If
        Next
        txtQuantidade = FormatStringMask("@V #####0,00", StrVal(AuxTotal))
    Else
        txtQuantidade.Enabled = True
        vscrQuantidade.Enabled = True
        cmdLotes.Enabled = False
        mÚltimoDígito = 0
    End If
    
    Aux = Trim(StrVal(TBLProduto("ICM")))
    txtICM = FormatStringMask("@V #0,00", Aux)
    
    Aux = Trim(StrVal(TBLProduto("IPI")))
    txtIPI = FormatStringMask("@V #0,00", Aux)
    
    Aux = Trim(StrVal(TBLProduto("DESCONTO DE PROMOÇÃO")))
    txtDescontoDePromoção = FormatStringMask("@V #0,00", Aux)
    
    If TBLProduto("INÍCIO") <> vbNull Then
        txtInício = FormatStringMask(CheckDataMask, TBLProduto("INÍCIO"))
        CorrigeData DataMask, txtInício, TBLProduto("INÍCIO")
    Else
        txtInício = DataNula
    End If
        
    If TBLProduto("TÉRMINO") <> vbNull Then
        txtTérmino = FormatStringMask(CheckDataMask, TBLProduto("TÉRMINO"))
        CorrigeData DataMask, txtTérmino, TBLProduto("TÉRMINO")
    Else
        txtTérmino = DataNula
    End If
    
    Aux = Trim(StrVal(TBLProduto("DESCONTO MÁXIMO")))
    txtDescontoMáximo = FormatStringMask("@V #0,00", Aux)
    
    lPula = False
    If Not lAllowEdit Then
        DesativaCampos
    End If
    
    Exit Sub
    
Erro:
    GeraMensagemDeErro "Produto - GetRecords "
    Resume Next
End Sub
Private Function SetRecords()
    On Error GoTo ErroProd
    
    Dim Msg$
    Dim Confirmação As Integer, Msg1$, Msg2$
    Dim SQL As String
    Dim Cont%, Recno%
    
    WS.BeginTrans 'Inicia uma Transação
    
    If lInserir Then
        TBLProduto.AddNew
    Else
        TBLProduto.Edit
    End If
    
    TBLProduto("CÓDIGO") = mCódigo
    TBLProduto("DESCRIÇÃO") = txtDescrição
    TBLProduto("PESO") = ValStr(txtPeso) ' 0,00
    TBLProduto("QUANTIDADE") = ValStr(txtQuantidade) '0,00
    TBLProduto("LOTES") = IIf(chkLotes.Value = 1, True, False)
    TBLProduto("QUANTIDADE DE LOTES") = Val(txtQuantidadeDeLotes) '0
    TBLProduto("IPI") = ValStr(txtIPI) '0,00
    TBLProduto("ICM") = ValStr(ArrayTipoDeICM(2, cmbTipoDeICM.ListIndex)) '0,00
    TBLProduto("DESCONTO DE PROMOÇÃO") = ValStr(txtDescontoDePromoção) '0,00
    TBLProduto("INÍCIO") = IIf(Trim(StrTran(txtInício, "/")) <> Empty, txtInício, vbNull)
    TBLProduto("TÉRMINO") = IIf(Trim(StrTran(txtTérmino, "/")) <> Empty, txtTérmino, vbNull)
    TBLProduto("DESCONTO MÁXIMO") = ValStr(txtDescontoMáximo) '0,00
    TBLProduto("TIPO DE EMBALAGEM") = ArrayTipoDeEmbalagem(cmbTipoDeEmbalagem.ListIndex)
    TBLProduto("TIPO DE ICM") = ArrayTipoDeICM(1, cmbTipoDeICM.ListIndex)
    TBLProduto("DEPTO - SEÇÃO") = ArrayDepartamento(cmbDepartamento.ListIndex) + ArraySeção(cmbSeção.ListIndex)
    TBLProduto("UNIDADES") = ArrayUnidades(cmbUnidades.ListIndex)
    TBLProduto("LOCAL") = ArrayLocal(cmbLocal.ListIndex)
    If lInserir Then
        TBLProduto("USERNAME - CRIA") = gUsuário
        TBLProduto("DATA - CRIA") = Date
        TBLProduto("HORA - CRIA") = Time
        TBLProduto("USERNAME - ALTERA") = "VAZIO"
        TBLProduto("DATA - ALTERA") = vbNull
        TBLProduto("HORA - ALTERA") = vbNull
    End If
    If lAlterar Then
        TBLProduto("USERNAME - ALTERA") = gUsuário
        TBLProduto("DATA - ALTERA") = Date
        TBLProduto("HORA - ALTERA") = Time
    End If
    TBLProduto.Update
    
ErroProd:
    If Err <> 0 Then
        TBLProduto.CancelUpdate
        GeraMensagemDeErro "Produto - SetRecords - ErroProd - " & txtDescrição, True
        SetRecords = False
        Exit Function
    End If
        
    On Error GoTo ErroCód
    
    If lAlterarArrayProduto Then
        SQL = "Delete * From [CÓDIGO DO PRODUTO] Where [CÓDIGO DO PRODUTO]= " & mCódigo
        DBCadastro.Execute SQL
        
        For Cont = 1 To ArrayProdutoTotal
            TBLCódigoDoProduto.AddNew

            TBLCódigoDoProduto("CÓDIGO DO PRODUTO") = mCódigo
            TBLCódigoDoProduto("FORNECEDOR") = ArrayProdutoFornecedor(Cont)
            TBLCódigoDoProduto("CÓDIGO DO FORNECEDOR") = ArrayProdutoCódigo(Cont)
            TBLCódigoDoProduto("USERNAME - CRIA") = gUsuário
            TBLCódigoDoProduto("DATA - CRIA") = Date
            TBLCódigoDoProduto("HORA - CRIA") = Time
            TBLCódigoDoProduto("USERNAME - ALTERA") = "VAZIO"
            TBLCódigoDoProduto("DATA - ALTERA") = vbNull
            TBLCódigoDoProduto("HORA - ALTERA") = vbNull
            TBLCódigoDoProduto.Update
        Next
    End If
    
ErroCód:
    If Err <> 0 Then
        TBLCódigoDoProduto.CancelUpdate
        GeraMensagemDeErro "Produto - SetRecords - ErroCód - " & txtDescrição, True
        SetRecords = False
        Exit Function
    End If
    
    On Error GoTo ErroPreço
    
    If lAlterarArrayPreço Then
        SQL = "Delete * From [PREÇO DO PRODUTO] Where [CÓDIGO DO PRODUTO]= " & mCódigo
        DBCadastro.Execute SQL
    
        For Cont = 1 To ArrayPreçoTotal
            TBLPreço.AddNew
            TBLPreço("CÓDIGO DO PRODUTO") = mCódigo
            TBLPreço("CÓDIGO DO FORNECEDOR") = ArrayPreçoFornecedor(Cont)
            TBLPreço("PREÇO DE CUSTO") = ValStr(ArrayPreçoCusto(Cont))
            TBLPreço("PREÇO DE VENDA") = ValStr(ArrayPreçoVenda(Cont))
            TBLPreço("MARGEM DE LUCRO") = ValStr(ArrayPreçoLucro(Cont))
            TBLPreço("USERNAME - CRIA") = gUsuário
            TBLPreço("DATA - CRIA") = Date
            TBLPreço("HORA - CRIA") = Time
            TBLPreço("USERNAME - ALTERA") = "VAZIO"
            TBLPreço("DATA - ALTERA") = vbNull
            TBLPreço("HORA - ALTERA") = vbNull
            TBLPreço.Update
        Next
    End If
    
ErroPreço:
    If Err <> 0 Then
        TBLPreço.CancelUpdate
        GeraMensagemDeErro "Produto - SetRecords - ErroPreço - " & txtDescrição, True
        SetRecords = False
        Exit Function
    End If
    
    On Error GoTo ErroLote
    
    If lAlterarArrayLote Then
        SQL = "Delete * From [LOTE DO PRODUTO] Where [CÓDIGO DO PRODUTO]= " & mCódigo
        DBCadastro.Execute SQL
        For Cont = 1 To ArrayLoteTotal
            TBLLote.AddNew
            TBLLote("CÓDIGO DO LOTE") = ArrayLote(1, Cont)
            TBLLote("DÍGITO DO LOTE") = ArrayLote(2, Cont)
            TBLLote("MÚLTIPLO") = ValStr(ArrayLote(3, Cont))
            TBLLote("QUANTIDADE") = ValStr(ArrayLote(4, Cont))
            TBLLote("CÓDIGO DO PRODUTO") = TBLProduto("CÓDIGO")
            TBLLote("USERNAME - CRIA") = gUsuário
            TBLLote("DATA - CRIA") = Date
            TBLLote("HORA - CRIA") = Time
            TBLLote("USERNAME - ALTERA") = "VAZIO"
            TBLLote("DATA - ALTERA") = vbNull
            TBLLote("HORA - ALTERA") = vbNull
            TBLLote.Update
        Next
    End If
    
ErroLote:
    If Err.Number <> 0 Then
        TBLLote.CancelUpdate
        GeraMensagemDeErro "Produto - SetRecords - ErroLote - " & txtDescrição, True
        SetRecords = False
        Exit Function
    End If

    WS.CommitTrans 'Grava as alterações ou inclusões se não houverem erros
        
    If lInserir Then
        Log gUsuário, "Inclusão - Produto " & txtDescrição
    Else
        Log gUsuário, "Alteração - Produto " & txtDescrição
    End If
    
    lAlterar = False
    lInserir = False
    ClearArrayLote
    ClearArrayProduto
    ClearArrayPreço
    
    Código = TBLProduto("CÓDIGO")
    
    SetRecords = True
End Function
Public Sub SetArrayLote(ByVal Item As Byte, ByVal Valor As String, ByVal Elemento As Integer)
    ArrayLote(Item, Elemento) = Valor
End Sub
Public Sub SetArrayPreço(ByVal Nome As String, ByVal Valor As String, ByVal Elemento As Integer)
    If Nome = "Fornecedor" Then
        ArrayPreçoFornecedor(Elemento) = Valor
    ElseIf Nome = "Custo" Then
        ArrayPreçoCusto(Elemento) = Valor
    ElseIf Nome = "Venda" Then
        ArrayPreçoVenda(Elemento) = Valor
    ElseIf Nome = "Lucro" Then
        ArrayPreçoLucro(Elemento) = Valor
    End If
End Sub
Public Sub SetArrayProduto(ByVal Nome As String, ByVal Valor As String, ByVal Elemento As Integer)
    If Nome = "Fornecedor" Then
        ArrayProdutoFornecedor(Elemento) = Valor
    ElseIf Nome = "Código" Then
        ArrayProdutoCódigo(Elemento) = Valor
    End If
End Sub
Public Sub SizeArrayLote(ByVal Tamanho As Integer)
    If Tamanho > 0 Then
        ArrayLoteTotal = Tamanho
        ReDim Preserve ArrayLote(MAXCOL, Tamanho)
    End If
End Sub
Public Sub SizeArrayPreço(ByVal Tamanho As Integer)
    ArrayPreçoTotal = Tamanho
    ASize Tamanho, ArrayPreçoFornecedor
    ASize Tamanho, ArrayPreçoCusto
    ASize Tamanho, ArrayPreçoVenda
    ASize Tamanho, ArrayPreçoLucro
End Sub
Public Sub SizeArrayProduto(ByVal Tamanho As Integer)
    ArrayProdutoTotal = Tamanho
    ASize Tamanho, ArrayProdutoFornecedor
    ASize Tamanho, ArrayProdutoCódigo
End Sub
Private Sub ZeraCampos()
    Código = Empty
    txtDescrição = Empty
    txtPeso = " 0,000"
    txtPeso_LostFocus
    txtQuantidade = "0,00"
    txtQuantidade.Locked = False
    txtQuantidadeDeLotes = "0"
    txtIPI = " 0,00"
    txtICM = " 0,00"
    txtDescontoDePromoção = " 0,00"
    txtInício = Empty
    txtTérmino = Empty
    txtDescontoMáximo = " 0,00"
    
    ArrayProdutoTotal = 0
    ArrayLoteTotal = 0
    ArrayPreçoTotal = 0
    
    ClearArrayProduto
    ClearArrayPreço
    ClearArrayLote
End Sub
Private Sub chkLotes_Click()
    If lPula Then
        Exit Sub
    End If
    If Not lInserir Then
        lAlterar = True
        StatusBarAviso = "Alteração"
        BarraDeStatus StatusBarAviso
    End If
    If chkLotes.Value = 1 Then
        txtQuantidade.Enabled = False
        vscrQuantidade.Enabled = False
        cmdLotes.Enabled = True
    Else
        txtQuantidade.Enabled = True
        vscrQuantidade.Enabled = True
        cmdLotes.Enabled = False
    End If
End Sub
Private Sub chkLotes_KeyPress(KeyAscii As Integer)
    If Not lInserir Then
        lAlterar = True
        StatusBarAviso = "Alteração"
        BarraDeStatus StatusBarAviso
    End If
End Sub
Private Sub cmbDepartamento_Click()
    If lPula Then
        Exit Sub
    End If
    TBLDepartamento.MoveFirst
    TBLDepartamento.Move cmbDepartamento.ListIndex
    FillSeção TBLDepartamento("CÓDIGO")
    If Not lInserir Then
        lAlterar = True
        StatusBarAviso = "Alteração"
        BarraDeStatus StatusBarAviso
    End If
End Sub
Private Sub cmbDepartamento_LostFocus()
    Dim Cont%, Encontrou As Boolean
    
    Encontrou = False
    
    For Cont = 0 To cmbDepartamento.ListCount - 1
        If UCase(cmbDepartamento.List(Cont)) = UCase(cmbDepartamento.Text) Then
            Encontrou = True
            cmbDepartamento.ListIndex = Cont
            Exit Sub
        End If
    Next
    
    For Cont = 0 To cmbDepartamento.ListCount - 1
        If InStr(UCase(cmbDepartamento.List(Cont)), UCase(cmbDepartamento.Text)) = 1 Then
            Encontrou = True
            cmbDepartamento.ListIndex = Cont
            Exit Sub
        End If
    Next
    
    If Not Encontrou Then
        cmbDepartamento.ListIndex = 0
    End If
End Sub
Private Sub cmbLocal_Click()
    If lPula Then
        Exit Sub
    End If
    If Not lInserir Then
        lAlterar = True
        StatusBarAviso = "Alteração"
        BarraDeStatus StatusBarAviso
    End If
End Sub
Private Sub cmbLocal_LostFocus()
    Dim Cont%, Encontrou As Boolean
    
    Encontrou = False
    
    For Cont = 0 To cmbLocal.ListCount - 1
        If UCase(cmbLocal.List(Cont)) = UCase(cmbLocal.Text) Then
            Encontrou = True
            cmbLocal.ListIndex = Cont
            Exit Sub
        End If
    Next
    
    For Cont = 0 To cmbLocal.ListCount - 1
        If InStr(UCase(cmbLocal.List(Cont)), UCase(cmbLocal.Text)) = 1 Then
            Encontrou = True
            cmbLocal.ListIndex = Cont
            Exit Sub
        End If
    Next
    
    If Not Encontrou Then
        cmbLocal.ListIndex = 0
    End If
End Sub
Private Sub cmbSeção_Click()
    If lPula Then
        Exit Sub
    End If
    If Not lInserir Then
        lAlterar = True
        StatusBarAviso = "Alteração"
        BarraDeStatus StatusBarAviso
    End If
End Sub
Private Sub cmbSeção_LostFocus()
    Dim Cont%, Encontrou As Boolean
    
    Encontrou = False
    
    For Cont = 0 To cmbSeção.ListCount - 1
        If UCase(cmbSeção.List(Cont)) = UCase(cmbSeção.Text) Then
            Encontrou = True
            cmbSeção.ListIndex = Cont
            Exit Sub
        End If
    Next
    
    For Cont = 0 To cmbSeção.ListCount - 1
        If InStr(UCase(cmbSeção.List(Cont)), UCase(cmbSeção.Text)) = 1 Then
            Encontrou = True
            cmbSeção.ListIndex = Cont
            Exit Sub
        End If
    Next
    
    If Not Encontrou Then
        cmbSeção.ListIndex = 0
    End If
End Sub
Private Sub cmbTipoDeEmbalagem_Click()
    If lPula Then
        Exit Sub
    End If
    If Not lInserir Then
        lAlterar = True
        StatusBarAviso = "Alteração"
        BarraDeStatus StatusBarAviso
    End If
End Sub
Private Sub cmbTipoDeEmbalagem_LostFocus()
    Dim Cont%, Encontrou As Boolean
    
    Encontrou = False
    
    For Cont = 0 To cmbTipoDeEmbalagem.ListCount - 1
        If UCase(cmbTipoDeEmbalagem.List(Cont)) = UCase(cmbTipoDeEmbalagem.Text) Then
            Encontrou = True
            cmbTipoDeEmbalagem.ListIndex = Cont
            Exit Sub
        End If
    Next
    
    For Cont = 0 To cmbTipoDeEmbalagem.ListCount - 1
        If InStr(UCase(cmbTipoDeEmbalagem.List(Cont)), UCase(cmbTipoDeEmbalagem.Text)) = 1 Then
            Encontrou = True
            cmbTipoDeEmbalagem.ListIndex = Cont
            Exit Sub
        End If
    Next
    
    If Not Encontrou Then
        cmbTipoDeEmbalagem.ListIndex = 0
    End If
End Sub
Private Sub cmbTipoDeICM_Click()
    If lPula Then
        Exit Sub
    End If
    If Not lInserir Then
        lAlterar = True
        StatusBarAviso = "Alteração"
        BarraDeStatus StatusBarAviso
    End If
End Sub
Private Sub cmbTipoDeICM_LostFocus()
    Dim Cont%, Encontrou As Boolean
    
    Encontrou = False
    
    For Cont = 0 To cmbTipoDeICM.ListCount - 1
        If UCase(cmbTipoDeICM.List(Cont)) = UCase(cmbTipoDeICM.Text) Then
            Encontrou = True
            cmbTipoDeICM.ListIndex = Cont
            txtICM.Text = ArrayTipoDeICM(2, cmbTipoDeICM.ListIndex)
            FormatMask "@V #0,00", txtICM
            Exit Sub
        End If
    Next
    
    For Cont = 0 To cmbTipoDeICM.ListCount - 1
        If InStr(UCase(cmbTipoDeICM.List(Cont)), UCase(cmbTipoDeICM.Text)) = 1 Then
            Encontrou = True
            cmbTipoDeICM.ListIndex = Cont
            txtICM.Text = ArrayTipoDeICM(2, cmbTipoDeICM.ListIndex)
            FormatMask "@V #0,00", txtICM
            Exit Sub
        End If
    Next
    
    If Not Encontrou Then
        cmbTipoDeICM.ListIndex = 0
    End If
End Sub
Private Sub cmbUnidades_Click()
    If lPula Then
        Exit Sub
    End If
    If Not lInserir Then
        lAlterar = True
        StatusBarAviso = "Alteração"
        BarraDeStatus StatusBarAviso
    End If
End Sub
Private Sub cmbUnidades_LostFocus()
    Dim Cont%, Encontrou As Boolean
    
    Encontrou = False
    
    For Cont = 0 To cmbUnidades.ListCount - 1
        If UCase(cmbUnidades.List(Cont)) = UCase(cmbUnidades.Text) Then
            Encontrou = True
            cmbUnidades.ListIndex = Cont
            Exit Sub
        End If
    Next
    
    For Cont = 0 To cmbUnidades.ListCount - 1
        If InStr(UCase(cmbUnidades.List(Cont)), UCase(cmbUnidades.Text)) = 1 Then
            Encontrou = True
            cmbUnidades.ListIndex = Cont
            Exit Sub
        End If
    Next
    
    If Not Encontrou Then
        cmbUnidades.ListIndex = 0
    End If
End Sub
Private Sub cmdCancelar_Click()
    Cancelamento
End Sub
Private Sub cmdCódigo_Click()
    If Not lInserir Then
        StatusBarAviso = "Alteração"
        BarraDeStatus StatusBarAviso
    End If
    If Not lAlterarArrayProduto Then
        FillArrayProduto
    End If
    frmCódigoDoProduto.Show 0
End Sub
Private Sub cmdGravar_Click()
    Gravar
End Sub
Private Sub cmdLotes_Click()
    If Not lInserir Then
        StatusBarAviso = "Alteração"
        BarraDeStatus StatusBarAviso
    End If
    If Not lAlterarArrayLote Then
        FillArrayLote
    End If
    Set frmLotes.mJanela = Me
    frmLotes.mÚltimoDígito = mÚltimoDígito
    frmLotes.Show 0
End Sub
Private Sub cmdPreços_Click()
    If Not lInserir Then
        StatusBarAviso = "Alteração"
        BarraDeStatus StatusBarAviso
    End If
    If Not lAlterarArrayPreço Then
        FillArrayPreço
    End If
    frmPreços.Show 0
End Sub
Private Sub Form_Activate()
    If mFechar Then
        Unload Me
        Exit Sub
    End If
    If Not ProdutoAberto Then
        Unload Me
        Exit Sub
    End If
    If Not CódigoDoProdutoAberto Then
        Unload Me
        Exit Sub
    End If
    If Not LoteAberto Then
        Unload Me
        Exit Sub
    End If
    If Not PreçoAberto Then
        Unload Me
        Exit Sub
    End If
    If Not DepartamentoAberto Then
        Unload Me
        Exit Sub
    End If
    If Not SeçãoAberto Then
        Unload Me
        Exit Sub
    End If
    If Not DepartamentoSeçãoAberto Then
        Unload Me
        Exit Sub
    End If
    If Not TipoDeICMAberto Then
        Unload Me
        Exit Sub
    End If
    If Not TipoDeEmbalagemAberto Then
        Unload Me
        Exit Sub
    End If
    If Not UnidadesAberto Then
        Unload Me
        Exit Sub
    End If
    If Not LocalAberto Then
        Unload Me
        Exit Sub
    End If
    If Not FornecedorAberto Then
        Unload Me
        Exit Sub
    End If
    
    If Not ParâmetrosAberto Then
        Unload Me
        Exit Sub
    End If
    
    TestaInferior TBLProduto, lAllowEdit, lAllowDelete, lAllowConsult
    TestaSuperior TBLProduto, lAllowEdit, lAllowDelete, lAllowConsult
    If TBLProduto.RecordCount = 0 Then
        BotãoGravar False
        cmdGravar.Enabled = False
        cmdCancelar.Enabled = False
        BotãoImprimir False
    Else
        BotãoGravar (lInserir Or lAllowEdit)
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
    
    ZeraCampos
        
    lAllowInsert = Allow("PRODUTO", "I")
    lAllowEdit = Allow("PRODUTO", "A")
    lAllowDelete = Allow("PRODUTO", "E")
    lAllowConsult = Allow("PRODUTO", "C")
    lAllowEditAmount = Allow("PRODUTO", "Q")
    
    lAtualizar = True 'Indica que o modulo possui a função atualizar
    
    lInserir = False
    lAlterar = False
    lPula = False
    
    'Abertura das tabelas
    ProdutoAberto = AbreTabela(Dicionário, "CADASTRO", "PRODUTO", DBCadastro, TBLProduto, TBLTabela, dbOpenTable)
    
    If ProdutoAberto Then
        IndiceProdutoAtivo = "PRODUTO1"
        TBLProduto.Index = IndiceProdutoAtivo
    Else
        MsgBox "Não consegui abrir a tabela 'Produto' !", vbCritical, "Erro"
        Exit Sub
    End If
    
    CódigoDoProdutoAberto = AbreTabela(Dicionário, "CADASTRO", "CÓDIGO DO PRODUTO", DBCadastro, TBLCódigoDoProduto, TBLTabela, dbOpenTable)
    
    If CódigoDoProdutoAberto Then
        IndiceCódigoDoProdutoAtivo = "CÓDIGODOPRODUTO2"
        TBLCódigoDoProduto.Index = IndiceCódigoDoProdutoAtivo
    Else
        MsgBox "Não consegui abrir a tabela 'Código do Produto' !", vbCritical, "Erro"
        Exit Sub
    End If
    
    LoteAberto = AbreTabela(Dicionário, "CADASTRO", "LOTE DO PRODUTO", DBCadastro, TBLLote, TBLTabela, dbOpenTable)
    
    If LoteAberto Then
        IndiceLoteAtivo = "LOTEDOPRODUTO2"
        TBLLote.Index = IndiceLoteAtivo
    Else
        MsgBox "Não consegui abrir a tabela 'Lote do Produto' !", vbCritical, "Erro"
        Exit Sub
    End If
    
    PreçoAberto = AbreTabela(Dicionário, "CADASTRO", "PREÇO DO PRODUTO", DBCadastro, TBLPreço, TBLTabela, dbOpenTable)
    
    If PreçoAberto Then
        IndicePreçoAtivo = "PREÇODOPRODUTO2"
        TBLPreço.Index = IndicePreçoAtivo
    Else
        MsgBox "Não consegui abrir a tabela 'Preço do Produto' !", vbCritical, "Erro"
        Exit Sub
    End If
    
    DepartamentoAberto = AbreTabela(Dicionário, "CADASTRO", "DEPARTAMENTO", DBCadastro, TBLDepartamento, TBLTabela, dbOpenTable)
        
    If DepartamentoAberto Then
        If TBLDepartamento.RecordCount = 0 Then
            MsgBox "Tabela 'Departamento' está vazia! " + vbCr + "Antes de tentar cadastrar um produto, primeiro cadastre esta tabela.", vbInformation, "Aviso"
            DepartamentoAberto = False
            Exit Sub
        End If
        IndiceDepartamentoAtivo = "DEPARTAMENTO2"
        TBLDepartamento.Index = IndiceDepartamentoAtivo
    Else
        MsgBox "Não consegui abrir a tabela 'Departamento' !", vbCritical, "Erro"
        Exit Sub
    End If
    
    SeçãoAberto = AbreTabela(Dicionário, "CADASTRO", "SEÇÃO", DBCadastro, TBLSeção, TBLTabela, dbOpenTable)
    
    If SeçãoAberto Then
        If TBLSeção.RecordCount = 0 Then
            MsgBox "Tabela 'Seção' está vazia! " + vbCr + "Antes de tentar cadastrar um produto, primeiro cadastre esta tabela.", vbInformation, "Aviso"
            SeçãoAberto = False
            Exit Sub
        End If
        IndiceSeçãoAtivo = "SEÇÃO2"
        TBLSeção.Index = IndiceSeçãoAtivo
    Else
        MsgBox "Não consegui abrir a tabela 'Seção' !", vbCritical, "Erro"
        Exit Sub
    End If
    
    DepartamentoSeçãoAberto = AbreTabela(Dicionário, "CADASTRO", "DEPARTAMENTO - SEÇÃO", DBCadastro, TBLDepartamentoSeção, TBLTabela, dbOpenTable)
    
    If DepartamentoSeçãoAberto Then
        If TBLDepartamentoSeção.RecordCount = 0 Then
            MsgBox "Tabela 'Departamento - Seção' está vazia! " + vbCr + "Antes de tentar cadastrar um produto, primeiro cadastre esta tabela.", vbInformation, "Aviso"
            DepartamentoSeçãoAberto = False
            Exit Sub
        End If
        IndiceDepartamentoSeçãoAtivo = "DEPARTAMENTOSEÇÃO1"
        TBLDepartamentoSeção.Index = IndiceDepartamentoSeçãoAtivo
    Else
        MsgBox "Não consegui abrir a tabela 'Departamento - Seção' !", vbCritical, "Erro"
        Exit Sub
    End If
    
    TipoDeICMAberto = AbreTabela(Dicionário, "CADASTRO", "TIPO DE ICM", DBCadastro, TBLTipoDeICM, TBLTabela, dbOpenTable)
    
    If TipoDeICMAberto Then
        If TBLTipoDeICM.RecordCount = 0 Then
            MsgBox "Tabela 'Tipo de ICM' está vazia! " + vbCr + "Antes de tentar cadastrar um produto, primeiro cadastre esta tabela.", vbInformation, "Aviso"
            TipoDeICMAberto = False
            Exit Sub
        End If
        IndiceTipoDeICMAtivo = "TIPODEICM2"
        TBLTipoDeICM.Index = IndiceTipoDeICMAtivo
    Else
        MsgBox "Não consegui abrir a tabela 'Tipo de ICM' !", vbCritical, "Erro"
        Exit Sub
    End If
    
    TipoDeEmbalagemAberto = AbreTabela(Dicionário, "CADASTRO", "TIPO DE EMBALAGEM", DBCadastro, TBLTipoDeEmbalagem, TBLTabela, dbOpenTable)
    
    If TipoDeEmbalagemAberto Then
        If TBLTipoDeEmbalagem.RecordCount = 0 Then
            MsgBox "Tabela 'Tipo de Embalagem' está vazia! " + vbCr + "Antes de tentar cadastrar um produto, primeiro cadastre esta tabela.", vbInformation, "Aviso"
            TipoDeEmbalagemAberto = False
            Exit Sub
        End If
        IndiceTipoDeEmbalagemAtivo = "TIPODEEMBALAGEM2"
        TBLTipoDeEmbalagem.Index = IndiceTipoDeEmbalagemAtivo
    Else
        MsgBox "Não consegui abrir a tabela 'Tipo de Embalagem' !", vbCritical, "Erro"
        Exit Sub
    End If
    
    UnidadesAberto = AbreTabela(Dicionário, "CADASTRO", "UNIDADES", DBCadastro, TBLUnidades, TBLTabela, dbOpenTable)
    
    If UnidadesAberto Then
        If TBLUnidades.RecordCount = 0 Then
            MsgBox "Tabela 'Unidades' está vazia! " + vbCr + "Antes de tentar cadastrar um produto, primeiro cadastre esta tabela.", vbInformation, "Aviso"
            UnidadesAberto = False
            Exit Sub
        End If
        IndiceUnidadesAtivo = "UNIDADES2"
        TBLUnidades.Index = IndiceUnidadesAtivo
    Else
        MsgBox "Não consegui abrir a tabela 'Unidades' !", vbCritical, "Erro"
        Exit Sub
    End If
    
    LocalAberto = AbreTabela(Dicionário, "CADASTRO", "LOCAL DO PRODUTO", DBCadastro, TBLLocal, TBLTabela, dbOpenTable)
    
    If LocalAberto Then
        If TBLLocal.RecordCount = 0 Then
            MsgBox "Tabela 'Local' está vazia! " + vbCr + "Antes de tentar cadastrar um produto, primeiro cadastre esta tabela.", vbInformation, "Aviso"
            LocalAberto = False
            Exit Sub
        End If
        IndiceLocalAtivo = "LOCALDOPRODUTO2"
        TBLLocal.Index = IndiceLocalAtivo
    Else
        MsgBox "Não consegui abrir a tabela 'Local do Produto' !", vbCritical, "Erro"
        Exit Sub
    End If
    
    FornecedorAberto = AbreTabela(Dicionário, "CADASTRO", "FORNECEDOR", DBCadastro, TBLFornecedor, TBLTabela, dbOpenTable)
        
    If FornecedorAberto Then
        If TBLFornecedor.RecordCount = 0 Then
            MsgBox "Tabela 'Fornecedor' está vazia! " + vbCr + "Antes de tentar cadastrar um produto, primeiro cadastre esta tabela.", vbInformation, "Aviso"
            FornecedorAberto = False
            Exit Sub
        End If
        IndiceFornecedorAtivo = "FORNECEDOR1"
        TBLFornecedor.Index = IndiceFornecedorAtivo
    Else
        MsgBox "Não consegui abrir a tabela 'Fornecedor' !", vbCritical, "Erro"
        Exit Sub
    End If
    
    ParâmetrosAberto = AbreTabela(Dicionário, "SISTEMA", "PARÂMETROS", DBSistema, TBLParâmetros, TBLTabela, dbOpenTable)
    
    If ParâmetrosAberto Then
    Else
        MsgBox "Não consegui abrir a tabela 'Parâmetros' !", vbCritical, "Erro"
        Exit Sub
    End If
    
    'Todas as tabelas foram abertas
    FillDepartamento
    TBLDepartamento.MoveFirst
    TBLDepartamento.Move cmbDepartamento.ListIndex
    
    FillSeção TBLDepartamento("CÓDIGO")
    
    FillUnidades
    
    FillTipoDeICM
    
    FillTipoDeEmbalagem
    
    FillLocal
    
    txtPeso = "0,000"
    txtPeso_LostFocus
    txtDescontoMáximo = " 0,00"
    txtDescontoDePromoção = " 0,00"
        
    BotãoIncluir lAllowInsert
 
    If TBLProduto.RecordCount = 0 Then
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
        
    If TBLProduto.RecordCount = 0 Or TBLProduto.RecordCount = 1 Then
        NavegaçãoSuperior False
    Else
        NavegaçãoInferior lAllowConsult
    End If
    
    lInserir = False
    lAlterar = False
    StatusBarAviso = "Pronto"
    Relatório = AddPath(AplicaçãoPath, "REPORT\PRODUTO.RPT")
    TotalDatabaseName = 1
    DataBaseName(1) = AddPath(AplicaçãoPath, "DATABASE\CADASTRO.MDB")
    mFechar = False
    Exit Sub
    
Erro:
    GeraMensagemDeErro "Produto - Load"
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
    
    Set frmProduto = Nothing
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Dim Cont As Byte
    
    If ProdutoAberto Then
        TBLProduto.Close
    End If
    If CódigoDoProdutoAberto Then
        TBLCódigoDoProduto.Close
    End If
    If LoteAberto Then
        TBLLote.Close
    End If
    If PreçoAberto Then
        TBLPreço.Close
    End If
    If DepartamentoAberto Then
        TBLDepartamento.Close
    End If
    If SeçãoAberto Then
        TBLSeção.Close
    End If
    If DepartamentoSeçãoAberto Then
        TBLDepartamentoSeção.Close
    End If
    If TipoDeICMAberto Then
        TBLTipoDeICM.Close
    End If
    If TipoDeEmbalagemAberto Then
        TBLTipoDeEmbalagem.Close
    End If
    If UnidadesAberto Then
        TBLUnidades.Close
    End If
    If LocalAberto Then
        TBLLocal.Close
    End If
    For Cont = 2 - 1 To Forms.Count - 1
        If Forms(Cont).Name = "frmEncontraProduto" Then
            Unload Forms(Cont)
            Exit For
        End If
    Next
    If Forms.Count = 2 Then
        AllBotões False
    End If
End Sub
Private Sub txtDescontoDePromoção_Change()
    FormatMask "@K 99,99", txtDescontoDePromoção
End Sub
Private Sub txtDescontoDePromoção_KeyPress(KeyAscii As Integer)
    If Not lInserir Then
        lAlterar = True
        StatusBarAviso = "Alteração"
        BarraDeStatus StatusBarAviso
    End If
End Sub
Private Sub txtDescontoDePromoção_LostFocus()
    FormatMask "@V #0,00", txtDescontoDePromoção
End Sub
Private Sub txtDescontoMáximo_Change()
    FormatMask "@K 99,99", txtDescontoMáximo
End Sub
Private Sub txtDescontoMáximo_KeyPress(KeyAscii As Integer)
    If Not lInserir Then
        lAlterar = True
        StatusBarAviso = "Alteração"
        BarraDeStatus StatusBarAviso
    End If
End Sub
Private Sub txtDescontoMáximo_LostFocus()
    FormatMask "@V #0,00", txtDescontoMáximo
End Sub
Private Sub txtDescrição_Change()
    If Not lPula Then
        FormatMask "@!S50", txtDescrição
    End If
End Sub
Private Sub txtDescrição_KeyPress(KeyAscii As Integer)
    If Not lInserir Then
        lAlterar = True
        StatusBarAviso = "Alteração"
        BarraDeStatus StatusBarAviso
    End If
End Sub
Private Sub txtInício_Change()
    If Not lPula Then
        lPula = True
        FormatMask DataMask, txtInício
        lPula = False
    End If
End Sub
Private Sub txtInício_KeyPress(KeyAscii As Integer)
    If Not lInserir Then
        lAlterar = True
        StatusBarAviso = "Alteração"
        BarraDeStatus StatusBarAviso
    End If
End Sub
Private Sub txtInício_LostFocus()
    If StrTran(txtInício.Text, "/") <> Space(8) Then
        lPula = True
        CorrigeData DataMask, txtInício, Date
        lPula = False
        If Not FormatMask(CheckDataMask, txtInício) Then
            Beep
            MsgBox "Data inválida !", vbCritical, "Erro"
            txtInício.SelStart = 0
            txtInício.SetFocus
        End If
    End If
End Sub
Private Sub txtIPI_Change()
    FormatMask "@K 99,99", txtIPI
End Sub
Private Sub txtIPI_KeyPress(KeyAscii As Integer)
    If Not lInserir Then
        lAlterar = True
        StatusBarAviso = "Alteração"
        BarraDeStatus StatusBarAviso
    End If
End Sub
Private Sub txtIPI_LostFocus()
    FormatMask "@V #0,00", txtIPI
End Sub
Private Sub txtPeso_Change()
    If Not lPula Then
        FormatMask "@K 999,999", txtPeso
    End If
End Sub
Private Sub txtPeso_KeyPress(KeyAscii As Integer)
    If Not lInserir Then
        lAlterar = True
        StatusBarAviso = "Alteração"
        BarraDeStatus StatusBarAviso
    End If
End Sub
Private Sub txtPeso_LostFocus()
    FormatMask "@V ##0,000", txtPeso
End Sub
Private Sub txtQuantidade_Change()
    If Not lPula Then
        FormatMask "@K 9999999,99", txtQuantidade
    End If
End Sub
Private Sub txtQuantidade_KeyPress(KeyAscii As Integer)
    If Not lInserir Then
        lAlterar = True
        StatusBarAviso = "Alteração"
        BarraDeStatus StatusBarAviso
    End If
End Sub
Private Sub txtQuantidade_LostFocus()
    FormatMask "@V ######0,00", txtQuantidade
End Sub
Private Sub txtTÉRMINO_Change()
    If Not lPula Then
        lPula = True
        FormatMask DataMask, txtTérmino
        lPula = False
    End If
End Sub
Private Sub txtTérmino_KeyPress(KeyAscii As Integer)
    If Not lInserir Then
        lAlterar = True
        StatusBarAviso = "Alteração"
        BarraDeStatus StatusBarAviso
    End If
End Sub
Private Sub txtTÉRMINO_LostFocus()
    If StrTran(txtTérmino.Text, "/") <> Space(8) Then
        lPula = True
        CorrigeData DataMask, txtTérmino, Date
        lPula = False
        If Not FormatMask(CheckDataMask, txtTérmino) Then
            Beep
            MsgBox "Data inválida !", vbCritical, "Erro"
            txtTérmino.SelStart = 0
            txtTérmino.SetFocus
        End If
    End If
End Sub
