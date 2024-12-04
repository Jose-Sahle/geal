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
      Begin VB.CommandButton cmdPre�os 
         Caption         =   "&Pre�os"
         Height          =   345
         Left            =   3420
         TabIndex        =   18
         Top             =   240
         Width           =   1245
      End
      Begin VB.CommandButton cmdC�digo 
         Caption         =   "C�&digos"
         Height          =   345
         Left            =   150
         TabIndex        =   17
         Top             =   240
         Width           =   1245
      End
   End
   Begin VB.Frame frDescontoPromo��o 
      Caption         =   " Promo��o/Desconto"
      Height          =   1230
      Left            =   0
      TabIndex        =   33
      Top             =   3810
      Width           =   8025
      Begin VB.VScrollBar vscrDescontoM�ximo 
         Height          =   315
         Left            =   2715
         TabIndex        =   43
         TabStop         =   0   'False
         Top             =   675
         Width           =   210
      End
      Begin VB.TextBox txtDescontoM�ximo 
         Height          =   285
         Left            =   1950
         TabIndex        =   16
         Text            =   " 0,00"
         Top             =   690
         Width           =   765
      End
      Begin VB.TextBox txtT�rmino 
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
      Begin VB.TextBox txtIn�cio 
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
      Begin VB.VScrollBar vscrDescontoDePromo��o 
         Height          =   315
         Left            =   2715
         TabIndex        =   38
         TabStop         =   0   'False
         Top             =   285
         Width           =   210
      End
      Begin VB.TextBox txtDescontoDePromo��o 
         Height          =   285
         Left            =   1950
         TabIndex        =   13
         Text            =   " 0,00"
         Top             =   300
         Width           =   765
      End
      Begin VB.Label lblDescontoM�ximoPorcentagem 
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
      Begin VB.Label lblDescontoM�ximo 
         Caption         =   "Desconto M�ximo"
         Height          =   195
         Left            =   150
         TabIndex        =   42
         Top             =   720
         Width           =   1380
      End
      Begin VB.Label lblT�rmino 
         Caption         =   "T�rmino"
         Height          =   195
         Left            =   5850
         TabIndex        =   41
         Top             =   330
         Width           =   570
      End
      Begin VB.Label lblIn�cio 
         Caption         =   "In�cio"
         Height          =   225
         Left            =   3600
         TabIndex        =   40
         Top             =   330
         Width           =   495
      End
      Begin VB.Label lblDescontoDePromo��oPorcentagem 
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
      Begin VB.Label lblDescontoDePromo��o 
         Caption         =   "Desconto de Promo��o"
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
   Begin VB.Frame frEspecifica��es 
      Caption         =   " Especifica��es "
      Height          =   2655
      Left            =   0
      TabIndex        =   22
      Top             =   0
      Width           =   8025
      Begin VB.CheckBox chkLotes 
         Caption         =   "Divis�o por Lotes"
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
      Begin VB.ComboBox cmbSe��o 
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
      Begin VB.TextBox txtDescri��o 
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
      Begin VB.Label lblSe��o 
         Caption         =   "Se��o"
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
      Begin VB.Label lblDepartamentoSe��o 
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
      Begin VB.Label lblDescri��o 
         Caption         =   "Descri��o"
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

Dim C�digo  As Long
Dim mC�digo As Long

Dim lAllowInsert     As Boolean
Dim lAllowEdit       As Boolean
Dim lAllowDelete     As Boolean
Dim lAllowConsult    As Boolean
Dim lAllowEditAmount As Boolean

Dim TBLProduto    As Table
Dim ProdutoAberto As Boolean
Dim IndiceProdutoAtivo$

Dim TBLC�digoDoProduto As Table
Dim C�digoDoProdutoAberto As Boolean
Dim IndiceC�digoDoProdutoAtivo$

Dim TBLPre�o As Table
Dim Pre�oAberto As Boolean
Dim IndicePre�oAtivo$

Dim TBLLote As Table
Dim LoteAberto As Boolean
Dim IndiceLoteAtivo$

Dim TBLDepartamento As Table
Dim DepartamentoAberto As Boolean
Dim IndiceDepartamentoAtivo$

Dim TBLSe��o As Table
Dim Se��oAberto As Boolean
Dim IndiceSe��oAtivo$

Dim TBLDepartamentoSe��o As Table
Dim DepartamentoSe��oAberto As Boolean
Dim IndiceDepartamentoSe��oAtivo$

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

Dim TBLPar�metros As Table
Dim Par�metrosAberto As Boolean

Dim ArrayDepartamento() As String
Dim ArraySe��o() As String
Dim ArrayTipoDeICM() As String
Dim ArrayTipoDeEmbalagem() As String
Dim ArrayUnidades() As String
Dim ArrayLocal() As String

'Matriz para a chamada da janela de inclus�o do c�digo do produto
Public ArrayProdutoTotal%
Dim ArrayProdutoFornecedor() As Variant
Dim ArrayProdutoC�digo() As Variant

'Matriz para a chamada da janela de inclus�o de lotes
Public ArrayLoteTotal%
Dim ArrayLote() As Variant

'Matriz para a chamada da janela de inclus�o de pre�o
Public ArrayPre�oTotal%
Dim ArrayPre�oFornecedor() As Variant
Dim ArrayPre�oCusto() As Variant
Dim ArrayPre�oVenda() As Variant
Dim ArrayPre�oLucro() As Variant

Public lInserir As Boolean
Public lAlterar As Boolean

Dim mFechar As Boolean

Public lAlterarArrayProduto As Boolean
Public lAlterarArrayLote As Boolean
Public lAlterarArrayPre�o As Boolean

Public StatusBarAviso$

Dim lPula As Boolean

Dim DataBaseName(1 To 1) As String
Public Relat�rio$
Public TotalDatabaseName%

Public m�ltimoD�gito As Byte

Public lAtualizar As Boolean
Private Sub AtivaCampos()
    Bot�oImprimir True
    frEspecifica��es.Enabled = True
    frImpostos.Enabled = True
    frDescontoPromo��o.Enabled = True
    frVariados.Enabled = True
    Bot�oGravar (lInserir Or lAllowEdit)
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
    FillSe��o TBLDepartamento("C�DIGO")
    FillTipoDeEmbalagem
    FillLocal
    FillUnidades
    FillTipoDeICM
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
    
    Bot�oIncluir lAllowInsert
    ClearArrayProduto
    ClearArrayPre�o
    ClearArrayLote
    
    If TBLProduto.RecordCount = 0 Then
        Navega��oInferior False
        Navega��oSuperior False
        Bot�oGravar False
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
    frmLotes.lAltera��o = False
    
    ArrayLoteTotal = 0
    ReDim ArrayLote(MAXCOL, 1)
End Sub
Public Sub ClearArrayPre�o()
    lAlterarArrayPre�o = False
    frmPre�os.lAltera��o = False
    
    ArrayPre�oTotal = 0
    ReDim ArrayPre�oFornecedor(1 To 1)
    ReDim ArrayPre�oCusto(1 To 1)
    ReDim ArrayPre�oVenda(1 To 1)
    ReDim ArrayPre�oLucro(1 To 1)
End Sub
Public Sub ClearArrayProduto()
    lAlterarArrayProduto = False
    frmC�digoDoProduto.lAltera��o = False
    
    ArrayProdutoTotal = 0
    ReDim ArrayProdutoFornecedor(1 To 1)
    ReDim ArrayProdutoC�digo(1 To 1)
End Sub
Private Sub DesativaCampos()
    Bot�oImprimir False
    frEspecifica��es.Enabled = False
    frImpostos.Enabled = False
    frDescontoPromo��o.Enabled = False
    frVariados.Enabled = False
    cmdGravar.Enabled = False
    cmdCancelar.Enabled = False
    Bot�oGravar False
End Sub
Public Sub Encontrar()
    If Not lAllowConsult Then
        Exit Sub
    End If
    If lInserir Then
        MsgBox "Voc� est� em uma inclus�o!", vbExclamation, Caption
        StatusBarAviso = "Finalize a inclus�o"
        Exit Sub
    End If
    If lAlterar Then
        MsgBox "Voc� est� em uma altera��o!", vbExclamation, Caption
        StatusBarAviso = "Finalize a altera��o"
        Exit Sub
    End If
    
    Set frmEncontraProduto.Janela = Me
    frmEncontraProduto.NoModal = True
    frmEncontraProduto.Show 0
End Sub
Public Sub Excluir()
    Dim Confirma��o As Integer, Msg1$, Msg2$, C�digoDoProduto As Variant
    Dim SQL As String

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
    
    C�digoDoProduto = TBLProduto("C�DIGO")
    TBLProduto.Delete
    
    SQL = "Delete * From [C�DIGO DO PRODUTO] Where [C�DIGO DO PRODUTO]= " + Str(C�digoDoProduto)
    DBCadastro.Execute SQL
    
    SQL = "Delete * From [PRE�O DO PRODUTO] Where [C�DIGO DO PRODUTO]= " + Str(C�digoDoProduto)
    DBCadastro.Execute SQL
    
    SQL = "Delete * From [LOTE DO PRODUTO] Where [C�DIGO DO PRODUTO]= " + Str(C�digoDoProduto)
    DBCadastro.Execute SQL
    
    If Err <> 0 Then
        GeraMensagemDeErro "Produto - Excluir - " & txtDescri��o, True
        StatusBarAviso = "Falha na exclus�o"
        BarraDeStatus StatusBarAviso
        Exit Sub
    End If
    
    Log gUsu�rio, "Exclus�o - Produto: " & txtDescri��o
    
    WS.CommitTrans
        
    ClearArrayProduto
    ClearArrayPre�o
    ClearArrayLote
    
    StatusBarAviso = "Exclus�o bem sucedida"
    BarraDeStatus StatusBarAviso
    
    If TBLProduto.RecordCount = 0 Then
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
    TBLLote.Seek "=", C�digo
    
    If TBLLote.NoMatch Then
        Exit Sub
    End If
    
    ArrayLoteTotal = 0
    
    Do While Not TBLLote.EOF
        If TBLLote("C�DIGO DO PRODUTO") <> TBLProduto("C�DIGO") Then
            Exit Do
        End If
        
        ArrayLoteTotal = ArrayLoteTotal + 1
        
        SizeArrayLote ArrayLoteTotal
        
        ArrayLote(1, ArrayLoteTotal) = TBLLote("C�DIGO DO LOTE")
        ArrayLote(2, ArrayLoteTotal) = TBLLote("D�GITO DO LOTE")
        ArrayLote(3, ArrayLoteTotal) = TBLLote("M�LTIPLO")
        ArrayLote(4, ArrayLoteTotal) = TBLLote("QUANTIDADE")
        
        TBLLote.MoveNext
    Loop
End Sub
Private Sub FillArrayPre�o()
    Dim Aux$, Posi��o%
    
    TBLPre�o.Seek "=", C�digo
    
    If TBLPre�o.NoMatch Then
        Exit Sub
    End If
    
    ArrayPre�oTotal = 0
    
    Do While Not TBLPre�o.EOF
        If TBLPre�o("C�DIGO DO PRODUTO") <> TBLProduto("C�DIGO") Then
            Exit Do
        End If
    
        ArrayPre�oTotal = ArrayPre�oTotal + 1
        
        SizeArrayPre�o ArrayPre�oTotal
        
        ArrayPre�oFornecedor(ArrayPre�oTotal) = TBLPre�o("C�DIGO DO FORNECEDOR")
        ArrayPre�oCusto(ArrayPre�oTotal) = FormatStringMask("@V ##.###.##0,00", StrVal(TBLPre�o("PRE�O DE CUSTO")))
        ArrayPre�oVenda(ArrayPre�oTotal) = FormatStringMask("@V ##.###.##0,00", StrVal(TBLPre�o("PRE�O DE VENDA")))
        ArrayPre�oLucro(ArrayPre�oTotal) = FormatStringMask("@V ##0,00", StrVal(TBLPre�o("MARGEM DE LUCRO")))
        
        TBLPre�o.MoveNext
    Loop
End Sub
Private Sub FillArrayProduto()
    TBLC�digoDoProduto.Seek "=", C�digo
    
    If TBLC�digoDoProduto.NoMatch Then
        Exit Sub
    End If
    
    ArrayProdutoTotal = 0
    
    Do While Not TBLC�digoDoProduto.EOF
        If TBLC�digoDoProduto("C�DIGO DO PRODUTO") <> TBLProduto("C�DIGO") Then
            Exit Do
        End If
        
        ArrayProdutoTotal = ArrayProdutoTotal + 1
        
        SizeArrayProduto ArrayProdutoTotal
        
            
        ArrayProdutoFornecedor(ArrayProdutoTotal) = TBLC�digoDoProduto("FORNECEDOR")
        ArrayProdutoC�digo(ArrayProdutoTotal) = TBLC�digoDoProduto("C�DIGO DO FORNECEDOR")
        
        TBLC�digoDoProduto.MoveNext
    Loop
End Sub
Private Sub FillDepartamento()
    Dim Cont%
    
    ReDim ArrayDepartamento(0 To TBLDepartamento.RecordCount - 1)
    
    Cont = 0
    
    cmbDepartamento.Clear
    
    TBLDepartamento.MoveFirst
    
    Do While Not TBLDepartamento.EOF
        cmbDepartamento.AddItem TBLDepartamento("DESCRI��O")
        ArrayDepartamento(Cont) = TBLDepartamento("C�DIGO")
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
        cmbLocal.AddItem TBLLocal("ENDERE�O")
        ArrayLocal(Cont) = TBLLocal("C�DIGO")
        Cont = Cont + 1
        TBLLocal.MoveNext
    Loop
    cmbLocal.ListIndex = 0
End Sub
Private Sub FillSe��o(ByVal C�digo)
    Dim SQL$, TBLAux As Recordset, Cont%
    
    SQL = "SELECT SE��O.C�DIGO,SE��O.DESCRI��O FROM (DEPARTAMENTO INNER JOIN [DEPARTAMENTO - SE��O] ON DEPARTAMENTO.C�DIGO = [DEPARTAMENTO - SE��O].[C�DIGO DO DEPTO]) INNER JOIN SE��O ON [DEPARTAMENTO - SE��O].[C�DIGO DA SE��O] = SE��O.C�DIGO Where (([DEPARTAMENTO - SE��O].[C�DIGO DO DEPTO] = '" + C�digo + "') And ([Se��o].[C�digo] = [DEPARTAMENTO - SE��O].[C�DIGO DA SE��O])) ORDER BY SE��O.DESCRI��O"
    
    Set TBLAux = DBCadastro.OpenRecordset(SQL)
    
    cmbSe��o.Clear
    
    If TBLAux.RecordCount = 0 Then
        Exit Sub
    End If
    
    ReDim ArraySe��o(0 To TBLAux.RecordCount - 1)
    
    Cont = 0
    
    TBLAux.MoveFirst
    
    Do While Not TBLAux.EOF
        cmbSe��o.AddItem TBLAux("DESCRI��O")
        ArraySe��o(Cont) = TBLAux("C�DIGO")
        Cont = Cont + 1
        TBLAux.MoveNext
    Loop
    
    cmbSe��o.ListIndex = 0
End Sub
Private Sub FillTipoDeEmbalagem()
    Dim Cont%
    
    ReDim ArrayTipoDeEmbalagem(0 To TBLTipoDeEmbalagem.RecordCount - 1)
    
    Cont = 0
    
    cmbTipoDeEmbalagem.Clear
    
    TBLTipoDeEmbalagem.MoveFirst
    
    Do While Not TBLTipoDeEmbalagem.EOF
        cmbTipoDeEmbalagem.AddItem TBLTipoDeEmbalagem("DESCRI��O")
        ArrayTipoDeEmbalagem(Cont) = TBLTipoDeEmbalagem("C�DIGO")
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
        cmbTipoDeICM.AddItem TBLTipoDeICM("DESCRI��O")
        ArrayTipoDeICM(1, Cont) = TBLTipoDeICM("C�DIGO")
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
        cmbUnidades.AddItem TBLUnidades("DESCRI��O")
        ArrayUnidades(Cont) = TBLUnidades("C�DIGO")
        Cont = Cont + 1
        TBLUnidades.MoveNext
    Loop
    cmbUnidades.ListIndex = 0
End Sub
Public Function GetArrayLote(ByVal Item As Byte, ByVal Elemento As Integer) As String
    GetArrayLote = ArrayLote(Item, Elemento)
End Function
Public Function GetArrayPre�o(ByVal Nome As String, ByVal Elemento As Integer) As String
    If Nome = "Fornecedor" Then
        GetArrayPre�o = ArrayPre�oFornecedor(Elemento)
    ElseIf Nome = "Custo" Then
        GetArrayPre�o = ArrayPre�oCusto(Elemento)
    ElseIf Nome = "Venda" Then
        GetArrayPre�o = ArrayPre�oVenda(Elemento)
    ElseIf Nome = "Lucro" Then
        GetArrayPre�o = ArrayPre�oLucro(Elemento)
    End If
End Function
Public Function GetArrayProduto(ByVal Nome As String, ByVal Elemento As Integer) As String
    If Nome = "Fornecedor" Then
        GetArrayProduto = ArrayProdutoFornecedor(Elemento)
    ElseIf Nome = "C�digo" Then
        GetArrayProduto = ArrayProdutoC�digo(Elemento)
    End If
End Function
Public Sub Gravar()
    If lInserir Then
        
        'Pega o novo c�digo interno do produto e atualiza na Tabela Par�metros
        TBLPar�metros.Edit
        mC�digo = TBLPar�metros("PRODUTO") + 1
        TBLPar�metros("PRODUTO") = mC�digo
        TBLPar�metros.Update
        
        If SetRecords Then
            PosRecords
            lInserir = False
            ClearArrayProduto
            ClearArrayPre�o
            ClearArrayLote
            StatusBarAviso = "Inclus�o bem sucedida"
        Else
            StatusBarAviso = "Falha na inclus�o"
        End If
    ElseIf lAlterar Then
        If TBLProduto.RecordCount > 0 And Not TBLProduto.BOF And Not TBLProduto.EOF Then
            mC�digo = TBLProduto("C�DIGO")
            If SetRecords Then
                PosRecords
                lAlterar = False
                ClearArrayProduto
                ClearArrayPre�o
                ClearArrayLote
                StatusBarAviso = "Altera��o bem sucedida"
            Else
                StatusBarAviso = "Falha na altera��o"
            End If
        End If
    End If
    
    BarraDeStatus StatusBarAviso
    
    TestaInferior TBLProduto, lAllowEdit, lAllowDelete, lAllowConsult
    TestaSuperior TBLProduto, lAllowEdit, lAllowDelete, lAllowConsult
    
    If TBLProduto.RecordCount = 0 Then
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
    
    If txtDescri��o.Enabled Then
        txtDescri��o.SetFocus
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
    
    txtDescri��o.SetFocus
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
    ClearArrayPre�o
    ClearArrayLote
    
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
    
    TBLProduto.MoveLast
    
    ClearArrayProduto
    ClearArrayPre�o
    ClearArrayLote
    
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
    
    TBLProduto.MoveNext
    
    If TBLProduto.EOF Then
        TBLProduto.MovePrevious
        Exit Sub
    End If
    
    ClearArrayProduto
    ClearArrayPre�o
    ClearArrayLote
    
    StatusBarAviso = "Pronto"
    BarraDeStatus StatusBarAviso
    
    Navega��oInferior lAllowConsult
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
    ClearArrayPre�o
    ClearArrayLote
    
    StatusBarAviso = "Pronto"
    BarraDeStatus StatusBarAviso
    
    Navega��oSuperior lAllowConsult
    TestaInferior TBLProduto, lAllowEdit, lAllowDelete, lAllowConsult
    
    GetRecords
End Sub
Public Sub Posicionar(ByVal C�digo)
    mC�digo = Val(C�digo)
    PosRecords
    SetFocus
End Sub
Public Sub PosRecords()
    If TBLProduto.RecordCount = 0 Then
            Exit Sub
    End If

    TBLProduto.Seek "=", mC�digo
    If TBLProduto.NoMatch Then
        'MsgBox "N�o consegui encontrar o cliente com CGC/CPF " + txtCgcCpf, vbExclamation, "Erro"
        TBLProduto.MoveFirst
        Navega��oInferior False
        Navega��oInferior lAllowConsult
    Else
        TestaInferior TBLProduto, lAllowEdit, lAllowDelete, lAllowConsult
        TestaSuperior TBLProduto, lAllowEdit, lAllowDelete, lAllowConsult
    End If
    GetRecords
End Sub
Public Function PushDataBaseName(ByVal Posi��o As Integer) As String
    PushDataBaseName = DataBaseName(Posi��o)
End Function
Private Sub GetRecords()
    On Error GoTo Erro
    
    Dim Aux$, Posi��o%, AuxBookMark, Cont As Integer, AuxTotal As Single
    
    lPula = True
    
    If Not lAllowConsult Then
        ZeraCampos
        DesativaCampos
        lPula = False
        Exit Sub
    End If
    
    C�digo = TBLProduto("C�DIGO")
    txtDescri��o = TBLProduto("DESCRI��O")
    
    Aux = Trim(StrVal(TBLProduto("PESO")))
    txtPeso = FormatStringMask("@V ##0,000", Aux)
    
    TBLDepartamento.Index = "DEPARTAMENTO1"
    TBLDepartamento.Seek "=", Mid(TBLProduto("DEPTO - SE��O"), 1, 4)
    cmbDepartamento.Text = TBLDepartamento("DESCRI��O")
    cmbDepartamento_LostFocus
    AuxBookMark = TBLDepartamento.Bookmark
    TBLDepartamento.Index = IndiceDepartamentoAtivo
    TBLDepartamento.Bookmark = AuxBookMark
    FillSe��o TBLDepartamento("C�DIGO")
        
    TBLSe��o.Index = "SE��O1"
    TBLSe��o.Seek "=", Mid(TBLProduto("DEPTO - SE��O"), 5, 4)
    cmbSe��o.Text = TBLSe��o("DESCRI��O")
    cmbSe��o_LostFocus
    AuxBookMark = TBLSe��o.Bookmark
    TBLSe��o.Index = IndiceSe��oAtivo
    TBLSe��o.Bookmark = AuxBookMark
    
    TBLUnidades.Index = "UNIDADES1"
    TBLUnidades.Seek "=", TBLProduto("UNIDADES")
    cmbUnidades.Text = TBLUnidades("DESCRI��O")
    cmbUnidades_LostFocus
    AuxBookMark = TBLUnidades.Bookmark
    TBLUnidades.Index = IndiceUnidadesAtivo
    TBLUnidades.Bookmark = AuxBookMark
    
    TBLTipoDeEmbalagem.Index = "TIPODEEMBALAGEM1"
    TBLTipoDeEmbalagem.Seek "=", TBLProduto("TIPO DE EMBALAGEM")
    cmbTipoDeEmbalagem.Text = TBLTipoDeEmbalagem("DESCRI��O")
    cmbTipoDeEmbalagem_LostFocus
    AuxBookMark = TBLTipoDeEmbalagem.Bookmark
    TBLTipoDeEmbalagem.Index = IndiceTipoDeEmbalagemAtivo
    TBLTipoDeEmbalagem.Bookmark = AuxBookMark
    
    TBLTipoDeICM.Index = "TIPODEICM1"
    TBLTipoDeICM.Seek "=", TBLProduto("TIPO DE ICM")
    cmbTipoDeICM.Text = TBLTipoDeICM("DESCRI��O")
    cmbTipoDeICM_LostFocus
    AuxBookMark = TBLTipoDeICM.Bookmark
    TBLTipoDeICM.Index = IndiceTipoDeICMAtivo
    TBLTipoDeICM.Bookmark = AuxBookMark
    
    TBLLocal.Index = "LOCALDOPRODUTO1"
    TBLLocal.Seek "=", TBLProduto("LOCAL")
    If TBLLocal.NoMatch Then
        TBLLocal.MoveFirst
    End If
    cmbLocal.Text = TBLLocal("ENDERE�O")
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
        m�ltimoD�gito = 0
        For Cont = 1 To ArrayLoteTotal
            AuxTotal = AuxTotal + ArrayLote(3, Cont) * ArrayLote(4, Cont)
            If m�ltimoD�gito < ArrayLote(2, Cont) Then
                m�ltimoD�gito = ArrayLote(2, Cont)
            End If
        Next
        txtQuantidade = FormatStringMask("@V #####0,00", StrVal(AuxTotal))
    Else
        txtQuantidade.Enabled = True
        vscrQuantidade.Enabled = True
        cmdLotes.Enabled = False
        m�ltimoD�gito = 0
    End If
    
    Aux = Trim(StrVal(TBLProduto("ICM")))
    txtICM = FormatStringMask("@V #0,00", Aux)
    
    Aux = Trim(StrVal(TBLProduto("IPI")))
    txtIPI = FormatStringMask("@V #0,00", Aux)
    
    Aux = Trim(StrVal(TBLProduto("DESCONTO DE PROMO��O")))
    txtDescontoDePromo��o = FormatStringMask("@V #0,00", Aux)
    
    If TBLProduto("IN�CIO") <> vbNull Then
        txtIn�cio = FormatStringMask(CheckDataMask, TBLProduto("IN�CIO"))
        CorrigeData DataMask, txtIn�cio, TBLProduto("IN�CIO")
    Else
        txtIn�cio = DataNula
    End If
        
    If TBLProduto("T�RMINO") <> vbNull Then
        txtT�rmino = FormatStringMask(CheckDataMask, TBLProduto("T�RMINO"))
        CorrigeData DataMask, txtT�rmino, TBLProduto("T�RMINO")
    Else
        txtT�rmino = DataNula
    End If
    
    Aux = Trim(StrVal(TBLProduto("DESCONTO M�XIMO")))
    txtDescontoM�ximo = FormatStringMask("@V #0,00", Aux)
    
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
    Dim Confirma��o As Integer, Msg1$, Msg2$
    Dim SQL As String
    Dim Cont%, Recno%
    
    WS.BeginTrans 'Inicia uma Transa��o
    
    If lInserir Then
        TBLProduto.AddNew
    Else
        TBLProduto.Edit
    End If
    
    TBLProduto("C�DIGO") = mC�digo
    TBLProduto("DESCRI��O") = txtDescri��o
    TBLProduto("PESO") = ValStr(txtPeso) ' 0,00
    TBLProduto("QUANTIDADE") = ValStr(txtQuantidade) '0,00
    TBLProduto("LOTES") = IIf(chkLotes.Value = 1, True, False)
    TBLProduto("QUANTIDADE DE LOTES") = Val(txtQuantidadeDeLotes) '0
    TBLProduto("IPI") = ValStr(txtIPI) '0,00
    TBLProduto("ICM") = ValStr(ArrayTipoDeICM(2, cmbTipoDeICM.ListIndex)) '0,00
    TBLProduto("DESCONTO DE PROMO��O") = ValStr(txtDescontoDePromo��o) '0,00
    TBLProduto("IN�CIO") = IIf(Trim(StrTran(txtIn�cio, "/")) <> Empty, txtIn�cio, vbNull)
    TBLProduto("T�RMINO") = IIf(Trim(StrTran(txtT�rmino, "/")) <> Empty, txtT�rmino, vbNull)
    TBLProduto("DESCONTO M�XIMO") = ValStr(txtDescontoM�ximo) '0,00
    TBLProduto("TIPO DE EMBALAGEM") = ArrayTipoDeEmbalagem(cmbTipoDeEmbalagem.ListIndex)
    TBLProduto("TIPO DE ICM") = ArrayTipoDeICM(1, cmbTipoDeICM.ListIndex)
    TBLProduto("DEPTO - SE��O") = ArrayDepartamento(cmbDepartamento.ListIndex) + ArraySe��o(cmbSe��o.ListIndex)
    TBLProduto("UNIDADES") = ArrayUnidades(cmbUnidades.ListIndex)
    TBLProduto("LOCAL") = ArrayLocal(cmbLocal.ListIndex)
    If lInserir Then
        TBLProduto("USERNAME - CRIA") = gUsu�rio
        TBLProduto("DATA - CRIA") = Date
        TBLProduto("HORA - CRIA") = Time
        TBLProduto("USERNAME - ALTERA") = "VAZIO"
        TBLProduto("DATA - ALTERA") = vbNull
        TBLProduto("HORA - ALTERA") = vbNull
    End If
    If lAlterar Then
        TBLProduto("USERNAME - ALTERA") = gUsu�rio
        TBLProduto("DATA - ALTERA") = Date
        TBLProduto("HORA - ALTERA") = Time
    End If
    TBLProduto.Update
    
ErroProd:
    If Err <> 0 Then
        TBLProduto.CancelUpdate
        GeraMensagemDeErro "Produto - SetRecords - ErroProd - " & txtDescri��o, True
        SetRecords = False
        Exit Function
    End If
        
    On Error GoTo ErroC�d
    
    If lAlterarArrayProduto Then
        SQL = "Delete * From [C�DIGO DO PRODUTO] Where [C�DIGO DO PRODUTO]= " & mC�digo
        DBCadastro.Execute SQL
        
        For Cont = 1 To ArrayProdutoTotal
            TBLC�digoDoProduto.AddNew

            TBLC�digoDoProduto("C�DIGO DO PRODUTO") = mC�digo
            TBLC�digoDoProduto("FORNECEDOR") = ArrayProdutoFornecedor(Cont)
            TBLC�digoDoProduto("C�DIGO DO FORNECEDOR") = ArrayProdutoC�digo(Cont)
            TBLC�digoDoProduto("USERNAME - CRIA") = gUsu�rio
            TBLC�digoDoProduto("DATA - CRIA") = Date
            TBLC�digoDoProduto("HORA - CRIA") = Time
            TBLC�digoDoProduto("USERNAME - ALTERA") = "VAZIO"
            TBLC�digoDoProduto("DATA - ALTERA") = vbNull
            TBLC�digoDoProduto("HORA - ALTERA") = vbNull
            TBLC�digoDoProduto.Update
        Next
    End If
    
ErroC�d:
    If Err <> 0 Then
        TBLC�digoDoProduto.CancelUpdate
        GeraMensagemDeErro "Produto - SetRecords - ErroC�d - " & txtDescri��o, True
        SetRecords = False
        Exit Function
    End If
    
    On Error GoTo ErroPre�o
    
    If lAlterarArrayPre�o Then
        SQL = "Delete * From [PRE�O DO PRODUTO] Where [C�DIGO DO PRODUTO]= " & mC�digo
        DBCadastro.Execute SQL
    
        For Cont = 1 To ArrayPre�oTotal
            TBLPre�o.AddNew
            TBLPre�o("C�DIGO DO PRODUTO") = mC�digo
            TBLPre�o("C�DIGO DO FORNECEDOR") = ArrayPre�oFornecedor(Cont)
            TBLPre�o("PRE�O DE CUSTO") = ValStr(ArrayPre�oCusto(Cont))
            TBLPre�o("PRE�O DE VENDA") = ValStr(ArrayPre�oVenda(Cont))
            TBLPre�o("MARGEM DE LUCRO") = ValStr(ArrayPre�oLucro(Cont))
            TBLPre�o("USERNAME - CRIA") = gUsu�rio
            TBLPre�o("DATA - CRIA") = Date
            TBLPre�o("HORA - CRIA") = Time
            TBLPre�o("USERNAME - ALTERA") = "VAZIO"
            TBLPre�o("DATA - ALTERA") = vbNull
            TBLPre�o("HORA - ALTERA") = vbNull
            TBLPre�o.Update
        Next
    End If
    
ErroPre�o:
    If Err <> 0 Then
        TBLPre�o.CancelUpdate
        GeraMensagemDeErro "Produto - SetRecords - ErroPre�o - " & txtDescri��o, True
        SetRecords = False
        Exit Function
    End If
    
    On Error GoTo ErroLote
    
    If lAlterarArrayLote Then
        SQL = "Delete * From [LOTE DO PRODUTO] Where [C�DIGO DO PRODUTO]= " & mC�digo
        DBCadastro.Execute SQL
        For Cont = 1 To ArrayLoteTotal
            TBLLote.AddNew
            TBLLote("C�DIGO DO LOTE") = ArrayLote(1, Cont)
            TBLLote("D�GITO DO LOTE") = ArrayLote(2, Cont)
            TBLLote("M�LTIPLO") = ValStr(ArrayLote(3, Cont))
            TBLLote("QUANTIDADE") = ValStr(ArrayLote(4, Cont))
            TBLLote("C�DIGO DO PRODUTO") = TBLProduto("C�DIGO")
            TBLLote("USERNAME - CRIA") = gUsu�rio
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
        GeraMensagemDeErro "Produto - SetRecords - ErroLote - " & txtDescri��o, True
        SetRecords = False
        Exit Function
    End If

    WS.CommitTrans 'Grava as altera��es ou inclus�es se n�o houverem erros
        
    If lInserir Then
        Log gUsu�rio, "Inclus�o - Produto " & txtDescri��o
    Else
        Log gUsu�rio, "Altera��o - Produto " & txtDescri��o
    End If
    
    lAlterar = False
    lInserir = False
    ClearArrayLote
    ClearArrayProduto
    ClearArrayPre�o
    
    C�digo = TBLProduto("C�DIGO")
    
    SetRecords = True
End Function
Public Sub SetArrayLote(ByVal Item As Byte, ByVal Valor As String, ByVal Elemento As Integer)
    ArrayLote(Item, Elemento) = Valor
End Sub
Public Sub SetArrayPre�o(ByVal Nome As String, ByVal Valor As String, ByVal Elemento As Integer)
    If Nome = "Fornecedor" Then
        ArrayPre�oFornecedor(Elemento) = Valor
    ElseIf Nome = "Custo" Then
        ArrayPre�oCusto(Elemento) = Valor
    ElseIf Nome = "Venda" Then
        ArrayPre�oVenda(Elemento) = Valor
    ElseIf Nome = "Lucro" Then
        ArrayPre�oLucro(Elemento) = Valor
    End If
End Sub
Public Sub SetArrayProduto(ByVal Nome As String, ByVal Valor As String, ByVal Elemento As Integer)
    If Nome = "Fornecedor" Then
        ArrayProdutoFornecedor(Elemento) = Valor
    ElseIf Nome = "C�digo" Then
        ArrayProdutoC�digo(Elemento) = Valor
    End If
End Sub
Public Sub SizeArrayLote(ByVal Tamanho As Integer)
    If Tamanho > 0 Then
        ArrayLoteTotal = Tamanho
        ReDim Preserve ArrayLote(MAXCOL, Tamanho)
    End If
End Sub
Public Sub SizeArrayPre�o(ByVal Tamanho As Integer)
    ArrayPre�oTotal = Tamanho
    ASize Tamanho, ArrayPre�oFornecedor
    ASize Tamanho, ArrayPre�oCusto
    ASize Tamanho, ArrayPre�oVenda
    ASize Tamanho, ArrayPre�oLucro
End Sub
Public Sub SizeArrayProduto(ByVal Tamanho As Integer)
    ArrayProdutoTotal = Tamanho
    ASize Tamanho, ArrayProdutoFornecedor
    ASize Tamanho, ArrayProdutoC�digo
End Sub
Private Sub ZeraCampos()
    C�digo = Empty
    txtDescri��o = Empty
    txtPeso = " 0,000"
    txtPeso_LostFocus
    txtQuantidade = "0,00"
    txtQuantidade.Locked = False
    txtQuantidadeDeLotes = "0"
    txtIPI = " 0,00"
    txtICM = " 0,00"
    txtDescontoDePromo��o = " 0,00"
    txtIn�cio = Empty
    txtT�rmino = Empty
    txtDescontoM�ximo = " 0,00"
    
    ArrayProdutoTotal = 0
    ArrayLoteTotal = 0
    ArrayPre�oTotal = 0
    
    ClearArrayProduto
    ClearArrayPre�o
    ClearArrayLote
End Sub
Private Sub chkLotes_Click()
    If lPula Then
        Exit Sub
    End If
    If Not lInserir Then
        lAlterar = True
        StatusBarAviso = "Altera��o"
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
        StatusBarAviso = "Altera��o"
        BarraDeStatus StatusBarAviso
    End If
End Sub
Private Sub cmbDepartamento_Click()
    If lPula Then
        Exit Sub
    End If
    TBLDepartamento.MoveFirst
    TBLDepartamento.Move cmbDepartamento.ListIndex
    FillSe��o TBLDepartamento("C�DIGO")
    If Not lInserir Then
        lAlterar = True
        StatusBarAviso = "Altera��o"
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
        StatusBarAviso = "Altera��o"
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
Private Sub cmbSe��o_Click()
    If lPula Then
        Exit Sub
    End If
    If Not lInserir Then
        lAlterar = True
        StatusBarAviso = "Altera��o"
        BarraDeStatus StatusBarAviso
    End If
End Sub
Private Sub cmbSe��o_LostFocus()
    Dim Cont%, Encontrou As Boolean
    
    Encontrou = False
    
    For Cont = 0 To cmbSe��o.ListCount - 1
        If UCase(cmbSe��o.List(Cont)) = UCase(cmbSe��o.Text) Then
            Encontrou = True
            cmbSe��o.ListIndex = Cont
            Exit Sub
        End If
    Next
    
    For Cont = 0 To cmbSe��o.ListCount - 1
        If InStr(UCase(cmbSe��o.List(Cont)), UCase(cmbSe��o.Text)) = 1 Then
            Encontrou = True
            cmbSe��o.ListIndex = Cont
            Exit Sub
        End If
    Next
    
    If Not Encontrou Then
        cmbSe��o.ListIndex = 0
    End If
End Sub
Private Sub cmbTipoDeEmbalagem_Click()
    If lPula Then
        Exit Sub
    End If
    If Not lInserir Then
        lAlterar = True
        StatusBarAviso = "Altera��o"
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
        StatusBarAviso = "Altera��o"
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
        StatusBarAviso = "Altera��o"
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
Private Sub cmdC�digo_Click()
    If Not lInserir Then
        StatusBarAviso = "Altera��o"
        BarraDeStatus StatusBarAviso
    End If
    If Not lAlterarArrayProduto Then
        FillArrayProduto
    End If
    frmC�digoDoProduto.Show 0
End Sub
Private Sub cmdGravar_Click()
    Gravar
End Sub
Private Sub cmdLotes_Click()
    If Not lInserir Then
        StatusBarAviso = "Altera��o"
        BarraDeStatus StatusBarAviso
    End If
    If Not lAlterarArrayLote Then
        FillArrayLote
    End If
    Set frmLotes.mJanela = Me
    frmLotes.m�ltimoD�gito = m�ltimoD�gito
    frmLotes.Show 0
End Sub
Private Sub cmdPre�os_Click()
    If Not lInserir Then
        StatusBarAviso = "Altera��o"
        BarraDeStatus StatusBarAviso
    End If
    If Not lAlterarArrayPre�o Then
        FillArrayPre�o
    End If
    frmPre�os.Show 0
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
    If Not C�digoDoProdutoAberto Then
        Unload Me
        Exit Sub
    End If
    If Not LoteAberto Then
        Unload Me
        Exit Sub
    End If
    If Not Pre�oAberto Then
        Unload Me
        Exit Sub
    End If
    If Not DepartamentoAberto Then
        Unload Me
        Exit Sub
    End If
    If Not Se��oAberto Then
        Unload Me
        Exit Sub
    End If
    If Not DepartamentoSe��oAberto Then
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
    
    If Not Par�metrosAberto Then
        Unload Me
        Exit Sub
    End If
    
    TestaInferior TBLProduto, lAllowEdit, lAllowDelete, lAllowConsult
    TestaSuperior TBLProduto, lAllowEdit, lAllowDelete, lAllowConsult
    If TBLProduto.RecordCount = 0 Then
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
    
    ZeraCampos
        
    lAllowInsert = Allow("PRODUTO", "I")
    lAllowEdit = Allow("PRODUTO", "A")
    lAllowDelete = Allow("PRODUTO", "E")
    lAllowConsult = Allow("PRODUTO", "C")
    lAllowEditAmount = Allow("PRODUTO", "Q")
    
    lAtualizar = True 'Indica que o modulo possui a fun��o atualizar
    
    lInserir = False
    lAlterar = False
    lPula = False
    
    'Abertura das tabelas
    ProdutoAberto = AbreTabela(Dicion�rio, "CADASTRO", "PRODUTO", DBCadastro, TBLProduto, TBLTabela, dbOpenTable)
    
    If ProdutoAberto Then
        IndiceProdutoAtivo = "PRODUTO1"
        TBLProduto.Index = IndiceProdutoAtivo
    Else
        MsgBox "N�o consegui abrir a tabela 'Produto' !", vbCritical, "Erro"
        Exit Sub
    End If
    
    C�digoDoProdutoAberto = AbreTabela(Dicion�rio, "CADASTRO", "C�DIGO DO PRODUTO", DBCadastro, TBLC�digoDoProduto, TBLTabela, dbOpenTable)
    
    If C�digoDoProdutoAberto Then
        IndiceC�digoDoProdutoAtivo = "C�DIGODOPRODUTO2"
        TBLC�digoDoProduto.Index = IndiceC�digoDoProdutoAtivo
    Else
        MsgBox "N�o consegui abrir a tabela 'C�digo do Produto' !", vbCritical, "Erro"
        Exit Sub
    End If
    
    LoteAberto = AbreTabela(Dicion�rio, "CADASTRO", "LOTE DO PRODUTO", DBCadastro, TBLLote, TBLTabela, dbOpenTable)
    
    If LoteAberto Then
        IndiceLoteAtivo = "LOTEDOPRODUTO2"
        TBLLote.Index = IndiceLoteAtivo
    Else
        MsgBox "N�o consegui abrir a tabela 'Lote do Produto' !", vbCritical, "Erro"
        Exit Sub
    End If
    
    Pre�oAberto = AbreTabela(Dicion�rio, "CADASTRO", "PRE�O DO PRODUTO", DBCadastro, TBLPre�o, TBLTabela, dbOpenTable)
    
    If Pre�oAberto Then
        IndicePre�oAtivo = "PRE�ODOPRODUTO2"
        TBLPre�o.Index = IndicePre�oAtivo
    Else
        MsgBox "N�o consegui abrir a tabela 'Pre�o do Produto' !", vbCritical, "Erro"
        Exit Sub
    End If
    
    DepartamentoAberto = AbreTabela(Dicion�rio, "CADASTRO", "DEPARTAMENTO", DBCadastro, TBLDepartamento, TBLTabela, dbOpenTable)
        
    If DepartamentoAberto Then
        If TBLDepartamento.RecordCount = 0 Then
            MsgBox "Tabela 'Departamento' est� vazia! " + vbCr + "Antes de tentar cadastrar um produto, primeiro cadastre esta tabela.", vbInformation, "Aviso"
            DepartamentoAberto = False
            Exit Sub
        End If
        IndiceDepartamentoAtivo = "DEPARTAMENTO2"
        TBLDepartamento.Index = IndiceDepartamentoAtivo
    Else
        MsgBox "N�o consegui abrir a tabela 'Departamento' !", vbCritical, "Erro"
        Exit Sub
    End If
    
    Se��oAberto = AbreTabela(Dicion�rio, "CADASTRO", "SE��O", DBCadastro, TBLSe��o, TBLTabela, dbOpenTable)
    
    If Se��oAberto Then
        If TBLSe��o.RecordCount = 0 Then
            MsgBox "Tabela 'Se��o' est� vazia! " + vbCr + "Antes de tentar cadastrar um produto, primeiro cadastre esta tabela.", vbInformation, "Aviso"
            Se��oAberto = False
            Exit Sub
        End If
        IndiceSe��oAtivo = "SE��O2"
        TBLSe��o.Index = IndiceSe��oAtivo
    Else
        MsgBox "N�o consegui abrir a tabela 'Se��o' !", vbCritical, "Erro"
        Exit Sub
    End If
    
    DepartamentoSe��oAberto = AbreTabela(Dicion�rio, "CADASTRO", "DEPARTAMENTO - SE��O", DBCadastro, TBLDepartamentoSe��o, TBLTabela, dbOpenTable)
    
    If DepartamentoSe��oAberto Then
        If TBLDepartamentoSe��o.RecordCount = 0 Then
            MsgBox "Tabela 'Departamento - Se��o' est� vazia! " + vbCr + "Antes de tentar cadastrar um produto, primeiro cadastre esta tabela.", vbInformation, "Aviso"
            DepartamentoSe��oAberto = False
            Exit Sub
        End If
        IndiceDepartamentoSe��oAtivo = "DEPARTAMENTOSE��O1"
        TBLDepartamentoSe��o.Index = IndiceDepartamentoSe��oAtivo
    Else
        MsgBox "N�o consegui abrir a tabela 'Departamento - Se��o' !", vbCritical, "Erro"
        Exit Sub
    End If
    
    TipoDeICMAberto = AbreTabela(Dicion�rio, "CADASTRO", "TIPO DE ICM", DBCadastro, TBLTipoDeICM, TBLTabela, dbOpenTable)
    
    If TipoDeICMAberto Then
        If TBLTipoDeICM.RecordCount = 0 Then
            MsgBox "Tabela 'Tipo de ICM' est� vazia! " + vbCr + "Antes de tentar cadastrar um produto, primeiro cadastre esta tabela.", vbInformation, "Aviso"
            TipoDeICMAberto = False
            Exit Sub
        End If
        IndiceTipoDeICMAtivo = "TIPODEICM2"
        TBLTipoDeICM.Index = IndiceTipoDeICMAtivo
    Else
        MsgBox "N�o consegui abrir a tabela 'Tipo de ICM' !", vbCritical, "Erro"
        Exit Sub
    End If
    
    TipoDeEmbalagemAberto = AbreTabela(Dicion�rio, "CADASTRO", "TIPO DE EMBALAGEM", DBCadastro, TBLTipoDeEmbalagem, TBLTabela, dbOpenTable)
    
    If TipoDeEmbalagemAberto Then
        If TBLTipoDeEmbalagem.RecordCount = 0 Then
            MsgBox "Tabela 'Tipo de Embalagem' est� vazia! " + vbCr + "Antes de tentar cadastrar um produto, primeiro cadastre esta tabela.", vbInformation, "Aviso"
            TipoDeEmbalagemAberto = False
            Exit Sub
        End If
        IndiceTipoDeEmbalagemAtivo = "TIPODEEMBALAGEM2"
        TBLTipoDeEmbalagem.Index = IndiceTipoDeEmbalagemAtivo
    Else
        MsgBox "N�o consegui abrir a tabela 'Tipo de Embalagem' !", vbCritical, "Erro"
        Exit Sub
    End If
    
    UnidadesAberto = AbreTabela(Dicion�rio, "CADASTRO", "UNIDADES", DBCadastro, TBLUnidades, TBLTabela, dbOpenTable)
    
    If UnidadesAberto Then
        If TBLUnidades.RecordCount = 0 Then
            MsgBox "Tabela 'Unidades' est� vazia! " + vbCr + "Antes de tentar cadastrar um produto, primeiro cadastre esta tabela.", vbInformation, "Aviso"
            UnidadesAberto = False
            Exit Sub
        End If
        IndiceUnidadesAtivo = "UNIDADES2"
        TBLUnidades.Index = IndiceUnidadesAtivo
    Else
        MsgBox "N�o consegui abrir a tabela 'Unidades' !", vbCritical, "Erro"
        Exit Sub
    End If
    
    LocalAberto = AbreTabela(Dicion�rio, "CADASTRO", "LOCAL DO PRODUTO", DBCadastro, TBLLocal, TBLTabela, dbOpenTable)
    
    If LocalAberto Then
        If TBLLocal.RecordCount = 0 Then
            MsgBox "Tabela 'Local' est� vazia! " + vbCr + "Antes de tentar cadastrar um produto, primeiro cadastre esta tabela.", vbInformation, "Aviso"
            LocalAberto = False
            Exit Sub
        End If
        IndiceLocalAtivo = "LOCALDOPRODUTO2"
        TBLLocal.Index = IndiceLocalAtivo
    Else
        MsgBox "N�o consegui abrir a tabela 'Local do Produto' !", vbCritical, "Erro"
        Exit Sub
    End If
    
    FornecedorAberto = AbreTabela(Dicion�rio, "CADASTRO", "FORNECEDOR", DBCadastro, TBLFornecedor, TBLTabela, dbOpenTable)
        
    If FornecedorAberto Then
        If TBLFornecedor.RecordCount = 0 Then
            MsgBox "Tabela 'Fornecedor' est� vazia! " + vbCr + "Antes de tentar cadastrar um produto, primeiro cadastre esta tabela.", vbInformation, "Aviso"
            FornecedorAberto = False
            Exit Sub
        End If
        IndiceFornecedorAtivo = "FORNECEDOR1"
        TBLFornecedor.Index = IndiceFornecedorAtivo
    Else
        MsgBox "N�o consegui abrir a tabela 'Fornecedor' !", vbCritical, "Erro"
        Exit Sub
    End If
    
    Par�metrosAberto = AbreTabela(Dicion�rio, "SISTEMA", "PAR�METROS", DBSistema, TBLPar�metros, TBLTabela, dbOpenTable)
    
    If Par�metrosAberto Then
    Else
        MsgBox "N�o consegui abrir a tabela 'Par�metros' !", vbCritical, "Erro"
        Exit Sub
    End If
    
    'Todas as tabelas foram abertas
    FillDepartamento
    TBLDepartamento.MoveFirst
    TBLDepartamento.Move cmbDepartamento.ListIndex
    
    FillSe��o TBLDepartamento("C�DIGO")
    
    FillUnidades
    
    FillTipoDeICM
    
    FillTipoDeEmbalagem
    
    FillLocal
    
    txtPeso = "0,000"
    txtPeso_LostFocus
    txtDescontoM�ximo = " 0,00"
    txtDescontoDePromo��o = " 0,00"
        
    Bot�oIncluir lAllowInsert
 
    If TBLProduto.RecordCount = 0 Then
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
        
    If TBLProduto.RecordCount = 0 Or TBLProduto.RecordCount = 1 Then
        Navega��oSuperior False
    Else
        Navega��oInferior lAllowConsult
    End If
    
    lInserir = False
    lAlterar = False
    StatusBarAviso = "Pronto"
    Relat�rio = AddPath(Aplica��oPath, "REPORT\PRODUTO.RPT")
    TotalDatabaseName = 1
    DataBaseName(1) = AddPath(Aplica��oPath, "DATABASE\CADASTRO.MDB")
    mFechar = False
    Exit Sub
    
Erro:
    GeraMensagemDeErro "Produto - Load"
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
    
    Set frmProduto = Nothing
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Dim Cont As Byte
    
    If ProdutoAberto Then
        TBLProduto.Close
    End If
    If C�digoDoProdutoAberto Then
        TBLC�digoDoProduto.Close
    End If
    If LoteAberto Then
        TBLLote.Close
    End If
    If Pre�oAberto Then
        TBLPre�o.Close
    End If
    If DepartamentoAberto Then
        TBLDepartamento.Close
    End If
    If Se��oAberto Then
        TBLSe��o.Close
    End If
    If DepartamentoSe��oAberto Then
        TBLDepartamentoSe��o.Close
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
        AllBot�es False
    End If
End Sub
Private Sub txtDescontoDePromo��o_Change()
    FormatMask "@K 99,99", txtDescontoDePromo��o
End Sub
Private Sub txtDescontoDePromo��o_KeyPress(KeyAscii As Integer)
    If Not lInserir Then
        lAlterar = True
        StatusBarAviso = "Altera��o"
        BarraDeStatus StatusBarAviso
    End If
End Sub
Private Sub txtDescontoDePromo��o_LostFocus()
    FormatMask "@V #0,00", txtDescontoDePromo��o
End Sub
Private Sub txtDescontoM�ximo_Change()
    FormatMask "@K 99,99", txtDescontoM�ximo
End Sub
Private Sub txtDescontoM�ximo_KeyPress(KeyAscii As Integer)
    If Not lInserir Then
        lAlterar = True
        StatusBarAviso = "Altera��o"
        BarraDeStatus StatusBarAviso
    End If
End Sub
Private Sub txtDescontoM�ximo_LostFocus()
    FormatMask "@V #0,00", txtDescontoM�ximo
End Sub
Private Sub txtDescri��o_Change()
    If Not lPula Then
        FormatMask "@!S50", txtDescri��o
    End If
End Sub
Private Sub txtDescri��o_KeyPress(KeyAscii As Integer)
    If Not lInserir Then
        lAlterar = True
        StatusBarAviso = "Altera��o"
        BarraDeStatus StatusBarAviso
    End If
End Sub
Private Sub txtIn�cio_Change()
    If Not lPula Then
        lPula = True
        FormatMask DataMask, txtIn�cio
        lPula = False
    End If
End Sub
Private Sub txtIn�cio_KeyPress(KeyAscii As Integer)
    If Not lInserir Then
        lAlterar = True
        StatusBarAviso = "Altera��o"
        BarraDeStatus StatusBarAviso
    End If
End Sub
Private Sub txtIn�cio_LostFocus()
    If StrTran(txtIn�cio.Text, "/") <> Space(8) Then
        lPula = True
        CorrigeData DataMask, txtIn�cio, Date
        lPula = False
        If Not FormatMask(CheckDataMask, txtIn�cio) Then
            Beep
            MsgBox "Data inv�lida !", vbCritical, "Erro"
            txtIn�cio.SelStart = 0
            txtIn�cio.SetFocus
        End If
    End If
End Sub
Private Sub txtIPI_Change()
    FormatMask "@K 99,99", txtIPI
End Sub
Private Sub txtIPI_KeyPress(KeyAscii As Integer)
    If Not lInserir Then
        lAlterar = True
        StatusBarAviso = "Altera��o"
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
        StatusBarAviso = "Altera��o"
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
        StatusBarAviso = "Altera��o"
        BarraDeStatus StatusBarAviso
    End If
End Sub
Private Sub txtQuantidade_LostFocus()
    FormatMask "@V ######0,00", txtQuantidade
End Sub
Private Sub txtT�RMINO_Change()
    If Not lPula Then
        lPula = True
        FormatMask DataMask, txtT�rmino
        lPula = False
    End If
End Sub
Private Sub txtT�rmino_KeyPress(KeyAscii As Integer)
    If Not lInserir Then
        lAlterar = True
        StatusBarAviso = "Altera��o"
        BarraDeStatus StatusBarAviso
    End If
End Sub
Private Sub txtT�RMINO_LostFocus()
    If StrTran(txtT�rmino.Text, "/") <> Space(8) Then
        lPula = True
        CorrigeData DataMask, txtT�rmino, Date
        lPula = False
        If Not FormatMask(CheckDataMask, txtT�rmino) Then
            Beep
            MsgBox "Data inv�lida !", vbCritical, "Erro"
            txtT�rmino.SelStart = 0
            txtT�rmino.SetFocus
        End If
    End If
End Sub
