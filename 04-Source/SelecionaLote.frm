VERSION 5.00
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#2.5#0"; "SS32X25.OCX"
Begin VB.Form frmSelecionaLote 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Selecinar Lote / Quantidade"
   ClientHeight    =   3975
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5220
   Icon            =   "SelecionaLote.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3975
   ScaleWidth      =   5220
   Begin VB.Frame frQuantidadeVendida 
      Caption         =   "Quantidade Vendida"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   0
      TabIndex        =   4
      Top             =   2730
      Width           =   5205
      Begin VB.Label lblQuantidadeVendida 
         Alignment       =   2  'Center
         Caption         =   "0,00"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   1380
         TabIndex        =   5
         Top             =   270
         Width           =   2685
      End
   End
   Begin VB.Frame frLote 
      Caption         =   "Lote"
      Height          =   2715
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   5205
      Begin FPSpread.vaSpread dbgrdLotes 
         Height          =   2445
         Left            =   60
         TabIndex        =   0
         Top             =   210
         Width           =   5085
         _Version        =   131077
         _ExtentX        =   8969
         _ExtentY        =   4313
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
         MaxCols         =   4
         MaxRows         =   9
         ScrollBars      =   2
         SelectBlockOptions=   0
         SpreadDesigner  =   "SelecionaLote.frx":030A
         UserResize      =   0
      End
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   345
      Left            =   3960
      TabIndex        =   2
      Top             =   3630
      Width           =   1245
   End
   Begin VB.CommandButton cmdGravar 
      Caption         =   "&Gravar"
      Height          =   345
      Left            =   2640
      TabIndex        =   1
      Top             =   3630
      Width           =   1245
   End
End
Attribute VB_Name = "frmSelecionaLote"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const MAXCOLS = 4

Dim lPula    As Boolean
Dim mlFechar As Boolean

Dim mTotalRows%

Dim TBLLoteDoProduto As Table
Dim LoteDoProdutoAberto As Boolean
Dim IndiceLoteDoProdutoAtivo$

Public MatrizLote     As ClassLote
Public mCódigoProduto As Long
Public mQuantidade    As Single
Private Sub FillGrid()
    Dim PosLote As Integer
    
    mTotalRows = 0
    
    TBLLoteDoProduto.Seek "=", mCódigoProduto
    If Not TBLLoteDoProduto.NoMatch Then
        Do While Not TBLLoteDoProduto.EOF And TBLLoteDoProduto("CÓDIGO DO PRODUTO") = mCódigoProduto
            mTotalRows = mTotalRows + 1
            dbgrdLotes.Row = mTotalRows
            
            dbgrdLotes.Col = 1
            dbgrdLotes.Text = TBLLoteDoProduto("CÓDIGO DO LOTE") & "-" & TBLLoteDoProduto("DÍGITO DO LOTE") 'Código do Lote
            
            dbgrdLotes.Col = 2
            dbgrdLotes.Text = FormatStringMask("@V ######0,00", StrVal(TBLLoteDoProduto("QUANTIDADE"))) 'Quantidade em estoque
            dbgrdLotes.CellType = SS_CELL_TYPE_FLOAT
            
            dbgrdLotes.Col = 3
            dbgrdLotes.Text = FormatStringMask("@V ######0,00", StrVal(TBLLoteDoProduto("MÚLTIPLO"))) 'Quantidade em estoque
            dbgrdLotes.CellType = SS_CELL_TYPE_FLOAT
            
            dbgrdLotes.Col = 4
            dbgrdLotes.CellType = SS_CELL_TYPE_INTEGER
            PosLote = MatrizLote.Ascan(TBLLoteDoProduto("CÓDIGO DO LOTE") & "-" & TBLLoteDoProduto("DÍGITO DO LOTE"))
            If PosLote > 0 Then
                dbgrdLotes.Text = FormatStringMask("@V ######0", StrVal(MatrizLote.GetQuantidade(PosLote)))   'Quantidade a ser vendida
            Else
                dbgrdLotes.Text = FormatStringMask("@V ######0", "0,00")   'Quantidade a ser vendida
            End If
            TBLLoteDoProduto.MoveNext

            If TBLLoteDoProduto.EOF Then
                Exit Do
            End If
        Loop
    End If
End Sub
Private Sub cmdCancelar_Click()
    Unload Me
End Sub
Private Sub cmdGravar_Click()
    Dim Cont As Integer, PosLote As Integer
    Dim Código As String, Múltiplo As Single, Quantidade As Single
    
    For Cont = 1 To mTotalRows
        dbgrdLotes.Row = Cont
        
        dbgrdLotes.Col = 1
        Código = dbgrdLotes.Text
        
        dbgrdLotes.Col = 3
        Múltiplo = ValStr(dbgrdLotes.Text)
        
        dbgrdLotes.Col = 4
        Quantidade = ValStr(dbgrdLotes.Text)
        
        PosLote = MatrizLote.Ascan(Código)
        If Quantidade = 0 Then
            If PosLote > 0 Then
                MatrizLote.RemoveItem PosLote
            End If
        Else
            If PosLote > 0 Then
                MatrizLote.SetQuantidade PosLote, Quantidade
            Else
                MatrizLote.AddNew Código, Múltiplo, Quantidade
            End If
        End If
    Next
    
    Unload Me
End Sub
Private Sub dbgrdLotes_EditMode(ByVal Col As Long, ByVal Row As Long, ByVal Mode As Integer, ByVal ChangeMade As Boolean)
    Dim Cont
    Dim QuantidadeVendida As Single
    Dim Múltiplo As Single, Quantidade As Single
    
    If ChangeMade Then
        QuantidadeVendida = 0
        For Cont = 1 To mTotalRows
            dbgrdLotes.Row = Cont
            
            dbgrdLotes.Col = 3
            Múltiplo = ValStr(dbgrdLotes.Text)
            
            dbgrdLotes.Col = 4
            Quantidade = ValStr(dbgrdLotes.Text)
            
            QuantidadeVendida = QuantidadeVendida + (Múltiplo * Quantidade)
        Next
        lblQuantidadeVendida = StrVal(QuantidadeVendida)
    End If
End Sub
Private Sub Form_Activate()
    If mlFechar Then
        Unload Me
        Exit Sub
    End If
    
    If Not LoteDoProdutoAberto Then
        Unload Me
        Exit Sub
    End If
    
    dbgrdLotes.Refresh
End Sub
Private Sub Form_Load()
    On Error GoTo Erro
    
    Dim Cont, Cont1
    
    LoteDoProdutoAberto = AbreTabela(Dicionário, "CADASTRO", "LOTE DO PRODUTO", DBCadastro, TBLLoteDoProduto, TBLTabela, dbOpenTable)
    
    If LoteDoProdutoAberto Then
        IndiceLoteDoProdutoAtivo = "LOTEDOPRODUTO2"
        TBLLoteDoProduto.Index = IndiceLoteDoProdutoAtivo
    Else
        MsgBox "Não consegui abrir a tabela 'Código do Produto' !", vbCritical, "Erro"
        Exit Sub
    End If
    
    lblQuantidadeVendida = StrVal(mQuantidade)
    
    dbgrdLotes.Col = 1
    dbgrdLotes.Row = 0
    dbgrdLotes.Text = "Código"
    dbgrdLotes.ColWidth(1) = 8.25
    dbgrdLotes.Lock = True

    dbgrdLotes.Col = 2
    dbgrdLotes.Row = 0
    dbgrdLotes.Text = "Qtd.de Estoque"
    dbgrdLotes.ColWidth(2) = 11.875
    dbgrdLotes.Lock = True
    
    dbgrdLotes.Col = 3
    dbgrdLotes.Row = 0
    dbgrdLotes.Text = "Múltiplo de"
    dbgrdLotes.ColWidth(3) = 9.125
    dbgrdLotes.Lock = True

    dbgrdLotes.Col = 4
    dbgrdLotes.Row = 0
    dbgrdLotes.Text = "Qtd. Vendida"
    dbgrdLotes.ColWidth(4) = 10.5
    dbgrdLotes.Lock = True
    
    For Cont = 1 To 3
        dbgrdLotes.Col = Cont
        For Cont1 = 1 To 9
            dbgrdLotes.Row = Cont1
            dbgrdLotes.Lock = True
        Next
    Next

    FillGrid
        
    Exit Sub
    
Erro:
    mlFechar = True
End Sub
Private Sub Form_Unload(Cancel As Integer)
    If LoteDoProdutoAberto Then
        TBLLoteDoProduto.Close
    End If
    mQuantidade = ValStr(lblQuantidadeVendida.Caption)
End Sub
