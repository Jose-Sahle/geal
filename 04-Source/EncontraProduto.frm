VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form frmEncontraProduto 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Encontra Produtos"
   ClientHeight    =   5655
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9180
   Icon            =   "EncontraProduto.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5655
   ScaleWidth      =   9180
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   345
      Left            =   7890
      TabIndex        =   14
      Top             =   5280
      Width           =   1245
   End
   Begin VB.Data dtProdutos 
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   315
      Left            =   30
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   5280
      Visible         =   0   'False
      Width           =   1230
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   345
      Left            =   6630
      TabIndex        =   7
      Top             =   5280
      Width           =   1245
   End
   Begin VB.Frame frProdutosEncontrados 
      Caption         =   "Produtos &Encontrados"
      Height          =   2655
      Left            =   30
      TabIndex        =   13
      Top             =   2580
      Width           =   9135
      Begin MSDBGrid.DBGrid dbgrdProduto 
         Bindings        =   "EncontraProduto.frx":030A
         Height          =   2385
         Left            =   60
         OleObjectBlob   =   "EncontraProduto.frx":0323
         TabIndex        =   6
         Top             =   210
         Width           =   9045
      End
   End
   Begin VB.Frame frCritério 
      Caption         =   "Critério de Busca"
      Height          =   1545
      Left            =   4590
      TabIndex        =   12
      Top             =   1020
      Width           =   4590
      Begin VB.OptionButton optIdentico 
         Caption         =   "&Idêntico"
         Height          =   255
         Left            =   330
         TabIndex        =   3
         Top             =   300
         Width           =   2265
      End
      Begin VB.OptionButton optIgual 
         Caption         =   "Ig&ual"
         Height          =   225
         Left            =   330
         TabIndex        =   4
         Top             =   720
         Width           =   2265
      End
      Begin VB.OptionButton optContido 
         Caption         =   "C&ontido"
         Height          =   225
         Left            =   330
         TabIndex        =   5
         Top             =   1080
         Value           =   -1  'True
         Width           =   2595
      End
   End
   Begin VB.Frame frTipoDeBusca 
      Caption         =   "Tipo de Busca "
      Height          =   1545
      Left            =   0
      TabIndex        =   11
      Top             =   1020
      Width           =   4590
      Begin VB.OptionButton optCódigoDoProduto 
         Caption         =   "Código do &Produto"
         Height          =   375
         Left            =   270
         TabIndex        =   2
         Top             =   870
         Width           =   2295
      End
      Begin VB.OptionButton optNomeDoProduto 
         Caption         =   "&Descrição do Produto"
         Height          =   285
         Left            =   270
         TabIndex        =   1
         Top             =   420
         Value           =   -1  'True
         Width           =   2655
      End
   End
   Begin VB.Frame frProduto 
      Height          =   1005
      Left            =   0
      TabIndex        =   9
      Top             =   0
      Width           =   9165
      Begin VB.CommandButton cmdBuscar 
         Caption         =   "&Buscar"
         Height          =   315
         Left            =   8190
         TabIndex        =   8
         Top             =   390
         Width           =   855
      End
      Begin VB.TextBox txtChave 
         Height          =   315
         Left            =   1500
         TabIndex        =   0
         Top             =   390
         Width           =   6555
      End
      Begin VB.Label lblChave 
         Caption         =   "&Chave de Busca"
         Height          =   255
         Left            =   180
         TabIndex        =   10
         Top             =   420
         Width           =   1245
      End
   End
End
Attribute VB_Name = "frmEncontraProduto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim lPula As Boolean

Dim mCampoChave As String

Public Código
Public Janela  As Form
Public NoModal As Boolean
Public CondiçãoSQL As String
Public TipoDeBusca As Byte
Public CritérioDeBusca As Byte
Private Sub Resize()
    dbgrdProduto.ReBind
    dbgrdProduto.Columns(0).Width = 4094
    dbgrdProduto.Columns(1).Width = 1170
    dbgrdProduto.Columns(2).Width = 0
    dbgrdProduto.Columns(3).Width = 1500
    dbgrdProduto.Columns(4).Width = 1700
End Sub
Private Sub cmdBuscar_Click()
    txtChave = Trim(txtChave)
    If optContido.Value Then
        dtProdutos.RecordSource = "Select [PRODUTO].[DESCRIÇÃO], [PRODUTO].[QUANTIDADE], [PRODUTO].[CÓDIGO], [PREÇO DO PRODUTO].[PREÇO DE VENDA], MAX([CÓDIGO DO PRODUTO].[CÓDIGO DO FORNECEDOR]) AS [CÓDIGO] FROM ([PRODUTO] LEFT JOIN [CÓDIGO DO PRODUTO] ON [PRODUTO].[CÓDIGO] = [CÓDIGO DO PRODUTO].[CÓDIGO DO PRODUTO])  LEFT JOIN [PREÇO DO PRODUTO] ON [PRODUTO].[CÓDIGO] = [PREÇO DO PRODUTO].[CÓDIGO DO PRODUTO] WHERE [CÓDIGO DO PRODUTO].[FORNECEDOR] = '" & gCGC & "' And " & mCampoChave & " LIKE " & " '*" & txtChave & "*' " & " GROUP BY  [PRODUTO].[DESCRIÇÃO], [PRODUTO].[QUANTIDADE], [PRODUTO].[CÓDIGO],[PREÇO DO PRODUTO].[PREÇO DE VENDA] ORDER BY [PRODUTO].[DESCRIÇÃO]"
    ElseIf optIdentico Then
        dtProdutos.RecordSource = "Select [PRODUTO].[DESCRIÇÃO], [PRODUTO].[QUANTIDADE], [PRODUTO].[CÓDIGO], [PREÇO DO PRODUTO].[PREÇO DE VENDA], MAX([CÓDIGO DO PRODUTO].[CÓDIGO DO FORNECEDOR]) AS [CÓDIGO] FROM ([PRODUTO] LEFT JOIN [CÓDIGO DO PRODUTO] ON [PRODUTO].[CÓDIGO] = [CÓDIGO DO PRODUTO].[CÓDIGO DO PRODUTO])  LEFT JOIN [PREÇO DO PRODUTO] ON [PRODUTO].[CÓDIGO] = [PREÇO DO PRODUTO].[CÓDIGO DO PRODUTO] WHERE [CÓDIGO DO PRODUTO].[FORNECEDOR] = '" & gCGC & "' And " & mCampoChave & " = " & " '" & txtChave & "' GROUP BY  [PRODUTO].[DESCRIÇÃO], [PRODUTO].[QUANTIDADE], [PRODUTO].[CÓDIGO],[PREÇO DO PRODUTO].[PREÇO DE VENDA] ORDER BY [PRODUTO].[DESCRIÇÃO]"
    ElseIf optIgual Then
        dtProdutos.RecordSource = "Select [PRODUTO].[DESCRIÇÃO], [PRODUTO].[QUANTIDADE], [PRODUTO].[CÓDIGO], [PREÇO DO PRODUTO].[PREÇO DE VENDA], MAX([CÓDIGO DO PRODUTO].[CÓDIGO DO FORNECEDOR]) AS [CÓDIGO] FROM ([PRODUTO] LEFT JOIN [CÓDIGO DO PRODUTO] ON [PRODUTO].[CÓDIGO] = [CÓDIGO DO PRODUTO].[CÓDIGO DO PRODUTO])  LEFT JOIN [PREÇO DO PRODUTO] ON [PRODUTO].[CÓDIGO] = [PREÇO DO PRODUTO].[CÓDIGO DO PRODUTO] WHERE [CÓDIGO DO PRODUTO].[FORNECEDOR] = '" & gCGC & "' And " & mCampoChave & " > " & " '" & txtChave & "' AND " & mCampoChave & " <= '" & txtChave & Chr(255) & "'" & " GROUP BY  [PRODUTO].[DESCRIÇÃO], [PRODUTO].[QUANTIDADE], [PRODUTO].[CÓDIGO],[PREÇO DO PRODUTO].[PREÇO DE VENDA] ORDER BY [PRODUTO].[DESCRIÇÃO]"
    End If
    
    dtProdutos.Refresh
    Resize
End Sub
Private Sub cmdCancelar_Click()
    Código = Empty
    Unload Me
End Sub
Private Sub cmdOK_Click()
    If dbgrdProduto.Row = -1 Then
        MsgBox "Nenhum produto foi selecionado !", vbInformation, "Aviso"
        Exit Sub
    End If
    Código = dbgrdProduto.Columns(2).Text
    If Not NoModal Then
        Unload Me
    Else
        Janela.Posicionar Código
    End If
End Sub
Private Sub dbgrdProduto_ColResize(ByVal ColIndex As Integer, Cancel As Integer)
    Cancel = 1
End Sub
Private Sub dbgrdProduto_DblClick()
    cmdOK_Click
End Sub
Private Sub dbgrdProduto_RowResize(Cancel As Integer)
    Cancel = 1
End Sub
Private Sub Form_Load()
    If TipoDeBusca = 0 Or TipoDeBusca = 1 Then
        mCampoChave = " [PRODUTO].[DESCRIÇÃO] "
        optNomeDoProduto.Value = True
        TipoDeBusca = 1
    ElseIf TipoDeBusca = 2 Then
        mCampoChave = " [CÓDIGO DO PRODUTO].[CÓDIGO DO FORNECEDOR] "
        optCódigoDoProduto.Value = True
    End If
        
    txtChave = CondiçãoSQL
    
    dtProdutos.DataBaseName = DBCadastro.Name
    If CritérioDeBusca = 0 Then
        dtProdutos.RecordSource = "SELECT [PRODUTO].[DESCRIÇÃO], [PRODUTO].[QUANTIDADE], [PRODUTO].[CÓDIGO], [PREÇO DO PRODUTO].[PREÇO DE VENDA],MAX([CÓDIGO DO PRODUTO].[CÓDIGO DO FORNECEDOR]) AS [CÓDIGO] FROM ( [PRODUTO] LEFT JOIN [CÓDIGO DO PRODUTO] ON [PRODUTO].[CÓDIGO] = [CÓDIGO DO PRODUTO].[CÓDIGO DO PRODUTO])  LEFT JOIN [PREÇO DO PRODUTO] ON [PRODUTO].[CÓDIGO] = [PREÇO DO PRODUTO].[CÓDIGO DO PRODUTO] WHERE [PRODUTO].[DESCRIÇÃO] = '9' And [CÓDIGO DO PRODUTO].[FORNECEDOR] = '" & gCGC & "' GROUP BY  [PRODUTO].[DESCRIÇÃO], [PRODUTO].[QUANTIDADE], [PRODUTO].[CÓDIGO], [PREÇO DO PRODUTO].[PREÇO DE VENDA] ORDER BY [PRODUTO].[DESCRIÇÃO]"
        CritérioDeBusca = 1
    ElseIf CritérioDeBusca = 1 Then
        optIdentico.Value = True
        dtProdutos.RecordSource = "Select [PRODUTO].[DESCRIÇÃO], [PRODUTO].[QUANTIDADE], [PRODUTO].[CÓDIGO], [PREÇO DO PRODUTO].[PREÇO DE VENDA], MAX([CÓDIGO DO PRODUTO].[CÓDIGO DO FORNECEDOR]) AS [CÓDIGO] FROM ([PRODUTO] LEFT JOIN [CÓDIGO DO PRODUTO] ON [PRODUTO].[CÓDIGO] = [CÓDIGO DO PRODUTO].[CÓDIGO DO PRODUTO])  LEFT JOIN [PREÇO DO PRODUTO] ON [PRODUTO].[CÓDIGO] = [PREÇO DO PRODUTO].[CÓDIGO DO PRODUTO] WHERE [CÓDIGO DO PRODUTO].[FORNECEDOR] = '" & gCGC & "' And " & mCampoChave & " = " & " '" & txtChave & "' GROUP BY  [PRODUTO].[DESCRIÇÃO], [PRODUTO].[QUANTIDADE], [PRODUTO].[CÓDIGO],[PREÇO DO PRODUTO].[PREÇO DE VENDA] ORDER BY [PRODUTO].[DESCRIÇÃO]"
    ElseIf CritérioDeBusca = 2 Then
        optIgual.Value = True
        dtProdutos.RecordSource = "Select [PRODUTO].[DESCRIÇÃO], [PRODUTO].[QUANTIDADE], [PRODUTO].[CÓDIGO], [PREÇO DO PRODUTO].[PREÇO DE VENDA], MAX([CÓDIGO DO PRODUTO].[CÓDIGO DO FORNECEDOR]) AS [CÓDIGO] FROM ([PRODUTO] LEFT JOIN [CÓDIGO DO PRODUTO] ON [PRODUTO].[CÓDIGO] = [CÓDIGO DO PRODUTO].[CÓDIGO DO PRODUTO])  LEFT JOIN [PREÇO DO PRODUTO] ON [PRODUTO].[CÓDIGO] = [PREÇO DO PRODUTO].[CÓDIGO DO PRODUTO] WHERE [CÓDIGO DO PRODUTO].[FORNECEDOR] = '" & gCGC & "' And " & mCampoChave & " > " & " '" & txtChave & "' AND " & mCampoChave & " <= '" & txtChave & Chr(255) & "'" & " GROUP BY  [PRODUTO].[DESCRIÇÃO], [PRODUTO].[QUANTIDADE], [PRODUTO].[CÓDIGO],[PREÇO DO PRODUTO].[PREÇO DE VENDA] ORDER BY [PRODUTO].[DESCRIÇÃO]"
    ElseIf CritérioDeBusca = 3 Then
        optContido.Value = True
        dtProdutos.RecordSource = "Select [PRODUTO].[DESCRIÇÃO], [PRODUTO].[QUANTIDADE], [PRODUTO].[CÓDIGO], [PREÇO DO PRODUTO].[PREÇO DE VENDA], MAX([CÓDIGO DO PRODUTO].[CÓDIGO DO FORNECEDOR]) AS [CÓDIGO] FROM ([PRODUTO] LEFT JOIN [CÓDIGO DO PRODUTO] ON [PRODUTO].[CÓDIGO] = [CÓDIGO DO PRODUTO].[CÓDIGO DO PRODUTO])  LEFT JOIN [PREÇO DO PRODUTO] ON [PRODUTO].[CÓDIGO] = [PREÇO DO PRODUTO].[CÓDIGO DO PRODUTO] WHERE [CÓDIGO DO PRODUTO].[FORNECEDOR] = '" & gCGC & "' And " & mCampoChave & " LIKE " & " '*" & txtChave & "*' " & " GROUP BY  [PRODUTO].[DESCRIÇÃO], [PRODUTO].[QUANTIDADE], [PRODUTO].[CÓDIGO],[PREÇO DO PRODUTO].[PREÇO DE VENDA] ORDER BY [PRODUTO].[DESCRIÇÃO]"
    End If
    dtProdutos.Refresh
    
    Resize
    
    If Not NoModal Then
        cmdCancelar.Visible = True
        cmdCancelar.Enabled = True
        cmdCancelar.Left = 7920
        cmdOK.Left = 6660
    Else
        cmdCancelar.Visible = False
        cmdCancelar.Enabled = False
        cmdOK.Left = 7920
    End If
End Sub
Private Sub frProdutosEncontrados_Click()
    dbgrdProduto.SetFocus
End Sub
Private Sub optCódigoDoProduto_Click()
    mCampoChave = " [CÓDIGO DO PRODUTO].[CÓDIGO DO FORNECEDOR] "
    TipoDeBusca = 2
End Sub
Private Sub optContido_Click()
    CritérioDeBusca = 3
End Sub
Private Sub optIdentico_Click()
    CritérioDeBusca = 1
End Sub
Private Sub optIgual_Click()
    CritérioDeBusca = 2
End Sub
Private Sub optNomeDoProduto_Click()
    mCampoChave = "[PRODUTO].[DESCRIÇÃO]"
    TipoDeBusca = 1
End Sub
Private Sub txtChave_Change()
    FormatMask "@!", txtChave
    CondiçãoSQL = txtChave
End Sub
