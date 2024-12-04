VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form frmVendaDevoluçãoTroca 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Venda - Devolução/Troca"
   ClientHeight    =   6300
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9540
   Icon            =   "VendaDevolução.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6300
   ScaleWidth      =   9540
   Begin VB.Frame frDadosCadastraisOrçamento 
      Height          =   1140
      Left            =   4770
      TabIndex        =   14
      Top             =   0
      Width           =   4770
      Begin VB.CommandButton cmdLocalizarOrçamento 
         Caption         =   "&Localizar"
         Height          =   375
         Left            =   2940
         TabIndex        =   4
         Top             =   450
         Width           =   1335
      End
      Begin VB.TextBox txtOrçamento 
         Height          =   285
         Left            =   1170
         TabIndex        =   2
         Top             =   270
         Width           =   975
      End
      Begin VB.TextBox txtDataOrçamento 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1170
         TabIndex        =   3
         Text            =   "  /  /"
         Top             =   690
         Width           =   990
      End
      Begin VB.Label lblOrçamento 
         Caption         =   "Orçamento"
         Height          =   180
         Left            =   150
         TabIndex        =   16
         Top             =   330
         Width           =   825
      End
      Begin VB.Label lblDataDoOrçamento 
         Caption         =   "Data"
         Height          =   210
         Left            =   150
         TabIndex        =   15
         Top             =   720
         Width           =   465
      End
   End
   Begin VB.Frame frItensDevolução 
      Caption         =   "Itens para devolução"
      Height          =   2355
      Left            =   0
      TabIndex        =   13
      Top             =   3510
      Width           =   9540
      Begin MSDBGrid.DBGrid dbgrdItensDevolução 
         Height          =   2025
         Left            =   90
         OleObjectBlob   =   "VendaDevolução.frx":030A
         TabIndex        =   6
         Top             =   210
         Width           =   9405
      End
   End
   Begin VB.Frame frDadosCadastraisDevolução 
      Height          =   1140
      Left            =   0
      TabIndex        =   10
      Top             =   0
      Width           =   4770
      Begin VB.TextBox txtDataDevolução 
         Enabled         =   0   'False
         Height          =   285
         Left            =   3510
         TabIndex        =   1
         Text            =   "  /  /"
         Top             =   510
         Width           =   990
      End
      Begin VB.TextBox txtDevolução 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1170
         TabIndex        =   0
         Top             =   510
         Width           =   975
      End
      Begin VB.Label lblDataDaDevolução 
         Caption         =   "Data"
         Height          =   210
         Left            =   2880
         TabIndex        =   12
         Top             =   540
         Width           =   465
      End
      Begin VB.Label lblDevolução 
         Caption         =   "Devolução"
         Height          =   180
         Left            =   150
         TabIndex        =   11
         Top             =   540
         Width           =   825
      End
   End
   Begin VB.Frame frItensOrçamento 
      Caption         =   "Itens do Orçamento"
      Height          =   2355
      Left            =   0
      TabIndex        =   9
      Top             =   1140
      Width           =   9540
      Begin MSDBGrid.DBGrid dbgrdItensOrçamento 
         Height          =   2025
         Left            =   60
         OleObjectBlob   =   "VendaDevolução.frx":0CE9
         TabIndex        =   5
         Top             =   210
         Width           =   9405
      End
   End
   Begin VB.CommandButton cmdGravar 
      Caption         =   "&Gravar"
      Height          =   345
      Left            =   6930
      TabIndex        =   7
      Top             =   5925
      Width           =   1245
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   345
      Left            =   8250
      TabIndex        =   8
      Top             =   5925
      Width           =   1245
   End
End
Attribute VB_Name = "frmVendaDevoluçãoTroca"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const MAXCOLS = 6

Const vbDescrição = 1
Const vbCódigo = 2
Const vbValorUnitário = 3
Const vbValValorUnitário = 4
Const vbIndice1 = 1
Const vbIndice2 = 2
Const vbIndice3 = 3
Const vbIndice4 = 4

Dim mUsuário As String

Dim mRecnoDevolução%
Dim mTotalRowsDevolução%
Dim dbgrdItensDevoluçãoArray() As String
Dim ArrayVendasItensRecnoDevolução() As Variant

Dim mRecnoOrçamento%
Dim mTotalRowsOrçamento%
Dim dbgrdItensOrçamentoArray() As String
Dim ArrayVendasItensRecnoOrçamento() As Variant

Dim lAllowInsert  As Boolean
Dim lAllowEdit    As Boolean
Dim lAllowDelete  As Boolean
Dim lAllowConsult As Boolean

Dim lPula As Boolean
Dim lFechar As Boolean
Dim lInserir As Boolean

Dim TBLVendas As Table
Dim VendasAberto As Boolean
Dim IndiceVendasAtivo$

Dim TBLVendasItens As Table
Dim VendasItensAberto As Boolean
Dim IndiceVendasItensAtivo$

Dim TBLParâmetros As Table
Dim ParâmetrosAberto As Boolean

Dim TBLVendasDevolução As Table
Dim VendasDevoluçãoAberto As Boolean
Dim IndiceVendasDevoluçãoAtivo$

Dim TBLVendasDevoluçãoItens As Table
Dim VendasDevoluçãoItensAberto As Boolean
Dim IndiceVendasDevoluçãoItensAtivo$

Dim TBLValeCliente As Table
Dim ValeClienteAberto As Boolean
Dim IndiceValeCliente$

Dim StatusBarAviso$

Dim DataBaseName(1 To 1) As String
Public Relatório$
Public TotalDatabaseName%

Public lAtualizar As Boolean
Private Sub AtivaCampos()
    frDadosCadastraisDevolução.Enabled = True
    frDadosCadastraisOrçamento.Enabled = True
    frItensOrçamento.Enabled = True
    frItensDevolução.Enabled = True
    'cmdIncluirItem.Enabled = True
    BotãoGravar (lInserir)
    cmdCancelar.Enabled = (lInserir)
    cmdGravar.Enabled = (lInserir)
End Sub
Private Function Cancelamento() As Boolean
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
    
    BarraDeStatus StatusBarAviso
    
    BotãoIncluir lAllowInsert
    
    BotãoGravar False
    DesativaCampos
    ZeraCampos
    Cancelamento = True
    lInserir = False
End Function
Private Sub DesativaCampos()
    frDadosCadastraisDevolução.Enabled = False
    frDadosCadastraisOrçamento.Enabled = False
    frItensOrçamento.Enabled = False
    frItensDevolução.Enabled = False
    cmdGravar.Enabled = False
    cmdCancelar.Enabled = False
    'cmdIncluirItem.Enabled = False
    cmdGravar.Enabled = False
    cmdCancelar.Enabled = False
    BotãoGravar False
End Sub
Private Sub FillGridOrçamento(ByVal Chave As Long)
    Dim pValor As Currency
    Dim pDesconto As Currency
    
    dbgrdItensOrçamento.ReBind
    
    ReDim dbgrdItensOrçamentoArray(MAXCOLS - 1, 0)
    ReDim ArrayVendasItensRecnoOrçamento(0)
    
    mTotalRowsOrçamento = 0
    mRecnoOrçamento = 0
    
    TBLVendasItens.Seek "=", Chave
    If Not TBLVendasItens.NoMatch Then
        Do While Not TBLVendasItens.EOF And TBLVendasItens("ORÇAMENTO") = Chave
            mRecnoOrçamento = mRecnoOrçamento + 1
            mTotalRowsOrçamento = mTotalRowsOrçamento + 1
            ReDim Preserve dbgrdItensOrçamentoArray(MAXCOLS - 1, mTotalRowsOrçamento - 1)
            ReDim Preserve ArrayVendasItensRecnoOrçamento(mTotalRowsOrçamento - 1)
            
            ArrayVendasItensRecnoOrçamento(mTotalRowsOrçamento - 1) = TBLVendasItens.Bookmark
            dbgrdItensOrçamentoArray(0, mTotalRowsOrçamento - 1) = SearchProduto(TBLVendasItens("CÓDIGO DO PRODUTO")) 'Nome do Produto
            dbgrdItensOrçamentoArray(1, mTotalRowsOrçamento - 1) = FormatStringMask("@V ######0", StrVal(TBLVendasItens("QUANTIDADE"))) 'Quantidade
            dbgrdItensOrçamentoArray(2, mTotalRowsOrçamento - 1) = FormatStringMask("@V ##.###.##0,00", StrVal(TBLVendasItens("VALOR UNITÁRIO"))) 'Preço Unitário
            dbgrdItensOrçamentoArray(3, mTotalRowsOrçamento - 1) = FormatStringMask("@V ##.###.##0,00", StrVal(TBLVendasItens("DESCONTO"))) 'Desconto no valor do produto
            
            pValor = TBLVendasItens("VALOR UNITÁRIO") * TBLVendasItens("QUANTIDADE")
            pValor = pValor - (pValor * TBLVendasItens("DESCONTO") / 100)
            dbgrdItensOrçamentoArray(4, mTotalRowsOrçamento - 1) = FormatStringMask("@V ##.###.##0,00", StrVal(pValor)) 'Valor total
            
            dbgrdItensOrçamentoArray(5, mTotalRowsOrçamento - 1) = TBLVendasItens("CÓDIGO DO PRODUTO") 'Código do Produto

            TBLVendasItens.MoveNext

            If TBLVendasItens.EOF Then
                Exit Do
            End If
        Loop
    End If
    
    dbgrdItensOrçamento.Refresh
End Sub
Private Sub Gravar()
    If SetRecords Then
        ZeraCampos
        DesativaCampos
        
        StatusBarAviso = "Pronto"
        
        BarraDeStatus StatusBarAviso
    End If
End Sub
Public Sub Incluir()
    
    ZeraCampos
    AtivaCampos
    
    lInserir = True
    
    BotãoGravar (lInserir)
    BotãoIncluir False
    BotãoGravar (lInserir)
    cmdCancelar.Enabled = (lInserir)
    
    StatusBarAviso = "Inclusão"
    BarraDeStatus StatusBarAviso
    
    txtOrçamento.SetFocus
End Sub
Private Sub Localizar()
    If PosRecords Then
        GetRecords
    End If
End Sub
Private Function PosRecords() As Boolean
    TBLVendas.Seek "=", Val(txtOrçamento)
    If TBLVendas.NoMatch Then
        PosRecords = False
        MsgBox "Não encontrei o orçamento " & txtOrçamento, vbInformation, "Aviso"
    Else
        If TBLVendas("TIPO") <> "V" Then
            MsgBox "Este orçamento não pode ser editado!", vbInformation, "Aviso"
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

    If Not lAllowConsult Then
        ZeraCampos
        DesativaCampos
        lPula = False
        Exit Sub
    End If
    
    txtDataOrçamento = TBLVendas("DATA DO ORÇAMENTO")
    CorrigeData DataMask, txtDataOrçamento, TBLVendas("DATA DO ORÇAMENTO")
    
    FillGridOrçamento TBLVendas("CÓDIGO")
    
    lPula = False
    If Not lAllowEdit Then
        DesativaCampos
    End If
    If Not lAllowEdit Then
        DesativaCampos
    End If
End Sub
Private Function SetRecords()
    On Error GoTo ErroVendasDevolução
    
    Dim Cont As Integer
    
    WS.BeginTrans 'Inicia uma Transação
            
    TBLParâmetros.Edit
    TBLParâmetros("DEVOLUÇÃO") = TBLParâmetros("DEVOLUÇÃO") + 1
    TBLParâmetros.Update
    TBLParâmetros.MoveFirst
    
    TBLVendasDevolução.AddNew
    
    TBLVendasDevolução("CÓDIGO") = TBLParâmetros("DEVOLUÇÃO")
    TBLVendasDevolução("ORÇAMENTO") = txtOrçamento
    TBLVendasDevolução("MOTIVO") = 0
    TBLVendasDevolução("OBSERVAÇÃO") = vbNull
    TBLVendasDevolução("BAIXADO") = False
    
    TBLVendasDevolução("USERNAME - CRIA") = mUsuário
    TBLVendasDevolução("DATA - CRIA") = Date
    TBLVendasDevolução("HORA - CRIA") = Time
    TBLVendasDevolução("USERNAME - ALTERA") = "VAZIO"
    TBLVendasDevolução("DATA - ALTERA") = vbNull
    TBLVendasDevolução("HORA - ALTERA") = vbNull
    
    TBLVendasDevolução.Update
            
    On Error GoTo ErroVendasDevoluçãoItens
    
    For Cont = 0 To mTotalRowsDevolução - 1
        TBLVendasDevoluçãoItens.AddNew
        
        TBLVendasDevoluçãoItens("CÓDIGO DA DEVOLUÇÃO") = TBLParâmetros("DEVOLUÇÃO")
        TBLVendasDevoluçãoItens("CÓDIGO DO PRODUTO") = dbgrdItensDevoluçãoArray(5, Cont)
        TBLVendasDevoluçãoItens("QUANTIDADE") = dbgrdItensDevoluçãoArray(1, Cont)
        TBLVendasDevoluçãoItens("DESTINO") = vbNull
        TBLVendasDevoluçãoItens("BAIXADO") = vbNull
        
        TBLVendasDevoluçãoItens.Update
    Next
    
    WS.CommitTrans 'Grava as alterações ou inclusões se não houverem erros
    
    If lInserir Then
        Log gUsuário, "Inclusão - (Venda) Devolução/Troca: " & txtOrçamento
    Else
        Log gUsuário, "Alteração - (Venda) Devolução/Troca: " & txtOrçamento
    End If
    
    SetRecords = True
    
    Exit Function
    
ErroVendasDevolução:
    TBLVendasDevolução.CancelUpdate
    GeraMensagemDeErro "Venda - Devolução/Troca - SetRecords - ErroVendasDevolução - " & txtOrçamento, True
    SetRecords = False
    Exit Function
    
ErroVendasDevoluçãoItens:
    TBLVendasDevoluçãoItens.CancelUpdate
    GeraMensagemDeErro "Venda - Devolução/Troca - SetRecords - ErroVendasDevoluçãoItens - " & txtOrçamento, True
    SetRecords = False
    
    Exit Function
End Function
Private Sub ZeraCampos()
    On Error Resume Next

    lPula = True
    lInserir = False
    
    txtDevolução = Empty
    txtDataDevolução = Empty
    
    txtOrçamento = Empty
    txtDataOrçamento = Empty
    
    ReDim dbgrdItensOrçamentoArray(MAXCOLS - 1, 0)
    ReDim dbgrdItensDevoluçãoArray(MAXCOLS - 1, 0)
    
    mTotalRowsOrçamento = 0
    mTotalRowsDevolução = 0
    
    mRecnoOrçamento = 0
    mRecnoDevolução = 0
    
    dbgrdItensOrçamento.ReBind
    dbgrdItensDevolução.ReBind
    
    lPula = False
End Sub
Private Sub cmdGravar_Click()
    'Valida Usuário
    frmValidaUsuário.Show 1
    
    mUsuário = frmValidaUsuário.Usuário
    
    Set frmValidaUsuário = Nothing
    
    If mUsuário = Empty Then
        Exit Sub
    End If
    
    Gravar
End Sub
Private Sub cmdCancelar_Click()
    Cancelamento
End Sub
Private Sub cmdLocalizarOrçamento_Click()
    Localizar
End Sub
Private Sub dbgrdItensDevolução_BeforeColUpdate(ByVal ColIndex As Integer, oldvalue As Variant, Cancel As Integer)
    Dim pCódigo As String, pQuantidade As Integer, pQuantidadeAtual As Integer
    Dim pValor As Currency, pDesconto As Currency
    Dim Cont As Integer
    
    If ColIndex = 1 Then
        pQuantidadeAtual = dbgrdItensDevolução.Text
        dbgrdItensDevolução.Col = 5
        pCódigo = dbgrdItensDevolução.Text
        For Cont = 0 To mTotalRowsOrçamento - 1
            If dbgrdItensOrçamentoArray(5, Cont) = pCódigo Then
                pQuantidade = dbgrdItensOrçamentoArray(1, Cont)
                pValor = ValStr(dbgrdItensOrçamentoArray(2, Cont))
                pDesconto = ValStr(dbgrdItensOrçamentoArray(3, Cont))
                Exit For
            End If
        Next
        
        If pQuantidade < pQuantidadeAtual Then
            MsgBox "A quantidade de devolução não pode ser maior que a quantidade de compra!", vbCritical, "Aviso"
            Cancel = 1
            Exit Sub
        End If
        
                
        pValor = pQuantidadeAtual * pValor
        pValor = pValor - (pValor * (pDesconto / 100))
        
        dbgrdItensDevolução.Col = 4
        dbgrdItensDevolução.Text = FormatStringMask("@V ##.###.##0,00", StrVal(pValor))
        
        dbgrdItensDevolução.Col = 1
        dbgrdItensDevolução.Text = pQuantidadeAtual
    End If
End Sub
Private Sub dbgrdItensDevolução_UnboundAddData(ByVal RowBuf As MSDBGrid.RowBuffer, NewRowBookmark As Variant)
    Dim Col%
        
    mTotalRowsDevolução = mTotalRowsDevolução + 1
    ReDim Preserve dbgrdItensDevoluçãoArray(MAXCOLS - 1, mTotalRowsDevolução - 1)
    
    'Sets the bookmark to the last row.
    NewRowBookmark = mTotalRowsDevolução - 1
    
    ' The following loop adds a new record to the database.
    For Col = 0 To UBound(dbgrdItensDevoluçãoArray, 1)
        If Not IsNull(RowBuf.Value(0, Col)) Then
            dbgrdItensDevoluçãoArray(Col, mTotalRowsDevolução - 1) = RowBuf.Value(0, Col)
        Else
            ' If no value set for column, then use the
            ' DefaultValue
            dbgrdItensDevoluçãoArray(Col, mTotalRowsDevolução - 1) = dbgrdItensDevolução.Columns(Col).DefaultValue
        End If
    Next
End Sub
Private Sub dbgrdItensDevolução_UnboundDeleteRow(Bookmark As Variant)
    Dim iCol As Integer, iRow As Integer
    
    ' Move all rows above the deleted row down in the
    ' array.
    
    For iRow = Bookmark + 1 To mTotalRowsDevolução - 1
        For iCol = 0 To MAXCOLS - 1
            dbgrdItensDevoluçãoArray(iCol, iRow - 1) = dbgrdItensDevoluçãoArray(iCol, iRow)
        Next iCol
    Next iRow
    
    mTotalRowsDevolução = mTotalRowsDevolução - 1
End Sub
Private Sub dbgrdItensDevolução_UnboundReadData(ByVal RowBuf As MSDBGrid.RowBuffer, StartLocation As Variant, ByVal ReadPriorRows As Boolean)
    Dim CurRow&, iRow As Integer, iCol As Integer, iRowsFetched As Integer, iIncr As Integer
    ' DBGrid is requesting rows so give them to it
    
    If mTotalRowsDevolução = 0 Then Exit Sub
    
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
        If CurRow < 0 Or CurRow >= mTotalRowsDevolução Then Exit For
        For iCol = 0 To UBound(dbgrdItensDevoluçãoArray, 1)
            RowBuf.Value(iRow, iCol) = dbgrdItensDevoluçãoArray(iCol, CurRow&)
        Next iCol
        ' Set bookmark using CurRow& which is also our
        ' array index
        RowBuf.Bookmark(iRow) = CStr(CurRow)
        CurRow = CurRow + iIncr
        iRowsFetched = iRowsFetched + 1
    Next iRow
    RowBuf.RowCount = iRowsFetched
End Sub
Private Sub dbgrdItensDevolução_UnboundWriteData(ByVal RowBuf As MSDBGrid.RowBuffer, WriteLocation As Variant)
    Dim iCol As Integer
    ' Data is being updated
    'MsgBox WriteLocation
    ' Update each column in the data set array
    For iCol = 0 To MAXCOLS - 1
        If Not IsNull(RowBuf.Value(0, iCol)) Then
            dbgrdItensDevoluçãoArray(iCol, WriteLocation) = RowBuf.Value(0, iCol)
        End If
    Next iCol
End Sub
Private Sub dbgrdItensOrçamento_BeforeColEdit(ByVal ColIndex As Integer, ByVal KeyAscii As Integer, Cancel As Integer)
    If ColIndex <> 1 Then
        Cancel = 1
    End If
End Sub
Private Sub dbgrdItensOrçamento_ColResize(ByVal ColIndex As Integer, Cancel As Integer)
    Cancel = 1
End Sub
Private Sub dbgrdItensOrçamento_DblClick()
    Dim pCódigo As String, Cont As Byte
    Dim pQuantidade As Integer
    Dim pValor As Currency
    Dim pDesconto As Currency
    
    dbgrdItensOrçamento.Col = MAXCOLS - 1
    pCódigo = dbgrdItensOrçamento.Text
    
    If mTotalRowsDevolução > 0 Then
        For Cont = 0 To mTotalRowsDevolução - 1
            If dbgrdItensDevoluçãoArray(MAXCOLS - 1, Cont) = pCódigo Then
                MsgBox "O item já foi incluído na tabela de devolução!", vbInformation, "Aviso"
                Exit Sub
            End If
        Next
    End If
    
    mTotalRowsDevolução = mTotalRowsDevolução + 1
    ReDim Preserve dbgrdItensDevoluçãoArray(MAXCOLS - 1, mTotalRowsDevolução - 1)
    
    For Cont = 0 To MAXCOLS - 1
        dbgrdItensOrçamento.Col = Cont
        If Cont = 1 Then
            dbgrdItensDevoluçãoArray(Cont, mTotalRowsDevolução - 1) = 1
        Else
            dbgrdItensDevoluçãoArray(Cont, mTotalRowsDevolução - 1) = dbgrdItensOrçamento.Text
        End If
    Next
    
    pQuantidade = 1
    
    dbgrdItensOrçamento.Col = 2
    pValor = ValStr(dbgrdItensOrçamento.Text)
    
    dbgrdItensOrçamento.Col = 3
    pDesconto = dbgrdItensOrçamento.Text
    
    pValor = pValor * pQuantidade
    
    pValor = pValor - (pValor * pDesconto / 100)
    
    dbgrdItensDevoluçãoArray(4, mTotalRowsDevolução - 1) = FormatStringMask("@V ##.###.##0,00", StrVal(pValor))
    
    dbgrdItensDevolução.ReBind
    dbgrdItensDevolução.Refresh
End Sub
Private Sub dbgrdItensOrçamento_RowResize(Cancel As Integer)
    Cancel = 1
End Sub
Private Sub dbgrdItensOrçamento_UnboundAddData(ByVal RowBuf As MSDBGrid.RowBuffer, NewRowBookmark As Variant)
    Dim Col%
        
    mTotalRowsOrçamento = mTotalRowsOrçamento + 1
    ReDim Preserve dbgrdItensOrçamentoArray(MAXCOLS - 1, mTotalRowsOrçamento - 1)
    
    'Sets the bookmark to the last row.
    NewRowBookmark = mTotalRowsOrçamento - 1
    
    ' The following loop adds a new record to the database.
    For Col = 0 To UBound(dbgrdItensOrçamentoArray, 1)
        If Not IsNull(RowBuf.Value(0, Col)) Then
            dbgrdItensOrçamentoArray(Col, mTotalRowsOrçamento - 1) = RowBuf.Value(0, Col)
        Else
            ' If no value set for column, then use the
            ' DefaultValue
            dbgrdItensOrçamentoArray(Col, mTotalRowsOrçamento - 1) = dbgrdItensOrçamento.Columns(Col).DefaultValue
        End If
    Next
End Sub
Private Sub dbgrdItensOrçamento_UnboundDeleteRow(Bookmark As Variant)
    Dim iCol As Integer, iRow As Integer
    
    ' Move all rows above the deleted row down in the
    ' array.
    
    For iRow = Bookmark + 1 To mTotalRowsOrçamento - 1
        For iCol = 0 To MAXCOLS - 1
            dbgrdItensOrçamentoArray(iCol, iRow - 1) = dbgrdItensOrçamentoArray(iCol, iRow)
        Next iCol
    Next iRow
    
    mTotalRowsOrçamento = mTotalRowsOrçamento - 1
End Sub
Private Sub dbgrdItensOrçamento_UnboundReadData(ByVal RowBuf As MSDBGrid.RowBuffer, StartLocation As Variant, ByVal ReadPriorRows As Boolean)
    Dim CurRow&, iRow As Integer, iCol As Integer, iRowsFetched As Integer, iIncr As Integer
    ' DBGrid is requesting rows so give them to it
    
    If mTotalRowsOrçamento = 0 Then Exit Sub
    
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
        If CurRow < 0 Or CurRow >= mTotalRowsOrçamento Then Exit For
        For iCol = 0 To UBound(dbgrdItensOrçamentoArray, 1)
            RowBuf.Value(iRow, iCol) = dbgrdItensOrçamentoArray(iCol, CurRow&)
        Next iCol
        ' Set bookmark using CurRow& which is also our
        ' array index
        RowBuf.Bookmark(iRow) = CStr(CurRow)
        CurRow = CurRow + iIncr
        iRowsFetched = iRowsFetched + 1
    Next iRow
    RowBuf.RowCount = iRowsFetched
End Sub
Private Sub dbgrdItensOrçamento_UnboundWriteData(ByVal RowBuf As MSDBGrid.RowBuffer, WriteLocation As Variant)
    Dim iCol As Integer
    ' Data is being updated
    'MsgBox WriteLocation
    ' Update each column in the data set array
    For iCol = 0 To MAXCOLS - 1
        If Not IsNull(RowBuf.Value(0, iCol)) Then
            dbgrdItensOrçamentoArray(iCol, WriteLocation) = RowBuf.Value(0, iCol)
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
    If Not ParâmetrosAberto Then
        Unload Me
        Exit Sub
    End If
    
    
    NavegaçãoInferior False
    NavegaçãoSuperior False
    
    BotãoGravar False
    BotãoExcluir False
    BotãoImprimir False
    
    BotãoIncluir lAllowInsert
    
    BarraDeStatus StatusBarAviso

    If lAtualizar Then
        BotãoAtualizar True
    Else
        BotãoAtualizar False
    End If
End Sub
Private Sub Form_Deactivate()
    cmdGravar.Enabled = False
    BotãoImprimir False
End Sub
Private Sub Form_Load()
    On Error GoTo Erro
    
    Dim Cont%
    
    lAtualizar = False
    
    lAllowInsert = Allow("DEVOLUÇÃO/TROCA (VENDA)", "I")
    lAllowEdit = Allow("DEVOLUÇÃO/TROCA (VENDA)", "A")
    lAllowDelete = Allow("DEVOLUÇÃO/TROCA (VENDA)", "E")
    lAllowConsult = Allow("DEVOLUÇÃO/TROCA (VENDA)", "C")
    
    ZeraCampos
    
    VendasAberto = AbreTabela(Dicionário, "FINANCEIRO", "VENDA", DBFinanceiro, TBLVendas, TBLTabela, dbOpenTable)
    
    If VendasAberto Then
        IndiceVendasAtivo = "VENDA1"
        TBLVendas.Index = IndiceVendasAtivo
    Else
        MsgBox "Não consegui abrir a tabela 'Vendas' !", vbCritical, "Erro"
        Exit Sub
    End If
    
    VendasItensAberto = AbreTabela(Dicionário, "FINANCEIRO", "VENDA - ITENS", DBFinanceiro, TBLVendasItens, TBLTabela, dbOpenTable)
    
    If VendasItensAberto Then
        IndiceVendasItensAtivo = "VENDAITENS1"
        TBLVendasItens.Index = IndiceVendasItensAtivo
    Else
        MsgBox "Não consegui abrir a tabela 'Itens de Venda' !", vbCritical, "Erro"
        Exit Sub
    End If
    
    VendasDevoluçãoAberto = AbreTabela(Dicionário, "FINANCEIRO", "VENDA - DEVOLUÇÃO", DBFinanceiro, TBLVendasDevolução, TBLTabela, dbOpenTable)
    
    If VendasDevoluçãoAberto Then
'        IndiceVendasAtivo = "VENDADEVOLUÇÃO1"
'        TBLVendas.Index = IndiceVendasAtivo
    Else
        MsgBox "Não consegui abrir a tabela 'Vendas' !", vbCritical, "Erro"
        Exit Sub
    End If
    
    VendasDevoluçãoItensAberto = AbreTabela(Dicionário, "FINANCEIRO", "VENDA - DEVOLUÇÃO - ITENS", DBFinanceiro, TBLVendasDevoluçãoItens, TBLTabela, dbOpenTable)
    
    If VendasDevoluçãoItensAberto Then
'        IndiceVendasItensAtivo = "VENDAITENS1"
'        TBLVendasItens.Index = IndiceVendasItensAtivo
    Else
        MsgBox "Não consegui abrir a tabela 'Itens de Venda' !", vbCritical, "Erro"
        Exit Sub
    End If
       
    ParâmetrosAberto = AbreTabela(Dicionário, "SISTEMA", "PARÂMETROS", DBSistema, TBLParâmetros, TBLTabela, dbOpenTable)
    
    If ParâmetrosAberto Then
    Else
        MsgBox "Não consegui abrir a tabela 'Parâmetros' !", vbCritical, "Erro"
        Exit Sub
    End If
    
    'Monta a grade de Orçamento
    dbgrdItensOrçamento.Columns.Add 1
    dbgrdItensOrçamento.Columns.Add 1
    dbgrdItensOrçamento.Columns.Add 1
    dbgrdItensOrçamento.Columns.Add 1
    
    For Cont = 0 To dbgrdItensOrçamento.Columns.Count - 1
        dbgrdItensOrçamento.Columns(Cont).Visible = True
    Next
       
    dbgrdItensOrçamento.Columns(0).Caption = "Produto"
    dbgrdItensOrçamento.Columns(0).Width = 3045
    dbgrdItensOrçamento.Columns(0).DefaultValue = " "
    dbgrdItensOrçamento.Columns(0).Alignment = dbgLeft
    
    dbgrdItensOrçamento.Columns(1).Caption = "Quantidade"
    dbgrdItensOrçamento.Columns(1).Width = 1000
    dbgrdItensOrçamento.Columns(1).DefaultValue = "0"
    dbgrdItensOrçamento.Columns(1).Alignment = dbgRight
    
    dbgrdItensOrçamento.Columns(2).Caption = "Valor Unitário"
    dbgrdItensOrçamento.Columns(2).Width = 1910
    dbgrdItensOrçamento.Columns(2).DefaultValue = "0,00"
    dbgrdItensOrçamento.Columns(2).Alignment = dbgRight
    
    dbgrdItensOrçamento.Columns(3).Caption = "Desconto"
    dbgrdItensOrçamento.Columns(3).Width = 1000
    dbgrdItensOrçamento.Columns(3).DefaultValue = "0,00"
    dbgrdItensOrçamento.Columns(3).Alignment = dbgRight
    
    dbgrdItensOrçamento.Columns(4).Caption = "Valor Total"
    dbgrdItensOrçamento.Columns(4).Width = 1910
    dbgrdItensOrçamento.Columns(4).DefaultValue = "0,00"
    dbgrdItensOrçamento.Columns(4).Alignment = dbgRight
    
    dbgrdItensOrçamento.Columns(5).Caption = "" 'Código do Produto
    dbgrdItensOrçamento.Columns(5).Width = 1
    dbgrdItensOrçamento.Columns(5).DefaultValue = "0"
    
    dbgrdItensOrçamento.ReBind
    dbgrdItensOrçamento.Refresh
    
    'Monta a grade de Devolução
    dbgrdItensDevolução.Columns.Add 1
    dbgrdItensDevolução.Columns.Add 1
    dbgrdItensDevolução.Columns.Add 1
    dbgrdItensDevolução.Columns.Add 1
    
    For Cont = 0 To dbgrdItensDevolução.Columns.Count - 1
        dbgrdItensDevolução.Columns(Cont).Visible = True
    Next
       
    dbgrdItensDevolução.Columns(0).Caption = "Produto"
    dbgrdItensDevolução.Columns(0).Width = 3045
    dbgrdItensDevolução.Columns(0).DefaultValue = " "
    dbgrdItensDevolução.Columns(0).Alignment = dbgLeft
    
    dbgrdItensDevolução.Columns(1).Caption = "Quantidade"
    dbgrdItensDevolução.Columns(1).Width = 1000
    dbgrdItensDevolução.Columns(1).DefaultValue = "0"
    dbgrdItensDevolução.Columns(1).Alignment = dbgRight
    
    dbgrdItensDevolução.Columns(2).Caption = "Valor Unitário"
    dbgrdItensDevolução.Columns(2).Width = 1910
    dbgrdItensDevolução.Columns(2).DefaultValue = "0,00"
    dbgrdItensDevolução.Columns(2).Alignment = dbgRight
    
    dbgrdItensDevolução.Columns(3).Caption = "Desconto"
    dbgrdItensDevolução.Columns(3).Width = 1000
    dbgrdItensDevolução.Columns(3).DefaultValue = "0,00"
    dbgrdItensDevolução.Columns(3).Alignment = dbgRight
    
    dbgrdItensDevolução.Columns(4).Caption = "Valor Total"
    dbgrdItensDevolução.Columns(4).Width = 1910
    dbgrdItensDevolução.Columns(4).DefaultValue = "0,00"
    dbgrdItensDevolução.Columns(4).Alignment = dbgRight
    
    dbgrdItensDevolução.Columns(5).Caption = "" 'Código do Produto
    dbgrdItensDevolução.Columns(5).Width = 1
    dbgrdItensDevolução.Columns(5).DefaultValue = "0"
    
    dbgrdItensDevolução.ReBind
    dbgrdItensDevolução.Refresh
    
    NavegaçãoInferior False
        
    StatusBarAviso = "Pronto"
    
    DesativaCampos
    
    Exit Sub
    
Erro:
    GeraMensagemDeErro "Devolução/Troca - Load"
    lFechar = True
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
    mdiGeal.StatusBar.Panels("Posição").Visible = False
    ResizeStatusBar
    
    Set frmVendaDevoluçãoTroca = Nothing
End Sub
Private Sub Form_Unload(Cancel As Integer)
    If VendasAberto Then
        TBLVendas.Close
    End If
    If VendasItensAberto Then
        TBLVendasItens.Close
    End If
    If ParâmetrosAberto Then
        TBLParâmetros.Close
    End If
    If Forms.Count = 2 Then
        AllBotões False
    End If
End Sub
