VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form frmVendaDevolu��oTroca 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Venda - Devolu��o/Troca"
   ClientHeight    =   6300
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9540
   Icon            =   "VendaDevolu��o.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6300
   ScaleWidth      =   9540
   Begin VB.Frame frDadosCadastraisOr�amento 
      Height          =   1140
      Left            =   4770
      TabIndex        =   14
      Top             =   0
      Width           =   4770
      Begin VB.CommandButton cmdLocalizarOr�amento 
         Caption         =   "&Localizar"
         Height          =   375
         Left            =   2940
         TabIndex        =   4
         Top             =   450
         Width           =   1335
      End
      Begin VB.TextBox txtOr�amento 
         Height          =   285
         Left            =   1170
         TabIndex        =   2
         Top             =   270
         Width           =   975
      End
      Begin VB.TextBox txtDataOr�amento 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1170
         TabIndex        =   3
         Text            =   "  /  /"
         Top             =   690
         Width           =   990
      End
      Begin VB.Label lblOr�amento 
         Caption         =   "Or�amento"
         Height          =   180
         Left            =   150
         TabIndex        =   16
         Top             =   330
         Width           =   825
      End
      Begin VB.Label lblDataDoOr�amento 
         Caption         =   "Data"
         Height          =   210
         Left            =   150
         TabIndex        =   15
         Top             =   720
         Width           =   465
      End
   End
   Begin VB.Frame frItensDevolu��o 
      Caption         =   "Itens para devolu��o"
      Height          =   2355
      Left            =   0
      TabIndex        =   13
      Top             =   3510
      Width           =   9540
      Begin MSDBGrid.DBGrid dbgrdItensDevolu��o 
         Height          =   2025
         Left            =   90
         OleObjectBlob   =   "VendaDevolu��o.frx":030A
         TabIndex        =   6
         Top             =   210
         Width           =   9405
      End
   End
   Begin VB.Frame frDadosCadastraisDevolu��o 
      Height          =   1140
      Left            =   0
      TabIndex        =   10
      Top             =   0
      Width           =   4770
      Begin VB.TextBox txtDataDevolu��o 
         Enabled         =   0   'False
         Height          =   285
         Left            =   3510
         TabIndex        =   1
         Text            =   "  /  /"
         Top             =   510
         Width           =   990
      End
      Begin VB.TextBox txtDevolu��o 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1170
         TabIndex        =   0
         Top             =   510
         Width           =   975
      End
      Begin VB.Label lblDataDaDevolu��o 
         Caption         =   "Data"
         Height          =   210
         Left            =   2880
         TabIndex        =   12
         Top             =   540
         Width           =   465
      End
      Begin VB.Label lblDevolu��o 
         Caption         =   "Devolu��o"
         Height          =   180
         Left            =   150
         TabIndex        =   11
         Top             =   540
         Width           =   825
      End
   End
   Begin VB.Frame frItensOr�amento 
      Caption         =   "Itens do Or�amento"
      Height          =   2355
      Left            =   0
      TabIndex        =   9
      Top             =   1140
      Width           =   9540
      Begin MSDBGrid.DBGrid dbgrdItensOr�amento 
         Height          =   2025
         Left            =   60
         OleObjectBlob   =   "VendaDevolu��o.frx":0CE9
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
Attribute VB_Name = "frmVendaDevolu��oTroca"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const MAXCOLS = 6

Const vbDescri��o = 1
Const vbC�digo = 2
Const vbValorUnit�rio = 3
Const vbValValorUnit�rio = 4
Const vbIndice1 = 1
Const vbIndice2 = 2
Const vbIndice3 = 3
Const vbIndice4 = 4

Dim mUsu�rio As String

Dim mRecnoDevolu��o%
Dim mTotalRowsDevolu��o%
Dim dbgrdItensDevolu��oArray() As String
Dim ArrayVendasItensRecnoDevolu��o() As Variant

Dim mRecnoOr�amento%
Dim mTotalRowsOr�amento%
Dim dbgrdItensOr�amentoArray() As String
Dim ArrayVendasItensRecnoOr�amento() As Variant

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

Dim TBLPar�metros As Table
Dim Par�metrosAberto As Boolean

Dim TBLVendasDevolu��o As Table
Dim VendasDevolu��oAberto As Boolean
Dim IndiceVendasDevolu��oAtivo$

Dim TBLVendasDevolu��oItens As Table
Dim VendasDevolu��oItensAberto As Boolean
Dim IndiceVendasDevolu��oItensAtivo$

Dim TBLValeCliente As Table
Dim ValeClienteAberto As Boolean
Dim IndiceValeCliente$

Dim StatusBarAviso$

Dim DataBaseName(1 To 1) As String
Public Relat�rio$
Public TotalDatabaseName%

Public lAtualizar As Boolean
Private Sub AtivaCampos()
    frDadosCadastraisDevolu��o.Enabled = True
    frDadosCadastraisOr�amento.Enabled = True
    frItensOr�amento.Enabled = True
    frItensDevolu��o.Enabled = True
    'cmdIncluirItem.Enabled = True
    Bot�oGravar (lInserir)
    cmdCancelar.Enabled = (lInserir)
    cmdGravar.Enabled = (lInserir)
End Sub
Private Function Cancelamento() As Boolean
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
    
    BarraDeStatus StatusBarAviso
    
    Bot�oIncluir lAllowInsert
    
    Bot�oGravar False
    DesativaCampos
    ZeraCampos
    Cancelamento = True
    lInserir = False
End Function
Private Sub DesativaCampos()
    frDadosCadastraisDevolu��o.Enabled = False
    frDadosCadastraisOr�amento.Enabled = False
    frItensOr�amento.Enabled = False
    frItensDevolu��o.Enabled = False
    cmdGravar.Enabled = False
    cmdCancelar.Enabled = False
    'cmdIncluirItem.Enabled = False
    cmdGravar.Enabled = False
    cmdCancelar.Enabled = False
    Bot�oGravar False
End Sub
Private Sub FillGridOr�amento(ByVal Chave As Long)
    Dim pValor As Currency
    Dim pDesconto As Currency
    
    dbgrdItensOr�amento.ReBind
    
    ReDim dbgrdItensOr�amentoArray(MAXCOLS - 1, 0)
    ReDim ArrayVendasItensRecnoOr�amento(0)
    
    mTotalRowsOr�amento = 0
    mRecnoOr�amento = 0
    
    TBLVendasItens.Seek "=", Chave
    If Not TBLVendasItens.NoMatch Then
        Do While Not TBLVendasItens.EOF And TBLVendasItens("OR�AMENTO") = Chave
            mRecnoOr�amento = mRecnoOr�amento + 1
            mTotalRowsOr�amento = mTotalRowsOr�amento + 1
            ReDim Preserve dbgrdItensOr�amentoArray(MAXCOLS - 1, mTotalRowsOr�amento - 1)
            ReDim Preserve ArrayVendasItensRecnoOr�amento(mTotalRowsOr�amento - 1)
            
            ArrayVendasItensRecnoOr�amento(mTotalRowsOr�amento - 1) = TBLVendasItens.Bookmark
            dbgrdItensOr�amentoArray(0, mTotalRowsOr�amento - 1) = SearchProduto(TBLVendasItens("C�DIGO DO PRODUTO")) 'Nome do Produto
            dbgrdItensOr�amentoArray(1, mTotalRowsOr�amento - 1) = FormatStringMask("@V ######0", StrVal(TBLVendasItens("QUANTIDADE"))) 'Quantidade
            dbgrdItensOr�amentoArray(2, mTotalRowsOr�amento - 1) = FormatStringMask("@V ##.###.##0,00", StrVal(TBLVendasItens("VALOR UNIT�RIO"))) 'Pre�o Unit�rio
            dbgrdItensOr�amentoArray(3, mTotalRowsOr�amento - 1) = FormatStringMask("@V ##.###.##0,00", StrVal(TBLVendasItens("DESCONTO"))) 'Desconto no valor do produto
            
            pValor = TBLVendasItens("VALOR UNIT�RIO") * TBLVendasItens("QUANTIDADE")
            pValor = pValor - (pValor * TBLVendasItens("DESCONTO") / 100)
            dbgrdItensOr�amentoArray(4, mTotalRowsOr�amento - 1) = FormatStringMask("@V ##.###.##0,00", StrVal(pValor)) 'Valor total
            
            dbgrdItensOr�amentoArray(5, mTotalRowsOr�amento - 1) = TBLVendasItens("C�DIGO DO PRODUTO") 'C�digo do Produto

            TBLVendasItens.MoveNext

            If TBLVendasItens.EOF Then
                Exit Do
            End If
        Loop
    End If
    
    dbgrdItensOr�amento.Refresh
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
    
    Bot�oGravar (lInserir)
    Bot�oIncluir False
    Bot�oGravar (lInserir)
    cmdCancelar.Enabled = (lInserir)
    
    StatusBarAviso = "Inclus�o"
    BarraDeStatus StatusBarAviso
    
    txtOr�amento.SetFocus
End Sub
Private Sub Localizar()
    If PosRecords Then
        GetRecords
    End If
End Sub
Private Function PosRecords() As Boolean
    TBLVendas.Seek "=", Val(txtOr�amento)
    If TBLVendas.NoMatch Then
        PosRecords = False
        MsgBox "N�o encontrei o or�amento " & txtOr�amento, vbInformation, "Aviso"
    Else
        If TBLVendas("TIPO") <> "V" Then
            MsgBox "Este or�amento n�o pode ser editado!", vbInformation, "Aviso"
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
    
    txtDataOr�amento = TBLVendas("DATA DO OR�AMENTO")
    CorrigeData DataMask, txtDataOr�amento, TBLVendas("DATA DO OR�AMENTO")
    
    FillGridOr�amento TBLVendas("C�DIGO")
    
    lPula = False
    If Not lAllowEdit Then
        DesativaCampos
    End If
    If Not lAllowEdit Then
        DesativaCampos
    End If
End Sub
Private Function SetRecords()
    On Error GoTo ErroVendasDevolu��o
    
    Dim Cont As Integer
    
    WS.BeginTrans 'Inicia uma Transa��o
            
    TBLPar�metros.Edit
    TBLPar�metros("DEVOLU��O") = TBLPar�metros("DEVOLU��O") + 1
    TBLPar�metros.Update
    TBLPar�metros.MoveFirst
    
    TBLVendasDevolu��o.AddNew
    
    TBLVendasDevolu��o("C�DIGO") = TBLPar�metros("DEVOLU��O")
    TBLVendasDevolu��o("OR�AMENTO") = txtOr�amento
    TBLVendasDevolu��o("MOTIVO") = 0
    TBLVendasDevolu��o("OBSERVA��O") = vbNull
    TBLVendasDevolu��o("BAIXADO") = False
    
    TBLVendasDevolu��o("USERNAME - CRIA") = mUsu�rio
    TBLVendasDevolu��o("DATA - CRIA") = Date
    TBLVendasDevolu��o("HORA - CRIA") = Time
    TBLVendasDevolu��o("USERNAME - ALTERA") = "VAZIO"
    TBLVendasDevolu��o("DATA - ALTERA") = vbNull
    TBLVendasDevolu��o("HORA - ALTERA") = vbNull
    
    TBLVendasDevolu��o.Update
            
    On Error GoTo ErroVendasDevolu��oItens
    
    For Cont = 0 To mTotalRowsDevolu��o - 1
        TBLVendasDevolu��oItens.AddNew
        
        TBLVendasDevolu��oItens("C�DIGO DA DEVOLU��O") = TBLPar�metros("DEVOLU��O")
        TBLVendasDevolu��oItens("C�DIGO DO PRODUTO") = dbgrdItensDevolu��oArray(5, Cont)
        TBLVendasDevolu��oItens("QUANTIDADE") = dbgrdItensDevolu��oArray(1, Cont)
        TBLVendasDevolu��oItens("DESTINO") = vbNull
        TBLVendasDevolu��oItens("BAIXADO") = vbNull
        
        TBLVendasDevolu��oItens.Update
    Next
    
    WS.CommitTrans 'Grava as altera��es ou inclus�es se n�o houverem erros
    
    If lInserir Then
        Log gUsu�rio, "Inclus�o - (Venda) Devolu��o/Troca: " & txtOr�amento
    Else
        Log gUsu�rio, "Altera��o - (Venda) Devolu��o/Troca: " & txtOr�amento
    End If
    
    SetRecords = True
    
    Exit Function
    
ErroVendasDevolu��o:
    TBLVendasDevolu��o.CancelUpdate
    GeraMensagemDeErro "Venda - Devolu��o/Troca - SetRecords - ErroVendasDevolu��o - " & txtOr�amento, True
    SetRecords = False
    Exit Function
    
ErroVendasDevolu��oItens:
    TBLVendasDevolu��oItens.CancelUpdate
    GeraMensagemDeErro "Venda - Devolu��o/Troca - SetRecords - ErroVendasDevolu��oItens - " & txtOr�amento, True
    SetRecords = False
    
    Exit Function
End Function
Private Sub ZeraCampos()
    On Error Resume Next

    lPula = True
    lInserir = False
    
    txtDevolu��o = Empty
    txtDataDevolu��o = Empty
    
    txtOr�amento = Empty
    txtDataOr�amento = Empty
    
    ReDim dbgrdItensOr�amentoArray(MAXCOLS - 1, 0)
    ReDim dbgrdItensDevolu��oArray(MAXCOLS - 1, 0)
    
    mTotalRowsOr�amento = 0
    mTotalRowsDevolu��o = 0
    
    mRecnoOr�amento = 0
    mRecnoDevolu��o = 0
    
    dbgrdItensOr�amento.ReBind
    dbgrdItensDevolu��o.ReBind
    
    lPula = False
End Sub
Private Sub cmdGravar_Click()
    'Valida Usu�rio
    frmValidaUsu�rio.Show 1
    
    mUsu�rio = frmValidaUsu�rio.Usu�rio
    
    Set frmValidaUsu�rio = Nothing
    
    If mUsu�rio = Empty Then
        Exit Sub
    End If
    
    Gravar
End Sub
Private Sub cmdCancelar_Click()
    Cancelamento
End Sub
Private Sub cmdLocalizarOr�amento_Click()
    Localizar
End Sub
Private Sub dbgrdItensDevolu��o_BeforeColUpdate(ByVal ColIndex As Integer, oldvalue As Variant, Cancel As Integer)
    Dim pC�digo As String, pQuantidade As Integer, pQuantidadeAtual As Integer
    Dim pValor As Currency, pDesconto As Currency
    Dim Cont As Integer
    
    If ColIndex = 1 Then
        pQuantidadeAtual = dbgrdItensDevolu��o.Text
        dbgrdItensDevolu��o.Col = 5
        pC�digo = dbgrdItensDevolu��o.Text
        For Cont = 0 To mTotalRowsOr�amento - 1
            If dbgrdItensOr�amentoArray(5, Cont) = pC�digo Then
                pQuantidade = dbgrdItensOr�amentoArray(1, Cont)
                pValor = ValStr(dbgrdItensOr�amentoArray(2, Cont))
                pDesconto = ValStr(dbgrdItensOr�amentoArray(3, Cont))
                Exit For
            End If
        Next
        
        If pQuantidade < pQuantidadeAtual Then
            MsgBox "A quantidade de devolu��o n�o pode ser maior que a quantidade de compra!", vbCritical, "Aviso"
            Cancel = 1
            Exit Sub
        End If
        
                
        pValor = pQuantidadeAtual * pValor
        pValor = pValor - (pValor * (pDesconto / 100))
        
        dbgrdItensDevolu��o.Col = 4
        dbgrdItensDevolu��o.Text = FormatStringMask("@V ##.###.##0,00", StrVal(pValor))
        
        dbgrdItensDevolu��o.Col = 1
        dbgrdItensDevolu��o.Text = pQuantidadeAtual
    End If
End Sub
Private Sub dbgrdItensDevolu��o_UnboundAddData(ByVal RowBuf As MSDBGrid.RowBuffer, NewRowBookmark As Variant)
    Dim Col%
        
    mTotalRowsDevolu��o = mTotalRowsDevolu��o + 1
    ReDim Preserve dbgrdItensDevolu��oArray(MAXCOLS - 1, mTotalRowsDevolu��o - 1)
    
    'Sets the bookmark to the last row.
    NewRowBookmark = mTotalRowsDevolu��o - 1
    
    ' The following loop adds a new record to the database.
    For Col = 0 To UBound(dbgrdItensDevolu��oArray, 1)
        If Not IsNull(RowBuf.Value(0, Col)) Then
            dbgrdItensDevolu��oArray(Col, mTotalRowsDevolu��o - 1) = RowBuf.Value(0, Col)
        Else
            ' If no value set for column, then use the
            ' DefaultValue
            dbgrdItensDevolu��oArray(Col, mTotalRowsDevolu��o - 1) = dbgrdItensDevolu��o.Columns(Col).DefaultValue
        End If
    Next
End Sub
Private Sub dbgrdItensDevolu��o_UnboundDeleteRow(Bookmark As Variant)
    Dim iCol As Integer, iRow As Integer
    
    ' Move all rows above the deleted row down in the
    ' array.
    
    For iRow = Bookmark + 1 To mTotalRowsDevolu��o - 1
        For iCol = 0 To MAXCOLS - 1
            dbgrdItensDevolu��oArray(iCol, iRow - 1) = dbgrdItensDevolu��oArray(iCol, iRow)
        Next iCol
    Next iRow
    
    mTotalRowsDevolu��o = mTotalRowsDevolu��o - 1
End Sub
Private Sub dbgrdItensDevolu��o_UnboundReadData(ByVal RowBuf As MSDBGrid.RowBuffer, StartLocation As Variant, ByVal ReadPriorRows As Boolean)
    Dim CurRow&, iRow As Integer, iCol As Integer, iRowsFetched As Integer, iIncr As Integer
    ' DBGrid is requesting rows so give them to it
    
    If mTotalRowsDevolu��o = 0 Then Exit Sub
    
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
        If CurRow < 0 Or CurRow >= mTotalRowsDevolu��o Then Exit For
        For iCol = 0 To UBound(dbgrdItensDevolu��oArray, 1)
            RowBuf.Value(iRow, iCol) = dbgrdItensDevolu��oArray(iCol, CurRow&)
        Next iCol
        ' Set bookmark using CurRow& which is also our
        ' array index
        RowBuf.Bookmark(iRow) = CStr(CurRow)
        CurRow = CurRow + iIncr
        iRowsFetched = iRowsFetched + 1
    Next iRow
    RowBuf.RowCount = iRowsFetched
End Sub
Private Sub dbgrdItensDevolu��o_UnboundWriteData(ByVal RowBuf As MSDBGrid.RowBuffer, WriteLocation As Variant)
    Dim iCol As Integer
    ' Data is being updated
    'MsgBox WriteLocation
    ' Update each column in the data set array
    For iCol = 0 To MAXCOLS - 1
        If Not IsNull(RowBuf.Value(0, iCol)) Then
            dbgrdItensDevolu��oArray(iCol, WriteLocation) = RowBuf.Value(0, iCol)
        End If
    Next iCol
End Sub
Private Sub dbgrdItensOr�amento_BeforeColEdit(ByVal ColIndex As Integer, ByVal KeyAscii As Integer, Cancel As Integer)
    If ColIndex <> 1 Then
        Cancel = 1
    End If
End Sub
Private Sub dbgrdItensOr�amento_ColResize(ByVal ColIndex As Integer, Cancel As Integer)
    Cancel = 1
End Sub
Private Sub dbgrdItensOr�amento_DblClick()
    Dim pC�digo As String, Cont As Byte
    Dim pQuantidade As Integer
    Dim pValor As Currency
    Dim pDesconto As Currency
    
    dbgrdItensOr�amento.Col = MAXCOLS - 1
    pC�digo = dbgrdItensOr�amento.Text
    
    If mTotalRowsDevolu��o > 0 Then
        For Cont = 0 To mTotalRowsDevolu��o - 1
            If dbgrdItensDevolu��oArray(MAXCOLS - 1, Cont) = pC�digo Then
                MsgBox "O item j� foi inclu�do na tabela de devolu��o!", vbInformation, "Aviso"
                Exit Sub
            End If
        Next
    End If
    
    mTotalRowsDevolu��o = mTotalRowsDevolu��o + 1
    ReDim Preserve dbgrdItensDevolu��oArray(MAXCOLS - 1, mTotalRowsDevolu��o - 1)
    
    For Cont = 0 To MAXCOLS - 1
        dbgrdItensOr�amento.Col = Cont
        If Cont = 1 Then
            dbgrdItensDevolu��oArray(Cont, mTotalRowsDevolu��o - 1) = 1
        Else
            dbgrdItensDevolu��oArray(Cont, mTotalRowsDevolu��o - 1) = dbgrdItensOr�amento.Text
        End If
    Next
    
    pQuantidade = 1
    
    dbgrdItensOr�amento.Col = 2
    pValor = ValStr(dbgrdItensOr�amento.Text)
    
    dbgrdItensOr�amento.Col = 3
    pDesconto = dbgrdItensOr�amento.Text
    
    pValor = pValor * pQuantidade
    
    pValor = pValor - (pValor * pDesconto / 100)
    
    dbgrdItensDevolu��oArray(4, mTotalRowsDevolu��o - 1) = FormatStringMask("@V ##.###.##0,00", StrVal(pValor))
    
    dbgrdItensDevolu��o.ReBind
    dbgrdItensDevolu��o.Refresh
End Sub
Private Sub dbgrdItensOr�amento_RowResize(Cancel As Integer)
    Cancel = 1
End Sub
Private Sub dbgrdItensOr�amento_UnboundAddData(ByVal RowBuf As MSDBGrid.RowBuffer, NewRowBookmark As Variant)
    Dim Col%
        
    mTotalRowsOr�amento = mTotalRowsOr�amento + 1
    ReDim Preserve dbgrdItensOr�amentoArray(MAXCOLS - 1, mTotalRowsOr�amento - 1)
    
    'Sets the bookmark to the last row.
    NewRowBookmark = mTotalRowsOr�amento - 1
    
    ' The following loop adds a new record to the database.
    For Col = 0 To UBound(dbgrdItensOr�amentoArray, 1)
        If Not IsNull(RowBuf.Value(0, Col)) Then
            dbgrdItensOr�amentoArray(Col, mTotalRowsOr�amento - 1) = RowBuf.Value(0, Col)
        Else
            ' If no value set for column, then use the
            ' DefaultValue
            dbgrdItensOr�amentoArray(Col, mTotalRowsOr�amento - 1) = dbgrdItensOr�amento.Columns(Col).DefaultValue
        End If
    Next
End Sub
Private Sub dbgrdItensOr�amento_UnboundDeleteRow(Bookmark As Variant)
    Dim iCol As Integer, iRow As Integer
    
    ' Move all rows above the deleted row down in the
    ' array.
    
    For iRow = Bookmark + 1 To mTotalRowsOr�amento - 1
        For iCol = 0 To MAXCOLS - 1
            dbgrdItensOr�amentoArray(iCol, iRow - 1) = dbgrdItensOr�amentoArray(iCol, iRow)
        Next iCol
    Next iRow
    
    mTotalRowsOr�amento = mTotalRowsOr�amento - 1
End Sub
Private Sub dbgrdItensOr�amento_UnboundReadData(ByVal RowBuf As MSDBGrid.RowBuffer, StartLocation As Variant, ByVal ReadPriorRows As Boolean)
    Dim CurRow&, iRow As Integer, iCol As Integer, iRowsFetched As Integer, iIncr As Integer
    ' DBGrid is requesting rows so give them to it
    
    If mTotalRowsOr�amento = 0 Then Exit Sub
    
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
        If CurRow < 0 Or CurRow >= mTotalRowsOr�amento Then Exit For
        For iCol = 0 To UBound(dbgrdItensOr�amentoArray, 1)
            RowBuf.Value(iRow, iCol) = dbgrdItensOr�amentoArray(iCol, CurRow&)
        Next iCol
        ' Set bookmark using CurRow& which is also our
        ' array index
        RowBuf.Bookmark(iRow) = CStr(CurRow)
        CurRow = CurRow + iIncr
        iRowsFetched = iRowsFetched + 1
    Next iRow
    RowBuf.RowCount = iRowsFetched
End Sub
Private Sub dbgrdItensOr�amento_UnboundWriteData(ByVal RowBuf As MSDBGrid.RowBuffer, WriteLocation As Variant)
    Dim iCol As Integer
    ' Data is being updated
    'MsgBox WriteLocation
    ' Update each column in the data set array
    For iCol = 0 To MAXCOLS - 1
        If Not IsNull(RowBuf.Value(0, iCol)) Then
            dbgrdItensOr�amentoArray(iCol, WriteLocation) = RowBuf.Value(0, iCol)
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
    If Not Par�metrosAberto Then
        Unload Me
        Exit Sub
    End If
    
    
    Navega��oInferior False
    Navega��oSuperior False
    
    Bot�oGravar False
    Bot�oExcluir False
    Bot�oImprimir False
    
    Bot�oIncluir lAllowInsert
    
    BarraDeStatus StatusBarAviso

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
    
    lAtualizar = False
    
    lAllowInsert = Allow("DEVOLU��O/TROCA (VENDA)", "I")
    lAllowEdit = Allow("DEVOLU��O/TROCA (VENDA)", "A")
    lAllowDelete = Allow("DEVOLU��O/TROCA (VENDA)", "E")
    lAllowConsult = Allow("DEVOLU��O/TROCA (VENDA)", "C")
    
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
    
    VendasDevolu��oAberto = AbreTabela(Dicion�rio, "FINANCEIRO", "VENDA - DEVOLU��O", DBFinanceiro, TBLVendasDevolu��o, TBLTabela, dbOpenTable)
    
    If VendasDevolu��oAberto Then
'        IndiceVendasAtivo = "VENDADEVOLU��O1"
'        TBLVendas.Index = IndiceVendasAtivo
    Else
        MsgBox "N�o consegui abrir a tabela 'Vendas' !", vbCritical, "Erro"
        Exit Sub
    End If
    
    VendasDevolu��oItensAberto = AbreTabela(Dicion�rio, "FINANCEIRO", "VENDA - DEVOLU��O - ITENS", DBFinanceiro, TBLVendasDevolu��oItens, TBLTabela, dbOpenTable)
    
    If VendasDevolu��oItensAberto Then
'        IndiceVendasItensAtivo = "VENDAITENS1"
'        TBLVendasItens.Index = IndiceVendasItensAtivo
    Else
        MsgBox "N�o consegui abrir a tabela 'Itens de Venda' !", vbCritical, "Erro"
        Exit Sub
    End If
       
    Par�metrosAberto = AbreTabela(Dicion�rio, "SISTEMA", "PAR�METROS", DBSistema, TBLPar�metros, TBLTabela, dbOpenTable)
    
    If Par�metrosAberto Then
    Else
        MsgBox "N�o consegui abrir a tabela 'Par�metros' !", vbCritical, "Erro"
        Exit Sub
    End If
    
    'Monta a grade de Or�amento
    dbgrdItensOr�amento.Columns.Add 1
    dbgrdItensOr�amento.Columns.Add 1
    dbgrdItensOr�amento.Columns.Add 1
    dbgrdItensOr�amento.Columns.Add 1
    
    For Cont = 0 To dbgrdItensOr�amento.Columns.Count - 1
        dbgrdItensOr�amento.Columns(Cont).Visible = True
    Next
       
    dbgrdItensOr�amento.Columns(0).Caption = "Produto"
    dbgrdItensOr�amento.Columns(0).Width = 3045
    dbgrdItensOr�amento.Columns(0).DefaultValue = " "
    dbgrdItensOr�amento.Columns(0).Alignment = dbgLeft
    
    dbgrdItensOr�amento.Columns(1).Caption = "Quantidade"
    dbgrdItensOr�amento.Columns(1).Width = 1000
    dbgrdItensOr�amento.Columns(1).DefaultValue = "0"
    dbgrdItensOr�amento.Columns(1).Alignment = dbgRight
    
    dbgrdItensOr�amento.Columns(2).Caption = "Valor Unit�rio"
    dbgrdItensOr�amento.Columns(2).Width = 1910
    dbgrdItensOr�amento.Columns(2).DefaultValue = "0,00"
    dbgrdItensOr�amento.Columns(2).Alignment = dbgRight
    
    dbgrdItensOr�amento.Columns(3).Caption = "Desconto"
    dbgrdItensOr�amento.Columns(3).Width = 1000
    dbgrdItensOr�amento.Columns(3).DefaultValue = "0,00"
    dbgrdItensOr�amento.Columns(3).Alignment = dbgRight
    
    dbgrdItensOr�amento.Columns(4).Caption = "Valor Total"
    dbgrdItensOr�amento.Columns(4).Width = 1910
    dbgrdItensOr�amento.Columns(4).DefaultValue = "0,00"
    dbgrdItensOr�amento.Columns(4).Alignment = dbgRight
    
    dbgrdItensOr�amento.Columns(5).Caption = "" 'C�digo do Produto
    dbgrdItensOr�amento.Columns(5).Width = 1
    dbgrdItensOr�amento.Columns(5).DefaultValue = "0"
    
    dbgrdItensOr�amento.ReBind
    dbgrdItensOr�amento.Refresh
    
    'Monta a grade de Devolu��o
    dbgrdItensDevolu��o.Columns.Add 1
    dbgrdItensDevolu��o.Columns.Add 1
    dbgrdItensDevolu��o.Columns.Add 1
    dbgrdItensDevolu��o.Columns.Add 1
    
    For Cont = 0 To dbgrdItensDevolu��o.Columns.Count - 1
        dbgrdItensDevolu��o.Columns(Cont).Visible = True
    Next
       
    dbgrdItensDevolu��o.Columns(0).Caption = "Produto"
    dbgrdItensDevolu��o.Columns(0).Width = 3045
    dbgrdItensDevolu��o.Columns(0).DefaultValue = " "
    dbgrdItensDevolu��o.Columns(0).Alignment = dbgLeft
    
    dbgrdItensDevolu��o.Columns(1).Caption = "Quantidade"
    dbgrdItensDevolu��o.Columns(1).Width = 1000
    dbgrdItensDevolu��o.Columns(1).DefaultValue = "0"
    dbgrdItensDevolu��o.Columns(1).Alignment = dbgRight
    
    dbgrdItensDevolu��o.Columns(2).Caption = "Valor Unit�rio"
    dbgrdItensDevolu��o.Columns(2).Width = 1910
    dbgrdItensDevolu��o.Columns(2).DefaultValue = "0,00"
    dbgrdItensDevolu��o.Columns(2).Alignment = dbgRight
    
    dbgrdItensDevolu��o.Columns(3).Caption = "Desconto"
    dbgrdItensDevolu��o.Columns(3).Width = 1000
    dbgrdItensDevolu��o.Columns(3).DefaultValue = "0,00"
    dbgrdItensDevolu��o.Columns(3).Alignment = dbgRight
    
    dbgrdItensDevolu��o.Columns(4).Caption = "Valor Total"
    dbgrdItensDevolu��o.Columns(4).Width = 1910
    dbgrdItensDevolu��o.Columns(4).DefaultValue = "0,00"
    dbgrdItensDevolu��o.Columns(4).Alignment = dbgRight
    
    dbgrdItensDevolu��o.Columns(5).Caption = "" 'C�digo do Produto
    dbgrdItensDevolu��o.Columns(5).Width = 1
    dbgrdItensDevolu��o.Columns(5).DefaultValue = "0"
    
    dbgrdItensDevolu��o.ReBind
    dbgrdItensDevolu��o.Refresh
    
    Navega��oInferior False
        
    StatusBarAviso = "Pronto"
    
    DesativaCampos
    
    Exit Sub
    
Erro:
    GeraMensagemDeErro "Devolu��o/Troca - Load"
    lFechar = True
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
    mdiGeal.StatusBar.Panels("Posi��o").Visible = False
    ResizeStatusBar
    
    Set frmVendaDevolu��oTroca = Nothing
End Sub
Private Sub Form_Unload(Cancel As Integer)
    If VendasAberto Then
        TBLVendas.Close
    End If
    If VendasItensAberto Then
        TBLVendasItens.Close
    End If
    If Par�metrosAberto Then
        TBLPar�metros.Close
    End If
    If Forms.Count = 2 Then
        AllBot�es False
    End If
End Sub
