VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form frmFormaDePagamento 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Forma de Pagamento"
   ClientHeight    =   2910
   ClientLeft      =   1605
   ClientTop       =   1530
   ClientWidth     =   6285
   Icon            =   "FormaDePagamento.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2910
   ScaleWidth      =   6285
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   345
      Left            =   5040
      TabIndex        =   9
      Top             =   2550
      Width           =   1245
   End
   Begin VB.CommandButton cmdGravar 
      Caption         =   "&Gravar"
      Height          =   345
      Left            =   3720
      TabIndex        =   8
      Top             =   2550
      Width           =   1245
   End
   Begin VB.Frame frDatas 
      Caption         =   "Vencimentos"
      Height          =   1365
      Left            =   0
      TabIndex        =   6
      Top             =   1125
      Width           =   6285
      Begin MSDBGrid.DBGrid dbgrdPagamento 
         Height          =   1125
         Left            =   30
         OleObjectBlob   =   "FormaDePagamento.frx":030A
         TabIndex        =   7
         Top             =   180
         Width           =   6225
      End
   End
   Begin VB.Frame frDados 
      Height          =   1110
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6285
      Begin VB.ComboBox cmbPlanoDePagamento 
         Height          =   315
         Left            =   1680
         Locked          =   -1  'True
         TabIndex        =   10
         Top             =   270
         Width           =   4485
      End
      Begin VB.TextBox txtValorAPrazo 
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   4485
         Locked          =   -1  'True
         TabIndex        =   5
         Top             =   660
         Width           =   1680
      End
      Begin VB.TextBox txtValorAVista 
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1680
         TabIndex        =   3
         Top             =   660
         Width           =   1680
      End
      Begin VB.Label lblValorAPrazo 
         Caption         =   "Valor à prazo"
         Height          =   180
         Left            =   3465
         TabIndex        =   4
         Top             =   720
         Width           =   945
      End
      Begin VB.Label lblValorAVista 
         Caption         =   "Valor à vista"
         Height          =   210
         Left            =   150
         TabIndex        =   2
         Top             =   720
         Width           =   1215
      End
      Begin VB.Label lblPlanoDePagamento 
         Caption         =   "Plano de Pagamento"
         Height          =   225
         Left            =   150
         TabIndex        =   1
         Top             =   330
         Width           =   1500
      End
   End
End
Attribute VB_Name = "frmFormaDePagamento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const MAXCOLS = 3

Dim mRecno%
Dim mTotalRows%

Dim mCódigoAtual%, mCarência%, mIntervaloDePagamentos%, mQuantidadeDeVencimentos As Integer
Dim mCustoFinaceiro As Currency, mRepasse As Currency
Dim mAutoInclusão As Boolean, mPermiteAlterarAutoInclusão As Boolean

Dim dbgrdPagamentoArray() As String
Dim CódigoDoPlanoDePagamento() As Variant

Dim StatusBarAviso As String

Dim lPula As Boolean

Public lAtualizar As Boolean

Public lInserir As Boolean
Public lAlterar As Boolean

Public mTotalPagamentos As Integer
Public mValorAVista As String
Public mValorAPrazo As String
Public mTipoDePagamento As Long
Public mData As String
Public ptrForm As Object

Public lEdit As Boolean
Public lNotCancel As Boolean
Public lCaixa As Boolean

Public lCompra As Boolean

Public TBLPlanoDePagamento As Recordset
Public Sub Atualizar()
    FillPlanoDePagamento
End Sub
Private Sub FillGridPg()
    Dim Cont As Integer, Cont1 As Integer
    
    dbgrdPagamento.ReBind
    
    mTotalRows = mTotalPagamentos
    If mTotalRows = 0 Then Exit Sub
    
    ReDim Preserve dbgrdPagamentoArray(MAXCOLS - 1, mTotalRows - 1)
    
    If mTotalPagamentos > 0 Then
        For Cont = 0 To mTotalPagamentos - 1
            For Cont1 = 0 To MAXCOLS - 1
                dbgrdPagamentoArray(Cont1, Cont) = ptrForm.GetPagamentos(Cont1, Cont)
            Next
        Next
    End If
    
    dbgrdPagamento.Refresh
End Sub
Private Sub FillPlanoDePagamento()
    Dim Tamanho%
    
    lPula = True
    
    If TBLPlanoDePagamento.RecordCount = 0 Then Exit Sub
    
    TBLPlanoDePagamento.MoveFirst
    
    cmbPlanoDePagamento.Clear
    
    If Not lCaixa Then
        cmbPlanoDePagamento.AddItem "(Nenhum)"
        Tamanho = 1
    Else
        Tamanho = 0
    End If
    
    Do While Not TBLPlanoDePagamento.EOF
        If Not lCaixa Or TBLPlanoDePagamento("CAIXA") Then
            If (lCompra And TBLPlanoDePagamento("COMPRA")) Or (Not lCompra And TBLPlanoDePagamento("VENDA")) Then
                cmbPlanoDePagamento.AddItem TBLPlanoDePagamento("DESCRIÇÃO")
                Tamanho = Tamanho + 1
                ASize Tamanho, CódigoDoPlanoDePagamento()
                CódigoDoPlanoDePagamento(UBound(CódigoDoPlanoDePagamento)) = TBLPlanoDePagamento("CÓDIGO")
            End If
        End If
        TBLPlanoDePagamento.MoveNext
    Loop
    
    cmbPlanoDePagamento.ListIndex = 0
    
    lPula = False
End Sub
Private Sub GetRecords()
    On Error GoTo Erro
    
    mCódigoAtual = TBLPlanoDePagamento("CÓDIGO")
    mCarência% = TBLPlanoDePagamento("CARÊNCIA")
    mIntervaloDePagamentos% = TBLPlanoDePagamento("INTERVALO DE PAGAMENTOS")
    mQuantidadeDeVencimentos = TBLPlanoDePagamento("QUANTIDADE DE VENCIMENTOS")
    mCustoFinaceiro = TBLPlanoDePagamento("CUSTO FINANCEIRO") / 100
    mRepasse = TBLPlanoDePagamento("REPASSE") / 100
    mAutoInclusão = TBLPlanoDePagamento("AUTO-INCLUSÃO")
    mPermiteAlterarAutoInclusão = TBLPlanoDePagamento("PERMITE ALTERAR AUTO-INCLUSÃO")
    
    Exit Sub
    
Erro:
    GeraMensagemDeErro "Forma de Pagamento - GetRecords "
    Resume Next
End Sub
Private Sub InicializaValores()
    lPula = True
    txtValorAVista = mValorAVista
    lPula = False
    txtValorAVista_LostFocus
    lPula = True
    txtValorAPrazo = FormatStringMask("@V ##.###.##0,00", "00.000.000,00")
    lPula = False
End Sub
Private Sub PosRecords(ByVal Valor As Integer)
    If TBLPlanoDePagamento.RecordCount = 0 Then
        Exit Sub
    End If
    TBLPlanoDePagamento.Seek "=", Valor
    
    If TBLPlanoDePagamento.NoMatch Then
        MsgBox "O código " & Valor & " não foi encontrado " & Chr(13) & "na tabela Plano de Pagamento !", vbCritical, "Erro: Código não encontrado"
        Unload Me
    End If
End Sub
Private Function PosTipoPagamento(ByVal Código As Integer) As Integer
    On Error Resume Next
    
    Dim Cont As Integer
    
    PosTipoPagamento = 0
    
    For Cont = 1 To UBound(CódigoDoPlanoDePagamento)
        If CódigoDoPlanoDePagamento(Cont) = Código Then
            PosTipoPagamento = Cont - 1
            Exit For
        End If
    Next
End Function
Private Sub RefazPagamento()
    On Error GoTo Erro
    
    Dim Cont As Byte
    Dim Valor As Currency
    Dim Data As Date
    Dim szData As String
    Dim Dia As Byte, Mes As Byte, Ano As Integer
    
    If lPula Then Exit Sub
    
    mTotalRows = 0
    ReDim dbgrdPagamentoArray(MAXCOLS - 1, 0)
        
    If mAutoInclusão Then
        mTotalRows = mQuantidadeDeVencimentos
        ReDim dbgrdPagamentoArray(MAXCOLS - 1, mTotalRows - 1)
        
        Valor = ValStr(txtValorAVista)
        Valor = Valor * (1 + mRepasse)
        Valor = Valor / mQuantidadeDeVencimentos
        
        If mRepasse > 0 Then
            txtValorAPrazo = FormatStringMask("@V ##.###.##0,00", StrVal(Valor * mQuantidadeDeVencimentos))
        End If
        
        Data = Date + mCarência
        szData = FormatStringMask(CheckDataMask, Data)
        
        For Cont = 0 To mTotalRows - 1
            dbgrdPagamentoArray(1, Cont) = FormatStringMask(CheckDataMask, CorrigeStringData(DataMask, szData, Date))
            If mIntervaloDePagamentos > 0 Then
                Data = Data + mIntervaloDePagamentos
            Else
                Dia = Day(Data)
                Mes = Month(Data) + 1
                Ano = Year(Data)
                If Mes > 12 Then
                    Mes = 1
                    Ano = Ano + 1
                End If
                Data = CDate(Str(Dia) & "/" & Str(Mes) & "/" & Str(Ano))
                szData = FormatStringMask(CheckDataMask, Data)
            End If
            dbgrdPagamentoArray(2, Cont) = FormatStringMask("@V ###.###.##0,00", StrVal(Valor))
        Next
    End If
    
    dbgrdPagamento.ReBind
    dbgrdPagamento.Refresh
    
    frDatas.Enabled = mPermiteAlterarAutoInclusão
    
    Exit Sub
    
Erro:
    GeraMensagemDeErro "Forma de Pagamento - RefazPagamento"
End Sub
Private Sub ZeraCampos()
    mCódigoAtual = 0
    mCarência% = 0
    mQuantidadeDeVencimentos = 0
    mCustoFinaceiro = 0
    mRepasse = 0
    mAutoInclusão = False
    mPermiteAlterarAutoInclusão = False
End Sub
Private Sub cmbPlanoDePagamento_Click()
    cmbPlanoDePagamento.Locked = True
        
    If cmbPlanoDePagamento.ListIndex > 0 Or lCaixa Then
        PosRecords CódigoDoPlanoDePagamento(cmbPlanoDePagamento.ListIndex + 1)
        GetRecords
    Else
        ZeraCampos
    End If
    
    RefazPagamento
End Sub
Private Sub cmbPlanoDePagamento_DropDown()
    cmbPlanoDePagamento.Locked = False
End Sub
Private Sub cmbPlanoDePagamento_KeyUp(KeyCode As Integer, Shift As Integer)
    Dim TotalNoCombo%, PosiçãoAtual%
    
    TotalNoCombo = cmbPlanoDePagamento.ListCount - 1
    
    PosiçãoAtual = cmbPlanoDePagamento.ListIndex
    
    If KeyCode = 38 Then
        If PosiçãoAtual > 0 Then
            PosiçãoAtual = PosiçãoAtual - 1
            cmbPlanoDePagamento.ListIndex = PosiçãoAtual
        End If
    ElseIf KeyCode = 40 Then
        If PosiçãoAtual < TotalNoCombo Then
            PosiçãoAtual = PosiçãoAtual + 1
            cmbPlanoDePagamento.ListIndex = PosiçãoAtual
        End If
    End If
End Sub
Private Sub cmdCancelar_Click()
    Unload Me
End Sub
Private Sub cmdGravar_Click()
    Dim Cont As Byte
    Dim Valor As Currency
    Dim Diferença As Currency
        
    If mTotalRows > 0 Then
        Valor = 0
        
        For Cont = 0 To mTotalRows - 1
            Valor = Valor + ValStr(dbgrdPagamentoArray(2, Cont))
        Next
        
        
        Diferença = ValStr(txtValorAPrazo) - Valor
        If Diferença < 0 Then
            Diferença = Diferença * (-1)
        End If
        
        If mRepasse > 0 Then
            If Diferença > 1 Then
                MsgBox "Valores incorretos!" & vbCr & "Total à prazo:  " & txtValorAPrazo & vbCr & "Total descrito: " & FormatStringMask("@V ##.###.##0,00", StrVal(Valor)), vbCritical, "Valores Incorretos"
                Exit Sub
            End If
        Else
            If ValStr(txtValorAVista) - Valor > 1 Then
                MsgBox "Valores incorretos!" & vbCr & "Total à vista:  " & txtValorAVista & vbCr & "Total descrito: " & FormatStringMask("@V ##.###.##0,00", StrVal(Valor)), vbCritical, "Valores Incorretos"
                Exit Sub
            End If
        End If
    End If
    
    mTotalPagamentos = mTotalRows
    mTipoDePagamento = CódigoDoPlanoDePagamento(cmbPlanoDePagamento.ListIndex + 1)
    mValorAPrazo = txtValorAPrazo
    If mTotalPagamentos > 0 Then
        If Not ptrForm.GravaPagamento(dbgrdPagamentoArray) Then
            Exit Sub
        End If
    Else
        If Not ptrForm.ExcluirPagamento Then
            Exit Sub
        End If
    End If
    Unload Me
End Sub
Private Sub dbgrdPagamento_AfterColEdit(ByVal ColIndex As Integer)
    If ColIndex = 0 Then 'Documento
    ElseIf ColIndex = 1 Then 'Data
        If lPula Then
            Exit Sub
        End If
        lPula = True
        CorrigeData DataMask, dbgrdPagamento, Date
        FormatMask CheckDataMask, dbgrdPagamento
        lPula = False
    ElseIf ColIndex = 2 Then 'Valor
        If lPula Then
            Exit Sub
        End If
        lPula = True
        FormatMask "@V ###.###.##0,00", dbgrdPagamento
        lPula = False
    End If
End Sub
Private Sub dbgrdPagamento_AfterUpdate()
'    Dim Cont As Byte
'    Dim Valor As Currency
    
    dbgrdPagamento.Refresh
    
'    If mRepasse = 0 Then Exit Sub
'
'    Valor = 0
'    For Cont = 0 To mTotalRows - 1
'        Valor = Valor + ValStr(dbgrdPagamentoArray(2, Cont))
'    Next
'
'    txtValorAPrazo = FormatStringMask("@V ##.###.##0,00", StrVal(Valor))
End Sub
Private Sub dbgrdPagamento_BeforeInsert(Cancel As Integer)
    If mQuantidadeDeVencimentos > 0 And mTotalRows = mQuantidadeDeVencimentos Then
        Cancel = 1
    End If
End Sub
Private Sub dbgrdPagamento_Change()
    If Not lInserir Then
        lAlterar = True
        StatusBarAviso = "Alteração da Forma de Pagamento"
        BarraDeStatus StatusBarAviso
    End If
    If dbgrdPagamento.Col = 0 Then
        FormatMask "@S30", dbgrdPagamento
    ElseIf dbgrdPagamento.Col = 1 Then
        FormatMask CheckDataMask, dbgrdPagamento
    ElseIf dbgrdPagamento.Col = 2 Then
        FormatMask "@K 999.999.999,99", dbgrdPagamento
    End If
End Sub
Private Sub dbgrdPagamento_ColResize(ByVal ColIndex As Integer, Cancel As Integer)
    Cancel = 1
End Sub
Private Sub dbgrdPagamento_RowResize(Cancel As Integer)
    Cancel = 1
End Sub
Private Sub dbgrdPagamento_UnboundAddData(ByVal RowBuf As MSDBGrid.RowBuffer, NewRowBookmark As Variant)
    Dim Col%
        
    mTotalRows = mTotalRows + 1
    ReDim Preserve dbgrdPagamentoArray(MAXCOLS - 1, mTotalRows - 1)
    
    'Sets the bookmark to the last row.
    NewRowBookmark = mTotalRows - 1
    
    ' The following loop adds a new record to the database.
    For Col = 0 To UBound(dbgrdPagamentoArray, 1)
        If Not IsNull(RowBuf.Value(0, Col)) Then
            dbgrdPagamentoArray(Col, mTotalRows - 1) = RowBuf.Value(0, Col)
        Else
            ' If no value set for column, then use the
            ' DefaultValue
            dbgrdPagamentoArray(Col, mTotalRows - 1) = dbgrdPagamento.Columns(Col).DefaultValue
        End If
    Next
End Sub
Private Sub dbgrdPagamento_UnboundDeleteRow(Bookmark As Variant)
    Dim iCol As Integer, iRow As Integer
    
    ' Move all rows above the deleted row down in the
    ' array.
    
    For iRow = Bookmark + 1 To mTotalRows - 1
        For iCol = 0 To MAXCOLS - 1
            dbgrdPagamentoArray(iCol, iRow - 1) = dbgrdPagamentoArray(iCol, iRow)
        Next iCol
    Next iRow
    
    mTotalRows = mTotalRows - 1
End Sub
Private Sub dbgrdPagamento_UnboundReadData(ByVal RowBuf As MSDBGrid.RowBuffer, StartLocation As Variant, ByVal ReadPriorRows As Boolean)
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
        For iCol = 0 To UBound(dbgrdPagamentoArray, 1)
            RowBuf.Value(iRow, iCol) = dbgrdPagamentoArray(iCol, CurRow&)
        Next iCol
        ' Set bookmark using CurRow& which is also our
        ' array index
        RowBuf.Bookmark(iRow) = CStr(CurRow)
        CurRow = CurRow + iIncr
        iRowsFetched = iRowsFetched + 1
    Next iRow
    RowBuf.RowCount = iRowsFetched
End Sub
Private Sub dbgrdPagamento_UnboundWriteData(ByVal RowBuf As MSDBGrid.RowBuffer, WriteLocation As Variant)
    Dim iCol As Integer
    ' Data is being updated
    'MsgBox WriteLocation
    ' Update each column in the data set array
    For iCol = 0 To MAXCOLS - 1
        If Not IsNull(RowBuf.Value(0, iCol)) Then
            dbgrdPagamentoArray(iCol, WriteLocation) = RowBuf.Value(0, iCol)
        End If
    Next iCol
End Sub
Private Sub Form_Activate()
    Dim Valor As Currency
    
    FillGridPg
    dbgrdPagamento.Refresh
    
    If mRepasse > 0 Then
        Valor = ValStr(txtValorAVista) * (1 + mRepasse)
    End If
    
    txtValorAPrazo = FormatStringMask("@V ##.###.##0,00", StrVal(Valor))
End Sub
Private Sub Form_Load()
    Dim Cont As Integer
    
    If lNotCancel Then
        cmdCancelar.Enabled = False
    End If
    
    dbgrdPagamento.Columns.Add 0

    dbgrdPagamento.Columns(0).Alignment = dbgLeft
    dbgrdPagamento.Columns(0).Caption = "Documento"
    dbgrdPagamento.Columns(0).DefaultValue = ""
    dbgrdPagamento.Columns(0).Visible = True
    dbgrdPagamento.Columns(0).Width = 2745
    
    dbgrdPagamento.Columns(1).Alignment = dbgCenter
    dbgrdPagamento.Columns(1).Caption = "Data"
    dbgrdPagamento.Columns(1).DefaultValue = "00/00/00"
    dbgrdPagamento.Columns(1).Visible = True
    dbgrdPagamento.Columns(1).Width = 1095
    
    dbgrdPagamento.Columns(2).Caption = "Valor"
    dbgrdPagamento.Columns(2).Alignment = dbgRight
    dbgrdPagamento.Columns(2).DefaultValue = "0,00"
    dbgrdPagamento.Columns(2).Visible = True
    dbgrdPagamento.Columns(2).Width = 1845
    
    dbgrdPagamento.LeftCol = 0
    
    dbgrdPagamento.AllowAddNew = lEdit
    dbgrdPagamento.AllowArrows = lEdit
    dbgrdPagamento.AllowDelete = lEdit
    dbgrdPagamento.AllowRowSizing = lEdit
    dbgrdPagamento.AllowUpdate = lEdit
    cmdGravar.Enabled = lEdit
    cmbPlanoDePagamento.Enabled = lEdit
    txtValorAVista.Enabled = lEdit
    
    
    dbgrdPagamento.ReBind
    
    lAtualizar = True
        
    FillPlanoDePagamento
    
    InicializaValores
    
    lPula = True
    cmbPlanoDePagamento.ListIndex = PosTipoPagamento(mTipoDePagamento)
    lPula = False
    
    frDatas.Enabled = mPermiteAlterarAutoInclusão
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Set TBLPlanoDePagamento = Nothing
End Sub
Private Sub txtValorAVista_Change()
    If Not lPula Then
        FormatMask "@K 99.999.999,99", txtValorAVista
    End If
End Sub
Private Sub txtValorAVista_KeyPress(KeyAscii As Integer)
    If Not lInserir Then
        lAlterar = True
        StatusBarAviso = "Alteração da Forma de Pagamento"
        BarraDeStatus StatusBarAviso
    End If
End Sub
Private Sub txtValorAVista_LostFocus()
    Dim Valor As Currency
    
    If Not lPula Then
        lPula = True
        FormatMask "@V ##.###.##0,00", txtValorAVista
        If ValStr(txtValorAVista) < ValStr(mValorAVista) Then
            Valor = ValStr(mValorAVista)
            Valor = Valor - ValStr(txtValorAVista)
            Valor = Valor * (1 + (mRepasse / 100))
            txtValorAPrazo = FormatStringMask("@V ##.###.##0,00", StrVal(Valor))
        End If
        lPula = False
    End If
End Sub
