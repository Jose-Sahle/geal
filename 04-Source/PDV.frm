VERSION 5.00
Begin VB.Form frmPDV 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "PDV"
   ClientHeight    =   4410
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7140
   Icon            =   "PDV.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4410
   ScaleWidth      =   7140
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame frReimpress�o 
      Caption         =   "Reimpress�o do Cupom"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   825
      Left            =   0
      TabIndex        =   13
      Top             =   2400
      Width           =   7125
      Begin VB.CommandButton cmdReimpress�o 
         Caption         =   "Reimpress�o de Cupom"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   4500
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   240
         Width           =   2295
      End
      Begin VB.TextBox txtOr�amento 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1290
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   330
         Width           =   945
      End
      Begin VB.Label lbl 
         Caption         =   "Or�amento"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   120
         TabIndex        =   14
         Top             =   360
         Width           =   1065
      End
   End
   Begin VB.CommandButton cmdFechar 
      Caption         =   "&Fechar"
      Default         =   -1  'True
      Height          =   345
      Left            =   5850
      TabIndex        =   11
      Top             =   4050
      Width           =   1245
   End
   Begin VB.Frame frStatus 
      Caption         =   "Status"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   765
      Left            =   0
      TabIndex        =   10
      Top             =   3240
      Width           =   7125
      Begin VB.TextBox txtStatus 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         ForeColor       =   &H0000FFFF&
         Height          =   315
         Left            =   120
         TabIndex        =   12
         Top             =   270
         Width           =   6915
      End
   End
   Begin VB.Frame frCancelarCupom 
      Height          =   795
      Left            =   0
      TabIndex        =   8
      Top             =   1590
      Width           =   7125
      Begin VB.CommandButton cmdCancelarCupom 
         Caption         =   "Cancelar Cupom Anterior"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   2400
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   240
         Width           =   2295
      End
   End
   Begin VB.Frame frRedu��oZ 
      Caption         =   "Redu��o Z"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   3570
      TabIndex        =   1
      Top             =   0
      Width           =   3555
      Begin VB.CommandButton cmdRedu��oZ 
         Caption         =   "Redu��o Z"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   690
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   1050
         Width           =   2205
      End
      Begin VB.TextBox txtData 
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   1680
         TabIndex        =   6
         TabStop         =   0   'False
         Text            =   "99/99/99"
         Top             =   630
         Width           =   1065
      End
      Begin VB.CheckBox chkRelat�rioGerencialZ 
         Caption         =   "Relat�rio Gerencial"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   720
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   330
         Width           =   1965
      End
      Begin VB.Label lblData 
         Caption         =   "Data"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   720
         TabIndex        =   5
         Top             =   720
         Width           =   525
      End
   End
   Begin VB.Frame frLeituraX 
      Caption         =   "Leitura X"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3555
      Begin VB.CommandButton cmdLeituraX 
         Caption         =   "Leitura X"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   630
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   1050
         Width           =   2205
      End
      Begin VB.CheckBox chkRelat�rioGerencialX 
         Caption         =   "Relat�rio Gerencial"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   750
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   420
         Width           =   1965
      End
   End
End
Attribute VB_Name = "frmPDV"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim lFechar As Boolean
Dim lPula As Boolean

Dim mAbriPorta As Boolean

Dim lAllowPDV As Boolean

Dim TBLVendas As Table
Dim VendasAberto As Boolean
Dim IndiceVendasAtivo$

Dim TBLVendasItens As Table
Dim VendasItensAberto As Boolean
Dim IndiceVendasItensAtivo$
Private Sub cmdCancelarCupom_Click()
    CancelarCupom
    txtStatus = VerStatusECF
End Sub
Private Sub cmdFechar_Click()
    Unload Me
End Sub
Private Sub cmdLeituraX_Click()
    If chkRelat�rioGerencialX.Value = 0 Then
        LeituraX "N"
    Else
        LeituraX "S"
    End If
    txtStatus = VerStatusECF
End Sub
Private Sub cmdRedu��oZ_Click()
    If chkRelat�rioGerencialX.Value = 0 Then
        Redu��oZ "N"
    Else
        Redu��oZ "S"
    End If
    txtStatus = VerStatusECF
End Sub
Private Sub cmdReimpress�o_Click()
    On Error GoTo ErroPDV
    
    Dim C�digo$, Quantidade$, Pre�oUnit�rio$, Pre�oTotal$, Descri��o$, Tributa��o$, Total$
    Dim Status$, ValorTotal As Currency, DescontoTotal As Currency
    Dim AuxValor$, AuxTexto$
    
    If txtOr�amento = Empty Then
        Exit Sub
    End If
    
    VendasAberto = AbreTabela(Dicion�rio, "FINANCEIRO", "VENDA", DBFinanceiro, TBLVendas, TBLTabela, dbOpenTable)
    
    If VendasAberto Then
        IndiceVendasAtivo = "VENDA1"
        TBLVendas.Index = IndiceVendasAtivo
    Else
        MsgBox "N�o consegui abrir a tabela 'Vendas' !", vbCritical, "Erro"
        GoTo ErroPDV
    End If
    
    VendasItensAberto = AbreTabela(Dicion�rio, "FINANCEIRO", "VENDA - ITENS", DBFinanceiro, TBLVendasItens, TBLTabela, dbOpenTable)
    
    If VendasItensAberto Then
        IndiceVendasItensAtivo = "VENDAITENS1"
        TBLVendasItens.Index = IndiceVendasItensAtivo
    Else
        MsgBox "N�o consegui abrir a tabela 'Itens de Venda' !", vbCritical, "Erro"
        GoTo ErroPDV
    End If
    
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
    
    ValorTotal = 0
    
    TBLVendas.Seek "=", txtOr�amento
    TBLVendasItens.Seek "=", txtOr�amento
    
    DescontoTotal = TBLVendas("DESCONTO TOTAL DA VENDA") + TBLVendas("VALOR DO BONUS")
    
    Do While Not TBLVendasItens.EOF And TBLVendasItens("OR�AMENTO") = txtOr�amento
        C�digo = LeftBlankString(SearchAdvancedProduto(TBLVendasItens("C�DIGO DO PRODUTO"), vbC�digoDoFornecedor, vbIndice2), 13)
        Quantidade = LeftZeroString(Str(TBLVendasItens("QUANTIDADE")), 4) & "000"
        Pre�oUnit�rio = "0" & StrTran(FormatStringMask("@V 000000,00", StrVal(TBLVendasItens("VALOR UNIT�RIO"))), ",")
        Pre�oTotal = "0" & StrTran(FormatStringMask("@V 000000000,00", StrVal(TBLVendasItens("VALOR UNIT�RIO") * TBLVendasItens("QUANTIDADE"))), ",")
        Descri��o = RightBlankString(SearchAdvancedProduto(TBLVendasItens("C�DIGO DO PRODUTO"), vbDescri��o, vbIndice2), 24)
        Tributa��o = "I  "
        
        RegistrarItemVendido C�digo, Quantidade, Pre�oUnit�rio, Pre�oTotal, Descri��o, Tributa��o
        Status = VerStatusECF
        
        If Mid(Status, 1, 2) = ".-" Then
            AuxTexto = Mid(Status, 3, 4)
            Status = Mid(Status, 7, Len(Status) - 7)
            MsgBox Status, vbCritical, "Erro #" & AuxTexto
            GoTo ErroPDV
        End If
        
        ValorTotal = ValorTotal + TBLVendasItens("VALOR UNIT�RIO") * TBLVendasItens("QUANTIDADE")
        
        TBLVendasItens.MoveNext
        If TBLVendasItens.EOF Then
            Exit Do
        End If
    Loop
    
    If DescontoTotal > 0 Then
        AuxValor = StrTran(FormatStringMask("@V 0000000000,00", StrVal(DescontoTotal)), ",")
        AuxTexto = FormatStringMask("@V ##%", StrVal(DescontoTotal * 100 / ValorTotal))
        AuxTexto = RightBlankString(AuxTexto, 10)
        DescontoSobreCupomFiscal AuxTexto, AuxValor
    End If
    
    frmTotal.ValorAPagar = ValorTotal - DescontoTotal
    frmTotal.Show 1
    Total = frmTotal.Total
    
    Set frmTotal = Nothing
    
    TotalizarCupomFiscal Total
    Status = VerStatusECF
    
    If Mid(Status, 1, 2) = ".-" Then
        AuxTexto = Mid(Status, 3, 4)
        Status = Mid(Status, 7, Len(Status) - 7)
        MsgBox Status, vbCritical, "Erro #" & AuxTexto
        GoTo ErroPDV
    End If
    
    FecharCupomFiscal
    Status = VerStatusECF
    
    If Mid(Status, 1, 2) = ".-" Then
        AuxTexto = Mid(Status, 3, 4)
        Status = Mid(Status, 7, Len(Status) - 7)
        MsgBox Status, vbCritical, "Erro #" & AuxTexto
        GoTo ErroPDV
    End If
        
    txtStatus = Status
    
ErroPDV:
    If VendasAberto Then
        TBLVendas.Close
    End If
    If VendasItensAberto Then
        TBLVendasItens.Close
    End If
    txtOr�amento = Empty
End Sub
Private Sub Form_Activate()
    If lFechar Then
        Unload Me
    End If
End Sub
Private Sub Form_Load()
    lFechar = False
    lPula = False
        
    Dim Usu�rio As String
    
    frmValidaUsu�rio.Show 1
    
    Usu�rio = frmValidaUsu�rio.Usu�rio
    
    Set frmValidaUsu�rio = Nothing
    
    If Usu�rio = Empty Then
        Exit Sub
    End If
    
    lAllowPDV = Allow("PDV", "A", Usu�rio)
    
    If Not lAllowPDV Then
        MsgBox "Acesso negado!", vbInformation, "Aviso"
        lFechar = True
        Exit Sub
    End If
    
    If Not AbrirPorta(mAbriPorta) Then
        lFechar = True
        Exit Sub
    End If
End Sub
Private Sub Form_Unload(Cancel As Integer)
    If mAbriPorta Then
        FecharPorta
    End If
End Sub
Private Sub txtOr�amento_Change()
    If Not lPula Then
        FormatMask "999999", txtOr�amento
    End If
End Sub
Private Sub txtOr�amento_LostFocus()
    lPula = True
    txtOr�amento = Trim(txtOr�amento)
    lPula = False
End Sub

