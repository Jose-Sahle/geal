VERSION 5.00
Begin VB.Form frmFechamentoOperaçãoDoCaixa 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Fechamento de Operação do Caixa"
   ClientHeight    =   3435
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5280
   Icon            =   "FechamentoOperaçãoDoCaixa.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3435
   ScaleWidth      =   5280
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   795
      Left            =   0
      TabIndex        =   8
      Top             =   2220
      Width           =   5265
      Begin VB.CommandButton cmdFecharOperação 
         Caption         =   "&Fechar Operação"
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
         Left            =   1440
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   240
         Width           =   2295
      End
   End
   Begin VB.CommandButton cmdFechar 
      Caption         =   "&OK"
      Height          =   345
      Left            =   3990
      TabIndex        =   7
      Top             =   3060
      Width           =   1245
   End
   Begin VB.Frame frCancelarCupom 
      Height          =   735
      Left            =   0
      TabIndex        =   5
      Top             =   690
      Width           =   5265
      Begin VB.CommandButton cmdSangria 
         Caption         =   "&Sangria"
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
         Left            =   1440
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   180
         Width           =   2295
      End
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
      TabIndex        =   3
      Top             =   1440
      Width           =   5265
      Begin VB.TextBox txtStatus 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         ForeColor       =   &H0000FFFF&
         Height          =   315
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   270
         Width           =   5025
      End
   End
   Begin VB.Frame frAberturaDoCaixa 
      Height          =   705
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5265
      Begin VB.Label lblDataFechamento 
         Caption         =   "00/00/00"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2400
         TabIndex        =   2
         Top             =   270
         Width           =   1155
      End
      Begin VB.Label lblData 
         Caption         =   "Data"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1440
         TabIndex        =   1
         Top             =   240
         Width           =   585
      End
   End
End
Attribute VB_Name = "frmFechamentoOperaçãoDoCaixa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim lFechar      As Boolean
Dim lAbriPorta   As Boolean
Dim lAllowAcesso As Boolean
Dim lPula        As Boolean

Dim TBLConfiguraçãoCaixa    As Table
Dim ConfiguraçãoCaixaAberto As Boolean

Dim TBLCaixaAbertura          As Table
Dim CaixaAberturaAberto       As Boolean
Dim IndiceCaixaAberturaAtivo$

Dim NúmeroCaixa   As Byte

Public CódigoDoCaixa As Long
Public CódigoDaAbertura As Long
Public lSuccessfull As Boolean
Private Sub cmdFechar_Click()
    Unload Me
End Sub
Private Sub cmdFecharOperação_Click()
    On Error GoTo Erro
    
    Dim Bookmark As Variant
    Dim Código   As Integer
    
    Dim TBLValorTotal   As Recordset
    Dim TBLSangriaTotal As Recordset
    Dim TBLEntradaTotal As Recordset
    
    Dim ValorTotal      As Currency
    Dim SangriaTotal    As Currency
    Dim EntradaTotal    As Currency
    Dim SQL             As String

    SQL = "SELECT SUM(B.[VALOR TOTAL DA VENDA]) As [VALOR TOTAL] From [CAIXA - MOVIMENTO] AS A LEFT JOIN [VENDA] AS B ON A.[ORÇAMENTO] = B.[CÓDIGO]Where A.[CÓDIGO DA ABERTURA] = " & CódigoDaAbertura
    Set TBLValorTotal = DBFinanceiro.OpenRecordset(SQL)
    
    SQL = "SELECT SUM(VALOR) AS [SANGRIA TOTAL] FROM [CAIXA - SANGRIA - ENTRADA] WHERE [CÓDIGO DA ABERTURA] = " & CódigoDaAbertura & " AND [TIPO] = 'S'"
    Set TBLSangriaTotal = DBFinanceiro.OpenRecordset(SQL)
    
    SQL = "SELECT SUM(VALOR)AS [ENTRADA TOTAL] FROM [CAIXA - SANGRIA - ENTRADA] WHERE [CÓDIGO DA ABERTURA] = " & CódigoDaAbertura & " AND [TIPO] = 'E'"
    Set TBLEntradaTotal = DBFinanceiro.OpenRecordset(SQL)
    
    ValorTotal = IIf(IsNull(TBLValorTotal("VALOR TOTAL")), 0, TBLValorTotal("VALOR TOTAL"))
    SangriaTotal = IIf(IsNull(TBLSangriaTotal("SANGRIA TOTAL")), 0, TBLSangriaTotal("SANGRIA TOTAL"))
    EntradaTotal = IIf(IsNull(TBLEntradaTotal("ENTRADA TOTAL")), 0, TBLEntradaTotal("ENTRADA TOTAL"))
    
    TBLValorTotal.Close
    TBLSangriaTotal.Close
    TBLEntradaTotal.Close
    
    If (ValorTotal + EntradaTotal - SangriaTotal) <> 0 Then
        MsgBox "Valores do caixa não estão corretos !" & vbCr & "Venda Total  : " & FormatStringMask("@V ##.###.##0,00", StrVal(ValorTotal)) _
                                                       & vbCr & "Entrada Total: " & FormatStringMask("@V ##.###.##0,00", StrVal(EntradaTotal)) _
                                                       & vbCr & "Sangria Total: " & FormatStringMask("@V ##.###.##0,00", StrVal(SangriaTotal)) _
                                                       & vbCr & "Diferença    : " & FormatStringMask("@V ##.###.##0,00", StrVal((ValorTotal + EntradaTotal - SangriaTotal))), vbInformation, "Aviso"
        Exit Sub
    End If
    
    On Error GoTo ErroTrans
    
    TBLCaixaAbertura.Seek "=", CódigoDaAbertura
    
    If TBLCaixaAbertura.NoMatch Then
        MsgBox "Código de Abertura " & CódigoDaAbertura & " não foi localizado!", vbCritical, "Aviso"
        Exit Sub
    End If
    
    WS.BeginTrans

    TBLCaixaAbertura.Edit
    TBLCaixaAbertura("HORA DE FECHAMENTO") = Time
    TBLCaixaAbertura("VALOR TOTAL") = ValorTotal
    TBLCaixaAbertura("SANGRIA") = SangriaTotal
    TBLCaixaAbertura("ENTRADA") = EntradaTotal
    TBLCaixaAbertura.Update
    
    If LeituraX("S") Then
        cmdFecharOperação.Enabled = False
        If StatusOkECF Then
            WS.CommitTrans
            lSuccessfull = True
        Else
            WS.Rollback
        End If
    Else
        WS.Rollback
    End If
    
    txtStatus = VerStatusECF
    
    Exit Sub
    
Erro:
    GeraMensagemDeErro "Fechamento de Caixa - Abrir", False
    lSuccessfull = False
    Exit Sub

ErroTrans:
    GeraMensagemDeErro "Fechamento de Caixa - Abrir", True
    lSuccessfull = False
End Sub
Private Sub cmdSangria_Click()
    frmSangriaEntrada.Tipo = "S"
    frmSangriaEntrada.CódigoDaAbertura = CódigoDaAbertura
    frmSangriaEntrada.NúmeroCaixa = NúmeroCaixa
    frmSangriaEntrada.Título = "Sangrida do Caixa"
    frmSangriaEntrada.Show 1
End Sub
Private Sub Form_Activate()
    If lFechar Then
        Unload Me
    End If
End Sub
Private Sub Form_Load()
    lFechar = True
    lSuccessfull = False
    
    ConfiguraçãoCaixaAberto = AbreTabela(Dicionário, "SISTEMA", "CAIXA", DBSistema, TBLConfiguraçãoCaixa, TBLTabela, dbOpenTable)
    
    If ConfiguraçãoCaixaAberto Then
        TBLConfiguraçãoCaixa.Index = "CAIXA1"
    Else
        MsgBox "Não consegui abrir a tabela 'Configuração de Caixa' !", vbCritical, "Erro"
        Exit Sub
    End If
    
    NúmeroCaixa = GetRegistryString("Caixa", "Posto", "Número", "")

    If NúmeroCaixa <> 0 Then
        TBLConfiguraçãoCaixa.Seek "=", NúmeroCaixa
        If TBLConfiguraçãoCaixa.NoMatch Then
            MsgBox "Existe uma inconsistência no Posto de Caixa " & NúmeroCaixa, vbCritical, "Inconsistência"
            Exit Sub
        End If
    Else
        MsgBox "Nenhum Posto de Caixa foi configurado para este computador!", vbCritical, "Inconsistência"
        Exit Sub
    End If
    
    CaixaAberturaAberto = AbreTabela(Dicionário, "FINANCEIRO", "CAIXA - ABERTURA", DBFinanceiro, TBLCaixaAbertura, TBLTabela, dbOpenTable)
    
    If CaixaAberturaAberto Then
        IndiceCaixaAberturaAtivo = "CAIXAABERTURA1"
        TBLCaixaAbertura.Index = IndiceCaixaAberturaAtivo
    Else
        MsgBox "Não consegui abrir a tabela 'CAIXA - ABERTURA' !", vbCritical, "Erro"
        Exit Sub
    End If
    
    If Not AbrirPorta(lAbriPorta) Then
        Exit Sub
    End If
    
    lblDataFechamento.Caption = Date
    
    lFechar = False
End Sub
Private Sub Form_Unload(Cancel As Integer)
    If ConfiguraçãoCaixaAberto Then
        TBLConfiguraçãoCaixa.Close
    End If
    If CaixaAberturaAberto Then
        TBLCaixaAbertura.Close
    End If
    
    If lAbriPorta Then
        FecharPorta
    End If
End Sub
