VERSION 5.00
Begin VB.Form frmSangriaEntrada 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   3420
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5235
   Icon            =   "SangriaEntrada.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3420
   ScaleWidth      =   5235
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   345
      Left            =   3990
      TabIndex        =   3
      Top             =   3060
      Width           =   1245
   End
   Begin VB.Frame frMotivo 
      Caption         =   "Motivo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1605
      Left            =   0
      TabIndex        =   8
      Top             =   1410
      Width           =   5235
      Begin VB.TextBox txtMotivo 
         Height          =   1215
         Left            =   90
         MultiLine       =   -1  'True
         TabIndex        =   1
         Top             =   300
         Width           =   5055
      End
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   345
      Left            =   2700
      TabIndex        =   2
      Top             =   3060
      Width           =   1245
   End
   Begin VB.Frame frAberturaDoCaixa 
      Height          =   705
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   5265
      Begin VB.Label lblDataAbertura 
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
         TabIndex        =   7
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
         TabIndex        =   6
         Top             =   240
         Width           =   585
      End
   End
   Begin VB.Frame frValor 
      Caption         =   "Valor"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Left            =   0
      TabIndex        =   4
      Top             =   720
      Width           =   5235
      Begin VB.TextBox txtValor 
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1590
         TabIndex        =   0
         Text            =   "##.###.##0,00"
         Top             =   240
         Width           =   1875
      End
   End
End
Attribute VB_Name = "frmSangriaEntrada"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim lFechar      As Boolean
Dim lAbriPorta   As Boolean
Dim lPula        As Boolean

Dim Usuário      As String

Dim lAllowSangriaEntrada As Boolean

Dim TBLCaixaSangriaEntrada          As Table
Dim CaixaSangriaEntradaAberto       As Boolean
Dim IndiceCaixaSangriaEntradaAtivo  As String

Dim TBLValoresDoSistema             As Table
Dim ValoresDoSistemaAberto          As Boolean
Dim IndiceValoresDoSistemaAtivo     As String

Public Tipo             As String
Public Título           As String
Public NúmeroCaixa      As Byte
Public CódigoDaAbertura As Long
Private Function SetRecords() As Boolean
    On Error GoTo Erro
    
    WS.BeginTrans
    
    TBLCaixaSangriaEntrada.AddNew
    TBLCaixaSangriaEntrada("CÓDIGO DA ABERTURA") = CódigoDaAbertura
    TBLCaixaSangriaEntrada("USERNAME") = Usuário
    TBLCaixaSangriaEntrada("HORA") = Time
    TBLCaixaSangriaEntrada("VALOR") = ValStr(txtValor)
    TBLCaixaSangriaEntrada("TIPO") = Tipo
    TBLCaixaSangriaEntrada("MOTIVO") = txtMotivo
    TBLCaixaSangriaEntrada.Update
    
    On Error Resume Next
    TBLValoresDoSistema.MoveFirst
    On Error GoTo Erro
    
    If TBLValoresDoSistema.BOF Or TBLValoresDoSistema.EOF Then
        TBLValoresDoSistema.AddNew
    Else
        TBLValoresDoSistema.Edit
    End If
    
    If Tipo = "E" Then
        TBLValoresDoSistema("CAIXA") = TBLValoresDoSistema("CAIXA") - ValStr(txtValor)
    Else
        TBLValoresDoSistema("CAIXA") = TBLValoresDoSistema("CAIXA") + ValStr(txtValor)
    End If
    TBLValoresDoSistema.Update
    
    WS.CommitTrans
    
    SetRecords = True
    
    Exit Function
    
Erro:
    GeraMensagemDeErro "Sangria Entrada - SetRecords", True
    SetRecords = False
End Function
Private Sub cmdCancelar_Click()
    Unload Me
End Sub
Private Sub cmdOK_Click()
    Dim TBLValorTotal   As Recordset
    Dim TBLSangriaTotal As Recordset
    Dim TBLEntradaTotal As Recordset
    
    Dim Valor           As Currency
    Dim ValorTotal      As Currency
    Dim SangriaTotal    As Currency
    Dim EntradaTotal    As Currency
    Dim SQL             As String

    If Tipo = "E" Then
        TBLValoresDoSistema.MoveFirst
        Valor = TBLValoresDoSistema("CAIXA")
    Else
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
        
        Valor = ValorTotal + EntradaTotal - SangriaTotal
    End If

    If Valor < ValStr(txtValor) Then
        MsgBox "Não há valores suficientes para esta operação!" _
                                                       & vbCr & "Venda Total  : " & FormatStringMask("@V ##.###.##0,00", StrVal(ValorTotal)) _
                                                       & vbCr & "Entrada Total: " & FormatStringMask("@V ##.###.##0,00", StrVal(EntradaTotal)) _
                                                       & vbCr & "Sangria Total: " & FormatStringMask("@V ##.###.##0,00", StrVal(SangriaTotal)) _
                                                       & vbCr & "Diferença    : " & FormatStringMask("@V ##.###.##0,00", StrVal((ValorTotal + EntradaTotal - SangriaTotal) - ValStr(txtValor))), vbInformation, "Aviso"
        Exit Sub
    End If
    
    If SetRecords Then
        Unload Me
    End If
End Sub
Private Sub Form_Activate()
    If lFechar Then
        Unload Me
    End If
End Sub
Private Sub Form_Load()
    lFechar = True
    
    Caption = Título
    
    frmValidaUsuário.Show 1
    
    Usuário = frmValidaUsuário.Usuário
    
    Set frmValidaUsuário = Nothing
    
    If Usuário = Empty Then
        Exit Sub
    End If
    
    If Tipo = "E" Then
        lAllowSangriaEntrada = Allow("CAIXA", "E", Usuário)
    ElseIf Tipo = "S" Then
        lAllowSangriaEntrada = Allow("CAIXA", "S", Usuário)
    End If
    
    If Not lAllowSangriaEntrada Then
        MsgBox "Acesso negado!", vbInformation, "Aviso"
        Exit Sub
    End If
    
    CaixaSangriaEntradaAberto = AbreTabela(Dicionário, "FINANCEIRO", "CAIXA - SANGRIA - ENTRADA", DBFinanceiro, TBLCaixaSangriaEntrada, TBLTabela, dbOpenTable)
    
    If CaixaSangriaEntradaAberto Then
        IndiceCaixaSangriaEntradaAtivo = "CAIXASANGRIAENTRADA1"
        TBLCaixaSangriaEntrada.Index = IndiceCaixaSangriaEntradaAtivo
    Else
        MsgBox "Não consegui abrir a tabela 'CAIXA' !", vbCritical, "Erro"
        Exit Sub
    End If
    
    ValoresDoSistemaAberto = AbreTabela(Dicionário, "FINANCEIRO", "VALORES DO SISTEMA", DBFinanceiro, TBLValoresDoSistema, TBLTabela, dbOpenTable)
    
    If ValoresDoSistemaAberto Then
    Else
        MsgBox "Não consegui abrir a tabela 'VALORES DO SISTEMA' !", vbCritical, "Erro"
        Exit Sub
    End If
    
    lblDataAbertura.Caption = Date
    lPula = True
    txtValor = "0,00"
    txtValor_LostFocus
    lPula = False
    
    lFechar = False
End Sub
Private Sub Form_Unload(Cancel As Integer)
    If CaixaSangriaEntradaAberto Then
        TBLCaixaSangriaEntrada.Close
    End If
    
    Set frmSangriaEntrada = Nothing
End Sub
Private Sub txtValor_Change()
    If Not lPula Then
        FormatMask "@K 99.999.999,99", txtValor
    End If
End Sub
Private Sub txtValor_LostFocus()
    lPula = True
    FormatMask "@V ##.###.##0,00", txtValor
    lPula = False
End Sub
