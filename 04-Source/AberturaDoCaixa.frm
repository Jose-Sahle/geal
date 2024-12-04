VERSION 5.00
Begin VB.Form frmAberturaDoCaixa 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Abertura do Caixa"
   ClientHeight    =   2775
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5265
   Icon            =   "AberturaDoCaixa.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2775
   ScaleWidth      =   5265
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdFechar 
      Caption         =   "&Fechar"
      Default         =   -1  'True
      Height          =   345
      Left            =   3990
      TabIndex        =   7
      Top             =   2370
      Width           =   1245
   End
   Begin VB.Frame frCancelarCupom 
      Height          =   795
      Left            =   0
      TabIndex        =   5
      Top             =   1500
      Width           =   5265
      Begin VB.CommandButton cmdAbrirCaixa 
         Caption         =   "&Abrir Caixa"
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
         Top             =   240
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
      TabIndex        =   2
      Top             =   720
      Width           =   5265
      Begin VB.TextBox txtStatus 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         ForeColor       =   &H0000FFFF&
         Height          =   315
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   3
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
         TabIndex        =   4
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
Attribute VB_Name = "frmAberturaDoCaixa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim lFechar      As Boolean
Dim lAbriPorta   As Boolean
Dim lAllowAcesso As Boolean

Dim TBLConfigura��oCaixa    As Table
Dim Configura��oCaixaAberto As Boolean

Dim TBLPar�metros    As Table
Dim Par�metrosAberto As Boolean

Dim TBLCaixa          As Table
Dim CaixaAberto       As Boolean
Dim IndiceCaixaAtivo$

Dim N�meroCaixa As Byte

Dim Usu�rio As String

Public AbrirCaixa As Boolean
Private Sub cmdAbrirCaixa_Click()
    On Error GoTo Erro
    
    Dim Bookmark As Variant
    Dim C�digo   As Integer
        
    TBLPar�metros.Edit
    If IsNull(TBLPar�metros("C�DIGO DO CAIXA")) Then
        C�digo = 1
    Else
        C�digo = TBLPar�metros("C�DIGO DO CAIXA") + 1
    End If
    TBLPar�metros("C�DIGO DO CAIXA") = C�digo
    TBLPar�metros.Update
    
    WS.BeginTrans

    TBLCaixa.AddNew
    TBLCaixa("C�DIGO") = C�digo
    TBLCaixa("POSTO") = N�meroCaixa
    TBLCaixa("USERNAME") = Usu�rio
    TBLCaixa("DATA DE ABERTURA") = Date
    TBLCaixa("HORA DE ABERTURA") = Time
    TBLCaixa("ABERTO") = True
    TBLCaixa.Update
        
    If LeituraX("S") Then
        cmdAbrirCaixa.Enabled = False
        If StatusOkECF Then
            AbrirCaixa = True
            WS.CommitTrans
        Else
            WS.Rollback
        End If
    Else
        WS.Rollback
    End If
    
    txtStatus = VerStatusECF
    
    Exit Sub
    
Erro:
    GeraMensagemDeErro "Abertura de Caixa - Abrir", True
    AbrirCaixa = False
End Sub
Private Sub cmdFechar_Click()
    Unload Me
End Sub
Private Sub Form_Activate()
    If lFechar Then
        Unload Me
    End If
End Sub
Private Sub Form_Load()

    AbrirCaixa = False
    
    lFechar = True
    
    frmValidaUsu�rio.Show 1
    
    Usu�rio = frmValidaUsu�rio.Usu�rio
    
    Set frmValidaUsu�rio = Nothing
    
    If Usu�rio = Empty Then
        Exit Sub
    End If
    
    lAllowAcesso = Allow("ABERTURA DO CAIXA", "A", Usu�rio)
    
    If Not lAllowAcesso Then
        MsgBox "Acesso negado!", vbInformation, "Aviso"
        Exit Sub
    End If
    
    Configura��oCaixaAberto = AbreTabela(Dicion�rio, "SISTEMA", "CAIXA", DBSistema, TBLConfigura��oCaixa, TBLTabela, dbOpenTable)
    
    If Configura��oCaixaAberto Then
        TBLConfigura��oCaixa.Index = "CAIXA1"
    Else
        MsgBox "N�o consegui abrir a tabela 'Configura��o de Caixa' !", vbCritical, "Erro"
        Exit Sub
    End If
    
    N�meroCaixa = Val(GetRegistryString("Caixa", "Posto", "N�mero", ""))

    If N�meroCaixa <> 0 Then
        TBLConfigura��oCaixa.Seek "=", N�meroCaixa
        If TBLConfigura��oCaixa.NoMatch Then
            MsgBox "Existe uma inconsist�ncia no Posto de Caixa " & N�meroCaixa, vbCritical, "Inconsist�ncia"
            Exit Sub
        End If
    Else
        MsgBox "Nenhum Posto de Caixa foi configurado para este computador!", vbCritical, "Inconsist�ncia"
        Exit Sub
    End If
    
    CaixaAberto = AbreTabela(Dicion�rio, "FINANCEIRO", "CAIXA", DBFinanceiro, TBLCaixa, TBLTabela, dbOpenTable)
    
    If CaixaAberto Then
        IndiceCaixaAtivo = "CAIXA3"
        TBLCaixa.Index = IndiceCaixaAtivo
    Else
        MsgBox "N�o consegui abrir a tabela 'CAIXA' !", vbCritical, "Erro"
        Exit Sub
    End If
    
    TBLCaixa.Seek "=", N�meroCaixa, True, False
    
    If Not TBLCaixa.NoMatch Then
        MsgBox "Caixa dever ser fechado," & vbCr & "antes que uma nova opera��o seja iniciada!", vbInformation, "Aviso"
        Exit Sub
    End If

    Par�metrosAberto = AbreTabela(Dicion�rio, "SISTEMA", "PAR�METROS", DBSistema, TBLPar�metros, TBLTabela, dbOpenTable)
    
    If Par�metrosAberto Then
    Else
        MsgBox "N�o consegui abrir a tabela 'Par�metros' !", vbCritical, "Erro"
        Exit Sub
    End If
    
    If Not AbrirPorta(lAbriPorta) Then
        Exit Sub
    End If
    
    lblDataAbertura.Caption = Date
    
    lFechar = False
End Sub
Private Sub Form_Unload(Cancel As Integer)
    If Configura��oCaixaAberto Then
        TBLConfigura��oCaixa.Close
    End If
    If CaixaAberto Then
        TBLCaixa.Close
    End If
    If Par�metrosAberto Then
        TBLPar�metros.Close
    End If
    
    If lAbriPorta Then
        FecharPorta
    End If
    
    Set frmAberturaDoCaixa = Nothing
End Sub
