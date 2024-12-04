VERSION 5.00
Begin VB.Form frmPar�metrosCaixa 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Par�metros de Caixa"
   ClientHeight    =   1155
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3240
   Icon            =   "Par�metrosCaixa.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1155
   ScaleWidth      =   3240
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdExcluir 
      Caption         =   "&Excluir"
      Height          =   345
      Left            =   1950
      TabIndex        =   5
      Top             =   420
      Width           =   1245
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   345
      Left            =   1950
      TabIndex        =   2
      Top             =   780
      Width           =   1245
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&Ok"
      Height          =   345
      Left            =   1950
      TabIndex        =   1
      Top             =   60
      Width           =   1245
   End
   Begin VB.Frame frPosto 
      Caption         =   "Posto"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1035
      Left            =   0
      TabIndex        =   3
      Top             =   60
      Width           =   1875
      Begin VB.TextBox txtN�mero 
         Height          =   315
         Left            =   1230
         TabIndex        =   0
         Top             =   390
         Width           =   375
      End
      Begin VB.Label lblN�mero 
         Caption         =   "N�mero"
         Height          =   255
         Left            =   210
         TabIndex        =   4
         Top             =   450
         Width           =   585
      End
   End
End
Attribute VB_Name = "frmPar�metrosCaixa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim TBLConfigura��oCaixa As Table
Dim Configura��oCaixaAberto As Boolean

Dim lFechar As Boolean
Dim lNewRegister As Boolean

Dim mN�mero As Byte
Dim lPula As Boolean
Private Function ApagaCaixa() As Boolean
    On Error GoTo Erro
    
    If Not lNewRegister Then
        If PosRecords(mN�mero) Then
            TBLConfigura��oCaixa.Delete
        End If
    End If
    
    ApagaCaixa = True
    
    Exit Function
Erro:
    GeraMensagemDeErro "Par�metrosCaixa - ApagaCaixa"
    ApagaCaixa = False
End Function
Private Function GravaCaixa() As Boolean
    On Error GoTo Erro
        
    WS.BeginTrans
    If lNewRegister Then
        TBLConfigura��oCaixa.AddNew
    Else
        TBLConfigura��oCaixa.Edit
    End If
    
    TBLConfigura��oCaixa("N�MERO") = txtN�mero
    
    TBLConfigura��oCaixa.Update
    
    WS.CommitTrans
    
    GravaCaixa = True
    
    Exit Function
    
Erro:
    GeraMensagemDeErro "Par�metrosCaixa - GravaCaixa", True
    GravaCaixa = False
End Function
Private Function PosRecords(ByVal N�meroCaixa As Byte) As Boolean
    TBLConfigura��oCaixa.Seek "=", N�meroCaixa
    If TBLConfigura��oCaixa.NoMatch Then
        PosRecords = False
    Else
        PosRecords = True
    End If
End Function
Private Sub cmdCancelar_Click()
    Unload Me
End Sub
Private Sub cmdExcluir_Click()
    If PosRecords(Val(Trim(txtN�mero))) Then
        TBLConfigura��oCaixa.Delete
        DeleteRegistryString "Caixa", "Posto"
        cmdOk.Caption = "&Fechar"
        cmdExcluir.Enabled = False
        cmdCancelar.Enabled = False
        txtN�mero.Enabled = False
        txtN�mero.Text = Empty
    End If
End Sub
Private Sub cmdOK_Click()
    On Error GoTo Erro
    
    If cmdOk.Caption = "&Fechar" Then
        Unload Me
        Exit Sub
    End If
    
    If Trim(txtN�mero) <> Empty Then
        If GravaCaixa Then
            SetRegistryString "Caixa", "Posto", "N�mero", Trim(txtN�mero)
        End If
    Else
        If ApagaCaixa Then
            DeleteRegistryString "Caixa", "Posto"
        End If
    End If
    
    cmdOk.Caption = "&Fechar"
    cmdExcluir.Enabled = False
    cmdCancelar.Enabled = False
    txtN�mero.Enabled = False
    
    Exit Sub
    
Erro:
    GeraMensagemDeErro "Par�metrosCaixa - cmdOK"
End Sub
Private Sub Form_Activate()
    If lFechar Then
        Unload Me
    End If
    
    cmdOk.Caption = "&Ok"
End Sub
Private Sub Form_Load()
    lFechar = True
    
    Configura��oCaixaAberto = AbreTabela(Dicion�rio, "SISTEMA", "CAIXA", DBSistema, TBLConfigura��oCaixa, TBLTabela, dbOpenTable)
    
    If Configura��oCaixaAberto Then
        TBLConfigura��oCaixa.Index = "CAIXA1"
    Else
        MsgBox "N�o consegui abrir a tabela 'Configura��o de Caixa' !", vbCritical, "Erro"
        Exit Sub
    End If
    
    lPula = True
    txtN�mero = GetRegistryString("Caixa", "Posto", "N�mero", "")
    mN�mero = Val(txtN�mero)
    lPula = False
    
    If mN�mero <> 0 Then
        TBLConfigura��oCaixa.Seek "=", mN�mero
        If TBLConfigura��oCaixa.NoMatch Then
            MsgBox "Existe uma inconsist�ncia no Posto de Caixa " & txtN�mero, vbCritical, "Inconsist�ncia"
            lNewRegister = True
        Else
            lNewRegister = False
        End If
    Else
        lNewRegister = True
    End If
    
    lFechar = False
End Sub
Private Sub Form_Unload(Cancel As Integer)
    If Configura��oCaixaAberto Then
        TBLConfigura��oCaixa.Close
    End If
End Sub
Private Sub txtN�mero_Change()
    FormatMask "99", txtN�mero
End Sub
