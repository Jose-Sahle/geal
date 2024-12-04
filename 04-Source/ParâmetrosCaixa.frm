VERSION 5.00
Begin VB.Form frmParâmetrosCaixa 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Parâmetros de Caixa"
   ClientHeight    =   1155
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3240
   Icon            =   "ParâmetrosCaixa.frx":0000
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
      Begin VB.TextBox txtNúmero 
         Height          =   315
         Left            =   1230
         TabIndex        =   0
         Top             =   390
         Width           =   375
      End
      Begin VB.Label lblNúmero 
         Caption         =   "Número"
         Height          =   255
         Left            =   210
         TabIndex        =   4
         Top             =   450
         Width           =   585
      End
   End
End
Attribute VB_Name = "frmParâmetrosCaixa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim TBLConfiguraçãoCaixa As Table
Dim ConfiguraçãoCaixaAberto As Boolean

Dim lFechar As Boolean
Dim lNewRegister As Boolean

Dim mNúmero As Byte
Dim lPula As Boolean
Private Function ApagaCaixa() As Boolean
    On Error GoTo Erro
    
    If Not lNewRegister Then
        If PosRecords(mNúmero) Then
            TBLConfiguraçãoCaixa.Delete
        End If
    End If
    
    ApagaCaixa = True
    
    Exit Function
Erro:
    GeraMensagemDeErro "ParâmetrosCaixa - ApagaCaixa"
    ApagaCaixa = False
End Function
Private Function GravaCaixa() As Boolean
    On Error GoTo Erro
        
    WS.BeginTrans
    If lNewRegister Then
        TBLConfiguraçãoCaixa.AddNew
    Else
        TBLConfiguraçãoCaixa.Edit
    End If
    
    TBLConfiguraçãoCaixa("NÚMERO") = txtNúmero
    
    TBLConfiguraçãoCaixa.Update
    
    WS.CommitTrans
    
    GravaCaixa = True
    
    Exit Function
    
Erro:
    GeraMensagemDeErro "ParâmetrosCaixa - GravaCaixa", True
    GravaCaixa = False
End Function
Private Function PosRecords(ByVal NúmeroCaixa As Byte) As Boolean
    TBLConfiguraçãoCaixa.Seek "=", NúmeroCaixa
    If TBLConfiguraçãoCaixa.NoMatch Then
        PosRecords = False
    Else
        PosRecords = True
    End If
End Function
Private Sub cmdCancelar_Click()
    Unload Me
End Sub
Private Sub cmdExcluir_Click()
    If PosRecords(Val(Trim(txtNúmero))) Then
        TBLConfiguraçãoCaixa.Delete
        DeleteRegistryString "Caixa", "Posto"
        cmdOk.Caption = "&Fechar"
        cmdExcluir.Enabled = False
        cmdCancelar.Enabled = False
        txtNúmero.Enabled = False
        txtNúmero.Text = Empty
    End If
End Sub
Private Sub cmdOK_Click()
    On Error GoTo Erro
    
    If cmdOk.Caption = "&Fechar" Then
        Unload Me
        Exit Sub
    End If
    
    If Trim(txtNúmero) <> Empty Then
        If GravaCaixa Then
            SetRegistryString "Caixa", "Posto", "Número", Trim(txtNúmero)
        End If
    Else
        If ApagaCaixa Then
            DeleteRegistryString "Caixa", "Posto"
        End If
    End If
    
    cmdOk.Caption = "&Fechar"
    cmdExcluir.Enabled = False
    cmdCancelar.Enabled = False
    txtNúmero.Enabled = False
    
    Exit Sub
    
Erro:
    GeraMensagemDeErro "ParâmetrosCaixa - cmdOK"
End Sub
Private Sub Form_Activate()
    If lFechar Then
        Unload Me
    End If
    
    cmdOk.Caption = "&Ok"
End Sub
Private Sub Form_Load()
    lFechar = True
    
    ConfiguraçãoCaixaAberto = AbreTabela(Dicionário, "SISTEMA", "CAIXA", DBSistema, TBLConfiguraçãoCaixa, TBLTabela, dbOpenTable)
    
    If ConfiguraçãoCaixaAberto Then
        TBLConfiguraçãoCaixa.Index = "CAIXA1"
    Else
        MsgBox "Não consegui abrir a tabela 'Configuração de Caixa' !", vbCritical, "Erro"
        Exit Sub
    End If
    
    lPula = True
    txtNúmero = GetRegistryString("Caixa", "Posto", "Número", "")
    mNúmero = Val(txtNúmero)
    lPula = False
    
    If mNúmero <> 0 Then
        TBLConfiguraçãoCaixa.Seek "=", mNúmero
        If TBLConfiguraçãoCaixa.NoMatch Then
            MsgBox "Existe uma inconsistência no Posto de Caixa " & txtNúmero, vbCritical, "Inconsistência"
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
    If ConfiguraçãoCaixaAberto Then
        TBLConfiguraçãoCaixa.Close
    End If
End Sub
Private Sub txtNúmero_Change()
    FormatMask "99", txtNúmero
End Sub
