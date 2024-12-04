VERSION 5.00
Begin VB.Form frmMudançaDeSenha 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Mudança de Senha"
   ClientHeight    =   2115
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4215
   Icon            =   "MudançaDeSenha.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2115
   ScaleWidth      =   4215
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   2880
      TabIndex        =   4
      Top             =   690
      Width           =   1155
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   2880
      TabIndex        =   3
      Top             =   270
      Width           =   1155
   End
   Begin VB.TextBox txtConfirmação 
      Height          =   345
      IMEMode         =   3  'DISABLE
      Left            =   1410
      MaxLength       =   6
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   1350
      Width           =   1155
   End
   Begin VB.TextBox txtNovaSenha 
      Height          =   345
      IMEMode         =   3  'DISABLE
      Left            =   1410
      MaxLength       =   6
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   810
      Width           =   1155
   End
   Begin VB.TextBox txtSenhaAtual 
      Height          =   345
      IMEMode         =   3  'DISABLE
      Left            =   1410
      MaxLength       =   6
      PasswordChar    =   "*"
      TabIndex        =   0
      Top             =   270
      Width           =   1155
   End
   Begin VB.Label lblConfirmação 
      Caption         =   "Confirmação"
      Height          =   195
      Left            =   120
      TabIndex        =   7
      Top             =   1440
      Width           =   975
   End
   Begin VB.Label lblNovaSenha 
      Caption         =   "Nova Senha"
      Height          =   195
      Left            =   150
      TabIndex        =   6
      Top             =   900
      Width           =   945
   End
   Begin VB.Label lblSenhaAtual 
      Caption         =   "Senha Atual"
      Height          =   195
      Left            =   150
      TabIndex        =   5
      Top             =   360
      Width           =   975
   End
End
Attribute VB_Name = "frmMudançaDeSenha"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim TBLUsuário As Table
Dim UsuárioAberto As Boolean
Dim IndiceAtivoUsuário As String

Dim lPula As Boolean

Public Usuário$

Public lAtualizar As Boolean
Private Sub cmdCancelar_Click()
    Unload Me
End Sub
Private Sub cmdOk_Click()
    On Error GoTo Erro
    
    Dim Senha$
    
    TBLUsuário.Seek "=", Usuário
    If TBLUsuário.NoMatch Then
        MsgBox "Usuário não cadastrado!", vbInformation, "Aviso"
        Unload Me
        Exit Sub
    End If
    
    Senha = Trim(UCase(txtSenhaAtual))
    If TBLUsuário("SENHA") <> ValidaSenha(Senha) Then
        MsgBox "Senha inválida!", vbCritical, "Aviso"
        lPula = True
        txtSenhaAtual = Empty
        txtSenhaAtual.SetFocus
        lPula = False
        Exit Sub
    End If
    
    If Trim(UCase(txtNovaSenha)) <> Trim(UCase(txtConfirmação)) Then
        MsgBox "A nova senha e a confirmação da senha não coincidem. Digite-as novamente !", vbInformation, "Aviso"
        lPula = True
        txtNovaSenha = Empty
        txtConfirmação = Empty
        lPula = False
        Exit Sub
    End If
    
    TBLUsuário.Edit
    TBLUsuário("SENHA") = ValidaSenha(Trim(UCase(txtNovaSenha)))
    TBLUsuário.Update
    
    MsgBox "Senha alterada com sucesso !", , "Aviso"
    
    Unload Me
    
    Exit Sub
    
Erro:
    GeraMensagemDeErro "Mudançã de Usuário - " & Usuário
End Sub
Private Sub Form_Load()
    UsuárioAberto = AbreTabela(Dicionário, "USUÁRIO", "USUÁRIO", DBUsuário, TBLUsuário, TBLTabela, dbOpenTable)
    If UsuárioAberto Then
        IndiceAtivoUsuário = "USUÁRIO1"
        TBLUsuário.Index = IndiceAtivoUsuário
    Else
        MsgBox "Não consegui abrir a tabela 'Usuário' !", vbCritical, "Erro"
        Exit Sub
    End If
End Sub
Private Sub Form_Unload(Cancel As Integer)
    If UsuárioAberto Then
        TBLUsuário.Close
    End If
    
    Set frmMudançaDeSenha = Nothing
End Sub
