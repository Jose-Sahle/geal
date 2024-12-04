VERSION 5.00
Begin VB.Form frmMudan�aDeSenha 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Mudan�a de Senha"
   ClientHeight    =   2115
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4215
   Icon            =   "Mudan�aDeSenha.frx":0000
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
   Begin VB.TextBox txtConfirma��o 
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
   Begin VB.Label lblConfirma��o 
      Caption         =   "Confirma��o"
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
Attribute VB_Name = "frmMudan�aDeSenha"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim TBLUsu�rio As Table
Dim Usu�rioAberto As Boolean
Dim IndiceAtivoUsu�rio As String

Dim lPula As Boolean

Public Usu�rio$

Public lAtualizar As Boolean
Private Sub cmdCancelar_Click()
    Unload Me
End Sub
Private Sub cmdOk_Click()
    On Error GoTo Erro
    
    Dim Senha$
    
    TBLUsu�rio.Seek "=", Usu�rio
    If TBLUsu�rio.NoMatch Then
        MsgBox "Usu�rio n�o cadastrado!", vbInformation, "Aviso"
        Unload Me
        Exit Sub
    End If
    
    Senha = Trim(UCase(txtSenhaAtual))
    If TBLUsu�rio("SENHA") <> ValidaSenha(Senha) Then
        MsgBox "Senha inv�lida!", vbCritical, "Aviso"
        lPula = True
        txtSenhaAtual = Empty
        txtSenhaAtual.SetFocus
        lPula = False
        Exit Sub
    End If
    
    If Trim(UCase(txtNovaSenha)) <> Trim(UCase(txtConfirma��o)) Then
        MsgBox "A nova senha e a confirma��o da senha n�o coincidem. Digite-as novamente !", vbInformation, "Aviso"
        lPula = True
        txtNovaSenha = Empty
        txtConfirma��o = Empty
        lPula = False
        Exit Sub
    End If
    
    TBLUsu�rio.Edit
    TBLUsu�rio("SENHA") = ValidaSenha(Trim(UCase(txtNovaSenha)))
    TBLUsu�rio.Update
    
    MsgBox "Senha alterada com sucesso !", , "Aviso"
    
    Unload Me
    
    Exit Sub
    
Erro:
    GeraMensagemDeErro "Mudan�� de Usu�rio - " & Usu�rio
End Sub
Private Sub Form_Load()
    Usu�rioAberto = AbreTabela(Dicion�rio, "USU�RIO", "USU�RIO", DBUsu�rio, TBLUsu�rio, TBLTabela, dbOpenTable)
    If Usu�rioAberto Then
        IndiceAtivoUsu�rio = "USU�RIO1"
        TBLUsu�rio.Index = IndiceAtivoUsu�rio
    Else
        MsgBox "N�o consegui abrir a tabela 'Usu�rio' !", vbCritical, "Erro"
        Exit Sub
    End If
End Sub
Private Sub Form_Unload(Cancel As Integer)
    If Usu�rioAberto Then
        TBLUsu�rio.Close
    End If
    
    Set frmMudan�aDeSenha = Nothing
End Sub
