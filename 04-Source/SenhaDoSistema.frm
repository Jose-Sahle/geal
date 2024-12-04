VERSION 5.00
Begin VB.Form frmSenhaDoSistema 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Senha do Sistema"
   ClientHeight    =   1695
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4095
   Icon            =   "SenhaDoSistema.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1695
   ScaleWidth      =   4095
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtSenhaAtual 
      Height          =   345
      IMEMode         =   3  'DISABLE
      Left            =   1410
      MaxLength       =   10
      PasswordChar    =   "*"
      TabIndex        =   0
      Top             =   90
      Width           =   1155
   End
   Begin VB.TextBox txtNovaSenha 
      Height          =   345
      IMEMode         =   3  'DISABLE
      Left            =   1410
      MaxLength       =   10
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   630
      Width           =   1155
   End
   Begin VB.TextBox txtConfirmação 
      Height          =   345
      IMEMode         =   3  'DISABLE
      Left            =   1410
      MaxLength       =   10
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   1170
      Width           =   1155
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   2880
      TabIndex        =   3
      Top             =   90
      Width           =   1155
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   2880
      TabIndex        =   4
      Top             =   510
      Width           =   1155
   End
   Begin VB.Label lblSenhaAtual 
      Caption         =   "Senha Atual"
      Height          =   195
      Left            =   150
      TabIndex        =   7
      Top             =   180
      Width           =   975
   End
   Begin VB.Label lblNovaSenha 
      Caption         =   "Nova Senha"
      Height          =   195
      Left            =   150
      TabIndex        =   6
      Top             =   720
      Width           =   945
   End
   Begin VB.Label lblConfirmação 
      Caption         =   "Confirmação"
      Height          =   195
      Left            =   120
      TabIndex        =   5
      Top             =   1260
      Width           =   975
   End
End
Attribute VB_Name = "frmSenhaDoSistema"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim lFechar As Boolean
Dim lFechado As Boolean
Private Sub cmdCancelar_Click()
    Unload Me
End Sub
Private Sub cmdOK_Click()
    On Error GoTo Erro
    
    Dim Bookmark As Variant
        
    If txtNovaSenha <> txtConfirmação Then
        MsgBox "A nova senha e a confirmação da senha não coincidem. Digite-as novamente !", vbInformation, "Aviso"
        txtNovaSenha = Empty
        txtConfirmação = Empty
        Exit Sub
    End If
    
    WS.BeginTrans
    
    DBCadastro.NewPassword ValidaSenha(txtSenhaAtual), ValidaSenha(txtNovaSenha)
    DBUsuário.NewPassword ValidaSenha(txtSenhaAtual), ValidaSenha(txtNovaSenha)
    DBFinanceiro.NewPassword ValidaSenha(txtSenhaAtual), ValidaSenha(txtNovaSenha)
    DBSistema.NewPassword ValidaSenha(txtSenhaAtual), ValidaSenha(txtNovaSenha)
    DBUtilitário.NewPassword ValidaSenha(txtSenhaAtual), ValidaSenha(txtNovaSenha)
    
    TBLArquivo.MoveFirst
    
    Do While Not TBLArquivo.EOF
        Bookmark = TBLArquivo.Bookmark
        
        TBLArquivo.Edit
        TBLArquivo("Senha") = ValidaSenha(txtNovaSenha)
        TBLArquivo.Update
        
        TBLArquivo.Bookmark = Bookmark
        
        TBLArquivo.MoveNext
    Loop
    
    WS.CommitTrans
    
    MsgBox "Senha alterada com sucesso!", , "Aviso"
    
    Unload Me
    
    Exit Sub
Erro:
    GeraMensagemDeErro "Senha do Sistema - Ok", True
End Sub
Private Sub Form_Activate()
    If lFechar Then
        Unload Me
        Exit Sub
    End If
    
    txtSenhaAtual.SetFocus
End Sub
Private Sub Form_Load()
    lFechar = True
    lFechado = False
    
    If Forms.Count > 2 Then
        MsgBox "Feche todas as janelas do sistema antes de executar esta opção", vbInformation, "Aviso"
        Exit Sub
    End If
    
    FechaBaseDeDados
    lFechado = True
    If Not AbreBaseDeDados(False, True) Then
        MsgBox "Não foi possível abrir a base de dados no modo exclusivo!" & vbCr & "Por isso a Senha do Sistema não pode ser acessada.", vbInformation, "Aviso"
        Exit Sub
    End If
    
    lFechar = False
End Sub
Private Sub Form_Unload(Cancel As Integer)
    If lFechado Then
        FechaBaseDeDados
        If Not AbreBaseDeDados(False, False) Then
            MsgBox "Não foi possível reabrir a base de dados!" & vbCr & "O Sistema será encerrado!", vbCritical, "Aviso"
            Unload mdiGeal
        End If
    End If
End Sub
