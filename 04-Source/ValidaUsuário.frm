VERSION 5.00
Begin VB.Form frmValidaUsuário 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Senha"
   ClientHeight    =   1650
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3825
   Icon            =   "ValidaUsuário.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1650
   ScaleWidth      =   3825
   Begin VB.CommandButton cmdOk 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   690
      TabIndex        =   6
      Top             =   1230
      Width           =   1245
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   1980
      TabIndex        =   5
      Top             =   1230
      Width           =   1245
   End
   Begin VB.Frame frSenha 
      Height          =   1185
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3825
      Begin VB.TextBox txtSenha 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   900
         MaxLength       =   6
         PasswordChar    =   "*"
         TabIndex        =   4
         Top             =   660
         Width           =   1275
      End
      Begin VB.TextBox txtUserName 
         Height          =   285
         Left            =   900
         TabIndex        =   2
         Text            =   "Admin"
         Top             =   240
         Width           =   1275
      End
      Begin VB.Label lblSenha 
         Caption         =   "Senha"
         Height          =   255
         Left            =   180
         TabIndex        =   3
         Top             =   690
         Width           =   585
      End
      Begin VB.Label lblUserName 
         Caption         =   "Usuário"
         Height          =   225
         Left            =   180
         TabIndex        =   1
         Top             =   270
         Width           =   795
      End
   End
End
Attribute VB_Name = "frmValidaUsuário"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim TBLUsuário As Table
Dim UsuárioAberto As Boolean
Dim IndiceAtivoUsuário As String

Dim lPula As Boolean

Public Bloquear     As Boolean
Public Fechado      As Boolean
Public Usuário      As String
Public GravaUsuário As Boolean
Public WindowTop    As Long
Public WindowHeight As Long
Private Sub cmdCancelar_Click()
    Usuário = ""
    Unload Me
End Sub
Private Sub cmdOK_Click()
    Dim Senha$
    TBLUsuário.Seek "=", UCase(Trim(txtUserName))
    If TBLUsuário.NoMatch Then
        MsgBox "Usuário não cadastrado!", vbInformation, "Aviso"
        txtUserName.SetFocus
        Exit Sub
    End If
    
    Senha = Trim(UCase(txtSenha))
    
    If TBLUsuário("SENHA") <> ValidaSenha(Senha) Then
        MsgBox "Senha inválida!", vbCritical, "Aviso"
        txtSenha.SetFocus
        Exit Sub
    End If
    
    Usuário = UCase(Trim(txtUserName))
    Unload Me
End Sub
Private Sub Form_Activate()
    If Not UsuárioAberto Then
        Usuário = ""
        Unload Me
    End If
End Sub
Private Sub Form_Load()
    Fechado = False
        
    'Se é para bloquear sistema
    If Bloquear Then
        cmdCancelar.Enabled = False
        cmdCancelar.Visible = False
        cmdOK.Left = 1260
        txtUserName.Enabled = False
    End If
    
    'Posição Inicial da Janela
    If WindowTop = 0 Then
        WindowTop = 0
        WindowHeight = Height
    End If
    
    Move (Screen.Width - Width) / 2, WindowTop + WindowHeight
    
    Usuário = "" 'Campo de retorno para o programa GEAL, se este campo estiver vazio, não existe permissão no programa
    
    UsuárioAberto = AbreTabela(Dicionário, "USUÁRIO", "USUÁRIO", DBUsuário, TBLUsuário, TBLTabela, dbOpenTable)
    If UsuárioAberto Then
        IndiceAtivoUsuário = "USUÁRIO1"
        TBLUsuário.Index = IndiceAtivoUsuário
    Else
        MsgBox "Não consegui abrir a tabela 'Usuário' !", vbCritical, "Erro"
        Exit Sub
    End If
    
    'Verifica o ultimo usuário do sistema
    lPula = True
    txtUserName = Trim(GetRegistryString("Geal", "Geral", "Usuário"))
    lPula = False
    
    If txtUserName = "" Then
        lPula = True
        txtUserName = "ADMIN"
        lPula = False
    End If
    txtUserName.SelStart = 0
    txtUserName.SelLength = Len(txtUserName.Text)
End Sub
Private Sub Form_Unload(Cancel As Integer)
    If UsuárioAberto Then
        TBLUsuário.Close
    End If
    
    If Usuário <> Empty And GravaUsuário Then
        SetRegistryString "Geal", "Geral", "Usuário", Usuário
    End If
    
    Fechado = True
End Sub
Private Sub txtUserName_Change()
    If lPula Then
        Exit Sub
    End If
    FormatMask "@! AAAAAA", txtUserName
End Sub
Private Sub txtUserName_GotFocus()
    txtUserName.SelStart = 0
    txtUserName.SelLength = Len(txtUserName.Text)
End Sub
