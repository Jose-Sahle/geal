VERSION 5.00
Begin VB.Form frmUsuárioCadastro 
   Caption         =   "Cadastro de Usuário"
   ClientHeight    =   1785
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6540
   Icon            =   "IncluirFuncionario.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   1785
   ScaleWidth      =   6540
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   5250
      TabIndex        =   4
      Top             =   1380
      Width           =   1245
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "&OK"
      Height          =   375
      Left            =   3960
      TabIndex        =   3
      Top             =   1380
      Width           =   1245
   End
   Begin VB.Frame frUserName 
      Height          =   1305
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   6525
      Begin VB.TextBox txtNomeDoFuncionário 
         Height          =   285
         Left            =   2580
         Locked          =   -1  'True
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   330
         Width           =   3825
      End
      Begin VB.TextBox txtCódigoDoFuncionário 
         Height          =   285
         Left            =   1950
         TabIndex        =   0
         Top             =   330
         Width           =   585
      End
      Begin VB.TextBox txtUserName 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   1110
         TabIndex        =   2
         Top             =   780
         Width           =   1425
      End
      Begin VB.Label lblCódigoDeFuncionário 
         Caption         =   "Código do Funcionário"
         Height          =   195
         Left            =   180
         TabIndex        =   7
         Top             =   360
         Width           =   1635
      End
      Begin VB.Label lblUserName 
         Caption         =   "Usuário"
         Height          =   225
         Left            =   180
         TabIndex        =   6
         Top             =   810
         Width           =   735
      End
   End
End
Attribute VB_Name = "frmUsuárioCadastro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim TBLUsuário As Table
Dim UsuárioAberto As Boolean
Dim IndiceAtivoUsuário As String

Dim lInserir As Boolean
Dim lAlterar As Boolean
Dim lPush As Boolean
Dim mFechar As Boolean
Dim mFirst As Boolean
Dim mPula As Boolean

Dim StatusBarAviso$

Public TipoOperação As Integer
Public CampoChave As String
Public Cancel As Boolean

Public lAtualizar As Boolean
Public Sub Gravar()
    If lInserir Then
        If SetRecords Then
            StatusBarAviso = "Inclusão bem sucedida"
            Unload Me
            GoTo Fim
            Exit Sub
        Else
            StatusBarAviso = "Falha na inclusão"
        End If
    Else
        If SetRecords Then
            StatusBarAviso = "Alteração bem sucedida"
            Unload Me
            GoTo Fim
            Exit Sub
        Else
            StatusBarAviso = "Falha na alteração"
        End If
    End If
    
    If txtCódigoDoFuncionário.Enabled Then
        txtCódigoDoFuncionário.SetFocus
    End If
Fim:
    BarraDeStatus StatusBarAviso
End Sub
Private Function PosRecords(ByVal Chave$)
    TBLUsuário.Seek "=", Chave
    If TBLUsuário.NoMatch Then
        MsgBox "Não consegui encontrar o UserName " + Chave, vbExclamation, "Erro"
        PosRecords = False
    Else
        PosRecords = True
    End If
End Function
Private Sub GetRecords()
    lPush = True
    mPula = True
    txtCódigoDoFuncionário = TBLUsuário("CÓDIGO DE FUNCIONÁRIO")
    txtNomeDoFuncionário = BuscaFuncionário(TBLUsuário("CÓDIGO DE FUNCIONÁRIO"))
    txtUserName = CampoChave
    lPush = False
    mPula = False
End Sub
Private Function SetRecords() As Boolean
    On Error GoTo Erro

    WS.BeginTrans
    
    If lInserir Then
        TBLUsuário.AddNew
    Else
        TBLUsuário.Edit
    End If
    
    If lInserir Then
        TBLUsuário("CÓDIGO DE FUNCIONÁRIO") = txtCódigoDoFuncionário
        TBLUsuário("SENHA") = ValidaSenha("GEAL")
    End If
    TBLUsuário("USERNAME") = txtUserName
    CampoChave = txtUserName
    
    If lInserir Then
        TBLUsuário("USERNAME - CRIA") = gUsuário
        TBLUsuário("DATA - CRIA") = Date
        TBLUsuário("HORA - CRIA") = Time
        TBLUsuário("USERNAME - ALTERA") = "VAZIO"
        TBLUsuário("DATA - ALTERA") = vbNull
        TBLUsuário("HORA - ALTERA") = vbNull
    End If
    If lAlterar Then
        TBLUsuário("USERNAME - ALTERA") = gUsuário
        TBLUsuário("DATA - ALTERA") = Date
        TBLUsuário("HORA - ALTERA") = Time
    End If
    TBLUsuário.Update
    
Erro:
    If Err <> 0 Then
        TBLUsuário.CancelUpdate
        GeraMensagemDeErro "UsuárioCadastro - SetRecords - " & txtUserName, True
        SetRecords = False
        Exit Function
    End If
    
    WS.CommitTrans 'Grava as alterações ou inclusões se não houverem erros
    
    SetRecords = True
End Function
Private Sub ZeraCampos()
    lPush = True
    mPula = True
    txtCódigoDoFuncionário = Empty
    txtNomeDoFuncionário = Empty
    txtUserName = ""
    mPula = False
    lPush = False
End Sub
Private Sub cmdCancelar_Click()
    Cancel = True
    Unload Me
End Sub
Private Sub cmdOK_Click()
    Gravar
    Cancel = False
End Sub
Private Sub Form_Activate()
    If Not UsuárioAberto And Not mFechar Then
        Unload Me
        Exit Sub
    End If
    
    If mFirst And TipoOperação = vbIncluir Then
        txtCódigoDoFuncionário.SetFocus
    Else
        txtUserName.SetFocus
    End If
    
    mFirst = False
    
    If lAtualizar Then
        BotãoAtualizar True
    Else
        BotãoAtualizar False
    End If
End Sub
Private Sub Form_Load()

    mFechar = False
    mPula = False
    
    UsuárioAberto = AbreTabela(Dicionário, "USUÁRIO", "USUÁRIO", DBUsuário, TBLUsuário, TBLTabela, dbOpenTable)
        
    If UsuárioAberto Then
        IndiceAtivoUsuário = "USUÁRIO1"
        TBLUsuário.Index = IndiceAtivoUsuário
    Else
        MsgBox "Não consegui abrir a tabela 'Usuário' !", vbCritical, "Erro"
        Exit Sub
    End If
    
    If TipoOperação = vbIncluir Then
        ZeraCampos
        lInserir = True
        lAlterar = False
    ElseIf TipoOperação = vbAlterar Then
        txtCódigoDoFuncionário.Enabled = False
        If PosRecords(CampoChave) Then
            GetRecords
            lInserir = False
            lAlterar = True
        Else
            mFechar = True
            Exit Sub
        End If
    End If
    mFirst = True
End Sub
Private Sub Form_Unload(Cancel As Integer)
    If UsuárioAberto Then
        TBLUsuário.Close
    End If
    
    Set frmUsuárioCadastro = Nothing
End Sub
Private Sub txtCódigoDoFuncionário_Change()
    FormatMask "99", txtCódigoDoFuncionário
End Sub
Private Sub txtCódigoDoFuncionário_LostFocus()
    If txtCódigoDoFuncionário = Empty Then
        Exit Sub
    End If
    If Not IsCorrectFuncionário(txtCódigoDoFuncionário) Then
        MsgBox "Funcionário não cadastrado!", vbInformation, "Aviso"
        Set frmEncontrar.DBBancoDeDados = DBUsuário
        frmEncontrar.LabelDescription = "Nome"
        frmEncontrar.NomeDaJanela = "Funcionário"
        frmEncontrar.Mensagem = "Nenhum funcionário foi selecionado!"
        frmEncontrar.BancoDeDados = "USUÁRIO"
        frmEncontrar.Tabela = "FUNCIONÁRIO"
        frmEncontrar.Indice = "1"
        frmEncontrar.CampoChave = "CÓDIGO"
        frmEncontrar.CampoPreencheLista = "NOME"
        frmEncontrar.Show vbModal
        txtCódigoDoFuncionário = frmEncontrar.Chave
        txtNomeDoFuncionário = frmEncontrar.Nome
    Else
        txtNomeDoFuncionário = BuscaFuncionário(Val(txtCódigoDoFuncionário))
    End If
End Sub
Private Sub txtUserName_Change()
    FormatMask "@! AAAAAA", txtUserName
End Sub
Private Sub txtUserName_KeyPress(KeyAscii As Integer)
    If Not lInserir Then
        If Not lPush Then
            lAlterar = True
            StatusBarAviso = "Alteração"
            BarraDeStatus StatusBarAviso
        End If
    End If
End Sub
