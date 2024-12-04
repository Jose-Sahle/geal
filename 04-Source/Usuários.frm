VERSION 5.00
Begin VB.Form frmUsuários 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Usuários"
   ClientHeight    =   3360
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7500
   Icon            =   "Usuários.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3360
   ScaleWidth      =   7500
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox lstGrupo 
      Height          =   2400
      Left            =   4860
      TabIndex        =   7
      Top             =   540
      Width           =   2415
   End
   Begin VB.CommandButton cmdExcluir 
      Caption         =   "&Excluir"
      Height          =   345
      Left            =   2880
      TabIndex        =   3
      Top             =   1230
      Width           =   1695
   End
   Begin VB.CommandButton cmdAlterar 
      Caption         =   "&Alterar"
      Height          =   345
      Left            =   2880
      TabIndex        =   2
      Top             =   870
      Width           =   1695
   End
   Begin VB.Frame frUsuários 
      Height          =   3345
      Left            =   0
      TabIndex        =   8
      Top             =   0
      Width           =   7485
      Begin VB.CommandButton cmdExcluirGrupo 
         Caption         =   "Excl&uir Grupo"
         Height          =   345
         Left            =   2880
         TabIndex        =   11
         Top             =   2310
         Width           =   1695
      End
      Begin VB.CommandButton cmdIncluirGrupo 
         Caption         =   "Incluir &Grupo"
         Height          =   345
         Left            =   2880
         TabIndex        =   5
         Top             =   1950
         Width           =   1695
      End
      Begin VB.CommandButton cmdFechar 
         Caption         =   "&Fechar"
         Height          =   345
         Left            =   2880
         TabIndex        =   6
         Top             =   2670
         Width           =   1695
      End
      Begin VB.CommandButton cmdMudançaDeSenha 
         Caption         =   "&Mudança de Senha"
         Height          =   345
         Left            =   2880
         TabIndex        =   4
         Top             =   1590
         Width           =   1695
      End
      Begin VB.CommandButton cmdIncluir 
         Caption         =   "&Incluir"
         Height          =   345
         Left            =   2880
         TabIndex        =   1
         Top             =   510
         Width           =   1695
      End
      Begin VB.ListBox lstUsuário 
         Height          =   2400
         Left            =   180
         TabIndex        =   0
         Top             =   540
         Width           =   2415
      End
      Begin VB.Label lblGrupo 
         Caption         =   "Grupo"
         Height          =   225
         Left            =   4980
         TabIndex        =   10
         Top             =   270
         Width           =   1455
      End
      Begin VB.Label lblUsuários 
         Caption         =   "Usuário"
         Height          =   195
         Left            =   300
         TabIndex        =   9
         Top             =   270
         Width           =   705
      End
   End
End
Attribute VB_Name = "frmUsuários"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim TBLUsuário As Table
Dim UsuárioAberto As Boolean
Dim IndiceAtivoUsuário As String

Dim TBLGrupos As Table
Dim GruposAberto As Boolean
Dim IndiceGruposAtivo$

Dim TBLUsuárioGrupo As Table
Dim UsuárioGrupoAberto As Boolean
Dim IndiceUsuárioGrupoAtivo$

Dim lInserir As Boolean
Dim lAlterar As Boolean
Dim mFechar As Boolean

Dim StatusBarAviso$

Public lAtualizar As Boolean
Private Sub Botões(ByVal Valor As Boolean)
    cmdAlterar.Enabled = Valor
    cmdExcluir.Enabled = Valor
    cmdMudançaDeSenha.Enabled = Valor
End Sub
Private Function BuscaGrupo(ByVal Código As Long) As String
    TBLGrupos.Seek "=", Código
    
    If TBLGrupos.NoMatch Then
        BuscaGrupo = ""
        Exit Function
    End If
    
    BuscaGrupo = TBLGrupos("DESCRIÇÃO")
End Function
Private Function BuscaGrupoDescrição(ByVal Descrição As String) As Long
    Dim Bookmark
    
    
    TBLGrupos.Index = "GRUPO2"
    TBLGrupos.Seek "=", Descrição
    If TBLGrupos.NoMatch Then
        BuscaGrupoDescrição = 0
        Exit Function
    End If
    Bookmark = TBLGrupos.Bookmark
    TBLGrupos.Index = IndiceGruposAtivo
    TBLGrupos.Bookmark = Bookmark
    
    BuscaGrupoDescrição = TBLGrupos("CÓDIGO")
End Function
Private Sub FillUsuário(Optional ByVal CampoChave)
    Dim Cont%
    
    lstUsuário.Clear
    
    If IsMissing(CampoChave) Then
        CampoChave = ""
    End If
    
    If TBLUsuário.RecordCount <= 0 Then
        Botões False
        Exit Sub
    End If
    
    TBLUsuário.MoveFirst
    
    If TBLUsuário.EOF Or TBLUsuário.BOF Then
        Botões False
        Exit Sub
    End If
    
    Botões True
    
    Do While Not TBLUsuário.EOF
        lstUsuário.AddItem TBLUsuário("USERNAME")
        TBLUsuário.MoveNext
    Loop
    
    If CampoChave = "" Then
        lstUsuário.ListIndex = 0
        Exit Sub
    End If
    
    For Cont = 0 To lstUsuário.ListCount - 1
        If lstUsuário.List(Cont) = CampoChave Then
            lstUsuário.ListIndex = Cont
            Exit For
        End If
    Next
    
End Sub
Private Sub Excluir()
    On Error Resume Next
    
    Dim Confirmação As Integer, Msg1$, Msg2$, CampoChave$
    
    StatusBarAviso = "Exclusão"
    BarraDeStatus StatusBarAviso
    
    Msg1 = "Você está preste a apagar um usuário !"
    Msg2 = "Tem certeza?"
    Msg2 = String(((Len(Msg1) - Len(Msg2)) / 2), " ") + Msg2
    Confirmação = MsgBox(Msg1 + vbCr + Msg2, vbYesNo + vbQuestion + vbDefaultButton2, "Confirmação")
    
    If Confirmação = vbNo Then
        Exit Sub
    End If
    
    If Not PosRecords(lstUsuário.List(lstUsuário.ListIndex)) Then
        Exit Sub
    End If
    
    If TBLUsuário("USERNAME") = "ADMIN" Then
        MsgBox "O usuário 'ADMIN' não pode ser excluído!", vbCritical, "Aviso"
        Exit Sub
    End If
    
    WS.BeginTrans
    
    TBLUsuário.Delete
    
    If Err <> 0 Then
        GeraMensagemDeErro "Usuários - Excluir " & TBLUsuário("USERNAME"), True
        StatusBarAviso = "Falha na exclusão"
        BarraDeStatus StatusBarAviso
        Exit Sub
    End If
    
    WS.CommitTrans
    
    StatusBarAviso = "Exclusão bem sucedida"
    BarraDeStatus StatusBarAviso
    
    If TBLUsuário.RecordCount = 0 Then
        GoTo Fim
    End If
    
    If TBLUsuário.BOF Then
        TBLUsuário.MoveFirst
    ElseIf TBLUsuário.EOF Then
        TBLUsuário.MoveLast
    Else
        TBLUsuário.MovePrevious
        If TBLUsuário.BOF Then
            TBLUsuário.MoveNext
        End If
    End If
Fim:
    If TBLUsuário.RecordCount > 0 Then
        CampoChave = TBLUsuário("USERNAME")
        FillUsuário CampoChave
    Else
        FillUsuário
    End If
End Sub
Public Sub ExcluirGrupo()
    On Error GoTo Erro
    
    Dim Chave As String, Chave1 As String, Chave2 As Long
    Dim Bookmark
    
    Chave1 = lstUsuário.List(lstUsuário.ListIndex)
    Chave2 = BuscaGrupoDescrição(lstGrupo.List(lstGrupo.ListIndex))
    
    Chave = Chave1 & Chave2
    
    TBLUsuárioGrupo.Index = "USUÁRIOGRUPO2"
    TBLUsuárioGrupo.Seek "=", Chave1, Chave2
    If TBLUsuárioGrupo.NoMatch Then
        MsgBox "Não encontrei !" & vbCr & "USUÁRIO: " & lstUsuário.List(lstUsuário.ListIndex) & vbCr & "GRUPO: " & lstGrupo.List(lstGrupo.ListIndex), vbCritical, "Aviso"
        TBLUsuárioGrupo.Index = IndiceUsuárioGrupoAtivo
        Exit Sub
    End If
    
    Bookmark = TBLUsuárioGrupo.Bookmark
    TBLUsuárioGrupo.Index = IndiceUsuárioGrupoAtivo
    TBLUsuárioGrupo.Bookmark = Bookmark
    
    TBLUsuárioGrupo.Delete
    
    lstUsuário_Click
    
    Exit Sub
    
Erro:
    GeraMensagemDeErro "ExcluirGrupo - " & Chave

End Sub
Public Function GetCountGrupo() As Integer
    GetCountGrupo = lstGrupo.ListCount
End Function
Public Function GetGrupo(ByVal Item) As String
    Item = Item - 1
    
    GetGrupo = lstGrupo.List(Item)
End Function
Private Function PosRecords(ByVal Chave$)
    TBLUsuário.Seek "=", Chave
    If TBLUsuário.NoMatch Then
        MsgBox "Não consegui encontrar o UserName " + Chave, vbExclamation, "Erro"
        PosRecords = False
    Else
        PosRecords = True
    End If
End Function
Private Sub cmdAlterar_Click()
    frmUsuárioCadastro.TipoOperação = vbAlterar
    frmUsuárioCadastro.CampoChave = lstUsuário.List(lstUsuário.ListIndex)
    frmUsuárioCadastro.Show vbModal
    If Not frmUsuárioCadastro.Cancel Then
        FillUsuário lstUsuário.List(lstUsuário.ListIndex)
    End If
End Sub
Private Sub cmdExcluir_Click()
    Excluir
End Sub
Private Sub cmdExcluirGrupo_Click()
    ExcluirGrupo
End Sub
Private Sub cmdFechar_Click()
    Unload Me
End Sub
Private Sub cmdIncluir_Click()
    frmUsuárioCadastro.TipoOperação = vbIncluir
    frmUsuárioCadastro.Show vbModal
    If Not frmUsuárioCadastro.Cancel Then
        FillUsuário frmUsuárioCadastro.CampoChave
    End If
End Sub
Private Sub cmdIncluirGrupo_Click()
    frmIncluirGrupo.Show 1
    If frmIncluirGrupo.GrupoEscolhido <> Empty Then
        lstGrupo.AddItem frmIncluirGrupo.GrupoEscolhido
        TBLUsuárioGrupo.AddNew
        TBLUsuárioGrupo("USERNAME") = lstUsuário.List(lstUsuário.ListIndex)
        TBLUsuárioGrupo("CÓDIGO DO GRUPO") = frmIncluirGrupo.GrupoCódigo
        TBLUsuárioGrupo.Update
    End If
    Set frmIncluirGrupo = Nothing
End Sub
Private Sub cmdMudançaDeSenha_Click()
    frmMudançaDeSenha.Usuário = lstUsuário.List(lstUsuário.ListIndex)
    frmMudançaDeSenha.Show 1
End Sub
Private Sub Form_Activate()
    If Not UsuárioAberto Then
        Unload Me
        Exit Sub
    End If
    
    If Not GruposAberto Then
        Unload Me
        Exit Sub
    End If
    
    If Not UsuárioGrupoAberto Then
        Unload Me
        Exit Sub
    End If
    
    If lAtualizar Then
        BotãoAtualizar True
    Else
        BotãoAtualizar False
    End If
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
    
    GruposAberto = AbreTabela(Dicionário, "USUÁRIO", "GRUPO", DBUsuário, TBLGrupos, TBLTabela, dbOpenTable)
    
    If GruposAberto Then
        IndiceGruposAtivo = "GRUPO1"
        TBLGrupos.Index = IndiceGruposAtivo
    Else
        MsgBox "Não consegui abrir a tabela 'GRUPO' !", vbCritical, "Erro"
        Exit Sub
    End If
    
    UsuárioGrupoAberto = AbreTabela(Dicionário, "USUÁRIO", "USUÁRIO - GRUPO", DBUsuário, TBLUsuárioGrupo, TBLTabela, dbOpenTable)
    
    If UsuárioGrupoAberto Then
        IndiceUsuárioGrupoAtivo = "USUÁRIOGRUPO1"
        TBLUsuárioGrupo.Index = IndiceUsuárioGrupoAtivo
    Else
        MsgBox "Não consegui abrir a tabela 'GRUPO' !", vbCritical, "Erro"
        Exit Sub
    End If
    
    FillUsuário
End Sub
Private Sub Form_Unload(Cancel As Integer)
    If UsuárioAberto Then
        TBLUsuário.Close
    End If
    
    If GruposAberto Then
        TBLGrupos.Close
    End If
    
    Set frmUsuários = Nothing
End Sub
Private Sub lstUsuário_Click()
    Dim Usuário As String
    Dim Cont As Integer
    
    Usuário = lstUsuário.List(lstUsuário.ListIndex)
    
    If Trim(Usuário) = "ADMIN" Then
        cmdAlterar.Enabled = False
        cmdExcluir.Enabled = False
        cmdIncluirGrupo.Enabled = False
        cmdExcluirGrupo.Enabled = False
        lstGrupo.Enabled = False
    Else
        cmdAlterar.Enabled = True
        cmdExcluir.Enabled = True
        cmdIncluirGrupo.Enabled = True
        cmdExcluirGrupo.Enabled = True
        lstGrupo.Enabled = True
    End If
    
    TBLUsuárioGrupo.Seek "=", Usuário
    lstGrupo.Clear
    
    If TBLUsuárioGrupo.NoMatch Then
        Exit Sub
    End If
    
    Do While Not TBLUsuárioGrupo.EOF And TBLUsuárioGrupo("USERNAME") = Usuário
        lstGrupo.AddItem BuscaGrupo(TBLUsuárioGrupo("CÓDIGO DO GRUPO"))
        TBLUsuárioGrupo.MoveNext
        If TBLUsuárioGrupo.EOF Then
            Exit Do
        End If
    Loop
End Sub
Private Sub lstUsuário_Scroll()
    lstUsuário_Click
End Sub
