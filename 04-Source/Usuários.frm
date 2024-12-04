VERSION 5.00
Begin VB.Form frmUsu�rios 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Usu�rios"
   ClientHeight    =   3360
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7500
   Icon            =   "Usu�rios.frx":0000
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
   Begin VB.Frame frUsu�rios 
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
      Begin VB.CommandButton cmdMudan�aDeSenha 
         Caption         =   "&Mudan�a de Senha"
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
      Begin VB.ListBox lstUsu�rio 
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
      Begin VB.Label lblUsu�rios 
         Caption         =   "Usu�rio"
         Height          =   195
         Left            =   300
         TabIndex        =   9
         Top             =   270
         Width           =   705
      End
   End
End
Attribute VB_Name = "frmUsu�rios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim TBLUsu�rio As Table
Dim Usu�rioAberto As Boolean
Dim IndiceAtivoUsu�rio As String

Dim TBLGrupos As Table
Dim GruposAberto As Boolean
Dim IndiceGruposAtivo$

Dim TBLUsu�rioGrupo As Table
Dim Usu�rioGrupoAberto As Boolean
Dim IndiceUsu�rioGrupoAtivo$

Dim lInserir As Boolean
Dim lAlterar As Boolean
Dim mFechar As Boolean

Dim StatusBarAviso$

Public lAtualizar As Boolean
Private Sub Bot�es(ByVal Valor As Boolean)
    cmdAlterar.Enabled = Valor
    cmdExcluir.Enabled = Valor
    cmdMudan�aDeSenha.Enabled = Valor
End Sub
Private Function BuscaGrupo(ByVal C�digo As Long) As String
    TBLGrupos.Seek "=", C�digo
    
    If TBLGrupos.NoMatch Then
        BuscaGrupo = ""
        Exit Function
    End If
    
    BuscaGrupo = TBLGrupos("DESCRI��O")
End Function
Private Function BuscaGrupoDescri��o(ByVal Descri��o As String) As Long
    Dim Bookmark
    
    
    TBLGrupos.Index = "GRUPO2"
    TBLGrupos.Seek "=", Descri��o
    If TBLGrupos.NoMatch Then
        BuscaGrupoDescri��o = 0
        Exit Function
    End If
    Bookmark = TBLGrupos.Bookmark
    TBLGrupos.Index = IndiceGruposAtivo
    TBLGrupos.Bookmark = Bookmark
    
    BuscaGrupoDescri��o = TBLGrupos("C�DIGO")
End Function
Private Sub FillUsu�rio(Optional ByVal CampoChave)
    Dim Cont%
    
    lstUsu�rio.Clear
    
    If IsMissing(CampoChave) Then
        CampoChave = ""
    End If
    
    If TBLUsu�rio.RecordCount <= 0 Then
        Bot�es False
        Exit Sub
    End If
    
    TBLUsu�rio.MoveFirst
    
    If TBLUsu�rio.EOF Or TBLUsu�rio.BOF Then
        Bot�es False
        Exit Sub
    End If
    
    Bot�es True
    
    Do While Not TBLUsu�rio.EOF
        lstUsu�rio.AddItem TBLUsu�rio("USERNAME")
        TBLUsu�rio.MoveNext
    Loop
    
    If CampoChave = "" Then
        lstUsu�rio.ListIndex = 0
        Exit Sub
    End If
    
    For Cont = 0 To lstUsu�rio.ListCount - 1
        If lstUsu�rio.List(Cont) = CampoChave Then
            lstUsu�rio.ListIndex = Cont
            Exit For
        End If
    Next
    
End Sub
Private Sub Excluir()
    On Error Resume Next
    
    Dim Confirma��o As Integer, Msg1$, Msg2$, CampoChave$
    
    StatusBarAviso = "Exclus�o"
    BarraDeStatus StatusBarAviso
    
    Msg1 = "Voc� est� preste a apagar um usu�rio !"
    Msg2 = "Tem certeza?"
    Msg2 = String(((Len(Msg1) - Len(Msg2)) / 2), " ") + Msg2
    Confirma��o = MsgBox(Msg1 + vbCr + Msg2, vbYesNo + vbQuestion + vbDefaultButton2, "Confirma��o")
    
    If Confirma��o = vbNo Then
        Exit Sub
    End If
    
    If Not PosRecords(lstUsu�rio.List(lstUsu�rio.ListIndex)) Then
        Exit Sub
    End If
    
    If TBLUsu�rio("USERNAME") = "ADMIN" Then
        MsgBox "O usu�rio 'ADMIN' n�o pode ser exclu�do!", vbCritical, "Aviso"
        Exit Sub
    End If
    
    WS.BeginTrans
    
    TBLUsu�rio.Delete
    
    If Err <> 0 Then
        GeraMensagemDeErro "Usu�rios - Excluir " & TBLUsu�rio("USERNAME"), True
        StatusBarAviso = "Falha na exclus�o"
        BarraDeStatus StatusBarAviso
        Exit Sub
    End If
    
    WS.CommitTrans
    
    StatusBarAviso = "Exclus�o bem sucedida"
    BarraDeStatus StatusBarAviso
    
    If TBLUsu�rio.RecordCount = 0 Then
        GoTo Fim
    End If
    
    If TBLUsu�rio.BOF Then
        TBLUsu�rio.MoveFirst
    ElseIf TBLUsu�rio.EOF Then
        TBLUsu�rio.MoveLast
    Else
        TBLUsu�rio.MovePrevious
        If TBLUsu�rio.BOF Then
            TBLUsu�rio.MoveNext
        End If
    End If
Fim:
    If TBLUsu�rio.RecordCount > 0 Then
        CampoChave = TBLUsu�rio("USERNAME")
        FillUsu�rio CampoChave
    Else
        FillUsu�rio
    End If
End Sub
Public Sub ExcluirGrupo()
    On Error GoTo Erro
    
    Dim Chave As String, Chave1 As String, Chave2 As Long
    Dim Bookmark
    
    Chave1 = lstUsu�rio.List(lstUsu�rio.ListIndex)
    Chave2 = BuscaGrupoDescri��o(lstGrupo.List(lstGrupo.ListIndex))
    
    Chave = Chave1 & Chave2
    
    TBLUsu�rioGrupo.Index = "USU�RIOGRUPO2"
    TBLUsu�rioGrupo.Seek "=", Chave1, Chave2
    If TBLUsu�rioGrupo.NoMatch Then
        MsgBox "N�o encontrei !" & vbCr & "USU�RIO: " & lstUsu�rio.List(lstUsu�rio.ListIndex) & vbCr & "GRUPO: " & lstGrupo.List(lstGrupo.ListIndex), vbCritical, "Aviso"
        TBLUsu�rioGrupo.Index = IndiceUsu�rioGrupoAtivo
        Exit Sub
    End If
    
    Bookmark = TBLUsu�rioGrupo.Bookmark
    TBLUsu�rioGrupo.Index = IndiceUsu�rioGrupoAtivo
    TBLUsu�rioGrupo.Bookmark = Bookmark
    
    TBLUsu�rioGrupo.Delete
    
    lstUsu�rio_Click
    
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
    TBLUsu�rio.Seek "=", Chave
    If TBLUsu�rio.NoMatch Then
        MsgBox "N�o consegui encontrar o UserName " + Chave, vbExclamation, "Erro"
        PosRecords = False
    Else
        PosRecords = True
    End If
End Function
Private Sub cmdAlterar_Click()
    frmUsu�rioCadastro.TipoOpera��o = vbAlterar
    frmUsu�rioCadastro.CampoChave = lstUsu�rio.List(lstUsu�rio.ListIndex)
    frmUsu�rioCadastro.Show vbModal
    If Not frmUsu�rioCadastro.Cancel Then
        FillUsu�rio lstUsu�rio.List(lstUsu�rio.ListIndex)
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
    frmUsu�rioCadastro.TipoOpera��o = vbIncluir
    frmUsu�rioCadastro.Show vbModal
    If Not frmUsu�rioCadastro.Cancel Then
        FillUsu�rio frmUsu�rioCadastro.CampoChave
    End If
End Sub
Private Sub cmdIncluirGrupo_Click()
    frmIncluirGrupo.Show 1
    If frmIncluirGrupo.GrupoEscolhido <> Empty Then
        lstGrupo.AddItem frmIncluirGrupo.GrupoEscolhido
        TBLUsu�rioGrupo.AddNew
        TBLUsu�rioGrupo("USERNAME") = lstUsu�rio.List(lstUsu�rio.ListIndex)
        TBLUsu�rioGrupo("C�DIGO DO GRUPO") = frmIncluirGrupo.GrupoC�digo
        TBLUsu�rioGrupo.Update
    End If
    Set frmIncluirGrupo = Nothing
End Sub
Private Sub cmdMudan�aDeSenha_Click()
    frmMudan�aDeSenha.Usu�rio = lstUsu�rio.List(lstUsu�rio.ListIndex)
    frmMudan�aDeSenha.Show 1
End Sub
Private Sub Form_Activate()
    If Not Usu�rioAberto Then
        Unload Me
        Exit Sub
    End If
    
    If Not GruposAberto Then
        Unload Me
        Exit Sub
    End If
    
    If Not Usu�rioGrupoAberto Then
        Unload Me
        Exit Sub
    End If
    
    If lAtualizar Then
        Bot�oAtualizar True
    Else
        Bot�oAtualizar False
    End If
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
    
    GruposAberto = AbreTabela(Dicion�rio, "USU�RIO", "GRUPO", DBUsu�rio, TBLGrupos, TBLTabela, dbOpenTable)
    
    If GruposAberto Then
        IndiceGruposAtivo = "GRUPO1"
        TBLGrupos.Index = IndiceGruposAtivo
    Else
        MsgBox "N�o consegui abrir a tabela 'GRUPO' !", vbCritical, "Erro"
        Exit Sub
    End If
    
    Usu�rioGrupoAberto = AbreTabela(Dicion�rio, "USU�RIO", "USU�RIO - GRUPO", DBUsu�rio, TBLUsu�rioGrupo, TBLTabela, dbOpenTable)
    
    If Usu�rioGrupoAberto Then
        IndiceUsu�rioGrupoAtivo = "USU�RIOGRUPO1"
        TBLUsu�rioGrupo.Index = IndiceUsu�rioGrupoAtivo
    Else
        MsgBox "N�o consegui abrir a tabela 'GRUPO' !", vbCritical, "Erro"
        Exit Sub
    End If
    
    FillUsu�rio
End Sub
Private Sub Form_Unload(Cancel As Integer)
    If Usu�rioAberto Then
        TBLUsu�rio.Close
    End If
    
    If GruposAberto Then
        TBLGrupos.Close
    End If
    
    Set frmUsu�rios = Nothing
End Sub
Private Sub lstUsu�rio_Click()
    Dim Usu�rio As String
    Dim Cont As Integer
    
    Usu�rio = lstUsu�rio.List(lstUsu�rio.ListIndex)
    
    If Trim(Usu�rio) = "ADMIN" Then
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
    
    TBLUsu�rioGrupo.Seek "=", Usu�rio
    lstGrupo.Clear
    
    If TBLUsu�rioGrupo.NoMatch Then
        Exit Sub
    End If
    
    Do While Not TBLUsu�rioGrupo.EOF And TBLUsu�rioGrupo("USERNAME") = Usu�rio
        lstGrupo.AddItem BuscaGrupo(TBLUsu�rioGrupo("C�DIGO DO GRUPO"))
        TBLUsu�rioGrupo.MoveNext
        If TBLUsu�rioGrupo.EOF Then
            Exit Do
        End If
    Loop
End Sub
Private Sub lstUsu�rio_Scroll()
    lstUsu�rio_Click
End Sub
